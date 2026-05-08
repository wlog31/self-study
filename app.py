"""
고등학생 자율학습 관리 시스템
Flask 웹 애플리케이션 진입점
"""

import logging
import os
import sqlite3
import uuid
from datetime import timedelta, date, datetime
from logging.handlers import RotatingFileHandler
from flask import Flask, redirect, url_for, session as flask_session
from flask_login import LoginManager
from flask_wtf.csrf import CSRFProtect
from models import db, User, StudyPeriodSetting
from constants import DEFAULT_PERIODS
from settings import init_default_settings
from audit import log_audit
from sqlalchemy import event
from sqlalchemy.engine import Engine


def _configure_logging(app):
    """감사 로거 + Flask 앱 로거를 파일 회전 핸들러에 연결한다.

    logs/audit.log: 구조화된 감사 이벤트 (5MB * 5개 회전)
    stdout: 콘솔 표시용 (Waitress 표준 출력으로 따라감)
    """
    log_dir = os.path.join(os.path.dirname(__file__), 'logs')
    os.makedirs(log_dir, exist_ok=True)

    fmt = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    file_handler = RotatingFileHandler(
        os.path.join(log_dir, 'audit.log'),
        maxBytes=5 * 1024 * 1024, backupCount=5, encoding='utf-8',
    )
    file_handler.setFormatter(fmt)
    file_handler.setLevel(logging.INFO)

    console = logging.StreamHandler()
    console.setFormatter(fmt)
    console.setLevel(logging.INFO)

    for name in ('self_study.audit', 'self_study.app'):
        lg = logging.getLogger(name)
        lg.setLevel(logging.INFO)
        # 중복 핸들러 방지 (reload 시)
        if not lg.handlers:
            lg.addHandler(file_handler)
            lg.addHandler(console)
        lg.propagate = False


# 부팅 시 기존 DB에 적용 여부를 검증할 핵심 무결성 제약 마커.
# (SQLAlchemy db.create_all()은 기존 테이블을 변경하지 않으므로,
#  models.py에 CheckConstraint를 추가해도 기존 운영 DB에는 자동 반영되지 않는다.
#  반드시 migrate_add_constraints_v2.py를 1회 실행해야 적용된다.)
_EXPECTED_DB_CONSTRAINTS = [
    ('users',         'ck_users_role'),
    ('users',         'ck_users_gender'),
    ('attendance',    'ck_attendance_status'),
    # NOTE: student_rooms.uq_room_seat은 의도적으로 제거됨 (남/여 zone이 같은
    # 좌석 번호를 공유하는 구조라 (room, seat) UNIQUE는 정상 배정도 막아버림).
    # 자세한 배경은 migrate_drop_room_seat_uq.py 참조.
]


def _verify_db_constraints(app):
    """기존 DB에 핵심 무결성 제약이 적용됐는지 검사한다.

    누락 시:
      - logs/audit.log에 WARNING (system.constraints_missing) 기록
      - 콘솔에 즉시 식별 가능한 경고 박스 출력
      - 앱은 정상 부팅 (운영 중단 X) — 관리자가 마이그레이션을 수행하도록 유도
    """
    db_path = os.path.join(os.path.dirname(__file__), 'instance', 'self_study.db')
    if not os.path.exists(db_path):
        return  # 신규 설치 — db.create_all()이 직전에 만들었으면 제약 포함됨

    missing = []
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        for table, marker in _EXPECTED_DB_CONSTRAINTS:
            cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (table,))
            row = cur.fetchone()
            if not row or not row[0] or marker not in row[0]:
                missing.append(f'{table}.{marker}')
    finally:
        conn.close()

    if missing:
        log_audit(
            'system.constraints_missing',
            level='warning',
            constraints=','.join(missing),
            action='run_migrate_add_constraints_v2',
        )
        print('\n' + '!' * 60)
        print('  [경고] 데이터 무결성 제약 누락 발견 - 다음 항목이 DB에 없음:')
        for m in missing:
            print(f'      - {m}')
        print('  -> 서버를 일시 중지한 뒤 다음 명령을 실행하십시오:')
        print('      python migrate_add_constraints_v2.py')
        print('!' * 60 + '\n')


def _backfill_session_tokens():
    """기존 사용자에게 NULL/빈 session_token이 있으면 UUID로 채운다.

    이 컬럼은 v1.7부터 도입됐지만 옛 DB에서 ALTER TABLE만 수행되고
    백필이 누락된 경우, 또는 직접 SQL로 사용자가 추가된 경우 NULL일 수 있다.
    NULL이면 다음 두 가지 문제가 발생한다.
      1) login_manager의 user_loader가 None 비교에서 항상 불일치로 판정 → 로그인 무한 실패
      2) migrate_add_constraints_v2.py의 NOT NULL 제약 재구축 시 SQL 오류

    부팅 시점에 멱등하게 채워서 두 위험을 차단한다.
    """
    try:
        rows = User.query.filter(
            (User.session_token.is_(None)) | (User.session_token == '')
        ).all()
    except Exception as e:
        # 컬럼 자체가 없는 옛 DB는 migrate.py를 먼저 실행해야 함
        log_audit('system.session_token_check_failed', level='error',
                  error=str(e), action='run_migrate_py_first')
        return

    if not rows:
        return

    for u in rows:
        u.session_token = str(uuid.uuid4())
    db.session.commit()
    log_audit('system.session_token_backfilled', level='warning', count=len(rows))
    print(f'\n[참고] {len(rows)}명 사용자의 session_token이 비어 있어 UUID로 채웠습니다.\n')


def reinitialize_after_db_change():
    """DB가 외부에서 교체된 직후 (예: db_restore) 호출되어 다음을 수행한다.

      - db.create_all(): 새 DB가 옛 스키마라 SystemSetting 등 누락 테이블이 있으면 생성
      - init_default_period_settings(): 자습 시간 기본값 시드
      - init_default_settings(): system_settings 8개 시드
      - _backfill_session_tokens(): NULL session_token 채움

    이 모두는 멱등하므로 정상 DB(이미 모두 있는)에서 호출해도 변경 없다.
    호출자는 반드시 app context 내에서 실행해야 한다.
    """
    db.create_all()
    init_default_period_settings()
    init_default_settings()
    _backfill_session_tokens()


def init_default_period_settings():
    """기본 자습 시간 설정을 DB에 초기화"""
    for day_type, periods in DEFAULT_PERIODS.items():
        for period, (start_time, end_time) in periods.items():
            # 이미 설정이 있는지 확인
            existing = StudyPeriodSetting.query.filter_by(
                day_type=day_type, period=period
            ).first()

            if not existing:
                setting = StudyPeriodSetting(
                    day_type=day_type,
                    period=period,
                    start_time=start_time,
                    end_time=end_time,
                    is_active=True
                )
                db.session.add(setting)

    db.session.commit()


@event.listens_for(Engine, 'connect')
def _set_sqlite_pragmas(dbapi_conn, _):
    """SQLite 연결마다 FK 강제 + WAL 모드 활성화"""
    if isinstance(dbapi_conn, sqlite3.Connection):
        cur = dbapi_conn.cursor()
        cur.execute('PRAGMA foreign_keys = ON')
        cur.execute('PRAGMA journal_mode = WAL')
        cur.close()


def create_app():
    app = Flask(__name__)

    # 로깅 먼저 구성 (초기화 과정에서 이미 이벤트 기록 가능하도록)
    _configure_logging(app)

    # 설정
    basedir = os.path.abspath(os.path.dirname(__file__))
    secret_key_file = os.path.join(basedir, 'instance', 'secret_key.txt')
    if os.path.exists(secret_key_file):
        with open(secret_key_file) as f:
            secret_key = f.read().strip()
    else:
        secret_key = os.urandom(24).hex()
        os.makedirs(os.path.dirname(secret_key_file), exist_ok=True)
        with open(secret_key_file, 'w') as f:
            f.write(secret_key)
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', secret_key)
    app.config['SQLALCHEMY_DATABASE_URI'] = \
        'sqlite:///' + os.path.join(basedir, 'instance', 'self_study.db')
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    # 세션 보안 설정
    app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=12)  # 12시간 후 자동 만료
    app.config['SESSION_COOKIE_HTTPONLY'] = True   # JS에서 세션 쿠키 접근 차단
    app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'  # CSRF 방어
    # 업로드 크기 상한 (메모리 압박 방어). 학교 DB는 수십 MB 수준.
    app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024   # 50 MB

    # instance 폴더 생성
    os.makedirs(os.path.join(basedir, 'instance'), exist_ok=True)

    # DB 초기화
    db.init_app(app)

    # CSRF 보호
    CSRFProtect(app)

    # Flask-Login 설정
    login_manager = LoginManager()
    login_manager.login_view = 'auth.login'
    login_manager.login_message = '로그인이 필요합니다.'
    login_manager.session_protection = 'basic'  # 브라우저 변경 시 세션 재발급
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(username):
        # username 기반 조회 (숫자 PK 의존 제거)
        user = User.query.filter_by(username=username).first()
        if user is None:
            return None
        # session_token 검증: DB 교체·복원 시 UUID가 달라지므로 기존 세션 즉시 차단
        if flask_session.get('_session_token') != user.session_token:
            return None
        return user

    # Blueprint 등록
    from auth import auth_bp
    from routes_student import student_bp
    from routes_teacher import teacher_bp
    from routes_admin import admin_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(student_bp, url_prefix='/student')
    app.register_blueprint(teacher_bp, url_prefix='/teacher')
    app.register_blueprint(admin_bp, url_prefix='/admin')

    # 메인 페이지 -> 로그인으로 리다이렉트
    @app.route('/')
    def index():
        return redirect(url_for('auth.login'))

    # DB 테이블 생성 및 기본 설정 초기화
    with app.app_context():
        db.create_all()
        init_default_period_settings()
        init_default_settings()
        _init_admin_account()
        # 옛 DB에서 session_token이 NULL인 행을 안전하게 채움 (login 꼬임 방지)
        _backfill_session_tokens()
        # 기존 DB가 마이그레이션 안 된 경우를 가시화 (db.create_all은 기존 테이블 미변경)
        _verify_db_constraints(app)

    # 야간 자동 조퇴 처리 스케줄러 시작
    _start_scheduler(app)

    return app


def _init_admin_account():
    """관리자 계정이 없으면 기본 admin 계정을 생성한다."""
    if User.query.filter_by(role='admin').first():
        return  # 이미 관리자 존재

    default_password = 'Admin1234!'
    admin = User(
        username='admin',
        name='관리자',
        role='admin',
        is_approved=True,
    )
    admin.set_password(default_password)
    db.session.add(admin)
    db.session.commit()

    log_audit('system.admin_account_created', username='admin')
    # 콘솔 안내는 유지 - 학교 관리자가 최초 설치 시 비밀번호 메모를 남겨야 함
    print("\n" + "=" * 50)
    print("  [관리자 계정 자동 생성]")
    print(f"  아이디: admin")
    print(f"  비밀번호: {default_password}")
    print("  ※ 첫 로그인 후 반드시 비밀번호를 변경하세요!")
    print("=" * 50 + "\n")


def _auto_early_leave(app):
    """매일 23:59 — 입실 QR을 찍었지만 퇴실 QR을 찍지 않은 학생을 조퇴로 처리.
    present뿐 아니라 late(지각 입실)도 대상에 포함한다."""
    from models import Attendance, AttendanceLog
    with app.app_context():
        today = date.today()
        changed = 0
        try:
            targets = Attendance.query.filter(
                Attendance.date == today,
                Attendance.status.in_(['present', 'late']),
                Attendance.checked_at.isnot(None),
                Attendance.checked_out_at.is_(None),
            ).all()
            for att in targets:
                old_status = att.status   # 'present' 또는 'late' — 동적 캡처
                att.status = 'early_leave'
                att.checked_out_at = datetime.combine(today, datetime.strptime('23:59', '%H:%M').time())
                att.study_minutes = 0  # 퇴실 미확인 — 학습시간 0으로 명시
                att.early_leave_note = (att.early_leave_note or '') or '퇴실미확인(자동)'
                db.session.add(AttendanceLog(
                    attendance_id=att.id,
                    changed_by=None,
                    old_status=old_status,
                    new_status='early_leave',
                    note='퇴실미확인(자동)',
                ))
                changed += 1
            db.session.commit()
            if changed:
                log_audit('system.auto_early_leave', date=str(today), count=changed)
        except Exception as e:
            db.session.rollback()
            log_audit('system.auto_early_leave_failed', level='error', error=str(e))


def _start_scheduler(app):
    """APScheduler로 매일 23:59에 자동 조퇴 처리 잡을 등록한다.
    waitress는 단일 프로세스이므로 스케줄러가 중복 실행되지 않는다."""
    try:
        from apscheduler.schedulers.background import BackgroundScheduler
        scheduler = BackgroundScheduler(timezone='Asia/Seoul')
        scheduler.add_job(
            _auto_early_leave,
            trigger='cron',
            hour=23, minute=59,
            args=[app],
            id='auto_early_leave',
            replace_existing=True,
        )
        scheduler.start()
        log_audit('system.scheduler_started')
    except Exception as e:
        log_audit('system.scheduler_failed', level='error', error=str(e))


def _get_lan_ip():
    """LAN 인터페이스의 실제 IP를 반환한다.
    gethostbyname(hostname)은 루프백이나 엉뚱한 주소를 돌려줄 수 있으므로,
    외부 연결 시도 소켓으로 라우팅되는 인터페이스 IP를 읽는다."""
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))   # 실제 패킷 전송 없음, 라우팅만 확인
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'


if __name__ == '__main__':
    from waitress import serve
    app = create_app()
    lan_ip = _get_lan_ip()
    print("=" * 50)
    print("  자율학습 관리 시스템")
    print(f"  http://{lan_ip}:5000  ← 학생/교사 모두 이 주소로 접속")
    print("=" * 50)
    serve(app, host='0.0.0.0', port=5000, threads=4)
