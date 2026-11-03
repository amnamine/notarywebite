"""
Database models and extensions setup.
- User (admin/staff, password hashing with werkzeug.security)
- ActivityLog (user, tool, original_filename, pages/tables count, warning count, timestamp)
- Job (background OCR progress tracking)
"""
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin, LoginManager
from flask_wtf.csrf import CSRFProtect
from flask_babel import Babel
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()
login_manager = LoginManager()
csrf = CSRFProtect()
babel = Babel()

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), default='staff')  # 'admin' or 'staff'
    is_active_account = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def is_active(self):
        return self.is_active_account

    @property
    def is_admin(self):
        return self.role == 'admin'

    def set_password(self, password: str):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)


class ActivityLog(db.Model):
    __tablename__ = 'activity_logs'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    username = db.Column(db.String(80), nullable=False)
    tool_name = db.Column(db.String(100), nullable=False)  # 'ijraa' or 'tables556'
    original_filename = db.Column(db.String(255), nullable=False)
    pages_count = db.Column(db.Integer, default=0)
    tables_count = db.Column(db.Integer, default=0)
    warnings_count = db.Column(db.Integer, default=0)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User', backref=db.backref('logs', lazy=True))


class Job(db.Model):
    __tablename__ = 'jobs'
    id = db.Column(db.String(64), primary_key=True)  # uuid
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    original_filename = db.Column(db.String(255), nullable=False)
    status = db.Column(db.String(32), default='pending')  # pending, processing, completed, stopped, failed
    current_page = db.Column(db.Integer, default=0)
    total_pages = db.Column(db.Integer, default=0)
    read_ok = db.Column(db.Integer, default=0)
    failed_count = db.Column(db.Integer, default=0)
    warnings_count = db.Column(db.Integer, default=0)
    error_message = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
