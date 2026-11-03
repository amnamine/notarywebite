"""
Application entry point and runner.
Run this directly with: python app.py
"""
import os
import click
from flask import Flask, request, session
from dotenv import load_dotenv

load_dotenv()

from core_engine.models import db, login_manager, csrf, babel, User
from routes.main import routes_bp

def create_app():
    app = Flask(__name__)
    
    # Configuration
    base_dir = os.path.abspath(os.path.dirname(__file__))
    app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'notary-boukhari-secret-key-2026')
    app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL', f"sqlite:///{os.path.join(base_dir, 'data', 'app.db')}")
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB max upload
    app.config['DATA_DIR'] = os.path.join(base_dir, 'data')
    app.config['DOCX_TEMPLATES_DIR'] = os.path.join(base_dir, 'docx_templates')
    app.config['ANTHROPIC_API_KEY'] = os.getenv('ANTHROPIC_API_KEY', '')
    app.config['AI_MODEL'] = os.getenv('AI_MODEL', 'claude-3-5-sonnet-latest')

    os.makedirs(app.config['DATA_DIR'], exist_ok=True)
    os.makedirs(os.path.join(app.config['DATA_DIR'], 'jobs'), exist_ok=True)

    # Initialize extensions
    db.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'routes.login'
    login_manager.login_message = 'يرجى تسجيل الدخول للوصول إلى هذه الصفحة.'
    csrf.init_app(app)

    def get_locale():
        if 'locale' in session:
            return session['locale']
        cookie_lang = request.cookies.get('edd_locale')
        if cookie_lang in ['ar', 'fr', 'en']:
            return cookie_lang
        return request.accept_languages.best_match(['ar', 'fr', 'en']) or 'ar'

    babel.init_app(app, locale_selector=get_locale)

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    # Context processors for template i18n and RTL direction
    @app.context_processor
    def inject_locale_and_meta():
        curr_lang = session.get('locale') or request.cookies.get('edd_locale') or 'ar'
        is_rtl = (curr_lang == 'ar')
        return {
            'curr_lang': curr_lang,
            'dir': 'rtl' if is_rtl else 'ltr',
            'is_rtl': is_rtl
        }

    # Register routes
    app.register_blueprint(routes_bp)

    # Database initialization & default admin
    with app.app_context():
        db.create_all()
        # Ensure default admin exists if no users
        if User.query.count() == 0:
            default_admin = User(username='admin', role='admin')
            default_admin.set_password('admin123')
            db.session.add(default_admin)
            db.session.commit()
            print("Created default admin: admin / admin123")

    return app

# CLI Commands
app = create_app()

@app.cli.command("create-admin")
@click.argument("username")
@click.argument("password")
def create_admin_cmd(username, password):
    """Create a new admin user."""
    with app.app_context():
        u = User.query.filter_by(username=username).first()
        if u:
            u.role = 'admin'
            u.set_password(password)
            u.is_active_account = True
            db.session.commit()
            print(f"Admin '{username}' updated successfully.")
        else:
            u = User(username=username, role='admin')
            u.set_password(password)
            db.session.add(u)
            db.session.commit()
            print(f"Admin '{username}' created successfully.")

if __name__ == '__main__':
    print("=" * 60)
    print("  Office of Maitre Boukhari Meriem - Notary")
    print("  Server running at: http://127.0.0.1:5000")
    print("  Default Login: admin / admin123")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=False)
