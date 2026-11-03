"""
Web routes for Authentication, Tool A, Tool B, and Admin operations.
"""
import os
import json
import uuid
import threading
from datetime import datetime
from flask import (
    Blueprint, render_template, request, redirect, url_for,
    flash, session, send_file, jsonify, current_app, make_response
)
from flask_login import login_user, logout_user, login_required, current_user
import io

from core_engine.models import db, User, ActivityLog, Job
from core_engine.ijraa_converter import analyse_docx, convert_ijraa

routes_bp = Blueprint('routes', __name__)

# In-memory simple login rate limiter: ip -> (failed_count, wait_until_timestamp)
LOGIN_LIMITS = {}

@routes_bp.route('/set-locale/<lang>')
def set_locale(lang):
    if lang in ['ar', 'fr', 'en']:
        session['locale'] = lang
        response = redirect(request.referrer or url_for('routes.index'))
        response.set_cookie('edd_locale', lang, max_age=30*86400)
        return response
    return redirect(url_for('routes.index'))


@routes_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('routes.index'))

    client_ip = request.remote_addr or 'unknown'
    now_ts = datetime.utcnow().timestamp()

    if client_ip in LOGIN_LIMITS:
        failed_count, wait_until = LOGIN_LIMITS[client_ip]
        if now_ts < wait_until:
            wait_sec = int(wait_until - now_ts)
            flash(f"محاولات كثيرة خاطئة. يرجى الانتظار {wait_sec} ثانية.", 'warn')
            return render_template('login.html')

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()

        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            if not user.is_active:
                flash("تم تعطيل هذا الحساب. يرجى مراجعة المسؤول.", 'warn')
                return render_template('login.html')
            
            LOGIN_LIMITS.pop(client_ip, None)
            login_user(user)
            return redirect(url_for('routes.index'))
        else:
            failed, wait_until = LOGIN_LIMITS.get(client_ip, (0, 0))
            failed += 1
            if failed >= 5:
                LOGIN_LIMITS[client_ip] = (0, now_ts + 30)
                flash("رمز خاطئ 5 مرات. يرجى الانتظار 30 ثانية.", 'warn')
            else:
                LOGIN_LIMITS[client_ip] = (failed, 0)
                flash("اسم المستخدم أو كلمة المرور غير صحيحة.", 'warn')

    return render_template('login.html')


@routes_bp.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('routes.login'))


@routes_bp.route('/')
@login_required
def index():
    return render_template('index.html')


# ----------------- Tool A Endpoints -----------------
@routes_bp.route('/api/ijraa/analyse', methods=['POST'])
@login_required
def ijraa_analyse():
    if 'file' not in request.files:
        return jsonify({'error': 'لم يتم اختيار أي ملف'}), 400
    file = request.files['file']
    filename = file.filename or ''

    if filename.lower().endswith('.doc'):
        return jsonify({'error': 'هذا ملف .doc قديم. افتحه في Word ثم «حفظ باسم» واختر .docx، وضعه هنا من جديد.'}), 400
    if not filename.lower().endswith('.docx'):
        return jsonify({'error': 'الملف يجب أن يكون Word بصيغة .docx.'}), 400

    try:
        data = file.read()
        res = analyse_docx(data)
        
        n_p = sum(1 for b in res['blocks'] if b['t'] == 'p' and b['x'].strip())
        n_tbl = sum(1 for b in res['blocks'] if b['t'] == 'tbl')

        return jsonify({
            'filename': filename,
            'paras': n_p,
            'tables': n_tbl,
            'boxed': res['boxed'],
            'box': res['box']
        })
    except Exception as e:
        return jsonify({'error': f'تعذّر فتح الملف: {str(e)}'}), 400


@routes_bp.route('/api/ijraa/convert', methods=['POST'])
@login_required
def ijraa_convert():
    if 'file' not in request.files:
        return jsonify({'error': 'لم يتم اختيار أي ملف'}), 400
    file = request.files['file']
    filename = file.filename or 'deed.docx'
    base_name = os.path.splitext(filename)[0]

    try:
        box_values = {
            'qism': request.form.get('qism', ''),
            'baladiya': request.form.get('baladiya', ''),
            'majmoua': request.form.get('majmoua', ''),
            'hissa': request.form.get('hissa', ''),
            'wilaya': request.form.get('wilaya', ''),
            'idaHajm': request.form.get('idaHajm', ''),
            'idaRaqm': request.form.get('idaRaqm', ''),
            'fi': request.form.get('fi', ''),
            'mujallad': request.form.get('mujallad', ''),
            'mujRaqm': request.form.get('mujRaqm', ''),
            'rasm': request.form.get('rasm', '')
        }

        data = file.read()
        analysed = analyse_docx(data)
        tpl_dir = os.path.join(current_app.config['DOCX_TEMPLATES_DIR'], 'ijraa_ichhar_template')
        
        out_bytes, stats = convert_ijraa(analysed, tpl_dir, box_values)

        # Log activity
        log = ActivityLog(
            user_id=current_user.id,
            username=current_user.username,
            tool_name='إجــراء إشهار عقاري',
            original_filename=filename,
            pages_count=stats['paras'],
            tables_count=stats['tables'],
            warnings_count=0
        )
        db.session.add(log)
        db.session.commit()

        out_name = f"{base_name} - إجراء إشهار عقاري.docx"
        return send_file(
            io.BytesIO(out_bytes),
            as_attachment=True,
            download_name=out_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return jsonify({'error': f'حدث خطأ أثناء الإنشاء: {str(e)}'}), 500


# ----------------- Admin Endpoints -----------------
@routes_bp.route('/admin/users', methods=['GET', 'POST'])
@login_required
def admin_users():
    if not current_user.is_admin:
        flash("هذه الصفحة خاصة بالمسؤول فقط.", 'warn')
        return redirect(url_for('routes.index'))

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'create':
            username = request.form.get('username', '').strip()
            password = request.form.get('password', '').strip()
            role = request.form.get('role', 'staff')
            if User.query.filter_by(username=username).first():
                flash("اسم المستخدم مستعمل بالفعل.", 'warn')
            elif username and password:
                new_u = User(username=username, role=role)
                new_u.set_password(password)
                db.session.add(new_u)
                db.session.commit()
                flash("تم إنشاء المستخدم بنجاح.", 'ok')
        elif action == 'toggle':
            uid = request.form.get('user_id')
            u = User.query.get(uid)
            if u and u.id != current_user.id:
                u.is_active_account = not u.is_active_account
                db.session.commit()
                flash("تم تحديث حالة الحساب.", 'ok')
        elif action == 'reset':
            uid = request.form.get('user_id')
            new_pass = request.form.get('new_password', '').strip()
            u = User.query.get(uid)
            if u and new_pass:
                u.set_password(new_pass)
                db.session.commit()
                flash("تم تغيير كلمة المرور بنجاح.", 'ok')

    users = User.query.order_by(User.created_at.desc()).all()
    return render_template('admin_users.html', users=users)


@routes_bp.route('/admin/logs')
@login_required
def admin_logs():
    if not current_user.is_admin:
        flash("هذه الصفحة خاصة بالمسؤول فقط.", 'warn')
        return redirect(url_for('routes.index'))

    logs = ActivityLog.query.order_by(ActivityLog.timestamp.desc()).limit(200).all()
    return render_template('admin_logs.html', logs=logs)
