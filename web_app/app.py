"""Bulk Email Pro v3.5 — Flask + SocketIO Web Application with campaign persistence"""
import os, sys, json, uuid, hashlib, threading, time
from pathlib import Path
from datetime import datetime
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file
from flask_socketio import SocketIO, emit
from werkzeug.utils import secure_filename
from config import Config

# Add parent dir to import shared modules
sys.path.insert(0, str(Path(__file__).parent.parent / 'desktop_app'))
from account_manager import AccountManager
from excel_processor import ExcelProcessor
from smtp_engine import SMTPEngine

app = Flask(__name__)
app.config.from_object(Config)
Config.init_dirs()
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

# ─── In-memory session stores ─────────────────────────────────
user_processors = {}   # user_id → ExcelProcessor
user_campaigns = {}    # user_id → campaign state dict
user_engines = {}      # user_id → SMTPEngine

def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def get_user_dir(user_id):
    d = Config.USERS_DIR / user_id
    d.mkdir(parents=True, exist_ok=True)
    return d

def get_am(user_id):
    return AccountManager(str(get_user_dir(user_id)))

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

# ─── AUTH ──────────────────────────────────────────────────────
@app.route('/')
def index():
    return redirect(url_for('dashboard') if 'user_id' in session else url_for('login'))

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','')
        if not username or not password:
            flash('Fill all fields','danger'); return render_template('login.html', mode='login')
        users_file = Config.DATA_DIR / 'users.json'
        users = {}
        if users_file.exists():
            with open(users_file) as f: users = json.load(f)
        pw_hash = hash_password(password)
        if username in users:
            if users[username]['password'] == pw_hash:
                session['user_id'] = users[username]['id']
                session['username'] = username
                return redirect(url_for('dashboard'))
            flash('Invalid password','danger')
        else:
            flash('User not found. Register first.','warning')
        return render_template('login.html', mode='login')
    return render_template('login.html', mode='login')

@app.route('/register', methods=['GET','POST'])
def register():
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','')
        if not username or not password:
            flash('Fill all fields','danger'); return render_template('login.html', mode='register')
        users_file = Config.DATA_DIR / 'users.json'
        users = {}
        if users_file.exists():
            with open(users_file) as f: users = json.load(f)
        if username in users:
            flash('Username taken','danger'); return render_template('login.html', mode='register')
        uid = str(uuid.uuid4())[:8]
        users[username] = {'id': uid, 'password': hash_password(password), 'created': datetime.now().isoformat()}
        with open(users_file,'w') as f: json.dump(users, f, indent=2)
        get_user_dir(uid)
        session['user_id'] = uid; session['username'] = username
        flash('Account created!','success')
        return redirect(url_for('dashboard'))
    return render_template('login.html', mode='register')

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

# ─── DASHBOARD ─────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    am = get_am(session['user_id'])
    stats = am.get_stats()
    return render_template('dashboard.html', stats=stats, username=session.get('username'))

# ─── ACCOUNTS ─────────────────────────────────────────────────
@app.route('/accounts')
@login_required
def accounts():
    am = get_am(session['user_id'])
    return render_template('accounts.html', accounts=am.get_all_accounts(), stats=am.get_stats())

@app.route('/accounts/add', methods=['POST'])
@login_required
def add_account():
    am = get_am(session['user_id'])
    try:
        am.add_account(
            nickname=request.form['nickname'], email=request.form['email'],
            smtp_host=request.form['smtp_host'], smtp_port=int(request.form['smtp_port']),
            smtp_security=request.form['smtp_security'], password=request.form['password'],
            daily_limit=int(request.form.get('daily_limit', 500)))
        flash('Account added!','success')
    except Exception as e:
        flash(str(e),'danger')
    return redirect(url_for('accounts'))

@app.route('/accounts/<aid>/test', methods=['POST'])
@login_required
def test_account(aid):
    am = get_am(session['user_id'])
    result = am.test_account(aid)
    return jsonify(result)

@app.route('/accounts/<aid>/delete', methods=['POST'])
@login_required
def delete_account(aid):
    am = get_am(session['user_id'])
    am.delete_account(aid)
    flash('Account deleted','info')
    return redirect(url_for('accounts'))

@app.route('/accounts/test-all', methods=['POST'])
@login_required
def test_all():
    am = get_am(session['user_id'])
    results = am.test_all_accounts()
    return jsonify(results)

# ─── CAMPAIGN ─────────────────────────────────────────────────
@app.route('/campaign/new')
@login_required
def campaign_new():
    am = get_am(session['user_id'])
    return render_template('campaign_new.html', accounts=am.get_all_accounts())

@app.route('/campaign/upload-excel', methods=['POST'])
@login_required
def upload_excel():
    uid = session['user_id']
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "No file uploaded"})
    f = request.files['file']
    if not f.filename:
        return jsonify({"success": False, "error": "No file selected"})
    ext = f.filename.rsplit('.',1)[-1].lower()
    if ext not in Config.ALLOWED_EXTENSIONS:
        return jsonify({"success": False, "error": f"Invalid file type: .{ext}"})
    fname = secure_filename(f.filename)
    save_path = get_user_dir(uid) / 'uploads'
    save_path.mkdir(exist_ok=True)
    fpath = save_path / fname
    f.save(str(fpath))
    ep = ExcelProcessor()
    result = ep.load(str(fpath))
    if result["success"]:
        user_processors[uid] = ep
        detected = ep.auto_detect_email_column()
        result["detected_column"] = detected
        result["columns"] = ep.get_columns()
        if detected:
            stats = ep.validate_and_load(detected)
            result["email_stats"] = stats
            result["personalization_vars"] = ep.get_personalization_vars()
            result["preview"] = ep.get_preview(20)
    return jsonify(result)

@app.route('/campaign/validate-emails', methods=['POST'])
@login_required
def validate_emails():
    uid = session['user_id']
    ep = user_processors.get(uid)
    if not ep: return jsonify({"error": "Upload a file first"})
    col = request.json.get('column')
    stats = ep.validate_and_load(col)
    return jsonify({"stats": stats, "vars": ep.get_personalization_vars(), "preview": ep.get_preview(20)})

@app.route('/campaign/upload-attachments', methods=['POST'])
@login_required
def upload_attachments():
    uid = session['user_id']
    if 'attachments' not in request.files:
        return jsonify({"success": False, "error": "No files uploaded"})
    
    files = request.files.getlist('attachments')
    saved_paths = []
    
    save_dir = get_user_dir(uid) / 'attachments'
    save_dir.mkdir(exist_ok=True)
    
    for f in files:
        if f.filename:
            fname = secure_filename(f.filename)
            fpath = save_dir / fname
            f.save(str(fpath))
            saved_paths.append(str(fpath))
            
    return jsonify({"success": True, "paths": saved_paths})

@app.route('/campaign/launch', methods=['POST'])
@login_required
def launch_campaign():
    uid = session['user_id']
    ep = user_processors.get(uid)
    if not ep or not ep.valid_emails:
        return jsonify({"success": False, "error": "No valid recipients"})
    data = request.json
    campaign_id = str(uuid.uuid4())[:8]
    am = get_am(uid)
    engine = SMTPEngine()
    user_engines[uid] = engine
    user_campaigns[uid] = {
        "id": campaign_id, "status": "running", "is_paused": False,
        "total": len(ep.valid_emails), "sent": 0, "failed": 0, "results": []
    }
    def run():
        camp = user_campaigns[uid]
        def progress_cb(idx, total, result):
            camp["sent"] = sum(1 for r in camp["results"] if r.get("status")=="sent")
            camp["failed"] = sum(1 for r in camp["results"] if r.get("status")=="failed")
            if result.get("status") in ("sent","failed"):
                camp["results"].append(result)
            socketio.emit('progress', {
                "index": idx, "total": total, "result": result,
                "sent": camp["sent"], "failed": camp["failed"]
            }, namespace='/', room=uid)
        results = engine.send_bulk(
            email_list=ep.valid_emails, account_manager=am,
            subject_template=data.get('subject',''), body_template=data.get('body',''),
            is_html=data.get('is_html',True), delay=float(data.get('delay',1.5)),
            use_rotation=data.get('rotation',False), attachment_paths=data.get('attachment_paths', []),
            progress_callback=progress_cb,
            stop_flag=lambda: camp["status"]=="stopped",
            pause_flag=lambda: camp["is_paused"])
        camp["status"] = "complete"
        camp["results"] = results
        # Save campaign to history
        save_campaign_history(uid, camp)
        socketio.emit('complete', {"sent": camp["sent"], "failed": camp["failed"],
                                    "total": len(results)}, namespace='/', room=uid)
    threading.Thread(target=run, daemon=True).start()
    return jsonify({"success": True, "campaign_id": campaign_id})

@app.route('/campaign/pause', methods=['POST'])
@login_required
def pause_campaign():
    uid = session['user_id']
    camp = user_campaigns.get(uid)
    if camp: camp["is_paused"] = not camp["is_paused"]
    return jsonify({"paused": camp["is_paused"] if camp else False})

@app.route('/campaign/stop', methods=['POST'])
@login_required
def stop_campaign():
    uid = session['user_id']
    camp = user_campaigns.get(uid)
    if camp: camp["status"] = "stopped"
    return jsonify({"stopped": True})

@app.route('/campaign/live')
@login_required
def campaign_live():
    uid = session['user_id']
    camp = user_campaigns.get(uid, {})
    return render_template('campaign_live.html', campaign=camp)

# ─── REPORTS ──────────────────────────────────────────────────
@app.route('/reports')
@login_required
def reports():
    uid = session['user_id']
    camp = user_campaigns.get(uid, {})
    return render_template('report.html', campaign=camp, results=camp.get('results',[]))

@app.route('/reports/export')
@login_required
def export_report():
    uid = session['user_id']
    camp = user_campaigns.get(uid,{})
    results = camp.get('results',[])
    filt = request.args.get('filter','all')
    if filt == 'sent': results = [r for r in results if r.get('status')=='sent']
    elif filt == 'failed': results = [r for r in results if r.get('status')=='failed']
    import csv, io
    output = io.StringIO()
    w = csv.DictWriter(output, fieldnames=['_email','status','account_used','timestamp','error'])
    w.writeheader(); w.writerows(results)
    from flask import Response
    return Response(output.getvalue(), mimetype='text/csv',
                   headers={'Content-Disposition': f'attachment;filename=report_{filt}.csv'})

@socketio.on('connect')
def handle_connect():
    if 'user_id' in session:
        from flask_socketio import join_room
        join_room(session['user_id'])

def save_campaign_history(user_id, campaign):
    """Save campaign results to persistent JSON history."""
    hist_path = get_user_dir(user_id) / 'campaign_history.json'
    history = []
    if hist_path.exists():
        try:
            with open(hist_path) as f: history = json.load(f)
        except: pass
    sent = sum(1 for r in campaign.get('results',[]) if r.get('status')=='sent')
    failed = sum(1 for r in campaign.get('results',[]) if r.get('status')=='failed')
    total = sent + failed
    entry = {
        'id': campaign.get('id',''),
        'date': datetime.now().isoformat(),
        'total': total,
        'sent': sent,
        'failed': failed,
        'rate': round(sent/max(total,1)*100,1),
    }
    history.append(entry)
    with open(hist_path,'w') as f: json.dump(history, f, indent=2)

@app.route('/campaign/history')
@login_required
def campaign_history():
    uid = session['user_id']
    hist_path = get_user_dir(uid) / 'campaign_history.json'
    history = []
    if hist_path.exists():
        try:
            with open(hist_path) as f: history = json.load(f)
        except: pass
    return jsonify(history)

@app.route('/accounts/export')
@login_required
def export_accounts():
    uid = session['user_id']
    am = get_am(uid)
    import tempfile
    fp = tempfile.NamedTemporaryFile(delete=False, suffix='.json', mode='w')
    am.export_accounts(fp.name)
    return send_file(fp.name, as_attachment=True, download_name='accounts_backup.json')

@app.route('/accounts/import', methods=['POST'])
@login_required
def import_accounts():
    uid = session['user_id']
    if 'file' not in request.files: return jsonify({"error": "No file"})
    f = request.files['file']
    save_path = get_user_dir(uid) / 'import_accounts.json'
    f.save(str(save_path))
    am = get_am(uid)
    result = am.import_accounts(str(save_path))
    return jsonify(result)

if __name__ == '__main__':
    socketio.run(app, host='0.0.0.0', port=5000, debug=True)
