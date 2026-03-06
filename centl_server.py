import sqlite3
import json
import uuid
import os
import configparser
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, redirect, url_for, session
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_public_key
import base64
from functools import wraps

version = '2.5.312'

# 配置文件路径
CONFIG_FILE = 'server_config.ini'

# 默认配置
DEFAULT_CONFIG = {
    'Server': {
        'host': '0.0.0.0',
        'port': '8393',
        'debug': 'False',
        'server_name': '默认控制台',
        'admin_password': 'admin123'
    }
}

def load_config():
    """加载配置文件，如果不存在则创建默认配置"""
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE, encoding='utf-8')
    else:
        # 创建默认配置文件
        for section, options in DEFAULT_CONFIG.items():
            config[section] = options
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            config.write(f)
        print(f"配置文件 {CONFIG_FILE} 不存在，已创建默认配置")
    return config

def get_config_value(section, key, fallback=None):
    """获取配置值"""
    config = load_config()
    return config.get(section, key, fallback=fallback)

# 加载配置
config = load_config()

app = Flask(__name__)
app.secret_key = get_config_value('Server', 'secret_key', 'your-secret-key-change-in-production')

# 从配置文件读取服务器设置
SERVER_PASSWORD = get_config_value('Server', 'admin_password', 'admin123')
SERVER_NAME = get_config_value('Server', 'server_name', '默认控制台')
SERVER_HOST = get_config_value('Server', 'host', '0.0.0.0')
SERVER_PORT = int(get_config_value('Server', 'port', '8393'))
SERVER_DEBUG = get_config_value('Server', 'debug', 'False').lower() in ('true', 'yes', '1')

# 密码验证装饰器（支持header或JSON中的password字段）
def require_password(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # 从header获取
        pwd_header = request.headers.get('X-Server-Password')
        # 从JSON获取（如果有）
        pwd_json = request.json.get('password') if request.is_json else None
        password = pwd_header or pwd_json
        if not password or password != SERVER_PASSWORD:
            return jsonify({'error': 'invalid password'}), 403
        return f(*args, **kwargs)
    return decorated_function

@app.before_request
def log_request_info():
    print(f"请求: {request.method} {request.path}")

# 基础模板（含激活页面的链接）
BASE_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{{ server_name }} - 打卡中央控制平台</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background-color: #f3f3f3; margin: 0; padding: 20px; }
        .container { max-width: 1200px; margin: auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        h1 { color: #333; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #60cdff; color: black; }
        tr:hover { background-color: #f5f5f5; }
        .btn { display: inline-block; padding: 8px 16px; background-color: #60cdff; color: black; text-decoration: none; border-radius: 4px; margin-right: 5px; cursor: pointer; }
        .btn:hover { background-color: #3da9f5; }
        .btn-danger { background-color: #f44336; color: white; }
        .btn-danger:hover { background-color: #d32f2f; }
        .status-online { color: green; font-weight: bold; }
        .status-offline { color: red; }
        .grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 10px; margin-top: 20px; }
        .grid-item { background-color: white; border: 1px solid #ccc; padding: 10px; text-align: center; border-radius: 4px; cursor: pointer; }
        .grid-item.punched { background-color: #60cdff; }
        .grid-item:hover { background-color: #e0e0e0; }
        .modal { display: none; position: fixed; z-index: 1; left: 0; top: 0; width: 100%; height: 100%; overflow: auto; background-color: rgba(0,0,0,0.4); }
        .modal-content { background-color: #fefefe; margin: 15% auto; padding: 20px; border: 1px solid #888; width: 400px; border-radius: 8px; }
        .close { color: #aaa; float: right; font-size: 28px; font-weight: bold; cursor: pointer; }
        .close:hover { color: black; }
        input, select { width: 100%; padding: 8px; margin: 5px 0; border: 1px solid #ccc; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>{{ server_name }} - 打卡中央控制平台</h1>
        <div>
            <a href="/" class="btn">主页</a>
            {% if session.get('admin') %}
                <a href="/logout" class="btn">登出</a>
                <a href="/activate" class="btn">激活设置</a>
            {% else %}
                <a href="/login" class="btn">管理员登录</a>
            {% endif %}
        </div>
        <hr>
        {{ content|safe }}
    </div>

    <!-- 通用模态框（用于编辑配置、清除数据确认等） -->
    <div id="modal" class="modal">
        <div class="modal-content">
            <span class="close">&times;</span>
            <div id="modal-body"></div>
        </div>
    </div>

    <!-- 打卡/取消打卡模态框（单独保留） -->
    <div id="punchModal" class="modal">
        <div class="modal-content">
            <span class="close">&times;</span>
            <h3>网页操作</h3>
            <p>机器: <span id="modalMachine"></span></p>
            <p>学生: <span id="modalStudent"></span></p>
            <p>当前状态: <span id="modalStatus"></span></p>
            <label>管理员密码: <input type="password" id="adminPwd"></label>
            <br><br>
            <button id="confirmPunch" class="btn" style="display:none;">确认打卡</button>
            <button id="confirmCancel" class="btn btn-danger" style="display:none;">确认取消打卡</button>
        </div>
    </div>

    <script>
        var modal = document.getElementById('modal');
        var span = document.getElementsByClassName("close")[0];
        span.onclick = function() { modal.style.display = "none"; }
        window.onclick = function(event) {
            if (event.target == modal) { modal.style.display = "none"; }
        }

        // 打卡模态框控制
        var punchModal = document.getElementById('punchModal');
        var punchSpan = punchModal.getElementsByClassName("close")[0];
        punchSpan.onclick = function() { punchModal.style.display = "none"; }
        window.onclick = function(event) {
            if (event.target == punchModal) { punchModal.style.display = "none"; }
        }

        var confirmPunchBtn = document.getElementById('confirmPunch');
        var confirmCancelBtn = document.getElementById('confirmCancel');
        var currentMachine = '';
        var currentStudent = '';
        var currentPunched = false;

        function openPunchModal(machine, student, punched) {
            currentMachine = machine;
            currentStudent = student;
            currentPunched = punched;
            document.getElementById('modalMachine').innerText = machine;
            document.getElementById('modalStudent').innerText = student;
            document.getElementById('modalStatus').innerText = punched ? '已打卡' : '未打卡';
            if (punched) {
                confirmPunchBtn.style.display = 'none';
                confirmCancelBtn.style.display = 'inline-block';
            } else {
                confirmPunchBtn.style.display = 'inline-block';
                confirmCancelBtn.style.display = 'none';
            }
            punchModal.style.display = "block";
        }

        confirmPunchBtn.onclick = function() {
            var pwd = document.getElementById('adminPwd').value;
            fetch('/api/web_punch', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    machine_uuid: currentMachine,
                    student_name: currentStudent,
                    password: pwd
                })
            })
            .then(response => response.json())
            .then(data => {
                if (data.status === 'ok') {
                    alert('打卡成功！');
                    location.reload();
                } else {
                    alert('打卡失败：' + data.error);
                }
                punchModal.style.display = "none";
            })
            .catch(err => {
                alert('请求失败：' + err);
                punchModal.style.display = "none";
            });
        }

        confirmCancelBtn.onclick = function() {
            var pwd = document.getElementById('adminPwd').value;
            fetch('/api/web_cancel_punch', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    machine_uuid: currentMachine,
                    student_name: currentStudent,
                    password: pwd
                })
            })
            .then(response => response.json())
            .then(data => {
                if (data.status === 'ok') {
                    alert('取消打卡成功！');
                    location.reload();
                } else {
                    alert('取消打卡失败：' + data.error);
                }
                punchModal.style.display = "none";
            })
            .catch(err => {
                alert('请求失败：' + err);
                punchModal.style.display = "none";
            });
        }

        // 编辑配置模态框
        function openEditConfigModal(machineUuid, currentConfig) {
            var html = `
                <h3>编辑机器配置</h3>
                <form id="editConfigForm">
                    <label>学校: <input type="text" name="school" value="${currentConfig.school || ''}"></label>
                    <label>年级: <input type="text" name="nj" value="${currentConfig.nj || ''}"></label>
                    <label>班级: <input type="text" name="class_id" value="${currentConfig.class_id || ''}"></label>
                    <label>科目: <input type="text" name="km" value="${currentConfig.km || ''}"></label>
                    <label>行数: <input type="number" name="z" value="${currentConfig.z || 6}"></label>
                    <label>列数: <input type="number" name="l" value="${currentConfig.l || 6}"></label>
                    <label>管理员密码: <input type="password" name="password" required></label>
                    <br>
                    <button type="submit" class="btn">保存</button>
                </form>
            `;
            document.getElementById('modal-body').innerHTML = html;
            modal.style.display = "block";

            document.getElementById('editConfigForm').onsubmit = function(e) {
                e.preventDefault();
                var formData = new FormData(e.target);
                var data = {
                    machine_uuid: machineUuid,
                    config: {
                        school: formData.get('school'),
                        nj: formData.get('nj'),
                        class_id: formData.get('class_id'),
                        km: formData.get('km'),
                        z: parseInt(formData.get('z')),
                        l: parseInt(formData.get('l'))
                    },
                    password: formData.get('password')
                };
                fetch('/api/update_machine_config', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data)
                })
                .then(response => response.json())
                .then(data => {
                    if (data.status === 'ok') {
                        alert('配置更新成功');
                        location.reload();
                    } else {
                        alert('更新失败：' + data.error);
                    }
                    modal.style.display = "none";
                })
                .catch(err => {
                    alert('请求失败：' + err);
                    modal.style.display = "none";
                });
            };
        }

        // 清除数据模态框
        function openClearDataModal(machineUuid) {
            var html = `
                <h3>清除打卡数据</h3>
                <p>确定要清除该机器的所有打卡记录吗？此操作不可撤销！</p>
                <form id="clearDataForm">
                    <label>管理员密码: <input type="password" name="password" required></label>
                    <br>
                    <button type="submit" class="btn btn-danger">确认清除</button>
                    <button type="button" class="btn" onclick="modal.style.display='none'">取消</button>
                </form>
            `;
            document.getElementById('modal-body').innerHTML = html;
            modal.style.display = "block";

            document.getElementById('clearDataForm').onsubmit = function(e) {
                e.preventDefault();
                var pwd = e.target.password.value;
                fetch('/api/clear_attendance', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        machine_uuid: machineUuid,
                        password: pwd
                    })
                })
                .then(response => response.json())
                .then(data => {
                    if (data.status === 'ok') {
                        alert('清除成功');
                        location.reload();
                    } else {
                        alert('清除失败：' + data.error);
                    }
                    modal.style.display = "none";
                })
                .catch(err => {
                    alert('请求失败：' + err);
                    modal.style.display = "none";
                });
            };
        }
    </script>
</body>
</html>
'''

def render_page(content):
    return render_template_string(BASE_TEMPLATE, content=content, session=session, server_name=SERVER_NAME)

# 数据库初始化
def init_db():
    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS machines
                 (uuid TEXT PRIMARY KEY,
                  name TEXT,
                  public_key TEXT,
                  last_seen TIMESTAMP,
                  config TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS attendance
                 (machine_uuid TEXT,
                  data TEXT,
                  updated_at TIMESTAMP,
                  FOREIGN KEY(machine_uuid) REFERENCES machines(uuid))''')
    conn.commit()
    conn.close()

init_db()

# ---------- 辅助函数 ----------
def verify_signature(public_key_pem, message, signature_b64):
    try:
        public_key = load_pem_public_key(public_key_pem.encode())
        signature = base64.b64decode(signature_b64)
        public_key.verify(
            signature,
            message.encode(),
            padding.PKCS1v15(),
            hashes.SHA256()
        )
        return True
    except Exception as e:
        print(f"签名验证失败: {e}")
        return False

def get_machine_public_key(uuid):
    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT public_key FROM machines WHERE uuid=?", (uuid,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else None

# ---------- 登录（管理员密码验证）----------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        pwd = request.form.get('password')
        if pwd == SERVER_PASSWORD:
            session['admin'] = True
            return redirect(url_for('index'))
        else:
            return render_page('<h2>密码错误</h2><a href="/login">返回</a>')
    form = '''
    <h2>管理员登录</h2>
    <form method="post">
        <label>密码: <input type="password" name="password"></label>
        <button type="submit" class="btn">登录</button>
    </form>
    '''
    return render_page(form)

@app.route('/logout')
def logout():
    session.pop('admin', None)
    return redirect(url_for('index'))

# ---------- 激活页面（设置服务器密码和控制台名称）----------
@app.route('/activate', methods=['GET', 'POST'])
def activate():
    global SERVER_PASSWORD, SERVER_NAME
    if request.method == 'POST':
        new_pwd = request.form.get('password')
        new_name = request.form.get('server_name')
        new_debug = request.form.get('debug_mode')
        new_host = request.form.get('host')
        new_port = request.form.get('port')

        # 更新配置文件
        config = load_config()
        if new_pwd:
            config['Server']['admin_password'] = new_pwd
            SERVER_PASSWORD = new_pwd
        if new_name:
            config['Server']['server_name'] = new_name
            SERVER_NAME = new_name
        if new_debug:
            config['Server']['debug'] = new_debug
        if new_host:
            config['Server']['host'] = new_host
        if new_port:
            config['Server']['port'] = new_port

        # 保存配置文件
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            config.write(f)

        return redirect(url_for('index'))

    # 读取当前配置
    config = load_config()
    current_debug = config.get('Server', 'debug', fallback='False')
    current_host = config.get('Server', 'host', fallback='0.0.0.0')
    current_port = config.get('Server', 'port', fallback='8393')

    form = f'''
    <h2>服务器激活设置</h2>
    <form method="post">
        <label>服务器密码: <input type="password" name="password" value="{SERVER_PASSWORD}"></label><br><br>
        <label>控制台名称: <input type="text" name="server_name" value="{SERVER_NAME}"></label><br><br>
        <label>服务器IP: <input type="text" name="host" value="{current_host}" placeholder="0.0.0.0"></label><br><br>
        <label>端口号: <input type="number" name="port" value="{current_port}" min="1" max="65535"></label><br><br>
        <label>Flask调试模式:
            <select name="debug_mode">
                <option value="True" {"selected" if current_debug.lower() == 'true' else ""}>开启</option>
                <option value="False" {"selected" if current_debug.lower() != 'true' else ""}>关闭</option>
            </select>
        </label><br><br>
        <button type="submit" class="btn">保存配置</button>
    </form>
    <p style="color: #666; font-size: 12px;">注意：修改IP和端口后需要重启服务器才能生效</p>
    '''
    return render_page(form)

# ---------- 主页 ----------
@app.route('/')
def index():
    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT uuid, name, last_seen FROM machines ORDER BY last_seen DESC")
    machines = c.fetchall()
    conn.close()
    now = datetime.now()
    machine_list = []
    for uuid, name, last_seen in machines:
        last = datetime.fromisoformat(last_seen) if last_seen else None
        online = last and (now - last).total_seconds() < 300
        machine_list.append((uuid, name or '未命名', online, last_seen))

    rows = ''.join(
        f'<tr><td>{uuid[:8]}...</td><td>{name}</td><td class="status-{"online" if online else "offline"}">{"在线" if online else "离线"}</td><td>{last_seen}</td><td><a href="/machine/{uuid}" class="btn">查看</a></td></tr>'
        for uuid, name, online, last_seen in machine_list
    )
    table = f'<table><tr><th>UUID</th><th>名称</th><th>状态</th><th>最后在线</th><th>操作</th></tr>{rows}</table>'
    return render_page(f'<h2>已注册机器</h2>{table}')

# ---------- 机器详情页 ----------
@app.route('/machine/<uuid>')
def machine_detail(uuid):
    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT name, config FROM machines WHERE uuid=?", (uuid,))
    machine = c.fetchone()
    if not machine:
        return render_page('<h2>机器不存在</h2>')
    name, config_json = machine
    config = json.loads(config_json) if config_json else {}

    c.execute("SELECT data, updated_at FROM attendance WHERE machine_uuid=? ORDER BY updated_at DESC LIMIT 1", (uuid,))
    row = c.fetchone()
    conn.close()

    attendance_data = {}
    update_time = "从未同步"
    if row:
        attendance_data = json.loads(row[0])
        update_time = row[1]

    students = list(attendance_data.keys())
    punched = [(n, d['first_time']) for n, d in attendance_data.items() if d.get('first_time')]
    punched.sort(key=lambda x: x[1])

    rank_rows = ''.join(
        f'<tr><td>{i}</td><td>{n}</td><td>{t}</td></tr>'
        for i, (n, t) in enumerate(punched, 1)
    )
    rank_table = f'<table><tr><th>排名</th><th>姓名</th><th>打卡时间</th></tr>{rank_rows}</table>' if rank_rows else '<p>暂无打卡记录</p>'

    # 生成学生网格
    grid_items = ''
    for s in students:
        punched_class = 'punched' if attendance_data.get(s, {}).get('first_time') else ''
        punched_status = 'true' if attendance_data.get(s, {}).get('first_time') else 'false'
        if session.get('admin'):
            grid_items += f'<div class="grid-item {punched_class}" onclick="openPunchModal(\'{uuid}\', \'{s}\', {punched_status})">{s}</div>'
        else:
            grid_items += f'<div class="grid-item {punched_class}">{s}</div>'
    grid = f'<div class="grid">{grid_items}</div>'

    # 配置展示（美化）
    config_display = f'''
        <li>学校：{config.get('school', '未设置')}</li>
        <li>年级：{config.get('nj', '未设置')}</li>
        <li>班级：{config.get('class_id', '未设置')}</li>
        <li>科目：{config.get('km', '未设置')}</li>
        <li>网格行数：{config.get('z', 6)}</li>
        <li>网格列数：{config.get('l', 6)}</li>
    '''

    admin_buttons = ''
    if session.get('admin'):
        admin_buttons = f'''
            <button class="btn" onclick="openEditConfigModal('{uuid}', {json.dumps(config)})">编辑配置</button>
            <button class="btn btn-danger" onclick="openClearDataModal('{uuid}')">清除打卡数据</button>
        '''

    content = f'''
    <h2>机器详情 - {name or "未命名"}</h2>
    <p>UUID: {uuid}</p>
    <p>最后数据同步: {update_time}</p>
    <h3>当前配置</h3>
    <ul>{config_display}</ul>
    {admin_buttons}
    <h3>打卡排名 (最早打卡)</h3>
    {rank_table}
    <h3>学生打卡状态</h3>
    {grid}
    <a href="/" class="btn">返回</a>
    '''
    return render_page(content)

# ---------- API 端点（需验证密码）----------

@app.route('/api/register', methods=['POST'])
@require_password
def api_register():
    data = request.json
    public_key = data.get('public_key')
    name = data.get('name', '')
    if not public_key:
        return jsonify({'error': 'public_key required'}), 400

    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT uuid FROM machines WHERE public_key=?", (public_key,))
    existing = c.fetchone()
    if existing:
        machine_uuid = existing[0]
        c.execute("UPDATE machines SET last_seen=? WHERE uuid=?", (datetime.now().isoformat(), machine_uuid))
        conn.commit()
        conn.close()
        return jsonify({'uuid': machine_uuid, 'existing': True})
    else:
        machine_uuid = str(uuid.uuid4())
        c.execute("INSERT INTO machines (uuid, name, public_key, last_seen, config) VALUES (?,?,?,?,?)",
                  (machine_uuid, name, public_key, datetime.now().isoformat(), '{}'))
        conn.commit()
        conn.close()
        return jsonify({'uuid': machine_uuid, 'existing': False})

@app.route('/api/sync_data', methods=['POST'])
@require_password
def api_sync_data():
    data = request.json
    machine_uuid = data.get('uuid')
    signature = data.get('signature')
    payload = data.get('data')
    if not all([machine_uuid, signature, payload]):
        return jsonify({'error': 'missing fields'}), 400

    public_key_pem = get_machine_public_key(machine_uuid)
    if not public_key_pem:
        return jsonify({'error': 'unknown machine'}), 403

    if not verify_signature(public_key_pem, payload, signature):
        return jsonify({'error': 'invalid signature'}), 403

    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("INSERT INTO attendance (machine_uuid, data, updated_at) VALUES (?,?,?)",
              (machine_uuid, payload, datetime.now().isoformat()))
    c.execute("UPDATE machines SET last_seen=? WHERE uuid=?", (datetime.now().isoformat(), machine_uuid))
    conn.commit()
    conn.close()
    print(f"数据同步成功: {machine_uuid}")
    return jsonify({'status': 'ok'})

@app.route('/api/load_data', methods=['POST'])
@require_password
def api_load_data():
    data = request.json
    machine_uuid = data.get('uuid')
    signature = data.get('signature')
    challenge = data.get('challenge', '')
    if not all([machine_uuid, signature, challenge]):
        return jsonify({'error': 'missing fields'}), 400

    public_key_pem = get_machine_public_key(machine_uuid)
    if not public_key_pem:
        return jsonify({'error': 'unknown machine'}), 403

    if not verify_signature(public_key_pem, challenge, signature):
        return jsonify({'error': 'invalid signature'}), 403

    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT data FROM attendance WHERE machine_uuid=? ORDER BY updated_at DESC LIMIT 1", (machine_uuid,))
    row = c.fetchone()
    c.execute("UPDATE machines SET last_seen=? WHERE uuid=?", (datetime.now().isoformat(), machine_uuid))
    conn.commit()
    conn.close()
    if row:
        return jsonify({'data': json.loads(row[0])})
    else:
        return jsonify({'data': {}})

@app.route('/api/status', methods=['GET'])
def api_status():
    # 公开接口，无需密码，但可加可选密码（为保持兼容，不强制）
    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT uuid, name, last_seen FROM machines")
    rows = c.fetchall()
    machines = []
    now = datetime.now()
    for uuid, name, last_seen in rows:
        last = datetime.fromisoformat(last_seen) if last_seen else None
        online = last and (now - last).total_seconds() < 300
        machines.append({
            'uuid': uuid,
            'name': name,
            'online': online,
            'last_seen': last_seen
        })
    conn.close()
    return jsonify(machines)

@app.route('/api/get_config', methods=['POST'])
@require_password
def api_get_config():
    data = request.json
    machine_uuid = data.get('uuid')
    signature = data.get('signature')
    challenge = data.get('challenge', '')
    if not all([machine_uuid, signature, challenge]):
        return jsonify({'error': 'missing fields'}), 400

    public_key_pem = get_machine_public_key(machine_uuid)
    if not public_key_pem:
        return jsonify({'error': 'unknown machine'}), 403

    if not verify_signature(public_key_pem, challenge, signature):
        return jsonify({'error': 'invalid signature'}), 403

    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT config FROM machines WHERE uuid=?", (machine_uuid,))
    row = c.fetchone()
    c.execute("UPDATE machines SET last_seen=? WHERE uuid=?", (datetime.now().isoformat(), machine_uuid))
    conn.commit()
    conn.close()
    if row and row[0]:
        config = json.loads(row[0])
    else:
        config = {}
    return jsonify({'config': config})

@app.route('/api/update_config', methods=['POST'])
@require_password
def api_update_config():
    data = request.json
    machine_uuid = data.get('uuid')
    signature = data.get('signature')
    config_data = data.get('config')
    if not all([machine_uuid, signature, config_data]):
        return jsonify({'error': 'missing fields'}), 400

    config_str = json.dumps(config_data, ensure_ascii=False) if isinstance(config_data, dict) else config_data

    public_key_pem = get_machine_public_key(machine_uuid)
    if not public_key_pem:
        return jsonify({'error': 'unknown machine'}), 403

    if not verify_signature(public_key_pem, config_str, signature):
        return jsonify({'error': 'invalid signature'}), 403

    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("UPDATE machines SET config=?, last_seen=? WHERE uuid=?",
              (config_str, datetime.now().isoformat(), machine_uuid))
    conn.commit()
    conn.close()
    print(f"配置更新成功: {machine_uuid}")
    return jsonify({'status': 'ok'})

# 新增：网页端更新机器配置（管理员操作）
@app.route('/api/update_machine_config', methods=['POST'])
def api_update_machine_config():
    data = request.json
    machine_uuid = data.get('machine_uuid')
    new_config = data.get('config')
    password = data.get('password')
    if not all([machine_uuid, new_config, password]) or password != SERVER_PASSWORD:
        return jsonify({'error': 'invalid password or missing data'}), 403

    config_str = json.dumps(new_config, ensure_ascii=False)
    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("UPDATE machines SET config=? WHERE uuid=?", (config_str, machine_uuid))
    conn.commit()
    conn.close()
    return jsonify({'status': 'ok'})

# 新增：清除机器所有打卡数据
@app.route('/api/clear_attendance', methods=['POST'])
def api_clear_attendance():
    data = request.json
    machine_uuid = data.get('machine_uuid')
    password = data.get('password')
    if not all([machine_uuid, password]) or password != SERVER_PASSWORD:
        return jsonify({'error': 'invalid password'}), 403

    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    # 删除该机器的所有打卡记录，并插入一条空数据（可选）
    c.execute("DELETE FROM attendance WHERE machine_uuid=?", (machine_uuid,))
    # 插入一条空数据，表示已清除
    empty_data = json.dumps({})
    c.execute("INSERT INTO attendance (machine_uuid, data, updated_at) VALUES (?,?,?)",
              (machine_uuid, empty_data, datetime.now().isoformat()))
    conn.commit()
    conn.close()
    return jsonify({'status': 'ok'})

@app.route('/api/web_punch', methods=['POST'])
def api_web_punch():
    data = request.json
    machine_uuid = data.get('machine_uuid')
    student_name = data.get('student_name')
    password = data.get('password')
    if not all([machine_uuid, student_name, password]) or password != SERVER_PASSWORD:
        return jsonify({'error': 'invalid password or missing data'}), 403

    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT data FROM attendance WHERE machine_uuid=? ORDER BY updated_at DESC LIMIT 1", (machine_uuid,))
    row = c.fetchone()
    if row:
        attendance_data = json.loads(row[0])
    else:
        attendance_data = {}

    if student_name not in attendance_data:
        attendance_data[student_name] = {"count": 0, "first_time": None, "history": []}

    if attendance_data[student_name].get('first_time'):
        conn.close()
        return jsonify({'error': '该学生已经打卡'}), 400

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    attendance_data[student_name]['first_time'] = current_time
    attendance_data[student_name]['count'] = attendance_data[student_name].get('count', 0) + 1
    if 'history' not in attendance_data[student_name]:
        attendance_data[student_name]['history'] = []
    attendance_data[student_name]['history'].append(current_time)

    new_data_str = json.dumps(attendance_data, ensure_ascii=False)
    c.execute("INSERT INTO attendance (machine_uuid, data, updated_at) VALUES (?,?,?)",
              (machine_uuid, new_data_str, datetime.now().isoformat()))
    c.execute("UPDATE machines SET last_seen=? WHERE uuid=?", (datetime.now().isoformat(), machine_uuid))
    conn.commit()
    conn.close()
    print(f"网页打卡成功: {machine_uuid} - {student_name}")
    return jsonify({'status': 'ok'})

@app.route('/api/web_cancel_punch', methods=['POST'])
def api_web_cancel_punch():
    data = request.json
    machine_uuid = data.get('machine_uuid')
    student_name = data.get('student_name')
    password = data.get('password')
    if not all([machine_uuid, student_name, password]) or password != SERVER_PASSWORD:
        return jsonify({'error': 'invalid password or missing data'}), 403

    conn = sqlite3.connect('central.db')
    c = conn.cursor()
    c.execute("SELECT data FROM attendance WHERE machine_uuid=? ORDER BY updated_at DESC LIMIT 1", (machine_uuid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': '该机器无打卡数据'}), 404

    attendance_data = json.loads(row[0])

    if student_name not in attendance_data:
        conn.close()
        return jsonify({'error': '学生不存在'}), 404

    if not attendance_data[student_name].get('first_time'):
        conn.close()
        return jsonify({'error': '该学生未打卡，无法取消'}), 400

    if attendance_data[student_name]['history']:
        removed_time = attendance_data[student_name]['history'].pop()
        if attendance_data[student_name]['first_time'] == removed_time:
            if attendance_data[student_name]['history']:
                attendance_data[student_name]['first_time'] = attendance_data[student_name]['history'][0]
            else:
                attendance_data[student_name]['first_time'] = None
        attendance_data[student_name]['count'] = max(0, attendance_data[student_name].get('count', 0) - 1)
    else:
        attendance_data[student_name]['first_time'] = None
        attendance_data[student_name]['count'] = 0

    new_data_str = json.dumps(attendance_data, ensure_ascii=False)
    c.execute("INSERT INTO attendance (machine_uuid, data, updated_at) VALUES (?,?,?)",
              (machine_uuid, new_data_str, datetime.now().isoformat()))
    c.execute("UPDATE machines SET last_seen=? WHERE uuid=?", (datetime.now().isoformat(), machine_uuid))
    conn.commit()
    conn.close()
    print(f"网页取消打卡成功: {machine_uuid} - {student_name}")
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    print(fr'''
      o__ __o     o                                   o           __o__                                             
     /v     v\   <|>                                 <|>            |                                              
    />       <\  / >                                 / \           / \                                             
  o/             \o__ __o      o__  __o       __o__  \o/  o/       \o/   \o__ __o                                  
 <|               |     v\    /v      |>     />  \    |  /v         |     |     |>                                 
  \\             / \     <\  />      //    o/        / \/>         < >   / \   / \                                 
    \         /  \o/     o/  \o    o/     <|         \o/\o          |    \o/   \o/                                 
     o       o    |     <|    v\  /v __o   \\         |  v\         o     |     |                                  
     <\__ __/>   / \    / \    <\/> __/>    _\o__</  / \  <\      __|>_  / \   / \                                 
                                                                                                                    
                                                                                                                    
                                                                                                                    
                                            o__ __o                                                                
                                           /v     v\                                                               
                                          />       <\                                                              
                                         _\o____          o__  __o   \o__ __o    o      o     o__  __o   \o__ __o
                                              \_\__o__   /v      |>   |     |>  <|>    <|>   /v      |>   |     |> 
                                                    \   />      //   / \   < >  < >    < >  />      //   / \   < > 
                                          \         /   \o    o/     \o/         \o    o/   \o    o/     \o/       
                                           o       o     v\  /v __o   |           v\  /v     v\  /v __o   |        
                                           <\__ __/>      <\/> __/>  / \           <\/>       <\/> __/>  / \       
                                                                                                                   
    @2026 - {str(int(datetime.today().strftime("%Y"))+1)} 刘宇晨 保留全部对该版本的权力                                                                                               
    @2026 - {str(int(datetime.today().strftime("%Y"))+1)} Liu Yuchen reserves all rights to this version.
    |  Version: Server v{version}
    |  Developer: Liu Yuchen
    ''')


    print(f"服务器配置:")
    print(f"  IP: {SERVER_HOST}")
    print(f"  端口: {SERVER_PORT}")
    print(f"  调试模式: {SERVER_DEBUG}")
    print(f"  控制台名称: {SERVER_NAME}")
    print(f"  管理员密码: {SERVER_PASSWORD}")
    print(f"配置文件: {CONFIG_FILE}")
    app.run(host=SERVER_HOST, port=SERVER_PORT, debug=SERVER_DEBUG)
