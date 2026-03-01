# online.py
# 大屏打卡程序（中央服务器客户端）
# 版权所有 (c) 2025 刘宇晨

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import time
import json
import uuid
import base64
import requests
import sys
import webbrowser
import hashlib
from configparser import ConfigParser
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import serialization, hashes
from cryptography.hazmat.primitives.asymmetric import padding

import show_ui
from mode_sever_start import *   # 本地服务器模式（可选）

# 全局变量
version = 'v2.5.312'
directory = os.path.dirname(__file__)
conf = ConfigParser()
conf.read('config.ini', encoding='UTF-8')

# 从配置文件读取
nj = conf['config']['nj']
class_id = conf['config']['class_id']
z = int(conf['config']['z'])
l = int(conf['config']['l'])
km = conf['config']['km']
school = conf['config']['school']
online_mode = conf.getboolean('config', 'online', fallback=False)
bd_online = conf.getboolean('config', 'bd_online', fallback=False)
online_ip = conf.get('config', 'online_ip', fallback='')
server_port = conf.getint('config', 'server_port', fallback=5000)
server_password = conf.get('config', 'server_password', fallback='')
admin_password_hash = conf.get('config', 'admin_password', fallback='')  # 管理员密码哈希

file_path = ""  # 将在 load_student_data 中设置


def check_version():
    webbrowser.open('https://github.com/liuyuchen012/daikai/releases')


def open_url_bz():
    webbrowser.open('https://github.com/liuyuchen012/daikai/tree/' + version)


class AttendanceApp:
    def __init__(self, window):
        # 初始化窗口
        self.window = window
        self.version = version
        self.dark_mode = False          # 暂时只支持浅色模式
        self.maximized = False
        self.prev_geometry = ""

        # 从全局变量初始化

        self.nj = nj
        self.class_id = class_id
        self.z = z
        self.l = l
        self.km = km
        self.school = school
        self.online_mode = online_mode
        self.bd_online = bd_online
        self.online_ip = online_ip
        self.server_port = server_port
        self.server_password = server_password
        self.admin_password_hash = admin_password_hash
        # 密钥与UUID
        self.client_uuid = None
        self.private_key = None
        self.public_key = None
        self.load_or_create_keys()

        # 在线状态变量
        self.server_status = "未连接"
        self.server_last_check = "从未检查"
        self.connection_attempts = 0

        # 尝试加载图标
        try:
            self.window.iconbitmap('icon.ico')
        except:
            pass

        # 初始化数据结构
        self.student_data = {}
        self.load_student_data()

        # 应用主题（设置颜色）
        self.apply_theme()

        # 设置窗口样式（自定义标题栏等）
        show_ui.setup_window_style(self.window)
        self.create_custom_titlebar()
        self.create_menu()

        # 创建主界面
        self.show_main_interface()

        # 启动API服务器（如果配置为服务器）
        if self.bd_online:
            self.start_api_server()

        # 如果配置了在线模式，连接到中央服务器
        if self.online_mode:
            self.register_with_central()
            self.window.after(1000, self.load_data_from_server)   # 首次加载数据
            self.window.after(2000, self.sync_config_to_server)   # 推送本地配置
            self.window.after(60000, self.periodic_load_config)   # 每分钟拉取配置
            self.window.after(30000, self.check_server_status)    # 状态检查
            self.window.after(30000, self.periodic_load_data)     # 每30秒拉取打卡数据

    # ---------- 主题应用 ----------
    def theme(self):
        """应用主题颜色（基于 self.dark_mode）"""
        if True == False:
            # 深色模式（参考 old.py 注释掉的配色）
            self.bg_color = '#2b2b2b'
            self.fg_color = '#a9b7c6'
            self.status_fg_color = '#a9b7c6'
            self.button_bg = '#3c3f41'
            self.button_fg = '#a9b7c6'
            self.button_hover = '#4e5254'
            self.active_button_bg = '#4e5254'
            self.active_button_fg = '#ffffff'
            self.frame_bg = '#3c3f41'
            self.frame_border = '#515151'
            self.tree_bg = '#2b2b2b'
            self.tree_fg = '#a9b7c6'
            self.status_bg = '#3c3f41'
            self.menu_bg = '#3c3f41'
            self.menu_fg = '#a9b7c6'
            self.menu_active_bg = '#4e5254'
            self.menu_active_fg = '#ffffff'
            self.online_color = '#6a8759'
            self.offline_color = '#bc3f3c'
        else:
            # 浅色模式（old.py 默认）
            self.bg_color = '#f0f0f0'
            self.fg_color = '#000000'
            self.status_fg_color = '#333333'
            self.button_bg = '#e0e0e0'
            self.button_fg = '#000000'
            self.button_hover = '#d0d0d0'
            self.active_button_bg = '#4285f4'
            self.active_button_fg = '#ffffff'
            self.frame_bg = '#e8e8e8'
            self.frame_border = '#cccccc'
            self.tree_bg = '#ffffff'
            self.tree_fg = '#000000'
            self.status_bg = '#e0e0e0'
            self.menu_bg = '#e8e8e8'
            self.menu_fg = '#000000'
            self.menu_active_bg = '#d0d0d0'
            self.menu_active_fg = '#000000'
            self.online_color = '#34a853'
            self.offline_color = '#ea4335'

        # 设置窗口背景
        self.window.config(bg=self.bg_color)

        # 更新自定义标题栏（如果存在）- 标题栏现在是白色
        if hasattr(self, 'titlebar'):
            titlebar_bg = '#ffffff'
            titlebar_fg = '#333333'
            self.titlebar.config(bg=titlebar_bg)
            self.title_label.config(bg=titlebar_bg, fg=titlebar_fg)
            for btn in [self.min_btn, self.max_btn, self.close_btn]:
                btn.config(bg=titlebar_bg, fg=titlebar_fg)

        # 菜单已集成到标题栏中，不再需要单独更新

        # 更新在线状态标签
        if hasattr(self, 'online_status_label'):
            if self.server_status == "在线":
                self.online_status_label.config(fg=self.online_color)
            else:
                self.online_status_label.config(fg=self.offline_color)

        # 更新主容器背景
        if hasattr(self, 'main_container'):
            self.main_container.config(bg=self.bg_color)

        # 更新 Treeview 样式
        self.update_treeview_style()
    # ---------- 密码相关 ----------
    def hash_password(self, password):
        return hashlib.sha256(password.encode()).hexdigest()

    def verify_admin_password(self, action_name):
        if not self.admin_password_hash:
            return True
        password = self.ask_password(f"需要管理员权限执行: {action_name}")
        if password and self.hash_password(password) == self.admin_password_hash:
            return True
        else:
            messagebox.showerror("错误", "管理员密码错误")
            return False

    def ask_password(self, title):
        password_window = tk.Toplevel(self.window)
        password_window.title(title)
        password_window.geometry("300x150")
        password_window.resizable(False, False)
        password_window.transient(self.window)
        password_window.grab_set()
        password_window.config(bg=self.bg_color)

        tk.Label(password_window, text="请输入管理员密码:",
                 bg=self.bg_color, fg=self.fg_color, font=("楷体", 10)).pack(pady=10)

        password_var = tk.StringVar()
        password_entry = tk.Entry(password_window, textvariable=password_var, show="*",
                                  bg=self.bg_color, fg=self.fg_color, font=("楷体", 10))
        password_entry.pack(pady=5, padx=20, fill=tk.X)
        password_entry.focus()

        result = [None]

        def on_ok():
            result[0] = password_var.get()
            password_window.destroy()

        def on_cancel():
            result[0] = None
            password_window.destroy()

        button_frame = tk.Frame(password_window, bg=self.bg_color)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="确定", command=on_ok, width=10,
                  bg=self.button_bg, fg=self.button_fg).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="取消", command=on_cancel, width=10,
                  bg=self.button_bg, fg=self.button_fg).pack(side=tk.LEFT, padx=5)

        password_window.wait_window()
        return result[0]

    # ---------- 清空打卡记录 ----------
    def clear_attendance_records(self):
        if not self.verify_admin_password("清空打卡记录"):
            return
        for i in range(3, 0, -1):
            result = messagebox.askyesno("警告",
                                         f"确定要清空所有打卡记录吗？此操作不可恢复！\n还有 {i} 次警告")
            if not result:
                return
        confirm = messagebox.askyesno("最终确认",
                                      "这是最后一次确认！确定要清空所有打卡记录吗？")
        if not confirm:
            return

        for name in self.student_data:
            self.student_data[name]['count'] = 0
            self.student_data[name]['first_time'] = None
            self.student_data[name]['history'] = []

        self.update_ui_from_data()
        self.save_student_data()
        self.status_var.set("所有打卡记录已清空")
        messagebox.showinfo("完成", "所有打卡记录已成功清空")

        if self.online_mode and self.server_status == "在线":
            self.sync_data_to_server()

    # ---------- 管理员设置 ----------
    def show_admin_settings(self):
        show_ui.show_admin_settings(self,tk)

    def save_all_settings(self, settings_window):
        try:
            self.school = self.school_entry.get()
            self.nj = self.nj_entry.get()
            self.class_id = self.class_entry.get()
            self.km = self.km_entry.get()
            self.z = int(self.row_entry.get())
            self.l = int(self.col_entry.get())
            if self.z <= 0 or self.l <= 0:
                raise ValueError("行数和列数必须为正数")

            new_password = self.new_password_entry.get()
            confirm_password = self.confirm_password_entry.get()
            if new_password:
                if new_password != confirm_password:
                    raise ValueError("两次输入的密码不一致")
                self.admin_password_hash = self.hash_password(new_password)

            self.save_config()
            settings_window.destroy()

            # 应用主题并重启界面
            self.apply_theme()
            # 重新创建主界面
            for widget in self.main_container.winfo_children():
                widget.destroy()
            self.show_main_interface()
            messagebox.showinfo("提示", "所有设置已保存并生效，请手动重启程序以完全生效")
        except ValueError as e:
            messagebox.showerror("错误", f"输入无效: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {str(e)}")

    def clear_admin_password(self, settings_window=None):
        confirm = messagebox.askyesno("确认", "确定要清除管理员密码吗？")
        if confirm:
            self.admin_password_hash = ''
            self.save_config()
            if settings_window:
                settings_window.destroy()
            messagebox.showinfo("成功", "管理员密码已清除")

    # ---------- 窗口控制 ----------
    def minimize_window(self):
        self.window.iconify()

    def toggle_maximize(self):
        if self.maximized:
            self.window.geometry(self.prev_geometry)
            self.max_btn.config(text="□")
            self.maximized = False
        else:
            self.prev_geometry = self.window.geometry()
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            self.window.geometry(f"{screen_width}x{screen_height}+0+0")
            self.max_btn.config(text="❐")
            self.maximized = True

    # ---------- UI 辅助方法 ----------
    def apply_theme(self):
        show_ui.theme(self, tk)

    def create_custom_titlebar(self):
        show_ui.create_custom_titlebar(self, tk)

    def create_menu(self):
        # 将外部函数绑定到 self，供菜单使用
        self.open_url_bz = open_url_bz
        self.check_version = check_version
        show_ui.create_menu(self, tk)

    def show_main_interface(self):
        show_ui.create_widgets(self, tk, ttk, self.school, self.nj, self.class_id, self.km, self.z, self.l)

    def show_about(self):
        show_ui.show_about(self)

    def on_button_enter(self, event, button, name):
        show_ui.on_button_enter(self, event, button, name)

    def on_button_leave(self, event, button, name):
        show_ui.on_button_leave(self, event, button, name)

    def show_context_menu(self, event, name):
        show_ui.show_context_menu(self, event, name)

    def cancel_attendance(self, name):
        show_ui.cancel_attendance(self, name)

    def import_data(self):
        show_ui.import_data(self)

    def update_ui_from_data(self):
        show_ui.update_ui_from_data(self)

    # ---------- Treeview 样式更新 ----------
    def update_treeview_style(self):
        """专门更新Treeview的样式"""
        if hasattr(self, 'ranking_tree'):
            style = ttk.Style()
            style.configure("Custom.Treeview",
                            background=self.tree_bg,
                            foreground=self.tree_fg,
                            fieldbackground=self.tree_bg,
                            rowheight=25,
                            font=("楷体", 9))
            style.configure("Custom.Treeview.Heading",
                            background=self.menu_bg,
                            foreground=self.menu_fg,
                            font=("楷体", 9, 'bold'),
                            relief=tk.FLAT)
            style.map("Custom.Treeview",
                      background=[('selected', self.active_button_bg)],
                      foreground=[('selected', self.active_button_fg)])
            self.ranking_tree.configure(style="Custom.Treeview")

    # ---------- 密钥管理 ----------
    def load_or_create_keys(self):
        key_file = 'client_key.pem'
        uuid_file = 'client_uuid.txt'
        if os.path.exists(key_file) and os.path.exists(uuid_file):
            with open(key_file, 'rb') as f:
                self.private_key = serialization.load_pem_private_key(f.read(), password=None)
            with open(uuid_file, 'r') as f:
                self.client_uuid = f.read().strip()
            self.public_key = self.private_key.public_key()
        else:
            self.private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
            self.public_key = self.private_key.public_key()
            with open(key_file, 'wb') as f:
                f.write(self.private_key.private_bytes(
                    encoding=serialization.Encoding.PEM,
                    format=serialization.PrivateFormat.PKCS8,
                    encryption_algorithm=serialization.NoEncryption()
                ))
            self.client_uuid = str(uuid.uuid4())
            with open(uuid_file, 'w') as f:
                f.write(self.client_uuid)

    def get_public_key_pem(self):
        return self.public_key.public_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PublicFormat.SubjectPublicKeyInfo
        ).decode()

    def sign_message(self, message):
        signature = self.private_key.sign(
            message.encode(),
            padding.PKCS1v15(),
            hashes.SHA256()
        )
        return base64.b64encode(signature).decode()

    # ---------- 服务器通信（增加密码header）----------
    def _request_with_auth(self, method, url, **kwargs):
        headers = kwargs.pop('headers', {})
        headers['X-Server-Password'] = self.server_password
        try:
            response = requests.request(method, url, headers=headers, timeout=kwargs.pop('timeout', 5), **kwargs)
            if response.status_code == 403:
                self.status_var.set("服务器密码错误，请检查设置")
                messagebox.showerror("认证失败", "服务器密码错误，请检查远程设置中的密码")
                return None
            return response
        except Exception as e:
            self.status_var.set(f"请求异常: {e}")
            return None

    def register_with_central(self):
        if not self.online_mode or not self.online_ip:
            return
        url = f"http://{self.online_ip}:{self.server_port}/api/register"
        response = self._request_with_auth('POST', url, json={
            'public_key': self.get_public_key_pem(),
            'name': f'{self.school}{self.nj}年{self.class_id}班'
        })
        if response is None:
            return
        if response.status_code == 200:
            data = response.json()
            server_uuid = data.get('uuid')
            if server_uuid:
                if server_uuid != self.client_uuid:
                    self.client_uuid = server_uuid
                    uuid_file = 'client_uuid.txt'
                    with open(uuid_file, 'w') as f:
                        f.write(server_uuid)
                self.server_status = "在线"
                print(f"注册成功，UUID: {server_uuid}")
            else:
                self.server_status = "注册失败（无UUID）"
                print("注册失败：服务器未返回UUID")
        else:
            self.server_status = "注册失败"
            print(f"注册失败，HTTP {response.status_code}")

    def check_server_status(self):
        if not self.online_mode or not self.online_ip:
            self.server_status = "离线 (未配置)"
            self.server_last_check = time.strftime("%Y-%m-%d %H:%M:%S")
            if hasattr(self, 'online_status_label'):
                color = self.offline_color
                self.online_status_label.config(text=f"服务器状态: {self.server_status}", fg=color)
            return

        response = self._request_with_auth('GET', f"http://{self.online_ip}:{self.server_port}/api/status")
        if response is None:
            self.server_status = "离线"
        elif response.status_code == 200:
            machines = response.json()
            found = False
            for m in machines:
                if m['uuid'] == self.client_uuid:
                    self.server_status = "在线" if m['online'] else "离线"
                    found = True
                    break
            if not found:
                self.server_status = "未注册"
                self.register_with_central()
        else:
            self.server_status = "离线"

        self.server_last_check = time.strftime("%Y-%m-%d %H:%M:%S")
        if hasattr(self, 'online_status_label'):
            color = self.online_color if self.server_status == "在线" else self.offline_color
            self.online_status_label.config(text=f"服务器状态: {self.server_status}", fg=color)
        if hasattr(self, 'last_check_label'):
            self.last_check_label.config(text=f"最后检查: {self.server_last_check}")

        self.window.after(30000, self.check_server_status)

    def sync_data_to_server(self):
        if not self.online_mode or self.server_status != "在线":
            self.status_var.set("无法同步: 服务器离线")
            return
        data_str = json.dumps(self.student_data, ensure_ascii=False)
        signature = self.sign_message(data_str)
        response = self._request_with_auth('POST',
            f"http://{self.online_ip}:{self.server_port}/api/sync_data",
            json={'uuid': self.client_uuid, 'signature': signature, 'data': data_str}
        )
        if response is None:
            return
        if response.status_code == 200:
            self.status_var.set("数据同步成功")
            print("同步成功")
        else:
            self.status_var.set(f"同步失败: {response.text}")
            print(f"同步失败: {response.status_code} {response.text}")

    def load_data_from_server(self):
        if not self.online_mode or self.server_status != "在线":
            return
        challenge = str(time.time())
        signature = self.sign_message(challenge)
        response = self._request_with_auth('POST',
            f"http://{self.online_ip}:{self.server_port}/api/load_data",
            json={'uuid': self.client_uuid, 'signature': signature, 'challenge': challenge}
        )
        if response is None:
            return
        if response.status_code == 200:
            data = response.json().get('data', {})
            if data:
                for name, d in data.items():
                    if name in self.student_data:
                        self.student_data[name] = d
                self.save_student_data()
                self.update_ui_from_data()
            self.status_var.set("数据加载成功")
        else:
            self.status_var.set(f"加载失败: {response.text}")

    def sync_config_to_server(self):
        if not self.online_mode or self.server_status != "在线":
            return
        config = {
            'school': self.school,
            'nj': self.nj,
            'class_id': self.class_id,
            'km': self.km,
            'z': self.z,
            'l': self.l
        }
        config_str = json.dumps(config, ensure_ascii=False)
        signature = self.sign_message(config_str)
        response = self._request_with_auth('POST',
            f"http://{self.online_ip}:{self.server_port}/api/update_config",
            json={'uuid': self.client_uuid, 'signature': signature, 'config': config}
        )
        if response is None:
            return
        if response.status_code == 200:
            print("配置同步成功")
        else:
            print(f"配置同步失败: {response.text}")

    def load_config_from_server(self):
        if not self.online_mode or self.server_status != "在线":
            return
        challenge = str(time.time())
        signature = self.sign_message(challenge)
        response = self._request_with_auth('POST',
            f"http://{self.online_ip}:{self.server_port}/api/get_config",
            json={'uuid': self.client_uuid, 'signature': signature, 'challenge': challenge}
        )
        if response is None:
            return
        if response.status_code == 200:
            remote_config = response.json().get('config', {})
            if remote_config:
                self.update_local_config(remote_config)
            print("配置加载成功")
        else:
            print(f"配置加载失败: {response.text}")

    def update_local_config(self, remote_config):
        global school, nj, class_id, km, z, l
        self.school = remote_config.get('school', self.school)
        self.nj = remote_config.get('nj', self.nj)
        self.class_id = remote_config.get('class_id', self.class_id)
        self.km = remote_config.get('km', self.km)
        self.z = remote_config.get('z', self.z)
        self.l = remote_config.get('l', self.l)

        conf.set('config', 'school', self.school)
        conf.set('config', 'nj', self.nj)
        conf.set('config', 'class_id', self.class_id)
        conf.set('config', 'km', self.km)
        conf.set('config', 'z', str(self.z))
        conf.set('config', 'l', str(self.l))
        with open('config.ini', 'w', encoding='UTF-8') as f:
            conf.write(f)

        self.window.title(f'{self.school}{self.nj}年{self.class_id}班{self.km}打卡 {version} 作者: 刘宇晨')

    def periodic_load_config(self):
        if self.online_mode and self.server_status == "在线":
            self.load_config_from_server()
        self.window.after(60000, self.periodic_load_config)

    def periodic_load_data(self):
        if self.online_mode and self.server_status == "在线":
            self.load_data_from_server()
        self.window.after(5000, self.periodic_load_data)

    # ---------- 本地数据管理 ----------
    def load_student_data(self):
        global file_path
        if not file_path:
            file_path = 'attendance.dat'
        if not os.path.exists("name.txt"):
            with open("name.txt", "w", encoding="UTF-8") as f:
                for i in range(1, 41):
                    f.write(f"学生{i}\n")
        with open("name.txt", encoding="UTF-8") as f:
            names = [line.strip() for line in f.readlines()]
        for name in names:
            self.student_data[name] = {"count": 0, "first_time": None, "history": []}
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="UTF-8") as f:
                for line in f:
                    if ':' in line:
                        parts = line.strip().split(':')
                        name = parts[0]
                        if name in self.student_data:
                            self.student_data[name]["count"] = int(parts[1])
                            self.student_data[name]["first_time"] = parts[2] if len(parts) > 2 and parts[2] else None
                            if len(parts) > 3 and parts[3]:
                                self.student_data[name]["history"] = parts[3].split('|')

    def save_student_data(self):
        with open(file_path, "w", encoding="UTF-8") as f:
            for name, data in self.student_data.items():
                history_str = '|'.join(data["history"])
                f.write(f"{name}:{data['count']}:{data['first_time'] or ''}:{history_str}\n")

    def save_config(self):
        """保存配置到config.ini"""
        conf.set('config', 'online', str(int(self.online_mode)))
        conf.set('config', 'bd_online', str(int(self.bd_online)))
        conf.set('config', 'online_ip', self.online_ip)
        conf.set('config', 'server_port', str(self.server_port))
        conf.set('config', 'server_password', self.server_password)
        conf.set('config', 'admin_password', self.admin_password_hash)
        conf.set('config', 'school', self.school)
        conf.set('config', 'nj', self.nj)
        conf.set('config', 'class_id', self.class_id)
        conf.set('config', 'km', self.km)
        conf.set('config', 'z', str(self.z))
        conf.set('config', 'l', str(self.l))
        with open('config.ini', 'w', encoding='UTF-8') as configfile:
            conf.write(configfile)

    def show_remote_settings(self):
        """显示远程设置窗口"""
        show_ui.show_settings_remot(self, tk)

    def save_remote_settings(self, settings_window):
        """保存远程设置"""
        self.online_ip = self.ip_entry.get()
        try:
            self.server_port = int(self.port_entry.get())
        except ValueError:
            messagebox.showerror("错误", "端口必须是整数")
            return
        self.server_password = self.password_entry.get()
        self.save_config()
        settings_window.destroy()
        self.check_server_status()
        messagebox.showinfo("提示", "设置已保存，点击确定重启程序生效")
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def start_api_server(self):
        """启动api服务器"""
        from mode_sever_start import start_api
        from flask import Flask
        api_app = Flask(__name__)
        api_queue = []
        start_api(self, api_app, api_queue, self.school, self.nj, self.class_id, self.km)

    def export_data(self):
        """导出csv数据文件函数"""
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
            )
            if not file_path:
                return
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("姓名,打卡时间,打卡次数,历史记录\n")
                for name, data in self.student_data.items():
                    first_time = data['first_time'] if data['first_time'] else "未打卡"
                    history = "|".join(data['history'])
                    f.write(f"{name},{first_time},{data['count']},{history}\n")
            self.status_var.set(f"数据已导出到: {file_path}")
            messagebox.showinfo("导出成功", "打卡数据已成功导出！")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出数据时出错: {str(e)}")

    def update_stats(self):
        total = len(self.student_data)
        punched = sum(1 for data in self.student_data.values() if data['first_time'])
        self.total_var.set(f"总人数: {total}")
        self.punched_var.set(f"已打卡: {punched} ({punched / total * 100:.1f}%)")

    def mark_attendance(self, name):
        """打卡函数"""

        current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        if self.student_data[name]['first_time']:
            self.status_var.set(f"{name} 已打卡，右键点击可取消打卡")
            return
        self.student_data[name]['first_time'] = current_time
        self.student_data[name]['count'] += 1
        self.student_data[name]['history'].append(current_time)
        self.buttons[name].config(bg=self.active_button_bg, fg=self.active_button_fg)
        self.status_var.set(f"{name} 打卡成功! 打卡时间: {current_time}")
        self.update_ranking()
        self.update_stats()
        self.save_student_data()
        if self.online_mode and self.server_status == "在线":
            self.sync_data_to_server()

    def update_ranking(self):
        """更新排行榜"""
        for item in self.ranking_tree.get_children():
            self.ranking_tree.delete(item)
        punched_students = []
        for name, data in self.student_data.items():
            if data['first_time']:
                punched_students.append((name, data['first_time']))
        punched_students.sort(key=lambda x: x[1])
        for rank, (name, first_time) in enumerate(punched_students, start=1):
            tags = ()
            if rank == 1:
                tags = ('first',)
            elif rank == 2:
                tags = ('second',)
            elif rank == 3:
                tags = ('third',)
            formatted_time = first_time
            if len(first_time) > 16:
                formatted_time = first_time[11:16]
            self.ranking_tree.insert("", "end", values=(rank, name, formatted_time), tags=tags)
        self.ranking_tree.tag_configure('first', background='#ffd700', foreground='black')
        self.ranking_tree.tag_configure('second', background='#c0c0c0', foreground='black')
        self.ranking_tree.tag_configure('third', background='#cd7f32', foreground='white')


def gat_file():
    return True


if __name__ == '__main__':
    try:
        if not gat_file():
            tk.messagebox.showerror('错误', '未授权')
    except Exception:
        tk.messagebox.showerror('错误', '无法连接到中心监管服务器')

    if len(sys.argv) > 1:
        impfile = True
        file_path = sys.argv[1]

    window = tk.Tk()
    window.geometry("1431x800")
    app = AttendanceApp(window)

    window.mainloop()
