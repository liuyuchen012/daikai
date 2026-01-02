# 开发者: 刘宇晨
# 联系邮箱: liuyuchen032901@outlook.com
# 开发环境: Python 3.11
# 系统环境: Windows 11
# 软件环境: PyCharm 2025.1.3
# 版本信息: 1.9
# 功能描述: 大屏打卡
# 本程序在中国大陆申请的版权仅意味这我是该版本的原始开发者
'''
    Check_in
    Copyright (C) 2026 En-us:LiuYuchen Zh-cn:刘宇晨

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'''
from pystray import MenuItem
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import time
from configparser import ConfigParser
import threading
import json
from datetime import datetime
import requests
from flask import Flask, request, jsonify
import sys
import webbrowser
from tkinter import font
import ctypes
import hashlib
from PIL import Image

# 定义Windows常量
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_DRAWFRAME = 0x0020
SWP_NOACTIVATE = 0x0010
SWP_FRAMECHANGED = 0x0020
GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_CAPTION = 0x00C00000
WS_SYSMENU = 0x00080000
WS_EX_APPWINDOW = 0x00040000
WS_EX_DLGMODALFRAME = 0x00000001
WM_SYSCOMMAND = 0x0112
SC_MOVE = 0xF010
HTCAPTION = 2
NULL = 0

##impfile = False
file_path = ""
# 版本变量
version = '2.0'  # 更新版本号

# 加载配置
directory = os.path.dirname(__file__)
conf = ConfigParser()
print(directory)
conf.read('config.ini', encoding='UTF-8')
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
# 读取管理员密码（使用SHA256哈希存储）
admin_password_hash = conf.get('config', 'admin_password', fallback='')
impfile = False

# 创建Flask应用用于API
api_app = Flask(__name__)
api_queue = []  # 用于存储API请求的队列
# Flask 2.0+ 兼容性设置
api_app.config['JSONIFY_PRETTYPRINT_REGULAR'] = False


def check_version():
    webbrowser.open(f'https://jay.615mc.cn/version/index.html?v={version}')


def open_url_bz():
    webbrowser.open('https://github.com/liuyuchen012/daikai/tree/' + version)


# 密码哈希函数
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


# 改进的窗口样式设置，修复任务栏图标问题
def setup_window_style(window):
    """设置窗口样式，保持任务栏图标"""
    try:
        # 获取窗口句柄
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        
        # 获取当前窗口样式
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
        
        # 移除默认的标题栏样式，但保留系统菜单以便在任务栏显示
        style = style & ~WS_CAPTION  # 移除标题栏
        style = style | WS_SYSMENU   # 添加系统菜单，以便任务栏显示
        
        # 设置新的窗口样式
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_STYLE, style)
        
        # 设置扩展样式，确保窗口在任务栏显示
        ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        ex_style = ex_style | WS_EX_APPWINDOW  # 强制窗口在任务栏显示
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex_style)
        
        # 更新窗口
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 
                                          SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
        
        return True
    except Exception as e:
        print(f"设置窗口样式失败: {e}")
        # 回退方案：使用常规窗口但移除标题栏
        window.overrideredirect(True)
        return False


# 优化后的窗口移动函数，解决花屏问题
def windowMove(widget, window):
    """优化的窗口移动函数，减少花屏和闪烁"""
    class DragManager:
        def __init__(self, window):
            self.window = window
            self.dragging = False
            self.start_x = 0
            self.start_y = 0
            self.window_x = 0
            self.window_y = 0
            self.last_update_time = 0
            self.update_interval = 16  # 约60fps，减少更新频率
            
        def start_drag(self, event):
            self.dragging = True
            self.start_x = event.x_root
            self.start_y = event.y_root
            self.window_x = self.window.winfo_x()
            self.window_y = self.window.winfo_y()
            self.last_update_time = time.time()
            
            # 在拖动开始时锁定窗口重绘
            try:
                hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, 
                    ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE) | 0x02000000)  # WS_EX_COMPOSITED
            except:
                pass
                
            # 设置光标为移动样式
            self.window.config(cursor="fleur")
            
        def do_drag(self, event):
            if not self.dragging:
                return
                
            current_time = time.time()
            if current_time - self.last_update_time < self.update_interval / 1000.0:
                return  # 限制更新频率
                
            dx = event.x_root - self.start_x
            dy = event.y_root - self.start_y
            new_x = self.window_x + dx
            new_y = self.window_y + dy
            
            # 使用Win32 API平滑移动窗口
            try:
                hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
                # 使用SWP_NOCOPYBITS避免复制背景，减少花屏
                ctypes.windll.user32.SetWindowPos(hwnd, 0, new_x, new_y, 0, 0,
                                                  SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | 0x0100)  # SWP_NOCOPYBITS
            except:
                # 如果Win32 API失败，回退到tkinter方法
                self.window.geometry(f"+{new_x}+{new_y}")
            
            self.last_update_time = current_time
            
        def end_drag(self, event=None):
            if not self.dragging:
                return
                
            self.dragging = False
            
            # 恢复光标样式
            self.window.config(cursor="")
            
            # 恢复窗口重绘
            try:
                hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, 
                    ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE) & ~0x02000000)  # 移除WS_EX_COMPOSITED
                # 强制重绘窗口
                ctypes.windll.user32.RedrawWindow(hwnd, None, None, 0x0001 | 0x0004 | 0x0100)  # RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME
            except:
                pass
                
            # 更新窗口位置
            self.window.update_idletasks()
    
    # 创建拖拽管理器
    drag_manager = DragManager(window)
    
    # 绑定事件
    def on_drag_start(event):
        drag_manager.start_drag(event)
        return "break"  # 阻止事件继续传播
        
    def on_drag_motion(event):
        drag_manager.do_drag(event)
        return "break"
        
    def on_drag_end(event):
        drag_manager.end_drag()
        return "break"
    
    # 绑定事件到widget
    widget.bind("<ButtonPress-1>", on_drag_start)
    widget.bind("<B1-Motion>", on_drag_motion)
    widget.bind("<ButtonRelease-1>", on_drag_end)
    
    # 确保鼠标离开窗口时也能结束拖拽
    widget.bind("<Leave>", lambda e: drag_manager.end_drag())
    
    # 也绑定到窗口本身，防止在某些情况下失去焦点
    window.bind("<ButtonRelease-1>", lambda e: drag_manager.end_drag())


class AttendanceApp:
    def __init__(self, window):
        self.window = window

        # 配置属性作为实例变量
        self.nj = nj
        self.class_id = class_id
        self.z = z
        self.l = l
        self.km = km
        self.school = school
        self.admin_password_hash = admin_password_hash

        self.window.title(f'{school}{nj}年{class_id}班{km}打卡 {version} 作者: 刘宇晨')
        self.window.geometry("1300x790")
        self.window.minsize(1024, 600)
        
        # 设置窗口图标
        self.set_window_icon()

        # 窗口状态变量
        self.maximized = False
        self.prev_geometry = ""

        # 使用PyCharm风格的深色主题
        self.dark_mode = False
        self.apply_theme()

        # 设置窗口样式（修复任务栏图标问题）
        self.setup_custom_window()

        # 在线状态变量
        self.online_mode = online_mode
        self.bd_online = bd_online
        self.online_ip = online_ip
        self.server_port = server_port
        self.server_status = "未连接"
        self.server_last_check = "从未检查"
        self.connection_attempts = 0

        # 初始化数据结构
        self.student_data = {}
        self.load_student_data()

        # 创建菜单栏
        self.create_menu()

        # 创建主容器
        self.main_container = tk.Frame(self.window, bg=self.bg_color)
        self.main_container.pack(fill=tk.BOTH, expand=True, pady=(30, 0))

        # 显示主界面
        self.show_main_interface()

        # 启动API服务器（如果配置为服务器）
        if self.bd_online:
            self.start_api_server()

        # 启动服务器状态检查（如果配置为客户端）
        if self.online_mode:
            self.check_server_status()
            self.window.after(1000, self.load_data_from_server)

        # 设置窗口样式
        self.set_pycharm_style()
        
        # 绑定窗口事件
        self.bind_window_events()

    def set_window_icon(self):
        """设置窗口图标"""
        try:
            icon_path = 'icon.ico'
            if os.path.exists(icon_path):
                # 尝试多种方法设置图标
                try:
                    self.window.iconbitmap(icon_path)
                except:
                    # 使用PhotoImage作为备选方案
                    img = tk.PhotoImage(file=icon_path)
                    self.window.iconphoto(False, img)
        except Exception as e:
            print(f"设置窗口图标失败: {e}")

    def setup_custom_window(self):
        """设置自定义窗口样式，修复任务栏图标问题"""
        # 创建自定义标题栏
        self.create_custom_titlebar()
        
        # 设置窗口样式，保持任务栏图标
        setup_window_style(self.window)
        
        # 设置应用程序ID，帮助Windows识别应用程序
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"AttendanceSystem.{school}.{nj}.{class_id}")
        except:
            pass

    def bind_window_events(self):
        """绑定窗口事件"""
        # 绑定窗口状态变化事件
        self.window.bind('<Map>', self.on_window_mapped)
        self.window.bind('<Unmap>', self.on_window_unmapped)
        
        # 绑定焦点事件
        self.window.bind('<FocusIn>', self.on_window_focus)
        self.window.bind('<FocusOut>', self.on_window_blur)
        
        # 绑定窗口大小变化事件
        self.window.bind('<Configure>', self.on_window_configure)

    def on_window_mapped(self, event):
        """窗口被映射时调用（显示）"""
        # 确保窗口在任务栏显示
        self.window.after(100, self.ensure_taskbar_visibility)

    def on_window_unmapped(self, event):
        """窗口被取消映射时调用（隐藏）"""
        pass

    def on_window_focus(self, event):
        """窗口获得焦点"""
        if hasattr(self, 'titlebar'):
            self.titlebar.config(bg=self.menu_bg)

    def on_window_blur(self, event):
        """窗口失去焦点"""
        if hasattr(self, 'titlebar'):
            self.titlebar.config(bg=self.menu_bg)

    def on_window_configure(self, event):
        """窗口大小或位置变化时调用"""
        # 如果窗口不是最大化状态，更新prev_geometry
        if not self.maximized and event.widget == self.window:
            self.prev_geometry = self.window.geometry()

    def ensure_taskbar_visibility(self):
        """确保窗口在任务栏可见"""
        try:
            hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
            
            # 检查窗口是否在任务栏
            ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            
            # 如果不在任务栏，强制添加APPWINDOW样式
            if not (ex_style & WS_EX_APPWINDOW):
                ex_style = ex_style | WS_EX_APPWINDOW
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex_style)
                ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 
                                                  SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
                
            # 确保窗口标题可见
            title = f'{school}{nj}年{class_id}班{km}打卡 {version} 作者: 刘宇晨'
            ctypes.windll.user32.SetWindowTextW(hwnd, title)
            
        except Exception as e:
            print(f"确保任务栏可见性失败: {e}")

    def set_pycharm_style(self):
        """设置PyCharm风格的窗口样式"""
        self.window.config(bg=self.bg_color)

        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="楷体", size=9)

        text_font = font.nametofont("TkTextFont")
        text_font.configure(family="Consolas", size=10)

        style = ttk.Style()
        style.theme_use('clam')

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
                        relief="flat")

        style.map("Custom.Treeview",
                  background=[('selected', self.active_button_bg)],
                  foreground=[('selected', self.active_button_fg)])

        style.map("Custom.Treeview.Heading",
                  background=[('active', self.menu_active_bg)])

    def reboot(self):
        messagebox.showinfo("提示", "设置已保存,请手动重新打开生效")
        self.window.quit()

    def create_custom_titlebar(self):
        """创建自定义标题栏，类似PyCharm风格"""
        self.titlebar = tk.Frame(self.window, bg=self.menu_bg, height=30, relief=tk.FLAT, bd=0)
        self.titlebar.pack(fill=tk.X)

        windowMove(self.titlebar, self.window)

        title_text = f'{school}{nj}年{class_id}班{km}打卡 {version}'
        self.title_label = tk.Label(
            self.titlebar,
            text=title_text,
            bg=self.menu_bg,
            fg=self.menu_fg,
            font=("楷体", 10)
        )
        self.title_label.pack(side=tk.LEFT, padx=10, pady=5)
        windowMove(self.title_label, self.window)

        controls = tk.Frame(self.titlebar, bg=self.menu_bg)
        controls.pack(side=tk.RIGHT)

        self.min_btn = tk.Button(
            controls,
            text="−",
            bg=self.menu_bg,
            fg=self.menu_fg,
            width=3,
            height=1,
            bd=0,
            relief=tk.FLAT,
            command=self.minimize_window
        )
        self.min_btn.pack(side=tk.LEFT)
        self.min_btn.bind("<Enter>", lambda e: self.min_btn.config(bg=self.menu_active_bg))
        self.min_btn.bind("<Leave>", lambda e: self.min_btn.config(bg=self.menu_bg))

        self.max_btn = tk.Button(
            controls,
            text="□",
            bg=self.menu_bg,
            fg=self.menu_fg,
            width=3,
            height=1,
            bd=0,
            relief=tk.FLAT,
            command=self.toggle_maximize
        )
        self.max_btn.pack(side=tk.LEFT)
        self.max_btn.bind("<Enter>", lambda e: self.max_btn.config(bg=self.menu_active_bg))
        self.max_btn.bind("<Leave>", lambda e: self.max_btn.config(bg=self.menu_bg))

        self.close_btn = tk.Button(
            controls,
            text="✕",
            bg=self.menu_bg,
            fg=self.menu_fg,
            width=3,
            height=1,
            bd=0,
            relief=tk.FLAT,
            command=self.window.quit
        )
        self.close_btn.pack(side=tk.LEFT)
        self.close_btn.bind("<Enter>", lambda e: self.close_btn.config(bg="#ea4335", fg="white"))
        self.close_btn.bind("<Leave>", lambda e: self.close_btn.config(bg=self.menu_bg, fg=self.menu_fg))

    def minimize_window(self):
        """最小化窗口"""
        self.window.iconify()

    def toggle_maximize(self):
        """切换窗口最大化/还原状态"""
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

    def create_menu(self):
        """创建菜单栏"""
        self.menubar = tk.Frame(self.window, bg=self.menu_bg, height=25, relief=tk.FLAT, bd=0)
        self.menubar.pack(fill=tk.X)
        windowMove(self.menubar, self.window)

        menu_items = [
            ("文件", ["导出打卡数据", "导入打卡数据", "清空打卡记录", "退出"]),
            ("远程", ["远程服务器设置", "检查服务器状态", "从服务器加载数据", "同步数据到服务器"]),
            ("设置", ["管理员设置"]),#("设置", ["管理员设置", "切换主题样式"]),我们删除了黑色主题
            ("帮助", ["Github", "检查版本列表", "关于"])
        ]

        commands = {
            "导出打卡数据": self.export_data,
            "导入打卡数据": self.import_data,
            "清空打卡记录": self.clear_attendance_records,
            "退出": self.window.quit,
            "远程服务器设置": self.show_remote_settings,
            "检查服务器状态": self.check_server_status,
            "从服务器加载数据": self.load_data_from_server,
            "同步数据到服务器": self.sync_data_to_server,
            "管理员设置": self.show_admin_settings,
            "切换主题样式": self.toggle_theme,
            "Github": open_url_bz,
            "检查版本列表": check_version,
            "关于": self.show_about
        }

        for item, subitems in menu_items:
            btn = tk.Menubutton(self.menubar, text=item, bg=self.menu_bg, fg=self.menu_fg,
                                activebackground=self.menu_active_bg, activeforeground=self.menu_active_fg,
                                bd=0, padx=10, pady=3, font=("楷体", 12), relief=tk.FLAT)
            btn.pack(side=tk.LEFT)

            submenu = tk.Menu(btn, tearoff=0, bg=self.menu_bg, fg=self.menu_fg,
                              activebackground=self.menu_active_bg, activeforeground=self.menu_active_fg,
                              font=("楷体", 9))
            btn.config(menu=submenu)

            for subitem in subitems:
                if subitem == "-":
                    submenu.add_separator()
                else:
                    submenu.add_command(label=subitem, command=commands.get(subitem))

    def toggle_theme(self):
        """切换主题样式"""
        self.dark_mode = not self.dark_mode
        self.apply_theme()
        self.show_main_interface()
        self.save_config()

    def verify_admin_password(self, action_name):
        """验证管理员密码"""
        if not self.admin_password_hash:
            # 如果没有设置密码，直接返回成功
            return True

        password = self.ask_password(f"需要管理员权限执行: {action_name}")
        if password and hash_password(password) == self.admin_password_hash:
            return True
        else:
            messagebox.showerror("错误", "管理员密码错误")
            return False

    def ask_password(self, title):
        """弹出密码输入对话框"""
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

        result = [None]  # 使用列表来存储结果，以便在内部函数中修改

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

    def clear_attendance_records(self):
        """清空打卡记录"""
        if not self.verify_admin_password("清空打卡记录"):
            return

        # 显示三次警告
        for i in range(3, 0, -1):
            result = messagebox.askyesno("警告",
                                         f"确定要清空所有打卡记录吗？此操作不可恢复！\n还有 {i} 次警告")
            if not result:
                return

        # 确认清空
        confirm = messagebox.askyesno("最终确认",
                                      "这是最后一次确认！确定要清空所有打卡记录吗？")
        if not confirm:
            return

        # 清空所有打卡记录
        for name in self.student_data:
            self.student_data[name]['count'] = 0
            self.student_data[name]['first_time'] = None
            self.student_data[name]['history'] = []

        # 更新UI
        self.update_ui_from_data()
        self.save_student_data()

        if hasattr(self, 'status_var'):
            self.status_var.set("所有打卡记录已清空")

        messagebox.showinfo("完成", "所有打卡记录已成功清空")

    def show_admin_settings(self):
        """显示管理员设置对话框，包含系统设置和密码设置"""
        if not self.verify_admin_password("访问管理员设置"):
            return

        settings_window = tk.Toplevel(self.window)
        settings_window.title("管理员设置")
        settings_window.geometry("600x500")
        settings_window.resizable(False, False)
        settings_window.transient(self.window)
        settings_window.grab_set()
        settings_window.config(bg=self.bg_color)

        # 添加标题
        title_font = font.Font(family="楷体", size=14, weight='bold')
        tk.Label(settings_window, text="管理员设置", font=title_font,
                 bg=self.bg_color, fg=self.fg_color).pack(pady=15)

        # 创建选项卡控件
        notebook = ttk.Notebook(settings_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 系统设置选项卡
        system_frame = ttk.Frame(notebook, padding=10)
        notebook.add(system_frame, text="系统设置")

        # 密码设置选项卡
        password_frame = ttk.Frame(notebook, padding=10)
        notebook.add(password_frame, text="密码设置")

        # 填充系统设置选项卡
        self.create_system_settings_tab(system_frame)

        # 填充密码设置选项卡
        self.create_password_settings_tab(password_frame)

        # 按钮区域
        button_frame = tk.Frame(settings_window, bg=self.bg_color)
        button_frame.pack(fill=tk.X, padx=20, pady=20)

        # 保存按钮
        save_btn = tk.Button(button_frame, text="保存所有设置",
                             command=lambda: self.save_all_settings(settings_window),
                             width=15, bg=self.button_bg, fg=self.button_fg,
                             activebackground=self.button_hover, relief=tk.FLAT,
                             font=("楷体", 9))
        save_btn.pack(side=tk.LEFT, padx=10)
        save_btn.bind("<Enter>", lambda e: save_btn.config(bg=self.button_hover))
        save_btn.bind("<Leave>", lambda e: save_btn.config(bg=self.button_bg))

        # 取消按钮
        cancel_btn = tk.Button(button_frame, text="取消",
                               command=settings_window.destroy,
                               width=15, bg=self.button_bg, fg=self.button_fg,
                               activebackground=self.button_hover, relief=tk.FLAT,
                               font=("楷体", 9))
        cancel_btn.pack(side=tk.RIGHT, padx=10)
        cancel_btn.bind("<Enter>", lambda e: cancel_btn.config(bg=self.button_hover))
        cancel_btn.bind("<Leave>", lambda e: cancel_btn.config(bg=self.button_bg))

    def create_system_settings_tab(self, parent):
        """创建系统设置选项卡内容"""
        # 学校信息设置
        tk.Label(parent, text="学校信息设置", font=("楷体", 12, 'bold')).pack(anchor=tk.W, pady=(0, 10))

        # 学校名称
        school_frame = tk.Frame(parent)
        school_frame.pack(fill=tk.X, pady=5)
        tk.Label(school_frame, text="学校名称:", width=15).pack(side=tk.LEFT)
        self.school_entry = tk.Entry(school_frame)
        self.school_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.school_entry.insert(0, self.school)

        # 年级
        nj_frame = tk.Frame(parent)
        nj_frame.pack(fill=tk.X, pady=5)
        tk.Label(nj_frame, text="年级:", width=15).pack(side=tk.LEFT)
        self.nj_entry = tk.Entry(nj_frame)
        self.nj_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.nj_entry.insert(0, self.nj)

        # 班级
        class_frame = tk.Frame(parent)
        class_frame.pack(fill=tk.X, pady=5)
        tk.Label(class_frame, text="班级:", width=15).pack(side=tk.LEFT)
        self.class_entry = tk.Entry(class_frame)
        self.class_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.class_entry.insert(0, self.class_id)

        # 课程
        km_frame = tk.Frame(parent)
        km_frame.pack(fill=tk.X, pady=5)
        tk.Label(km_frame, text="课程名称:", width=15).pack(side=tk.LEFT)
        self.km_entry = tk.Entry(km_frame)
        self.km_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.km_entry.insert(0, self.km)

        # 界面设置
        tk.Label(parent, text="界面设置", font=("楷体", 12, 'bold')).pack(anchor=tk.W, pady=(20, 10))

        # 行数
        row_frame = tk.Frame(parent)
        row_frame.pack(fill=tk.X, pady=5)
        tk.Label(row_frame, text="按钮行数:", width=15).pack(side=tk.LEFT)
        self.row_entry = tk.Entry(row_frame)
        self.row_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.row_entry.insert(0, str(self.z))

        # 列数
        col_frame = tk.Frame(parent)
        col_frame.pack(fill=tk.X, pady=5)
        tk.Label(col_frame, text="按钮列数:", width=15).pack(side=tk.LEFT)
        self.col_entry = tk.Entry(col_frame)
        self.col_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.col_entry.insert(0, str(self.l))

    def create_password_settings_tab(self, parent):
        """创建密码设置选项卡内容"""
        # 当前密码状态
        status_frame = tk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=10)

        status_text = "已设置管理员密码" if self.admin_password_hash else "未设置管理员密码"
        tk.Label(status_frame, text=f"当前状态: {status_text}",
                 font=("楷体", 9)).pack(side=tk.LEFT)

        # 新密码
        password_frame = tk.Frame(parent)
        password_frame.pack(fill=tk.X, pady=10)

        tk.Label(password_frame, text="新密码:", font=("楷体", 9)).pack(side=tk.LEFT)
        self.new_password_entry = tk.Entry(password_frame, show="*")
        self.new_password_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # 确认密码
        confirm_frame = tk.Frame(parent)
        confirm_frame.pack(fill=tk.X, pady=10)

        tk.Label(confirm_frame, text="确认密码:", font=("楷体", 9)).pack(side=tk.LEFT)
        self.confirm_password_entry = tk.Entry(confirm_frame, show="*")
        self.confirm_password_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # 清除密码按钮
        clear_btn = tk.Button(parent, text="清除密码",
                              command=lambda: self.clear_admin_password(settings_window=None),
                              width=15, bg=self.button_bg, fg=self.button_fg,
                              activebackground=self.button_hover, relief=tk.FLAT,
                              font=("楷体", 9))
        clear_btn.pack(pady=20)
        clear_btn.bind("<Enter>", lambda e: clear_btn.config(bg=self.button_hover))
        clear_btn.bind("<Leave>", lambda e: clear_btn.config(bg=self.button_bg))

    def save_all_settings(self, settings_window):
        """保存所有设置（系统设置和密码设置）"""
        try:
            # 保存系统设置
            self.school = self.school_entry.get()
            self.nj = self.nj_entry.get()
            self.class_id = self.class_entry.get()
            self.km = self.km_entry.get()

            # 验证行列数
            self.z = int(self.row_entry.get())
            self.l = int(self.col_entry.get())

            if self.z <= 0 or self.l <= 0:
                raise ValueError("行数和列数必须为正数")

            # 保存密码设置（如果有输入新密码）
            new_password = self.new_password_entry.get()
            confirm_password = self.confirm_password_entry.get()

            if new_password:
                if new_password != confirm_password:
                    raise ValueError("两次输入的密码不一致")
                self.admin_password_hash = hash_password(new_password)

            # 保存配置
            self.save_config()
            settings_window.destroy()

            # 应用主题和更新界面
            self.apply_theme()
            self.show_main_interface()

            messagebox.showinfo("提示", "所有设置已保存并生效")
            self.reboot()
        except ValueError as e:
            messagebox.showerror("错误", f"输入无效: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {str(e)}")

    def clear_admin_password(self, settings_window=None):
        """清除管理员密码"""
        confirm = messagebox.askyesno("确认", "确定要清除管理员密码吗？")
        if confirm:
            self.admin_password_hash = ''
            self.save_config()
            if settings_window:
                settings_window.destroy()
            messagebox.showinfo("成功", "管理员密码已清除")

    def apply_theme(self):
        """应用主题颜色"""
        if 0 == 1:#self.dark_mode: # 我们删除了黑色主题
            # PyCharm深色主题配置
            self.bg_color = '#2b2b2b'  # 深灰色背景
            self.fg_color = '#a9b7c6'  # 浅灰色文字
            self.status_fg_color = '#a9b7c6'  # 状态文字
            self.button_bg = '#3c3f41'  # 按钮背景
            self.button_fg = '#a9b7c6'  # 按钮文字
            self.button_hover = '#4e5254'  # 按钮悬停效果
            self.active_button_bg = '#4e5254'  # 已打卡按钮背景
            self.active_button_fg = '#ffffff'  # 已打卡按钮文字
            self.frame_bg = '#3c3f41'  # 框架背景
            self.frame_border = '#515151'  # 框架边框
            self.tree_bg = '#2b2b2b'  # 列表背景
            self.tree_fg = '#a9b7c6'  # 列表文字
            self.status_bg = '#3c3f41'  # 状态栏背景
            self.menu_bg = '#3c3f41'  # 菜单背景
            self.menu_fg = '#a9b7c6'  # 菜单文字
            self.menu_active_bg = '#4e5254'  # 菜单选中背景
            self.menu_active_fg = '#ffffff'  # 菜单选中文字
            self.online_color = '#6a8759'  # 绿色表示在线
            self.offline_color = '#bc3f3c'  # 红色表示离线
        else:
            # 旧式浅色主题配置
            self.bg_color = '#f0f0f0'  # 浅灰色背景
            self.fg_color = '#000000'  # 黑色文字
            self.status_fg_color = '#333333'  # 状态文字稍浅
            self.button_bg = '#e0e0e0'  # 按钮背景
            self.button_fg = '#000000'  # 按钮文字
            self.button_hover = '#d0d0d0'  # 按钮悬停效果
            self.active_button_bg = '#4285f4'  # 蓝色表示已打卡
            self.active_button_fg = '#ffffff'  # 已打卡按钮文字
            self.frame_bg = '#e8e8e8'  # 框架背景
            self.frame_border = '#cccccc'  # 框架边框
            self.tree_bg = '#ffffff'  # 列表背景
            self.tree_fg = '#000000'  # 列表文字
            self.status_bg = '#e0e0e0'  # 状态栏背景
            self.menu_bg = '#e8e8e8'  # 菜单背景
            self.menu_fg = '#000000'  # 菜单文字
            self.menu_active_bg = '#d0d0d0'  # 菜单选中背景
            self.menu_active_fg = '#000000'  # 菜单选中文字
            self.online_color = '#34a853'  # 绿色表示在线
            self.offline_color = '#ea4335'  # 红色表示离线

        # 设置窗口背景
        self.window.config(bg=self.bg_color)

        # 更新自定义标题栏
        if hasattr(self, 'titlebar'):
            self.titlebar.config(bg=self.menu_bg, borderwidth=0, relief=tk.FLAT)
            self.title_label.config(bg=self.menu_bg, fg=self.menu_fg)
            # 更新标题栏按钮
            for btn in [self.min_btn, self.max_btn, self.close_btn]:
                btn.config(bg=self.menu_bg, fg=self.menu_fg)

        # 更新菜单栏颜色
        if hasattr(self, 'menubar'):
            self.menubar.config(bg=self.menu_bg, borderwidth=0, relief=tk.FLAT)
            for child in self.menubar.winfo_children():
                if isinstance(child, tk.Menubutton):
                    child.config(bg=self.menu_bg, fg=self.menu_fg,
                                 activebackground=self.menu_active_bg,
                                 activeforeground=self.menu_active_fg)

        # 更新在线状态标签颜色
        if hasattr(self, 'online_status_label'):
            if self.server_status == "在线":
                self.online_status_label.config(fg=self.online_color)
            else:
                self.online_status_label.config(fg=self.offline_color)

        # 更新主容器背景
        if hasattr(self, 'main_container'):
            self.main_container.config(bg=self.bg_color)

        # 更新所有控件样式
        self.update_widget_style(self.window)

        # 特别处理Treeview样式
        self.update_treeview_style()

    def update_widget_style(self, widget):
        """递归更新所有控件的样式"""
        if isinstance(widget, tk.Label):
            # 区分状态标签和其他标签
            if hasattr(widget, 'is_status_label') and widget.is_status_label:
                widget.config(bg=self.bg_color, fg=self.status_fg_color)
            else:
                widget.config(bg=self.bg_color, fg=self.fg_color)
        elif isinstance(widget, (tk.Button, tk.Menubutton)):
            widget.config(bg=self.button_bg, fg=self.button_fg,
                          activebackground=self.menu_active_bg,
                          activeforeground=self.menu_active_fg,
                          relief=tk.FLAT)
        elif isinstance(widget, (tk.Frame, tk.LabelFrame)):
            widget.config(bg=self.frame_bg)
            # 递归处理子控件
            for child in widget.winfo_children():
                self.update_widget_style(child)
        elif isinstance(widget, tk.Entry):
            widget.config(bg=self.bg_color, fg=self.fg_color,
                          insertbackground=self.fg_color,
                          relief=tk.FLAT)
        elif isinstance(widget, tk.Checkbutton):
            widget.config(bg=self.frame_bg, fg=self.fg_color,
                          selectcolor=self.frame_bg)

    def update_treeview_style(self):
        """专门更新Treeview的样式"""
        if hasattr(self, 'ranking_tree'):
            style = ttk.Style()

            # 重置Treeview样式
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

            # 列表选中行样式
            style.map("Custom.Treeview",
                      background=[('selected', self.active_button_bg)],
                      foreground=[('selected', self.active_button_fg)])

            # 应用自定义样式
            self.ranking_tree.configure(style="Custom.Treeview")

    def save_config(self):
        """保存配置到config.ini"""
        conf.set('config', 'online', str(int(self.online_mode)))
        conf.set('config', 'bd_online', str(int(self.bd_online)))
        conf.set('config', 'online_ip', self.online_ip)
        conf.set('config', 'server_port', str(self.server_port))
        conf.set('config', 'nj', self.nj)
        conf.set('config', 'class_id', self.class_id)
        conf.set('config', 'z', str(self.z))
        conf.set('config', 'l', str(self.l))
        conf.set('config', 'km', self.km)
        conf.set('config', 'school', self.school)
        conf.set('config', 'admin_password', self.admin_password_hash)

        with open('config.ini', 'w', encoding='UTF-8') as configfile:
            conf.write(configfile)

    def show_remote_settings(self):
        """显示远程服务器设置对话框"""
        settings_window = tk.Toplevel(self.window)
        settings_window.title("远程服务器设置")
        settings_window.geometry("450x320")
        settings_window.resizable(False, False)
        settings_window.transient(self.window)
        settings_window.grab_set()
        settings_window.config(bg=self.bg_color)

        # 添加标题
        title_font = font.Font(family="楷体", size=14, weight='bold')
        tk.Label(settings_window, text="远程连接设置", font=title_font,
                 bg=self.bg_color, fg=self.fg_color).pack(pady=15)

        # 创建设置框架
        frame = tk.Frame(settings_window, bg=self.frame_bg, bd=1, relief=tk.FLAT,
                         highlightbackground=self.frame_border)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 在线模式
        online_frame = tk.Frame(frame, bg=self.frame_bg)
        online_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(online_frame, text="客户端模式:", bg=self.frame_bg, fg=self.fg_color,
                 font=("楷体", 9)).pack(side=tk.LEFT)
        self.online_var = tk.BooleanVar(value=self.online_mode)
        online_check = tk.Checkbutton(online_frame, variable=self.online_var,
                                      bg=self.frame_bg, fg=self.fg_color, selectcolor=self.frame_bg,
                                      command=lambda: setattr(self, 'online_mode', self.online_var.get()))
        online_check.pack(side=tk.LEFT, padx=10)

        # 服务器模式
        server_frame = tk.Frame(frame, bg=self.frame_bg)
        server_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(server_frame, text="服务器模式:", bg=self.frame_bg, fg=self.fg_color,
                 font=("楷体", 9)).pack(side=tk.LEFT)
        self.server_var = tk.BooleanVar(value=self.bd_online)
        server_check = tk.Checkbutton(server_frame, variable=self.server_var,
                                      bg=self.frame_bg, fg=self.fg_color, selectcolor=self.frame_bg,
                                      command=lambda: setattr(self, 'bd_online', self.server_var.get()))
        server_check.pack(side=tk.LEFT, padx=10)

        # 服务器地址
        ip_frame = tk.Frame(frame, bg=self.frame_bg)
        ip_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(ip_frame, text="服务器地址:", bg=self.frame_bg, fg=self.fg_color,
                 font=("楷体", 9)).pack(side=tk.LEFT)
        self.ip_entry = tk.Entry(ip_frame, width=20, bg=self.bg_color, fg=self.fg_color,
                                 insertbackground=self.fg_color, relief=tk.FLAT)
        self.ip_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.ip_entry.insert(0, self.online_ip)

        # 服务器端口
        port_frame = tk.Frame(frame, bg=self.frame_bg)
        port_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(port_frame, text="服务器端口:", bg=self.frame_bg, fg=self.fg_color,
                 font=("楷体", 9)).pack(side=tk.LEFT)
        self.port_entry = tk.Entry(port_frame, width=10, bg=self.bg_color, fg=self.fg_color,
                                   insertbackground=self.fg_color, relief=tk.FLAT)
        self.port_entry.pack(side=tk.LEFT, padx=10)
        self.port_entry.insert(0, str(self.server_port))
        # 按钮区域
        button_frame = tk.Frame(settings_window, bg=self.bg_color)
        button_frame.pack(fill=tk.X, padx=20, pady=20)

        # 创建带悬停效果的按钮
        save_btn = tk.Button(button_frame, text="保存设置",
                             command=lambda: self.save_remote_settings(settings_window),
                             width=12, bg=self.button_bg, fg=self.button_fg,
                             activebackground=self.button_hover, relief=tk.FLAT,
                             font=("楷体", 9))
        save_btn.pack(side=tk.LEFT, padx=10)
        save_btn.bind("<Enter>", lambda e: save_btn.config(bg=self.button_hover))
        save_btn.bind("<Leave>", lambda e: save_btn.config(bg=self.button_bg))

        cancel_btn = tk.Button(button_frame, text="取消",
                               command=settings_window.destroy,
                               width=12, bg=self.button_bg, fg=self.button_fg,
                               activebackground=self.button_hover, relief=tk.FLAT,
                               font=("楷体", 9))
        cancel_btn.pack(side=tk.RIGHT, padx=10)
        cancel_btn.bind("<Enter>", lambda e: cancel_btn.config(bg=self.button_hover))
        cancel_btn.bind("<Leave>", lambda e: cancel_btn.config(bg=self.button_bg))

    def save_remote_settings(self, settings_window):
        """保存远程设置并关闭窗口"""
        self.online_ip = self.ip_entry.get()
        try:
            self.server_port = int(self.port_entry.get())
            if self.server_port < 1 or self.server_port > 65535:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "端口必须是1-65535之间的整数")
            return

        self.bd_online = self.server_var.get()
        self.online_mode = self.online_var.get()
        self.save_config()
        settings_window.destroy()
        self.check_server_status()
        self.reboot()

    def start_api_server(self):
        """启动API服务器线程"""

        def run_api():
            # 设置API路由
            @api_app.route('/punch', methods=['POST'])
            def api_punch():
                try:
                    data = request.get_json()
                    name = data.get('name')
                    if name:
                        # 将打卡请求添加到队列
                        api_queue.append({
                            'name': name,
                            'time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            'ip': request.remote_addr
                        })
                        return jsonify({
                            'status': 'success',
                            'message': f'{name}的打卡请求已接收，请等待处理'
                        }), 200
                    else:
                        return jsonify({
                            'status': 'error',
                            'message': '缺少姓名参数'
                        }), 400
                except Exception as e:
                    return jsonify({
                        'status': 'error',
                        'message': f'服务器错误: {str(e)}'
                    }), 500

            @api_app.route("/pull_name", methods=['GET'])
            def api_pull_name():
                return jsonify(self.student_data), 200

            @api_app.route('/status', methods=['GET'])
            def api_status():
                total = len(self.student_data)
                punched = sum(1 for data in self.student_data.values() if data['first_time'])
                return jsonify({
                    'status': 'running',
                    'class': f'{school}{nj}年{class_id}班',
                    'subject': km,
                    'total': total,
                    'punched': punched,
                    'punched_percent': f'{punched / total * 100:.1f}%' if total > 0 else '0.0%'
                }), 200

            @api_app.route('/get_data', methods=['GET'])
            def api_get_data():
                """获取所有打卡数据"""
                return jsonify(self.student_data), 200

            @api_app.route('/get_names', methods=['GET'])
            def api_get_names():
                """获取学生名单"""
                return jsonify(list(self.student_data.keys())), 200

            @api_app.route('/update_data', methods=['POST'])
            def api_update_data():
                """更新所有打卡数据"""
                try:
                    new_data = request.get_json()
                    self.student_data = new_data
                    self.save_student_data()

                    # 更新UI
                    self.update_ui_from_data()

                    return jsonify({
                        'status': 'success',
                        'message': '数据更新成功'
                    }), 200
                except Exception as e:
                    return jsonify({
                        'status': 'error',
                        'message': f'更新数据失败: {str(e)}'
                    }), 500

            # 启动Flask应用
            api_app.run(host='0.0.0.0', port=self.server_port, threaded=True)

        # 启动API线程
        api_thread = threading.Thread(target=run_api, daemon=True)
        api_thread.start()

        # 更新服务器状态
        self.server_status = "在线"
        if hasattr(self, 'online_status_label'):
            self.online_status_label.config(text=f"服务器状态: {self.server_status}", fg=self.online_color)
        if hasattr(self, 'status_var'):
            self.status_var.set("API服务器已启动，等待连接...")

        # 定期检查API队列
        self.check_api_queue()

    def check_api_queue(self):
        """定期检查并处理API队列中的请求"""
        while api_queue:
            req = api_queue.pop(0)
            name = req['name']
            remote_time = req['time']
            ip = req['ip']

            if name in self.student_data:
                # 如果学生已打卡，则忽略重复请求
                if not self.student_data[name]['first_time']:
                    current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                    self.student_data[name]['first_time'] = current_time
                    self.student_data[name]['count'] += 1
                    self.student_data[name]['history'].append(current_time)

                    # 更新按钮颜色
                    if name in self.buttons:
                        self.buttons[name].config(bg=self.active_button_bg, fg=self.active_button_fg)

                    # 更新状态
                    if hasattr(self, 'status_var'):
                        self.status_var.set(f"远程打卡: {name} (来自 {ip})")

                    # 更新排名和统计
                    self.update_ranking()
                    self.update_stats()
                    self.save_student_data()
            else:
                if hasattr(self, 'status_var'):
                    self.status_var.set(f"无效的远程打卡请求: {name} (来自 {ip})")

        # 每500毫秒检查一次队列
        self.window.after(500, self.check_api_queue)

    def check_server_status(self):
        """检查远程服务器状态"""
        if not self.online_mode or not self.online_ip:
            self.server_status = "离线 (客户端模式未启用)"
            self.server_last_check = time.strftime("%Y-%m-%d %H:%M:%S")
            if hasattr(self, 'online_status_label'):
                self.online_status_label.config(text=f"服务器状态: {self.server_status}", fg=self.offline_color)
            return

        try:
            # 尝试连接服务器
            response = requests.get(f"http://{self.online_ip}:{self.server_port}/status", timeout=2)
            if response.status_code == 200:
                self.server_status = "在线"
                self.connection_attempts = 0
            else:
                self.server_status = f"离线 (HTTP {response.status_code})"
                self.connection_attempts += 1
        except Exception as e:
            self.server_status = f"离线 ({str(e)})"
            self.connection_attempts += 1

        self.server_last_check = time.strftime("%Y-%m-%d %H:%M:%S")

        # 更新UI
        if hasattr(self, 'online_status_label'):
            if self.server_status == "在线":
                self.online_status_label.config(text=f"服务器状态: {self.server_status}", fg=self.online_color)
            else:
                self.online_status_label.config(text=f"服务器状态: {self.server_status}", fg=self.offline_color)

        # 如果连接失败超过5次，显示警告
        if hasattr(self, 'status_var') and self.connection_attempts > 5:
            self.status_var.set(f"警告: 无法连接服务器，已尝试 {self.connection_attempts} 次")

        # 每30秒检查一次服务器状态
        self.window.after(30000, self.check_server_status)

    def load_data_from_server(self):
        """从远程服务器加载数据，包括学生名单"""
        if not self.online_mode or not self.server_status == "在线":
            if hasattr(self, 'status_var'):
                self.status_var.set("无法从服务器加载数据: 未启用在线模式或服务器离线")
            return

        try:
            # 1. 从服务器获取学生名单
            response = requests.get(f"http://{self.online_ip}:{self.server_port}/get_names", timeout=5)
            if response.status_code == 200:
                server_names = response.json()

                # 更新本地name.txt文件
                with open('name.txt', 'w', encoding='UTF-8') as f:
                    for name in server_names:
                        f.write(f"{name}\n")

                # 重新初始化学生数据结构
                self.student_data = {}
                for name in server_names:
                    self.student_data[name] = {"count": 0, "first_time": None, "history": []}
            else:
                if hasattr(self, 'status_var'):
                    self.status_var.set(f"从服务器获取学生名单失败: HTTP {response.status_code}")

            # 2. 从服务器获取打卡数据
            response = requests.get(f"http://{self.online_ip}:{self.server_port}/get_data", timeout=5)
            if response.status_code == 200:
                server_data = response.json()

                # 合并服务器数据到本地
                for name, data in server_data.items():
                    if name in self.student_data:
                        self.student_data[name] = data

                self.save_student_data()
                self.update_ui_from_data()
                if hasattr(self, 'status_var'):
                    self.status_var.set("数据已从服务器成功加载")
            else:
                if hasattr(self, 'status_var'):
                    self.status_var.set(f"从服务器加载打卡数据失败: HTTP {response.status_code}")
        except Exception as e:
            if hasattr(self, 'status_var'):
                self.status_var.set(f"从服务器加载数据失败: {str(e)}")

    def sync_data_to_server(self):
        """将本地数据同步到服务器"""
        if not self.online_mode or not self.server_status == "在线":
            if hasattr(self, 'status_var'):
                self.status_var.set("无法同步数据到服务器: 未启用在线模式或服务器离线")
            return

        try:
            response = requests.post(
                f"http://{self.online_ip}:{self.server_port}/update_data",
                json=self.student_data,
                timeout=5
            )
            if response.status_code == 200:
                if hasattr(self, 'status_var'):
                    self.status_var.set("数据已成功同步到服务器")
            else:
                if hasattr(self, 'status_var'):
                    self.status_var.set(f"同步数据到服务器失败: HTTP {response.status_code}")
        except Exception as e:
            if hasattr(self, 'status_var'):
                self.status_var.set(f"同步数据到服务器失败: {str(e)}")

    def export_data(self):
        """导出打卡数据到CSV文件"""
        try:
            # 弹出文件保存对话框
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
            )
            if not file_path:
                return  # 用户取消了保存

            with open(file_path, 'w', encoding='utf-8') as f:
                # 写入CSV表头
                f.write("姓名,打卡时间,打卡次数,历史记录\n")

                # 写入每个学生的数据
                for name, data in self.student_data.items():
                    first_time = data['first_time'] if data['first_time'] else "未打卡"
                    history = "|".join(data['history'])
                    f.write(f"{name},{first_time},{data['count']},{history}\n")

            if hasattr(self, 'status_var'):
                self.status_var.set(f"数据已导出到: {file_path}")
            messagebox.showinfo("导出成功", "打卡数据已成功导出！")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出数据时出错: {str(e)}")

    def import_data(self):
        """从CSV文件导入打卡数据"""
        try:
            # 弹出文件选择对话框
            file_path = filedialog.askopenfilename(
                filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
            )
            if not file_path:
                return  # 用户取消了选择

            with open(file_path, 'r', encoding='utf-8') as f:
                # 跳过表头
                next(f)

                new_data = {}
                for line in f:
                    parts = line.strip().split(',')
                    if len(parts) < 4:
                        continue

                    name = parts[0]
                    first_time = parts[1] if parts[1] != "未打卡" else None
                    count = int(parts[2])
                    history = parts[3].split('|') if parts[3] else []

                    new_data[name] = {
                        'count': count,
                        'first_time': first_time,
                        'history': history
                    }

            # 更新学生数据
            self.student_data = new_data
            self.save_student_data()
            self.update_ui_from_data()

            if hasattr(self, 'status_var'):
                self.status_var.set(f"数据已从 {file_path} 导入")
            messagebox.showinfo("导入成功", "打卡数据已成功导入！")

            # 如果是客户端模式，同步到服务器
            if self.online_mode and self.server_status == "在线":
                self.sync_data_to_server()
        except Exception as e:
            messagebox.showerror("导入失败", f"导入数据时出错: {str(e)}")

    def update_ui_from_data(self):
        """根据最新数据更新UI"""
        # 更新按钮状态
        if hasattr(self, 'buttons'):
            for name, data in self.student_data.items():
                if name in self.buttons:
                    if data['first_time']:
                        self.buttons[name].config(bg=self.active_button_bg, fg=self.active_button_fg)
                    else:
                        self.buttons[name].config(bg=self.button_bg, fg=self.button_fg)

        # 更新统计和排名
        self.update_stats()
        self.update_ranking()

    def load_student_data(self):
        global impfile
        global file_path
        # 尝试加载历史数据
        if not impfile:
            file_path = 'attendance.dat'

        """加载学生数据"""
        if not os.path.exists("name.txt"):
            # 创建示例数据文件
            with open("name.txt", "w", encoding="UTF-8") as f:
                for i in range(1, 41):
                    f.write(f"学生{i}\n")

        # 读取学生姓名
        with open("name.txt", encoding="UTF-8") as f:
            names = [line.strip() for line in f.readlines()]

        # 初始化学生数据
        for name in names:
            self.student_data[name] = {"count": 0, "first_time": None, "history": []}

        if os.path.exists(file_path):
            with open(file_path, "r", encoding="UTF-8") as f:
                for line in f:
                    if ':' in line:
                        parts = line.strip().split(':')
                        name = parts[0]
                        if name in self.student_data:
                            # 格式: name:count:first_time:history
                            self.student_data[name]["count"] = int(parts[1])
                            self.student_data[name]["first_time"] = parts[2] if len(parts) > 2 and parts[2] else None

                            # 加载历史记录
                            if len(parts) > 3 and parts[3]:
                                history_records = parts[3].split('|')
                                self.student_data[name]["history"] = history_records

    def save_student_data(self):
        """保存学生数据"""
        with open(file_path, "w", encoding="UTF-8") as f:
            for name, data in self.student_data.items():
                # 保存格式: name:count:first_time:history
                history_str = '|'.join(data["history"])
                f.write(f"{name}:{data['count']}:{data['first_time'] or ''}:{history_str}\n")

    def show_main_interface(self):
        """显示主界面"""
        for widget in self.main_container.winfo_children():
            widget.destroy()

        main_frame = tk.Frame(self.main_container, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 左侧排名区域
        left_frame = tk.LabelFrame(main_frame, text="打卡排名", font=("楷体", 14, 'bold'),
                                   bg=self.frame_bg, fg=self.fg_color, padx=10, pady=10,
                                   bd=1, relief=tk.FLAT, highlightbackground=self.frame_border)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 20))

        tk.Label(left_frame, text="打卡时间排名 (最早打卡)", font=("楷体", 12, 'bold'),
                 bg=self.frame_bg, fg=self.fg_color).pack(pady=(0, 10))

        columns = ("排名", "姓名", "打卡时间")
        self.ranking_tree = ttk.Treeview(left_frame, columns=columns, show="headings",
                                         height=20, style="Custom.Treeview")

        self.ranking_tree.heading("排名", text="排名")
        self.ranking_tree.heading("姓名", text="姓名")
        self.ranking_tree.heading("打卡时间", text="打卡时间")

        self.ranking_tree.column("排名", width=50, anchor=tk.CENTER)
        self.ranking_tree.column("姓名", width=100, anchor=tk.CENTER)
        self.ranking_tree.column("打卡时间", width=150, anchor=tk.CENTER)

        self.ranking_tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.ranking_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.ranking_tree.configure(yscrollcommand=scrollbar.set)

        # 右侧学生区域
        right_frame = tk.Frame(main_frame, bg=self.bg_color)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        header_frame = tk.Frame(right_frame, bg=self.bg_color)
        header_frame.pack(fill=tk.X, pady=(0, 20))

        title_text = f'{school}{nj}年{class_id}班{km}打卡'
        if self.online_mode:
            title_text += " (在线模式)"
        elif self.bd_online:
            title_text += " (服务器模式)"

        tk.Label(header_frame, text=title_text, font=("楷体", 16, 'bold'),
                 bg=self.bg_color, fg=self.fg_color).pack(pady=(0, 10))

        status_frame = tk.Frame(header_frame, bg=self.bg_color)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        status_color = self.online_color if self.server_status == "在线" else self.offline_color
        self.online_status_label = tk.Label(
            status_frame,
            text=f"服务器状态: {self.server_status}",
            font=("楷体", 9),
            bg=self.bg_color,
            fg=status_color
        )
        self.online_status_label.is_status_label = True
        self.online_status_label.pack(side=tk.LEFT, padx=10)

        self.last_check_label = tk.Label(
            status_frame,
            text=f"最后检查: {self.server_last_check}",
            font=("楷体", 9),
            bg=self.bg_color,
            fg=self.status_fg_color
        )
        self.last_check_label.is_status_label = True
        self.last_check_label.pack(side=tk.RIGHT, padx=10)

        self.stats_frame = tk.Frame(header_frame, bg=self.bg_color)
        self.stats_frame.pack(fill=tk.X)

        self.total_var = tk.StringVar()
        self.punched_var = tk.StringVar()
        self.update_stats()

        tk.Label(self.stats_frame, textvariable=self.total_var, font=("楷体", 10),
                 bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT, padx=10)
        tk.Label(self.stats_frame, textvariable=self.punched_var, font=("楷体", 10),
                 bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT, padx=10)

        button_container = tk.Frame(right_frame, bg=self.frame_bg, bd=1, relief=tk.FLAT,
                                    highlightbackground=self.frame_border)
        button_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        button_frame = tk.Frame(button_container, bg=self.frame_bg)
        button_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.buttons = {}
        names = list(self.student_data.keys())

        total_students = len(names)
        self.z = z if z > 0 else 6
        self.l = l if l > 0 else 6

        for i in range(self.z):
            for j in range(self.l):
                idx = i * self.l + j
                if idx < total_students:
                    name = names[idx]
                    if self.student_data[name]['first_time']:
                        bg_color = self.active_button_bg
                        fg_color = self.active_button_fg
                    else:
                        bg_color = self.button_bg
                        fg_color = self.button_fg

                    btn = tk.Button(
                        button_frame,
                        text=name,
                        font=("楷体", 15),
                        bg=bg_color,
                        fg=fg_color,
                        height=2,
                        width=12,
                        relief=tk.FLAT,
                        bd=1,
                        command=lambda n=name: self.mark_attendance(n)
                    )
                    btn.bind("<Enter>", lambda e, b=btn, n=name: self.on_button_enter(e, b, n))
                    btn.bind("<Leave>", lambda e, b=btn, n=name: self.on_button_leave(e, b, n))
                    btn.bind("<Button-3>", lambda event, n=name: self.show_context_menu(event, n))
                    btn.grid(row=i, column=j, padx=5, pady=5, sticky="nsew")
                    self.buttons[name] = btn

        for i in range(self.z):
            button_frame.grid_rowconfigure(i, weight=1)
        for j in range(self.l):
            button_frame.grid_columnconfigure(j, weight=1)

        self.status_var = tk.StringVar()
        self.status_var.set("就绪 - 左键点击学生姓名进行打卡，右键点击取消打卡")
        status_bar = tk.Label(self.window, textvariable=self.status_var,
                              bd=1, relief=tk.SUNKEN, anchor=tk.W,
                              bg=self.status_bg, fg=self.status_fg_color,
                              font=("楷体", 9))
        status_bar.is_status_label = True
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.update_ranking()

    def on_button_enter(self, event, button, name):
        """按钮鼠标悬停效果"""
        if self.student_data[name]['first_time']:
            button.config(bg=self.active_button_bg, fg=self.active_button_fg)
        else:
            button.config(bg=self.button_hover)

    def on_button_leave(self, event, button, name):
        """按钮鼠标离开效果"""
        if self.student_data[name]['first_time']:
            button.config(bg=self.active_button_bg, fg=self.active_button_fg)
        else:
            button.config(bg=self.button_bg)

    def show_context_menu(self, event, name):
        """显示右键菜单（取消打卡）"""
        if not self.student_data[name]['first_time']:
            return  # 未打卡的学生不需要取消打卡菜单

        menu = tk.Menu(self.window, tearoff=0, bg=self.menu_bg, fg=self.menu_fg,
                       activebackground=self.menu_active_bg, activeforeground=self.menu_active_fg)
        menu.add_command(label=f"取消 {name} 的打卡",
                         command=lambda: self.cancel_attendance(name))
        menu.post(event.x_root, event.y_root)

    def cancel_attendance(self, name):
        """取消学生的打卡"""
        if not self.verify_admin_password(f"取消 {name} 的打卡"):
            return

        if not self.student_data[name]['first_time']:
            self.status_var.set(f"{name} 尚未打卡，无法取消")
            return

        # 确认对话框
        confirm = messagebox.askyesno("确认取消打卡",
                                      f"确定要取消 {name} 的打卡吗？")
        if not confirm:
            return

        # 恢复学生数据
        self.student_data[name]['first_time'] = None
        self.student_data[name]['count'] = max(0, self.student_data[name]['count'] - 1)

        # 更新按钮颜色
        self.buttons[name].config(bg=self.button_bg, fg=self.button_fg)

        # 更新状态
        self.status_var.set(f"{name} 的打卡已取消")

        # 更新排名和统计
        self.update_ranking()
        self.update_stats()
        self.save_student_data()

        # 如果是客户端模式，同步到服务器
        if self.online_mode and self.server_status == "在线":
            self.sync_data_to_server()

    def update_stats(self):
        """更新统计信息"""
        total = len(self.student_data)
        punched = sum(1 for data in self.student_data.values() if data['first_time'])
        percentage = f"{punched / total * 100:.1f}%" if total > 0 else "0.0%"
        self.total_var.set(f"总人数: {total}")
        self.punched_var.set(f"已打卡: {punched} ({percentage})")

    def mark_attendance(self, name):
        """记录学生打卡"""
        current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

        # 如果已经打卡，则提示
        if self.student_data[name]['first_time']:
            self.status_var.set(f"{name} 已打卡，右键点击可取消打卡")
            return

        # 记录打卡
        self.student_data[name]['first_time'] = current_time
        self.student_data[name]['count'] += 1

        # 添加历史记录
        self.student_data[name]['history'].append(current_time)

        # 更新按钮颜色
        self.buttons[name].config(bg=self.active_button_bg, fg=self.active_button_fg)

        # 更新状态
        self.status_var.set(f"{name} 打卡成功! 打卡时间: {current_time}")

        # 更新排名和统计
        self.update_ranking()
        self.update_stats()
        self.save_student_data()

        # 如果是客户端模式，同步到服务器
        if self.online_mode and self.server_status == "在线":
            self.sync_data_to_server()

    def update_ranking(self):
        """更新排名显示 - 按首次打卡时间排序"""
        # 清空当前排名
        for item in self.ranking_tree.get_children():
            self.ranking_tree.delete(item)

        # 获取已打卡学生数据（按首次打卡时间排序）
        punched_students = []
        for name, data in self.student_data.items():
            if data['first_time']:
                punched_students.append((name, data['first_time']))

        # 按打卡时间升序排序（最早打卡的排最前）
        punched_students.sort(key=lambda x: x[1])

        # 添加排名数据
        for rank, (name, first_time) in enumerate(punched_students, start=1):
            # 设置前三名特殊颜色
            tags = ()
            if rank == 1:
                tags = ('first',)
            elif rank == 2:
                tags = ('second',)
            elif rank == 3:
                tags = ('third',)

            # 格式化时间
            formatted_time = first_time
            if len(first_time) > 16:  # 简化显示
                formatted_time = first_time[11:16]  # 只显示小时和分钟

            self.ranking_tree.insert("", "end", values=(rank, name, formatted_time), tags=tags)

        # 配置前三名样式
        self.ranking_tree.tag_configure('first', background='#ffd700', foreground='black')  # 金色
        self.ranking_tree.tag_configure('second', background='#c0c0c0', foreground='black')  # 银色
        self.ranking_tree.tag_configure('third', background='#cd7f32', foreground='white')  # 铜色

    def show_about(self):
        """显示关于对话框"""
        about_window = tk.Toplevel(self.window)
        about_window.title("关于")
        about_window.geometry("400x300")
        about_window.resizable(False, False)
        about_window.transient(self.window)
        about_window.grab_set()
        about_window.config(bg=self.bg_color)

        # 程序图标或标题
        tk.Label(about_window, text="打卡系统", font=("楷体", 18, 'bold'),
                 bg=self.bg_color, fg=self.fg_color).pack(pady=20)

        # 版本信息
        tk.Label(about_window, text=f"版本: {version}", font=("楷体", 12),
                 bg=self.bg_color, fg=self.fg_color).pack(pady=5)

        # 开发者信息
        tk.Label(about_window, text="开发者: 刘宇晨", font=("楷体", 12),
                 bg=self.bg_color, fg=self.fg_color).pack(pady=5)

        # 联系信息
        tk.Label(about_window, text="联系邮箱: liuyuchen032901@outlook.com", font=("楷体", 12),
                 bg=self.bg_color, fg=self.fg_color).pack(pady=5)

        # 版权信息
        tk.Label(about_window, text="© 2026 版权所有", font=("楷体", 10),
                 bg=self.bg_color, fg=self.fg_color).pack(side=tk.BOTTOM, pady=20)

        # 关闭按钮
        close_btn = tk.Button(about_window, text="关闭", command=about_window.destroy,
                              bg=self.button_bg, fg=self.button_fg, width=10,
                              relief=tk.FLAT, font=("楷体", 9))
        close_btn.pack(pady=20)
        close_btn.bind("<Enter>", lambda e: close_btn.config(bg=self.button_hover))
        close_btn.bind("<Leave>", lambda e: close_btn.config(bg=self.button_bg))


def gat_file():
    try:
        url = 'https://jay.615mc.cn/daikaconnect.txt'
        response = requests.get(url, verify=False, timeout=10)
        if response.status_code == 200:
            if response.text.strip() == 'true':
                return True, response.text
            else:
                return False, response.text
        else:
            return False, None
    except Exception as e:
        print(f"连接中心服务器失败: {e}")
        return False, None


if __name__ == '__main__':
    try:
        authorized, msg = gat_file()
        if not authorized:
            tk.messagebox.showerror('错误', '未授权')
            sys.exit(1)
    except Exception as e:
        tk.messagebox.showerror('错误', f'无法连接到中心监管服务器: {str(e)}')
        sys.exit(1)

    if len(sys.argv) > 1:
        impfile = True
        file_path = sys.argv[1]

    window = tk.Tk()
    app = AttendanceApp(window)

    def exit_action(icon, item):
        icon.stop()
        window.quit()

    def show_window(icon, item):
        icon.stop()
        window.deiconify()
        window.state('normal')
        window.focus_force()

    def hide_window(icon, item):
        window.withdraw()

    def create_system_tray():
        try:
            if os.path.exists("icon.ico"):
                image = Image.open("icon.ico")
                menu = (
                    MenuItem('显示窗口', show_window),
                    MenuItem('隐藏窗口', hide_window),
                    MenuItem('退出', exit_action)
                )
                icon = pystray.Icon("attendance_system", image, "打卡系统", menu)
                return icon
        except Exception as e:
            print(f"系统托盘创建失败: {e}")
        return None

    tray_icon = create_system_tray()
    
    def run_tray():
        if tray_icon:
            tray_icon.run()

    tray_thread = threading.Thread(target=run_tray, daemon=True)
    tray_thread.start()

    # 确保窗口在初始化后显示在任务栏
    window.after(100, app.ensure_taskbar_visibility)
    
    window.mainloop()