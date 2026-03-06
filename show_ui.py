# show_ui.py
# 采用 old.py 的 UI 风格，移植管理员设置功能

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkFont
import ctypes
import time

# 导入 online 中的函数（需避免循环导入，这里使用局部导入）
# open_url_bz 和 check_version 将在需要时从 online 模块导入

# ---------- Windows API 常量 ----------
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_FRAMECHANGED = 0x0020
GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_CAPTION = 0x00C00000
WS_SYSMENU = 0x00080000
WS_THICKFRAME = 0x00040000
WS_EX_APPWINDOW = 0x00040000

# 窗口状态常量
SW_MAXIMIZE = 3  # ShowWindow 参数
SW_SHOWMAXIMIZED = 3  # 等同于 SW_MAXIMIZE
SW_RESTORE = 9
SW_SHOWNORMAL = 1

# 键盘相关常量
VK_LWIN = 0x5B
VK_UP = 0x26
KEYEVENTF_KEYUP = 0x0002

# DWM API 相关
S_OK = 0
DWMWA_EXTENDED_FRAME_BOUNDS = 9


# ---------- 窗口拖动辅助函数 ----------
def windowMove(widget, window):
    """使用 Windows API SetWindowPos 移动窗口，绕过 Tkinter 重绘"""

    class DragManager:
        def __init__(self, window):
            self.window = window
            self.dragging = False
            self.start_x = 0
            self.start_y = 0
            self.hwnd = None

        def start_drag(self, event):
            self.dragging = True
            self.start_x = event.x_root
            self.start_y = event.y_root

            # 获取窗口句柄和当前位置
            try:
                self.hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
                rect = ctypes.wintypes.RECT()
                ctypes.windll.user32.GetWindowRect(self.hwnd, ctypes.byref(rect))
                self.window_x = rect.left
                self.window_y = rect.top
            except:
                self.hwnd = None

            self.window.config(cursor="fleur")

        def do_drag(self, event):
            if not self.dragging or not self.hwnd:
                return

            # 使用 SetWindowPos 直接移动窗口，不触发 Tkinter 重绘
            dx = event.x_root - self.start_x
            dy = event.y_root - self.start_y
            new_x = self.window_x + dx
            new_y = self.window_y + dy

            # 使用 SetWindowPos 移动窗口
            # SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE = 0x0001 | 0x0004 | 0x0010 = 0x0015
            ctypes.windll.user32.SetWindowPos(
                self.hwnd,
                0,
                new_x,
                new_y,
                0, 0,
                0x0015  # SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE
            )

        def end_drag(self, event=None):
            if not self.dragging:
                return
            self.dragging = False
            self.window.config(cursor="")

            # 拖拽结束后同步一次 Tkinter 的 geometry
            if self.hwnd:
                try:
                    rect = ctypes.wintypes.RECT()
                    ctypes.windll.user32.GetWindowRect(self.hwnd, ctypes.byref(rect))
                    self.window.geometry(f"+{rect.left}+{rect.top}")
                except:
                    pass

    drag_manager = DragManager(window)

    # 绑定到主容器
    widget.bind("<ButtonPress-1>", drag_manager.start_drag)
    widget.bind("<B1-Motion>", drag_manager.do_drag)
    widget.bind("<ButtonRelease-1>", drag_manager.end_drag)

    # 同时也绑定到窗口级别，确保在任何地方释放都能结束拖拽
    window.bind("<ButtonRelease-1>", lambda e: drag_manager.end_drag())


# ---------- 设置窗口样式（保留任务栏图标）----------
def setup_window_style(window):
    """移除系统标题栏，但保留 WS_THICKFRAME 以支持最大化"""
    try:
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())

        # 获取当前样式
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
        ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)

        # 保留 WS_THICKFRAME 和 WS_SYSMENU，移除 WS_CAPTION
        # 这是让 Win11 正确识别可最大化窗口的关键
        style = (style & ~WS_CAPTION) | WS_THICKFRAME | WS_SYSMENU

        # 确保有 WS_EX_APPWINDOW 以显示在任务栏
        ex_style = ex_style | WS_EX_APPWINDOW

        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_STYLE, style)
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex_style)

        # 更新窗口位置 - 不使用 SWP_FRAMECHANGED，避免触发不必要的重绘
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                          SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)

        return True
    except Exception as e:
        print(f"设置窗口样式失败: {e}")
        window.overrideredirect(True)
        return False


# ---------- 主题应用 ----------
def theme(self, tk):
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


# ================== 远程设置对话框 ==================
def show_settings_remot(self, tk):
    """远程服务器设置对话框（完全使用 ttk 组件）"""
    from tkinter import ttk
    settings_window = tk.Toplevel(self.window)
    settings_window.title("远程服务器设置")
    settings_window.geometry("450x320")
    settings_window.resizable(False, False)
    settings_window.transient(self.window)
    settings_window.grab_set()

    # 获取当前主题背景
    bg = self.bg_color if self.dark_mode else '#f3f3f3'
    settings_window.config(bg=bg)

    main_frame = ttk.Frame(settings_window, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 客户端模式
    online_frame = ttk.Frame(main_frame)
    online_frame.pack(fill=tk.X, pady=(0, 15))
    ttk.Label(online_frame, text="客户端模式:").pack(side=tk.LEFT)
    self.online_var = tk.BooleanVar(value=self.online_mode)
    ttk.Checkbutton(online_frame, variable=self.online_var,
                    command=lambda: setattr(self, 'online_mode', self.online_var.get())).pack(side=tk.LEFT, padx=10)

    # 服务器模式
    server_frame = ttk.Frame(main_frame)
    server_frame.pack(fill=tk.X, pady=(0, 15))
    ttk.Label(server_frame, text="服务器模式:").pack(side=tk.LEFT)
    self.server_var = tk.BooleanVar(value=self.bd_online)
    ttk.Checkbutton(server_frame, variable=self.server_var,
                    command=lambda: setattr(self, 'bd_online', self.server_var.get())).pack(side=tk.LEFT, padx=10)

    # 服务器地址
    ip_frame = ttk.Frame(main_frame)
    ip_frame.pack(fill=tk.X, pady=(0, 15))
    ttk.Label(ip_frame, text="服务器地址:").pack(side=tk.LEFT)
    self.ip_entry = ttk.Entry(ip_frame, width=25)
    self.ip_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
    self.ip_entry.insert(0, self.online_ip)

    # 服务器端口
    port_frame = ttk.Frame(main_frame)
    port_frame.pack(fill=tk.X, pady=(0, 20))
    ttk.Label(port_frame, text="服务器端口:").pack(side=tk.LEFT)
    self.port_entry = ttk.Entry(port_frame, width=10)
    self.port_entry.pack(side=tk.LEFT, padx=10)
    self.port_entry.insert(0, str(self.server_port))

    # 服务器密码
    password_frame = ttk.Frame(main_frame)
    password_frame.pack(fill=tk.X, pady=(0, 20))
    ttk.Label(password_frame, text="服务器密码:").pack(side=tk.LEFT)
    self.password_entry = ttk.Entry(password_frame, width=20, show="*")
    self.password_entry.pack(side=tk.LEFT, padx=10)
    self.password_entry.insert(0, self.server_password)

    # 按钮
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(10, 0))
    ttk.Button(button_frame, text="保存设置",
               command=lambda: self.save_remote_settings(settings_window)).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame, text="取消",
               command=settings_window.destroy).pack(side=tk.RIGHT, padx=5)


# ---------- 创建自定义标题栏 ----------
def create_custom_titlebar(self, tk):
    """创建白色统一的标题栏（包含菜单）"""
    # 白色主题
    titlebar_bg = '#ffffff'
    titlebar_fg = '#333333'
    titlebar_hover = '#f3f3f3'
    titlebar_active = '#e0e0e0'

    # 统一的标题栏容器
    self.titlebar = tk.Frame(self.window, bg=titlebar_bg, height=40, bd=0, highlightthickness=0)
    self.titlebar.pack(fill=tk.X)
    windowMove(self.titlebar, self.window)

    # 左侧：应用图标 + 标题 + 菜单
    left_section = tk.Frame(self.titlebar, bg=titlebar_bg)
    left_section.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # 绑定拖拽到 left_section（占据标题栏大部分空间）
    windowMove(left_section, self.window)

    # 应用图标（蓝色圆点）
    app_icon = tk.Frame(left_section, bg='#4285f4', width=14, height=14)
    app_icon.pack(side=tk.LEFT, padx=(12, 8), pady=13)
    app_icon.pack_propagate(False)

    # 项目名称
    title_text = f'{self.school}{self.nj}年{self.class_id}班{self.km}打卡'
    self.title_label = tk.Label(
        left_section,
        text=title_text,
        bg=titlebar_bg,
        fg=titlebar_fg,
        font=("微软雅黑", 9)
    )
    self.title_label.pack(side=tk.LEFT, pady=11)

    # 垂直分隔符
    separator = tk.Frame(left_section, bg='#e0e0e0', width=1)
    separator.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0), pady=8)

    # 菜单区域
    menu_container = tk.Frame(left_section, bg=titlebar_bg)
    menu_container.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0))

    # 从 online 导入函数（避免循环导入）
    from online import open_url_bz, check_version

    menu_items = [
        ("文件", ["导出打卡数据", "导入打卡数据", "清空打卡记录", "退出"]),
        ("远程", ["远程服务器设置", "检查服务器状态", "从服务器加载数据", "同步数据到服务器"]),
        ("设置", ["管理员设置"]),
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
        "Github": open_url_bz,
        "检查版本列表": check_version,
        "关于": self.show_about
    }

    # 创建菜单项
    for item, subitems in menu_items:
        btn = tk.Menubutton(
            menu_container,
            text=item,
            bg=titlebar_bg,
            fg=titlebar_fg,
            activebackground=titlebar_active,
            activeforeground='#333333',
            bd=0,
            padx=10,
            pady=8,
            font=("微软雅黑", 9),
            relief=tk.FLAT
        )
        btn.pack(side=tk.LEFT)

        # 悬停效果
        def on_enter(e, b=btn):
            b.config(bg=titlebar_hover)

        def on_leave(e, b=btn):
            b.config(bg=titlebar_bg)

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

        submenu = tk.Menu(btn, tearoff=0, bg='white', fg='#333333',
                          activebackground='#e0e0e0', activeforeground='#333333',
                          font=("微软雅黑", 9), bd=1, relief=tk.SOLID)
        btn.config(menu=submenu)

        for subitem in subitems:
            if subitem == "-":
                submenu.add_separator()
            else:
                submenu.add_command(label=subitem, command=commands.get(subitem))

    # 右侧控制按钮
    controls = tk.Frame(self.titlebar, bg=titlebar_bg)
    controls.pack(side=tk.RIGHT, fill=tk.Y)

    # 定义控制按钮
    def create_control_btn(parent, text, command, close_btn=False):
        btn = tk.Button(
            parent,
            text=text,
            bg=titlebar_bg,
            fg=titlebar_fg,
            width=6,
            height=2,
            bd=0,
            relief=tk.FLAT,
            font=("微软雅黑", 11),
            command=command
        )
        btn.pack(side=tk.LEFT)

        if close_btn:
            btn.bind("<Enter>", lambda e: btn.config(bg='#e81123', fg='white'))
            btn.bind("<Leave>", lambda e: btn.config(bg=titlebar_bg, fg=titlebar_fg))
        else:
            btn.bind("<Enter>", lambda e: btn.config(bg=titlebar_hover))
            btn.bind("<Leave>", lambda e: btn.config(bg=titlebar_bg))

        return btn

    # 最小化按钮
    self.min_btn = create_control_btn(controls, "─", self.minimize_window)

    # 最大化按钮
    self.max_btn = create_control_btn(controls, "☐", self.toggle_maximize)

    # 关闭按钮
    self.close_btn = create_control_btn(controls, "✕", self.window.quit, close_btn=True)


# ---------- 最小化/最大化 ----------
def minimize_window(self):
    self.window.iconify()


def toggle_maximize(self):
    if hasattr(self, 'maximized') and self.maximized:
        # 恢复窗口 - 模拟 Win+下键
        try:
            hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
            if hwnd:
                # 确保窗口激活
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                # 模拟 Win+下键 (恢复)
                ctypes.windll.user32.keybd_event(VK_LWIN, 0, 0, 0)  # Win 按下
                ctypes.windll.user32.keybd_event(VK_UP, 0, 0, 0)  # 上键按下
                ctypes.windll.user32.keybd_event(VK_UP, 0, KEYEVENTF_KEYUP, 0)  # 上键释放
                ctypes.windll.user32.keybd_event(VK_UP, 0, 0, 0)  # 再按上键
                ctypes.windll.user32.keybd_event(VK_UP, 0, KEYEVENTF_KEYUP, 0)  # 上键释放
                ctypes.windll.user32.keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)  # Win 释放

                self.max_btn.config(text="□")
                self.maximized = False
        except Exception as e:
            print(f"恢复窗口失败: {e}")
            # 回退到 Tkinter 方法
            self.window.geometry(self.prev_geometry)
            self.max_btn.config(text="□")
            self.maximized = False
    else:
        # 最大化窗口 - 模拟 Win+上键
        self.prev_geometry = self.window.geometry()
        try:
            hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
            if hwnd:
                # 确保窗口激活
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                # 模拟 Win+上键
                ctypes.windll.user32.keybd_event(VK_LWIN, 0, 0, 0)  # Win 按下
                ctypes.windll.user32.keybd_event(VK_UP, 0, 0, 0)  # 上键按下
                ctypes.windll.user32.keybd_event(VK_UP, 0, KEYEVENTF_KEYUP, 0)  # 上键释放
                ctypes.windll.user32.keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)  # Win 释放

                self.max_btn.config(text="❐")
                self.maximized = True
        except Exception as e:
            print(f"最大化窗口失败: {e}")
            # 回退到 Tkinter 方法
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            self.window.geometry(f"{screen_width}x{screen_height}+0+0")
            self.max_btn.config(text="❐")
            self.maximized = True


# ---------- 创建菜单栏 ----------
def create_menu(self, tk):
    """菜单已集成到标题栏中，此函数保留为空以兼容"""
    pass


# ---------- 更新 Treeview 样式 ----------
def update_treeview_style(self):
    """应用 Treeview 样式"""
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


# ---------- 构建主界面 ----------
def create_widgets(self, tk, ttk, school, nj, class_id, km, z, l):
    """创建主界面（左右分区）"""
    # 主容器
    self.main_container = tk.Frame(self.window, bg=self.bg_color)
    self.main_container.pack(fill=tk.BOTH, expand=True, pady=(30, 0))

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

    # 右侧区域
    right_frame = tk.Frame(main_frame, bg=self.bg_color)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    # 标题和状态
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
    self.online_status_label.pack(side=tk.LEFT, padx=10)

    self.last_check_label = tk.Label(
        status_frame,
        text=f"最后检查: {self.server_last_check}",
        font=("楷体", 9),
        bg=self.bg_color,
        fg=self.status_fg_color
    )
    self.last_check_label.pack(side=tk.RIGHT, padx=10)

    # 统计
    self.stats_frame = tk.Frame(header_frame, bg=self.bg_color)
    self.stats_frame.pack(fill=tk.X)
    self.total_var = tk.StringVar()
    self.punched_var = tk.StringVar()
    self.update_stats()  # 需要在 online 中定义

    tk.Label(self.stats_frame, textvariable=self.total_var, font=("楷体", 10),
             bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT, padx=10)
    tk.Label(self.stats_frame, textvariable=self.punched_var, font=("楷体", 10),
             bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT, padx=10)

    # 按钮网格
    button_container = tk.Frame(right_frame, bg=self.frame_bg, bd=1, relief=tk.FLAT,
                                highlightbackground=self.frame_border)
    button_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    button_frame = tk.Frame(button_container, bg=self.frame_bg)
    button_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    self.buttons = {}
    names = list(self.student_data.keys())
    total_students = len(names)
    rows = z if z > 0 else 6
    cols = l if l > 0 else 6

    for i in range(rows):
        for j in range(cols):
            idx = i * cols + j
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

    for i in range(rows):
        button_frame.grid_rowconfigure(i, weight=1)
    for j in range(cols):
        button_frame.grid_columnconfigure(j, weight=1)

    # 状态栏
    self.status_var = tk.StringVar()
    self.status_var.set("就绪 - 左键点击学生姓名进行打卡，右键点击取消打卡")
    status_bar = tk.Label(self.window, textvariable=self.status_var,
                          bd=1, relief=tk.SUNKEN, anchor=tk.W,
                          bg=self.status_bg, fg=self.status_fg_color,
                          font=("楷体", 9))
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # 初始化排名
    self.update_ranking()


# ---------- 按钮悬停效果 ----------
def on_button_enter(self, event, button, name):
    if self.student_data[name]['first_time']:
        button.config(bg=self.active_button_bg, fg=self.active_button_fg)
    else:
        button.config(bg=self.button_hover)


def on_button_leave(self, event, button, name):
    if self.student_data[name]['first_time']:
        button.config(bg=self.active_button_bg, fg=self.active_button_fg)
    else:
        button.config(bg=self.button_bg)


# ---------- 右键菜单 ----------
def show_context_menu(self, event, name):
    if not self.student_data[name]['first_time']:
        return
    menu = tk.Menu(self.window, tearoff=0, bg=self.menu_bg, fg=self.menu_fg,
                   activebackground=self.menu_active_bg, activeforeground=self.menu_active_fg)
    menu.add_command(label=f"取消 {name} 的打卡",
                     command=lambda: self.cancel_attendance(name))
    menu.post(event.x_root, event.y_root)


# ---------- 取消打卡 ----------
def cancel_attendance(self, name):
    if not self.student_data[name]['first_time']:
        self.status_var.set(f"{name} 尚未打卡，无法取消")
        return

    confirm = messagebox.askyesno("确认取消打卡",
                                  f"确定要取消 {name} 的打卡吗？")
    if not confirm:
        return

    self.student_data[name]['first_time'] = None
    self.student_data[name]['count'] = max(0, self.student_data[name]['count'] - 1)
    self.buttons[name].config(bg=self.button_bg, fg=self.button_fg)
    self.status_var.set(f"{name} 的打卡已取消")
    self.update_ranking()
    self.update_stats()
    self.save_student_data()

    if self.online_mode and self.server_status == "在线":
        self.sync_data_to_server()


# ---------- 导入数据 ----------
def import_data(self):
    """从 CSV 导入打卡数据"""
    try:
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        with open(file_path, 'r', encoding='utf-8') as f:
            next(f)  # 跳过表头
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

        self.student_data = new_data
        self.save_student_data()
        self.update_ui_from_data()
        self.status_var.set(f"数据已从 {file_path} 导入")
        messagebox.showinfo("导入成功", "打卡数据已成功导入！")

        if self.online_mode and self.server_status == "在线":
            self.sync_data_to_server()
    except Exception as e:
        messagebox.showerror("导入失败", f"导入数据时出错: {str(e)}")


# ---------- 更新 UI 从数据 ----------
def update_ui_from_data(self):
    try:
        # 检查窗口是否已关闭
        if not hasattr(self, 'window') or not self.window.winfo_exists():
            return

        # 检查按钮是否存在
        if not hasattr(self, 'buttons') or not self.buttons:
            return

        for name, data in self.student_data.items():
            if name in self.buttons:
                try:
                    button = self.buttons[name]
                    # 检查按钮是否仍然存在
                    if button.winfo_exists():
                        if data['first_time']:
                            button.config(bg=self.active_button_bg, fg=self.active_button_fg)
                        else:
                            button.config(bg=self.button_bg, fg=self.button_fg)
                except Exception:
                    continue

        # 更新统计和排名
        if hasattr(self, 'update_stats'):
            try:
                self.update_stats()
            except Exception:
                pass

        if hasattr(self, 'update_ranking'):
            try:
                self.update_ranking()
            except Exception:
                pass
    except Exception:
        pass


# ---------- 关于对话框 ----------
def show_about(self, tk):
    about_window = tk.Toplevel(self.window)
    about_window.title("关于")
    about_window.geometry("400x300")
    about_window.resizable(False, False)
    about_window.transient(self.window)
    about_window.grab_set()
    about_window.config(bg=self.bg_color)

    tk.Label(about_window, text="打卡系统", font=("楷体", 18, 'bold'),
             bg=self.bg_color, fg=self.fg_color).pack(pady=20)
    tk.Label(about_window, text=f"版本: {self.version}", font=("楷体", 12),
             bg=self.bg_color, fg=self.fg_color).pack(pady=5)
    tk.Label(about_window, text="开发者: 刘宇晨", font=("楷体", 12),
             bg=self.bg_color, fg=self.fg_color).pack(pady=5)
    tk.Label(about_window, text="联系邮箱: liuyuchen032901@outlook.com", font=("楷体", 12),
             bg=self.bg_color, fg=self.fg_color).pack(pady=5)
    tk.Label(about_window, text="© 2026 版权所有", font=("楷体", 10),
             bg=self.bg_color, fg=self.fg_color).pack(side=tk.BOTTOM, pady=20)

    close_btn = tk.Button(about_window, text="关闭", command=about_window.destroy,
                          bg=self.button_bg, fg=self.button_fg, width=10,
                          relief=tk.FLAT, font=("楷体", 9))
    close_btn.pack(pady=20)
    close_btn.bind("<Enter>", lambda e: close_btn.config(bg=self.button_hover))
    close_btn.bind("<Leave>", lambda e: close_btn.config(bg=self.button_bg))


def show_admin_settings(self, tk):
    if not self.verify_admin_password("访问管理员设置"):
        return

    settings_window = tk.Toplevel(self.window)
    settings_window.title("管理员设置")
    settings_window.geometry("600x500")
    settings_window.resizable(False, False)
    settings_window.transient(self.window)
    settings_window.grab_set()
    settings_window.config(bg=self.bg_color)

    title_font = tk.font.Font(family="楷体", size=14, weight='bold')
    tk.Label(settings_window, text="管理员设置", font=title_font,
             bg=self.bg_color, fg=self.fg_color).pack(pady=15)

    notebook = ttk.Notebook(settings_window)
    notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    # 系统设置选项卡
    system_frame = ttk.Frame(notebook, padding=10)
    notebook.add(system_frame, text="系统设置")

    # 密码设置选项卡
    password_frame = ttk.Frame(notebook, padding=10)
    notebook.add(password_frame, text="密码设置")

    # 系统设置内容
    tk.Label(system_frame, text="学校名称:", width=15).grid(row=0, column=0, pady=5, sticky=tk.W)
    self.school_entry = tk.Entry(system_frame)
    self.school_entry.grid(row=0, column=1, pady=5, sticky=tk.W + tk.E)
    self.school_entry.insert(0, self.school)

    tk.Label(system_frame, text="年级:", width=15).grid(row=1, column=0, pady=5, sticky=tk.W)
    self.nj_entry = tk.Entry(system_frame)
    self.nj_entry.grid(row=1, column=1, pady=5, sticky=tk.W + tk.E)
    self.nj_entry.insert(0, self.nj)

    tk.Label(system_frame, text="班级:", width=15).grid(row=2, column=0, pady=5, sticky=tk.W)
    self.class_entry = tk.Entry(system_frame)
    self.class_entry.grid(row=2, column=1, pady=5, sticky=tk.W + tk.E)
    self.class_entry.insert(0, self.class_id)

    tk.Label(system_frame, text="课程名称:", width=15).grid(row=3, column=0, pady=5, sticky=tk.W)
    self.km_entry = tk.Entry(system_frame)
    self.km_entry.grid(row=3, column=1, pady=5, sticky=tk.W + tk.E)
    self.km_entry.insert(0, self.km)

    tk.Label(system_frame, text="按钮行数:", width=15).grid(row=4, column=0, pady=5, sticky=tk.W)
    self.row_entry = tk.Entry(system_frame)
    self.row_entry.grid(row=4, column=1, pady=5, sticky=tk.W + tk.E)
    self.row_entry.insert(0, str(self.z))

    tk.Label(system_frame, text="按钮列数:", width=15).grid(row=5, column=0, pady=5, sticky=tk.W)
    self.col_entry = tk.Entry(system_frame)
    self.col_entry.grid(row=5, column=1, pady=5, sticky=tk.W + tk.E)
    self.col_entry.insert(0, str(self.l))

    # 密码设置内容
    status_text = "已设置管理员密码" if self.admin_password_hash else "未设置管理员密码"
    tk.Label(password_frame, text=f"当前状态: {status_text}",
             font=("楷体", 9)).grid(row=0, column=0, columnspan=2, pady=10, sticky=tk.W)

    tk.Label(password_frame, text="新密码:", font=("楷体", 9)).grid(row=1, column=0, pady=5, sticky=tk.W)
    self.new_password_entry = tk.Entry(password_frame, show="*")
    self.new_password_entry.grid(row=1, column=1, pady=5, sticky=tk.W + tk.E)

    tk.Label(password_frame, text="确认密码:", font=("楷体", 9)).grid(row=2, column=0, pady=5, sticky=tk.W)
    self.confirm_password_entry = tk.Entry(password_frame, show="*")
    self.confirm_password_entry.grid(row=2, column=1, pady=5, sticky=tk.W + tk.E)

    # 清除密码按钮
    clear_btn = tk.Button(password_frame, text="清除密码",
                          command=lambda: self.clear_admin_password(settings_window),
                          bg=self.button_bg, fg=self.button_fg)
    clear_btn.grid(row=3, column=0, columnspan=2, pady=20)

    # 保存和取消按钮
    button_frame = tk.Frame(settings_window, bg=self.bg_color)
    button_frame.pack(fill=tk.X, padx=20, pady=20)

    save_btn = tk.Button(button_frame, text="保存所有设置",
                         command=lambda: self.save_all_settings(settings_window),
                         width=15, bg=self.button_bg, fg=self.button_fg)
    save_btn.pack(side=tk.LEFT, padx=10)

    cancel_btn = tk.Button(button_frame, text="取消",
                           command=settings_window.destroy,
                           width=15, bg=self.button_bg, fg=self.button_fg)
    cancel_btn.pack(side=tk.RIGHT, padx=10)
