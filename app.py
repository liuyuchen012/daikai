import os
import tkinter as tk
from tkinter import ttk, messagebox
import time
from configparser import ConfigParser
from config import *


# 版权所有(c) 2025年7月14日
# 刘宇晨保留所有权利
# 开发者: 刘宇晨
# 联系邮箱: liuyuchen032901@outlook.com
# 开发环境: Python 3.11
# 系统环境: Windows 10
# 软件环境: PyCharm 2025.1.3
# 版本信息: 1.3
# 功能描述: 天津港保税区空港学校2018级10班数学打卡程序
# 更新日志:
#   2025年7月14日，版本1.0，程序初版
#   2025年7月15日，版本1.1，添加打卡排名功能
#   2025年7月16日，版本1.2，按打卡时间排名，未打卡学生不计入排名
#   2025年7月17日，版本1.3，添加取消打卡功能
conf = ConfigParser()

conf.read('config.ini',encoding='UTF-8')
nj = conf['config']['nj']
class_id = conf['config']['class_id']
z = int(conf['config']['z'])
l = int(conf['config']['l'])
km = conf['config']['km']
school = conf['config']['school']
class AttendanceApp:
    def __init__(self, window):
        self.window = window
        self.window.title(f'{school}{nj}年{class_id}班{km}打卡 v1.3离线版 作者: 刘宇晨')
        self.window.geometry("1200x700")
        self.window.config(bg='white')

        # 尝试加载图标
        try:
            self.window.iconbitmap('icon.ico')
        except:
            pass

        # 初始化数据结构
        self.student_data = {}
        self.load_student_data()

        # 创建界面
        self.create_widgets()

    def load_student_data(self):
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

        # 尝试加载历史数据
        if os.path.exists("attendance.dat"):
            with open("attendance.dat", "r", encoding="UTF-8") as f:
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
        with open("attendance.dat", "w", encoding="UTF-8") as f:
            for name, data in self.student_data.items():
                # 保存格式: name:count:first_time:history
                history_str = '|'.join(data["history"])
                f.write(f"{name}:{data['count']}:{data['first_time'] or ''}:{history_str}\n")

    def create_widgets(self):
        """创建界面组件"""
        # 创建主框架
        main_frame = tk.Frame(self.window, bg='white')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 左侧排名区域
        left_frame = tk.LabelFrame(main_frame, text="打卡排名", font=("Arial Bold", 14),
                                   bg='white', padx=10, pady=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 20))

        # 排名标题
        tk.Label(left_frame, text="打卡时间排名 (最早打卡)", font=("Arial Bold", 16),
                 bg='white').pack(pady=(0, 10))

        # 创建排名表格
        columns = ("排名", "姓名", "打卡时间")
        self.ranking_tree = ttk.Treeview(left_frame, columns=columns, show="headings", height=20)

        # 设置列
        self.ranking_tree.heading("排名", text="排名")
        self.ranking_tree.heading("姓名", text="姓名")
        self.ranking_tree.heading("打卡时间", text="打卡时间")

        self.ranking_tree.column("排名", width=50, anchor=tk.CENTER)
        self.ranking_tree.column("姓名", width=100, anchor=tk.CENTER)
        self.ranking_tree.column("打卡时间", width=150, anchor=tk.CENTER)

        self.ranking_tree.pack(fill=tk.BOTH, expand=True)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.ranking_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.ranking_tree.configure(yscrollcommand=scrollbar.set)

        # 右侧学生区域
        right_frame = tk.Frame(main_frame, bg='white')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 标题和统计信息
        header_frame = tk.Frame(right_frame, bg='white')
        header_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Label(header_frame,
                 text=f'{school}{nj}年{class_id}班{km}打卡',
                 font=("Arial Bold", 25),
                 bg='white').pack(pady=(0, 10))

        # 统计信息
        self.stats_frame = tk.Frame(header_frame, bg='white')
        self.stats_frame.pack(fill=tk.X)

        self.total_var = tk.StringVar()
        self.punched_var = tk.StringVar()
        self.update_stats()

        tk.Label(self.stats_frame, textvariable=self.total_var, font=("Arial", 12),
                 bg='white').pack(side=tk.LEFT, padx=10)
        tk.Label(self.stats_frame, textvariable=self.punched_var, font=("Arial", 12),
                 bg='white').pack(side=tk.LEFT, padx=10)

        # 创建按钮网格容器
        button_frame = tk.Frame(right_frame, bg='white')
        button_frame.pack(fill=tk.BOTH, expand=True)

        # 创建按钮网格
        self.buttons = {}
        names = list(self.student_data.keys())

        # 计算行列数
        total_students = len(names)
        rows = z if z > 0 else 6
        cols = l if l > 0 else 6

        for i in range(rows):
            for j in range(cols):
                idx = i * cols + j
                if idx < total_students:
                    name = names[idx]
                    # 根据打卡状态设置按钮颜色
                    bg_color = 'lightgreen' if self.student_data[name]['first_time'] else 'white'
                    btn = tk.Button(
                        button_frame,
                        text=name,
                        font=("Arial", 14),
                        bg=bg_color,
                        height=2,
                        width=15,
                        command=lambda n=name: self.mark_attendance(n)
                    )
                    # 绑定右键菜单
                    btn.bind("<Button-3>", lambda event, n=name: self.show_context_menu(event, n))
                    btn.grid(row=i, column=j, padx=5, pady=5, sticky="nsew")
                    self.buttons[name] = btn

        # 设置网格权重，使按钮均匀分布
        for i in range(rows):
            button_frame.grid_rowconfigure(i, weight=1)
        for j in range(cols):
            button_frame.grid_columnconfigure(j, weight=1)

        # 底部状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪 - 左键点击学生姓名进行打卡，右键点击取消打卡")
        status_bar = tk.Label(self.window, textvariable=self.status_var,
                              bd=1, relief=tk.SUNKEN, anchor=tk.W, bg='lightgray')
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 初始化排名
        self.update_ranking()

    def show_context_menu(self, event, name):
        """显示右键菜单（取消打卡）"""
        if not self.student_data[name]['first_time']:
            return  # 未打卡的学生不需要取消打卡菜单

        menu = tk.Menu(self.window, tearoff=0)
        menu.add_command(label=f"取消 {name} 的打卡",
                         command=lambda: self.cancel_attendance(name))
        menu.post(event.x_root, event.y_root)

    def cancel_attendance(self, name):
        """取消学生的打卡"""
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
        self.buttons[name].config(bg='white')

        # 更新状态
        self.status_var.set(f"{name} 的打卡已取消")

        # 更新排名和统计
        self.update_ranking()
        self.update_stats()
        self.save_student_data()

    def update_stats(self):
        """更新统计信息"""
        total = len(self.student_data)
        punched = sum(1 for data in self.student_data.values() if data['first_time'])
        self.total_var.set(f"总人数: {total}")
        self.punched_var.set(f"已打卡: {punched} ({punched / total * 100:.1f}%)")

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
        self.buttons[name].config(bg='lightgreen')

        # 更新状态
        self.status_var.set(f"{name} 打卡成功! 打卡时间: {current_time}")

        # 更新排名和统计
        self.update_ranking()
        self.update_stats()
        self.save_student_data()

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
        self.ranking_tree.tag_configure('first', background='gold', font=('Arial', 11, 'bold'))
        self.ranking_tree.tag_configure('second', background='silver', font=('Arial', 11))
        self.ranking_tree.tag_configure('third', background='#cd7f32', font=('Arial', 11))


if __name__ == '__main__':
    window = tk.Tk()
    app = AttendanceApp(window)
    window.mainloop()