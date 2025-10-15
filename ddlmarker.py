import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os
import re
from datetime import datetime, timedelta
import threading
import time
from tkinter import scrolledtext
import winshell
from win32com.client import Dispatch

class DDLMarker:
    def __init__(self, root):
        self.root = root
        self.root.title("DDL Marker")
        self.root.geometry("650x700")
        self.root.attributes('-topmost', True)
        self.root.configure(bg="#f0f0f0")
        
        # 确保中文显示正常
        self.style = ttk.Style()
        
        # 设置默认提醒时间
        self.default_reminders = [
            (30, "minutes", "提前30分钟"),
            (1, "hours", "提前1小时"),
            (1, "days", "提前1天")
        ]
        
        # 开机自启动设置默认值
        self.startup_enabled = False
        
        # 数据存储
        self.data_file = "ddl_data.json"
        self.load_data()
        
        # 创建界面
        self.create_widgets()
        
        # 启动提醒检查线程
        self.stop_event = threading.Event()
        self.reminder_thread = threading.Thread(target=self.check_reminders, daemon=True)
        self.reminder_thread.start()
        
        # 窗口关闭时保存数据
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 初始化任务专属提醒设置（为现有任务添加默认设置）
        for task in self.tasks:
            if "reminders" not in task:
                task["reminders"] = []
    
    def load_data(self):
        """加载任务和标签数据"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.tasks = data.get("tasks", [])
                    self.tags = data.get("tags", ["学习", "工作", "生活", "其他"])
                    self.custom_reminders = data.get("reminders", self.default_reminders)
            except:
                self.tasks = []
                self.tags = ["学习", "工作", "生活", "其他"]
                self.custom_reminders = self.default_reminders
                self.startup_enabled = data.get("startup_enabled", False)
        else:
            self.tasks = []
            self.tags = ["学习", "工作", "生活", "其他"]
            self.custom_reminders = self.default_reminders
            self.startup_enabled = False
        
        # 解析和验证任务日期
        for task in self.tasks:
            if isinstance(task["due_date"], str):
                try:
                    task["due_date"] = datetime.fromisoformat(task["due_date"])
                except:
                    pass
            
            # 确保任务有reminders字段
            if "reminders" not in task:
                task["reminders"] = []
            
            # 转换reminders中的时间单位为英文（如果是中文的话）
            updated_reminders = []
            for reminder in task["reminders"]:
                if len(reminder) == 3:
                    value, unit, text = reminder
                    # 转换中文单位
                    if unit == "分钟":
                        unit = "minutes"
                    elif unit == "小时":
                        unit = "hours"
                    elif unit == "天":
                        unit = "days"
                    updated_reminders.append((value, unit, text))
            task["reminders"] = updated_reminders
    
    def save_data(self):
        """保存任务和标签数据"""
        data = {
            "tasks": [],
            "tags": self.tags,
            "reminders": self.custom_reminders,
            "startup_enabled": self.startup_enabled
        }
        
        # 转换日期为字符串格式
        for task in self.tasks:
            task_copy = task.copy()
            task_copy["due_date"] = task_copy["due_date"].isoformat()
            if isinstance(task_copy["created_at"], datetime):
                task_copy["created_at"] = task_copy["created_at"].isoformat()
            data["tasks"].append(task_copy)
        
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    def create_widgets(self):
        """创建GUI组件"""
        # 标题栏
        title_frame = ttk.Frame(self.root, padding="10")
        title_frame.pack(fill=tk.X)
        
        ttk.Label(title_frame, text="DDL Marker", font=("SimHei", 16, "bold")).pack()
        
        # 输入区域
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.pack(fill=tk.X)
        
        ttk.Label(input_frame, text="任务:", font=("SimHei", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.task_entry = ttk.Entry(input_frame, width=30)
        self.task_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(input_frame, text="日期时间:", font=("SimHei", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.date_entry = ttk.Entry(input_frame, width=30)
        self.date_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        self.date_entry.insert(0, "示例: 2024-12-31 23:59 或 明天 14:30")
        
        ttk.Label(input_frame, text="标签:", font=('SimHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.tag_frame = ttk.Frame(input_frame)
        self.tag_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 创建标签单选按钮
        self.tag_var = tk.StringVar(value="其他")  # 默认选择"其他"
        self.create_tag_radiobuttons()
        
        # 添加任务按钮
        self.add_button = ttk.Button(input_frame, text="添加任务", command=self.add_task)
        self.add_button.grid(row=3, column=0, columnspan=2, pady=10)
        
        # 标签管理按钮
        self.tag_button = ttk.Button(input_frame, text="管理标签", command=self.manage_tags)
        self.tag_button.grid(row=3, column=1, pady=10, sticky=tk.E)
        
        # 提醒设置按钮
        self.reminder_button = ttk.Button(input_frame, text="提醒设置", command=self.manage_reminders)
        self.reminder_button.grid(row=3, column=2, pady=10)
        
        # 开机自启动选项
        settings_frame = ttk.Frame(self.root, padding="10")
        settings_frame.pack(fill=tk.X)
        
        self.startup_var = tk.BooleanVar(value=self.startup_enabled)
        self.startup_checkbox = ttk.Checkbutton(
            settings_frame, 
            text="开机自启动", 
            variable=self.startup_var, 
            command=self.toggle_startup
        )
        self.startup_checkbox.pack(anchor=tk.W)
        
        # 任务列表区域
        list_frame = ttk.Frame(self.root, padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建任务列表树状视图
        columns = ("completed", "task", "due_date", "tag", "actions")
        self.task_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        
        # 设置列宽和标题
        self.task_tree.column("completed", width=50, anchor=tk.CENTER)
        self.task_tree.column("task", width=120, anchor=tk.W)
        self.task_tree.column("due_date", width=100, anchor=tk.CENTER)
        self.task_tree.column("tag", width=60, anchor=tk.CENTER)
        self.task_tree.column("actions", width=60, anchor=tk.CENTER)
        
        self.task_tree.heading("completed", text="完成")
        self.task_tree.heading("task", text="任务")
        self.task_tree.heading("due_date", text="截止日期")
        self.task_tree.heading("tag", text="标签")
        self.task_tree.heading("actions", text="操作")
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        self.task_tree.configure(yscroll=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.task_tree.pack(fill=tk.BOTH, expand=True)
        
        # 更新任务列表
        self.update_task_list()
        
        # 批量导入按钮
        self.import_button = ttk.Button(list_frame, text="批量导入", command=self.import_tasks)
        self.import_button.pack(pady=5)
    
    def update_task_list(self):
        """更新任务列表显示"""
        # 清空现有列表
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        
        # 按截止日期排序（未完成的优先，然后按日期排序）
        sorted_tasks = sorted(
            self.tasks, 
            key=lambda x: (x["completed"], x["due_date"])
        )
        
        # 添加任务到列表
        for i, task in enumerate(sorted_tasks):
            status = "✓" if task["completed"] else ""
            due_date_str = task["due_date"].strftime("%Y-%m-%d %H:%M")
            
            # 根据时间设置不同颜色
            now = datetime.now()
            time_diff = task["due_date"] - now
            
            # 插入行
            item_id = self.task_tree.insert("", tk.END, values=(status, task["name"], due_date_str, task["tag"], "编辑"))
            
            # 设置行颜色
            if task["completed"]:
                self.task_tree.item(item_id, tags=("completed",))
            elif time_diff.total_seconds() <= 0:
                self.task_tree.item(item_id, tags=("overdue",))
            elif time_diff.total_seconds() <= 86400:  # 24小时内
                self.task_tree.item(item_id, tags=("urgent",))
            else:
                self.task_tree.item(item_id, tags=("default",))  # 普通状态任务应用默认样式
        
        # 配置标签样式
        self.task_tree.tag_configure("default", font=("SimHei", 12))  # 普通状态字体
        self.task_tree.tag_configure("completed", foreground="#888888", font=("SimHei", 12, "italic"))
        self.task_tree.tag_configure("overdue", foreground="#ff0000", font=("SimHei", 12, "bold"))
        self.task_tree.tag_configure("urgent", foreground="#ff8c00", font=("SimHei", 12))
        
        # 绑定双击事件
        self.task_tree.bind("<Double-1>", self.on_task_double_click)
    
    def create_tag_radiobuttons(self):
        """创建标签单选按钮"""
        # 清空现有的标签单选按钮
        for widget in self.tag_frame.winfo_children():
            widget.destroy()
        
        # 创建新的标签单选按钮
        for i, tag in enumerate(self.tags):
            ttk.Radiobutton(
                self.tag_frame, 
                text=tag, 
                variable=self.tag_var, 
                value=tag
            ).grid(row=i//3, column=i%3, sticky=tk.W, padx=5)
        
    
    def parse_date_time(self, date_str):
        """解析日期时间字符串"""
        # 清除示例文本
        if date_str == "示例: 2024-12-31 23:59 或 明天 14:30":
            return None
        
        # 尝试解析多种格式
        formats = [
            "%Y-%m-%d %H:%M",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
            "%m/%d/%Y %H:%M",
            "%m/%d/%Y",
            "%d/%m/%Y %H:%M",
            "%d/%m/%Y"
        ]
        
        # 处理相对日期
        today = datetime.now().date()
        now = datetime.now()
        
        # 明天
        if "明天" in date_str:
            tomorrow = today + timedelta(days=1)
            time_part = re.search(r"(\d{1,2}):(\d{2})", date_str)
            if time_part:
                hour, minute = map(int, time_part.groups())
                return datetime(tomorrow.year, tomorrow.month, tomorrow.day, hour, minute)
            else:
                return datetime(tomorrow.year, tomorrow.month, tomorrow.day, 23, 59)
        
        # 后天
        if "后天" in date_str:
            day_after_tomorrow = today + timedelta(days=2)
            time_part = re.search(r"(\d{1,2}):(\d{2})", date_str)
            if time_part:
                hour, minute = map(int, time_part.groups())
                return datetime(day_after_tomorrow.year, day_after_tomorrow.month, day_after_tomorrow.day, hour, minute)
            else:
                return datetime(day_after_tomorrow.year, day_after_tomorrow.month, day_after_tomorrow.day, 23, 59)
        
        # 数字+n天
        days_match = re.search(r"(\d+)\s*天", date_str)
        if days_match:
            days = int(days_match.group(1))
            target_date = today + timedelta(days=days)
            time_part = re.search(r"(\d{1,2}):(\d{2})", date_str)
            if time_part:
                hour, minute = map(int, time_part.groups())
                return datetime(target_date.year, target_date.month, target_date.day, hour, minute)
            else:
                return datetime(target_date.year, target_date.month, target_date.day, 23, 59)
        
        # 尝试标准格式
        for fmt in formats:
            try:
                date_time = datetime.strptime(date_str, fmt)
                # 如果没有时间部分，设置为23:59
                if len(fmt) <= 10:
                    date_time = datetime(date_time.year, date_time.month, date_time.day, 23, 59)
                return date_time
            except ValueError:
                continue
        
        return None
    
    def add_task(self):
        """添加新任务"""
        task_name = self.task_entry.get().strip()
        date_str = self.date_entry.get().strip()
        
        if not task_name:
            messagebox.showwarning("警告", "任务名称不能为空！")
            return
        
        # 解析日期
        due_date = self.parse_date_time(date_str)
        if not due_date:
            messagebox.showwarning("警告", "日期格式不正确！请使用类似'2024-12-31 23:59'或'明天 14:30'的格式。")
            return
        
        # 获取选中的标签
        selected_tag = self.tag_var.get()
        
        # 创建新任务，包含空的reminders列表
        new_task = {
            "name": task_name,
            "due_date": due_date,
            "tag": selected_tag,
            "completed": False,
            "created_at": datetime.now(),
            "reminders": []
        }
        
        self.tasks.append(new_task)
        self.save_data()
        self.update_task_list()
        
        # 清空输入框
        self.task_entry.delete(0, tk.END)
        self.date_entry.delete(0, tk.END)
    
    def on_task_double_click(self, event):
        """双击任务行事件"""
        item = self.task_tree.identify_row(event.y)
        if not item:
            return
        
        column = self.task_tree.identify_column(event.x)
        
        # 获取任务名称用于查找正确的任务索引
        task_name = self.task_tree.item(item, "values")[1]
        
        # 查找对应任务的索引
        task_index = None
        for i, task in enumerate(self.tasks):
            if task["name"] == task_name:
                task_index = i
                break
        
        if task_index is None:
            return
        
        # 如果点击的是完成列
        if column == "#1":
            self.tasks[task_index]["completed"] = not self.tasks[task_index]["completed"]
            self.save_data()
            self.update_task_list()
        # 如果点击的是操作列
        elif column == "#5":
            self.edit_task(task_index)
    
    def edit_task(self, task_index):
        """编辑任务"""
        task = self.tasks[task_index]
        
        # 创建编辑对话框
        edit_window = tk.Toplevel(self.root)
        edit_window.title("编辑任务")
        edit_window.geometry("450x250")
        edit_window.resizable(False, False)
        edit_window.configure(bg="#f0f0f0")
        
        # 任务名称
        ttk.Label(edit_window, text="任务名称:", font=("SimHei", 10)).grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        task_entry = ttk.Entry(edit_window, width=25)
        task_entry.grid(row=0, column=1, padx=10, pady=10)
        task_entry.insert(0, task["name"])
        
        # 截止日期
        ttk.Label(edit_window, text="截止日期:", font=("SimHei", 10)).grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        date_entry = ttk.Entry(edit_window, width=25)
        date_entry.grid(row=1, column=1, padx=10, pady=10)
        date_entry.insert(0, task["due_date"].strftime("%Y-%m-%d %H:%M"))
        
        # 标签
        ttk.Label(edit_window, text="标签:", font=("SimHei", 10)).grid(row=2, column=0, sticky=tk.W, padx=10, pady=10)
        tag_var = tk.StringVar(value=task["tag"])
        tag_frame = ttk.Frame(edit_window)
        tag_frame.grid(row=2, column=1, padx=10, pady=10, sticky=tk.W)
        
        for tag in self.tags:
            ttk.Radiobutton(tag_frame, text=tag, variable=tag_var, value=tag).pack(side=tk.LEFT, padx=5)
        
        # 提醒设置区域
        reminder_frame = ttk.LabelFrame(edit_window, text="提醒设置")
        reminder_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky=tk.W+tk.E)
        
        # 提醒设置状态
        reminder_var = tk.BooleanVar(value=len(task.get("reminders", [])) > 0)
        
        def toggle_reminder_settings():
            if reminder_var.get():
                reminder_settings_frame.pack(fill=tk.X, padx=10, pady=5)
            else:
                reminder_settings_frame.pack_forget()
                task["reminders"] = []
        
        ttk.Checkbutton(reminder_frame, text="使用自定义提醒设置", variable=reminder_var, command=toggle_reminder_settings).pack(anchor=tk.W, padx=5, pady=5)
        
        # 提醒设置选项
        reminder_settings_frame = ttk.Frame(reminder_frame)
        
        # 复制任务当前的提醒设置
        task_reminders = task.get("reminders", [])
        temp_reminders = task_reminders.copy()
        
        # 提醒列表显示
        reminder_list_frame = ttk.Frame(reminder_settings_frame)
        reminder_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        def update_reminder_list_display():
            # 清空现有列表
            for widget in reminder_list_frame.winfo_children():
                widget.destroy()
            
            # 显示当前提醒设置
            for i, (value, unit, text) in enumerate(temp_reminders):
                item_frame = ttk.Frame(reminder_list_frame)
                item_frame.pack(fill=tk.X, pady=2)
                
                ttk.Label(item_frame, text=text, width=20).pack(side=tk.LEFT, padx=5)
                ttk.Button(item_frame, text="删除", command=lambda idx=i: remove_reminder(idx)).pack(side=tk.RIGHT, padx=5)
        
        def remove_reminder(index):
            temp_reminders.pop(index)
            update_reminder_list_display()
        
        # 添加提醒选项
        add_reminder_frame = ttk.Frame(reminder_settings_frame)
        add_reminder_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(add_reminder_frame, text="提前量：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        reminder_value_var = tk.StringVar()
        reminder_value_entry = ttk.Entry(add_reminder_frame, textvariable=reminder_value_var, width=5)
        reminder_value_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        reminder_unit_var = tk.StringVar(value="分钟")
        reminder_unit_frame = ttk.Frame(add_reminder_frame)
        reminder_unit_frame.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        ttk.Radiobutton(reminder_unit_frame, text="分钟", variable=reminder_unit_var, value="分钟").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(reminder_unit_frame, text="小时", variable=reminder_unit_var, value="小时").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(reminder_unit_frame, text="天", variable=reminder_unit_var, value="天").pack(side=tk.LEFT, padx=5)
        
        def add_reminder():
            try:
                value = int(reminder_value_var.get())
                if value <= 0:
                    messagebox.showwarning("警告", "提前量必须大于0！")
                    return
                
                unit = reminder_unit_var.get()
                if unit == "分钟":
                    reminder = (value, "minutes", f"提前{value}分钟")
                elif unit == "小时":
                    reminder = (value, "hours", f"提前{value}小时")
                else:  # 天
                    reminder = (value, "days", f"提前{value}天")
                
                if reminder not in temp_reminders:
                    temp_reminders.append(reminder)
                    update_reminder_list_display()
                    reminder_value_var.set("")  # 清空输入框
                else:
                    messagebox.showinfo("提示", "该提醒已存在！")
            except ValueError:
                messagebox.showwarning("警告", "请输入有效的数字！")
        
        ttk.Button(add_reminder_frame, text="添加", command=add_reminder).grid(row=0, column=3, padx=5, pady=5)
        
        # 快捷提醒按钮
        quick_reminder_frame = ttk.LabelFrame(reminder_settings_frame, text="快捷添加")
        quick_reminder_frame.pack(fill=tk.X, pady=5)
        
        quick_buttons = [
            (10, "分钟", "提前10分钟"),
            (30, "分钟", "提前30分钟"),
            (1, "小时", "提前1小时"),
            (3, "小时", "提前3小时"),
            (1, "天", "提前1天")
        ]
        
        for value, unit, text in quick_buttons:
            def create_quick_button_cmd(v=value, u=unit, t=text):
                def cmd():
                    if u == "分钟":
                        reminder = (v, "minutes", t)
                    elif u == "小时":
                        reminder = (v, "hours", t)
                    else:  # 天
                        reminder = (v, "days", t)
                    
                    if reminder not in temp_reminders:
                        temp_reminders.append(reminder)
                        update_reminder_list_display()
                return cmd
            
            ttk.Button(quick_reminder_frame, text=text, command=create_quick_button_cmd()).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 初始化提醒设置显示
        if reminder_var.get():
            update_reminder_list_display()
            reminder_settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 按钮
        button_frame = ttk.Frame(edit_window)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        def save_changes():
            new_name = task_entry.get().strip()
            new_date_str = date_entry.get().strip()
            new_tag = tag_var.get()
            
            if not new_name:
                messagebox.showwarning("警告", "任务名称不能为空！")
                return
            
            new_date = self.parse_date_time(new_date_str)
            if not new_date:
                messagebox.showwarning("警告", "日期格式不正确！")
                return
            
            task["name"] = new_name
            task["due_date"] = new_date
            task["tag"] = new_tag
            
            # 更新提醒设置
            if reminder_var.get():
                task["reminders"] = temp_reminders
            else:
                task["reminders"] = []
            
            self.save_data()
            self.update_task_list()
            edit_window.destroy()
        
        def delete_task():
            if messagebox.askyesno("确认删除", "确定要删除这个任务吗？"):
                del self.tasks[task_index]
                self.save_data()
                self.update_task_list()
                edit_window.destroy()
        
        ttk.Button(button_frame, text="保存", command=save_changes).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="删除", command=delete_task).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=edit_window.destroy).pack(side=tk.LEFT, padx=10)
        
        # 居中显示
        edit_window.geometry("+")
        edit_window.update_idletasks()
        width = edit_window.winfo_width()
        height = edit_window.winfo_height()
        x = (edit_window.winfo_screenwidth() // 2) - (width // 2)
        y = (edit_window.winfo_screenheight() // 2) - (height // 2)
        edit_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def manage_tags(self):
        """管理标签"""
        tag_window = tk.Toplevel(self.root)
        tag_window.title("管理标签")
        tag_window.geometry("300x300")
        tag_window.resizable(False, False)
        
        # 标签列表
        tag_listbox = tk.Listbox(tag_window, width=30, height=10)
        tag_listbox.pack(pady=10)
        
        # 填充标签
        for tag in self.tags:
            tag_listbox.insert(tk.END, tag)
        
        # 按钮
        button_frame = ttk.Frame(tag_window)
        button_frame.pack(pady=10)
        
        def add_tag():
            new_tag = simpledialog.askstring("添加标签", "请输入新标签名称：")
            if new_tag and new_tag.strip():
                new_tag = new_tag.strip()
                if new_tag not in self.tags:
                    self.tags.append(new_tag)
                    tag_listbox.insert(tk.END, new_tag)
                    self.save_data()
                    self.create_tag_radiobuttons()  # 更新标签单选按钮
                else:
                    messagebox.showinfo("提示", "该标签已存在！")
        
        def delete_tag():
            selected = tag_listbox.curselection()
            if selected:
                tag_index = selected[0]
                tag_name = self.tags[tag_index]
                
                # 检查是否有任务使用该标签
                tasks_with_tag = [t for t in self.tasks if t["tag"] == tag_name]
                if tasks_with_tag and len(self.tags) > 1:
                    if messagebox.askyesno("确认删除", f"标签 '{tag_name}' 已被使用，确定要删除吗？\n删除后相关任务的标签将被设置为'其他'。"):
                        # 更新使用该标签的任务
                        for task in tasks_with_tag:
                            task["tag"] = "其他" if "其他" in self.tags else self.tags[0]
                        
                        # 删除标签
                        del self.tags[tag_index]
                        tag_listbox.delete(tag_index)
                        self.save_data()
                        self.create_tag_radiobuttons()  # 更新标签单选按钮
                        self.update_task_list()
                elif len(self.tags) <= 1:
                    messagebox.showwarning("警告", "至少需要保留一个标签！")
                else:
                    del self.tags[tag_index]
                    tag_listbox.delete(tag_index)
                    self.save_data()
        
        ttk.Button(button_frame, text="添加标签", command=add_tag).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="删除标签", command=delete_tag).pack(side=tk.LEFT, padx=10)
        
        # 居中显示
        tag_window.update_idletasks()
        width = tag_window.winfo_width()
        height = tag_window.winfo_height()
        x = (tag_window.winfo_screenwidth() // 2) - (width // 2)
        y = (tag_window.winfo_screenheight() // 2) - (height // 2)
        tag_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def toggle_startup(self):
        """切换开机自启动状态"""
        try:
            if self.startup_enabled:
                # 禁用开机自启动
                self._remove_startup()
                self.startup_enabled = False
                self.startup_checkbox.configure(state=tk.NORMAL)
                messagebox.showinfo("成功", "已取消开机自启动设置")
            else:
                # 启用开机自启动
                self._add_startup()
                self.startup_enabled = True
                self.startup_checkbox.configure(state=tk.NORMAL)
                messagebox.showinfo("成功", "已设置开机自启动")
        except Exception as e:
            messagebox.showerror("错误", f"设置开机自启动时出错: {str(e)}")
        
        # 保存设置
        self.save_data()
    
    def _add_startup(self):
        """添加开机自启动快捷方式"""
        # 获取启动文件夹路径
        startup_folder = winshell.startup()
        
        # 获取当前脚本或批处理文件的路径
        if os.path.exists("启动DDLMarker.bat"):
            # 使用批处理文件
            target_path = os.path.abspath("启动DDLMarker.bat")
        else:
            # 使用Python脚本
            target_path = os.path.abspath(__file__)
        
        # 创建快捷方式
        shortcut_path = os.path.join(startup_folder, "DDLMarker.lnk")
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.IconLocation = target_path
        shortcut.save()
    
    def _remove_startup(self):
        """移除开机自启动快捷方式"""
        startup_folder = winshell.startup()
        shortcut_path = os.path.join(startup_folder, "DDLMarker.lnk")
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
    
    def manage_reminders(self):
        """管理提醒设置"""
        reminder_window = tk.Toplevel(self.root)
        reminder_window.title("提醒设置")
        reminder_window.geometry("450x500")
        reminder_window.resizable(False, False)
        
        # 提醒列表
        reminder_frame = ttk.Frame(reminder_window)
        reminder_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 自定义提醒设置
        custom_reminder_frame = ttk.LabelFrame(reminder_window, text="自定义提醒")
        custom_reminder_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(custom_reminder_frame, text="提前量：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        value_var = tk.StringVar()
        value_entry = ttk.Entry(custom_reminder_frame, textvariable=value_var, width=5)
        value_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        unit_var = tk.StringVar(value="分钟")
        unit_frame = ttk.Frame(custom_reminder_frame)
        unit_frame.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        ttk.Radiobutton(unit_frame, text="分钟", variable=unit_var, value="分钟").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(unit_frame, text="小时", variable=unit_var, value="小时").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(unit_frame, text="天", variable=unit_var, value="天").pack(side=tk.LEFT, padx=5)
        
        def add_custom_reminder():
            try:
                value = int(value_var.get())
                if value <= 0:
                    messagebox.showwarning("警告", "提前量必须大于0！")
                    return
                
                unit = unit_var.get()
                if unit == "分钟":
                    reminder = (value, "minutes", f"提前{value}分钟")
                elif unit == "小时":
                    reminder = (value, "hours", f"提前{value}小时")
                else:  # 天
                    reminder = (value, "days", f"提前{value}天")
                
                if reminder not in self.custom_reminders:
                    self.custom_reminders.append(reminder)
                    update_reminder_list()
                    self.save_data()
                    value_var.set("")  # 清空输入框
                else:
                    messagebox.showinfo("提示", "该提醒已存在！")
            except ValueError:
                messagebox.showwarning("警告", "请输入有效的数字！")
        
        ttk.Button(custom_reminder_frame, text="添加", command=add_custom_reminder).grid(row=0, column=3, padx=5, pady=5)
        
        # 快捷提醒按钮
        quick_reminder_frame = ttk.LabelFrame(reminder_window, text="快捷添加提醒")
        quick_reminder_frame.pack(fill=tk.X, padx=10, pady=10)
        
        quick_buttons = [
            (1, "小时", "提前1小时"),
            (3, "小时", "提前3小时"),
            (6, "小时", "提前6小时"),
            (12, "小时", "提前12小时"),
            (1, "天", "提前1天"),
            (2, "天", "提前2天"),
            (7, "天", "提前7天")
        ]
        
        button_row = 0
        button_col = 0
        for value, unit, text in quick_buttons:
            def create_button_cmd(v=value, u=unit, t=text):
                return lambda: add_quick_reminder(v, u, t)
            
            ttk.Button(quick_reminder_frame, text=text, command=create_button_cmd()).grid(
                row=button_row, column=button_col, padx=5, pady=5
            )
            
            button_col += 1
            if button_col > 2:
                button_col = 0
                button_row += 1
        
        def add_quick_reminder(value, unit, text):
            if unit == "天":
                reminder = (value, "days", text)
            else:
                reminder = (value, "hours", text)
            
            if reminder not in self.custom_reminders:
                self.custom_reminders.append(reminder)
                update_reminder_list()
                self.save_data()
        
        def update_reminder_list():
            # 清空现有列表
            for widget in reminder_frame.winfo_children():
                widget.destroy()
            
            # 添加提醒项
            for i, (value, unit, text) in enumerate(self.custom_reminders):
                frame = ttk.Frame(reminder_frame)
                frame.pack(fill=tk.X, pady=5)
                
                ttk.Label(frame, text=text, width=20).pack(side=tk.LEFT, padx=5)
                
                def create_delete_cmd(index=i):
                    return lambda: delete_reminder(index)
                
                ttk.Button(frame, text="删除", command=create_delete_cmd()).pack(side=tk.RIGHT, padx=5)
        
        def delete_reminder(index):
            if len(self.custom_reminders) > 1:
                del self.custom_reminders[index]
                update_reminder_list()
                self.save_data()
            else:
                messagebox.showwarning("警告", "至少需要保留一个提醒设置！")
        
        # 初始更新列表
        update_reminder_list()
        
        # 居中显示
        reminder_window.update_idletasks()
        width = reminder_window.winfo_width()
        height = reminder_window.winfo_height()
        x = (reminder_window.winfo_screenwidth() // 2) - (width // 2)
        y = (reminder_window.winfo_screenheight() // 2) - (height // 2)
        reminder_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def import_tasks(self):
        """批量导入任务"""
        import_window = tk.Toplevel(self.root)
        import_window.title("批量导入任务")
        import_window.geometry("400x300")
        
        ttk.Label(import_window, text="请按格式输入多个任务（每行一个）:", font=("SimHei", 10)).pack(pady=10)
        ttk.Label(import_window, text="格式: 任务名称 | 日期时间 | 标签", font=("SimHei", 10)).pack(pady=5)
        
        text_area = scrolledtext.ScrolledText(import_window, width=40, height=10, wrap=tk.WORD)
        text_area.pack(pady=10, padx=10)
        
        # 添加示例
        example = "项目报告 | 2024-12-31 23:59 | 工作\n数学作业 | 明天 18:00 | 学习"
        text_area.insert(tk.END, example)
        
        def do_import():
            content = text_area.get(1.0, tk.END).strip()
            if not content:
                messagebox.showwarning("警告", "请输入任务内容！")
                return
            
            lines = content.split("\n")
            imported_count = 0
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                parts = line.split("|")
                if len(parts) < 2:
                    continue
                
                task_name = parts[0].strip()
                date_str = parts[1].strip()
                tag = parts[2].strip() if len(parts) > 2 else "其他"
                
                # 解析日期
                due_date = self.parse_date_time(date_str)
                if not due_date:
                    continue
                
                # 验证标签
                if tag not in self.tags:
                    tag = "其他"
                
                # 添加任务
                new_task = {
                    "name": task_name,
                    "due_date": due_date,
                    "tag": tag,
                    "completed": False,
                    "created_at": datetime.now()
                }
                
                self.tasks.append(new_task)
                imported_count += 1
            
            if imported_count > 0:
                # 确保数据保存到文件
                self.save_data()
                self.update_task_list()
                messagebox.showinfo("成功", f"成功导入 {imported_count} 个任务！")
                import_window.destroy()
            else:
                messagebox.showwarning("警告", "没有成功导入任何任务，请检查格式！")
        
        button_frame = ttk.Frame(import_window)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="导入", command=do_import).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=import_window.destroy).pack(side=tk.LEFT, padx=10)
    
    def check_reminders(self):
        """检查并显示提醒"""
        while not self.stop_event.is_set():
            now = datetime.now()
            
            # 需要提醒的任务
            tasks_to_remind = []
            
            for task in self.tasks:
                if task["completed"]:
                    continue
                
                due_date = task["due_date"]
                if due_date <= now:
                    continue  # 已经过期
                
                # 检查是否需要提醒
            # 优先使用任务专属的提醒设置
            if task.get("reminders"):
                reminders_to_check = task["reminders"]
            else:
                # 如果没有任务专属提醒，则使用全局设置
                reminders_to_check = self.custom_reminders
                
            for value, unit, text in reminders_to_check:
                if unit == "days":
                    reminder_time = due_date - timedelta(days=value)
                elif unit == "hours":
                    reminder_time = due_date - timedelta(hours=value)
                else:  # minutes
                    reminder_time = due_date - timedelta(minutes=value)
                
                # 检查是否在提醒时间范围内（5分钟内）
                time_diff = now - reminder_time
                if 0 <= time_diff.total_seconds() <= 300:  # 5分钟内
                    tasks_to_remind.append((task, text))
            
            # 显示提醒
            for task, reminder_text in tasks_to_remind:
                self.show_reminder(task, reminder_text)
            
            # 每分钟检查一次
            time.sleep(60)
    
    def show_reminder(self, task, reminder_text):
        """显示提醒弹窗"""
        # 在主线程中显示弹窗
        self.root.after(0, lambda: self._show_reminder_dialog(task, reminder_text))
    
    def _show_reminder_dialog(self, task, reminder_text):
        """提醒对话框"""
        reminder_window = tk.Toplevel(self.root)
        reminder_window.title("DDL提醒")
        reminder_window.geometry("300x200")
        reminder_window.attributes('-topmost', True)
        reminder_window.configure(bg="#fff3cd")
        
        # 添加图标（简单的提醒图标）
        icon_frame = ttk.Frame(reminder_window, width=60, height=60)
        icon_frame.pack(pady=10)
        
        canvas = tk.Canvas(icon_frame, width=50, height=50, bg="#fff3cd", highlightthickness=0)
        canvas.pack()
        canvas.create_oval(10, 10, 40, 40, fill="#ffc107", outline="#ffc107")
        canvas.create_text(25, 25, text="!")
        
        # 提醒信息
        message = f"{reminder_text}\n\n任务: {task['name']}\n截止时间: {task['due_date'].strftime('%Y-%m-%d %H:%M')}\n标签: {task['tag']}"
        ttk.Label(reminder_window, text=message, font=("SimHei", 10), wraplength=250).pack(pady=10)
        
        # 按钮
        button_frame = ttk.Frame(reminder_window)
        button_frame.pack(pady=10)
        
        def mark_completed():
            task["completed"] = True
            self.save_data()
            self.update_task_list()
            reminder_window.destroy()
        
        ttk.Button(button_frame, text="标记完成", command=mark_completed).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="知道了", command=reminder_window.destroy).pack(side=tk.LEFT, padx=10)
        
        # 居中显示
        reminder_window.update_idletasks()
        width = reminder_window.winfo_width()
        height = reminder_window.winfo_height()
        x = (reminder_window.winfo_screenwidth() // 2) - (width // 2)
        y = (reminder_window.winfo_screenheight() // 2) - (height // 2)
        reminder_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def on_closing(self):
        """窗口关闭时的处理"""
        self.save_data()
        self.stop_event.set()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = DDLMarker(root)
    root.mainloop()