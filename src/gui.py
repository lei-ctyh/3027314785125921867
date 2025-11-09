"""
GUI模块 - 使用tkinter收集用户输入
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class ConfigGUI:
    """配置输入界面"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("自动化表单填写系统 - 配置")
        self.root.geometry("550x800")
        self.root.resizable(True, True)  # 允许调整大小

        # 存储用户输入的配置
        self.config = None
        self.confirmed = False

        # 设置样式
        self._setup_styles()

        # 创建界面（使用滚动条）
        self._create_widgets_with_scroll()

        # 居中显示窗口
        self._center_window()

    def _setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        style.theme_use('clam')

        # 标题样式
        self.title_font = ('Microsoft YaHei UI', 12, 'bold')
        self.label_font = ('Microsoft YaHei UI', 10)
        self.entry_font = ('Microsoft YaHei UI', 10)

    def _center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _create_widgets_with_scroll(self):
        """创建带滚动条的界面组件"""
        # 创建Canvas和Scrollbar
        canvas = tk.Canvas(self.root, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)

        # 创建一个frame作为滚动区域
        scrollable_frame = ttk.Frame(canvas)

        # 配置滚动
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # 鼠标滚轮绑定
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # 布局Canvas和Scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 在scrollable_frame中创建实际内容
        self._create_content(scrollable_frame)

    def _create_widgets(self):
        """创建界面组件（原有方法，保留兼容性）"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        self._create_content(main_frame)

    def _create_content(self, parent_frame):
        """创建实际的界面内容"""
        # 内边距容器
        main_frame = ttk.Frame(parent_frame, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = tk.Label(
            main_frame,
            text="自动化表单填写系统",
            font=('Microsoft YaHei UI', 16, 'bold'),
            fg='#2c3e50'
        )
        title_label.pack(pady=(0, 20))

        # === 登录信息区域 ===
        login_frame = ttk.LabelFrame(main_frame, text="登录信息", padding="15")
        login_frame.pack(fill=tk.X, pady=(0, 15))

        # 用户名
        tk.Label(login_frame, text="用户名:", font=self.label_font).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.username_var = tk.StringVar(value="330621241001")
        self.username_entry = ttk.Entry(login_frame, textvariable=self.username_var, font=self.entry_font, width=30)
        self.username_entry.grid(row=0, column=1, pady=5, padx=(10, 0))

        # 密码
        tk.Label(login_frame, text="密码:", font=self.label_font).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar(value="Yaojike2022,")
        self.password_entry = ttk.Entry(login_frame, textvariable=self.password_var, font=self.entry_font, width=30, show="*")
        self.password_entry.grid(row=1, column=1, pady=5, padx=(10, 0))

        # === 报表配置区域 ===
        report_frame = ttk.LabelFrame(main_frame, text="报表配置", padding="15")
        report_frame.pack(fill=tk.X, pady=(0, 15))

        # 月份选择
        tk.Label(report_frame, text="报表月份:", font=self.label_font).grid(row=0, column=0, sticky=tk.W, pady=5)
        month_container = ttk.Frame(report_frame)
        month_container.grid(row=0, column=1, pady=5, padx=(10, 0), sticky=tk.W)

        current_month = datetime.now().strftime("%Y-%m")
        self.month_var = tk.StringVar(value=current_month)
        self.month_entry = ttk.Entry(month_container, textvariable=self.month_var, font=self.entry_font, width=15)
        self.month_entry.pack(side=tk.LEFT)

        tk.Label(month_container, text="(格式: YYYY-MM)", font=('Microsoft YaHei UI', 9), fg='gray').pack(side=tk.LEFT, padx=(5, 0))

        # 功能类型
        tk.Label(report_frame, text="功能类型:", font=self.label_font).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.function_type_var = tk.StringVar(value="outpatient")

        function_container = ttk.Frame(report_frame)
        function_container.grid(row=1, column=1, pady=5, padx=(10, 0), sticky=tk.W)

        ttk.Radiobutton(
            function_container,
            text="门诊处方用药录入",
            variable=self.function_type_var,
            value="outpatient",
            command=self._on_function_type_change
        ).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Radiobutton(
            function_container,
            text="急诊处方用药录入",
            variable=self.function_type_var,
            value="emergency",
            command=self._on_function_type_change
        ).pack(side=tk.LEFT)

        # === 数据文件区域 ===
        file_frame = ttk.LabelFrame(main_frame, text="数据文件", padding="15")
        file_frame.pack(fill=tk.X, pady=(0, 15))

        # 输入文件
        tk.Label(file_frame, text="输入文件:", font=self.label_font).grid(row=0, column=0, sticky=tk.W, pady=5)
        file_container = ttk.Frame(file_frame)
        file_container.grid(row=0, column=1, pady=5, padx=(10, 0), sticky=tk.EW)

        self.input_file_var = tk.StringVar(value="data/门诊2509.xlsx")
        self.input_file_entry = ttk.Entry(file_container, textvariable=self.input_file_var, font=self.entry_font, width=35)
        self.input_file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(file_container, text="浏览...", command=self._browse_input_file).pack(side=tk.LEFT, padx=(5, 0))

        file_frame.columnconfigure(1, weight=1)

        # === 浏览器配置区域 ===
        browser_frame = ttk.LabelFrame(main_frame, text="浏览器配置", padding="15")
        browser_frame.pack(fill=tk.X, pady=(0, 15))

        self.headless_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            browser_frame,
            text="无头模式（后台运行，不显示浏览器窗口）",
            variable=self.headless_var
        ).pack(anchor=tk.W)

        # === 手动操作提示区域 ===
        manual_frame = ttk.LabelFrame(main_frame, text="重要提示", padding="15")
        manual_frame.pack(fill=tk.X, pady=(0, 20))

        tips_text = """执行过程中需要两次手动操作：

1. 自动填写前：手动输入报表日期、门诊总量
2. 填写完成后：手动点击"上报"按钮"""

        tips_label = tk.Label(
            manual_frame,
            text=tips_text,
            font=('Microsoft YaHei UI', 9),
            fg='#e74c3c',
            justify=tk.LEFT,
            bg='#fef5e7',
            padx=10,
            pady=10
        )
        tips_label.pack(fill=tk.X)

        # === 按钮区域 ===
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))

        # 使用标准Button而不是ttk.Button，可以自定义颜色
        start_button = tk.Button(
            button_frame,
            text="开始执行",
            command=self._on_confirm,
            width=15,
            height=2,
            font=('Microsoft YaHei UI', 11, 'bold'),
            bg='#27ae60',
            fg='white',
            cursor='hand2',
            relief=tk.RAISED,
            borderwidth=2
        )
        start_button.pack(side=tk.RIGHT, padx=(10, 0))

        cancel_button = tk.Button(
            button_frame,
            text="取消",
            command=self._on_cancel,
            width=15,
            height=2,
            font=('Microsoft YaHei UI', 11),
            bg='#95a5a6',
            fg='white',
            cursor='hand2',
            relief=tk.RAISED,
            borderwidth=2
        )
        cancel_button.pack(side=tk.RIGHT)

    def _on_function_type_change(self):
        """功能类型改变时，自动更新默认文件路径"""
        function_type = self.function_type_var.get()
        if function_type == "outpatient":
            self.input_file_var.set("data/门诊2509.xlsx")
        elif function_type == "emergency":
            self.input_file_var.set("data/emergency_data.xlsx")

    def _browse_input_file(self):
        """浏览选择输入文件"""
        initial_dir = Path(__file__).parent.parent / "data"
        filename = filedialog.askopenfilename(
            title="选择输入数据文件",
            initialdir=str(initial_dir) if initial_dir.exists() else ".",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if filename:
            # 转换为相对路径
            try:
                rel_path = Path(filename).relative_to(Path.cwd())
                self.input_file_var.set(str(rel_path))
            except ValueError:
                # 如果无法转换为相对路径，使用绝对路径
                self.input_file_var.set(filename)

    def _validate_inputs(self) -> bool:
        """验证用户输入"""
        # 验证用户名
        if not self.username_var.get().strip():
            messagebox.showerror("输入错误", "请输入用户名")
            return False

        # 验证密码
        if not self.password_var.get().strip():
            messagebox.showerror("输入错误", "请输入密码")
            return False

        # 验证月份格式
        month = self.month_var.get().strip()
        if not month:
            messagebox.showerror("输入错误", "请输入报表月份")
            return False

        try:
            datetime.strptime(month, "%Y-%m")
        except ValueError:
            messagebox.showerror("输入错误", "月份格式错误，请使用 YYYY-MM 格式（如：2025-09）")
            return False

        # 验证输入文件
        input_file = self.input_file_var.get().strip()
        if not input_file:
            messagebox.showerror("输入错误", "请选择输入文件")
            return False

        if not Path(input_file).exists():
            messagebox.showerror("文件错误", f"输入文件不存在:\n{input_file}")
            return False

        return True

    def _on_confirm(self):
        """确认按钮点击"""
        if not self._validate_inputs():
            return

        # 收集配置
        self.config = {
            "username": self.username_var.get().strip(),
            "password": self.password_var.get().strip(),
            "month": self.month_var.get().strip(),
            "function_type": self.function_type_var.get(),
            "input_file": self.input_file_var.get().strip(),
            "headless": self.headless_var.get()
        }

        self.confirmed = True
        logger.info("用户确认配置")
        self.root.quit()  # 退出主循环
        self.root.destroy()  # 销毁窗口

    def _on_cancel(self):
        """取消按钮点击"""
        if messagebox.askokcancel("确认退出", "确定要退出程序吗？"):
            self.confirmed = False
            logger.info("用户取消操作")
            self.root.quit()
            self.root.destroy()

    def show(self) -> dict:
        """
        显示GUI并等待用户输入

        Returns:
            用户配置字典，如果用户取消则返回None
        """
        # 处理窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # 显示窗口并进入主循环
        self.root.mainloop()

        # 返回配置
        if self.confirmed:
            return self.config
        else:
            return None


class ConfirmationDialog:
    """确认对话框 - 用于执行过程中的提示"""

    def __init__(self, title: str, message: str):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("550x450")
        self.root.resizable(True, True)  # 允许调整大小

        self.confirmed = False

        self._create_widgets(message)
        self._center_window()

    def _center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _create_widgets(self, message: str):
        """创建界面组件"""
        # 创建Canvas和Scrollbar用于滚动
        canvas = tk.Canvas(self.root, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)

        # 创建滚动区域
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # 鼠标滚轮绑定
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # 布局
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 创建内容
        main_frame = ttk.Frame(scrollable_frame, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 图标
        icon_label = tk.Label(
            main_frame,
            text="⚠",
            font=('Arial', 48),
            fg='#f39c12'
        )
        icon_label.pack(pady=(0, 20))

        # 消息
        message_label = tk.Label(
            main_frame,
            text=message,
            font=('Microsoft YaHei UI', 11),
            justify=tk.LEFT,
            wraplength=450
        )
        message_label.pack(pady=(0, 30))

        # 按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))

        confirm_button = tk.Button(
            button_frame,
            text="我已完成，继续",
            command=self._on_confirm,
            width=20,
            height=2,
            font=('Microsoft YaHei UI', 11, 'bold'),
            bg='#3498db',
            fg='white',
            cursor='hand2',
            relief=tk.RAISED,
            borderwidth=2
        )
        confirm_button.pack()

    def _on_confirm(self):
        """确认按钮"""
        self.confirmed = True
        self.root.quit()
        self.root.destroy()

    def show(self) -> bool:
        """显示对话框"""
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)  # 禁用关闭按钮
        self.root.mainloop()
        return self.confirmed


def show_config_gui() -> dict:
    """
    显示配置输入界面

    Returns:
        用户配置字典，如果用户取消则返回None
    """
    gui = ConfigGUI()
    return gui.show()


def show_confirmation_dialog(title: str, message: str) -> bool:
    """
    显示确认对话框

    Args:
        title: 对话框标题
        message: 提示消息

    Returns:
        True如果用户确认
    """
    dialog = ConfirmationDialog(title, message)
    return dialog.show()
