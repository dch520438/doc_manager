"""
文档管家 - UI界面模块
基于 tkinter + ttkbootstrap 实现美观的桌面GUI
兼容 Windows、麒麟(Linux)、macOS
"""

import os
import sys
import platform
import threading
import webbrowser
from datetime import datetime

# tkinter 必须始终导入（ttkbootstrap 依赖它，且代码中直接使用 tk.Frame 等）
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText

try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    from ttkbootstrap.toast import ToastNotification
    from ttkbootstrap.scrolled import ScrolledFrame
    from ttkbootstrap.dialogs import Messagebox, MessageDialog
    HAS_TTB = True
except ImportError:
    HAS_TTB = False

from core import (
    Database, Rule, Document, FileOrganizer, DesktopWatcher,
    ContentExtractor, RuleEngine, GlobalSearcher, SUPPORTED_EXTENSIONS,
    EXTENSION_TYPE_MAP, format_size, get_file_icon, select_folder_dialog,
    open_file, locate_file, get_system_font, check_dependencies, test_extract,
    IS_WINDOWS, IS_LINUX, IS_MACOS
)


# ==================== 颜色主题 ====================
COLORS = {
    'primary': '#4A90D9',
    'success': '#5CB85C',
    'warning': '#F0AD4E',
    'danger': '#D9534F',
    'info': '#5BC0DE',
    'bg': '#F5F6FA',
    'card_bg': '#FFFFFF',
    'text': '#2C3E50',
    'text_secondary': '#7F8C8D',
    'border': '#E1E4E8',
    'btn_topbar': '#E8EEF6',      # 顶部栏按钮背景（替代 rgba）
    'btn_topbar_active': '#D0DAE8',  # 顶部栏按钮激活色
}

if HAS_TTB:
    THEME = "cosmo"

# ==================== 字体 ====================
FONT_FAMILY = get_system_font()


# ==================== 跨平台鼠标滚轮 ====================
def bind_mousewheel(widget, canvas):
    """为 canvas 绑定跨平台鼠标滚轮"""
    def _on_enter(event):
        if IS_WINDOWS or IS_MACOS:
            canvas.bind_all("<MouseWheel>",
                            lambda e: canvas.yview_scroll(-1 * (e.delta // (120 if IS_WINDOWS else 1)), "units"))
        else:
            # Linux: Button-4 = 上滚, Button-5 = 下滚
            canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
            canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

    def _on_leave(event):
        if IS_WINDOWS or IS_MACOS:
            canvas.unbind_all("<MouseWheel>")
        else:
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

    widget.bind("<Enter>", _on_enter)
    widget.bind("<Leave>", _on_leave)


# ==================== 主应用 ====================
class DocManagerApp:
    """文档管家主应用"""

    def __init__(self):
        self.db = Database()
        self.organizer = FileOrganizer(self.db)
        self.watcher = DesktopWatcher(self.db, callback=self._on_auto_organize)
        self.global_searcher = GlobalSearcher()
        self._current_folder = None
        self._search_results = None
        self._search_live_results = []
        self._search_live_lock = threading.Lock()

        self._init_ui()
        self._load_settings()
        self._refresh_all()

    # ---- 初始化 ----
    def _init_ui(self):
        """初始化主窗口"""
        if HAS_TTB:
            self.root = tb.Window(themename=THEME)
            self.root.title("文档管家 - 智能文档分类管理系统")
        else:
            self.root = tk.Tk()
            self.root.title("文档管家 - 智能文档分类管理系统")

        self.root.geometry("1200x750")
        self.root.minsize(1000, 600)

        # Windows DPI 感知
        if IS_WINDOWS:
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass

        # 设置窗口图标
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass

        self._build_styles()
        self._build_layout()

    def _build_styles(self):
        """构建自定义样式"""
        if not HAS_TTB:
            style = ttk.Style()
            style.theme_use('clam')

    def _build_layout(self):
        """构建主布局"""
        self._build_topbar()

        self.main_frame = tk.Frame(self.root, bg=COLORS['bg'])
        self.main_frame.pack(fill=BOTH, expand=True, padx=0, pady=0)

        self._build_left_panel()
        self._build_right_panel()
        self._build_statusbar()

    def _build_topbar(self):
        """构建顶部工具栏"""
        topbar = tk.Frame(self.root, bg=COLORS['primary'], height=56)
        topbar.pack(fill=X)
        topbar.pack_propagate(False)

        # 标题
        title_frame = tk.Frame(topbar, bg=COLORS['primary'])
        title_frame.pack(side=LEFT, padx=15, pady=10)

        tk.Label(title_frame, text="文档管家", font=(FONT_FAMILY, 16, "bold"),
                 bg=COLORS['primary'], fg='white').pack(side=LEFT)

        # 操作按钮区
        btn_frame = tk.Frame(topbar, bg=COLORS['primary'])
        btn_frame.pack(side=RIGHT, padx=15, pady=8)

        # 一键整理按钮
        self.btn_organize = tk.Button(
            btn_frame, text="一键整理桌面", font=(FONT_FAMILY, 11, "bold"),
            bg='#FFFFFF', fg=COLORS['primary'], relief=FLAT, padx=20, pady=5,
            cursor="hand2", command=self._on_organize_click,
            activebackground='#E8E8E8', activeforeground=COLORS['primary']
        )
        self.btn_organize.pack(side=LEFT, padx=5)

        # 实时监控开关
        self.monitor_var = tk.BooleanVar(value=False)
        self.btn_monitor = tk.Button(
            btn_frame, text="开启监控", font=(FONT_FAMILY, 10),
            bg='#2ECC71', fg='white', relief=FLAT, padx=15, pady=5,
            cursor="hand2", command=self._toggle_monitor,
            activebackground='#27AE60'
        )
        self.btn_monitor.pack(side=LEFT, padx=5)

        # 重建索引按钮
        self.btn_reindex = tk.Button(
            btn_frame, text="重建索引", font=(FONT_FAMILY, 10),
            bg='#F39C12', fg='white', relief=FLAT, padx=12, pady=5,
            cursor="hand2", command=self._on_reindex_click,
            activebackground='#E67E22'
        )
        self.btn_reindex.pack(side=LEFT, padx=5)

        # 设置按钮
        tk.Button(
            btn_frame, text="设置", font=(FONT_FAMILY, 10),
            bg=COLORS['btn_topbar'], fg='white', relief=FLAT, padx=12, pady=5,
            cursor="hand2", command=self._show_settings,
            activebackground=COLORS['btn_topbar_active']
        ).pack(side=LEFT, padx=5)

    def _build_left_panel(self):
        """构建左侧面板"""
        self.left_panel = tk.Frame(self.main_frame, bg=COLORS['card_bg'], width=280)
        self.left_panel.pack(side=LEFT, fill=Y, padx='0 0')
        self.left_panel.pack_propagate(False)

        # 搜索框
        search_frame = tk.Frame(self.left_panel, bg=COLORS['card_bg'], padx=10, pady=10)
        search_frame.pack(fill=X)

        tk.Label(search_frame, text="快速搜索", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS['card_bg'], fg=COLORS['text']).pack(anchor=W)

        search_input_frame = tk.Frame(search_frame, bg=COLORS['card_bg'])
        search_input_frame.pack(fill=X, pady='5 0')

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            search_input_frame, textvariable=self.search_var,
            font=(FONT_FAMILY, 10), relief=FLAT,
            bg='#F0F2F5', fg=COLORS['text'], insertbackground=COLORS['text']
        )
        self.search_entry.pack(side=LEFT, fill=X, expand=True, ipady=6, padx='8 0')
        # 回车搜索
        self.search_entry.bind('<Return>', lambda e: self._do_search())

        # 搜索选项
        self.search_content_var = tk.BooleanVar(value=True)
        search_opt_frame = tk.Frame(search_frame, bg=COLORS['card_bg'])
        search_opt_frame.pack(fill=X, pady='3 0')
        tk.Checkbutton(
            search_opt_frame, text="搜索文档内容", variable=self.search_content_var,
            font=(FONT_FAMILY, 9), bg=COLORS['card_bg'],
            fg=COLORS['text_secondary'], selectcolor=COLORS['card_bg'],
            activebackground=COLORS['card_bg']
        ).pack(side=LEFT)

        # 搜索范围
        self.search_scope_var = tk.StringVar(value="managed")
        scope_frame = tk.Frame(search_frame, bg=COLORS['card_bg'])
        scope_frame.pack(fill=X, pady='3 0')
        tk.Label(scope_frame, text="范围:", font=(FONT_FAMILY, 9),
                 bg=COLORS['card_bg'], fg=COLORS['text_secondary']).pack(side=LEFT)
        for text, val in [("已管理", "managed"), ("全盘", "global"), ("指定路径", "custom")]:
            tk.Radiobutton(
                scope_frame, text=text, variable=self.search_scope_var, value=val,
                font=(FONT_FAMILY, 9), bg=COLORS['card_bg'],
                fg=COLORS['text_secondary'], selectcolor=COLORS['card_bg'],
                activebackground=COLORS['card_bg']
            ).pack(side=LEFT, padx='2 0')

        # 自定义路径输入（默认隐藏）
        self.custom_path_frame = tk.Frame(search_frame, bg=COLORS['card_bg'])
        self.custom_path_var = tk.StringVar()
        self.custom_path_entry = tk.Entry(
            self.custom_path_frame, textvariable=self.custom_path_var,
            font=(FONT_FAMILY, 9), relief=FLAT,
            bg='#F0F2F5', fg=COLORS['text'], insertbackground=COLORS['text']
        )
        self.custom_path_entry.pack(side=LEFT, fill=X, expand=True, ipady=3, padx='0 3')

        def browse_custom_path():
            path = select_folder_dialog("选择搜索路径", parent=self.root)
            if path:
                self.custom_path_var.set(path)

        tk.Button(self.custom_path_frame, text="浏览", font=(FONT_FAMILY, 8),
                  bg=COLORS['primary'], fg='white', relief=FLAT, padx=6, pady=1,
                  cursor="hand2", command=browse_custom_path).pack(side=LEFT)

        def on_scope_change(*args):
            if self.search_scope_var.get() == "custom":
                self.custom_path_frame.pack(fill=X, pady='3 0')
            else:
                self.custom_path_frame.pack_forget()

        self.search_scope_var.trace_add('write', on_scope_change)

        # 搜索按钮 + 停止按钮
        search_btn_frame = tk.Frame(search_frame, bg=COLORS['card_bg'])
        search_btn_frame.pack(fill=X, pady='5 0')
        tk.Button(
            search_btn_frame, text="搜索", font=(FONT_FAMILY, 9, "bold"),
            bg=COLORS['primary'], fg='white', relief=FLAT, padx=15, pady=3,
            cursor="hand2", command=self._do_search
        ).pack(side=LEFT, padx='0 5')
        self.btn_stop_search = tk.Button(
            search_btn_frame, text="停止", font=(FONT_FAMILY, 9),
            bg=COLORS['danger'], fg='white', relief=FLAT, padx=15, pady=3,
            cursor="hand2", command=self._stop_search, state=DISABLED
        )
        self.btn_stop_search.pack(side=LEFT)

        # 搜索进度条（默认隐藏）
        self.search_progress_var = tk.DoubleVar(value=0)
        self.search_progress = tk.Canvas(search_frame, height=4, bg='#E1E4E8',
                                         highlightthickness=0)
        self.search_progress_label = tk.Label(search_frame, text="", font=(FONT_FAMILY, 8),
                                              bg=COLORS['card_bg'], fg=COLORS['text_secondary'])

        # 分隔线
        tk.Frame(self.left_panel, bg=COLORS['border'], height=1).pack(fill=X, padx=10)

        # 文件夹列表
        folder_label_frame = tk.Frame(self.left_panel, bg=COLORS['card_bg'], padx=10, pady=10)
        folder_label_frame.pack(fill=X, pady='0 5')

        tk.Label(folder_label_frame, text="文档分类", font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS['card_bg'], fg=COLORS['text']).pack(side=LEFT)

        tk.Button(folder_label_frame, text="管理规则 >", font=(FONT_FAMILY, 9),
                   bg=COLORS['card_bg'], fg=COLORS['primary'], relief=FLAT,
                   cursor="hand2", command=self._show_rules_dialog,
                   activebackground=COLORS['card_bg'], activeforeground=COLORS['primary'],
                   bd=0, highlightthickness=0).pack(side=RIGHT)

        # 文件夹树
        tree_frame = tk.Frame(self.left_panel, bg=COLORS['card_bg'])
        tree_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)

        self.folder_canvas = tk.Canvas(tree_frame, bg=COLORS['card_bg'], highlightthickness=0)
        folder_scrollbar = tk.Scrollbar(tree_frame, orient=VERTICAL, command=self.folder_canvas.yview)
        self.folder_inner = tk.Frame(self.folder_canvas, bg=COLORS['card_bg'])

        self.folder_inner.bind("<Configure>",
                               lambda e: self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all")))
        self.folder_canvas.create_window((0, 0), window=self.folder_inner, anchor=NW)
        self.folder_canvas.configure(yscrollcommand=folder_scrollbar.set)

        self.folder_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        folder_scrollbar.pack(side=RIGHT, fill=Y)

        # 跨平台鼠标滚轮
        bind_mousewheel(tree_frame, self.folder_canvas)

    def _build_right_panel(self):
        """构建右侧内容区"""
        self.right_panel = tk.Frame(self.main_frame, bg=COLORS['bg'])
        self.right_panel.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=10)

        self.content_title = tk.Label(
            self.right_panel, text="所有文档", font=(FONT_FAMILY, 14, "bold"),
            bg=COLORS['bg'], fg=COLORS['text']
        )
        self.content_title.pack(anchor=W, pady='0 10')

        self._build_doc_list()

    def _build_doc_list(self):
        """构建文档列表"""
        header_frame = tk.Frame(self.right_panel, bg=COLORS['card_bg'])
        header_frame.pack(fill=X)

        headers = [
            ("", 40), ("文件名", 350), ("类型", 70), ("大小", 80),
            ("分类规则", 120), ("整理时间", 150), ("操作", 120)
        ]
        for text, width in headers:
            tk.Label(header_frame, text=text, font=(FONT_FAMILY, 9, "bold"),
                     bg=COLORS['card_bg'], fg=COLORS['text_secondary'], width=width // 8,
                     anchor=W).pack(side=LEFT, padx=8, pady=8)

        list_container = tk.Frame(self.right_panel, bg=COLORS['card_bg'])
        list_container.pack(fill=BOTH, expand=True)

        self.doc_canvas = tk.Canvas(list_container, bg=COLORS['card_bg'], highlightthickness=0)
        doc_scrollbar = tk.Scrollbar(list_container, orient=VERTICAL, command=self.doc_canvas.yview)
        self.doc_inner = tk.Frame(self.doc_canvas, bg=COLORS['card_bg'])

        self.doc_inner.bind("<Configure>",
                            lambda e: self.doc_canvas.configure(scrollregion=self.doc_canvas.bbox("all")))
        self.doc_canvas.create_window((0, 0), window=self.doc_inner, anchor=NW)
        self.doc_canvas.configure(yscrollcommand=doc_scrollbar.set)

        self.doc_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        doc_scrollbar.pack(side=RIGHT, fill=Y)

        # 跨平台鼠标滚轮
        bind_mousewheel(list_container, self.doc_canvas)

    def _build_statusbar(self):
        """构建底部状态栏"""
        self.statusbar = tk.Frame(self.root, bg='#2C3E50', height=30)
        self.statusbar.pack(fill=X, side=BOTTOM)
        self.statusbar.pack_propagate(False)

        self.status_label = tk.Label(
            self.statusbar, text="就绪", font=(FONT_FAMILY, 9),
            bg='#2C3E50', fg='#BDC3C7', padx=10
        )
        self.status_label.pack(side=LEFT)

        self.doc_count_label = tk.Label(
            self.statusbar, text="已管理文档: 0", font=(FONT_FAMILY, 9),
            bg='#2C3E50', fg='#BDC3C7', padx=10
        )
        self.doc_count_label.pack(side=RIGHT)

        self.monitor_label = tk.Label(
            self.statusbar, text="", font=(FONT_FAMILY, 9),
            bg='#2C3E50', fg='#2ECC71', padx=10
        )
        self.monitor_label.pack(side=RIGHT)

    # ---- 数据加载 ----
    def _load_settings(self):
        """加载设置"""
        pass

    def _refresh_all(self):
        """刷新所有数据"""
        self._refresh_folders()
        self._refresh_documents()
        self._update_status()

    def _refresh_folders(self):
        """刷新文件夹列表"""
        for widget in self.folder_inner.winfo_children():
            widget.destroy()

        folder_counts = self.db.get_folder_counts()

        total = sum(folder_counts.values())
        all_btn = self._create_folder_item("全部文档", total, is_all=True)
        all_btn.pack(fill=X, pady=1)

        for folder, count in folder_counts.items():
            folder_name = os.path.basename(folder) or folder
            btn = self._create_folder_item(folder_name, count, folder=folder)
            btn.pack(fill=X, pady=1)

        if not folder_counts:
            tk.Label(
                self.folder_inner, text="暂无文档\n请先添加规则并整理桌面",
                font=(FONT_FAMILY, 10), bg=COLORS['card_bg'],
                fg=COLORS['text_secondary'], pady=30
            ).pack()

    def _create_folder_item(self, text: str, count: int, is_all=False, folder=None):
        """创建文件夹列表项"""
        frame = tk.Frame(self.folder_inner, bg=COLORS['card_bg'], cursor="hand2", padx=10, pady=6)

        def on_enter(e):
            frame.configure(bg='#E8F0FE')
            label.configure(bg='#E8F0FE')
            count_lbl.configure(bg='#E8F0FE')

        def on_leave(e):
            frame.configure(bg=COLORS['card_bg'])
            label.configure(bg=COLORS['card_bg'])
            count_lbl.configure(bg=COLORS['card_bg'])

        def on_click(e):
            if is_all:
                self._current_folder = None
                self.content_title.config(text="所有文档")
            else:
                self._current_folder = folder
                folder_name = os.path.basename(folder) or folder
                self.content_title.config(text=folder_name)
            self._refresh_documents()

        frame.bind("<Enter>", on_enter)
        frame.bind("<Leave>", on_leave)
        frame.bind("<Button-1>", on_click)

        label = tk.Label(frame, text=text, font=(FONT_FAMILY, 10),
                         bg=COLORS['card_bg'], fg=COLORS['text'], anchor=W)
        label.pack(side=LEFT, fill=X, expand=True)
        label.bind("<Enter>", on_enter)
        label.bind("<Leave>", on_leave)
        label.bind("<Button-1>", on_click)

        count_lbl = tk.Label(frame, text=str(count), font=(FONT_FAMILY, 9),
                             bg=COLORS['card_bg'], fg=COLORS['text_secondary'])
        count_lbl.pack(side=RIGHT)
        count_lbl.bind("<Enter>", on_enter)
        count_lbl.bind("<Leave>", on_leave)
        count_lbl.bind("<Button-1>", on_click)

        return frame

    def _refresh_documents(self, documents=None):
        """刷新文档列表"""
        for widget in self.doc_inner.winfo_children():
            widget.destroy()

        if documents is None:
            if self._current_folder:
                documents = self.db.get_documents_by_folder(self._current_folder)
            else:
                documents = self.db.get_all_documents()

        if not documents:
            tk.Label(
                self.doc_inner, text="暂无文档记录\n\n提示：\n1. 在左侧「管理规则」添加分类规则\n2. 点击「一键整理桌面」开始整理\n3. 或开启「实时监控」自动整理",
                font=(FONT_FAMILY, 10), bg=COLORS['card_bg'],
                fg=COLORS['text_secondary'], pady=40, justify=LEFT
            ).pack(anchor=W, padx=20)
            return

        for i, doc in enumerate(documents):
            self._create_doc_row(doc, i % 2 == 0)

    def _create_doc_row(self, doc: Document, is_even: bool):
        """创建文档行"""
        bg = '#F8F9FA' if is_even else COLORS['card_bg']
        row = tk.Frame(self.doc_inner, bg=bg, padx=5, pady=4)

        def on_enter(e):
            row.configure(bg='#E3F2FD')

        def on_leave(e):
            row.configure(bg=bg)

        row.bind("<Enter>", on_enter)
        row.bind("<Leave>", on_leave)

        icon = get_file_icon(doc.file_type)
        items = [
            (icon, 40),
            (doc.filename, 350),
            (doc.file_type, 70),
            (format_size(doc.file_size), 80),
            (doc.rule_name or "未分类", 120),
            (doc.organized_at, 150),
        ]

        for text, width in items:
            lbl = tk.Label(row, text=text, font=(FONT_FAMILY, 9),
                           bg=bg, fg=COLORS['text'], anchor=W, width=width // 8)
            lbl.pack(side=LEFT, padx=5)
            lbl.bind("<Enter>", on_enter)
            lbl.bind("<Leave>", on_leave)

        btn_frame = tk.Frame(row, bg=bg)
        btn_frame.pack(side=LEFT, padx=5)

        # 打开按钮
        open_btn = tk.Label(btn_frame, text="打开", font=(FONT_FAMILY, 8),
                            bg=COLORS['primary'], fg='white', padx=6, pady=1, cursor="hand2")
        open_btn.pack(side=LEFT, padx=2)
        open_btn.bind("<Button-1>", lambda e, d=doc: self._open_document(d))
        open_btn.bind("<Enter>", lambda e, b=open_btn: b.configure(bg='#3A7BC8'))
        open_btn.bind("<Leave>", lambda e, b=open_btn: b.configure(bg=COLORS['primary']))

        # 定位按钮
        loc_btn = tk.Label(btn_frame, text="定位", font=(FONT_FAMILY, 8),
                           bg='#95A5A6', fg='white', padx=6, pady=1, cursor="hand2")
        loc_btn.pack(side=LEFT, padx=2)
        loc_btn.bind("<Button-1>", lambda e, d=doc: self._locate_document(d))
        loc_btn.bind("<Enter>", lambda e, b=loc_btn: b.configure(bg='#7F8C8D'))
        loc_btn.bind("<Leave>", lambda e, b=loc_btn: b.configure(bg='#95A5A6'))

        row.pack(fill=X)

    def _update_status(self):
        """更新状态栏"""
        count = self.db.get_document_count()
        self.doc_count_label.config(text=f"已管理文档: {count}")

        if self.watcher.is_watching():
            self.monitor_label.config(text="● 监控中", fg='#2ECC71')
        else:
            self.monitor_label.config(text="○ 监控关闭", fg='#95A5A6')

    # ---- 事件处理 ----
    def _on_organize_click(self):
        """一键整理"""
        rules = self.db.get_enabled_rules()
        if not rules:
            self._show_message("提示", "没有可用的分类规则！\n请先在左侧点击「管理规则」添加规则。", "warning")
            return

        if not self._confirm("确认整理", "将扫描桌面并根据规则自动整理文档，是否继续？"):
            return

        self.btn_organize.config(state=DISABLED, text="整理中...")
        self.status_label.config(text="正在整理桌面文档...")

        def do_organize():
            def progress(current, total, msg):
                self.root.after(0, lambda: self.status_label.config(text=f"{msg} ({current}/{total})"))

            success, skip, errors = self.organizer.organize_files(progress_callback=progress)
            self.root.after(0, lambda: self._on_organize_done(success, skip, errors))

        threading.Thread(target=do_organize, daemon=True).start()

    def _on_organize_done(self, success: int, skip: int, errors: list):
        """整理完成回调"""
        self.btn_organize.config(state=NORMAL, text="一键整理桌面")
        self._refresh_all()

        msg = f"整理完成！\n\n成功整理: {success} 个文件\n跳过（无匹配规则）: {skip} 个文件"
        if errors:
            msg += f"\n\n错误:\n" + "\n".join(f"· {e}" for e in errors[:5])
        self._show_message("整理结果", msg, "success" if not errors else "warning")
        self.status_label.config(text="整理完成")

    def _toggle_monitor(self):
        """切换实时监控"""
        if self.watcher.is_watching():
            self.watcher.stop()
            self.btn_monitor.config(text="开启监控", bg='#2ECC71', activebackground='#27AE60')
            self._update_status()
            self.status_label.config(text="实时监控已关闭")
        else:
            rules = self.db.get_enabled_rules()
            if not rules:
                self._show_message("提示", "没有可用的分类规则！\n请先添加规则再开启监控。", "warning")
                return
            self.watcher.start()
            self.btn_monitor.config(text="关闭监控", bg='#E74C3C', activebackground='#C0392B')
            self._update_status()
            self.status_label.config(text="实时监控已开启 - 自动整理桌面新文件")

    def _on_reindex_click(self):
        """重建文档内容索引"""
        docs = self.db.get_all_documents()
        if not docs:
            self.status_label.config(text="没有需要索引的文档")
            return

        self.btn_reindex.config(state=DISABLED, text="索引中...")
        self.status_label.config(text=f"正在重建索引（共{len(docs)}个文档）...")

        def do_reindex():
            success = 0
            fail = 0
            no_content = 0
            for i, doc in enumerate(docs):
                if not os.path.exists(doc.filepath):
                    fail += 1
                    continue
                try:
                    content = ContentExtractor.extract_text(doc.filepath)
                    if content and content.strip():
                        self.db.update_document_content(doc.id, content)
                        success += 1
                    else:
                        no_content += 1
                except Exception:
                    fail += 1
                if (i + 1) % 5 == 0:
                    self.root.after(0, lambda c=i+1, s=success: self.status_label.config(
                        text=f"索引中: {c}/{len(docs)} (成功{s}个)"))

            self.root.after(0, lambda: self._on_reindex_done(success, fail, no_content))

        threading.Thread(target=do_reindex, daemon=True).start()

    def _on_reindex_done(self, success: int, fail: int, no_content: int):
        """索引重建完成"""
        self.btn_reindex.config(state=NORMAL, text="重建索引")
        self._refresh_all()
        msg = f"索引重建完成: 成功{success}个"
        if no_content > 0:
            msg += f", 无内容{no_content}个"
        if fail > 0:
            msg += f", 失败{fail}个"
        self.status_label.config(text=msg)

    def _on_auto_organize(self, filename: str, target_folder: str, success: bool, error: str):
        """自动整理回调"""
        def update():
            self._refresh_all()
            if success:
                self.status_label.config(text=f"自动整理: {filename} -> {os.path.basename(target_folder)}")
            else:
                self.status_label.config(text=f"自动整理失败: {filename}")
        self.root.after(0, update)

    def _on_search_change(self, *args):
        """搜索框内容变化"""
        pass

    def _do_search(self):
        """执行搜索"""
        keyword = self.search_var.get().strip()
        if not keyword:
            self._search_results = None
            self._current_folder = None
            self._search_live_results = None
            self.content_title.config(text="所有文档")
            self._refresh_documents()
            return

        scope = self.search_scope_var.get()
        search_content = self.search_content_var.get()

        if scope in ("global", "custom"):
            # 获取搜索路径
            if scope == "custom":
                custom = self.custom_path_var.get().strip()
                if not custom:
                    self.status_label.config(text="请输入或选择搜索路径")
                    return
                paths = [p.strip() for p in custom.split(';') if p.strip() and os.path.isdir(p.strip())]
                if not paths:
                    self.status_label.config(text="路径无效，请检查")
                    return
            else:
                paths = None  # 使用默认全部路径

            self._do_global_search(keyword, search_content, paths)
        else:
            # 已管理文档搜索
            results = self.db.search_documents(keyword, search_content)
            self._search_results = results
            self._search_live_results = None
            self.content_title.config(text=f"搜索: \"{keyword}\" ({len(results)}个)")
            self._refresh_documents(results)
            self.status_label.config(text=f"搜索完成: 找到 {len(results)} 个文件")

    def _do_global_search(self, keyword: str, search_content: bool, paths=None):
        """执行全盘搜索（批量收集 + 定时刷新，确保结果稳定显示）"""
        self.btn_stop_search.config(state=NORMAL)
        self._search_live_results = []
        self._search_live_lock = threading.Lock()
        self._search_running = True
        self._refresh_documents([])  # 先清空列表
        self.content_title.config(text=f"搜索中: \"{keyword}\"...")

        # 显示进度条
        self.search_progress.pack(fill=X, pady='3 0')
        self.search_progress_label.pack(anchor=W)
        self.search_progress.delete("all")
        self.search_progress.create_rectangle(0, 0, 0, 4, fill=COLORS['primary'], outline='')

        def on_found(doc):
            """收集结果（不直接操作UI，避免线程安全问题）"""
            with self._search_live_lock:
                self._search_live_results.append(doc)

        def on_progress(searched, found, current_path):
            """更新进度"""
            short_path = current_path[-40:] if len(current_path) > 40 else current_path
            self.root.after_idle(lambda s=searched, f=found, p=short_path: self._update_search_progress(s, f, p))

        def on_done(total):
            """搜索完成"""
            self.root.after(0, lambda: self._on_global_search_done(keyword))

        def do_search():
            try:
                self.global_searcher.search(
                    keyword,
                    search_paths=paths,
                    search_content=search_content,
                    on_found=on_found,
                    on_progress=on_progress,
                    on_done=on_done
                )
            except Exception as e:
                self.root.after(0, lambda: self.status_label.config(text=f"搜索出错: {e}"))
                self._search_running = False

        threading.Thread(target=do_search, daemon=True).start()

        # 启动定时刷新（每500ms从结果队列中取新结果显示）
        self._start_search_refresh()

    def _start_search_refresh(self):
        """定时刷新搜索结果到UI（每500ms）"""
        if not self._search_running:
            return

        # 取出当前所有结果
        with self._search_live_lock:
            current_results = list(self._search_live_results)
            current_count = len(current_results)

        # 检查是否需要刷新（结果数量变化了）
        prev_count = getattr(self, '_last_refresh_count', 0)
        if current_count != prev_count:
            self._last_refresh_count = current_count
            self._search_results = current_results
            self._refresh_documents(current_results)
            self.content_title.config(text=f"搜索中... 已找到 {current_count} 个")

        # 500ms 后再次刷新
        self._search_refresh_timer = self.root.after(500, self._start_search_refresh)

    def _stop_search_refresh(self):
        """停止定时刷新"""
        self._search_running = False
        timer = getattr(self, '_search_refresh_timer', None)
        if timer:
            try:
                self.root.after_cancel(timer)
            except Exception:
                pass

    def _update_search_progress(self, searched, found, path):
        """更新搜索进度"""
        self.search_progress.delete("all")
        w = self.search_progress.winfo_width()
        if w > 1:
            bar_w = (searched * 3) % w
            self.search_progress.create_rectangle(0, 0, bar_w, 4,
                                                  fill=COLORS['primary'], outline='')
        self.search_progress_label.config(text=f"已扫描 {searched} 个文件 | 找到 {found} 个 | {path}")
        self.status_label.config(text=f"搜索中... 已扫描{searched}个文件, 找到{found}个")

    def _on_global_search_done(self, keyword: str):
        """全盘搜索完成"""
        self._stop_search_refresh()
        self.btn_stop_search.config(state=DISABLED)
        # 隐藏进度条
        self.search_progress.pack_forget()
        self.search_progress_label.pack_forget()

        # 最终刷新：用完整结果列表
        with self._search_live_lock:
            final_results = list(self._search_live_results)
        self._search_results = final_results
        self._refresh_documents(final_results)
        total = len(final_results)
        self.content_title.config(text=f"搜索完成: \"{keyword}\" ({total}个)")
        self.status_label.config(text=f"搜索完成: 共找到 {total} 个文件")

    def _stop_search(self):
        """停止搜索"""
        self.global_searcher.stop()
        self._stop_search_refresh()
        self.btn_stop_search.config(state=DISABLED)
        self.search_progress.pack_forget()
        self.search_progress_label.pack_forget()
        found = len(self._search_live_results)
        self._search_results = list(self._search_live_results)
        self._refresh_documents(self._search_results)
        self.content_title.config(text=f"搜索已停止 ({found}个)")
        self.status_label.config(text="搜索已停止")

    def _open_document(self, doc: Document):
        """打开文档（跨平台）"""
        filepath = doc.filepath
        if not open_file(filepath):
            self._show_message("错误", f"文件不存在或无法打开:\n{filepath}", "error")

    def _locate_document(self, doc: Document):
        """在文件管理器中定位文档（跨平台）"""
        filepath = doc.filepath
        if not locate_file(filepath):
            self._show_message("错误", f"文件不存在:\n{filepath}", "error")

    # ---- 规则管理对话框 ----
    def _show_rules_dialog(self):
        """显示规则管理对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("分类规则管理")
        dialog.geometry("750x520")
        dialog.transient(self.root)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 750) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 520) // 2
        dialog.geometry(f"+{x}+{y}")

        title_bar = tk.Frame(dialog, bg=COLORS['primary'], height=45)
        title_bar.pack(fill=X)
        title_bar.pack_propagate(False)

        tk.Label(title_bar, text="分类规则管理", font=(FONT_FAMILY, 13, "bold"),
                 bg=COLORS['primary'], fg='white').pack(side=LEFT, padx=15, pady=8)

        tk.Button(title_bar, text="+ 添加规则", font=(FONT_FAMILY, 10),
                  bg='white', fg=COLORS['primary'], relief=FLAT, padx=12, cursor="hand2",
                  command=lambda: self._show_rule_editor(dialog)).pack(side=RIGHT, padx=15, pady=8)

        rules_container = tk.Frame(dialog, bg=COLORS['bg'])
        rules_container.pack(fill=BOTH, expand=True, padx=10, pady=10)

        canvas = tk.Canvas(rules_container, bg=COLORS['bg'], highlightthickness=0)
        scrollbar = tk.Scrollbar(rules_container, orient=VERTICAL, command=canvas.yview)
        inner = tk.Frame(canvas, bg=COLORS['bg'])

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor=NW)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        bind_mousewheel(rules_container, canvas)

        def refresh_rules():
            for w in inner.winfo_children():
                w.destroy()
            rules = self.db.get_all_rules()
            if not rules:
                tk.Label(inner, text="暂无规则，点击右上角「添加规则」开始设置",
                         font=(FONT_FAMILY, 10), bg=COLORS['bg'],
                         fg=COLORS['text_secondary'], pady=30).pack()
                return
            for i, rule in enumerate(rules):
                self._create_rule_card(inner, rule, dialog, refresh_rules)

        refresh_rules()

    def _create_rule_card(self, parent, rule: Rule, parent_dialog, refresh_fn):
        """创建规则卡片"""
        card = tk.Frame(parent, bg=COLORS['card_bg'], padx=12, pady=10,
                        highlightbackground=COLORS['border'], highlightthickness=1)

        status_color = COLORS['success'] if rule.enabled else COLORS['text_secondary']
        status_text = "● 启用" if rule.enabled else "○ 禁用"

        top_row = tk.Frame(card, bg=COLORS['card_bg'])
        top_row.pack(fill=X)

        tk.Label(top_row, text=rule.name, font=(FONT_FAMILY, 11, "bold"),
                 bg=COLORS['card_bg'], fg=COLORS['text']).pack(side=LEFT)

        tk.Label(top_row, text=status_text, font=(FONT_FAMILY, 9),
                 bg=COLORS['card_bg'], fg=status_color).pack(side=LEFT, padx=10)

        tk.Label(top_row, text=f"优先级: {rule.priority}", font=(FONT_FAMILY, 9),
                 bg=COLORS['card_bg'], fg=COLORS['text_secondary']).pack(side=LEFT, padx=5)

        btn_frame = tk.Frame(top_row, bg=COLORS['card_bg'])
        btn_frame.pack(side=RIGHT)

        def toggle():
            rule.enabled = not rule.enabled
            self.db.update_rule(rule)
            refresh_fn()

        tk.Label(btn_frame, text="切换", font=(FONT_FAMILY, 8),
                 bg='#3498DB', fg='white', padx=8, pady=2, cursor="hand2").pack(side=LEFT, padx=2)
        btn_frame.children[list(btn_frame.children.keys())[-1]].bind("<Button-1>", lambda e: toggle())

        tk.Label(btn_frame, text="编辑", font=(FONT_FAMILY, 8),
                 bg=COLORS['warning'], fg='white', padx=8, pady=2, cursor="hand2").pack(side=LEFT, padx=2)
        btn_frame.children[list(btn_frame.children.keys())[-1]].bind(
            "<Button-1>", lambda e: self._show_rule_editor(parent_dialog, rule, refresh_fn))

        tk.Label(btn_frame, text="删除", font=(FONT_FAMILY, 8),
                 bg=COLORS['danger'], fg='white', padx=8, pady=2, cursor="hand2").pack(side=LEFT, padx=2)
        btn_frame.children[list(btn_frame.children.keys())[-1]].bind(
            "<Button-1>", lambda e: self._delete_rule(rule, refresh_fn))

        detail_frame = tk.Frame(card, bg=COLORS['card_bg'])
        detail_frame.pack(fill=X, pady='5 0')

        details = []
        if rule.contain_keywords:
            details.append(f"包含关键词: {rule.contain_keywords}")
        if rule.exclude_keywords:
            details.append(f"排除关键词: {rule.exclude_keywords}")
        details.append(f"目标位置: {rule.target_folder}")

        detail_text = "  |  ".join(details)
        tk.Label(detail_frame, text=detail_text, font=(FONT_FAMILY, 9),
                 bg=COLORS['card_bg'], fg=COLORS['text_secondary'], wraplength=600,
                 justify=LEFT).pack(anchor=W)

        card.pack(fill=X, pady=3)

    def _show_rule_editor(self, parent_dialog, rule=None, refresh_fn=None):
        """显示规则编辑对话框（内联式，避免麒麟系统嵌套窗口问题）"""
        is_edit = rule is not None
        editor = tk.Toplevel(parent_dialog if parent_dialog else self.root)
        editor.title("编辑规则" if is_edit else "添加规则")
        editor.geometry("560x520")
        editor.transient(parent_dialog if parent_dialog else self.root)
        editor.resizable(False, False)

        editor.update_idletasks()
        px = (editor.winfo_screenwidth() - 560) // 2
        py = (editor.winfo_screenheight() - 520) // 2
        editor.geometry(f"+{px}+{py}")

        # 注意：不使用 grab_set()，避免麒麟系统死锁

        title_bar = tk.Frame(editor, bg=COLORS['primary'], height=45)
        title_bar.pack(fill=X)
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text=("编辑分类规则" if is_edit else "添加分类规则"),
                 font=(FONT_FAMILY, 12, "bold"),
                 bg=COLORS['primary'], fg='white').pack(side=LEFT, padx=15, pady=8)

        # 按钮区 — 必须先 pack，固定在底部，避免被 form 挤出可视区域
        btn_frame = tk.Frame(editor, bg=COLORS['bg'])
        btn_frame.pack(side=BOTTOM, fill=X, padx=25, pady=15)

        def save():
            try:
                error_var.set("")  # 清除错误

                name = fields['name'].get().strip()
                if not name:
                    error_var.set("请输入规则名称")
                    return
                target = folder_var.get().strip()
                if not target:
                    error_var.set("请输入目标文件夹路径（可直接输入，如 /data/报告）")
                    return

                # 自动创建目标文件夹
                if auto_create_var.get() and not os.path.exists(target):
                    try:
                        os.makedirs(target, exist_ok=True)
                    except Exception as e:
                        error_var.set(f"创建文件夹失败: {e}")
                        return

                new_rule = Rule(
                    id=rule.id if rule else 0,
                    name=name,
                    contain_keywords=fields['contain'].get().strip(),
                    exclude_keywords=fields['exclude'].get().strip(),
                    target_folder=target,
                    enabled=enabled_var.get(),
                    priority=priority_var.get()
                )

                if is_edit:
                    self.db.update_rule(new_rule)
                else:
                    self.db.add_rule(new_rule)

                editor.destroy()
                if refresh_fn:
                    refresh_fn()
                self._refresh_folders()
            except Exception as e:
                error_var.set(f"保存失败: {e}")

        tk.Button(btn_frame, text="保存", font=(FONT_FAMILY, 10, "bold"),
                  bg=COLORS['success'], fg='white', relief=FLAT, padx=25, pady=5,
                  cursor="hand2", command=save).pack(side=RIGHT, padx=5)

        tk.Button(btn_frame, text="取消", font=(FONT_FAMILY, 10),
                  bg='#95A5A6', fg='white', relief=FLAT, padx=25, pady=5,
                  cursor="hand2", command=editor.destroy).pack(side=RIGHT, padx=5)

        # 表单区域 — 放在按钮之后 pack，填充剩余空间
        form_container = tk.Frame(editor, bg=COLORS['bg'])
        form_container.pack(side=TOP, fill=BOTH, expand=True)

        # 表单内用 Canvas + Scrollbar 实现滚动
        form_canvas = tk.Canvas(form_container, bg=COLORS['bg'], highlightthickness=0)
        form_scrollbar = tk.Scrollbar(form_container, orient=VERTICAL, command=form_canvas.yview)
        form = tk.Frame(form_canvas, bg=COLORS['bg'], padx=25, pady=15)

        form.bind("<Configure>",
                  lambda e: form_canvas.configure(scrollregion=form_canvas.bbox("all")))
        form_canvas.create_window((0, 0), window=form, anchor=NW)
        form_canvas.configure(yscrollcommand=form_scrollbar.set)

        form_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        form_scrollbar.pack(side=RIGHT, fill=Y)

        # 鼠标滚轮
        bind_mousewheel(form_container, form_canvas)

        # 内联错误提示标签
        error_var = tk.StringVar(value="")
        error_label = tk.Label(form, textvariable=error_var, font=(FONT_FAMILY, 9),
                               bg=COLORS['bg'], fg=COLORS['danger'], wraplength=500, justify=LEFT)
        error_label.pack(anchor=W, pady='0 5')

        fields = {}

        def add_field(label, hint="", default=""):
            row = tk.Frame(form, bg=COLORS['bg'])
            row.pack(fill=X, pady=6)
            tk.Label(row, text=label, font=(FONT_FAMILY, 10, "bold"),
                     bg=COLORS['bg'], fg=COLORS['text'], width=12, anchor=E).pack(side=LEFT)
            var = tk.StringVar(value=default)
            entry = tk.Entry(row, textvariable=var, font=(FONT_FAMILY, 10),
                             relief=FLAT, bg='#F0F2F5', fg=COLORS['text'], insertbackground=COLORS['text'])
            entry.pack(side=LEFT, fill=X, expand=True, ipady=5, padx='10 0')
            if hint:
                tk.Label(row, text=hint, font=(FONT_FAMILY, 8),
                         bg=COLORS['bg'], fg=COLORS['text_secondary']).pack(side=LEFT, padx='8 0')
            return var

        fields['name'] = add_field("规则名称:", "如：定稿报告", rule.name if rule else "")
        fields['contain'] = add_field("包含关键词:", "逗号分隔，全部需满足", rule.contain_keywords if rule else "")
        fields['exclude'] = add_field("排除关键词:", "逗号分隔，任一即排除", rule.exclude_keywords if rule else "")

        # ---- 目标文件夹（内联式，不弹窗） ----
        tk.Label(form, text="目标文件夹:", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS['bg'], fg=COLORS['text'], anchor=E).pack(anchor=W)

        folder_var = tk.StringVar(value=rule.target_folder if rule else "")
        folder_entry = tk.Entry(form, textvariable=folder_var, font=(FONT_FAMILY, 10),
                                relief=FLAT, bg='#F0F2F5', fg=COLORS['text'], insertbackground=COLORS['text'])
        folder_entry.pack(fill=X, ipady=5, pady='3 0')

        # 快捷路径按钮
        quick_frame = tk.Frame(form, bg=COLORS['bg'])
        quick_frame.pack(fill=X, pady=3)

        home = os.path.expanduser("~")
        shortcuts = [("主目录", home), ("桌面", self.organizer.get_desktop_path())]
        for name, path in [("/data", "/data"), ("/home", "/home"), ("/tmp", "/tmp")]:
            if os.path.isdir(path):
                shortcuts.append((name, path))

        def set_folder(p):
            folder_var.set(p)

        for label, path in shortcuts:
            if path and os.path.isdir(path):
                tk.Button(quick_frame, text=label, font=(FONT_FAMILY, 8),
                          bg='#E1E4E8', fg=COLORS['text'], relief=FLAT, padx=8, pady=1,
                          cursor="hand2",
                          command=lambda p=path: set_folder(p)).pack(side=LEFT, padx=2)

        # 浏览按钮（内联，不弹新窗口）
        def browse_inline():
            try:
                from tkinter import filedialog
                path = filedialog.askdirectory(title="选择文件夹", parent=editor,
                                              initialdir=folder_var.get())
                if path:
                    folder_var.set(path)
            except Exception:
                error_var.set("系统文件对话框不可用，请手动输入路径或使用上方快捷按钮")

        tk.Button(quick_frame, text="浏览...", font=(FONT_FAMILY, 8),
                  bg=COLORS['primary'], fg='white', relief=FLAT, padx=8, pady=1,
                  cursor="hand2", command=browse_inline).pack(side=LEFT, padx=2)

        # 路径状态
        folder_status_var = tk.StringVar(value="")
        folder_status_label = tk.Label(form, textvariable=folder_status_var, font=(FONT_FAMILY, 8),
                                       bg=COLORS['bg'], fg=COLORS['text_secondary'])
        folder_status_label.pack(anchor=W, pady='2 0')

        def check_folder(*args):
            p = folder_var.get().strip()
            if not p:
                folder_status_var.set("")
            elif os.path.isdir(p):
                folder_status_var.set(f"[OK] 路径有效")
            else:
                folder_status_var.set(f"路径不存在，保存时将自动创建")

        folder_var.trace_add('write', check_folder)
        check_folder()

        # 自动创建
        auto_create_var = tk.BooleanVar(value=True)
        tk.Checkbutton(form, text="保存时自动创建文件夹（如不存在）", variable=auto_create_var,
                       font=(FONT_FAMILY, 9), bg=COLORS['bg'],
                       fg=COLORS['text'], selectcolor=COLORS['bg'],
                       activebackground=COLORS['bg']).pack(anchor=W, pady='5 0')

        # 优先级
        priority_row = tk.Frame(form, bg=COLORS['bg'])
        priority_row.pack(fill=X, pady=6)
        tk.Label(priority_row, text="优先级:", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS['bg'], fg=COLORS['text'], width=12, anchor=E).pack(side=LEFT)
        priority_var = tk.IntVar(value=rule.priority if rule else 0)
        tk.Spinbox(priority_row, from_=0, to=99, textvariable=priority_var,
                   font=(FONT_FAMILY, 10), width=5).pack(side=LEFT, padx='10 0')
        tk.Label(priority_row, text="数字越小优先级越高", font=(FONT_FAMILY, 8),
                 bg=COLORS['bg'], fg=COLORS['text_secondary']).pack(side=LEFT, padx='8 0')

        # 启用
        enabled_var = tk.BooleanVar(value=rule.enabled if rule else True)
        tk.Checkbutton(form, text="启用此规则", variable=enabled_var,
                       font=(FONT_FAMILY, 10), bg=COLORS['bg'],
                       fg=COLORS['text'], selectcolor=COLORS['bg']).pack(anchor=W, pady='10 0')

        # 说明
        help_text = (
            "规则说明：\n"
            "· 包含关键词：文件名必须同时包含所有关键词（AND关系）\n"
            "· 排除关键词：文件名包含任一关键词则不匹配（OR关系）\n"
            "· 匹配时仅检查文件名（不含扩展名）\n"
            "· 同一文件匹配多个规则时，使用优先级最高的规则"
        )
        tk.Label(form, text=help_text, font=(FONT_FAMILY, 8),
                 bg=COLORS['bg'], fg=COLORS['text_secondary'], justify=LEFT,
                 wraplength=480).pack(anchor=W, pady='10 0')

        # 回车保存
        editor.bind('<Return>', lambda e: save())

    def _delete_rule(self, rule: Rule, refresh_fn):
        """删除规则"""
        if self._confirm("确认删除", f"确定要删除规则「{rule.name}」吗？\n此操作不可撤销。"):
            self.db.delete_rule(rule.id)
            refresh_fn()
            self._refresh_folders()

    # ---- 设置对话框 ----
    def _show_settings(self):
        """显示设置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("系统设置")
        dialog.geometry("520x550")
        dialog.transient(self.root)
        dialog.resizable(False, False)

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 520) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 550) // 2
        dialog.geometry(f"+{x}+{y}")

        title_bar = tk.Frame(dialog, bg=COLORS['primary'], height=45)
        title_bar.pack(fill=X)
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text="系统设置", font=(FONT_FAMILY, 12, "bold"),
                 bg=COLORS['primary'], fg='white').pack(side=LEFT, padx=15, pady=8)

        # 可滚动区域
        canvas = tk.Canvas(dialog, bg=COLORS['bg'], highlightthickness=0)
        scrollbar = tk.Scrollbar(dialog, orient=VERTICAL, command=canvas.yview)
        form = tk.Frame(canvas, bg=COLORS['bg'], padx=25, pady=15)
        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor=NW)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # 桌面路径
        row1 = tk.Frame(form, bg=COLORS['bg'])
        row1.pack(fill=X, pady=8)
        tk.Label(row1, text="桌面路径:", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS['bg'], fg=COLORS['text'], width=12, anchor=E).pack(side=LEFT)
        desktop_var = tk.StringVar(value=self.db.get_setting("desktop_path", self.organizer.get_desktop_path()))
        desktop_entry = tk.Entry(row1, textvariable=desktop_var, font=(FONT_FAMILY, 10),
                                 relief=FLAT, bg='#F0F2F5', fg=COLORS['text'])
        desktop_entry.pack(side=LEFT, fill=X, expand=True, ipady=5, padx='10 0')
        tk.Button(row1, text="浏览", font=(FONT_FAMILY, 9),
                  bg=COLORS['primary'], fg='white', relief=FLAT, padx=10, cursor="hand2",
                  command=lambda: _browse()).pack(side=LEFT, padx='8 0')

        def _browse():
            path = select_folder_dialog("选择桌面文件夹", parent=dialog)
            if path:
                desktop_var.set(path)

        # 数据库位置
        row2 = tk.Frame(form, bg=COLORS['bg'])
        row2.pack(fill=X, pady=8)
        tk.Label(row2, text="数据库位置:", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS['bg'], fg=COLORS['text'], width=12, anchor=E).pack(side=LEFT)
        tk.Label(row2, text=self.db.db_path, font=(FONT_FAMILY, 9),
                 bg=COLORS['bg'], fg=COLORS['text_secondary']).pack(side=LEFT, padx='10 0')

        # 支持格式
        row3 = tk.Frame(form, bg=COLORS['bg'])
        row3.pack(fill=X, pady=8)
        tk.Label(row3, text="支持格式:", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS['bg'], fg=COLORS['text'], width=12, anchor=E).pack(side=LEFT)
        formats = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        tk.Label(row3, text=formats, font=(FONT_FAMILY, 9),
                 bg=COLORS['bg'], fg=COLORS['text_secondary']).pack(side=LEFT, padx='10 0')

        # ---- 依赖状态 ----
        tk.Label(form, text="内容提取依赖状态:", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS['bg'], fg=COLORS['text']).pack(anchor=W, pady='15 5')

        deps = check_dependencies()
        dep_frame = tk.Frame(form, bg=COLORS['bg'])
        dep_frame.pack(fill=X)

        for name, installed in deps.items():
            row = tk.Frame(dep_frame, bg=COLORS['bg'])
            row.pack(fill=X, pady=1)
            status = "[OK]" if installed else "[未安装-使用备用方案]"
            color = '#27AE60' if installed else '#E67E22'
            tk.Label(row, text=status, font=("TkDefaultFont", 9, "bold"),
                     bg=COLORS['bg'], fg=color, width=18, anchor=W).pack(side=LEFT)
            tk.Label(row, text=name, font=("TkDefaultFont", 9),
                     bg=COLORS['bg'], fg=COLORS['text']).pack(side=LEFT)

        # ---- 测试提取 ----
        tk.Label(form, text="测试文件内容提取:", font=(FONT_FAMILY, 10, "bold"),
                 bg=COLORS['bg'], fg=COLORS['text']).pack(anchor=W, pady='15 5')

        test_frame = tk.Frame(form, bg=COLORS['bg'])
        test_frame.pack(fill=X)
        test_var = tk.StringVar()
        test_entry = tk.Entry(test_frame, textvariable=test_var, font=(FONT_FAMILY, 9),
                              relief=FLAT, bg='#F0F2F5', fg=COLORS['text'])
        test_entry.pack(side=LEFT, fill=X, expand=True, ipady=4, padx='0 5')

        def browse_test():
            from tkinter import filedialog
            path = filedialog.askopenfilename(
                title="选择要测试的文件",
                filetypes=[("所有文件", "*.*"), ("文档", "*.docx *.doc *.pdf *.xlsx *.xls *.pptx *.ppt *.wps *.et")]
            )
            if path:
                test_var.set(path)

        tk.Button(test_frame, text="选择文件", font=(FONT_FAMILY, 9),
                  bg=COLORS['primary'], fg='white', relief=FLAT, padx=8, cursor="hand2",
                  command=browse_test).pack(side=LEFT, padx='0 5')

        test_result_var = tk.StringVar(value="")
        test_result_label = tk.Label(form, textvariable=test_result_var, font=("TkDefaultFont", 9),
                                     bg=COLORS['bg'], fg=COLORS['text_secondary'],
                                     wraplength=460, justify=LEFT)
        test_result_label.pack(anchor=W, pady='5 0')

        def do_test():
            path = test_var.get().strip()
            if not path:
                test_result_var.set("请先选择一个文件")
                return
            test_result_var.set("正在提取...")
            dialog.update()

            def run_test():
                result = test_extract(path)
                dialog.after(0, lambda: test_result_var.set(result))

            threading.Thread(target=run_test, daemon=True).start()

        tk.Button(test_frame, text="测试提取", font=(FONT_FAMILY, 9),
                  bg='#F39C12', fg='white', relief=FLAT, padx=8, cursor="hand2",
                  command=do_test).pack(side=LEFT)

        # 清空数据
        row4 = tk.Frame(form, bg=COLORS['bg'])
        row4.pack(fill=X, pady=15)

        def clear_data():
            if self._confirm("确认清空", "确定要清空所有文档记录吗？\n（不会删除实际文件）"):
                conn = self.db._get_conn()
                conn.execute("DELETE FROM documents")
                conn.commit()
                self._refresh_all()

        tk.Button(row4, text="清空文档记录", font=(FONT_FAMILY, 10),
                  bg=COLORS['danger'], fg='white', relief=FLAT, padx=15, pady=5,
                  cursor="hand2", command=clear_data).pack(side=LEFT)

        # 保存按钮
        btn_frame = tk.Frame(dialog, bg=COLORS['bg'])
        btn_frame.pack(fill=X, padx=25, pady=15)

        def save():
            self.db.set_setting("desktop_path", desktop_var.get().strip())
            dialog.destroy()
            self._refresh_all()
            self.status_label.config(text="设置已保存")

        tk.Button(btn_frame, text="保存设置", font=(FONT_FAMILY, 10, "bold"),
                  bg=COLORS['success'], fg='white', relief=FLAT, padx=25, pady=5,
                  cursor="hand2", command=save).pack(side=RIGHT, padx=5)

        tk.Button(btn_frame, text="关闭", font=(FONT_FAMILY, 10),
                  bg='#95A5A6', fg='white', relief=FLAT, padx=25, pady=5,
                  cursor="hand2", command=dialog.destroy).pack(side=RIGHT, padx=5)

    # ---- 工具方法 ----
    def _show_message(self, title: str, message: str, msg_type: str = "info", parent=None):
        """显示消息框"""
        icons = {"info": "[i]", "success": "[OK]", "warning": "[!]", "error": "[X]"}
        icon = icons.get(msg_type, "[i]")

        dialog = tk.Toplevel(parent or self.root)
        dialog.title(title)
        dialog.geometry("420x250")
        dialog.transient(parent or self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 420) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 250) // 2
        dialog.geometry(f"+{x}+{y}")

        tk.Label(dialog, text=icon, font=(FONT_FAMILY, 24), bg=COLORS['card_bg']).pack(pady='20 5')
        tk.Label(dialog, text=title, font=(FONT_FAMILY, 13, "bold"),
                 bg=COLORS['card_bg'], fg=COLORS['text']).pack()
        tk.Label(dialog, text=message, font=(FONT_FAMILY, 10),
                 bg=COLORS['card_bg'], fg=COLORS['text_secondary'], wraplength=360,
                 justify=LEFT).pack(pady=10, padx=30)

        tk.Button(dialog, text="确定", font=(FONT_FAMILY, 10),
                  bg=COLORS['primary'], fg='white', relief=FLAT, padx=30, pady=5,
                  cursor="hand2", command=dialog.destroy).pack(pady=10)

    def _confirm(self, title: str, message: str) -> bool:
        """确认对话框"""
        result = [False]

        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 400) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 200) // 2
        dialog.geometry(f"+{x}+{y}")

        tk.Label(dialog, text="[?]", font=(FONT_FAMILY, 20), bg=COLORS['card_bg']).pack(pady='15 5')
        tk.Label(dialog, text=message, font=(FONT_FAMILY, 10),
                 bg=COLORS['card_bg'], fg=COLORS['text'], wraplength=340,
                 justify=LEFT).pack(pady=5, padx=25)

        btn_frame = tk.Frame(dialog, bg=COLORS['card_bg'])
        btn_frame.pack(pady=15)

        def yes():
            result[0] = True
            dialog.destroy()

        def no():
            dialog.destroy()

        tk.Button(btn_frame, text="确定", font=(FONT_FAMILY, 10),
                  bg=COLORS['primary'], fg='white', relief=FLAT, padx=25, pady=4,
                  cursor="hand2", command=yes).pack(side=LEFT, padx=10)

        tk.Button(btn_frame, text="取消", font=(FONT_FAMILY, 10),
                  bg='#95A5A6', fg='white', relief=FLAT, padx=25, pady=4,
                  cursor="hand2", command=no).pack(side=LEFT, padx=10)

        self.root.wait_window(dialog)
        return result[0]

    def run(self):
        """运行应用"""
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        """关闭应用"""
        self.watcher.stop()
        self.root.destroy()


# ==================== 启动入口 ====================
if __name__ == "__main__":
    app = DocManagerApp()
    app.run()
