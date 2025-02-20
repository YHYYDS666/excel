import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from utils.formula_manager import FormulaManager
from utils.excel_manager import ExcelManager
from .window_manager import WindowManager
from .status_manager import StatusManager
from .toolbar_manager import ToolbarManager
from utils.history_manager import HistoryManager
from .formula_page import FormulaPage
from .apply_page import ApplyPage
import os
import sys
import logging

class MainWindow(WindowManager):
    def __init__(self, root, config):
        super().__init__(root, config)
        
        # 设置窗口属性
        self.setup_window()
        
        # 添加置顶状态变量
        self.is_topmost = tk.BooleanVar(value=False)
        
        # 初始化管理器
        self.excel_manager = ExcelManager()
        self.formula_manager = FormulaManager()
        self.history_manager = HistoryManager()
        self.toolbar_manager = ToolbarManager(root)  # 先创建工具栏
        self.status_manager = StatusManager(root)    # 再创建状态栏
        
        # 初始化页面
        self.init_pages()
        
        # 绑定工具栏命令
        self.toolbar_manager.bind_commands({
            'open': self.open_excel_file,
            'refresh': self.refresh_workbooks,
            'undo': self.undo,
            'redo': self.redo,
            'reset': self.reset,
            'workbook_selected': self.on_workbook_selected
        })
        
        # 绑定事件
        self.bind_events()
        
        # 显示公式列表页面并刷新
        self.show_formula_page()
        self.formula_page.refresh_formula_list()
        
        # 创建置顶按钮
        self.create_pin_button()
    
    def init_pages(self):
        """初始化页面"""
        self.current_page = None
        
        # 创建页面容器
        self.container = ttk.Frame(self.root)
        self.container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建页面但不立即显示
        self.formula_page = FormulaPage(
            parent=self.container,
            config=self.config,
            formula_manager=self.formula_manager,
            excel_manager=self.excel_manager
        )
        
        self.apply_page = ApplyPage(
            parent=self.container,
            config=self.config,
            formula_manager=self.formula_manager,
            excel_manager=self.excel_manager
        )
        
        # 绑定页面事件
        self.formula_page.bind_events({
            'on_formula_selected': self.on_formula_selected,
            'on_excel_open': self.open_excel_file,
            'on_workbook_refresh': self.refresh_workbooks
        })
        
        self.apply_page.bind_events({
            'on_back': self.show_formula_page,
            'on_apply': self.apply_formula,
            'on_range_selected': self.on_range_selected
        })
        
        # 定时检查Excel状态
        self.check_excel_status()
    
    def bind_events(self):
        """绑定事件"""
        # 窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 快捷键绑定
        self.bind_shortcuts()
        
        # 工具提示
        self.create_tooltips()
    
    def bind_shortcuts(self):
        """绑定快捷键"""
        self.root.bind('<Control-z>', lambda e: self.undo())
        self.root.bind('<Control-y>', lambda e: self.redo())
        self.root.bind('<Control-Shift-Z>', lambda e: self.reset())
        self.root.bind('<Control-o>', lambda e: self.open_excel_file())
        self.root.bind('<Control-r>', lambda e: self.refresh_workbooks())
    
    def create_tooltips(self):
        """创建工具提示"""
        tooltips = {
            self.toolbar_manager.open_btn: "打开Excel文件 (Ctrl+O)",
            self.toolbar_manager.undo_btn: "撤销上一步操作 (Ctrl+Z)",
            self.toolbar_manager.redo_btn: "恢复已撤销的操作 (Ctrl+Y)",
            self.toolbar_manager.reset_btn: "重置到初始状态 (Ctrl+Shift+Z)",
            self.toolbar_manager.workbook_combo: "选择要操作的工作簿",
            self.formula_page.search_entry: "搜索公式 (Ctrl+F)",
            self.formula_page.formula_tree: "双击公式可直接应用"
        }
        
        for widget, text in tooltips.items():
            self.create_tooltip(widget, text)

    def open_excel_file(self):
        """打开Excel文件"""
        try:
            # 如果已经连接，询问是否关闭当前文件
            if self.excel_manager.is_connected():
                if messagebox.askyesno(
                    "提示", 
                    "是否关闭当前打开的Excel文件？\n"
                    "选择“是”关闭当前文件并打开新文件，\n"
                    "选择“否”保持当前文件打开。",
                    parent=self.root
                ):
                    self.excel_manager.disconnect()
                    # 更新工具栏状态
                    self.toolbar_manager.update_excel_status(False)
                    self.toolbar_manager.update_workbook_list({})
            
            # 打开文件选择对话框
            file_path = filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[
                    ("Excel文件", "*.xlsx;*.xls"),
                    ("所有文件", "*.*")
                ]
            )
            
            if file_path:
                # 更新最近文件列表
                self.update_recent_files(file_path)
                
                # 打开Excel文件
                if self.excel_manager.open_file(file_path):
                    # 刷新工作簿列表
                    workbooks = self.excel_manager.refresh_workbooks()
                    if workbooks:
                        self.toolbar_manager.update_workbook_list(workbooks)
                        self.toolbar_manager.update_excel_status(True)
                        self.status_manager.update_status(f"已打开: {os.path.basename(file_path)}")
                        
                        # 显示提示对话框
                        messagebox.showinfo(
                            "提示",
                            "Excel文件已打开，请从左侧列表选择要使用的公式。\n"
                            "双击公式可以直接应用。",
                            parent=self.root
                        )
                    else:
                        self.status_manager.update_status("打开Excel文件失败")
                else:
                    self.status_manager.update_status("打开Excel文件失败")
            
        except Exception as e:
            print(f"打开文件失败: {str(e)}")
            self.status_manager.update_status("打开文件失败")

    def refresh_workbooks(self):
        """刷新工作簿列表"""
        try:
            if not self.excel_manager.is_connected():
                messagebox.showwarning(
                    "提示",
                    "请先打开Excel文件",
                    parent=self.root
                )
                return
            
            workbooks = self.excel_manager.refresh_workbooks()
            if workbooks:
                self.toolbar_manager.update_workbook_list(workbooks)
                self.status_manager.update_status("工作簿列表已刷新")
                
                # 如果在应用公式页面，清除之前的选择
                if self.current_page == self.apply_page:
                    self.apply_page.clear_selection()
            else:
                self.status_manager.update_status("没有找到工作簿")
            
        except Exception as e:
            print(f"刷新工作簿失败: {str(e)}")
            self.status_manager.update_status("刷新工作簿失败")

    def on_workbook_selected(self, event):
        """工作簿选择事件处理"""
        try:
            workbook_name = self.toolbar_manager.workbook_var.get()
            if workbook_name:
                # 获取工作簿对象
                workbook = self.excel_manager.get_workbook(workbook_name)
                if workbook:
                    # 激活工作簿
                    workbook.Activate()
                    self.excel_manager.active_workbook = workbook
                    self.excel_manager.active_sheet = workbook.ActiveSheet
                    
                    # 更新状态
                    self.status_manager.update_status(f"当前工作簿: {workbook_name}")
                    
                    # 如果在应用公式页面，清除之前的选择
                    if self.current_page == self.apply_page:
                        self.apply_page.clear_selection()
                else:
                    self.status_manager.update_status("切换工作簿失败")
        
        except Exception as e:
            print(f"切换工作簿失败: {str(e)}")
            self.status_manager.update_status("切换工作簿失败")

    def on_formula_selected(self, formula_name):
        """公式选择事件处理"""
        try:
            if not self.excel_manager.is_connected():
                messagebox.showwarning(
                    "提示",
                    "请先打开Excel文件",
                    parent=self.root
                )
                return
            
            # 切换到应用页面
            self.show_apply_page(formula_name)
            
            # 更新状态栏
            self.status_manager.update_status(f"当前公式: {formula_name}")
            
        except Exception as e:
            print(f"选择公式失败: {str(e)}")
            self.status_manager.update_status("选择公式失败")

    def apply_formula(self, formula_name, input_range, output_range=None):
        """应用公式"""
        try:
            if not input_range:
                messagebox.showwarning(
                    "提示",
                    "请先选择输入区域",
                    parent=self.root
                )
                return
            
            # 保存当前状态
            self.save_state()
            
            # 应用公式
            success = self.formula_manager.apply_formula(
                formula_name,
                input_range,
                output_range
            )
            
            if success:
                self.status_manager.update_status("公式应用成功")
                # 清除选择
                self.apply_page.clear_selection()
            else:
                self.status_manager.update_status("公式应用失败")
                messagebox.showerror(
                    "错误",
                    "公式应用失败，请检查选择区域是否正确",
                    parent=self.root
                )
            
        except Exception as e:
            print(f"应用公式失败: {str(e)}")
            self.status_manager.update_status("应用公式失败")
            messagebox.showerror(
                "错误",
                f"应用公式失败: {str(e)}",
                parent=self.root
            )

    def on_range_selected(self, mode, value=None):
        """区域选择事件处理
        Args:
            mode: 选择模式('input'/'output')
            value: 预设值
        """
        try:
            selected_range = self.excel_manager.select_range(mode, value)
            if selected_range:
                # 传入选择模式，以便正确更新UI状态
                self.apply_page.update_selection(selected_range, mode)
                return True
            return False
            
        except Exception as e:
            print(f"选择区域失败: {str(e)}")
            self.status_manager.update_status("选择区域失败")
            return False

    def show_formula_page(self):
        """显示公式列表页面"""
        if self.current_page:
            self.current_page.pack_forget()
        self.formula_page.show()
        self.current_page = self.formula_page

    def show_apply_page(self, formula_name):
        """显示应用公式页面"""
        if self.current_page:
            self.current_page.pack_forget()
        self.apply_page.show()
        self.current_page = self.apply_page
        
        # 更新公式信息
        formula = self.formula_manager.get_formula(formula_name)
        if formula:
            self.apply_page.update_formula_info(formula)
            # 清除之前的选择
            self.apply_page.clear_selection()

    def check_excel_status(self):
        """检查Excel状态"""
        try:
            is_connected = self.excel_manager.is_connected()
            self.toolbar_manager.update_excel_status(is_connected)
            
            # 每5秒检查一次
            self.root.after(5000, self.check_excel_status)
            
        except Exception as e:
            print(f"检查Excel状态失败: {str(e)}")

    def create_tooltip(self, widget, text):
        """创建工具提示"""
        def enter(event):
            try:
                # 获取widget位置
                x = widget.winfo_rootx() + 25
                y = widget.winfo_rooty() + 20
                
                # 创建工具提示窗口
                self.tooltip = tk.Toplevel(widget)
                self.tooltip.wm_overrideredirect(True)
                self.tooltip.wm_geometry(f"+{x}+{y}")
                
                label = ttk.Label(
                    self.tooltip,
                    text=text,
                    justify=tk.LEFT,
                    background="#ffffe0",
                    relief=tk.SOLID,
                    borderwidth=1,
                    padding=5
                )
                label.pack()
                
            except Exception as e:
                print(f"创建工具提示失败: {str(e)}")
                
        def leave(event):
            if hasattr(self, "tooltip"):
                self.tooltip.destroy()
                
        widget.bind('<Enter>', enter)
        widget.bind('<Leave>', leave)

    def save_state(self):
        """保存当前状态"""
        try:
            state = self.excel_manager.get_current_state()
            self.history_manager.save_state(state)
            self.update_history_buttons()
            
        except Exception as e:
            print(f"保存状态失败: {str(e)}")

    def undo(self):
        """撤销操作"""
        try:
            if self.history_manager.can_undo():
                state = self.history_manager.undo()
                self.excel_manager.restore_state(state)
                self.update_history_buttons()
                self.status_manager.update_status("已撤销")
                
        except Exception as e:
            print(f"撤销操作失败: {str(e)}")
            self.status_manager.update_status("撤销失败")

    def redo(self):
        """重做操作"""
        try:
            if self.history_manager.can_redo():
                state = self.history_manager.redo()
                self.excel_manager.restore_state(state)
                self.update_history_buttons()
                self.status_manager.update_status("已重做")
                
        except Exception as e:
            print(f"重做操作失败: {str(e)}")
            self.status_manager.update_status("重做失败")

    def reset(self):
        """重置状态"""
        try:
            if self.history_manager.can_reset():
                state = self.history_manager.reset()
                self.excel_manager.restore_state(state)
                self.update_history_buttons()
                self.status_manager.update_status("已重置")
                
        except Exception as e:
            print(f"重置状态失败: {str(e)}")
            self.status_manager.update_status("重置失败")

    def update_history_buttons(self):
        """更新历史操作按钮状态"""
        self.toolbar_manager.update_button_states({
            'undo': self.history_manager.can_undo(),
            'redo': self.history_manager.can_redo(),
            'reset': self.history_manager.can_reset()
        })

    def save_config(self):
        """保存配置"""
        try:
            # 保存窗口大小
            self.config.set('window.width', self.root.winfo_width())
            self.config.set('window.height', self.root.winfo_height())
            
            # 保存其他配置
            self.config.save()
            
        except Exception as e:
            print(f"保存配置失败: {str(e)}")

    def on_closing(self):
        """窗口关闭事件处理"""
        try:
            # 保存配置
            self.save_config()
            
            # 询问是否保存Excel文件
            if self.excel_manager.is_connected():
                if messagebox.askyesno(
                    "提示", 
                    "是否保存Excel文件的更改？",
                    parent=self.root
                ):
                    for wb in self.excel_manager.workbooks.values():
                        try:
                            wb.Save()
                        except:
                            pass
            
            # 断开Excel连接
            self.excel_manager.disconnect()
            
            # 销毁窗口
            self.root.destroy()
        
        except Exception as e:
            print(f"关闭窗口失败: {str(e)}")
            self.root.destroy()

    def create_pin_button(self):
        """创建置顶按钮"""
        # 创建不同状态的图标
        self.pin_normal = "📌"  # 未置顶状态
        self.pin_active = "📍"  # 置顶状态
        
        self.pin_btn = ttk.Button(
            self.toolbar_manager.toolbar,
            text=self.pin_normal,  # 初始状态使用未置顶图标
            width=3,
            command=self.toggle_topmost,
            style='TButton'  # 使用默认样式
        )
        self.pin_btn.pack(side=tk.RIGHT, padx=2)
        
        # 添加工具提示
        self.create_tooltip(self.pin_btn, "窗口置顶 (Alt+T)")
        
        # 绑定快捷键
        self.root.bind('<Alt-t>', lambda e: self.toggle_topmost())

    def toggle_topmost(self):
        """切换窗口置顶状态"""
        try:
            self.is_topmost.set(not self.is_topmost.get())
            self.root.attributes('-topmost', self.is_topmost.get())
            
            # 更新按钮图标和样式
            if self.is_topmost.get():
                self.pin_btn.configure(
                    text=self.pin_active,  # 切换为置顶图标
                    style='Accent.TButton'  # 使用突出样式
                )
                self.create_tooltip(self.pin_btn, "取消置顶 (Alt+T)")
            else:
                self.pin_btn.configure(
                    text=self.pin_normal,  # 切换为未置顶图标
                    style='TButton'  # 使用普通样式
                )
                self.create_tooltip(self.pin_btn, "窗口置顶 (Alt+T)")
            
        except Exception as e:
            print(f"切换窗口置顶失败: {str(e)}")

    def setup_window(self):
        """设置窗口属性"""
        # 创建自定义样式
        style = ttk.Style()
        
        # 使用系统默认主题
        if sys.platform.startswith('win'):
            style.theme_use('vista')
        else:
            style.theme_use('clam')
        
        # 设置自定义样式
        style.configure('TButton', 
            padding=5,
            relief='flat',
            background='#f0f0f0'
        )
        
        style.configure('Accent.TButton',
            padding=5,
            relief='flat',
            background='#007acc',
            foreground='white'
        )
        
        style.configure('TEntry', 
            padding=5,
            relief='flat',
            fieldbackground='white'
        )
        
        style.configure('Treeview', 
            background='white',
            fieldbackground='white',
            rowheight=25,
            relief='flat'
        )
        
        style.configure('Treeview.Heading',
            padding=5,
            relief='flat',
            background='#f0f0f0'
        )
        
        style.configure('TLabel', 
            padding=5
        )
        
        style.configure('TFrame', 
            background='white'
        )
        
        # 添加选择按钮样式
        style.configure('Select.TButton',
            padding=5,
            relief='flat',
            background='#28a745',  # 绿色
            foreground='white'
        )
        
        style.map('Select.TButton',
            background=[('active', '#218838')],  # 鼠标悬停时的颜色
            foreground=[('active', 'white')]
        )
        
        # 设置窗口标题
        self.root.title("Excel公式工具")
        
        # 设置窗口图标（如果存在）
        try:
            if os.path.exists('assets/icon.ico'):
                self.root.iconbitmap('assets/icon.ico')
        except:
            pass
        
        # 设置窗口大小和位置
        window_width = self.config.get('window.width', 800)
        window_height = self.config.get('window.height', 600)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置最小窗口大小
        self.root.minsize(600, 400)
        
        # 设置窗口背景色
        self.root.configure(background='white')

    def update_recent_files(self, file_path):
        """更新最近文件列表"""
        try:
            recent_files = self.config.get('recent_files', [])
            
            # 如果文件已在列表中，移到最前
            if file_path in recent_files:
                recent_files.remove(file_path)
            
            # 添加到列表开头
            recent_files.insert(0, file_path)
            
            # 保持最多10个记录
            recent_files = recent_files[:10]
            
            # 更新配置
            self.config.set('recent_files', recent_files)
            
        except Exception as e:
            print(f"更新最近文件列表失败: {str(e)}")

    def create_file_menu(self):
        """创建文件菜单"""
        file_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="文件", menu=file_menu)
        
        file_menu.add_command(label="打开", command=self.open_excel_file)
        file_menu.add_command(label="刷新", command=self.refresh_workbooks)
        
        # 添加最近文件子菜单
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="最近文件", menu=self.recent_menu)
        
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_closing)

    def show_error(self, title, message, error=None):
        """显示错误对话框"""
        if error:
            message = f"{message}\n\n详细信息：{str(error)}"
        
        messagebox.showerror(
            title,
            message,
            parent=self.root
        )
        
        # 记录错误日志
        if error:
            logging.error(f"{title}: {message}", exc_info=error)
        else:
            logging.error(f"{title}: {message}")