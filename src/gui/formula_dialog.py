import tkinter as tk
import ttkbootstrap as ttk  # 使用ttkbootstrap替代ttk
from tkinter import simpledialog, messagebox
from utils.formula_manager import FormulaManager  # 从utils包导入

class FormulaDialog(tk.Toplevel):
    def __init__(self, parent, title="添加公式", formula_data=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x350")  # 设置窗口大小
        
        # 设置模态和样式
        self.transient(parent)
        self.grab_set()
        
        # 获取主窗口实例
        self.main_window = parent.main_window
        self.formula_manager = self.main_window.formula_manager
        
        # 初始化结果
        self.result = None
        
        # 创建界面
        self.create_widgets(formula_data)
        
        # 居中显示
        self.center_window()
    
    def create_widgets(self, formula_data=None):
        """创建对话框组件"""
        # 主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 分类选择
        ttk.Label(main_frame, text="分类:").grid(row=0, column=0, sticky="w", pady=5)
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(main_frame, textvariable=self.category_var, state="readonly")
        self.category_combo.grid(row=0, column=1, sticky="ew", pady=5)
        
        # 设置分类列表
        categories = self.formula_manager.get_categories()
        self.category_combo['values'] = ["新建分类..."] + categories
        
        # 如果是编辑模式，设置当前分类
        if formula_data and "category" in formula_data:
            self.category_var.set(formula_data["category"])
        
        # 绑定分类选择事件
        self.category_combo.bind('<<ComboboxSelected>>', self.on_category_selected)
        
        # 公式名称
        ttk.Label(main_frame, text="名称:").grid(row=1, column=0, sticky="w", pady=5)
        self.name_var = tk.StringVar(value=formula_data["name"] if formula_data else "")
        self.name_entry = ttk.Entry(main_frame, textvariable=self.name_var)
        self.name_entry.grid(row=1, column=1, sticky="ew", pady=5)
        
        # 公式模板
        ttk.Label(main_frame, text="公式模板:").grid(row=2, column=0, sticky="w", pady=5)
        self.template_text = tk.Text(main_frame, height=5, font=("Microsoft YaHei UI", 9))
        self.template_text.grid(row=2, column=1, sticky="nsew", pady=5)
        if formula_data:
            self.template_text.insert("1.0", formula_data["template"])
        
        # 描述
        ttk.Label(main_frame, text="描述:").grid(row=3, column=0, sticky="w", pady=5)
        self.desc_text = tk.Text(main_frame, height=3, font=("Microsoft YaHei UI", 9))
        self.desc_text.grid(row=3, column=1, sticky="nsew", pady=5)
        if formula_data:
            self.desc_text.insert("1.0", formula_data["description"])
        
        # 提示信息
        ttk.Label(
            main_frame,
            text="提示: 使用{range}作为选区占位符"
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=5)
        
        # 按钮框架
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(
            btn_frame,
            text="确定",
            command=self.on_ok,
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="取消",
            command=self.on_cancel,
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        # 配置网格权重
        main_frame.columnconfigure(1, weight=1)
    
    def center_window(self):
        """将窗口居中显示"""
        # 更新窗口大小
        self.update_idletasks()
        
        # 获取主窗口位置和大小
        parent = self.master
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # 计算居中位置
        width = self.winfo_width()
        height = self.winfo_height()
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        # 设置窗口位置
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def on_category_selected(self, event):
        """分类选择事件处理"""
        if self.category_var.get() == "新建分类...":
            # 弹出对话框让用户输入新分类名称
            new_category = simpledialog.askstring("新建分类", "请输入新分类名称:")
            if new_category:
                # 更新分类列表
                current_values = list(self.category_combo['values'])
                current_values.append(new_category)
                self.category_combo['values'] = current_values
                self.category_var.set(new_category)
            else:
                # 如果用户取消，恢复之前的选择
                if len(self.category_combo['values']) > 1:
                    self.category_var.set(self.category_combo['values'][1])
    
    def on_ok(self):
        """确定按钮回调"""
        category = self.category_var.get()
        name = self.name_var.get().strip()
        template = self.template_text.get("1.0", "end-1c").strip()
        description = self.desc_text.get("1.0", "end-1c").strip()
        
        if not category or category == "新建分类...":
            messagebox.showwarning("警告", "请选择或创建分类！")
            return
        
        if not name:
            messagebox.showwarning("警告", "请输入公式名称！")
            return
        
        if not template:
            messagebox.showwarning("警告", "请输入公式模板！")
            return
        
        self.result = {
            "category": category,
            "name": name,
            "template": template,
            "description": description
        }
        self.destroy()
    
    def on_cancel(self):
        """取消按钮回调"""
        self.destroy() 