import tkinter as tk
from tkinter import ttk
from typing import List, Dict

class FormulaPage(ttk.Frame):
    """公式列表页面"""
    def __init__(self, parent, config, formula_manager, excel_manager):
        super().__init__(parent)
        self.parent = parent
        self.config = config
        self.formula_manager = formula_manager
        self.excel_manager = excel_manager
        
        # 初始化事件回调
        self.callbacks = {}
        
        # 创建界面组件
        self.create_widgets()
        
        # 初始化公式列表
        self.refresh_formula_list()
        
    def create_widgets(self):
        """创建界面组件"""
        # 创建主框架
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 搜索框架
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, padx=10, pady=(5,10))
        
        # 搜索图标和输入框
        search_icon = ttk.Label(
            search_frame, 
            text="🔍",
            font=('Segoe UI', 10)
        )
        search_icon.pack(side=tk.LEFT, padx=(0,5))
        
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=('Microsoft YaHei UI', 10),
            width=40
        )
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 清除搜索按钮
        self.clear_btn = ttk.Button(
            search_frame,
            text="✕",
            width=3,
            command=self.clear_search,
            style='Clear.TButton'
        )
        self.clear_btn.pack(side=tk.LEFT, padx=(5,0))
        self.clear_btn.pack_forget()  # 初始隐藏
        
        # 创建树形视图框架
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建树形视图样式
        style = ttk.Style()
        
        # 设置基本样式
        style.configure(
            'Formula.Treeview',
            background='white',
            fieldbackground='white',
            rowheight=35,  # 增加行高
            font=('Microsoft YaHei UI', 10)
        )
        
        # 设置标题样式
        style.configure(
            'Formula.Treeview.Heading',
            font=('Microsoft YaHei UI', 11, 'bold'),
            background='#f8f9fa',
            foreground='#2c3e50'
        )
        
        # 设置选中项样式
        style.map('Formula.Treeview',
            background=[
                ('selected', '#e3f2fd'),  # 浅蓝色背景
                ('hover', '#f5f5f5')  # 鼠标悬停效果
            ],
            foreground=[
                ('selected', '#1976d2'),  # 深蓝色文字
                ('hover', '#2196f3')  # 鼠标悬停文字颜色
            ]
        )
        
        # 清除按钮样式
        style.configure('Clear.TButton',
            padding=2,
            relief='flat',
            font=('Segoe UI', 8)
        )
        
        # 公式树形视图
        self.formula_tree = ttk.Treeview(
            tree_frame,
            columns=("description",),
            displaycolumns=("description",),
            selectmode='browse',
            style='Formula.Treeview'
        )
        
        # 设置列标题
        self.formula_tree.heading(
            '#0', 
            text='公式名称',
            anchor=tk.W
        )
        self.formula_tree.heading(
            'description',
            text='说明',
            anchor=tk.W
        )
        
        # 设置列宽度和最小宽度
        self.formula_tree.column(
            '#0',
            width=200,
            minwidth=150,
            stretch=False  # 固定宽度
        )
        self.formula_tree.column(
            'description',
            width=500,
            minwidth=300,
            stretch=True  # 自动拉伸
        )
        
        # 添加滚动条
        vsb = ttk.Scrollbar(
            tree_frame, 
            orient=tk.VERTICAL,
            command=self.formula_tree.yview
        )
        hsb = ttk.Scrollbar(
            tree_frame,
            orient=tk.HORIZONTAL,
            command=self.formula_tree.xview
        )
        
        self.formula_tree.configure(
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )
        
        # 布局
        self.formula_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # 配置grid权重
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # 设置标签样式
        self.formula_tree.tag_configure(
            'category',
            font=('Microsoft YaHei UI', 11, 'bold'),
            background='#f8f9fa',
            foreground='#37474f'
        )
        
        self.formula_tree.tag_configure(
            'formula',
            font=('Microsoft YaHei UI', 10)
        )
        
        # 添加交替行颜色
        self.formula_tree.tag_configure(
            'odd_row',
            background='#ffffff'
        )
        self.formula_tree.tag_configure(
            'even_row',
            background='#f8f9fa'
        )
        
        # 绑定事件
        self.search_entry.bind('<KeyRelease>', self.on_search)
        self.search_var.trace('w', self.on_search_changed)
        self.formula_tree.bind('<Double-1>', self.on_formula_double_click)
        self.formula_tree.bind('<<TreeviewSelect>>', self.on_formula_select)

    def bind_events(self, callbacks):
        """绑定事件回调"""
        self.callbacks = callbacks
        
        # 搜索框事件
        self.search_var.trace('w', self.on_search_changed)
        
        # 公式树事件
        self.formula_tree.bind('<<TreeviewSelect>>', self.on_formula_selected)
        self.formula_tree.bind('<Double-1>', self.on_formula_double_click)
        
    def on_search(self, event=None):
        """搜索框按键事件处理"""
        # 取消之前的延迟搜索
        if hasattr(self, 'search_after_id'):
            self.after_cancel(self.search_after_id)
        
        # 设置新的延迟搜索
        delay = self.config.get('formula.search_delay', 500)
        self.search_after_id = self.after(delay, self.do_search)

    def do_search(self):
        """执行搜索"""
        try:
            search_text = self.search_var.get().strip().lower()
            
            # 清空当前显示
            for item in self.formula_tree.get_children():
                self.formula_tree.delete(item)
            
            if not search_text:
                # 如果搜索框为空，显示所有公式
                self.refresh_formula_list()
                return
            
            # 搜索公式
            categories = self.formula_manager.get_categories()
            row_count = 0
            for category in categories:
                category_items = []
                formulas = self.formula_manager.get_formulas(category)
                
                for formula in formulas:
                    # 处理不同的公式数据格式
                    if isinstance(formula, dict):
                        name = formula.get('name', '').lower()
                        description = formula.get('description', '').lower()
                    else:
                        name = formula.lower()
                        description = self.formula_manager.get_formula_description(formula).lower()
                    
                    # 如果公式名称或描述包含搜索文本
                    if search_text in name or search_text in description:
                        category_items.append(formula)
                
                # 如果该分类下有匹配的公式
                if category_items:
                    # 添加分类节点
                    category_id = self.formula_tree.insert(
                        '', 'end',
                        text=category,
                        tags=('category',)
                    )
                    
                    # 添加匹配的公式
                    for formula in category_items:
                        row_count += 1
                        row_tags = ('formula', 'odd_row' if row_count % 2 else 'even_row')
                        
                        if isinstance(formula, dict):
                            name = formula.get('name', '')
                            description = formula.get('description', '')
                        else:
                            name = formula
                            description = self.formula_manager.get_formula_description(formula)
                        
                        self.formula_tree.insert(
                            category_id, 'end',
                            text=name,
                            values=(description,),
                            tags=row_tags
                        )
                    
                    # 展开分类
                    self.formula_tree.item(category_id, open=True)
            
        except Exception as e:
            print(f"搜索失败: {str(e)}")

    def on_formula_selected(self, event):
        """公式选择事件处理"""
        try:
            # 获取选中项
            selection = self.formula_tree.selection()
            if not selection:
                return
            
            # 获取选中项的信息
            item = selection[0]
            
            # 如果点击的是公式（有父节点）
            if self.formula_tree.parent(item):
                formula_name = self.formula_tree.item(item)['text']
                if 'on_formula_selected' in self.callbacks:
                    self.callbacks['on_formula_selected'](formula_name)
                
        except Exception as e:
            print(f"处理公式选择失败: {str(e)}")

    def on_formula_double_click(self, event):
        """公式双击事件处理"""
        try:
            # 获取点击的项
            item = self.formula_tree.identify('item', event.x, event.y)
            if not item:
                return
            
            # 如果点击的是公式（有父节点）
            if self.formula_tree.parent(item):
                formula_name = self.formula_tree.item(item)['text']
                if 'on_formula_selected' in self.callbacks:
                    self.callbacks['on_formula_selected'](formula_name)
                
        except Exception as e:
            print(f"处理公式双击失败: {str(e)}")

    def on_open_excel(self):
        """打开Excel按钮事件处理"""
        if 'on_excel_open' in self.callbacks:
            self.callbacks['on_excel_open']()

    def update_file_status(self, text: str, color: str = "black"):
        """更新文件状态
        Args:
            text: 状态文本
            color: 文本颜色
        """
        try:
            # 更新工作簿选择状态
            if not text or "未打开" in text or "失败" in text:
                self.workbook_combo['state'] = 'disabled'
                self.workbook_var.set('')
            else:
                self.workbook_combo['state'] = 'readonly'
            
        except Exception as e:
            print(f"更新文件状态失败: {str(e)}")

    def update_excel_status(self, is_connected: bool):
        """更新Excel连接状态
        Args:
            is_connected: 是否已连接
        """
        try:
            if is_connected:
                self.open_btn['state'] = 'disabled'
                self.refresh_btn['state'] = 'normal'
                self.workbook_combo['state'] = 'readonly'
            else:
                self.open_btn['state'] = 'normal'
                self.refresh_btn['state'] = 'disabled'
                self.workbook_combo['state'] = 'disabled'
                self.workbook_var.set('')
            
        except Exception as e:
            print(f"更新Excel状态失败: {str(e)}")

    def update_formula_list(self, formulas: List[Dict]):
        """更新公式列表
        Args:
            formulas: 公式列表
        """
        try:
            # 清空现有项目
            self.formula_tree.delete(*self.formula_tree.get_children())
            
            if not formulas:
                return
            
            # 按分类组织公式
            categories = {}
            for formula in formulas:
                category = formula['category']
                if category not in categories:
                    categories[category] = []
                categories[category].append(formula)
            
            # 添加到树形视图
            for category, formula_list in categories.items():
                category_id = self.formula_tree.insert(
                    '', 'end',
                    text=category,
                    tags=('category',)
                )
                
                for formula in formula_list:
                    self.formula_tree.insert(
                        category_id, 'end',
                        text=formula['name'],
                        values=(formula['description'],),
                        tags=('formula',)
                    )
                    
                # 展开分类
                self.formula_tree.item(category_id, open=True)
            
        except Exception as e:
            print(f"更新公式列表失败: {str(e)}")
            if 'on_status_update' in self.callbacks:
                self.callbacks['on_status_update']("更新公式列表失败")

    def show(self):
        """显示页面"""
        self.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def hide(self):
        """隐藏页面"""
        self.pack_forget()

    def refresh_formula_list(self):
        """刷新公式列表"""
        try:
            # 清空现有项目
            for item in self.formula_tree.get_children():
                self.formula_tree.delete(item)
            
            # 获取所有分类
            categories = self.formula_manager.get_categories()
            if not categories:
                print("没有找到公式分类")
                return
            
            # 添加分类和公式
            row_count = 0
            for category in categories:
                # 添加分类
                category_id = self.formula_tree.insert(
                    '', 'end',
                    text=category,
                    tags=('category',)
                )
                
                # 获取该分类下的公式
                formulas = self.formula_manager.get_formulas(category)
                if not formulas:
                    continue
                
                # 添加公式
                for formula in formulas:
                    row_count += 1
                    row_tags = ('formula', 'odd_row' if row_count % 2 else 'even_row')
                    
                    # 处理不同的公式数据格式
                    if isinstance(formula, dict):
                        name = formula.get('name', '')
                        description = formula.get('description', '')
                    else:
                        name = formula
                        description = self.formula_manager.get_formula_description(formula)
                    
                    self.formula_tree.insert(
                        category_id, 'end',
                        text=name,
                        values=(description,),
                        tags=row_tags
                    )
                
                # 根据配置决定是否自动展开
                if self.config.get('formula.auto_expand', True):
                    self.formula_tree.item(category_id, open=True)
                
        except Exception as e:
            print(f"刷新公式列表失败: {str(e)}")

    def on_search_changed(self, *args):
        """搜索框内容变化时的处理"""
        # 显示/隐藏清除按钮
        if self.search_var.get():
            self.clear_btn.pack(side=tk.LEFT, padx=(5,0))
        else:
            self.clear_btn.pack_forget()
        
        # 触发搜索
        self.on_search()

    def clear_search(self):
        """清除搜索"""
        self.search_var.set('')
        self.search_entry.focus()

    def on_formula_select(self, event):
        """公式选择事件处理"""
        try:
            # 获取选中项
            selection = self.formula_tree.selection()
            if not selection:
                return
            
            item = selection[0]
            item_tags = self.formula_tree.item(item)['tags']
            
            # 如果选中的是公式而不是分类
            if 'formula' in item_tags:
                formula_name = self.formula_tree.item(item)['text']
                formula = self.formula_manager.get_formula(formula_name)
                if formula:
                    # 更新状态栏
                    if 'on_status_update' in self.callbacks:
                        self.callbacks['on_status_update'](f"已选择: {formula_name}")
    
        except Exception as e:
            print(f"选择公式失败: {str(e)}")

    def on_formula_double_click(self, event):
        """公式双击事件处理"""
        try:
            # 获取点击的项
            item = self.formula_tree.identify('item', event.x, event.y)
            if not item:
                return
            
            item_tags = self.formula_tree.item(item)['tags']
            
            # 如果双击的是公式而不是分类
            if 'formula' in item_tags:
                formula_name = self.formula_tree.item(item)['text']
                
                # 如果设置了回调函数，调用它
                if 'on_formula_selected' in self.callbacks:
                    self.callbacks['on_formula_selected'](formula_name)
                    
                # 更新状态栏
                if 'on_status_update' in self.callbacks:
                    self.callbacks['on_status_update'](f"已选择: {formula_name}")
    
        except Exception as e:
            print(f"双击公式失败: {str(e)}") 