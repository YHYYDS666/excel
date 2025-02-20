# Excel公式工具

一个帮助Excel初学者快速应用常用公式的工具。

## 功能特点

- 预置丰富的Excel公式库
- 支持公式分类管理
- 简单直观的公式应用方式
- 支持公式的增删改查
- 实时显示Excel文件状态
- 完整的操作历史管理
- 支持公式名称和描述搜索

## 安装依赖

```bash
pip install -r requirements.txt
```

## 界面特性

- 使用 ttkbootstrap 提供现代化的界面风格
- 支持高分辨率显示和DPI缩放
- 自适应窗口大小和布局
- 统一的视觉设计和交互体验

## 界面布局

### 工具栏区域
- 文件操作：打开Excel、刷新工作簿
- 工作簿选择：切换当前操作的工作簿
- 历史操作：撤销、恢复、重置
- 公式操作：添加、编辑、删除、应用

### 状态栏区域
- 文件状态：显示当前打开的Excel文件
- 操作状态：显示最近的操作和进度
- 历史状态：显示可撤销/恢复的操作数量

### 公式区域
- 搜索框：快速查找公式
- 公式树：按分类显示所有公式
- 支持展开/折叠分类
- 显示公式名称和描述

## 使用方法

1. 运行工具：
```bash
python src/main.py
```

2. 打开Excel文件：
   - 点击"打开Excel"按钮
   - 选择要操作的Excel文件
   - 文件会在下拉框中显示

3. 应用公式：
   - 在Excel中选择要应用公式的区域
   - 在工具中选择要使用的公式
   - 点击"应用公式"按钮

4. 操作历史管理：
   - 撤销(Ctrl+Z)：返回到上一步操作
   - 恢复(Ctrl+Y)：返回到下一步操作
   - 重置(Ctrl+Shift+Z)：恢复到文件打开时的状态

5. 公式管理：
   - 添加：点击"添加公式"，选择分类或创建新分类
   - 编辑：选中公式后点击"编辑公式"修改
   - 删除：选中公式后点击"删除公式"移除
   - 搜索：在搜索框输入关键词快速查找

## 预置公式说明

1. 基础运算：
   - 求和：计算选定区域的和
   - 平均值：计算选定区域的平均值
   - 最大值：获取选定区域的最大值
   - 最小值：获取选定区域的最小值
   - 乘积：计算选定区域的乘积

2. 统计函数：
   - 计数：统计选定区域内数值的个数
   - 文本计数：统计选定区域内非空单元格的个数
   - 空值计数：统计选定区域内空单元格的个数
   - 标准差：计算选定区域的标准差
   - 方差：计算选定区域的方差

3. 条件函数：
   - 条件求和：计算选定区域内大于0的数值之和
   - 条件计数：统计选定区域内大于0的数值个数
   - 条件平均值：计算选定区域内大于0的数值平均值

4. 查找函数：
   - 查找最大值位置：查找最大值在区域中的位置
   - 查找最小值位置：查找最小值在区域中的位置

5. 文本函数：
   - 合并文本：合并选定区域内的文本
   - 提取左侧：提取文本左侧指定个数的字符
   - 提取右侧：提取文本右侧指定个数的字符

6. 百分比：
   - 百分比：计算选定单元格占总和的百分比
   - 百分比格式：以百分比格式显示占比

## 快捷键

### 文件操作
- Ctrl+O：打开Excel文件
- Ctrl+R：刷新工作簿列表

### 历史操作
- Ctrl+Z：撤销到上一步操作
- Ctrl+Y：恢复到下一步操作
- Ctrl+Shift+Z：重置到文件打开时的状态

### 公式操作
- Ctrl+N：添加新公式
- Ctrl+E：编辑选中的公式
- Delete：删除选中的公式
- Ctrl+Enter：应用选中的公式
- Ctrl+F：搜索公式

### 状态显示
- 文件状态：显示当前打开的Excel文件
- 操作状态：显示最近的操作和进度
- 历史状态：显示可撤销/恢复的操作数量

## 注意事项

1. 使用撤销功能前请确保Excel文件未保存，否则无法恢复之前的状态
2. 重置功能会将工作表恢复到文件刚打开时的状态
3. 公式应用前请确保选择了正确的区域
4. 添加新公式时，使用{range}作为选区占位符

## 待开发功能



## 项目结构

```
excel-formula-tool/
├── data/                    # 数据文件目录
│   ├── config.json         # 全局配置文件
│   └── formulas.json       # 预定义公式配置
├── src/                    # 源代码目录
│   ├── gui/               # 界面相关代码
│   │   ├── __init__.py    # GUI模块初始化
│   │   ├── main_window.py # 主窗口实现
│   │   ├── window_manager.py # 窗口基础管理
│   │   ├── toolbar_manager.py # 工具栏管理
│   │   ├── status_manager.py # 状态栏管理
│   │   ├── formula_page.py # 公式列表页面
│   │   ├── apply_page.py  # 公式应用页面
│   │   ├── formula_dialog.py # 公式编辑对话框
│   │   └── float_window.py # 浮动窗口基类
│   ├── utils/             # 工具类模块
│   │   ├── __init__.py    # 工具模块初始化
│   │   ├── config_manager.py # 配置管理
│   │   ├── excel_manager.py # Excel操作管理
│   │   ├── formula_manager.py # 公式管理
│   │   └── history_manager.py # 历史记录管理
│   └── main.py            # 程序入口
├── tests/                 # 测试代码目录
│   ├── __init__.py
│   ├── test_excel.py     # Excel相关测试
│   ├── test_formula.py   # 公式相关测试
│   └── test_gui.py       # 界面相关测试
├── docs/                  # 文档目录
│   ├── api/              # API文档
│   └── user/             # 用户手册
├── README.md             # 项目说明
├── requirements.txt      # 项目依赖
└── setup.py             # 安装配置
```

项目采用模块化设计，主要分为以下几个部分：

1. GUI模块 (src/gui/)
   - WindowManager: 窗口基础属性和行为管理
   - ToolbarManager: 工具栏按钮和事件管理
   - StatusManager: 状态栏显示和更新管理
   - FormulaPage: 公式列表展示和搜索功能
   - ApplyPage: 公式应用和区域选择功能
   - FloatWindow: 浮动窗口基类实现

2. 工具类模块 (src/utils/)
   - ConfigManager: 配置文件读写和管理
   - ExcelManager: Excel文件操作和状态管理
   - FormulaManager: 公式数据的增删改查管理
   - HistoryManager: 操作历史记录管理

3. 数据文件 (data/)
   - config.json: 全局配置信息
   - formulas.json: 预定义公式数据

4. 测试模块 (tests/)
   - 包含单元测试和集成测试
   - 覆盖核心功能和界面交互

5. 文档模块 (docs/)
   - API文档：详细的接口说明
   - 用户手册：使用说明和示例

## 开发环境配置

1. Python版本要求
   - Python 3.7+
   - 64位版本

2. 依赖项安装
```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

3. 开发工具推荐
   - Visual Studio Code
   - PyCharm

4. 代码规范
   - 遵循PEP 8规范
   - 使用类型注解
   - 编写完整的文档字符串

## 调试说明

1. 命令行参数
```bash
# 启用调试模式
python src/main.py --debug

# 指定配置文件
python src/main.py --config custom_config.json

# 设置日志级别
python src/main.py --log-level DEBUG
```

2. 日志文件
- 位置: logs/app.log
- 格式: 时间 - 模块 - 级别 - 消息
- 大小: 单个文件最大1MB
- 数量: 保留最近5个文件

3. 调试技巧
- 查看日志文件了解程序运行状态
- 使用DEBUG级别获取更详细信息
- 检查异常堆栈追踪错误来源

## 贡献指南

1. 提交代码
   - Fork 项目
   - 创建特性分支
   - 提交变更
   - 推送到分支
   - 创建 Pull Request

2. 代码规范
   - 遵循项目现有的代码风格
   - 添加必要的注释和文档
   - 确保所有测试通过
   - 更新相关文档

3. 问题反馈
   - 使用 GitHub Issues
   - 提供详细的问题描述
   - 附上错误日志和复现步骤
   - 标记相关标签
