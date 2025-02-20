import win32com.client
from typing import Dict, Optional, Tuple, Any
import os

class ExcelManager:
    """Excel管理器"""
    
    def __init__(self):
        self.app = None
        self.active_workbook = None
        self.active_sheet = None
        self.workbooks = {}
        
    def connect(self) -> bool:
        """连接到Excel应用程序"""
        try:
            if not self.app:
                # 先尝试获取已打开的Excel实例
                try:
                    self.app = win32com.client.GetObject(Class='Excel.Application')
                except:
                    # 如果没有打开的Excel，则创建新实例
                    self.app = win32com.client.Dispatch('Excel.Application')
                    self.app.Visible = True
                
            return True
            
        except Exception as e:
            print(f"连接Excel失败: {str(e)}")
            return False
            
    def is_connected(self) -> bool:
        """检查是否已连接到Excel"""
        try:
            if self.app:
                # 尝试访问属性以验证连接
                _ = self.app.Workbooks.Count
                return True
            return False
        except:
            self.app = None
            self.workbooks.clear()
            self.active_workbook = None
            self.active_sheet = None
            return False
            
    def open_file(self, file_path: str) -> bool:
        """打开Excel文件"""
        try:
            if not os.path.exists(file_path):
                print(f"文件不存在: {file_path}")
                return False
            
            if not self.connect():
                return False
            
            # 检查文件扩展名
            _, ext = os.path.splitext(file_path)
            if ext.lower() not in ('.xls', '.xlsx', '.xlsm'):
                print(f"不支持的文件类型: {ext}")
                return False
            
            # 检查文件是否已经打开
            for wb in self.app.Workbooks:
                if wb.FullName.lower() == file_path.lower():
                    wb.Activate()
                    self.active_workbook = wb
                    self.active_sheet = wb.ActiveSheet
                    # 添加到工作簿字典
                    self.workbooks[wb.Name] = wb
                    return True
            
            # 打开新文件
            workbook = self.app.Workbooks.Open(file_path)
            if not workbook:
                return False
            
            # 保存工作簿引用
            self.workbooks[workbook.Name] = workbook
            self.active_workbook = workbook
            self.active_sheet = workbook.ActiveSheet
            
            # 确保Excel窗口可见
            self.app.Visible = True
            workbook.Activate()
            
            return True
            
        except Exception as e:
            print(f"打开Excel文件失败: {str(e)}")
            return False
            
    def get_workbook(self, name: str) -> Optional[Any]:
        """获取工作簿
        Args:
            name: 工作簿名称
        Returns:
            Workbook对象或None
        """
        return self.workbooks.get(name)
            
    def refresh_workbooks(self) -> Dict[str, Any]:
        """刷新工作簿列表"""
        try:
            if not self.connect():
                return {}
            
            # 清空现有工作簿
            self.workbooks.clear()
            
            # 获取所有打开的工作簿
            for wb in self.app.Workbooks:
                self.workbooks[wb.Name] = wb
            
            return self.workbooks
            
        except Exception as e:
            print(f"刷新工作簿列表失败: {str(e)}")
            return {}
            
    def select_range(self, mode: str, value: Any = None) -> Optional[Any]:
        """选择区域
        Args:
            mode: 选择模式 ('input' 或 'output')
            value: 可选的预设值
        Returns:
            Range对象或None
        """
        try:
            if not self.active_sheet:
                return None
            
            if value:
                return self.active_sheet.Range(value)
            
            # 获取用户选择的区域
            selection = self.app.Selection
            if not selection:
                return None
            
            return selection
            
        except Exception as e:
            print(f"选择区域失败: {str(e)}")
            return None
            
    def apply_formula(self, formula: str, input_range: Any, output_range: Any = None) -> bool:
        """应用公式
        Args:
            formula: 公式模板
            input_range: 输入区域
            output_range: 输出区域（可选）
        Returns:
            bool: 是否成功
        """
        try:
            if not self.active_sheet:
                return False
            
            # 如果没有指定输出区域，使用输入区域
            if not output_range:
                output_range = input_range
            
            # 替换公式中的占位符
            formula = formula.replace('{range}', input_range.Address)
            
            # 应用公式
            output_range.Formula = formula
            return True
            
        except Exception as e:
            print(f"应用公式失败: {str(e)}")
            return False

    def get_current_state(self) -> Optional[Dict]:
        """获取当前状态"""
        try:
            if not self.active_sheet:
                return None
            
            used_range = self.active_sheet.UsedRange
            return {
                'values': used_range.Value,
                'formulas': used_range.Formula,
                'range': used_range.Address
            }
        except Exception as e:
            print(f"获取状态失败: {str(e)}")
            return None

    def restore_state(self, state: Dict) -> bool:
        """恢复状态
        Args:
            state: 状态字典
        Returns:
            bool: 是否成功
        """
        try:
            if not self.active_sheet or not state:
                return False
            
            range_obj = self.active_sheet.Range(state['range'])
            range_obj.Value = state['values']
            range_obj.Formula = state['formulas']
            return True
        
        except Exception as e:
            print(f"恢复状态失败: {str(e)}")
            return False

    def disconnect(self):
        """断开Excel连接"""
        try:
            if self.app:
                # 关闭所有打开的工作簿
                for wb in self.workbooks.values():
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        pass
                
                # 清空工作簿字典
                self.workbooks.clear()
                self.active_workbook = None
                self.active_sheet = None
                
                # 退出Excel应用
                self.app.Quit()
                self.app = None
                return True
                
        except Exception as e:
            print(f"断开Excel连接失败: {str(e)}")
            self.app = None
            return False

    def __del__(self):
        """析构函数，确保清理资源"""
        try:
            self.disconnect()
        except:
            pass 