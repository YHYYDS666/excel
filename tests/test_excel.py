import unittest
from src.utils.excel_manager import ExcelManager

class TestExcelManager(unittest.TestCase):
    def setUp(self):
        """测试前准备"""
        self.excel_manager = ExcelManager()
        
    def test_connect(self):
        """测试Excel连接"""
        # 测试连接成功
        result = self.excel_manager.connect()
        self.assertTrue(result)
        
        # 测试是否已连接
        self.assertTrue(self.excel_manager.is_connected())
        
    def test_workbooks(self):
        """测试工作簿操作"""
        # 测试刷新工作簿列表
        workbooks = self.excel_manager.refresh_workbooks()
        self.assertIsInstance(workbooks, dict)
        
    def test_select_range(self):
        """测试区域选择"""
        # 测试选择输入区域
        range_obj = self.excel_manager.select_range('input')
        self.assertIsNotNone(range_obj)
        
        # 测试选择输出区域
        range_obj = self.excel_manager.select_range('output')
        self.assertIsNotNone(range_obj)

    def test_state_operations(self):
        """测试状态操作"""
        # 测试获取状态
        state = self.excel_manager.get_current_state()
        self.assertIsInstance(state, dict)
        
        # 测试恢复状态
        result = self.excel_manager.restore_state(state)
        self.assertTrue(result)

    def test_file_operations(self):
        """测试文件操作"""
        # 测试打开文件
        result = self.excel_manager.open_file("test.xlsx")
        self.assertTrue(result)
        
        # 测试获取工作簿
        wb = self.excel_manager.get_workbook("test.xlsx")
        self.assertIsNotNone(wb)
        
        # 测试激活工作簿
        result = self.excel_manager.activate_workbook("test.xlsx")
        self.assertTrue(result)

    def test_formula_operations(self):
        """测试公式操作"""
        # 测试应用公式
        formula = "=SUM({range})"
        input_range = self.excel_manager.select_range('input')
        result = self.excel_manager.apply_formula(formula, input_range)
        self.assertTrue(result)
        
        # 测试获取公式结果
        result = input_range.Value
        self.assertIsNotNone(result)

    def tearDown(self):
        """测试后清理"""
        try:
            if self.excel_manager:
                self.excel_manager.disconnect()
        except:
            pass

    @classmethod
    def setUpClass(cls):
        """测试前准备测试数据"""
        try:
            # 创建测试用Excel文件
            import xlsxwriter
            workbook = xlsxwriter.Workbook('test.xlsx')
            worksheet = workbook.add_worksheet()
            
            # 写入测试数据
            test_data = [
                [1, 2, 3],
                [4, 5, 6],
                [7, 8, 9]
            ]
            for row, data in enumerate(test_data):
                worksheet.write_row(row, 0, data)
            
            workbook.close()
            
        except Exception as e:
            print(f"准备测试数据失败: {str(e)}")

    @classmethod
    def tearDownClass(cls):
        """清理测试数据"""
        try:
            # 清理测试文件
            import os
            if os.path.exists('test.xlsx'):
                os.remove('test.xlsx')
            
            # 关闭Excel进程
            import win32com.client
            excel = win32com.client.GetObject(Class='Excel.Application')
            excel.Quit()
            
        except:
            pass

if __name__ == '__main__':
    unittest.main() 