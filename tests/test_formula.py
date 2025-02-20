import unittest
from src.utils.formula_manager import FormulaManager

class TestFormulaManager(unittest.TestCase):
    def setUp(self):
        """测试前准备"""
        self.formula_manager = FormulaManager()
        
    def test_categories(self):
        """测试公式分类"""
        categories = self.formula_manager.get_categories()
        self.assertIsInstance(categories, list)
        self.assertGreater(len(categories), 0)
        
    def test_formulas(self):
        """测试公式操作"""
        # 测试获取公式
        formulas = self.formula_manager.get_formulas("基础运算")
        self.assertIsInstance(formulas, list)
        
        # 测试获取公式描述
        desc = self.formula_manager.get_formula_description("求和")
        self.assertIsInstance(desc, str)
        
    def test_search(self):
        """测试公式搜索"""
        results = self.formula_manager.search("求和")
        self.assertIsInstance(results, list)
        self.assertGreater(len(results), 0)

if __name__ == '__main__':
    unittest.main() 