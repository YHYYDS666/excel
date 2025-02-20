import json
import os
from typing import Dict, List, Optional, Any
import copy

class FormulaManager:
    def __init__(self):
        """初始化公式管理器"""
        self.formulas = {
            "基础运算": [
                {
                    "name": "求和",
                    "description": "计算选定区域的和",
                    "template": "=SUM({range})"
                },
                {
                    "name": "平均值",
                    "description": "计算选定区域的平均值",
                    "template": "=AVERAGE({range})"
                },
                {
                    "name": "最大值",
                    "description": "返回选定区域中的最大值",
                    "template": "=MAX({range})"
                },
                {
                    "name": "最小值",
                    "description": "返回选定区域中的最小值",
                    "template": "=MIN({range})"
                },
                {
                    "name": "乘积",
                    "description": "返回选定区域中所有数值的乘积",
                    "template": "=PRODUCT({range})"
                }
            ],
            "统计函数": [
                {
                    "name": "计数",
                    "description": "统计选定区域内数值的个数",
                    "template": "=COUNT({range})"
                },
                {
                    "name": "文本计数",
                    "description": "统计选定区域内非空单元格的个数",
                    "template": "=COUNTA({range})"
                },
                {
                    "name": "空值计数",
                    "description": "统计选定区域内空单元格的个数",
                    "template": "=COUNTBLANK({range})"
                },
                {
                    "name": "标准差",
                    "description": "返回选定区域中数值的标准差",
                    "template": "=STDEV({range})"
                },
                {
                    "name": "方差",
                    "description": "返回选定区域中数值的方差",
                    "template": "=VAR({range})"
                }
            ],
            "条件函数": [
                {
                    "name": "条件求和",
                    "description": "计算选定区域内符合条件的数值之和",
                    "template": "=SUMIF({range}, {criteria})"
                },
                {
                    "name": "条件计数",
                    "description": "统计选定区域内符合条件的数值个数",
                    "template": "=COUNTIF({range}, {criteria})"
                },
                {
                    "name": "条件平均",
                    "description": "计算选定区域内符合条件的数值的平均值",
                    "template": "=AVERAGEIF({range}, {criteria})"
                }
            ],
            "查找函数": [
                {
                    "name": "查找最大值位置",
                    "description": "查找最大值在区域中的位置",
                    "template": "=MATCH(MAX({range}),{range},0)"
                },
                {
                    "name": "查找最小值位置",
                    "description": "查找最小值在区域中的位置",
                    "template": "=MATCH(MIN({range}),{range},0)"
                }
            ]
        }
        self.load_formulas()
    
    def load_formulas(self) -> Dict:
        """加载公式配置"""
        try:
            # 获取公式配置文件路径
            config_path = os.path.join('data', 'formulas.json')
            
            # 读取配置文件
            with open(config_path, 'r', encoding='utf-8') as f:
                self.formulas = json.load(f)  # 直接赋值给self.formulas
                return self.formulas
                
        except Exception as e:
            print(f"加载公式配置失败: {str(e)}")
            self.formulas = {}  # 确保失败时也是字典
            return self.formulas
    
    def save_formulas(self) -> bool:
        """保存公式到文件"""
        try:
            # 获取当前文件的绝对路径
            current_file = os.path.abspath(__file__)
            # 获取utils目录的路径
            utils_dir = os.path.dirname(current_file)
            # 获取src目录的路径
            src_dir = os.path.dirname(utils_dir)
            # 获取项目根目录的路径
            root_dir = os.path.dirname(src_dir)
            # 构建formulas.json的完整路径
            json_path = os.path.join(root_dir, 'data', 'formulas.json')
            
            # 确保data目录存在
            os.makedirs(os.path.dirname(json_path), exist_ok=True)
            
            # 保存到JSON文件
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(self.formulas, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存公式失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_categories(self) -> List[str]:
        """获取所有公式分类"""
        return list(self.formulas.keys())
    
    def get_formulas(self, category: str) -> List[str]:
        """获取指定分类下的所有公式名称"""
        try:
            formulas = self.formulas.get(category, [])
            return [formula['name'] for formula in formulas]
        except Exception as e:
            print(f"获取公式列表失败: {str(e)}")
            return []
    
    def get_formula_description(self, formula_name: str) -> str:
        """获取公式描述"""
        try:
            # 遍历所有分类
            for category in self.formulas.values():
                # 遍历分类中的公式
                for formula in category:
                    if formula['name'] == formula_name:
                        return formula.get('description', '')
            return ''
        except Exception as e:
            print(f"获取公式描述失败: {str(e)}")
            return ''
    
    def get_formula(self, name: str) -> Optional[Dict]:
        """获取指定公式
        Args:
            name: 公式名称
        Returns:
            公式信息字典或None
        """
        try:
            # 如果传入的是字典，说明已经是公式对象了
            if isinstance(name, dict):
                return name
            
            # 遍历所有分类查找公式
            for category, formulas in self.formulas.items():
                for formula in formulas:
                    if formula['name'] == name:
                        return formula
            return None
        
        except Exception as e:
            print(f"获取公式失败: {str(e)}")
            return None
    
    def add_formula(self, category: str, formula: Dict) -> bool:
        """添加公式
        Args:
            category: 公式分类
            formula: 公式信息
        Returns:
            bool: 是否成功
        """
        try:
            # 如果分类不存在，创建新分类
            if category not in self.formulas:
                self.formulas[category] = []
                
            # 添加公式
            self.formulas[category].append(formula)
            
            # 保存配置
            return self.save_formulas()
        except Exception as e:
            print(f"添加公式失败: {str(e)}")
            return False
    
    def update_formula(self, old_name: str, new_formula: Dict) -> bool:
        """更新公式
        Args:
            old_name: 原公式名称
            new_formula: 新公式信息
        Returns:
            bool: 是否成功
        """
        try:
            # 遍历所有分类
            for category in self.formulas.values():
                # 遍历分类中的公式
                for i, formula in enumerate(category):
                    if formula['name'] == old_name:
                        # 更新公式
                        category[i] = new_formula
                        # 保存配置
                        return self.save_formulas()
            return False
        except Exception as e:
            print(f"更新公式失败: {str(e)}")
            return False
    
    def delete_formula(self, formula_name: str) -> bool:
        """删除公式
        Args:
            formula_name: 公式名称
        Returns:
            bool: 是否成功
        """
        try:
            # 遍历所有分类
            for category in self.formulas.values():
                # 遍历分类中的公式
                for i, formula in enumerate(category):
                    if formula['name'] == formula_name:
                        # 删除公式
                        category.pop(i)
                        # 保存配置
                        return self.save_formulas()
            return False
        except Exception as e:
            print(f"删除公式失败: {str(e)}")
            return False
    
    def get_formula_template(self, formula_name: str) -> Optional[str]:
        """获取公式模板
        Args:
            formula_name: 公式名称
        Returns:
            str: 公式模板
        """
        try:
            # 遍历所有分类
            for category in self.formulas.values():
                # 遍历分类中的公式
                for formula in category:
                    if formula['name'] == formula_name:
                        return formula.get('template', '')
            return None
        except Exception as e:
            print(f"获取公式模板失败: {str(e)}")
            return None
    
    def get_all_formulas(self) -> List[Dict]:
        """获取所有公式
        
        Returns:
            List[Dict]: 所有公式列表，每个公式包含:
                - category: 分类名称
                - name: 公式名称
                - template: 公式模板
                - description: 公式描述
        """
        try:
            results = []
            for category, formulas in self.formulas.items():
                for formula in formulas:
                    results.append({
                        "category": category,
                        "name": formula['name'],
                        "template": formula['template'],
                        "description": formula['description']
                    })
            # 按分类和名称排序
            results.sort(key=lambda x: (x["category"], x["name"]))
            return results
        except Exception as e:
            print(f"获取所有公式失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return []
    
    def search_formulas(self, keyword: str) -> List[Dict]:
        """搜索公式
        
        Args:
            keyword (str): 搜索关键词
            
        Returns:
            List[Dict]: 匹配的公式列表，每个公式包含:
                - category: 分类名称
                - name: 公式名称
                - template: 公式模板
                - description: 公式描述
        """
        keyword = keyword.lower()
        results = []
        
        try:
            for category, formulas in self.formulas.items():
                for formula in formulas:
                    # 在名称和描述中搜索
                    if (keyword in formula['name'].lower() or 
                        keyword in formula['description'].lower()):
                        results.append({
                            "category": category,
                            "name": formula['name'],
                            "template": formula['template'],
                            "description": formula['description']
                        })
            
            # 按分类和名称排序
            results.sort(key=lambda x: (x["category"], x["name"]))
            
            return results
        except Exception as e:
            print(f"搜索公式时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            return []
    
    def prepare_formula(self, category, formula_name, input_range, output_range=None):
        """准备公式
        Args:
            category: 公式分类
            formula_name: 公式名称
            input_range: 输入区域
            output_range: 输出区域（可选）
        Returns:
            str: 准备好的公式
        """
        try:
            # 获取公式模板
            if category not in self.formulas or formula_name not in self.formulas[category]:
                print(f"公式不存在: {category} - {formula_name}")
                return None
                
            formula_template = self.formulas[category][formula_name]
            
            # 如果没有指定输出区域，直接使用输入区域的公式
            if not output_range:
                return formula_template.replace("{range}", input_range)
                
            # 根据公式类型和区域生成最终公式
            if "AVERAGE" in formula_template:
                return f"=AVERAGE({input_range})"
            elif "SUM" in formula_template:
                return f"=SUM({input_range})"
            elif "MAX" in formula_template:
                return f"=MAX({input_range})"
            elif "MIN" in formula_template:
                return f"=MIN({input_range})"
            elif "COUNT" in formula_template:
                return f"=COUNT({input_range})"
            elif "LEFT" in formula_template or "RIGHT" in formula_template:
                # 文本函数需要处理每个单元格
                return f"=IF(ROW()=ROW({output_range}),{formula_template.replace('{range}', input_range)},\"\")"
            elif "CONCATENATE" in formula_template:
                # 连接函数需要特殊处理
                return f"=CONCATENATE({input_range})"
            elif "VLOOKUP" in formula_template:
                # 查找函数需要特殊处理
                lookup_range = input_range.split(",")[0] if "," in input_range else input_range
                return f"=VLOOKUP({lookup_range},{input_range},2,FALSE)"
            else:
                # 默认情况，直接替换范围
                return formula_template.replace("{range}", input_range)
                
        except Exception as e:
            print(f"准备公式失败: {str(e)}")
            return None
    
    def apply_formula(self, name, input_range, output_range=None):
        """应用公式
        Args:
            name: 公式名称
            input_range: 输入区域
            output_range: 输出区域（可选）
        Returns:
            bool: 是否应用成功
        """
        try:
            # 获取公式信息
            formula = self.get_formula(name)
            if not formula:
                print(f"未找到公式: {name}")
                return False

            # 获取区域地址
            input_address = input_range.Address
            if output_range:
                output_address = output_range.Address

            # 准备公式文本
            formula_text = formula['template']
            
            # 替换输入区域
            formula_text = formula_text.replace('{range}', input_address)
            
            # 如果有输出区域且公式模板包含输出占位符
            if output_range and '{output}' in formula_text:
                formula_text = formula_text.replace('{output}', output_address)
            
            # 如果是条件函数，添加条件参数
            if '{criteria}' in formula_text:
                # TODO: 后续添加条件输入对话框
                criteria = '">0"'  # 临时使用固定条件
                formula_text = formula_text.replace('{criteria}', criteria)

            # 应用公式到Excel
            target_range = output_range if output_range else input_range
            target_range.Formula = formula_text
            return True

        except Exception as e:
            print(f"应用公式失败: {str(e)}")
            return False 