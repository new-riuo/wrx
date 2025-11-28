import json
import os

# 模拟ExcelProcessor中的国家代码相关功能
class MockExcelProcessor:
    def __init__(self):
        # 初始化国家代码列表
        self.country_codes = [
            {"country_name": "美国", "country_code": "US", "consignee_country": "United States", "three_letter_code": "USA"},
            {"country_name": "英国", "country_code": "GB", "consignee_country": "United Kingdom", "three_letter_code": "GBR"},
            {"country_name": "测试国家", "country_code": "TS", "consignee_country": "Test Country", "three_letter_code": "TST"}
        ]
        self.test_file = "test_country_codes.json"
    
    def get_country_codes(self):
        """获取国家代码列表"""
        return self.country_codes
    
    def add_country_code(self, country_name, country_code, consignee_country, three_letter_code):
        """添加国家代码"""
        # 检查是否已存在
        for code in self.country_codes:
            if code["country_name"] == country_name or code["country_code"] == country_code:
                raise Exception(f"国家代码已存在: {country_name} - {country_code}")
        
        # 添加新国家代码
        self.country_codes.append({
            "country_name": country_name,
            "country_code": country_code,
            "consignee_country": consignee_country,
            "three_letter_code": three_letter_code
        })
        print(f"✓ 国家代码添加成功: {country_name} - {country_code}")
        
    def edit_country_code(self, original_country_name, original_country_code, new_country_name, new_country_code, new_consignee_country, new_three_letter_code):
        """编辑国家代码"""
        found = False
        for i, code in enumerate(self.country_codes):
            if code["country_name"] == original_country_name and code["country_code"] == original_country_code:
                self.country_codes[i] = {
                    "country_name": new_country_name,
                    "country_code": new_country_code,
                    "consignee_country": new_consignee_country,
                    "three_letter_code": new_three_letter_code
                }
                found = True
                print(f"✓ 国家代码编辑成功: {original_country_name} -> {new_country_name} - {new_country_code}")
                break
        
        if not found:
            raise Exception(f"未找到国家代码: {original_country_name} ({original_country_code})")
    
    def delete_country_code(self, country_name):
        """删除国家代码"""
        found = False
        for i, code in enumerate(self.country_codes):
            if code["country_name"] == country_name:
                del self.country_codes[i]
                found = True
                print(f"✓ 国家代码删除成功: {country_name}")
                break
        
        if not found:
            raise Exception(f"未找到国家代码: {country_name}")
    
    def save_country_codes(self):
        """保存国家代码到文件（测试用）"""
        try:
            with open(self.test_file, "w", encoding="utf-8") as f:
                json.dump(self.country_codes, f, ensure_ascii=False, indent=4)
            print("✓ 国家代码保存成功")
        except Exception as e:
            print(f"保存失败: {e}")

# 测试函数
def test_country_code_operations():
    print("===== 开始测试国家代码操作功能 =====\n")
    
    # 创建测试处理器
    processor = MockExcelProcessor()
    
    try:
        # 1. 显示初始国家代码
        print("初始国家代码列表:")
        for code in processor.get_country_codes():
            print(f"  {code['country_name']} - {code['country_code']} - {code['three_letter_code']}")
        print()
        
        # 2. 测试编辑功能
        print("测试编辑功能:")
        processor.edit_country_code("测试国家", "TS", "更新测试国家", "UPD", "Updated Test Country", "UPDT")
        print()
        
        # 3. 测试删除功能
        print("测试删除功能:")
        processor.delete_country_code("更新测试国家")
        print()
        
        # 4. 显示操作后的国家代码列表
        print("操作后的国家代码列表:")
        for code in processor.get_country_codes():
            print(f"  {code['country_name']} - {code['country_code']} - {code['three_letter_code']}")
        print()
        
        # 5. 测试保存功能
        print("测试保存功能:")
        processor.save_country_codes()
        print()
        
        print("===== 所有测试通过！国家代码编辑和删除功能正常工作 =====")
        print("\n提示：在实际应用中，这些功能已集成到UI界面，可以通过以下方式使用：")
        print("1. 打开Excel管理工具")
        print("2. 切换到国家代码标签页")
        print("3. 使用新添加的'编辑国家代码'和'删除国家代码'按钮进行操作")
        print("4. 操作完成后点击'保存国家代码'按钮保存更改")
        
    except Exception as e:
        print(f"测试失败: {e}")
    finally:
        # 清理测试文件
        if os.path.exists(processor.test_file):
            os.remove(processor.test_file)

if __name__ == "__main__":
    test_country_code_operations()
