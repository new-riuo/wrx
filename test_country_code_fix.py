import json
import os

# 模拟ExcelProcessor中的国家代码匹配逻辑
def test_country_code_matching():
    print("开始测试国家代码匹配功能...")
    
    # 加载修复后的国家代码数据
    country_codes = []
    if os.path.exists("country_codes.json"):
        with open("country_codes.json", "r", encoding="utf-8") as f:
            country_codes = json.load(f)
        print(f"成功加载 {len(country_codes)} 个国家代码")
    else:
        print("错误：找不到country_codes.json文件")
        return
    
    # 测试用例：包含不同格式的国家名称
    test_cases = [
        {"test_name": "美国完整名称", "consignee_country": "United States", "expected_code": "USA"},
        {"test_name": "美国简称", "consignee_country": "USA", "expected_code": "USA"},
        {"test_name": "英国完整名称", "consignee_country": "United Kingdom", "expected_code": "GBR"},
        {"test_name": "英国简称", "consignee_country": "GB", "expected_code": "GBR"},
        {"test_name": "英国大写名称", "consignee_country": "UNITED KINGDOM", "expected_code": "GBR"},
        {"test_name": "德国", "consignee_country": "Germany", "expected_code": "DEU"},
        {"test_name": "德国大写", "consignee_country": "GERMANY", "expected_code": "DEU"},
        {"test_name": "韩国完整名称", "consignee_country": "Korea, Republic of", "expected_code": "KOR"},
        {"test_name": "韩国简称", "consignee_country": "Korea", "expected_code": "KOR"},
        {"test_name": "俄罗斯完整名称", "consignee_country": "Russian Federation", "expected_code": "RUS"},
        {"test_name": "俄罗斯简称", "consignee_country": "Russia", "expected_code": "RUS"}
    ]
    
    # 运行测试
    passed_tests = 0
    failed_tests = 0
    
    for test_case in test_cases:
        # 获取测试数据
        test_name = test_case["test_name"]
        consignee_country = test_case["consignee_country"]
        expected_code = test_case["expected_code"]
        
        # 模拟匹配过程
        country_info = None
        matched_code = "USA"  # 默认值
        
        # 1. 精确匹配
        for country in country_codes:
            cc_in_data = country.get("consignee_country", "").strip()
            if cc_in_data and cc_in_data == consignee_country:
                country_info = country
                matched_code = country.get("three_letter_code", "USA")
                break
        
        # 2. 模糊匹配
        if not country_info:
            for country in country_codes:
                cc_in_data = country.get("consignee_country", "").lower().strip()
                if cc_in_data and cc_in_data in consignee_country.lower():
                    country_info = country
                    matched_code = country.get("three_letter_code", "USA")
                    break
        
        # 3. 反向模糊匹配
        if not country_info:
            for country in country_codes:
                cc_in_data = country.get("consignee_country", "").lower().strip()
                if cc_in_data and consignee_country.lower() in cc_in_data:
                    country_info = country
                    matched_code = country.get("three_letter_code", "USA")
                    break
        
        # 验证结果
        if matched_code == expected_code:
            print(f"✓ 通过 - {test_name}: '{consignee_country}' -> '{matched_code}' (预期: '{expected_code}')")
            passed_tests += 1
        else:
            print(f"✗ 失败 - {test_name}: '{consignee_country}' -> '{matched_code}' (预期: '{expected_code}')")
            failed_tests += 1
    
    # 总结
    print("\n=== 测试结果总结 ===")
    print(f"总测试用例: {len(test_cases)}")
    print(f"通过: {passed_tests}")
    print(f"失败: {failed_tests}")
    
    if failed_tests == 0:
        print("🎉 所有测试通过！国家代码修复成功。")
    else:
        print("❌ 部分测试失败，请检查country_codes.json文件。")

# 运行测试
if __name__ == "__main__":
    test_country_code_matching()