"""
测试物流跟踪号与订单匹配逻辑
"""

from excel_process import ExcelProcessor

def test_tracking_number_matching():
    """测试报关单包裹号（运单号）与物流信息的匹配"""
    print("===== 开始测试物流跟踪号匹配功能 =====")
    
    # 创建处理器实例
    processor = ExcelProcessor()
    
    # 初始化必要的数据结构
    processor.order_data = [
        {
            "订单编号": "TEST-ORDER-001",
            "Order Code": "TEST-ORDER-001",
            "目的国": "美国",
            "物流跟踪号": "TRACK-ORDER001",  # 订单中的跟踪号（带有前缀）
            "总销售金额": 100.0,
            "AmountPaid": 100.0,
            "店铺名称": "Test Shop",
            "Consignee Country": "United States"
        },
        {
            "订单编号": "TEST-ORDER-002",
            "Order Code": "TEST-ORDER-002",
            "目的国": "英国",
            "Tracking Number": "ORDER002-TRACK",  # 订单中的跟踪号
            "总销售金额": 200.0,
            "AmountPaid": 200.0,
            "店铺名称": "Test Shop",
            "Consignee Country": "United Kingdom"
        },
        {
            "订单编号": "TEST-ORDER-003",
            "Order Code": "TEST-ORDER-003",
            "目的国": "加拿大",
            "包裹号（运单号）": "",  # 订单中没有跟踪号
            "总销售金额": 150.0,
            "AmountPaid": 150.0,
            "店铺名称": "Test Shop",
            "Consignee Country": "Canada"
        }
    ]
    
    # 设置匹配的物流信息（包含更准确的跟踪号）
    processor.logistics_data = [
        {
            "平台订单号": "TEST-ORDER-001",
            "物流跟踪号": "L0G1ST1CS-TRACKING-001",  # 物流信息中的跟踪号（更准确）
            "is_matched": True
        },
        {
            "订单号": "TEST-ORDER-002",
            "运单号": "CORRECT-TRACKING-002",  # 使用不同字段名存储跟踪号
            "is_matched": True
        },
        {
            "包裹号": "LOGISTICS-TRACK-003",  # 使用包裹号字段
            "平台订单号": "TEST-ORDER-003",
            "is_matched": True
        }
    ]
    
    # 初始化其他必要数据
    processor.declaration_info = [
        {
            "提单号": "TEST-BL-001",
            "商品品名": "测试商品",
            "HS CODE": "123456",
            "规格型号": "测试规格",
            "包裹内单个SKC的商品数量": 1,
            "申报单位": "个",
            "申报币制": "502",
            "商品总净重(KG)": 0.5,
            "第一法定数量": 0.5,
            "第一法定单位": "035",
            "电商企业代码": "TEST-CODE",
            "电商企业名称": "Test Shop",
            "电商平台代码": "",
            "电商平台名称": "",
            "收款企业代码": "",
            "收款企业名称": "",
            "生产企业代码": "",
            "生产企业名称": "",
            "电商企业dxpId": ""
        }
    ]
    
    # 设置店铺对应公司数据
    processor.shop_company_data = [
        {
            "shop_name": "Test Shop",
            "company_name": "Test Shop"
        }
    ]
    
    # 设置国家代码
    processor.country_codes = [
        {
            "country_name": "美国",
            "consignee_country": "United States",
            "three_letter_code": "USA"
        },
        {
            "country_name": "英国",
            "consignee_country": "United Kingdom",
            "three_letter_code": "GBR"
        },
        {
            "country_name": "加拿大",
            "consignee_country": "Canada",
            "three_letter_code": "CAN"
        }
    ]
    
    # 设置申报金额规则
    processor.declaration_amount_rules = [
        {
            "country_name": "美国",
            "declaration_ratio": 0.5,
            "max_declaration_amount": 800
        },
        {
            "country_name": "英国",
            "declaration_ratio": 0.6,
            "max_declaration_amount": 700
        },
        {
            "country_name": "加拿大",
            "declaration_ratio": 0.55,
            "max_declaration_amount": 800
        }
    ]
    
    # 生成报关单数据
    declaration_data = processor.generate_declaration_data()
    
    # 打印结果
    print("\n===== 测试结果 =====")
    print(f"生成的报关单数据数量: {len(declaration_data)}")
    
    # 检查每个订单的跟踪号是否正确匹配
    for index, item in enumerate(declaration_data):
        order_no = item.get("订单编号", "")
        tracking_no = item.get("包裹号（运单号）", "")
        print(f"\n订单 {index+1}: 订单号={order_no}")
        print(f"  报关单包裹号（运单号）: {tracking_no}")
        
        # 验证跟踪号是否符合预期
        expected_tracking = ""
        if order_no == "TEST-ORDER-001":
            expected_tracking = "L0G1ST1CS-TRACKING-001"  # 应从物流信息获取
        elif order_no == "TEST-ORDER-002":
            expected_tracking = "CORRECT-TRACKING-002"  # 应从物流信息获取
        elif order_no == "TEST-ORDER-003":
            expected_tracking = "LOGISTICS-TRACK-003"  # 应从物流信息获取
        
        if tracking_no == expected_tracking:
            print(f"  ✅ 跟踪号匹配正确: 期望={expected_tracking}, 实际={tracking_no}")
        else:
            print(f"  ❌ 跟踪号匹配错误: 期望={expected_tracking}, 实际={tracking_no}")
    
    print("\n===== 测试完成 =====")

if __name__ == "__main__":
    test_tracking_number_matching()
