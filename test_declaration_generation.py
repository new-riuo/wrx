#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试报关单数据生成功能
"""

from excel_process import ExcelProcessor

def test_generate_declaration_data():
    """测试生成报关单数据"""
    print("开始测试报关单数据生成功能...")
    
    # 创建Excel处理器实例
    processor = ExcelProcessor()
    
    # 打印初始数据
    print(f"\n1. 初始订单数据数量: {len(processor.order_data)}")
    print(f"2. 初始报关信息数量: {len(processor.declaration_info)}")
    print(f"3. 初始国家代码数量: {len(processor.country_codes)}")
    
    # 如果没有订单数据，添加一些测试订单
    if not processor.order_data:
        print("\n4. 添加测试订单数据...")
        processor.order_data = [
            {
                "订单编号": "TEST-001",
                "Order Code": "TEST-001",
                "目的国": "美国",
                "物流跟踪号": "TRK-001",
                "总销售金额": 100.0,
                "AmountPaid": 100.0,
                "店铺名称": "Top Unique Hair",
                "Consignee Country": "United States"
            },
            {
                "订单编号": "TEST-002",
                "Order Code": "TEST-002",
                "目的国": "英国",
                "物流跟踪号": "TRK-002",
                "总销售金额": 200.0,
                "AmountPaid": 200.0,
                "店铺名称": "BS",
                "Consignee Country": "United Kingdom"
            },
            {
                "订单编号": "TEST-003",
                "Order Code": "TEST-003",
                "目的国": "德国",
                "物流跟踪号": "TRK-003",
                "总销售金额": 300.0,
                "AmountPaid": 300.0,
                "店铺名称": "假发3店",
                "Consignee Country": "Germany"
            }
        ]
        print(f"   添加了 {len(processor.order_data)} 条测试订单数据")
    
    # 生成报关单数据
    print("\n5. 开始生成报关单数据...")
    declaration_data = processor.generate_declaration_data()
    
    # 打印生成结果
    print(f"\n6. 生成结果:")
    print(f"   成功生成 {len(declaration_data)} 条报关单数据")
    
    # 打印详细数据
    if declaration_data:
        print("\n7. 详细数据:")
        for i, item in enumerate(declaration_data, 1):
            print(f"   第 {i} 条:")
            for key, value in item.items():
                print(f"      {key}: {value}")
            print()
    
    print("\n测试完成!")

if __name__ == "__main__":
    test_generate_declaration_data()
