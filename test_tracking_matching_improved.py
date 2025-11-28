"""
增强版物流跟踪号匹配测试工具
"""

class ExcelProcessorMock:
    """模拟ExcelProcessor类的关键功能"""
    def __init__(self):
        self.order_data = []
        self.logistics_data = []
        self.country_codes = []
        self.declaration_amount_rules = []
        self.declaration_info = []
        self.shop_company_data = []
    
    def get_logistics_tracking_no(self, order_row, order_no):
        """增强版物流跟踪号获取逻辑"""
        logistics_tracking_no = order_no  # 默认使用订单号
        try:
            # 首先从订单数据中获取物流跟踪号
            order_tracking_no = order_row.get("物流跟踪号", "") or \
                               order_row.get("Tracking Number", "") or \
                               order_row.get("包裹号（运单号）", "")
            
            # 如果订单中有物流跟踪号，先使用它
            if order_tracking_no:
                logistics_tracking_no = order_tracking_no
                print(f"从订单数据获取物流跟踪号: {logistics_tracking_no}")
            
            # 尝试从已匹配的物流信息中获取更准确的物流跟踪号
            if hasattr(self, 'logistics_data') and self.logistics_data:
                print(f"检查{len(self.logistics_data)}条物流信息行以匹配订单号: {order_no}")
                for logistics_row in self.logistics_data:
                    # 优化的匹配逻辑：支持多种订单号字段匹配
                    platform_order_no = logistics_row.get("平台订单号", "")
                    logistics_order_no = logistics_row.get("订单号", "")
                    order_code = logistics_row.get("Order Code", "")
                    
                    # 输出每条物流行的订单号信息以便调试
                    print(f"  检查物流行: 平台订单号='{platform_order_no}', 订单号='{logistics_order_no}', Order Code='{order_code}'")
                    
                    # 更宽松的匹配条件，只要任何一个订单号字段匹配就认为是同一订单
                    if platform_order_no == order_no or logistics_order_no == order_no or order_code == order_no:
                        print(f"找到匹配的物流信息行，订单号: {order_no}")
                        # 从物流信息中获取跟踪号（支持更多字段名）
                        logistics_info_tracking_no = logistics_row.get("物流跟踪号", "") or \
                                                  logistics_row.get("Tracking Number", "") or \
                                                  logistics_row.get("包裹号（运单号）", "") or \
                                                  logistics_row.get("运单号", "") or \
                                                  logistics_row.get("包裹号", "") or \
                                                  logistics_row.get("跟踪号", "")
                         
                        print(f"  从物流行获取的跟踪号: '{logistics_info_tracking_no}'")
                        
                        if logistics_info_tracking_no:
                            # 标准化处理：去除可能的前缀
                            if logistics_info_tracking_no.startswith("TRACK-"):
                                logistics_info_tracking_no = logistics_info_tracking_no[6:]
                                print(f"  移除TRACK-前缀后: {logistics_info_tracking_no}")
                            
                            # 优先使用物流信息中的跟踪号
                            print(f"  使用物流信息中的跟踪号替换: {logistics_info_tracking_no}")
                            logistics_tracking_no = logistics_info_tracking_no
                            break
            
            # 预处理跟踪号：去掉可能的前缀
            if logistics_tracking_no.startswith("TRACK-"):
                logistics_tracking_no = logistics_tracking_no[6:]
                print(f"移除TRACK-前缀后跟踪号: {logistics_tracking_no}")
            
            print(f"最终确定的物流跟踪号: {logistics_tracking_no}")
        except Exception as e:
            print(f"获取/匹配物流跟踪号时出错: {str(e)}")
            # 出错时使用订单号作为跟踪号
            logistics_tracking_no = order_no
            print(f"使用订单号作为跟踪号: {logistics_tracking_no}")
        
        return logistics_tracking_no
    
    def generate_test_declaration_data(self):
        """生成测试报关单数据"""
        declaration_data = []
        
        for order_row in self.order_data:
            order_no = order_row.get("订单编号", "") or order_row.get("Order Code", "")
            print(f"\n处理订单: {order_no}")
            
            # 获取物流跟踪号
            tracking_no = self.get_logistics_tracking_no(order_row, order_no)
            
            # 创建报关单数据项
            declaration_item = {
                "订单编号": order_no,
                "包裹号（运单号）": tracking_no,
                "目的国": order_row.get("目的国", "未知")
            }
            
            declaration_data.append(declaration_item)
        
        return declaration_data

def test_tracking_matching():
    """测试物流跟踪号匹配功能"""
    print("===== 增强版物流跟踪号匹配测试 =====")
    
    # 创建模拟处理器
    processor = ExcelProcessorMock()
    
    # 设置测试数据
    processor.order_data = [
        {
            "订单编号": "TEST-ORDER-001",
            "Order Code": "TEST-ORDER-001",
            "目的国": "美国",
            "物流跟踪号": "TRACK-ORDER001",  # 订单中的跟踪号（带有前缀）
            "总销售金额": 100.0,
            "Consignee Country": "United States"
        },
        {
            "订单编号": "TEST-ORDER-002",
            "Order Code": "TEST-ORDER-002",
            "目的国": "英国",
            "Tracking Number": "ORDER002-TRACK",  # 订单中的跟踪号
            "总销售金额": 200.0,
            "Consignee Country": "United Kingdom"
        },
        {
            "订单编号": "TEST-ORDER-003",
            "Order Code": "TEST-ORDER-003",
            "目的国": "加拿大",
            "包裹号（运单号）": "",  # 订单中没有跟踪号
            "总销售金额": 150.0,
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
            "订单号": "TEST-ORDER-002",  # 使用订单号字段而不是平台订单号
            "运单号": "CORRECT-TRACKING-002",  # 使用运单号字段
            "is_matched": True
        },
        {
            "包裹号": "LOGISTICS-TRACK-003",  # 使用包裹号字段
            "平台订单号": "TEST-ORDER-003",
            "is_matched": True
        }
    ]
    
    # 生成报关单数据
    declaration_data = processor.generate_test_declaration_data()
    
    # 打印结果
    print("\n===== 测试结果汇总 =====")
    print(f"生成的报关单数据数量: {len(declaration_data)}")
    
    for item in declaration_data:
        print(f"\n订单号: {item.get('订单编号')}")
        print(f"包裹号（运单号）: {item.get('包裹号（运单号）')}")
        
        # 验证结果
        order_no = item.get('订单编号')
        tracking_no = item.get('包裹号（运单号）')
        
        expected_tracking = ""
        if order_no == "TEST-ORDER-001":
            expected_tracking = "L0G1ST1CS-TRACKING-001"
        elif order_no == "TEST-ORDER-002":
            expected_tracking = "CORRECT-TRACKING-002"
        elif order_no == "TEST-ORDER-003":
            expected_tracking = "LOGISTICS-TRACK-003"
        
        if tracking_no == expected_tracking:
            print(f"✅ 跟踪号匹配正确: {tracking_no}")
        else:
            print(f"❌ 跟踪号匹配错误: 期望={expected_tracking}, 实际={tracking_no}")
    
    print("\n===== 测试完成 =====")

if __name__ == "__main__":
    test_tracking_matching()
