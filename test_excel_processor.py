from excel_process import ExcelProcessor

# 测试ExcelProcessor类
if __name__ == "__main__":
    # 创建ExcelProcessor实例
    processor = ExcelProcessor()
    
    print("测试ExcelProcessor类...")
    
    # 测试文件选择功能
    print("\n1. 测试文件选择功能（请手动取消选择）：")
    file_path = processor.select_file(title="测试文件选择")
    print(f"选择的文件：{file_path}")
    
    # 测试常量定义
    print("\n2. 测试常量定义：")
    print(f"当前工作簿名称：{processor.THIS_WORKBOOK_NAME}")
    print(f"模板配置数量：{len(processor.TEMPLATE_CONFIG)}")
    print(f"模板文件映射数量：{len(processor.TEMPLATE_FILE_MAP)}")
    
    # 测试鑫瑞祥和报关文件_Click方法
    print("\n3. 测试鑫瑞祥和报关文件_Click方法：")
    processor.鑫瑞祥和报关文件_Click()
    
    print("\n测试完成！")
    print("\n注意：")
    print("1. 报关订单文件_Click方法需要实际Excel文件进行测试")
    print("2. 特瑞福报关文件_Click方法需要先运行报关订单文件_Click方法")
    print("3. 请确保当前目录下存在'当前工作簿.xlsx'文件，且包含必要的工作表")
