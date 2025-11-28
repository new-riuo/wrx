import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from excel_process import ExcelProcessor

class ExcelManagerGUI:
    """报关单管理系统GUI窗口"""
    
    def __init__(self, root):
        """初始化GUI窗口"""
        self.root = root
        self.root.title("报关单管理系统")
        self.root.geometry("1000x700")
        
        # 创建ExcelProcessor实例
        self.processor = ExcelProcessor()
        
        # 创建状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建主要框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标签页
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 创建物流信息导入标签页
        self.logistics_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.logistics_tab, text="物流信息导入")
        
        # 创建订单信息导入标签页
        self.order_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.order_tab, text="订单信息导入")
        
        # 创建报关单模板设置标签页
        self.declaration_template_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.declaration_template_tab, text="报关单模板设置")
        
        # 创建报关信息维护标签页
        self.declaration_maintenance_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.declaration_maintenance_tab, text="报关信息维护")
        
        # 创建其他信息维护标签页
        self.other_info_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.other_info_tab, text="其他信息维护")
        
        # 初始化物流信息导入界面
        self.init_logistics_tab()
        
        # 初始化订单信息导入界面
        self.init_order_tab()
        
        # 初始化报关单模板设置界面
        self.init_declaration_template_tab()
        
        # 初始化报关信息维护界面
        self.init_declaration_maintenance_tab()
        
        # 初始化其他信息维护界面
        self.init_other_info_tab()
        
        # 创建报关单管理标签页
        self.customs_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.customs_tab, text="报关单管理")
        
        # 初始化报关单管理界面
        self.init_customs_tab()
    
    def init_logistics_tab(self):
        """初始化物流信息导入界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.logistics_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="导入物流信息表", command=self.import_logistics_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="处理物流信息", command=self.process_logistics_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新物流数据", command=self.refresh_logistics_data).pack(side=tk.LEFT, padx=5)
        
        # 添加筛选框架
        filter_frame = ttk.Frame(self.logistics_tab)
        filter_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Label(filter_frame, text="筛选字段：").pack(side=tk.LEFT, padx=5)
        self.logistics_filter_field = ttk.Combobox(filter_frame, width=20)
        self.logistics_filter_field.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(filter_frame, text="筛选值：").pack(side=tk.LEFT, padx=5)
        self.logistics_filter_value = ttk.Entry(filter_frame, width=20)
        self.logistics_filter_value.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(filter_frame, text="筛选", command=self.filter_logistics_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="重置", command=self.refresh_logistics_data).pack(side=tk.LEFT, padx=5)
        
        # 创建数据表格
        table_frame = ttk.Frame(self.logistics_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        self.logistics_tree = ttk.Treeview(table_frame, 
                                          columns=[], 
                                          show="headings",
                                          yscrollcommand=v_scrollbar.set,
                                          xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.logistics_tree.yview)
        h_scrollbar.config(command=self.logistics_tree.xview)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.logistics_tree.pack(fill=tk.BOTH, expand=True)
        
        # 刷新物流数据
        self.refresh_logistics_data()
    
    def init_order_tab(self):
        """初始化订单信息导入界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.order_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="导入订单信息表", command=self.import_order_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="匹配订单数据", command=self.match_order_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新订单数据", command=self.refresh_order_data).pack(side=tk.LEFT, padx=5)
        
        # 添加筛选框架
        filter_frame = ttk.Frame(self.order_tab)
        filter_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Label(filter_frame, text="筛选字段：").pack(side=tk.LEFT, padx=5)
        self.order_filter_field = ttk.Combobox(filter_frame, width=20)
        self.order_filter_field.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(filter_frame, text="筛选值：").pack(side=tk.LEFT, padx=5)
        self.order_filter_value = ttk.Entry(filter_frame, width=20)
        self.order_filter_value.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(filter_frame, text="筛选", command=self.filter_order_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="重置", command=self.refresh_order_data).pack(side=tk.LEFT, padx=5)
        
        # 创建数据表格
        table_frame = ttk.Frame(self.order_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        self.order_tree = ttk.Treeview(table_frame, 
                                      columns=[], 
                                      show="headings",
                                      yscrollcommand=v_scrollbar.set,
                                      xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.order_tree.yview)
        h_scrollbar.config(command=self.order_tree.xview)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.order_tree.pack(fill=tk.BOTH, expand=True)
        
        # 刷新订单数据
        self.refresh_order_data()
    
    def init_declaration_template_tab(self):
        """初始化报关单模板设置界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.declaration_template_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="刷新模板字段", command=self.refresh_template_fields).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存模板设置", command=self.save_template_settings).pack(side=tk.LEFT, padx=5)
        
        # 创建模板字段表格
        table_frame = ttk.Frame(self.declaration_template_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        self.template_fields_tree = ttk.Treeview(table_frame, 
                                               columns=["field_name", "field_type", "required"], 
                                               show="headings",
                                               yscrollcommand=v_scrollbar.set,
                                               xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.template_fields_tree.yview)
        h_scrollbar.config(command=self.template_fields_tree.xview)
        
        # 设置表格列标题和宽度
        self.template_fields_tree.heading("field_name", text="字段名称")
        self.template_fields_tree.heading("field_type", text="字段类型")
        self.template_fields_tree.heading("required", text="是否必填")
        
        self.template_fields_tree.column("field_name", width=200, anchor=tk.CENTER)
        self.template_fields_tree.column("field_type", width=150, anchor=tk.CENTER)
        self.template_fields_tree.column("required", width=100, anchor=tk.CENTER)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.template_fields_tree.pack(fill=tk.BOTH, expand=True)
        
        # 刷新模板字段
        self.refresh_template_fields()
    
    def init_declaration_maintenance_tab(self):
        """初始化报关信息维护界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.declaration_maintenance_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="刷新报关信息", command=self.refresh_declaration_info).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="添加报关信息", command=self.add_declaration_info).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="修改报关信息", command=self.edit_declaration_info).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除报关信息", command=self.delete_declaration_info).pack(side=tk.LEFT, padx=5)
        
        # 创建报关信息表格
        table_frame = ttk.Frame(self.declaration_maintenance_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        self.declaration_info_tree = ttk.Treeview(table_frame, 
                                                columns=["bill_of_lading", "skc", "product_name", "hs_code"], 
                                                show="headings",
                                                yscrollcommand=v_scrollbar.set,
                                                xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.declaration_info_tree.yview)
        h_scrollbar.config(command=self.declaration_info_tree.xview)
        
        # 设置表格列标题和宽度
        self.declaration_info_tree.heading("bill_of_lading", text="提单号")
        self.declaration_info_tree.heading("skc", text="SKC")
        self.declaration_info_tree.heading("product_name", text="商品品名")
        self.declaration_info_tree.heading("hs_code", text="HS CODE")
        
        self.declaration_info_tree.column("bill_of_lading", width=150, anchor=tk.CENTER)
        self.declaration_info_tree.column("skc", width=100, anchor=tk.CENTER)
        self.declaration_info_tree.column("product_name", width=200, anchor=tk.CENTER)
        self.declaration_info_tree.column("hs_code", width=100, anchor=tk.CENTER)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.declaration_info_tree.pack(fill=tk.BOTH, expand=True)
        
        # 刷新报关信息
        self.refresh_declaration_info()
    
    def init_other_info_tab(self):
        """初始化其他信息维护界面"""
        # 创建标签页
        self.other_info_notebook = ttk.Notebook(self.other_info_tab)
        self.other_info_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建国家代码维护子标签页
        self.country_code_tab = ttk.Frame(self.other_info_notebook)
        self.other_info_notebook.add(self.country_code_tab, text="国家代码维护")
        
        # 创建币种维护子标签页
        self.currency_tab = ttk.Frame(self.other_info_notebook)
        self.other_info_notebook.add(self.currency_tab, text="币种维护")
        
        # 创建店铺对应公司维护子标签页
        self.shop_company_tab = ttk.Frame(self.other_info_notebook)
        self.other_info_notebook.add(self.shop_company_tab, text="店铺对应公司维护")
        
        # 创建申报金额维护子标签页
        self.declaration_amount_tab = ttk.Frame(self.other_info_notebook)
        self.other_info_notebook.add(self.declaration_amount_tab, text="申报金额维护")
        
        # 初始化国家代码维护界面
        self.init_country_code_tab()
        
        # 初始化币种维护界面
        self.init_currency_tab()
        
        # 初始化店铺对应公司维护界面
        self.init_shop_company_tab()
        
        # 初始化申报金额维护界面
        self.init_declaration_amount_tab()
    
    def init_customs_tab(self):
        """初始化报关单管理界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.customs_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="生成报关单数据", command=self.generate_customs_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新报关单数据", command=self.refresh_customs_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导出报关单数据", command=self.export_customs_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="按公司导出报关单数据", command=self.export_customs_data_by_company).pack(side=tk.LEFT, padx=5)
        
        # 添加搜索框架
        search_frame = ttk.Frame(self.customs_tab)
        search_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Label(search_frame, text="搜索字段：").pack(side=tk.LEFT, padx=5)
        self.customs_search_field = ttk.Combobox(search_frame, width=20)
        self.customs_search_field.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(search_frame, text="搜索值：").pack(side=tk.LEFT, padx=5)
        self.customs_search_value = ttk.Entry(search_frame, width=20)
        self.customs_search_value.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(search_frame, text="搜索", command=self.filter_customs_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="重置", command=self.refresh_customs_data).pack(side=tk.LEFT, padx=5)
        
        # 创建报关单数据表格
        table_frame = ttk.Frame(self.customs_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        columns = [
            "提单号", "订单编号", "大箱号", "快件单号", "包裹号（运单号）", "目的国", "SKC", 
            "敏感品类别", "商品品名", "HS CODE", "规格型号", "包裹内单个SKC的商品数量", 
            "申报单位", "商品申报单价", "申报币制", "商品总净重(KG)", "第一法定数量", 
            "第一法定单位", "第二法定数量", "第二法定单位", "电商企业代码", "电商企业名称", 
            "电商平台代码", "电商平台名称", "收款企业代码", "收款企业名称", "生产企业代码", 
            "生产企业名称", "电商企业dxpId"
        ]
        
        self.customs_tree = ttk.Treeview(table_frame, 
                                      columns=columns, 
                                      show="headings",
                                      yscrollcommand=v_scrollbar.set,
                                      xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.customs_tree.yview)
        h_scrollbar.config(command=self.customs_tree.xview)
        
        # 设置表格列标题和宽度
        for column in columns:
            self.customs_tree.heading(column, text=column)
            self.customs_tree.column(column, width=120, anchor=tk.CENTER)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.customs_tree.pack(fill=tk.BOTH, expand=True)
        
        # 初始化报关单数据
        self.customs_data = []
        
        # 刷新报关单数据
        self.refresh_customs_data()
    
    def init_country_code_tab(self):
        """初始化国家代码维护界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.country_code_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="刷新国家代码", command=self.refresh_country_codes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="添加国家代码", command=self.add_country_code).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="编辑国家代码", command=self.edit_country_code).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除国家代码", command=self.delete_country_code).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存国家代码", command=self.save_country_codes).pack(side=tk.LEFT, padx=5)
        
        # 创建国家代码表格
        table_frame = ttk.Frame(self.country_code_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        self.country_code_tree = ttk.Treeview(table_frame, 
                                            columns=["consignee_country", "three_letter_code"], 
                                            show="headings",
                                            yscrollcommand=v_scrollbar.set,
                                            xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.country_code_tree.yview)
        h_scrollbar.config(command=self.country_code_tree.xview)
        
        # 设置表格列标题和宽度
        self.country_code_tree.heading("consignee_country", text="收件人国家/Consignee Country")
        self.country_code_tree.heading("three_letter_code", text="3字码")
        
        self.country_code_tree.column("consignee_country", width=300, anchor=tk.CENTER)
        self.country_code_tree.column("three_letter_code", width=150, anchor=tk.CENTER)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.country_code_tree.pack(fill=tk.BOTH, expand=True)
        
        # 刷新国家代码
        self.refresh_country_codes()
    
    def init_currency_tab(self):
        """初始化币种维护界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.currency_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="刷新币种数据", command=self.refresh_currency_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="添加币种", command=self.add_currency).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存币种数据", command=self.save_currency_data).pack(side=tk.LEFT, padx=5)
        
        # 创建币种表格
        table_frame = ttk.Frame(self.currency_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        self.currency_tree = ttk.Treeview(table_frame, 
                                         columns=["country_name", "currency_code"], 
                                         show="headings",
                                         yscrollcommand=v_scrollbar.set,
                                         xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.currency_tree.yview)
        h_scrollbar.config(command=self.currency_tree.xview)
        
        # 设置表格列标题和宽度
        self.currency_tree.heading("country_name", text="国家名称")
        self.currency_tree.heading("currency_code", text="币种代码")
        
        self.currency_tree.column("country_name", width=200, anchor=tk.CENTER)
        self.currency_tree.column("currency_code", width=150, anchor=tk.CENTER)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.currency_tree.pack(fill=tk.BOTH, expand=True)
        
        # 刷新币种数据
        self.refresh_currency_data()
    
    def init_shop_company_tab(self):
        """初始化店铺对应公司维护界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.shop_company_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="刷新店铺数据", command=self.refresh_shop_company_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="添加店铺", command=self.add_shop_company).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="编辑店铺", command=self.edit_shop_company).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存店铺数据", command=self.save_shop_company_data).pack(side=tk.LEFT, padx=5)
        
        # 创建店铺对应公司表格
        table_frame = ttk.Frame(self.shop_company_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        self.shop_company_tree = ttk.Treeview(table_frame, 
                                             columns=["shop_name", "company_name"], 
                                             show="headings",
                                             yscrollcommand=v_scrollbar.set,
                                             xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.shop_company_tree.yview)
        h_scrollbar.config(command=self.shop_company_tree.xview)
        
        # 设置表格列标题和宽度
        self.shop_company_tree.heading("shop_name", text="店铺名称")
        self.shop_company_tree.heading("company_name", text="所属公司")
        
        self.shop_company_tree.column("shop_name", width=200, anchor=tk.CENTER)
        self.shop_company_tree.column("company_name", width=200, anchor=tk.CENTER)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.shop_company_tree.pack(fill=tk.BOTH, expand=True)
        
        # 刷新店铺数据
        self.refresh_shop_company_data()
    
    def init_declaration_amount_tab(self):
        """初始化申报金额维护界面"""
        # 创建操作按钮
        button_frame = ttk.Frame(self.declaration_amount_tab)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="刷新申报金额规则", command=self.refresh_declaration_amount_rules).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="添加申报金额规则", command=self.add_declaration_amount_rule).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="编辑申报金额规则", command=self.edit_declaration_amount_rule).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除申报金额规则", command=self.delete_declaration_amount_rule).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存申报金额规则", command=self.save_declaration_amount_rules).pack(side=tk.LEFT, padx=5)
        
        # 创建申报金额规则表格
        table_frame = ttk.Frame(self.declaration_amount_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建水平和垂直滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建Treeview表格
        self.declaration_amount_tree = ttk.Treeview(table_frame, 
                                            columns=["country_name", "declaration_ratio", "max_declaration_amount"], 
                                            show="headings",
                                            yscrollcommand=v_scrollbar.set,
                                            xscrollcommand=h_scrollbar.set)
        
        # 配置滚动条
        v_scrollbar.config(command=self.declaration_amount_tree.yview)
        h_scrollbar.config(command=self.declaration_amount_tree.xview)
        
        # 设置表格列标题和宽度
        self.declaration_amount_tree.heading("country_name", text="国家名称")
        self.declaration_amount_tree.heading("declaration_ratio", text="申报比例")
        self.declaration_amount_tree.heading("max_declaration_amount", text="最高申报金额")
        
        self.declaration_amount_tree.column("country_name", width=200, anchor=tk.CENTER)
        self.declaration_amount_tree.column("declaration_ratio", width=150, anchor=tk.CENTER)
        self.declaration_amount_tree.column("max_declaration_amount", width=150, anchor=tk.CENTER)
        
        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", font=("Arial", 9))
        
        self.declaration_amount_tree.pack(fill=tk.BOTH, expand=True)
        
        # 刷新申报金额规则
        self.refresh_declaration_amount_rules()
    
    def import_logistics_file(self):
        """导入物流信息表 - 增强版"""
        try:
            print("GUI调试: 开始导入物流信息表")
            self.status_var.set("正在导入物流信息表...")
            
            # 获取根窗口进行UI更新
            root_window = self._get_root_window()
            if root_window:
                print("GUI调试: 执行UI更新")
                root_window.update_idletasks()
            
            # 调用处理器的导入方法并获取返回值
            import_success = self.processor.import_logistics_file()
            print(f"GUI调试: 导入结果: {import_success}")
            
            if import_success:
                # 重新获取数据并进行完整性检查
                logistics_data = self.processor.get_logistics_data()
                print(f"GUI调试: 导入后物流数据长度: {len(logistics_data)}")
                
                # 详细验证数据
                if logistics_data:
                    print(f"GUI调试: 数据验证 - 第一行数据: {logistics_data[0]}")
                    print(f"GUI调试: 数据验证 - 字段数量: {len(logistics_data[0])}")
                    print(f"GUI调试: 数据验证 - 字段列表: {list(logistics_data[0].keys())}")
                else:
                    print("GUI调试: 警告 - 导入返回成功但数据为空")
                    # 即使导入返回成功但数据为空，也显示警告
                    messagebox.showwarning("警告", "导入成功但未找到有效数据")
                    self.status_var.set("就绪")
                    return
                
                # 检查logistics_tree是否存在
                if hasattr(self, 'logistics_tree'):
                    print(f"GUI调试: logistics_tree存在，准备刷新")
                    
                    # 强制刷新表格 - 使用增强的刷新机制
                    refresh_result = self.refresh_logistics_data()
                    print(f"GUI调试: 表格刷新结果: {refresh_result}")
                    
                    # 强制更新UI多次，确保表格完全刷新
                    if root_window:
                        print(f"GUI调试: 开始强制UI更新循环")
                        # 多次更新确保刷新
                        for i in range(5):  # 增加更新次数
                            try:
                                print(f"GUI调试: 执行UI更新轮次 {i+1}/5")
                                root_window.update_idletasks()
                                root_window.update()
                                # 短暂延迟让UI有时间响应
                                import time
                                time.sleep(0.05)
                            except Exception as e_update:
                                print(f"GUI调试: UI更新轮次 {i+1} 失败: {str(e_update)}")
                    
                    # 验证刷新后的表格状态
                    try:
                        item_count = len(self.logistics_tree.get_children())
                        print(f"GUI调试: 刷新后表格中的行数: {item_count}")
                        
                        # 如果表格为空但应该有数据，显示警告
                        if item_count == 0 and len(logistics_data) > 0:
                            print("GUI调试: 错误 - 表格中没有显示数据，但应该有数据")
                            messagebox.showwarning("警告", "数据已导入但表格未能正确显示，请尝试手动点击刷新按钮")
                    except Exception as e_verify:
                        print(f"GUI调试: 验证表格状态时出错: {str(e_verify)}")
                else:
                    print("GUI调试: 警告 - logistics_tree不存在")
                    messagebox.showwarning("警告", "表格组件未初始化，请重启应用")
                
                self.status_var.set(f"物流信息表导入完成，共 {len(logistics_data)} 条记录")
                messagebox.showinfo("成功", f"物流信息表导入完成，共 {len(logistics_data)} 条记录")
            else:
                print("GUI调试: 导入失败或用户取消")
                self.status_var.set("就绪")
                # 如果是用户取消，不显示错误消息
        except Exception as e:
            print(f"GUI调试: 导入物流信息表异常: {str(e)}")
            import traceback
            traceback.print_exc()
            self.status_var.set(f"错误：{e}")
            messagebox.showerror("错误", f"导入物流信息表失败：{e}")
            
            # 发生异常后清空处理器数据，避免部分数据导致的问题
            try:
                if hasattr(self.processor, 'logistics_data'):
                    print("GUI调试: 清空处理器数据以避免数据不一致")
                    self.processor.logistics_data = []
            except Exception:
                pass
    
    def _get_root_window(self):
        """获取根窗口的辅助方法，提供多种方式确保能获取到root窗口"""
        # 方式1: 直接返回self.root
        if hasattr(self, 'root') and self.root is not None:
            print("GUI调试: 通过self.root获取根窗口成功")
            return self.root
        
        # 方式2: 如果logistics_tree存在，通过其master属性查找根窗口
        elif hasattr(self, 'logistics_tree'):
            print("GUI调试: 尝试通过logistics_tree查找根窗口")
            current = self.logistics_tree
            # 向上遍历组件树找到根窗口
            while current.master:
                current = current.master
            print("GUI调试: 通过组件树找到根窗口")
            return current
        
        # 方式3: 使用tk.Tk()的全局实例
        else:
            print("GUI调试: 尝试获取全局根窗口")
            try:
                import tkinter as tk
                # 尝试获取已存在的根窗口
                for window in tk._default_root.tk.call('winfo', 'children', '.'):
                    if window.startswith('.'):
                        return tk.Tk()
                # 如果找不到，创建一个新的根窗口（但这通常不会被使用）
                return tk.Tk()
            except Exception as e:
                print(f"GUI调试: 获取全局根窗口失败: {str(e)}")
                return None
    
    def process_logistics_data(self):
        """处理物流信息，抓取实际发货物流字段的"US"和红底的"DHL"并保存"""
        try:
            self.status_var.set("正在处理物流信息...")
            self.processor.process_logistics_data()
            self.refresh_logistics_data()
            self.status_var.set("物流信息处理完成")
            messagebox.showinfo("成功", "物流信息处理完成")
        except Exception as e:
            self.status_var.set(f"错误：{e}")
            messagebox.showerror("错误", f"处理物流信息失败：{e}")
    
    def refresh_logistics_data(self):
        """刷新物流数据表格 - 增强版"""
        print("GUI调试: 开始强制刷新物流数据表格")
        
        # 获取根窗口用于UI更新
        root_window = self._get_root_window()
        
        try:
            # 立即更新状态栏显示正在刷新
            self.status_var.set("正在强制刷新物流数据...")
            if root_window:
                root_window.update_idletasks()
            
            # 检查logistics_tree是否存在
            if not hasattr(self, 'logistics_tree'):
                print("GUI调试: 错误 - logistics_tree不存在")
                self.status_var.set("错误：表格组件未初始化")
                return
            
            # 步骤1: 强制重新获取数据（确保获取最新数据）
            print("GUI调试: 强制重新获取物流数据")
            logistics_data = self.processor.get_logistics_data()
            print(f"GUI调试: 获取到的物流数据长度: {len(logistics_data)}")
            
            # 数据完整性验证
            data_valid = False
            if logistics_data:
                print(f"GUI调试: 数据验证 - 存在{len(logistics_data)}条记录")
                # 验证数据结构
                if isinstance(logistics_data[0], dict) and logistics_data[0]:
                    data_valid = True
                    print(f"GUI调试: 数据验证 - 数据结构有效，包含{len(logistics_data[0])}个字段")
                else:
                    print("GUI调试: 数据验证 - 数据结构无效")
            
            # 强制刷新机制 - 完全重建表格
            print("GUI调试: 强制刷新机制：开始完全重建表格")
            
            # 步骤2: 保存当前表格的父容器
            parent_frame = self.logistics_tree.master
            
            # 步骤3: 完全销毁现有表格
            print("GUI调试: 销毁现有表格组件")
            self.logistics_tree.destroy()
            
            # UI更新以确保销毁生效
            if root_window:
                root_window.update_idletasks()
                try:
                    root_window.update()
                except Exception:
                    pass
            
            # 步骤4: 重新创建滚动条（如果需要）
            print("GUI调试: 重新创建滚动条")
            # 移除可能存在的旧滚动条
            for child in parent_frame.winfo_children():
                if isinstance(child, ttk.Scrollbar):
                    child.destroy()
            
            # 创建新的滚动条
            h_scrollbar = ttk.Scrollbar(parent_frame, orient=tk.HORIZONTAL)
            h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            
            v_scrollbar = ttk.Scrollbar(parent_frame, orient=tk.VERTICAL)
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 步骤5: 重新创建Treeview表格
            print("GUI调试: 重新创建Treeview表格")
            self.logistics_tree = ttk.Treeview(parent_frame,
                                            columns=[],
                                            show="headings",
                                            yscrollcommand=v_scrollbar.set,
                                            xscrollcommand=h_scrollbar.set)
            
            # 配置滚动条
            v_scrollbar.config(command=self.logistics_tree.yview)
            h_scrollbar.config(command=self.logistics_tree.xview)
            
            # 设置表格样式
            style = ttk.Style()
            style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
            style.configure("Treeview", font=("Arial", 9))
            
            # 重新放置表格
            self.logistics_tree.pack(fill=tk.BOTH, expand=True)
            
            # UI更新以确保创建生效
            if root_window:
                root_window.update_idletasks()
            
            # 步骤6: 检查数据并显示
            if data_valid and logistics_data:
                print("GUI调试: 数据有效，开始设置新表格")
                
                # 设置表格列
                headers = list(logistics_data[0].keys())
                print(f"GUI调试: 表头信息: {headers}")
                self.logistics_tree['columns'] = headers
                
                # 设置列标题和宽度
                for header in headers:
                    self.logistics_tree.heading(header, text=header)
                    self.logistics_tree.column(header, width=120, anchor=tk.CENTER)
                
                # 更新筛选字段下拉列表
                if hasattr(self, 'logistics_filter_field'):
                    self.logistics_filter_field['values'] = headers
                    if headers:
                        self.logistics_filter_field.current(0)
                
                # 批量插入数据行
                print("GUI调试: 开始批量插入数据行")
                row_count = 0
                
                try:
                    # 逐行插入并实时更新UI，避免大数量数据时界面卡顿
                    for i, row in enumerate(logistics_data):
                        try:
                            # 确保所有值都是字符串类型
                            row_values = [str(value) if value is not None else "" for value in row.values()]
                            is_matched = row.get("is_matched", True)
                            
                            # 插入数据
                            if not is_matched:
                                self.logistics_tree.insert("", tk.END, values=row_values, tags=("unmatched",))
                            else:
                                self.logistics_tree.insert("", tk.END, values=row_values)
                            
                            row_count += 1
                            
                            # 每插入50行更新一次UI
                            if i > 0 and i % 50 == 0 and root_window:
                                self.status_var.set(f"正在加载数据... {i+1}/{len(logistics_data)}")
                                root_window.update_idletasks()
                        except Exception as e_row:
                            print(f"GUI调试: 插入第{i+1}行数据时出错: {str(e_row)}")
                    
                    print(f"GUI调试: 数据插入完成，共插入 {row_count} 行")
                    
                    # 配置标签样式
                    try:
                        self.logistics_tree.tag_configure("unmatched", background="#FFCCCC")
                    except Exception as e_style:
                        print(f"GUI调试: 配置标签样式时出错: {str(e_style)}")
                    
                    # 最终更新状态栏
                    self.status_var.set(f"物流数据刷新完成，共显示 {row_count} 条记录")
                    
                except Exception as e_insert:
                    print(f"GUI调试: 数据插入过程中出错: {str(e_insert)}")
                    self.status_var.set(f"错误：数据加载失败: {str(e_insert)}")
            else:
                # 没有数据时，显示空表格和提示信息
                print("GUI调试: 没有有效数据，显示空表格")
                self.logistics_tree['columns'] = ["提示信息"]
                self.logistics_tree.heading("提示信息", text="提示信息")
                self.logistics_tree.column("提示信息", width=400, anchor=tk.CENTER)
                self.logistics_tree.insert("", tk.END, values=["暂无数据，请导入物流信息表"])
                self.status_var.set("暂无数据，请导入物流信息表")
            
            # 最终UI更新
            if root_window:
                for _ in range(3):  # 多轮UI更新确保界面正确刷新
                    root_window.update_idletasks()
                    try:
                        root_window.update()
                    except Exception:
                        pass
        except Exception as e:
            print(f"GUI调试: 刷新物流数据时发生错误: {str(e)}")
            import traceback
            traceback.print_exc()
            self.status_var.set(f"错误：刷新数据失败: {str(e)}")
            
            # 发生异常时，尝试显示错误信息
            try:
                if hasattr(self, 'logistics_tree') and self.logistics_tree.winfo_exists():
                    for item in self.logistics_tree.get_children():
                        self.logistics_tree.delete(item)
                    self.logistics_tree['columns'] = ["错误信息"]
                    self.logistics_tree.heading("错误信息", text="错误信息")
                    self.logistics_tree.column("错误信息", width=400, anchor=tk.CENTER)
                    self.logistics_tree.insert("", tk.END, values=[f"加载数据时发生错误: {str(e)}"])
            except Exception:
                pass
                # 数据为空时显示提示信息
                self.status_var.set("物流数据为空，请先导入物流信息表")
                # 设置一个空列，避免表格显示异常
                self.logistics_tree['columns'] = ["提示"]
                self.logistics_tree.heading("提示", text="提示信息")
                self.logistics_tree.column("提示", width=300, anchor=tk.CENTER)
                self.logistics_tree.insert("", tk.END, values=["暂无数据，请导入物流信息表"])
            
            # 步骤7: 多重UI强制更新
            print("GUI调试: 执行多重UI强制更新")
            if root_window:
                # 连续多次更新确保UI完全刷新
                for i in range(3):
                    try:
                        print(f"GUI调试: UI更新轮次 {i+1}")
                        root_window.update_idletasks()
                        root_window.update()
                        # 添加短暂延迟让UI有时间响应
                        import time
                        time.sleep(0.05)
                    except Exception as e_update:
                        print(f"GUI调试: UI更新轮次 {i+1} 失败: {str(e_update)}")
            
        except Exception as e:
            print(f"GUI调试: 强制刷新物流数据异常: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # 即使出错也要尝试恢复表格显示
            try:
                if hasattr(self, 'logistics_tree'):
                    # 清空表格并显示错误信息
                    for item in self.logistics_tree.get_children():
                        self.logistics_tree.delete(item)
                    
                    self.logistics_tree['columns'] = ["错误信息"]
                    self.logistics_tree.heading("错误信息", text="错误信息")
                    self.logistics_tree.column("错误信息", width=500, anchor=tk.CENTER)
                    self.logistics_tree.insert("", tk.END, values=[f"刷新数据时出错: {str(e)}"])
            except Exception:
                pass
            
            self.status_var.set(f"刷新物流数据失败：{e}")
            messagebox.showerror("错误", f"刷新物流数据失败：{e}")
        
        print("GUI调试: 强制刷新流程结束")
    
    def import_order_file(self):
        """导入订单信息表"""
        try:
            self.status_var.set("正在导入订单信息表...")
            self.processor.import_order_file()
            self.refresh_order_data()
            self.status_var.set("订单信息表导入完成")
            messagebox.showinfo("成功", "订单信息表导入完成")
        except Exception as e:
            self.status_var.set(f"错误：{e}")
            messagebox.showerror("错误", f"导入订单信息表失败：{e}")
    
    def match_order_data(self):
        """根据物流信息匹配订单数据"""
        try:
            self.status_var.set("正在匹配订单数据...")
            self.processor.match_order_data()
            self.refresh_order_data()
            self.status_var.set("订单数据匹配完成")
            messagebox.showinfo("成功", "订单数据匹配完成")
        except Exception as e:
            self.status_var.set(f"错误：{e}")
            messagebox.showerror("错误", f"匹配订单数据失败：{e}")
    
    def filter_logistics_data(self):
        """筛选物流数据"""
        try:
            self.status_var.set("正在筛选物流数据...")
            
            # 获取筛选条件
            filter_field = self.logistics_filter_field.get()
            filter_value = self.logistics_filter_value.get().strip()
            
            if not filter_field:
                messagebox.showwarning("警告", "请选择筛选字段")
                return
            
            # 清空现有表格
            for item in self.logistics_tree.get_children():
                self.logistics_tree.delete(item)
            
            # 加载物流数据
            logistics_data = self.processor.get_logistics_data()
            
            if logistics_data:
                # 过滤数据
                filtered_data = []
                for row in logistics_data:
                    if filter_field in row:
                        field_value = str(row[filter_field])
                        if filter_value in field_value:
                            filtered_data.append(row)
                
                # 插入过滤后的数据
                for row in filtered_data:
                    # 检查是否匹配
                    is_matched = row.get("is_matched", True)
                    # 根据匹配情况设置样式
                    if not is_matched:
                        self.logistics_tree.insert("", tk.END, values=list(row.values()), tags=("unmatched",))
                    else:
                        self.logistics_tree.insert("", tk.END, values=list(row.values()))
                
                self.status_var.set(f"物流数据筛选完成，共显示 {len(filtered_data)} 条记录")
        except Exception as e:
            self.status_var.set(f"筛选物流数据失败：{e}")
            messagebox.showerror("错误", f"筛选物流数据失败：{e}")
    
    def filter_order_data(self):
        """筛选订单数据"""
        try:
            self.status_var.set("正在筛选订单数据...")
            
            # 获取筛选条件
            filter_field = self.order_filter_field.get()
            filter_value = self.order_filter_value.get().strip()
            
            if not filter_field:
                messagebox.showwarning("警告", "请选择筛选字段")
                return
            
            # 清空现有表格
            for item in self.order_tree.get_children():
                self.order_tree.delete(item)
            
            # 加载订单数据
            order_data = self.processor.get_order_data()
            
            if order_data:
                # 过滤数据
                filtered_data = []
                for row in order_data:
                    if filter_field in row:
                        field_value = str(row[filter_field])
                        if filter_value in field_value:
                            filtered_data.append(row)
                
                # 插入过滤后的数据
                for row in filtered_data:
                    # 检查是否匹配
                    is_matched = row.get("is_matched", True)
                    # 根据匹配情况设置样式
                    if not is_matched:
                        self.order_tree.insert("", tk.END, values=list(row.values()), tags=("unmatched",))
                    else:
                        self.order_tree.insert("", tk.END, values=list(row.values()))
                
                self.status_var.set(f"订单数据筛选完成，共显示 {len(filtered_data)} 条记录")
        except Exception as e:
            self.status_var.set(f"筛选订单数据失败：{e}")
            messagebox.showerror("错误", f"筛选订单数据失败：{e}")
    
    def refresh_order_data(self):
        """刷新订单数据表格"""
        try:
            self.status_var.set("正在刷新订单数据...")
            
            # 清空现有表格
            for item in self.order_tree.get_children():
                self.order_tree.delete(item)
            
            # 清空现有列
            self.order_tree['columns'] = []
            
            # 加载订单数据
            order_data = self.processor.get_order_data()
            
            if order_data:
                # 设置表格列
                headers = list(order_data[0].keys())
                self.order_tree['columns'] = headers
                
                # 设置列标题和宽度
                for header in headers:
                    self.order_tree.heading(header, text=header)
                    self.order_tree.column(header, width=120, anchor=tk.CENTER)
                
                # 更新筛选字段下拉列表
                self.order_filter_field['values'] = headers
                if headers:
                    self.order_filter_field.current(0)
                
                # 插入数据行
                for row in order_data:
                    # 检查是否匹配
                    is_matched = row.get("is_matched", True)
                    # 根据匹配情况设置样式
                    if not is_matched:
                        self.order_tree.insert("", tk.END, values=list(row.values()), tags=("unmatched",))
                    else:
                        self.order_tree.insert("", tk.END, values=list(row.values()))
                
                # 配置标签样式
                self.order_tree.tag_configure("unmatched", background="#FFCCCC")
            
            self.status_var.set("订单数据刷新完成")
        except Exception as e:
            self.status_var.set(f"刷新订单数据失败：{e}")
            messagebox.showerror("错误", f"刷新订单数据失败：{e}")
    
    def refresh_template_fields(self):
        """刷新模板字段表格"""
        try:
            self.status_var.set("正在刷新模板字段...")
            
            # 清空现有表格
            for item in self.template_fields_tree.get_children():
                self.template_fields_tree.delete(item)
            
            # 加载模板字段
            template_fields = self.processor.get_declaration_template_fields()
            
            # 插入模板字段数据
            for field in template_fields:
                self.template_fields_tree.insert("", tk.END, values=(field["field_name"], field["field_type"], "是" if field["required"] else "否"))
            
            self.status_var.set("模板字段刷新完成")
        except Exception as e:
            self.status_var.set(f"刷新模板字段失败：{e}")
            messagebox.showerror("错误", f"刷新模板字段失败：{e}")
    
    def save_template_settings(self):
        """保存模板设置"""
        try:
            self.status_var.set("正在保存模板设置...")
            # 这里可以添加保存模板设置的逻辑
            self.status_var.set("模板设置保存成功")
            messagebox.showinfo("成功", "模板设置保存成功")
        except Exception as e:
            self.status_var.set(f"错误：{e}")
            messagebox.showerror("错误", f"保存模板设置失败：{e}")
    
    def refresh_declaration_info(self):
        """刷新报关信息表格"""
        try:
            self.status_var.set("正在刷新报关信息...")
            
            # 清空现有表格
            for item in self.declaration_info_tree.get_children():
                self.declaration_info_tree.delete(item)
            
            # 清空现有列
            self.declaration_info_tree['columns'] = []
            
            # 加载报关信息
            declaration_info = self.processor.get_declaration_info()
            
            if declaration_info:
                # 设置表格列
                headers = list(declaration_info[0].keys())
                self.declaration_info_tree['columns'] = headers
                
                # 设置列标题和宽度
                for header in headers:
                    self.declaration_info_tree.heading(header, text=header)
                    self.declaration_info_tree.column(header, width=120, anchor=tk.CENTER)
                
                # 插入报关信息数据
                for info in declaration_info:
                    self.declaration_info_tree.insert("", tk.END, values=list(info.values()))
            
            self.status_var.set("报关信息刷新完成")
        except Exception as e:
            self.status_var.set(f"刷新报关信息失败：{e}")
            messagebox.showerror("错误", f"刷新报关信息失败：{e}")
    
    def add_declaration_info(self):
        """添加报关信息"""
        # 创建添加报关信息对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("添加报关信息")
        dialog.geometry("800x600")
        dialog.resizable(True, True)
        
        # 设置对话框居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_reqwidth()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_reqheight()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 创建滚动条框架
        scroll_frame = ttk.Frame(dialog)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建滚动条
        canvas = tk.Canvas(scroll_frame)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 布局滚动条和画布
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 创建表单字段
        fields = [
            "提单号", "商品品名", "HS CODE", "规格型号",
            "包裹内单个SKC商品数量", "申报单位", "申报币制", "商品总净重(KG)",
            "第一法定数量", "第一法定单位", "电商企业代码", "电商企业名称",
            "电商平台代码", "电商平台名称", "收款企业代码", "收款企业名称",
            "生产企业代码", "生产企业名称", "电商企业dxpId"
        ]
        
        # 创建输入框字典
        entry_widgets = {}
        
        # 布局表单字段
        for i, field in enumerate(fields):
            row = i // 2
            col = i % 2
            
            # 创建标签
            label = ttk.Label(scrollable_frame, text=f"{field}：")
            label.grid(row=row, column=col*2, padx=10, pady=10, sticky=tk.W)
            
            # 创建输入框
            entry = ttk.Entry(scrollable_frame, width=30)
            entry.grid(row=row, column=col*2+1, padx=10, pady=10, sticky=tk.W)
            entry_widgets[field] = entry
        
        # 保存按钮回调
        def save_declaration_info():
            # 收集表单数据
            declaration_data = {}
            for field, entry in entry_widgets.items():
                declaration_data[field] = entry.get().strip()
            
            # 验证必填字段
            required_fields = ["提单号", "商品品名", "HS CODE", "申报单位", "申报币制", "电商企业代码", "电商企业名称", "电商平台代码", "电商平台名称", "收款企业代码", "收款企业名称", "生产企业代码", "生产企业名称", "电商企业dxpId"]
            for field in required_fields:
                if not declaration_data[field]:
                    messagebox.showwarning("警告", f"请填写必填字段：{field}")
                    return
            
            # 调用处理器保存报关信息
            try:
                self.processor.add_declaration_info(declaration_data)
                self.refresh_declaration_info()
                dialog.destroy()
                messagebox.showinfo("成功", "报关信息添加成功")
            except Exception as e:
                messagebox.showerror("错误", f"添加报关信息失败：{e}")
        
        # 创建按钮框架
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=len(fields)//2 + 1, column=0, columnspan=4, pady=20)
        
        # 创建保存和取消按钮
        ttk.Button(button_frame, text="保存", command=save_declaration_info).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
    
    def edit_declaration_info(self):
        """修改报关信息"""
        messagebox.showinfo("提示", "修改报关信息功能正在开发中")
    
    def delete_declaration_info(self):
        """删除报关信息"""
        messagebox.showinfo("提示", "删除报关信息功能正在开发中")
    
    def refresh_country_codes(self):
        """刷新国家代码表格"""
        try:
            self.status_var.set("正在刷新国家代码...")
            
            # 清空现有表格
            for item in self.country_code_tree.get_children():
                self.country_code_tree.delete(item)
            
            # 加载国家代码
            country_codes = self.processor.country_codes  # 直接访问属性，因为get_country_codes方法已移除
            
            # 插入国家代码数据，使用get方法安全地获取可能不存在的键
            for code in country_codes:
                consignee_country = code.get("consignee_country", "")  # 安全获取，不存在则返回空字符串
                three_letter_code = code.get("three_letter_code", "")  # 安全获取，不存在则返回空字符串
                self.country_code_tree.insert("", tk.END, values=(consignee_country, three_letter_code))
            
            self.status_var.set("国家代码刷新完成")
        except Exception as e:
            self.status_var.set(f"刷新国家代码失败：{e}")
            messagebox.showerror("错误", f"刷新国家代码失败：{e}")
    
    def add_country_code(self):
        """添加国家代码"""
        # 创建添加国家代码对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("添加国家代码")
        dialog.geometry("500x250")
        dialog.resizable(False, False)
        
        # 设置对话框居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_reqwidth()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_reqheight()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 创建表单框架
        form_frame = ttk.Frame(dialog, padding="10")
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建收件人国家字段
        ttk.Label(form_frame, text="收件人国家/Consignee Country：").grid(row=0, column=0, padx=10, pady=10, sticky=tk.E)
        consignee_country_entry = ttk.Entry(form_frame, width=30)
        consignee_country_entry.grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)
        
        # 创建3字码字段
        ttk.Label(form_frame, text="3字码：").grid(row=1, column=0, padx=10, pady=10, sticky=tk.E)
        three_letter_code_entry = ttk.Entry(form_frame, width=30)
        three_letter_code_entry.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)
        
        # 保存按钮回调
        def save_country_code():
            # 获取表单数据
            consignee_country = consignee_country_entry.get().strip()
            three_letter_code = three_letter_code_entry.get().strip()
            
            # 验证必填字段
            if not consignee_country or not three_letter_code:
                messagebox.showwarning("警告", "请填写所有必填字段")
                return
            
            # 调用处理器保存国家代码
            try:
                self.processor.add_country_code(consignee_country, three_letter_code)
                self.refresh_country_codes()
                dialog.destroy()
                messagebox.showinfo("成功", "国家代码添加成功")
            except Exception as e:
                messagebox.showerror("错误", f"添加国家代码失败：{e}")
        
        # 创建按钮框架
        button_frame = ttk.Frame(dialog, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 创建保存和取消按钮
        ttk.Button(button_frame, text="保存", command=save_country_code).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
    
    def delete_country_code(self):
        """删除国家代码"""
        try:
            # 获取选中项
            selection = self.country_code_tree.selection()
            if not selection:
                messagebox.showwarning("警告", "请选择要删除的国家代码")
                return
            
            # 获取选中项数据
            item = selection[0]
            values = self.country_code_tree.item(item, "values")
            consignee_country = values[0]  # 收件人国家用于删除
            
            # 显示确认对话框
            confirm = messagebox.askyesno("确认删除", f"确定要删除国家代码 '{consignee_country}' 吗？")
            if not confirm:
                return
            
            # 调用处理器删除国家代码
            self.processor.delete_country_code(consignee_country)
            
            # 刷新国家代码表格
            self.refresh_country_codes()
            
            # 显示成功消息
            messagebox.showinfo("成功", "国家代码删除成功")
            
        except Exception as e:
            messagebox.showerror("错误", f"删除国家代码失败：{e}")
    
    def edit_country_code(self):
        """编辑国家代码"""
        # 获取选中项
        selection = self.country_code_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请选择要编辑的国家代码")
            return
        
        # 获取选中项数据
        item = selection[0]
        values = self.country_code_tree.item(item, "values")
        original_consignee_country = values[0]
        original_three_letter_code = values[1]
        
        # 创建编辑国家代码对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑国家代码")
        dialog.geometry("500x250")
        dialog.resizable(False, False)
        
        # 设置对话框居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_reqwidth()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_reqheight()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 创建表单框架
        form_frame = ttk.Frame(dialog, padding="10")
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建收件人国家字段
        ttk.Label(form_frame, text="收件人国家/Consignee Country：").grid(row=0, column=0, padx=10, pady=10, sticky=tk.E)
        consignee_country_entry = ttk.Entry(form_frame, width=30)
        consignee_country_entry.grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)
        consignee_country_entry.insert(0, original_consignee_country)
        
        # 创建3字码字段
        ttk.Label(form_frame, text="3字码：").grid(row=1, column=0, padx=10, pady=10, sticky=tk.E)
        three_letter_code_entry = ttk.Entry(form_frame, width=30)
        three_letter_code_entry.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)
        three_letter_code_entry.insert(0, original_three_letter_code)
        
        # 保存按钮回调
        def save_edit_country_code():
            # 获取表单数据
            consignee_country = consignee_country_entry.get().strip()
            three_letter_code = three_letter_code_entry.get().strip()
            
            # 验证必填字段
            if not consignee_country or not three_letter_code:
                messagebox.showwarning("警告", "请填写所有必填字段")
                return
            
            # 调用处理器保存编辑后的国家代码
            try:
                self.processor.edit_country_code(original_consignee_country, consignee_country, three_letter_code)
                self.refresh_country_codes()
                dialog.destroy()
                messagebox.showinfo("成功", "国家代码编辑成功")
            except Exception as e:
                messagebox.showerror("错误", f"编辑国家代码失败：{e}")
        
        # 创建按钮框架
        button_frame = ttk.Frame(dialog, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 创建保存和取消按钮
        ttk.Button(button_frame, text="保存", command=save_edit_country_code).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
    
    def save_country_codes(self):
        """保存国家代码"""
        try:
            self.status_var.set("正在保存国家代码...")
            # 调用processor的save_country_codes方法保存国家代码
            self.processor.save_country_codes()
            self.status_var.set("国家代码保存成功")
            messagebox.showinfo("成功", "国家代码保存成功")
        except Exception as e:
            self.status_var.set(f"错误：{e}")
            messagebox.showerror("错误", f"保存国家代码失败：{e}")
    
    def refresh_currency_data(self):
        """刷新币种数据表格"""
        try:
            self.status_var.set("正在刷新币种数据...")
            
            # 清空现有表格
            for item in self.currency_tree.get_children():
                self.currency_tree.delete(item)
            
            # 加载币种数据
            currency_data = self.processor.get_currency_data()
            
            # 插入币种数据
            for data in currency_data:
                self.currency_tree.insert("", tk.END, values=(data["country_name"], data["currency_code"]))
            
            self.status_var.set("币种数据刷新完成")
        except Exception as e:
            self.status_var.set(f"刷新币种数据失败：{e}")
            messagebox.showerror("错误", f"刷新币种数据失败：{e}")
    
    def add_currency(self):
        """添加币种"""
        messagebox.showinfo("提示", "添加币种功能正在开发中")
    
    def save_currency_data(self):
        """保存币种数据"""
        try:
            self.status_var.set("正在保存币种数据...")
            # 这里可以添加保存币种数据的逻辑
            self.status_var.set("币种数据保存成功")
            messagebox.showinfo("成功", "币种数据保存成功")
        except Exception as e:
            self.status_var.set(f"错误：{e}")
            messagebox.showerror("错误", f"保存币种数据失败：{e}")
    
    def refresh_declaration_amount_rules(self):
        """刷新申报金额规则表格"""
        try:
            self.status_var.set("正在刷新申报金额规则...")
            
            # 清空现有表格
            for item in self.declaration_amount_tree.get_children():
                self.declaration_amount_tree.delete(item)
            
            # 加载申报金额规则
            declaration_rules = self.processor.get_declaration_amount_rules()
            
            # 插入申报金额规则数据
            for rule in declaration_rules:
                self.declaration_amount_tree.insert("", tk.END, values=(rule["country_name"], rule["declaration_ratio"], rule["max_declaration_amount"]))
            
            self.status_var.set("申报金额规则刷新完成")
        except Exception as e:
            self.status_var.set(f"刷新申报金额规则失败：{e}")
            messagebox.showerror("错误", f"刷新申报金额规则失败：{e}")
    
    def add_declaration_amount_rule(self):
        """添加申报金额规则"""
        messagebox.showinfo("提示", "添加申报金额规则功能正在开发中")
    
    def save_declaration_amount_rules(self):
        """保存申报金额规则"""
        try:
            self.status_var.set("正在保存申报金额规则...")
            # 调用processor的save_declaration_amount_rules方法保存申报金额规则
            self.processor.save_declaration_amount_rules()
            self.status_var.set("申报金额规则保存成功")
            messagebox.showinfo("成功", "申报金额规则保存成功")
        except Exception as e:
            self.status_var.set(f"错误：{e}")
            messagebox.showerror("错误", f"保存申报金额规则失败：{e}")
    
    def edit_declaration_amount_rule(self):
        """编辑申报金额规则"""
        try:
            # 获取选中的规则
            selection = self.declaration_amount_tree.selection()
            if not selection:
                messagebox.showwarning("警告", "请选择要编辑的申报金额规则")
                return
            
            # 获取选中规则的数据
            item = selection[0]
            values = self.declaration_amount_tree.item(item, "values")
            if not values:
                messagebox.showwarning("警告", "选中的规则数据为空")
                return
            
            # 创建编辑对话框
            edit_window = tk.Toplevel(self.root)
            edit_window.title("编辑申报金额规则")
            edit_window.geometry("400x300")
            edit_window.resizable(False, False)
            
            # 设置对话框居中
            edit_window.update_idletasks()
            x = (edit_window.winfo_screenwidth() - edit_window.winfo_reqwidth()) // 2
            y = (edit_window.winfo_screenheight() - edit_window.winfo_reqheight()) // 2
            edit_window.geometry(f"+{x}+{y}")
            
            # 创建输入框
            frame = ttk.Frame(edit_window, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # 国家名称
            ttk.Label(frame, text="国家名称:").grid(row=0, column=0, sticky=tk.W, pady=5)
            country_entry = ttk.Entry(frame, width=30)
            country_entry.grid(row=0, column=1, pady=5)
            country_entry.insert(0, values[0])
            
            # 申报比例
            ttk.Label(frame, text="申报比例:").grid(row=1, column=0, sticky=tk.W, pady=5)
            ratio_entry = ttk.Entry(frame, width=30)
            ratio_entry.grid(row=1, column=1, pady=5)
            ratio_entry.insert(0, str(values[1]))
            
            # 最高申报金额
            ttk.Label(frame, text="最高申报金额:").grid(row=2, column=0, sticky=tk.W, pady=5)
            max_amount_entry = ttk.Entry(frame, width=30)
            max_amount_entry.grid(row=2, column=1, pady=5)
            max_amount_entry.insert(0, str(values[2]))
            
            # 保存按钮
            def save_edit():
                try:
                    old_country = values[0]
                    new_country = country_entry.get().strip()
                    declaration_ratio = float(ratio_entry.get().strip())
                    max_declaration_amount = float(max_amount_entry.get().strip())
                    
                    if not new_country:
                        messagebox.showwarning("警告", "国家名称不能为空")
                        return
                    
                    # 更新规则
                    self.processor.update_declaration_amount_rule(old_country, new_country, declaration_ratio, max_declaration_amount)
                    
                    # 刷新表格
                    self.refresh_declaration_amount_rules()
                    
                    # 关闭对话框
                    edit_window.destroy()
                    
                    messagebox.showinfo("成功", "申报金额规则编辑成功")
                except ValueError:
                    messagebox.showerror("错误", "申报比例和最高申报金额必须是数字")
                except Exception as e:
                    messagebox.showerror("错误", f"编辑申报金额规则失败：{e}")
            
            # 取消按钮
            def cancel_edit():
                edit_window.destroy()
            
            # 创建按钮框架
            button_frame = ttk.Frame(frame)
            button_frame.grid(row=3, column=0, columnspan=2, pady=20)
            
            ttk.Button(button_frame, text="保存", command=save_edit).pack(side=tk.LEFT, padx=10)
            ttk.Button(button_frame, text="取消", command=cancel_edit).pack(side=tk.LEFT, padx=10)
            
            # 设置焦点
            country_entry.focus()
            
        except Exception as e:
            messagebox.showerror("错误", f"编辑申报金额规则失败：{e}")
    
    def delete_declaration_amount_rule(self):
        """删除申报金额规则"""
        try:
            # 获取选中的规则
            selection = self.declaration_amount_tree.selection()
            if not selection:
                messagebox.showwarning("警告", "请选择要删除的申报金额规则")
                return
            
            # 获取选中规则的数据
            item = selection[0]
            values = self.declaration_amount_tree.item(item, "values")
            if not values:
                messagebox.showwarning("警告", "选中的规则数据为空")
                return
            
            country_name = values[0]
            
            # 确认删除
            if messagebox.askyesno("确认", f"确定要删除{country_name}的申报金额规则吗？"):
                # 删除规则
                self.processor.delete_declaration_amount_rule(country_name)
                
                # 刷新表格
                self.refresh_declaration_amount_rules()
                
                messagebox.showinfo("成功", "申报金额规则删除成功")
        except Exception as e:
            messagebox.showerror("错误", f"删除申报金额规则失败：{e}")
    
    def refresh_shop_company_data(self):
        """刷新店铺对应公司数据表格"""
        try:
            self.status_var.set("正在刷新店铺数据...")
            
            # 清空现有表格
            for item in self.shop_company_tree.get_children():
                self.shop_company_tree.delete(item)
            
            # 加载店铺对应公司数据
            shop_company_data = self.processor.get_shop_company_data()
            
            # 插入店铺对应公司数据
            for data in shop_company_data:
                self.shop_company_tree.insert("", tk.END, values=(data["shop_name"], data["company_name"]))
            
            self.status_var.set("店铺数据刷新完成")
        except Exception as e:
            self.status_var.set(f"刷新店铺数据失败：{e}")
            messagebox.showerror("错误", f"刷新店铺数据失败：{e}")
    
    def add_shop_company(self):
        """添加店铺对应公司"""
        # 创建添加店铺对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("添加店铺对应公司")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        
        # 设置对话框居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_reqwidth()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_reqheight()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 创建表单框架
        form_frame = ttk.Frame(dialog, padding="10")
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建店铺名称字段
        ttk.Label(form_frame, text="店铺名称：").grid(row=0, column=0, padx=10, pady=15, sticky=tk.E)
        shop_name_entry = ttk.Entry(form_frame, width=30)
        shop_name_entry.grid(row=0, column=1, padx=10, pady=15, sticky=tk.W)
        
        # 创建所属公司字段
        ttk.Label(form_frame, text="所属公司：").grid(row=1, column=0, padx=10, pady=15, sticky=tk.E)
        company_name_entry = ttk.Entry(form_frame, width=30)
        company_name_entry.grid(row=1, column=1, padx=10, pady=15, sticky=tk.W)
        
        # 保存按钮回调
        def save_shop_company():
            # 获取表单数据
            shop_name = shop_name_entry.get().strip()
            company_name = company_name_entry.get().strip()
            
            # 验证必填字段
            if not shop_name or not company_name:
                messagebox.showwarning("警告", "请填写所有必填字段")
                return
            
            # 调用处理器保存店铺数据
            try:
                self.processor.add_shop_company(shop_name, company_name)
                self.refresh_shop_company_data()
                dialog.destroy()
                messagebox.showinfo("成功", "店铺对应公司添加成功")
            except Exception as e:
                messagebox.showerror("错误", f"添加店铺对应公司失败：{e}")
        
        # 创建按钮框架
        button_frame = ttk.Frame(dialog, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 创建保存和取消按钮
        ttk.Button(button_frame, text="保存", command=save_shop_company).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
    
    def edit_shop_company(self):
        """编辑店铺对应公司"""
        # 获取选中项
        selection = self.shop_company_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请选择要编辑的店铺")
            return
        
        # 获取选中项数据
        item = selection[0]
        values = self.shop_company_tree.item(item, "values")
        shop_name = values[0]
        company_name = values[1]
        
        # 创建编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑店铺对应公司")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        
        # 设置对话框居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_reqwidth()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_reqheight()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 创建输入框
        ttk.Label(dialog, text="店铺名称：").grid(row=0, column=0, padx=10, pady=10, sticky=tk.E)
        shop_entry = ttk.Entry(dialog, width=30)
        shop_entry.grid(row=0, column=1, padx=10, pady=10)
        shop_entry.insert(0, shop_name)
        # 店铺名称现在可以编辑
        
        ttk.Label(dialog, text="所属公司：").grid(row=1, column=0, padx=10, pady=10, sticky=tk.E)
        company_entry = ttk.Entry(dialog, width=30)
        company_entry.grid(row=1, column=1, padx=10, pady=10)
        company_entry.insert(0, company_name)
        
        # 保存按钮回调
        def save_edit():
            new_shop_name = shop_entry.get().strip()
            new_company_name = company_entry.get().strip()
            
            if not new_shop_name:
                messagebox.showwarning("警告", "请填写店铺名称")
                return
            
            if not new_company_name:
                messagebox.showwarning("警告", "请填写所属公司")
                return
            
            # 更新表格数据
            self.shop_company_tree.item(item, values=(new_shop_name, new_company_name))
            
            # 更新处理器中的数据
            shop_company_data = self.processor.get_shop_company_data()
            
            # 如果店铺名称被修改，需要删除旧数据并添加新数据
            if new_shop_name != shop_name:
                # 删除旧数据
                for i, data in enumerate(shop_company_data):
                    if data["shop_name"] == shop_name:
                        del shop_company_data[i]
                        break
                # 添加新数据
                shop_company_data.append({
                    "shop_name": new_shop_name,
                    "company_name": new_company_name
                })
            else:
                # 只更新所属公司
                for data in shop_company_data:
                    if data["shop_name"] == new_shop_name:
                        data["company_name"] = new_company_name
                        break
            
            # 关闭对话框
            dialog.destroy()
            messagebox.showinfo("成功", "店铺对应公司编辑成功")
        
        # 创建按钮
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=2, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="保存", command=save_edit).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
    
    def save_shop_company_data(self):
        """保存店铺对应公司数据"""
        try:
            self.status_var.set("正在保存店铺数据...")
            # 这里可以添加保存店铺对应公司数据的逻辑
            self.status_var.set("店铺数据保存成功")
            messagebox.showinfo("成功", "店铺数据保存成功")
        except Exception as e:
            self.status_var.set(f"保存店铺数据失败：{e}")
            messagebox.showerror("错误", f"保存店铺数据失败：{e}")
    
    def generate_customs_data(self):
        """生成报关单数据"""
        try:
            self.status_var.set("正在生成报关单数据...")
            
            # 生成报关单数据
            self.customs_data = self.processor.generate_declaration_data()
            
            # 刷新表格显示
            self.refresh_customs_data()
            
            self.status_var.set("报关单数据生成成功")
            messagebox.showinfo("成功", f"成功生成 {len(self.customs_data)} 条报关单数据")
        except Exception as e:
            self.status_var.set(f"生成报关单数据失败：{e}")
            messagebox.showerror("错误", f"生成报关单数据失败：{e}")
    
    def refresh_customs_data(self):
        """刷新报关单数据表格"""
        try:
            self.status_var.set("正在刷新报关单数据...")
            
            # 初始化搜索字段下拉框
            columns = self.customs_tree['columns']
            self.customs_search_field['values'] = columns
            if columns:
                self.customs_search_field.current(0)
            
            # 清空现有表格
            for item in self.customs_tree.get_children():
                self.customs_tree.delete(item)
            
            # 插入报关单数据
            for data in self.customs_data:
                # 获取所有列的值
                values = []
                for column in columns:
                    values.append(data.get(column, ""))
                
                # 插入行数据
                item = self.customs_tree.insert("", tk.END, values=values)
                
                # 根据国家有效性标记设置行颜色
                is_country_valid = data.get("is_country_valid", True)
                if not is_country_valid:
                    # 设置红色背景
                    self.customs_tree.item(item, tags=("invalid_country",))
                    # 配置标签样式
                    self.customs_tree.tag_configure("invalid_country", background="#FFCCCC")
            
            self.status_var.set("报关单数据刷新完成")
        except Exception as e:
            self.status_var.set(f"刷新报关单数据失败：{e}")
            messagebox.showerror("错误", f"刷新报关单数据失败：{e}")
    
    def filter_customs_data(self):
        """搜索报关单数据"""
        try:
            self.status_var.set("正在搜索报关单数据...")
            
            # 获取搜索条件
            search_field = self.customs_search_field.get()
            search_value = self.customs_search_value.get().strip()
            
            if not search_field:
                messagebox.showwarning("警告", "请选择搜索字段")
                return
            
            # 清空现有表格
            for item in self.customs_tree.get_children():
                self.customs_tree.delete(item)
            
            # 过滤数据
            filtered_data = []
            for data in self.customs_data:
                if search_field in data:
                    field_value = str(data[search_field])
                    if search_value in field_value:
                        filtered_data.append(data)
            
            # 插入过滤后的数据
            columns = self.customs_tree['columns']
            for data in filtered_data:
                values = []
                for column in columns:
                    values.append(data.get(column, ""))
                self.customs_tree.insert("", tk.END, values=values)
            
            self.status_var.set(f"报关单数据搜索完成，共显示 {len(filtered_data)} 条记录")
        except Exception as e:
            self.status_var.set(f"搜索报关单数据失败：{e}")
            messagebox.showerror("错误", f"搜索报关单数据失败：{e}")
    
    def export_customs_data(self):
        """导出报关单数据"""
        if not self.customs_data:
            messagebox.showwarning("警告", "没有可导出的报关单数据")
            return
        
        try:
            self.status_var.set("正在导出报关单数据...")
            
            # 选择导出文件路径
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
                title="选择导出文件路径"
            )
            
            if not file_path:
                return
            
            # 使用openpyxl导出数据
            from openpyxl import Workbook
            
            # 创建工作簿
            wb = Workbook()
            ws = wb.active
            ws.title = "报关单数据"
            
            # 写入表头
            columns = self.customs_tree['columns']
            ws.append(columns)
            
            # 写入数据
            for data in self.customs_data:
                row = []
                for column in columns:
                    row.append(data.get(column, ""))
                ws.append(row)
            
            # 保存文件
            wb.save(file_path)
            
            self.status_var.set("报关单数据导出成功")
            messagebox.showinfo("成功", f"报关单数据已导出到：{file_path}")
        except Exception as e:
            self.status_var.set(f"导出报关单数据失败：{e}")
            messagebox.showerror("错误", f"导出报关单数据失败：{e}")
    
    def export_customs_data_by_company(self):
        """以公司为单位导出报关单数据"""
        if not self.customs_data:
            messagebox.showwarning("警告", "没有可导出的报关单数据")
            return
        
        try:
            self.status_var.set("正在按公司导出报关单数据...")
            
            # 获取当天日期，格式：YYYY-MM-DD
            import datetime
            today_date = datetime.datetime.now().strftime("%Y-%m-%d")
            
            # 选择导出目录
            export_dir = filedialog.askdirectory(title="选择导出目录")
            if not export_dir:
                return
            
            # 按公司名称分组数据
            company_data_map = {}
            # 确定公司名称字段（可能是"电商企业名称"）
            company_field = "电商企业名称"  # 默认使用电商企业名称作为公司名称
            
            for data in self.customs_data:
                company_name = data.get(company_field, "未知公司")
                if company_name not in company_data_map:
                    company_data_map[company_name] = []
                company_data_map[company_name].append(data)
            
            # 使用openpyxl导出数据
            from openpyxl import Workbook
            
            # 为每个公司导出单独的文件
            exported_files = []
            for company_name, company_data in company_data_map.items():
                # 创建文件名：公司名称+当天日期
                # 替换文件名中可能的非法字符
                safe_company_name = company_name.replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace("\"", "_").replace("<", "_").replace(">", "_").replace("|", "_")
                file_name = f"{safe_company_name}_{today_date}.xlsx"
                file_path = os.path.join(export_dir, file_name)
                
                # 创建工作簿
                wb = Workbook()
                ws = wb.active
                ws.title = "报关单数据"
                
                # 写入表头
                columns = self.customs_tree['columns']
                ws.append(columns)
                
                # 写入数据
                for data in company_data:
                    row = []
                    for column in columns:
                        row.append(data.get(column, ""))
                    ws.append(row)
                
                # 保存文件
                wb.save(file_path)
                exported_files.append(file_path)
            
            self.status_var.set(f"按公司导出报关单数据成功，共导出 {len(exported_files)} 个文件")
            messagebox.showinfo("成功", f"已按公司成功导出 {len(exported_files)} 个文件到：{export_dir}\n\n导出的文件列表：\n" + "\n".join([os.path.basename(f) for f in exported_files[:5]]) + ("\n..." if len(exported_files) > 5 else ""))
        except Exception as e:
            self.status_var.set(f"按公司导出报关单数据失败：{e}")
            messagebox.showerror("错误", f"按公司导出报关单数据失败：{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelManagerGUI(root)
    root.mainloop()
