import tkinter as tk
from gui_base import ExcelManagerGUI


if __name__ == "__main__":
    # 创建主窗口
    root = tk.Tk()
    # 初始化应用
    app = ExcelManagerGUI(root)
    # 启动主循环
    root.mainloop()