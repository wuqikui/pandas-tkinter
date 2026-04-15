import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import logging
from datetime import datetime
import os
import traceback

class ExcelDataCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("检查工具")
        self.root.geometry("600x400")

        # 初始化变量
        self.excel_path = tk.StringVar(value="请选择Excel文件:")
        self.selected_area = tk.StringVar(value="全部")
        self.result_text = tk.StringVar()

        # 创建界面组件
        self.create_widgets()
        # 配置日志
        self.setup_logging()

    def setup_logging(self):
        """配置日志记录"""
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"error_{timestamp}.log")

        logging.basicConfig(
            filename=log_file,
            level=logging.ERROR,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def create_widgets(self):
        """创建界面组件"""
        # 文件选择部分
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Entry(file_frame, textvariable=self.excel_path, width=75).pack(side=tk.LEFT, padx=7)
        tk.Button(file_frame, text="选择", command=self.select_file).pack(side=tk.LEFT)

        # 区域选择部分
        area_frame = tk.Frame(self.root)
        area_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(area_frame, text="区域:").pack(side=tk.LEFT)
        area_options = ["全部", "A区", "B区", "C区"]
        ttk.Combobox(area_frame, textvariable=self.selected_area, values=area_options, state="readonly").pack(
            side=tk.LEFT, padx=5)

        # 结果显示
        result_frame = tk.Frame(self.root)
        result_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        tk.Label(result_frame, text="异常数据:").pack(anchor=tk.W)
        result_text = tk.Text(result_frame, wrap=tk.WORD, height=10)
        scrollbar = tk.Scrollbar(result_frame, command=result_text.yview)
        result_text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        result_text.pack(fill=tk.BOTH, expand=True)
        self.result_text_widget = result_text

        # 进度条
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=570, mode='determinate')
        self.progress.pack(pady=10)
        # 检查按钮
        tk.Button(self.root, text="检查", command=self.check_data, height=1).pack(side=tk.RIGHT,padx=10,pady=20)

    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="请选择Excel文件",
            filetypes=[("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.excel_path.set(file_path)

    def check_data(self):
        """检查数据"""
        # 重置界面
        self.result_text_widget.delete(1.0, tk.END)
        self.progress['value'] = 0
        self.root.update()

        # 验证文件
        excel_file = self.excel_path.get()
        if not excel_file:
            messagebox.showerror("错误", "程序发生异常！")
            return

        if not os.path.exists(excel_file):
            messagebox.showerror("错误", "程序发生异常！")
            return

        try:
            # 读取Excel数据
            self.progress['value'] = 20
            self.root.update()

            df = pd.read_excel(excel_file)
            if df.empty:
                messagebox.showinfo("提示", "Excel文件中没有数据！")
                return

            # 检查数据
            self.progress['value'] = 40
            self.root.update()

            selected_area = self.selected_area.get()
            if selected_area != "全部":
                area_code = selected_area  # 提取"A"从"A区"
                df = df[df['区域'] == area_code]

            errors = []
            for index, row in df.iterrows():
                error_msgs = []

                # 检查地号
                if pd.isna(row['地号']) or str(row['地号']).strip() == "":
                    error_msgs.append("地号为空")

                # 检查土地性质
                land_type = str(row['土地性质']).strip()
                if pd.isna(row['土地性质']) or land_type == "" or (land_type not in ["国有", "集体"]):
                    error_msgs.append("土地性质值不合法")

                if error_msgs:
                    errors.append(f"第{index + 2}行数据，{', '.join(error_msgs)}；")

            self.progress['value'] = 80
            self.root.update()

            # 显示结果
            if errors:
                result_text = "\n".join(errors)
                self.result_text_widget.insert(tk.END, result_text)
            else:
                self.result_text_widget.insert(tk.END, "没有发现不符合标准的数据")

            self.progress['value'] = 100
            messagebox.showinfo("提示", "检查完毕！")

        except Exception as e:
            logging.error(traceback.format_exc())
            messagebox.showerror("错误", "程序发生异常！")
            self.progress['value'] = 0


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDataCheckerApp(root)
    root.mainloop()