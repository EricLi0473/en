import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import subprocess
import os

class DictationGUI:
    def __init__(self, root):
        self.root = root
        root.title("Dictation PDF Generator with Wrongbook")

        # Excel 文件选择
        tk.Label(root, text="Excel 文件:").grid(row=0, column=0, sticky='e')
        self.excel_path_var = tk.StringVar()
        tk.Entry(root, textvariable=self.excel_path_var, width=50).grid(row=0, column=1)
        tk.Button(root, text="浏览", command=self.browse_excel).grid(row=0, column=2)

        # Lists 输入
        tk.Label(root, text="Lists (逗号分隔):").grid(row=1, column=0, sticky='e')
        self.lists_var = tk.StringVar()
        tk.Entry(root, textvariable=self.lists_var, width=50).grid(row=1, column=1)

        # 生成模式
        tk.Label(root, text="模式:").grid(row=2, column=0, sticky='e')
        self.mode_var = tk.StringVar(value='full')
        tk.OptionMenu(root, self.mode_var, 'full', 'sample').grid(row=2, column=1, sticky='w')

        # 抽查数量
        tk.Label(root, text="抽查数量 (sample 模式下):").grid(row=3, column=0, sticky='e')
        self.count_var = tk.StringVar(value='30')
        tk.Entry(root, textvariable=self.count_var, width=10).grid(row=3, column=1, sticky='w')

        # 包含错题本内容
        self.include_wb_var = tk.BooleanVar()
        tk.Checkbutton(root, text="包含错题本内容", variable=self.include_wb_var).grid(row=4, column=1, sticky='w')

        # 生成按钮
        tk.Button(root, text="生成 PDF", command=self.generate_pdf).grid(row=5, column=1, pady=10)

        # 错题本管理区
        tk.Label(root, text="错题本管理:").grid(row=6, column=0, sticky='e')
        frame_wb = tk.Frame(root)
        frame_wb.grid(row=6, column=1, columnspan=2, sticky='w')

        tk.Button(frame_wb, text="添加条目", command=self.wb_add).grid(row=0, column=0, padx=5)
        tk.Button(frame_wb, text="删除条目", command=self.wb_remove).grid(row=0, column=1, padx=5)
        tk.Button(frame_wb, text="导出错题本 PDF", command=self.wb_output).grid(row=0, column=2, padx=5)

        # 输出信息显示
        self.output_text = tk.Text(root, height=15, width=80)
        self.output_text.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

        # 错题本文件名，默认
        self.wb_file = 'wrongbook.txt'

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if path:
            self.excel_path_var.set(path)

    def run_cmd(self, cmd):
        self.output_text.insert(tk.END, f"运行命令: {' '.join(cmd)}\n")
        self.output_text.see(tk.END)
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            self.output_text.insert(tk.END, result.stdout + "\n")
        except subprocess.CalledProcessError as e:
            self.output_text.insert(tk.END, "错误:\n" + e.stderr + "\n")
        self.output_text.see(tk.END)

    def generate_pdf(self):
        excel_path = self.excel_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("错误", "请选择有效的 Excel 文件")
            return
        lists = self.lists_var.get()
        if not lists:
            messagebox.showerror("错误", "请输入有效的 List 号")
            return
        mode = self.mode_var.get()
        count = self.count_var.get()
        try:
            count_num = int(count)
        except:
            count_num = 0

        cmd = [
            "python", "dictation.py",
            "--excel", excel_path,
            "generate",
            "--mode", mode,
            "--lists", lists,
            "--output", "output"
        ]
        if mode == 'sample':
            if count_num <= 0:
                messagebox.showerror("错误", "抽查模式需要输入有效的抽查数量")
                return
            cmd.extend(["--count", str(count_num)])

        if self.include_wb_var.get():
            cmd.append("--include-wb")

        self.run_cmd(cmd)
        self.output_text.insert(tk.END, "生成完成！\n")

    def wb_add(self):
        excel_path = self.excel_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("错误", "请选择有效的 Excel 文件（添加错题本时需要）")
            return
        # 交互式多条输入，直到空行结束
        self.output_text.insert(tk.END, "请输入要添加的错题编号（格式：ListIndex-WordIndex），输入空行结束：\n")
        self.output_text.see(tk.END)
        while True:
            entry = simpledialog.askstring("添加错题本", "输入错题编号（如10-1），空行结束：")
            if not entry:
                break
            # 调用dictation.py wb add命令，用管道输入
            cmd = [
                "python", "dictation.py",
                "--excel", excel_path,
                "wb", "add",
                "--wb-file", self.wb_file
            ]
            # 因为原脚本交互式，需要我们用expect或者自己实现，简单起见用subprocess输入
            # 这里先用简易实现，提示用户使用命令行添加更灵活
            # 所以这里直接写入到错题本文件
            with open(self.wb_file, 'a', encoding='utf-8') as f:
                f.write(entry.strip() + '\n')
            self.output_text.insert(tk.END, f"已添加：{entry.strip()}\n")
            self.output_text.see(tk.END)
        self.output_text.insert(tk.END, "添加完成\n")

    def wb_remove(self):
        excel_path = self.excel_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("错误", "请选择有效的 Excel 文件（删除错题本时需要）")
            return
        self.output_text.insert(tk.END, "请输入要删除的错题编号（格式：ListIndex-WordIndex），输入空行结束：\n")
        self.output_text.see(tk.END)
        while True:
            entry = simpledialog.askstring("删除错题本", "输入错题编号（如10-1），空行结束：")
            if not entry:
                break
            # 读取文件，删除对应条目，再写回
            if not os.path.exists(self.wb_file):
                self.output_text.insert(tk.END, "错题本文件不存在，无需删除\n")
                break
            with open(self.wb_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            lines_new = [line.strip() for line in lines if line.strip() != entry.strip()]
            with open(self.wb_file, 'w', encoding='utf-8') as f:
                for l in lines_new:
                    f.write(l + '\n')
            self.output_text.insert(tk.END, f"已删除（如果存在）：{entry.strip()}\n")
            self.output_text.see(tk.END)
        self.output_text.insert(tk.END, "删除完成\n")

    def wb_output(self):
        excel_path = self.excel_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("错误", "请选择有效的 Excel 文件（导出错题本时需要）")
            return
        if not os.path.exists(self.wb_file):
            messagebox.showinfo("提示", "错题本为空，无法导出")
            return
        cmd = [
            "python", "dictation.py",
            "--excel", excel_path,
            "--wb-file", self.wb_file,
            "wb", "output",
            "--output", "output"
        ]
        self.run_cmd(cmd)
        self.output_text.insert(tk.END, "错题本导出完成！\n")

if __name__ == '__main__':
    root = tk.Tk()
    app = DictationGUI(root)
    root.mainloop()
