#!/usr/bin/env python3
import sys, os, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

try:
    from docx2pdf import convert
    HAS_DEP = True
except ImportError:
    HAS_DEP = False

class App:
    def __init__(self, root):
        self.root = root
        root.title('Word to PDF 转换工具 v1.0')
        root.geometry('600x400')
        self.files = []
        self.build_ui()
    
    def build_ui(self):
        f = tk.Frame(self.root, bg='#1f538d', height=60)
        f.pack(fill='x')
        tk.Label(f, text='Word → PDF 转换工具', font=('Arial',16,'bold'), fg='white', bg='#1f538d').pack(pady=15)
        main = tk.Frame(self.root, padx=20, pady=15)
        main.pack(fill='both', expand=True)
        bf = tk.Frame(main)
        bf.pack(fill='x', pady=5)
        tk.Button(bf, text='添加Word文件', command=self.add_files, bg='#1f538d', fg='white', padx=15).pack(side='left', padx=5)
        tk.Button(bf, text='清空列表', command=self.clear, bg='#d9534f', fg='white', padx=15).pack(side='left', padx=5)
        tk.Button(bf, text='开始转换', command=self.convert, bg='#5cb85c', fg='white', font=('Arial',10,'bold'), padx=20).pack(side='right', padx=5)
        self.lb = tk.Listbox(main, font=('Consolas',10), bg='#f8f9fa', height=12)
        self.lb.pack(fill='both', expand=True, pady=10)
        self.status = tk.Label(main, text='请添加Word文件（支持.doc/.docx）', font=('Arial',10), fg='gray', anchor='w')
        self.status.pack(fill='x')
    
    def add_files(self):
        fs = filedialog.askopenfilenames(title='选择Word文件', filetypes=[('Word文件','*.doc*')])
        for f in fs:
            if f not in self.files:
                self.files.append(f)
                self.lb.insert('end', Path(f).name)
        self.status.config(text=f'已添加 {len(self.files)} 个文件')
    
    def clear(self):
        self.files.clear()
        self.lb.delete(0, 'end')
        self.status.config(text='列表已清空')
    
    def convert(self):
        if not self.files:
            messagebox.showwarning('提示', '请先添加Word文件')
            return
        if not HAS_DEP:
            messagebox.showerror('缺少依赖', '请运行：pip install docx2pdf python-docx')
            return
        out_dir = filedialog.askdirectory(title='选择输出目录')
        if not out_dir: return
        ok = 0
        for f in self.files:
            try:
                out = str(Path(out_dir) / (Path(f).stem + '.pdf'))
                convert(f, out)
                ok += 1
            except Exception as e:
                messagebox.showerror('错误', f'{Path(f).name} 转换失败：{e}')
        messagebox.showinfo('完成', f'成功转换 {ok}/{len(self.files)} 个文件！')
        self.status.config(text=f'✅ 完成：{ok}个文件已转换')

if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
