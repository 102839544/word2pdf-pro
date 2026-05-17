#!/usr/bin/env python3
"""
Word2PDF Pro - 批量 Word 转 PDF 专业工具
功能：
1. 批量 Word 转 PDF
2. 批量重命名
3. 自动去除 Word 备注/批注
4. 收费授权验证
"""

import os
import sys
import json
import hashlib
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from docx import Document
from docx2pdf import convert
import pythoncom
import win32com.client
import threading

# ==================== 授权系统 ====================
LICENSE_FILE = Path.home() / ".word2pdf_pro_license"
APP_ID = "Word2PDF_Pro_v1.0"

def generate_machine_code():
    """生成机器码（基于硬件信息）"""
    import platform
    machine_info = f"{platform.node()}{platform.processor()}"
    return hashlib.md5(machine_info.encode()).hexdigest()[:16].upper()

def verify_license(license_key: str) -> bool:
    """验证授权码"""
    machine_code = generate_machine_code()
    expected = hashlib.md5(f"{APP_ID}{machine_code}".encode()).hexdigest()[:24].upper()
    return license_key.upper() == expected

def save_license(license_key: str):
    """保存授权信息"""
    license_data = {
        "key": license_key,
        "machine": generate_machine_code(),
        "activated_at": str(datetime.datetime.now())
    }
    with open(LICENSE_FILE, "w") as f:
        json.dump(license_data, f)

def load_license() -> bool:
    """加载并验证本地授权"""
    if not LICENSE_FILE.exists():
        return False
    try:
        with open(LICENSE_FILE, "r") as f:
            data = json.load(f)
        return verify_license(data["key"])
    except:
        return False

# ==================== Word 处理功能 ====================
def remove_comments(input_path: str, output_path: str = None) -> str:
    """
    去除 Word 文档中的批注和修订
    """
    if output_path is None:
        output_path = input_path
    
    # 使用 win32com 处理
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        doc_com = word.Documents.Open(input_path)
        
        # 接受所有修订
        if doc_com.Revisions.Count > 0:
            doc_com.Revisions.AcceptAll()
        
        # 删除所有批注
        if doc_com.Comments.Count > 0:
            for i in range(doc_com.Comments.Count, 0, -1):
                doc_com.Comments(i).Delete()
        
        # 保存
        doc_com.SaveAs(output_path)
        doc_com.Close()
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    
    return output_path

def convert_word_to_pdf(input_path: str, output_path: str = None, remove_comments_first: bool = True):
    """将 Word 转换为 PDF"""
    if output_path is None:
        output_path = str(Path(input_path).with_suffix(".pdf"))
    
    # 如果需要去除备注，先处理
    if remove_comments_first:
        temp_path = str(Path(input_path).with_suffix(".temp.docx"))
        remove_comments(input_path, temp_path)
        input_path = temp_path
    
    # 转换
    convert(input_path, output_path)
    
    # 删除临时文件
    if remove_comments_first and os.path.exists(temp_path):
        os.remove(temp_path)
    
    return output_path

# ==================== 批量重命名功能 ====================
def batch_rename(files: list, rule: str, prefix: str = "", suffix: str = "", start_num: int = 1):
    """
    批量重命名文件
    rule: "add_prefix", "add_suffix", "add_number", "replace"
    """
    results = []
    for i, file_path in enumerate(files):
        path_obj = Path(file_path)
        original_name = path_obj.stem
        ext = path_obj.suffix
        
        if rule == "add_prefix":
            new_name = f"{prefix}{original_name}{ext}"
        elif rule == "add_suffix":
            new_name = f"{original_name}{suffix}{ext}"
        elif rule == "add_number":
            new_name = f"{prefix}{start_num + i:03d}{suffix}{ext}"
        elif rule == "replace":
            new_name = f"{prefix}{ext}"
        else:
            new_name = f"{original_name}{ext}"
        
        new_path = path_obj.parent / new_name
        path_obj.rename(new_path)
        results.append(str(new_path))
    
    return results

# ==================== GUI 界面 ====================
class LicenseDialog(ctk.CTkToplevel):
    """授权对话框"""
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Word2PDF Pro - 授权激活")
        self.geometry("500x350")
        self.resizable(False, False)
        
        # 居中显示
        self.transient(parent)
        self.grab_set()
        
        # UI
        self.create_widgets()
    
    def create_widgets(self):
        # 标题
        title = ctk.CTkLabel(self, text="Word2PDF Pro", font=("Arial", 24, "bold"), text_color="#1f538d")
        title.pack(pady=20)
        
        # 说明
        info = ctk.CTkLabel(self, text="请输入授权码以继续使用", font=("Arial", 14))
        info.pack(pady=10)
        
        # 机器码显示
        machine_code = generate_machine_code()
        machine_label = ctk.CTkLabel(self, text=f"机器码: {machine_code}", font=("Courier", 12))
        machine_label.pack(pady=5)
        
        # 授权码输入
        self.key_var = tk.StringVar()
        key_entry = ctk.CTkEntry(self, textvariable=self.key_var, width=300, height=40,
                                  font=("Arial", 14), placeholder_text="输入授权码")
        key_entry.pack(pady=20)
        
        # 激活按钮
        activate_btn = ctk.CTkButton(self, text="激活", command=self.activate, width=200, height=40,
                                       font=("Arial", 14), fg_color="#1f538d", hover_color="#14375e")
        activate_btn.pack(pady=10)
        
        # 购买链接（模拟）
        buy_label = ctk.CTkLabel(self, text="获取授权码请联系：1028395447@qq.com", font=("Arial", 12), text_color="gray")
        buy_label.pack(pady=20)
    
    def activate(self):
        key = self.key_var.get().strip()
        if not key:
            messagebox.showerror("错误", "请输入授权码")
            return
        
        if verify_license(key):
            save_license(key)
            messagebox.showinfo("成功", "激活成功！")
            self.destroy()
        else:
            messagebox.showerror("错误", "授权码无效")

class Word2PDFApp(ctk.CTk):
    """主应用窗口"""
    def __init__(self):
        super().__init__()
        
        # 检查授权
        if not load_license():
            dialog = LicenseDialog(self)
            self.wait_window(dialog)
            if not load_license():
                messagebox.showerror("错误", "需要激活才能使用")
                self.quit()
                return
        
        # 初始化 UI
        self.title("Word2PDF Pro - 批量 Word 转 PDF 专业工具")
        self.geometry("900x700")
        
        # 文件列表
        self.files = []
        
        self.create_widgets()
    
    def create_widgets(self):
        # 标题栏
        header = ctk.CTkFrame(self, height=80, fg_color="#1f538d")
        header.pack(fill="x")
        header.pack_propagate(False)
        
        title = ctk.CTkLabel(header, text="Word2PDF Pro", font=("Arial", 28, "bold"), text_color="white")
        title.pack(side="left", padx=20, pady=20)
        
        subtitle = ctk.CTkLabel(header, text="批量转换 · 智能重命名 · 去除备注", font=("Arial", 14), text_color="white")
        subtitle.pack(side="left", padx=10)
        
        # 主容器
        main_container = ctk.CTkFrame(self)
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 左侧：文件列表
        left_frame = ctk.CTkFrame(main_container)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        file_label = ctk.CTkLabel(left_frame, text="文件列表", font=("Arial", 16, "bold"))
        file_label.pack(pady=10)
        
        # 文件列表框
        self.file_listbox = tk.Listbox(left_frame, width=50, height=25, font=("Arial", 10),
                                        selectmode=tk.EXTENDED, bg="#2b2b2b", fg="white",
                                        selectbackground="#1f538d")
        self.file_listbox.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 按钮区
        btn_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        add_btn = ctk.CTkButton(btn_frame, text="添加文件", command=self.add_files, width=100)
        add_btn.pack(side="left", padx=5)
        
        clear_btn = ctk.CTkButton(btn_frame, text="清空列表", command=self.clear_files, width=100, fg_color="#d9534f")
        clear_btn.pack(side="left", padx=5)
        
        # 右侧：功能面板
        right_frame = ctk.CTkFrame(main_container, width=300)
        right_frame.pack(side="right", fill="both", padx=(10, 0))
        right_frame.pack_propagate(False)
        
        # 转换设置
        convert_label = ctk.CTkLabel(right_frame, text="转换设置", font=("Arial", 16, "bold"))
        convert_label.pack(pady=10)
        
        self.remove_comments_var = tk.BooleanVar(value=True)
        remove_check = ctk.CTkCheckBox(right_frame, text="去除批注和修订", variable=self.remove_comments_var)
        remove_check.pack(pady=10, padx=20, anchor="w")
        
        # 重命名设置
        rename_label = ctk.CTkLabel(right_frame, text="重命名规则", font=("Arial", 16, "bold"))
        rename_label.pack(pady=(20, 10))
        
        self.rename_rule = tk.StringVar(value="none")
        rules = [
            ("不重命名", "none"),
            ("添加前缀", "add_prefix"),
            ("添加后缀", "add_suffix"),
            ("添加序号", "add_number"),
        ]
        for text, value in rules:
            rb = ctk.CTkRadioButton(right_frame, text=text, variable=self.rename_rule, value=value)
            rb.pack(pady=5, padx=20, anchor="w")
        
        # 前缀/后缀输入
        self.prefix_var = tk.StringVar()
        prefix_entry = ctk.CTkEntry(right_frame, textvariable=self.prefix_var, placeholder_text="前缀/后缀/序号前缀")
        prefix_entry.pack(fill="x", padx=20, pady=10)
        
        # 开始转换按钮
        convert_btn = ctk.CTkButton(right_frame, text="开始转换", command=self.start_convert,
                                     height=50, font=("Arial", 16, "bold"), fg_color="#5cb85c",
                                     hover_color="#4cae4c")
        convert_btn.pack(fill="x", padx=20, pady=30)
        
        # 进度条
        self.progress = ctk.CTkProgressBar(right_frame)
        self.progress.pack(fill="x", padx=20, pady=10)
        self.progress.set(0)
        
        # 状态栏
        self.status_label = ctk.CTkLabel(right_frame, text="就绪", font=("Arial", 12), text_color="gray")
        self.status_label.pack(pady=10)
    
    def add_files(self):
        """添加文件"""
        files = filedialog.askopenfilenames(
            title="选择 Word 文件",
            filetypes=[("Word 文件", "*.docx *.doc")]
        )
        if files:
            for f in files:
                if f not in self.files:
                    self.files.append(f)
                    self.file_listbox.insert("end", Path(f).name)
    
    def clear_files(self):
        """清空文件列表"""
        self.files.clear()
        self.file_listbox.delete(0, "end")
    
    def start_convert(self):
        """开始转换"""
        if not self.files:
            messagebox.showwarning("警告", "请先添加文件")
            return
        
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择输出目录")
        if not output_dir:
            return
        
        # 在新线程中执行转换
        thread = threading.Thread(target=self.convert_files, args=(output_dir,))
        thread.daemon = True
        thread.start()
    
    def convert_files(self, output_dir: str):
        """执行批量转换"""
        total = len(self.files)
        
        for i, file_path in enumerate(self.files):
            try:
                self.status_label.configure(text=f"正在处理: {Path(file_path).name}")
                
                # 重命名逻辑
                rule = self.rename_rule.get()
                output_name = Path(file_path).stem
                
                if rule == "add_prefix":
                    output_name = f"{self.prefix_var.get()}{output_name}"
                elif rule == "add_suffix":
                    output_name = f"{output_name}{self.prefix_var.get()}"
                elif rule == "add_number":
                    prefix = self.prefix_var.get() or "file_"
                    output_name = f"{prefix}{i+1:03d}"
                
                output_path = str(Path(output_dir) / f"{output_name}.pdf")
                
                # 转换
                convert_word_to_pdf(file_path, output_path, self.remove_comments_var.get())
                
                # 更新进度
                progress_value = (i + 1) / total
                self.progress.set(progress_value)
                
            except Exception as e:
                messagebox.showerror("错误", f"处理 {file_path} 时出错:\n{str(e)}")
                return
        
        self.status_label.configure(text="转换完成！")
        messagebox.showinfo("完成", f"成功转换 {total} 个文件！")

def main():
    """主函数"""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    app = Word2PDFApp()
    app.mainloop()

if __name__ == "__main__":
    main()
