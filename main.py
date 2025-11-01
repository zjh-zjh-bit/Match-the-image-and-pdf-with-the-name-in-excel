# ========== 强制的DLL修复 ==========
import os
import sys

# 强制设置DLL搜索路径
if hasattr(sys, '_MEIPASS'):
    # 打包后的情况
    base_path = sys._MEIPASS
    os.environ['PATH'] = base_path + os.pathsep + os.environ['PATH']
    # 添加必要的子目录到PATH
    for subdir in ['', 'poppler\\Library\\bin', 'tessdata']:
        path_to_add = os.path.join(base_path, subdir)
        if os.path.exists(path_to_add):
            os.environ['PATH'] = path_to_add + os.pathsep + os.environ['PATH']
else:
    # 开发环境
    base_path = os.path.dirname(sys.executable)
    os.environ['PATH'] = base_path + os.pathsep + os.environ['PATH']

# 添加当前工作目录
os.environ['PATH'] = os.path.abspath('.') + os.pathsep + os.environ['PATH']

# 强制预加载可能缺失的DLL
try:
    if hasattr(os, 'add_dll_directory'):
        dll_dirs = [base_path]
        if hasattr(sys, '_MEIPASS'):
            dll_dirs.append(sys._MEIPASS)
        for dll_dir in dll_dirs:
            if os.path.exists(dll_dir):
                os.add_dll_directory(dll_dir)
                # 添加子目录
                for subdir in ['poppler\\Library\\bin', 'tessdata']:
                    subdir_path = os.path.join(dll_dir, subdir)
                    if os.path.exists(subdir_path):
                        os.add_dll_directory(subdir_path)
except Exception:
    pass  # 如果失败继续运行
# ========== DLL修复结束 ==========

import os
import sys
import shutil
import pytesseract
import pandas as pd
import re
import tempfile
from PIL import Image, ImageEnhance
from pdf2image import convert_from_path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# ========== 资源路径函数 ==========
def resource_path(relative_path):
    """获取资源的绝对路径，用于PyInstaller打包后定位文件"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ========== 依赖检查和初始化 ==========
def setup_environment():
    """设置运行环境，确保所有依赖可用"""
    # 设置Tesseract路径 - 适配 tessdata/tesseract.exe 结构
    tesseract_path = resource_path("tessdata")
    if os.path.exists(tesseract_path):
        tesseract_exe = os.path.join(tesseract_path, "tesseract.exe")
        if os.path.exists(tesseract_exe):
            pytesseract.pytesseract.tesseract_cmd = tesseract_exe
            print(f"Tesseract路径设置为: {tesseract_exe}")
    
    # 设置Poppler路径 - 适配 poppler/Library/bin 结构
    poppler_path = resource_path("poppler")
    if os.path.exists(poppler_path):
        poppler_bin = os.path.join(poppler_path, "Library", "bin")
        if os.path.exists(poppler_bin):
            # 确保poppler在PATH中
            if poppler_bin not in os.environ['PATH']:
                os.environ['PATH'] = poppler_bin + os.pathsep + os.environ['PATH']
            print(f"Poppler路径设置为: {poppler_bin}")

def check_dependencies():
    """检查关键依赖是否可用"""
    missing_deps = []
    
    # 检查Tesseract
    try:
        pytesseract.get_tesseract_version()
    except:
        missing_deps.append("Tesseract OCR")
    
    # 检查Poppler
    poppler_path = resource_path("poppler")
    if not os.path.exists(os.path.join(poppler_path, "Library", "bin", "pdftoppm.exe")):
        missing_deps.append("Poppler PDF工具")
    
    if missing_deps:
        messagebox.showerror("依赖缺失", f"以下依赖缺失：{', '.join(missing_deps)}\n\n请确保程序完整安装。")
        return False
    return True

# ========== 初始化环境 ==========
setup_environment()

class CertificateProcessor:
    def __init__(self):
        # 设置poppler路径（使用打包的版本）
        poppler_base = resource_path("poppler")
        self.poppler_path = os.path.join(poppler_base, "Library", "bin")
        
        # 确保环境变量包含poppler路径
        if self.poppler_path not in os.environ['PATH']:
            os.environ['PATH'] = self.poppler_path + os.pathsep + os.environ['PATH']
        
    def load_target_names(self, table_path, name_column, year_column, target_years):
        try:
            df = pd.read_excel(table_path)
            if year_column not in df.columns:
                raise ValueError(f"未找到'{year_column}'列，请检查列名")
            df_filtered = df[df[year_column].isin(target_years)]
            names = df_filtered[name_column].dropna().unique().tolist()
            return names, None
        except Exception as e:
            return [], f"表格处理失败：{e}"
    
    def extract_names_from_filename(self, filename):
        """从文件名中提取可能的姓名"""
        filename_without_ext = os.path.splitext(filename)[0]
        
        patterns = [
            r'[（(]([\u4e00-\u9fa5]{2,4})[^）)]*[）)]',
            r'[\u4e00-\u9fa5]{2,4}',
            r'[\u4e00-\u9fa5]{2,4}[A-Za-z0-9+]+'
        ]
        
        found_names = []
        for pattern in patterns:
            matches = re.findall(pattern, filename_without_ext)
            for match in matches:
                if re.match(r'^[\u4e00-\u9fa5]{2,4}$', match):
                    found_names.append(match)
                else:
                    chinese_part = re.findall(r'[\u4e00-\u9fa5]{2,4}', match)
                    if chinese_part:
                        found_names.extend(chinese_part)
        
        return list(set(found_names))
    
    def image_contains_name(self, image_path, target_names):
        try:
            # 文件名匹配
            filename = os.path.basename(image_path)
            filename_without_suffix = os.path.splitext(filename)[0]
            
            extracted_names = self.extract_names_from_filename(filename_without_suffix)
            
            for extracted_name in extracted_names:
                for target_name in target_names:
                    clean_target = target_name.strip().replace(" ", "")
                    clean_extracted = extracted_name.strip().replace(" ", "")
                    
                    if clean_extracted == clean_target:
                        return True, target_name, "文件名匹配"
            
            # OCR识别匹配
            image = Image.open(image_path)
            image = image.convert("L")
            image = image.point(lambda x: 0 if x < 140 else 255)
            enhancer = ImageEnhance.Sharpness(image)
            image = enhancer.enhance(1.8)
            
            text = pytesseract.image_to_string(
                image,
                lang="chi_sim+eng",
                config="--psm 11"
            )
            clean_text = text.replace("：", "").replace("（", "").replace("）", "").replace("-", "").replace("_", "").replace(" ", "")
            
            for name in target_names:
                clean_name = name.strip().replace(" ", "")
                pattern = re.compile(rf"(?:^|[^a-zA-Z0-9\u4e00-\u9fa5]){re.escape(clean_name)}(?:$|[^a-zA-Z0-9\u4e00-\u9fa5])", re.IGNORECASE)
                if pattern.search(clean_text):
                    return True, name, "OCR识别匹配"
            
            return False, None, None
        
        except Exception as e:
            return False, None, f"图片处理失败：{e}"
    
    def pdf_to_images(self, pdf_path):
        """将PDF文件转换为图片列表"""
        try:
            # 检查poppler路径
            if not os.path.exists(self.poppler_path):
                return None, f"Poppler路径不存在：{self.poppler_path}"
                
            # 使用打包的poppler
            images = convert_from_path(
                pdf_path, 
                poppler_path=self.poppler_path,
                dpi=200
            )
            return images, None
        except Exception as e:
            return None, f"PDF转换失败：{e}"
    
    def process_pdf(self, pdf_path, target_names):
        """处理PDF文件，检查是否包含目标姓名"""
        try:
            # 先检查PDF文件名
            filename = os.path.basename(pdf_path)
            filename_without_suffix = os.path.splitext(filename)[0]
            
            extracted_names = self.extract_names_from_filename(filename_without_suffix)
            
            for extracted_name in extracted_names:
                for target_name in target_names:
                    clean_target = target_name.strip().replace(" ", "")
                    clean_extracted = extracted_name.strip().replace(" ", "")
                    
                    if clean_extracted == clean_target:
                        return True, target_name, "PDF文件名匹配"
            
            # 将PDF转换为图片并逐一检查
            images, error = self.pdf_to_images(pdf_path)
            if error:
                return False, None, error
            
            with tempfile.TemporaryDirectory() as temp_dir:
                for i, image in enumerate(images):
                    temp_image_path = os.path.join(temp_dir, f"page_{i+1}.png")
                    image.save(temp_image_path, "PNG")
                    
                    contains_name, matched_name, match_type = self.image_contains_name(temp_image_path, target_names)
                    if contains_name:
                        return True, matched_name, f"PDF内页匹配(第{i+1}页)"
            
            return False, None, None
        
        except Exception as e:
            return False, None, f"PDF处理失败：{e}"
    
    def process_files(self, table_path, root_folder, output_folder, name_column, year_column, target_years, progress_callback=None):
        """主处理函数"""
        try:
            # 检查依赖
            if not check_dependencies():
                return False, "依赖检查失败，无法继续处理"
            
            # 加载目标姓名
            target_names, error = self.load_target_names(table_path, name_column, year_column, target_years)
            if error:
                return False, error
            
            if not target_names:
                return False, "未找到匹配的姓名"
            
            os.makedirs(output_folder, exist_ok=True)
            
            matched_count = 0
            file_match_count = 0
            ocr_match_count = 0
            pdf_match_count = 0
            processed_count = 0
            total_files = 0
            
            # 统计总文件数
            for root, dirs, files in os.walk(root_folder):
                for filename in files:
                    if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.pdf')):
                        total_files += 1
            
            # 处理文件
            for root, dirs, files in os.walk(root_folder):
                for filename in files:
                    file_path = os.path.join(root, filename)
                    lower_filename = filename.lower()
                    
                    if progress_callback:
                        progress_callback(processed_count, total_files, f"正在处理: {filename}")
                    
                    # 处理图片文件
                    if lower_filename.endswith(('.jpg', '.jpeg', '.png', '.gif')):
                        contains_name, matched_name, match_type = self.image_contains_name(file_path, target_names)
                        if contains_name:
                            if match_type == "文件名匹配":
                                file_match_count += 1
                            else:
                                ocr_match_count += 1
                            
                            relative_path = os.path.relpath(file_path, root_folder)
                            output_dir = os.path.join(output_folder, os.path.dirname(relative_path))
                            os.makedirs(output_dir, exist_ok=True)
                            output_path = os.path.join(output_dir, filename)
                            shutil.copy2(file_path, output_path)
                            matched_count += 1
                    
                    # 处理PDF文件
                    elif lower_filename.endswith('.pdf'):
                        contains_name, matched_name, match_type = self.process_pdf(file_path, target_names)
                        if contains_name:
                            pdf_match_count += 1
                            
                            relative_path = os.path.relpath(file_path, root_folder)
                            output_dir = os.path.join(output_folder, os.path.dirname(relative_path))
                            os.makedirs(output_dir, exist_ok=True)
                            output_path = os.path.join(output_dir, filename)
                            shutil.copy2(file_path, output_path)
                            matched_count += 1
                    
                    processed_count += 1
            
            result_message = f"处理完成！共找到 {matched_count} 个匹配文件：\n"
            result_message += f"- 文件名匹配：{file_match_count} 个\n"
            result_message += f"- OCR识别匹配：{ocr_match_count} 个\n"
            result_message += f"- PDF文件匹配：{pdf_match_count} 个"
            
            return True, result_message
        
        except Exception as e:
            return False, f"处理过程中出错：{e}"

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("证书文件智能分类工具")
        self.geometry("600x500")
        self.processor = CertificateProcessor()
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置列权重
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="证书文件智能分类工具", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 表格文件选择
        ttk.Label(main_frame, text="学生信息表格:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.table_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.table_path, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_table).grid(row=1, column=2, padx=(5, 0))
        
        # 图片文件夹选择
        ttk.Label(main_frame, text="证书文件文件夹:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.root_folder = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.root_folder, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_root_folder).grid(row=2, column=2, padx=(5, 0))
        
        # 输出文件夹选择
        ttk.Label(main_frame, text="输出文件夹:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.output_folder = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.output_folder, width=50).grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_output_folder).grid(row=3, column=2, padx=(5, 0))
        
        # 配置参数
        config_frame = ttk.LabelFrame(main_frame, text="配置参数", padding="10")
        config_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        config_frame.columnconfigure(1, weight=1)
        
        ttk.Label(config_frame, text="姓名列名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_column = tk.StringVar(value="姓名")
        ttk.Entry(config_frame, textvariable=self.name_column).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(config_frame, text="筛选列名:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.year_column = tk.StringVar(value="专业")
        ttk.Entry(config_frame, textvariable=self.year_column).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(config_frame, text="筛选条件:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.target_years = tk.StringVar(value="（本）计算机科学与技术")
        ttk.Entry(config_frame, textvariable=self.target_years).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="准备就绪")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        # 开始按钮
        self.start_button = ttk.Button(main_frame, text="开始处理", command=self.start_processing)
        self.start_button.grid(row=7, column=0, columnspan=3, pady=10)
        
        # 结果文本框
        self.result_text = tk.Text(main_frame, height=8, width=70)
        self.result_text.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 设置默认输出文件夹
        default_output = os.path.join(os.path.expanduser("~"), "证书分类结果")
        self.output_folder.set(default_output)
    
    def browse_table(self):
        filename = filedialog.askopenfilename(
            title="选择学生信息表格",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.table_path.set(filename)
    
    def browse_root_folder(self):
        folder = filedialog.askdirectory(title="选择证书文件所在文件夹")
        if folder:
            self.root_folder.set(folder)
    
    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_folder.set(folder)
    
    def update_progress(self, current, total, status):
        if total > 0:
            progress_percent = (current / total) * 100
            self.progress['value'] = progress_percent
        self.status_label.config(text=status)
        self.update_idletasks()
    
    def start_processing(self):
        # 验证输入
        if not self.table_path.get():
            messagebox.showerror("错误", "请选择学生信息表格文件")
            return
        
        if not self.root_folder.get():
            messagebox.showerror("错误", "请选择证书文件文件夹")
            return
        
        if not self.output_folder.get():
            messagebox.showerror("错误", "请选择输出文件夹")
            return
        
        # 禁用开始按钮
        self.start_button.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self.result_text.delete(1.0, tk.END)
        
        # 在新线程中运行处理过程
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
    
    def run_processing(self):
        try:
            # 解析筛选条件（支持多个条件，用逗号分隔）
            target_years_list = [year.strip() for year in self.target_years.get().split(',')]
            
            # 调用处理函数
            success, result = self.processor.process_files(
                table_path=self.table_path.get(),
                root_folder=self.root_folder.get(),
                output_folder=self.output_folder.get(),
                name_column=self.name_column.get(),
                year_column=self.year_column.get(),
                target_years=target_years_list,
                progress_callback=self.update_progress
            )
            
            # 显示结果
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, result)
            
            if success:
                messagebox.showinfo("完成", "处理完成！")
            else:
                messagebox.showerror("错误", result)
        
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出错：{e}")
        
        finally:
            # 重新启用开始按钮
            self.start_button.config(state=tk.NORMAL)
            self.progress['value'] = 0
            self.status_label.config(text="处理完成")

if __name__ == "__main__":
    # 初始化环境检查
    if not check_dependencies():
        sys.exit(1)
        
    app = Application()
    app.mainloop()