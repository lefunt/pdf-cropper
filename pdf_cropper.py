import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import fitz
from PIL import Image
import numpy as np
import os
import ctypes
import subprocess  # 用于打开文件浏览器

def auto_crop_pdf(input_pdf_path, output_pdf_path, border_width=5):
    """
    自动裁剪 PDF 文件中的图表，去除周围空白，并忽略指定宽度的边框。

    Args:
        input_pdf_path: 输入 PDF 文件的路径。
        output_pdf_path: 输出 PDF 文件的路径。
        border_width: 需要忽略的边框宽度，默认为 5 像素。
    """
    pdf_document = fitz.open(input_pdf_path)
    output_pdf = fitz.open()

    for page_number in range(pdf_document.page_count):
        page = pdf_document[page_number]
        pix = page.get_pixmap()
        
        # 将 PDF 页面转换为 PIL 图像
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # 转换为灰度图
        image = image.convert("L")
        image_array = np.array(image)

        # 获取图像尺寸
        height, width = image_array.shape

        # 初始化裁剪边界
        left, top, right, bottom = width, height, 0, 0

        # 判断像素是否位于边框
        def is_border_pixel(x, y):
            return x < border_width or y < border_width or x >= width - border_width or y >= height - border_width
        
        # 寻找所有非白色像素，并且忽略边框
        for y in range(height):
            for x in range(width):
                if not is_border_pixel(x, y) and image_array[y, x] < 255:
                    left = min(left, x)
                    right = max(right, x)
                    top = min(top, y)
                    bottom = max(bottom, y)
                    
        # 如果整个页面都是白色或只有边框，则不裁剪
        if left == width and top == height and right == 0 and bottom == 0:
            new_page = output_pdf.new_page(width=pix.width, height=pix.height)
            new_page.show_pdf_page(new_page.rect, pdf_document, page_number)
            continue
        
        # 创建裁剪区域的矩形
        crop_rect = fitz.Rect(left, top, right + 1, bottom + 1)  # 注意加一操作

        # 创建新的 PDF 页面
        new_page = output_pdf.new_page(width=crop_rect.width, height=crop_rect.height)

        # 从原始页面提取并显示内容
        new_page.show_pdf_page(new_page.rect, pdf_document, page_number, clip=crop_rect)

    # 保存输出 PDF
    output_pdf.save(output_pdf_path)
    pdf_document.close()
    output_pdf.close()

def select_pdf_files():
    """打开文件对话框，选择多个PDF文件."""
    file_paths = filedialog.askopenfilenames(title="选择 PDF 文件", filetypes=[("PDF files", "*.pdf")])
    if file_paths:
        input_files_listbox.delete(0, tk.END)
        for file_path in file_paths:
            input_files_listbox.insert(tk.END, file_path)
        status_label.config(text=f"已选择 {len(file_paths)} 个文件")

def select_output_folder():
    """打开文件夹选择对话框，选择输出文件夹."""
    folder_path = filedialog.askdirectory(title="选择输出文件夹")
    if folder_path:
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder_path)
        status_label.config(text="输出文件夹已选择：" + folder_path)

def process_pdf_files():
    """处理选择的多个PDF文件."""
    file_paths = input_files_listbox.get(0, tk.END)
    border_width_str = border_width_entry.get()
    output_folder = output_folder_entry.get()

    if not file_paths:
         status_label.config(text="错误: 请选择 PDF 文件")
         return
    
    try:
         border_width = int(border_width_str)
    except ValueError:
         status_label.config(text="错误: 边框宽度必须是整数")
         return
    
    if not output_folder:
        output_folder = "output"  # 设置默认输出文件夹为 output
        os.makedirs(output_folder, exist_ok=True) # 如果文件夹不存在，创建它

    status_label.config(text="正在处理，请稍候...")

    for input_pdf_path in file_paths:
         file_name = os.path.basename(input_pdf_path)
         output_file_name = os.path.splitext(file_name)[0] + "_cropped.pdf"
         output_pdf_path = os.path.join(output_folder, output_file_name)
         try:
             auto_crop_pdf(input_pdf_path, output_pdf_path, border_width)
         except Exception as e:
             messagebox.showerror("错误", "处理 PDF 文件时发生错误: " + str(e))
             status_label.config(text="发生错误，停止处理："+ str(e))
             return

    status_label.config(text="处理完成")
    messagebox.showinfo("完成", "PDF 文件裁剪成功")
    
    # 打开输出文件夹
    try:
        if os.name == 'nt':  # windows
             subprocess.Popen(['explorer', os.path.abspath(output_folder)])
        else: # 其他系统
             subprocess.Popen(['open', os.path.abspath(output_folder)])  
    except Exception as e:
            messagebox.showerror("错误", "无法打开输出文件夹：" + str(e))
            status_label.config(text = "无法打开输出文件夹")

# 启用 DPI 感知
if os.name == 'nt': # 仅在 Windows 上启用 DPI 感知
   try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
   except Exception as e:
        print("Failed to set DPI awareness")
    
# 创建主窗口
window = tk.Tk()
window.title("PDF 自动裁剪工具")

# 获取屏幕的 DPI
screen_dpi = window.winfo_fpixels('1i')

# 计算字体大小的缩放比例，并设置最大值
base_font_size = 10 # 设置基准字体大小
font_scale = screen_dpi / 96
max_font_size = 16  # 设置字体大小的最大值
scaled_font_size = min(int(base_font_size * font_scale), max_font_size)

# 根据DPI调整窗口大小
base_width = 700
base_height = 350
window_width = int(base_width * screen_dpi / 96)
window_height = int(base_height * screen_dpi / 96)
window.geometry(f"{window_width}x{window_height}")

default_font = ("Arial", scaled_font_size) # 使用计算后的字体大小

# 使用 ttk 的主题样式
style = ttk.Style(window)
style.theme_use('clam')
style.configure('TLabel', font=default_font)
style.configure('TButton', font=default_font)
style.configure('TEntry', font=default_font)
style.configure('TLabelFrame', font=default_font)
style.configure("Listbox", font=default_font)

# PDF 文件选择框架
input_frame = ttk.LabelFrame(window, text="选择 PDF 文件")
input_frame.pack(padx=20, pady=10, fill="x")

# PDF 文件列表框
input_files_listbox = tk.Listbox(input_frame,  height=5)
input_files_listbox.pack(side=tk.LEFT, padx=5, pady=5, fill="x", expand=True)

# PDF 文件选择按钮
select_file_button = ttk.Button(input_frame, text="选择文件", command=select_pdf_files)
select_file_button.pack(side=tk.LEFT, padx=5, pady=5)

# 输出文件夹框架
output_frame = ttk.LabelFrame(window, text="选择输出文件夹(默认为output文件夹)")
output_frame.pack(padx=20, pady=10, fill="x")

# 输出文件夹路径输入框
output_folder_entry = ttk.Entry(output_frame)
output_folder_entry.pack(side=tk.LEFT, padx=5, pady=5, fill = "x", expand=True)

# 输出文件夹选择按钮
select_output_button = ttk.Button(output_frame, text="选择文件夹", command=select_output_folder)
select_output_button.pack(side=tk.LEFT, padx=5, pady=5)

# 边框宽度框架
border_frame = ttk.LabelFrame(window, text="设置边框宽度")
border_frame.pack(padx=20, pady=10, fill="x")

# 边框宽度输入框
border_width_entry = ttk.Entry(border_frame, width=10)
border_width_entry.insert(0, "5")  # 默认值
border_width_entry.pack(side=tk.LEFT, padx=5, pady=5)

# 处理按钮
process_button = ttk.Button(window, text="开始裁剪", command=process_pdf_files)
process_button.pack(pady=10)

# 状态标签
status_label = ttk.Label(window, text="等待操作...")
status_label.pack(pady=10)

# 运行 GUI 窗口
window.mainloop()