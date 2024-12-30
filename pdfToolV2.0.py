import fitz  # PyMuPDF
from PIL import Image
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import ctypes
import subprocess  # 用于打开文件浏览器
import pythoncom
import win32com.client


# 判断文件类型并转换为 PDF
def convert_to_pdf(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()
    output_pdf = os.path.splitext(file_path)[0] + ".pdf"  # 设置默认输出 PDF 文件名

    if file_extension == ".pdf":
        return file_path  # 如果已经是 PDF，直接返回文件路径

    elif file_extension in [".jpg", ".jpeg", ".png", ".bmp"]:
        # 图片文件转换为 PDF
        images_to_pdf([file_path], output_pdf)
        return output_pdf

    elif file_extension in [".vsd", ".vsdx"]:
        # Visio 文件转换为 PDF
        visio_to_pdf(file_path, output_pdf)
        return output_pdf

    elif file_extension == ".docx":
        # Word 文件转换为 PDF
        word_to_pdf(file_path, output_pdf)
        return output_pdf

    else:
        messagebox.showerror("错误", f"不支持的文件类型: {file_extension}")
        return None


# 合并 PDF 文件的函数
def merge_pdfs(pdf_list, output_pdf):
    merged_pdf = fitz.open()
    for pdf in pdf_list:
        current_pdf = fitz.open(pdf)
        merged_pdf.insert_pdf(current_pdf)
    merged_pdf.save(output_pdf)
    messagebox.showinfo("成功", f"PDFs已成功合并并保存为 {output_pdf}")


# 图片转换为 PDF 的函数
def images_to_pdf(image_list, output_pdf):
    images = []
    for img in image_list:
        try:
            img_obj = Image.open(img)
            img_obj = img_obj.convert("RGB")  # 转换为RGB模式
            images.append(img_obj)
        except Exception as e:
            messagebox.showerror("错误", f"图片文件转换失败: {img}\n{str(e)}")
            return

    if images:
        # 保存为单页PDF
        images[0].save(output_pdf, save_all=True, append_images=images[1:])
        messagebox.showinfo("成功", f"图片已成功转换并保存为 {output_pdf}")


# Visio转换为PDF
def visio_to_pdf(visio_file, output_pdf):
    try:
        pythoncom.CoInitialize()  # 初始化COM库
        visio = win32com.client.Dispatch("Visio.Application")
        visio.Visible = False  # 不显示Visio界面
        doc = visio.Documents.Open(visio_file)
        doc.SaveAs(output_pdf, 32)  # 32表示保存为PDF格式
        doc.Close()
        visio.Quit()
        messagebox.showinfo("成功", f"Visio文件已成功转换为 {output_pdf}")
    except Exception as e:
        messagebox.showerror("错误", f"Visio文件转换失败: {str(e)}")


# Word转换为PDF
def word_to_pdf(word_file, output_pdf):
    try:
        pythoncom.CoInitialize()  # 初始化COM库
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 不显示Word界面
        doc = word.Documents.Open(word_file)
        doc.SaveAs(output_pdf, FileFormat=17)  # 17表示保存为PDF格式
        doc.Close()
        word.Quit()
        messagebox.showinfo("成功", f"Word文件已成功转换为 {output_pdf}")
    except Exception as e:
        messagebox.showerror("错误", f"Word文件转换失败: {str(e)}")



# 自动裁剪 PDF 文件的函数
def auto_crop_pdf(input_pdf_path, output_pdf_path, border_width=5):
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


# 合并 PDF 按钮的点击事件
def merge_button_click():
    pdf_files = file_list.get(0, tk.END)
    if not pdf_files:
        messagebox.showwarning("警告", "请先选择至少一个文件！")
        return

    # 转换所有非 PDF 文件为 PDF
    pdf_list = []
    for file in pdf_files:
        pdf_file = convert_to_pdf(file)
        if pdf_file:
            pdf_list.append(pdf_file)

    if not pdf_list:
        messagebox.showwarning("警告", "没有可合并的PDF文件！")
        return

    output_pdf = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")],
        title="保存合并后的PDF"
    )

    if output_pdf:
        merge_pdfs(pdf_list, output_pdf)


# 处理 PDF 文件裁切的点击事件
def process_pdf_files():
    """处理选择的多个PDF文件进行自动裁剪"""
    file_paths = file_list.get(0, tk.END)
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
        output_folder = filedialog.askdirectory(title="选择输出文件夹")  # 选择输出文件夹
        if not output_folder:
            messagebox.showwarning("警告", "请选择输出文件夹!")
            return

    status_label.config(text="正在处理，请稍候...")

    for input_pdf_path in file_paths:
        file_name = os.path.basename(input_pdf_path)
        output_file_name = os.path.splitext(file_name)[0] + "_cropped.pdf"
        output_pdf_path = os.path.join(output_folder, output_file_name)
        try:
            auto_crop_pdf(input_pdf_path, output_pdf_path, border_width)
        except Exception as e:
            messagebox.showerror("错误", "处理 PDF 文件时发生错误: " + str(e))
            status_label.config(text="发生错误，停止处理：" + str(e))
            return

    status_label.config(text="处理完成")
    messagebox.showinfo("完成", "PDF 文件裁剪成功")

    # 打开输出文件夹
    try:
        if os.name == 'nt':  # windows
            subprocess.Popen(['explorer', os.path.abspath(output_folder)])
        else:  # macOS or Linux
            subprocess.Popen(['open', os.path.abspath(output_folder)])
    except Exception as e:
        status_label.config(text=f"无法打开文件夹: {str(e)}")


# 选择文件并显示在列表框中
def browse_files():
    files = filedialog.askopenfilenames(
        title="选择文件",
        filetypes=(("所有文件", "*.*"),  # 允许选择所有文件格式
                   ("PDF Files", "*.pdf"),
                   ("Image Files", "*.jpg;*.jpeg;*.png;*.bmp"),
                   ("Visio Files", "*.vsd;*.vsdx"),
                   ("Word Files", "*.docx"))
    )
    if files:
        file_list.delete(0, tk.END)
        for file in files:
            file_list.insert(tk.END, file)

# 选择输出文件夹
def select_output_folder():
    folder_path = filedialog.askdirectory(title="选择输出文件夹")
    if folder_path:
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder_path)
        status_label.config(text="输出文件夹已选择：" + folder_path)


# 启用 DPI 感知
if os.name == 'nt':  # 仅在 Windows 上启用 DPI 感知
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e:
        print("Failed to set DPI awareness")

# 创建主窗口
window = tk.Tk()
window.title("PDF 自动裁剪与文件转换工具")

# 获取屏幕的 DPI
screen_dpi = window.winfo_fpixels('1i')

# 计算字体大小的缩放比例，并设置最大值
base_font_size = 10  # 设置基准字体大小
font_scale = screen_dpi / 96
max_font_size = 16  # 设置字体大小的最大值
scaled_font_size = min(int(base_font_size * font_scale), max_font_size)

# 根据DPI调整窗口大小
base_width = 700
base_height = 500
window_width = int(base_width * screen_dpi / 96)
window_height = int(base_height * screen_dpi / 96)
window.geometry(f"{window_width}x{window_height}")

default_font = ("Arial", scaled_font_size)  # 使用计算后的字体大小

# 使用 ttk 的主题样式
style = ttk.Style(window)
style.theme_use('clam')
style.configure('TLabel', font=default_font)
style.configure('TButton', font=default_font)
style.configure('TEntry', font=default_font)
style.configure('TLabelFrame', font=default_font)
style.configure("Listbox", font=default_font)

# 文件选择框架
file_frame = ttk.LabelFrame(window, text="选择文件")
file_frame.pack(padx=20, pady=10, fill="x")

# 文件列表框
file_list = tk.Listbox(file_frame, height=5)
file_list.pack(side=tk.LEFT, padx=5, pady=5, fill="x", expand=True)

# 选择文件按钮
select_file_button = ttk.Button(file_frame, text="选择文件", command=browse_files)
select_file_button.pack(side=tk.LEFT, padx=5, pady=5)

# 合并 PDF 按钮
merge_button = ttk.Button(window, text="合并或转换 PDF", command=merge_button_click)
merge_button.pack(pady=10)

# 处理 PDF 裁切按钮
process_pdf_button = ttk.Button(window, text="自动裁切 PDF", command=process_pdf_files)
process_pdf_button.pack(pady=10)

# 状态标签
status_label = ttk.Label(window, text="等待操作...")
status_label.pack(pady=10)

# 输出文件夹路径框
output_folder_frame = ttk.LabelFrame(window, text="输出文件夹路径")
output_folder_frame.pack(padx=20, pady=10, fill="x")

output_folder_entry = ttk.Entry(output_folder_frame)
output_folder_entry.pack(side=tk.LEFT, padx=5, pady=5, fill="x", expand=True)

# 选择文件夹按钮
select_output_button = ttk.Button(output_folder_frame, text="选择文件夹", command=select_output_folder)
select_output_button.pack(side=tk.LEFT, padx=5, pady=5)

# 设置边框宽度
border_width_frame = ttk.LabelFrame(window, text="边框宽度")
border_width_frame.pack(padx=20, pady=10, fill="x")

border_width_entry = ttk.Entry(border_width_frame)
border_width_entry.insert(0, "5")  # 默认边框宽度
border_width_entry.pack(side=tk.LEFT, padx=5, pady=5, fill="x", expand=True)

# 运行 GUI 窗口
window.mainloop()
