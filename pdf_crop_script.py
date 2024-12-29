import fitz
from PIL import Image
import numpy as np
import os
import argparse
import glob


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

def process_pdf_files(input_files, output_folder, border_width=5):
    """
    处理多个 PDF 文件。

    Args:
        input_files: 一个包含输入 PDF 文件路径的列表。
        output_folder: 输出文件夹路径。
        border_width: 需要忽略的边框宽度，默认为 5 像素。
    """
    if not input_files:
        print("错误: 没有提供任何 PDF 文件路径.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder) # 如果输出文件夹不存在，创建它
    
    for input_pdf_path in input_files:
        if not os.path.exists(input_pdf_path):
            print(f"错误: 输入文件不存在: {input_pdf_path}")
            continue

        file_name = os.path.basename(input_pdf_path)
        output_file_name = os.path.splitext(file_name)[0] + "_cropped.pdf"
        output_pdf_path = os.path.join(output_folder, output_file_name)
        try:
            auto_crop_pdf(input_pdf_path, output_pdf_path, border_width)
            print(f"已裁剪: {input_pdf_path}  ->  {output_pdf_path}")
        except Exception as e:
            print(f"处理 {input_pdf_path} 时发生错误：{e}")
    
    print(f"处理完成。裁剪后的文件保存在：{os.path.abspath(output_folder)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="自动裁剪 PDF 文件中的图表。")
    parser.add_argument(
        "input_files",
        nargs="*",
        help="要处理的 PDF 文件路径（如果没有提供，则处理当前目录下的所有 PDF 文件）。",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="output",
        help="输出文件夹路径，默认是当前目录下的 'output' 文件夹。",
    )
    parser.add_argument(
        "-b",
        "--border",
        type=int,
        default=5,
        help="需要忽略的边框宽度（像素），默认为 5。",
    )
    
    args = parser.parse_args()

    input_files = args.input_files
    output_folder = args.output
    border_width = args.border

    if not input_files:  # 如果没有提供输入文件，则处理当前目录下的所有 PDF 文件
       input_files = glob.glob("*.pdf")
    
    process_pdf_files(input_files, output_folder, border_width)