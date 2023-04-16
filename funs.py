import os
import re
import io

# 导入PyPDF2模块
from PyPDF2 import PdfReader, PdfWriter, PdfMerger, PageRange
import img2pdf
from PIL import Image
import pikepdf


def make_dir(f_path):
    output_path = os.path.join(os.path.dirname(f_path), 'pdf_kit')
    if not os.path.exists(output_path):
        os.mkdir(output_path)
    return output_path


def make_newpage(page_string, interval=1, group=1):
    pattern = r',|，|\\|/|:|;|；'
    page_list = re.split(pattern, page_string)
    mid_list = []
    for g in range(1, group):
        for part in page_list:
            if '-' in part:
                start, end = part.split('-')
                p = f'{int(start) + interval * g}-{int(end) + interval * g}'
            else:
                p = int(part) + interval * g
            mid_list.append(p)
    page_list.extend(mid_list)
    new_page = ','.join(map(str, page_list))
    return new_page


def make_page_list(page_range):
    pattern = r',|，|\\|/|:|;|；'
    page_range_list = []
    for part in re.split(pattern, page_range):
        if "-" in part:
            start, end = part.split("-")
            page_range_list.extend(range(int(start) - 1, int(end)))
        else:
            page_range_list.append(int(part) - 1)

    return page_range_list


def get_max_page(file):
    """获取pdf最大页数"""
    pdf_reader = PdfReader(file)
    pdf_pages = len(pdf_reader.pages)

    return pdf_pages


def odd_merge_even(odd_file, output_path, even_file):
    """pdf合并奇偶页"""
    writer = PdfWriter()
    odd = PdfReader(odd_file)
    even = PdfReader(even_file)

    # 循环遍历两个pdf文件的每一页，按照奇偶顺序将其添加到writer对象中
    for i in range(max(len(odd.pages), len(even.pages))):
        if i < len(odd.pages):
            writer.add_page(odd.pages[i])
        if i < len(even.pages):
            writer.add_page(even.pages[i])

    output_file_path = os.path.join(output_path, 'odd-merged-even.pdf')

    # 将writer对象写入到一个新的pdf文件中
    with open(output_file_path, 'wb') as f:
        writer.write(f)


def reverse_pdf(file, output_path):
    """倒序"""
    output_pdf = PdfWriter()
    with open(file, 'rb') as readfile:
        input_pdf = PdfReader(readfile)
        for page in reversed(input_pdf.pages):
            output_pdf.add_page(page)
    output_file_path = os.path.join(output_path, f'{os.path.basename(os.path.splitext(file)[0])}_reversed.pdf')
    with open(output_file_path, 'wb') as writefile:
        output_pdf.write(writefile)


def rotate_page(file, page_list, mark, angle, output_path):
    """旋转指定页面和旋转角度"""

    # 创建一个PDF读取器对象
    pdf_reader = PdfReader(file)

    # 创建一个PDF写入器对象
    pdf_writer = PdfWriter()

    max_pages = len(pdf_reader.pages)

    if mark == 0:
        new_page = page_list
    if mark == 1:
        new_page = page_list[1::2]
    if mark == 2:
        new_page = page_list[::2]

    for i in range(max_pages):
        page = pdf_reader.pages[i]
        if i in new_page:
            page.rotate(angle)

        pdf_writer.add_page(page)

    output_file_path = os.path.join(output_path, f'{os.path.basename(os.path.splitext(file)[0])}_rotated.pdf')
    # 创建一个新的PDF文件并将旋转后的页面添加到其中
    with open(output_file_path, 'wb') as f:
        pdf_writer.write(f)


def extract_pages(pdf_file, output_path, pages_string, seperate):
    """
    提取指定页码的PDF页面
    :param seperate: 是否独立成文件
    :param output_path: 输出路径
    :param pdf_file: PDF文件路径
    :param pages_string: 需要提取的页面范围，格式为"1,3-5,7"
    :return:
    """
    if seperate:
        page_list = re.split(r',|，|；|;', pages_string)
        for part in page_list:
            pdf_reader = PdfReader(pdf_file)
            pdf_writer = PdfWriter()
            if '-' in part:
                start, end = part.split("-")
                for page in range(int(start) - 1, int(end)):
                    pdf_writer.add_page(pdf_reader.pages[page])
                    output_file_path = os.path.join(output_path,
                                                    f'{os.path.basename(os.path.splitext(pdf_file)[0])}_{part}.pdf')
                    with open(output_file_path, 'wb') as f:
                        pdf_writer.write(f)
            else:
                pdf_writer.add_page(pdf_reader.pages[int(part) - 1])
                output_file_path = os.path.join(output_path,
                                                f'{os.path.basename(os.path.splitext(pdf_file)[0])}_{part}.pdf')
                with open(output_file_path, 'wb') as f:
                    pdf_writer.write(f)
    else:
        # 将字符串转换为页码列表
        page_nums = make_page_list(pages_string)

        # 打开PDF文件并提取指定页面
        pdf_reader = PdfReader(pdf_file)
        pdf_writer = PdfWriter()
        for page_num in page_nums:
            pdf_writer.add_page(pdf_reader.pages[page_num])

        output_file_path = os.path.join(output_path, f'{os.path.basename(os.path.splitext(pdf_file)[0])}_extracted.pdf')
        with open(output_file_path, 'wb') as f:
            pdf_writer.write(f)


def split_pdf_by_page_number(input_path, output_path, num_pages_per_split):
    with open(input_path, 'rb') as input_pdf:
        pdf_reader = PdfReader(input_pdf)
        num_pages = len(pdf_reader.pages)
        if len(pdf_reader.pages) >= num_pages_per_split > 0:
            for i in range(0, num_pages, num_pages_per_split):
                output_file_path = os.path.join(
                    output_path,
                    f'{os.path.basename(os.path.splitext(input_path)[0])}_{i + 1}-{min(i + num_pages_per_split, num_pages)}.pdf'
                )
                pdf_writer = PdfWriter()

                for j in range(i, min(i + num_pages_per_split, num_pages)):
                    pdf_writer.add_page(pdf_reader.pages[j])

                with open(output_file_path, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)
        else:
            print("输入错误!!!")


def merge_pdfs(pdf_list, page_list, output_path):
    # 初始化一个PdfFileMerger对象
    pdf_writer = PdfWriter()

    # 将所有PDF文件添加到PdfMerger对象中
    for pdf_file, page_string in zip(pdf_list, page_list):
        pdf = PdfReader(pdf_file)
        page_num = []
        for part in page_string.split(','):
            if '-' in part:
                start, end = part.split('-')
                page_num.extend(range(int(start) - 1, int(end)))
            else:
                page_num.append(int(part) - 1)
        for p in page_num:
            pdf_writer.add_page(pdf.pages[p])
    output_file = os.path.join(output_path, 'multi-merged.pdf')
    with open(output_file, 'wb') as f:
        pdf_writer.write(f)


def merge_imgs(img_list, output_path):
    for img_file in img_list:
        img = Image.open(img_file)
        img.convert('RGB')
        img.save(img_file, quality=95)
    a4inpt = (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
    layout_fun = img2pdf.get_layout_fun(a4inpt)
    output_file = os.path.join(output_path, 'multi-merged.pdf')
    with open(output_file, "wb") as f:
        f.write(img2pdf.convert(img_list, layout_fun=layout_fun, rotation=img2pdf.Rotation.ifvalid))


def unlock_file(file, output_path):
    output_file_path = os.path.join(output_path, os.path.basename(file))
    pdf = pikepdf.open(file)
    pdf.save(output_file_path)




def insert_page(file1, file2, page_list, output_path, position=0):
    """
    实现插入页面功能
    :param file1: 原始文件
    :param file2: 要插入的文件
    :param output_path: 保存的路径
    :return:
    """
    pdf1 = PdfReader(file1)  # 原始文件
    pdf2 = PdfReader(file2)
    mid_writer = PdfWriter()

    for i in page_list:
        page = pdf2.pages[i]
        mid_writer.add_page(page)
    pdf_buffer = io.BytesIO()
    mid_writer.write(pdf_buffer)

    reader = PdfReader(pdf_buffer)
    pdf_merger = PdfMerger()
    pdf_merger.append(pdf1)
    ran = PageRange(":")
    pdf_merger.merge(page_number=position, fileobj=reader, pages=ran)

    output_file_path = os.path.join(output_path, f'{os.path.basename(os.path.splitext(file1)[0])}_insert.pdf')
    pdf_merger.write(output_file_path)
    pdf_merger.close()


def delete_pages(pdf_file, pages_list, output_path):
    pdf_reader = PdfReader(pdf_file)
    pdf_writer = PdfWriter()

    for num in range(len(pdf_reader.pages)):
        if num in pages_list:
            continue
        page = pdf_reader.pages[num]
        pdf_writer.add_page(page)
    output_file_path = os.path.join(output_path,
                                    f'{os.path.basename(os.path.splitext(pdf_file)[0])}_deleted.pdf')
    with open(output_file_path, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)


def replace_pages(origin_f, place_f, origin_l, place_l, output_path):
    origin_reader = PdfReader(origin_f)
    place_reader = PdfReader(place_f)

    new_pdf = PdfWriter()

    for n in range(len(origin_reader.pages)):
        page = origin_reader.pages[n]
        for i, j in zip(origin_l, place_l):
            if n == i:
                page = place_reader.pages[j]
        new_pdf.add_page(page)

    output_file_path = os.path.join(output_path,
                                    f'{os.path.basename(os.path.splitext(origin_f)[0])}_replaced.pdf')
    with open(output_file_path, 'wb') as f:
        new_pdf.write(f)

# if __name__ == '__main__':
#     file1 = r"C:\Users\kindred\Desktop\pdf-kit-test\1\odd.PDF"
#     file2 = r"C:\Users\kindred\Desktop\pdf-kit-test\1\even.pdf"
#
#     insert_page(file1, file2)
