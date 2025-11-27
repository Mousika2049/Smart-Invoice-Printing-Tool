import os
import time
import win32api
import win32print
import decimal
import shutil
import re
from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import A4

#发票源文件夹相对路径
SOURCE_FOLDER = "source_pdf"
 
#文件输出文件夹
OUTPUT_FOLDER = "processed_pdf"

#设定A4纸张尺寸
A4_WIDTH, A4_HEIGHT = A4

# 缩放比例范围
SCALE_MIN = decimal.Decimal("0.70")
SCALE_MAX = decimal.Decimal("1.00")
SCALE_STEP = decimal.Decimal("-0.01")

# 单独成页的默认缩放比例
DEFAULT_STANDALONE_SCALE = 0.7

# 打印机名称 (可选, 如果留空，会使用系统默认打印机)
PRINTER_NAME = "" # 例如 "HP LaserJet Pro M404n"


def get_pdf_files(folder_path):
    """获取文件夹内所有PDF文件的路径"""
    return [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

def get_pdf_metadata(pdf_path):
    """一次性读取PDF文件的所有必要信息。返回一个包含宽度和高度的字典，如果读取失败则返回None。"""
    try:
        reader = PdfReader(pdf_path)
        if not reader.pages:
            print(f"警告: 文件 '{os.path.basename(pdf_path)}' 没有页面，已跳过。")
            return None
        page = reader.pages[0]
        return {
            #"path": pdf_path,
            "width": page.mediabox.width,
            "height": page.mediabox.height
        }
    except Exception as e:
        print(f"无法读取文件 '{os.path.basename(pdf_path)}' 的尺寸信息: {e}")
        return None

def find_optimal_scales(long_pdf_path, short_pdf_path):
    """查找最佳缩放比例"""
    short_scale = SCALE_MAX
    while short_scale >= SCALE_MIN:
        long_scale = SCALE_MAX
        while long_scale >= SCALE_MIN:
            # 计算缩放后的高度
            scaled_l_height = get_pdf_metadata(long_pdf_path)['height'] * float(long_scale)
            scaled_s_height = get_pdf_metadata(short_pdf_path)['height'] * float(short_scale)

            # 检查是否能纵向拼接进A4页面
            total_height = scaled_l_height + scaled_s_height 

            if total_height <= A4_HEIGHT:
                print(f"  > 找到可合并的缩放值: [L] 缩放比例 {long_scale*100:.0f}%, [S] 缩放比例 {short_scale*100:.0f}%")
                return float(long_scale), float(short_scale)
            
            long_scale += SCALE_STEP
        
        short_scale += SCALE_STEP
    
    return None, None # 未找到合适的缩放比例

def create_merged_pdf(pdf1_path, pdf1_scale, pdf2_path, pdf2_scale, output_path):
    """将两个发票PDF按缩放比例合并到一页A4上"""

    try:
        # 读取源PDF文件
        pdf1 = PdfReader(pdf1_path)
        pdf2 = PdfReader(pdf2_path)
        # 创建输出PDF
        output = PdfWriter()
        # 创建空白A4页面
        new_page = output.add_blank_page(width=A4_WIDTH, height=A4_HEIGHT)
        # 将第一个PDF放置在上半部分
        page1 = pdf1.pages[0]
        new_page.merge_transformed_page(page1, 
                                        [pdf1_scale, 0, 0, pdf1_scale, 
                                         (A4_WIDTH-get_pdf_metadata(pdf1_path)['width']*pdf1_scale)/2, 
                                         A4_HEIGHT-get_pdf_metadata(pdf1_path)['height']*pdf1_scale
                                         ])
        # 将第二个PDF放置在下半部分
        page2 = pdf2.pages[0]
        new_page.merge_transformed_page(page2, 
                                        [pdf2_scale, 0, 0, pdf2_scale, 
                                         (A4_WIDTH-get_pdf_metadata(pdf2_path)['width']*pdf2_scale)/2, 0
                                         ])
        # 保存结果
        with open(output_path, "wb") as f_out:
            output.write(f_out)
            print(f"此PDF文件合并成功！")
        return True 
    except Exception as e:
        print(f"文件合并失败: {os.path.basename(pdf1_path)}, {os.path.basename(pdf2_path)}. 错误: {e}")
        return False

def create_standalone_pdf(single_pdf_path, output_path):
    # 将未匹配的长发票单独打印
    single_pdf=PdfReader(single_pdf_path)

    output = PdfWriter()
    # 创建空白A4页面
    new_page = output.add_blank_page(width=A4_WIDTH, height=A4_HEIGHT)

    # 将PDF放置在上半部分
    page = single_pdf.pages[0]
    new_page.merge_transformed_page(page, [DEFAULT_STANDALONE_SCALE, 0, 0, DEFAULT_STANDALONE_SCALE, 
                                           (A4_WIDTH-get_pdf_metadata(single_pdf_path)['width']*DEFAULT_STANDALONE_SCALE)/2, 
                                           A4_HEIGHT-get_pdf_metadata(single_pdf_path)['height']*DEFAULT_STANDALONE_SCALE
                                           ])
    # 保存结果
    with open(output_path, "wb") as f_out:
        output.write(f_out)
        print(f"单只发票处理成功！")
    return True 

def print_pdf(pdf_path, printer_name):
    """调用系统命令打印PDF文件"""
    try:
        if not printer_name:
            printer_name = win32print.GetDefaultPrinter()
        
        print(f"\n正在发送打印任务:\n{os.path.basename(pdf_path)}\n-> 打印机: {printer_name}")
        win32api.ShellExecute(0, "print", pdf_path, f'"{printer_name}"', ".", 0)
        # 等待一小段时间，防止打印任务发送过快
        time.sleep(2)
    except Exception as e:
        print(f"打印时发生错误: {e}")


def main():
    """主函数体"""
    print(">>> 智能发票打印工具 <<<\n")
    
    # 检查源文件夹
    if not os.path.exists(SOURCE_FOLDER):
        print(f"错误：源文件夹 '{SOURCE_FOLDER}' 不存在。请创建该文件夹并放入PDF文件。")
        return
    else:
        print(f" > PDF源文件夹 '{SOURCE_FOLDER}'")
    
    # 删除并重建输出文件夹
    if os.path.exists(OUTPUT_FOLDER):
        print(f"检测到旧的输出文件夹 '{OUTPUT_FOLDER}'，正在清空...")
        try:
            shutil.rmtree(OUTPUT_FOLDER) # 删除整个文件夹树
        except OSError as e:
            # 如果删除失败（例如文件被占用），则打印错误并终止
            print(f"错误：无法删除文件夹 {OUTPUT_FOLDER}。请检查文件是否被占用。")
            print(f"错误信息: {e}")
            return
    
    # 删除后创建一个新的空文件夹
    os.makedirs(OUTPUT_FOLDER)
    print(f"已重新创建输出文件夹 '{OUTPUT_FOLDER}'")

    # 获取所有PDF文件
    pdf_path_list = get_pdf_files(SOURCE_FOLDER)
    if not pdf_path_list:
        print(f" > 在文件夹 '{SOURCE_FOLDER}' 中没有找到任何PDF文件。")
        return
    print(f" > 在文件夹 '{SOURCE_FOLDER}' 中找到 {len(pdf_path_list)} 个发票PDF文件。")
    

    #按PDF高度降序排列文件
    sorted_pdfs = sorted(pdf_path_list, key=lambda k:get_pdf_metadata(k)['height'], reverse=True)
    
    # 处理并准备打印列表
    files_to_print = []
    # 可配对文件计数
    merged_pair_count = 0
    # 独立处理文件计数
    standalone_file_count = 0

    # 如果文件个数是奇数，直接单独打印最长的PDF
    if len(sorted_pdfs) % 2 != 0:
        print(f"\n发票PDF文件个数为奇数，先单独处理最长PDF文件：")
        longest_pdf_path = sorted_pdfs.pop(0)
        longest_pdf_name = os.path.basename(longest_pdf_path)
        output_filename = f"single_{longest_pdf_name}.pdf"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        print(f"[SA] {longest_pdf_name}")
        create_standalone_pdf(longest_pdf_path, output_path)
        files_to_print.append(output_path)
        standalone_file_count += 1
        

    while len(sorted_pdfs) >= 2:
        long_pdf_path = sorted_pdfs.pop(0)  # 最长的
        short_pdf_path = sorted_pdfs.pop(-1) # 最短的
        long_pdf_name = os.path.basename(long_pdf_path)
        short_pdf_name = os.path.basename(short_pdf_path)
        print(f"\n正在尝试配对可合并文件:\n[L] {long_pdf_name} \n[S] {short_pdf_name}")

        # 查找最佳缩放比例
        l_scale, s_scale = find_optimal_scales(long_pdf_path, short_pdf_path)

        if l_scale and s_scale:
            # 配对成功，合并
            # 正则表达式提取发票号码简化output_filename
            #match = re.search(r'\d{20}')
            output_filename = f"{long_pdf_name}_{short_pdf_name}.pdf"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            create_merged_pdf(long_pdf_path, l_scale, short_pdf_path, s_scale, output_path)
            merged_pair_count+=1
            files_to_print.append(output_path)
        
        else:
            # 无可配对文件，单独处理长PDF，将短PDF放回待处理列表
            print(f"  > 该文件无可配对文件...单独处理 {long_pdf_name}")
            output_filename = f"single_long_{long_pdf_name}.pdf"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            create_standalone_pdf(long_pdf_path, output_path)
            files_to_print.append(output_path)
            standalone_file_count += 1

            # 将短的PDF放回原列表并重新排序
            print(f"  > 将较短的PDF '{short_pdf_name}' 放回原列表以待下次配对")
            sorted_pdfs.append(short_pdf_path)
            sorted_pdfs.sort(key=lambda k:get_pdf_metadata(k)['height'], reverse=True)
            
    # 处理最后剩下的PDF文件（配对失败剩下的）
    if sorted_pdfs:
        print("\n--- 开始处理剩余的文件 ---")
        remaining_pdf_path = sorted_pdfs[0]
        pdf_name = os.path.basename(remaining_pdf_path)
        print(f"处理单个文件: {pdf_name}")
        output_filename = f"single_{pdf_name.split('.')[0]}.pdf"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        create_standalone_pdf(remaining_pdf_path, output_path)
        files_to_print.append(output_path)
        standalone_file_count += 1
    
    # 汇总处理情况
    print("\n所有文件处理完毕！")
    print(f"\n--- 已配对可合并文件计数 {merged_pair_count*2} ---")      
    print(f"\n--- 已单独处理文件计数 {standalone_file_count} ---")  
    
    # 执行打印
    if not files_to_print:
        print("没有需要打印的文件。")
        return

    print("\n--- 开始发送打印任务 ---")
    try:
        user_input = input(f"准备打印 {len(files_to_print)} 个目标文档。是否继续？(y/n): ")
        if user_input.lower() == 'y':
            for pdf_file in files_to_print:
                print_pdf(pdf_file, PRINTER_NAME)
            print("\n所有打印任务已发送完毕！")
        else:
            print(f"打印操作已取消。\n处理好的文件保存在 '{OUTPUT_FOLDER}' 文件夹中，您可以手动打印或重试。")
    except KeyboardInterrupt:
        print("\n操作被用户中断。您可以重试。")

if __name__ == "__main__":
    main()