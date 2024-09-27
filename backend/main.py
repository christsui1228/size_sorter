import pandas as pd
import re
import os
import sys
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 定义尺码排序顺序
SIZE_ORDER = ['100','110','120','130','140','150','XS','S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL','6XL','7XL','8XL','9XL','10XL']

def get_input_path():
    return input("请输入要处理的Excel文件的完整路径：").strip()

def get_output_dir(input_path):
    output_dir = input("请输入保存结果的目录路径（按回车键使用默认输出目录）：").strip()
    return output_dir if output_dir else os.path.dirname(input_path)

def clean_and_order_size(size):
    size = str(size).upper().strip()
    if size in SIZE_ORDER:
        return SIZE_ORDER.index(size)
    elif size.endswith('XL'):
        x_count = size.count('X')
        if x_count > 1:
            return SIZE_ORDER.index('XL') + x_count
    return len(SIZE_ORDER)  # 未知尺码放到最后

def process_excel(input_path, output_dir):
    try:
        # 读取Excel文件
        df = pd.read_excel(input_path)
        logging.info(f"成功读取Excel文件，共 {len(df)} 行")
        
        # 确保列名正确
        if len(df.columns) >= 2:
            df = df.iloc[:, :2]  # 只取前两列
            df.columns = ['姓名', '尺码']
        else:
            raise ValueError("Excel文件应至少包含两列数据（姓名和尺码）")

        # 数据清洗和排序
        df['尺码排序'] = df['尺码'].apply(clean_and_order_size)
        df = df.sort_values('尺码排序')
        df = df.drop('尺码排序', axis=1)

        # 添加序号列
        df['序号'] = range(1, len(df) + 1)
        df = df[['序号', '姓名', '尺码']]

        # 创建新的Excel文件
        source_filename = os.path.splitext(os.path.basename(input_path))[0]
        output_filename = f'{source_filename}_sorted.xlsx'
        output_path = os.path.join(output_dir, output_filename)
        
        # 使用openpyxl保存并格式化
        wb = Workbook()
        ws = wb.active
        ws.title = "排序后数据"

        # 写入标题
        for col, header in enumerate(['序号', '姓名', '尺码'], start=1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        # 写入数据并设置格式
        for row, (idx, name, size) in enumerate(df.values, start=2):
            ws.cell(row=row, column=1, value=idx).alignment = Alignment(horizontal='center')
            ws.cell(row=row, column=2, value=name).alignment = Alignment(horizontal='left')
            ws.cell(row=row, column=3, value=size).alignment = Alignment(horizontal='center')

        # 调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column_letter].width = adjusted_width

        wb.save(output_path)
        logging.info(f"文件已成功保存: {output_path}")

    except Exception as e:
        logging.error(f"处理过程中出现错误: {str(e)}")
        raise

def main():
    logging.info(f"Python version: {sys.version}")
    logging.info(f"Pandas version: {pd.__version__}")

    input_path = get_input_path()
    output_dir = get_output_dir(input_path)
    os.makedirs(output_dir, exist_ok=True)

    try:
        process_excel(input_path, output_dir)
        logging.info("处理完成。请检查输出文件。")
    except Exception as e:
        logging.error(f"程序执行失败: {str(e)}")

if __name__ == "__main__":
    main()