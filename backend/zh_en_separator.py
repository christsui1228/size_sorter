import pandas as pd
import re
import os
from openpyxl import load_workbook
import xlrd

def get_input_path():
    return input("请输入要处理的Excel文件的完整路径：").strip()

def get_output_dir(input_path):
    output_dir = input("请输入保存结果的目录路径（按回车键使用默认输出目录）：").strip()
    return output_dir if output_dir else os.path.dirname(input_path)

def read_excel(file_path):
    _, file_extension = os.path.splitext(file_path)
    if file_extension.lower() == '.xlsx':
        # 使用 openpyxl 读取 .xlsx 文件
        return pd.read_excel(file_path, engine='openpyxl')
    elif file_extension.lower() == '.xls':
        # 使用 xlrd 读取 .xls 文件
        return pd.read_excel(file_path, engine='xlrd')
    else:
        raise ValueError("不支持的文件格式。请使用 .xlsx 或 .xls 文件。")

def clean_data(df):
    df['A'] = df['A'].apply(lambda x: re.sub(r'([\u4e00-\u9fa5]+)([A-Za-z])', r'\1 \2', str(x)))
    return df

def split_name(full_name):
    match = re.match(r'^([\u4e00-\u9fa5]+)\s*([A-Za-z\s.]+)$', str(full_name).strip())
    if match:
        return pd.Series({'chinese_name': match.group(1).strip(), 'english_name': match.group(2).strip()})
    return pd.Series({'chinese_name': '', 'english_name': ''})

def process_names(input_file, output_file):
    # 读取输入文件
    df = read_excel(input_file)

    # 确保文件包含 A 和 B 列
    if len(df.columns) < 2:
        raise ValueError("输入文件应至少包含两列（A 和 B）")

    # 重命名列
    df.columns = ['A', 'B'] + list(df.columns[2:])

    # 清理数据
    df = clean_data(df)

    # 应用分离函数到 A 列
    df[['chinese_name', 'english_name']] = df['A'].apply(split_name)

    # 创建新的 DataFrame，按照要求的顺序排列列
    new_df = pd.DataFrame({
        'English': df['english_name'],
        'B': df['B'],
        'Chinese': df['chinese_name']
    })

    # 将结果写入 Excel 文件
    new_df.to_excel(output_file, index=False, engine='openpyxl')

    print(f"处理完成。结果已保存到 {output_file}")
    print(f"总共处理了 {len(df)} 条记录")

def main():
    input_path = get_input_path()
    output_dir = get_output_dir(input_path)
    
    # 生成输出文件名
    input_filename = os.path.basename(input_path)
    output_filename = f"{os.path.splitext(input_filename)[0]}_separated.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    # 处理数据
    process_names(input_path, output_path)

if __name__ == "__main__":
    main()