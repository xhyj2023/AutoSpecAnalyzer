import subprocess
import time
import os
import pandas as pd
import re
import sys
from datetime import datetime


# 获取当前脚本所在目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# 文件路径（全部默认在脚本同级目录）
recognize_script = os.path.join(BASE_DIR, "recognize_spec.py")
crawl_script = os.path.join(BASE_DIR, "crawl_price.py")
txt_path = os.path.join(BASE_DIR, "specs.txt")
excel_path = os.path.join(BASE_DIR, "compare_data.xlsx")
output_file = os.path.join(BASE_DIR, "matched_result.xlsx")


# 启动识别脚本
print(">>> 启动识别脚本，正在监听图片文件夹...")
subprocess.Popen(["python", recognize_script])  # 非阻塞启动识别脚本


# 等待 specs.txt 文件生成
print("等待 specs.txt 文件生成...处理中[99%] ")
while not os.path.exists(txt_path):
    time.sleep(1)
time.sleep(0.5)  # 确保文件写入完成
print(f"Done！ {os.path.basename(txt_path)} 已生成，开始 Excel 筛选")


# 检查 Excel 文件是否已存在且是当天生成的
def is_file_today(file_path):
    """检查文件是否是今天生成"""
    if not os.path.exists(file_path):
        return False
    file_modified_time = os.path.getmtime(file_path)
    file_date = datetime.fromtimestamp(file_modified_time).strftime("%Y-%m-%d")
    today = datetime.now().strftime("%Y-%m-%d")
    return file_date == today

if is_file_today(excel_path):
    print(f"太好了！ {os.path.basename(excel_path)} 已存在，跳过爬取")
else:
    print("[\] 正在运行爬取脚本以获取最新 Excel 数据...")
    subprocess.run(["python", crawl_script], check=True)

if not os.path.exists(excel_path):
    raise FileNotFoundError(f"❌ 找不到 Excel 文件：{excel_path}")


# 设置输出编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')


def mm_to_feet(mm_value):
    """
    将毫米转换为英尺
    1英尺 = 304.8mm
    """
    feet = mm_value / 304.8
    return round(feet)  # 四舍五入到整数英尺


def parse_specs_from_table(file_path):
    """
    从specs.txt中解析表格数据，提取厚度、长宽、材质信息
    返回规格列表，每个规格包含：材质、厚度、宽度（英尺）、长度（英尺）
    """
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    specs_list = []

    # 提取材质信息（304）
    material_pattern = r'(\d{3})[#\s]'
    materials = re.findall(material_pattern, content)
    material = materials[0] if materials else '304'

    # 从表格中提取规格数据
    # 匹配表格行，格式如：| 3 | ...材质... | 2440 × 1220 | 0.94 | 1 |
    # 需要提取：规格（长×宽）、实厚、厚度
    table_pattern = r'\|\s*\d+\s*\|[^|]+\|\s*(\d+)\s*[×x]\s*(\d+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|'
    matches = re.findall(table_pattern, content)

    if matches:
        print(f"从表格中提取到 {len(matches)} 组规格数据：")
        for match in matches:
            length_mm = int(match[0])  # 长度（mm）
            width_mm = int(match[1])  # 宽度（mm）
            real_thickness = float(match[2])  # 实厚
            thickness = float(match[3])  # 厚度

            # 转换为英尺
            length_feet = mm_to_feet(length_mm)
            width_feet = mm_to_feet(width_mm)
            
            # 计算筛选目标厚度：厚度 - 0.05
            filter_thickness = thickness - 0.05

            spec = {
                'material': material,
                'real_thickness': real_thickness,
                'thickness': thickness,
                'filter_thickness': filter_thickness,  # 新增：筛选用的厚度值
                'length_mm': length_mm,
                'width_mm': width_mm,
                'length_feet': length_feet,
                'width_feet': width_feet,
                'spec_str': f"{width_feet}*{length_feet}"  # 宽*长（英尺格式）
            }
            specs_list.append(spec)

            print(f"  - 材质:{material} 厚度:{thickness}mm (厚度-0.05={filter_thickness:.2f}mm) "
                  f"长宽:{length_mm}*{width_mm}mm → {length_feet}*{width_feet}英尺")
    else:
        print("警告: 未能从specs.txt中提取到表格数据")

    return specs_list


def filter_by_specs(specs_list, input_file, output_file):
    
    # 读取Excel文件
    print(f"\n正在读取 {input_file}...") 
    df = pd.read_excel(input_file)

    print(f"原始数据行数: {len(df)}")
    print(f"数据列名: {df.columns.tolist()}")

    # 创建空的DataFrame用于存储所有匹配结果
    all_filtered_df = pd.DataFrame()

    # 遍历每个规格进行筛选
    for idx, spec in enumerate(specs_list, 1):
        print(f"\n{'=' * 60}")
        print(f"[规格 {idx}/{len(specs_list)}] 筛选条件:")
        print(f"  材质: {spec['material']}")
        print(f"  厚度: {spec['thickness']}mm → 筛选厚度: {spec['filter_thickness']:.2f}mm")
        print(f"  规格: {spec['spec_str']} (英尺)")
        print(f"  原始尺寸: {spec['length_mm']}*{spec['width_mm']}mm")

        # 创建筛选条件
        mask = pd.Series([True] * len(df))

        # 条件1: 材质匹配
        if '材质' in df.columns:
            material_mask = df['材质'].astype(str).str.contains(spec['material'], na=False, case=False)
            mask = mask & material_mask
            print(f"  → 材质匹配: {material_mask.sum()} 条")

        # 条件2: 使用厚度 - 0.05的确切数值匹配（允许±0.01容差）
        if '厚度' in df.columns:
            filter_thick = spec['filter_thickness']
            thickness_mask = (
                (pd.to_numeric(df['厚度'], errors='coerce') >= filter_thick - 0.01) &
                (pd.to_numeric(df['厚度'], errors='coerce') <= filter_thick + 0.01)
            )
            mask = mask & thickness_mask
            print(f"  → 厚度匹配 (目标{filter_thick:.2f}±0.01): {thickness_mask.sum()} 条")

        # 条件3: 规格匹配（英尺格式）
        if '规格' in df.columns:
            # 完全匹配规格字符串
            spec_mask = df['规格'].astype(str) == spec['spec_str']
            mask = mask & spec_mask
            print(f"  → 规格匹配: {spec_mask.sum()} 条")

        # 应用筛选
        filtered_df = df[mask].copy()
        print(f"  ✓ 本次筛选结果: {len(filtered_df)} 条")

        # 添加标记列，标识这是哪个规格的匹配结果
        if len(filtered_df) > 0:
            filtered_df['匹配规格'] = f"{spec['material']} 厚度{spec['thickness']}mm {spec['length_mm']}*{spec['width_mm']}mm"
            filtered_df['匹配规格_英尺'] = f"{spec['spec_str']}"
            filtered_df['筛选厚度'] = f"{spec['filter_thickness']:.2f}mm"

            # 合并到总结果中
            all_filtered_df = pd.concat([all_filtered_df, filtered_df], ignore_index=True)

    print(f"\n{'=' * 60}")
    print(f"所有规格筛选完成！")
    print(f"总共筛选出: {len(all_filtered_df)} 条数据")

    # 去重
    original_count = len(all_filtered_df)
    all_filtered_df = all_filtered_df.drop_duplicates(subset=['编号'], keep='first')
    if len(all_filtered_df) < original_count:
        print(f"去重后: {len(all_filtered_df)} 条数据 (去除了 {original_count - len(all_filtered_df)} 条重复)")

    # 保存结果
    if len(all_filtered_df) > 0:
        all_filtered_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\n✓ 结果已保存到: {output_file}")

        # 显示结果预览
        print(f"\n筛选结果预览 (前10行):")
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 30)
        print(all_filtered_df.head(10).to_string())

        # 统计信息
        print(f"\n按匹配规格分组统计:")
        if '匹配规格' in all_filtered_df.columns:
            print(all_filtered_df['匹配规格'].value_counts().to_string())
    else:
        print("\n⚠ 警告: 没有找到匹配的数据！")
        print("请检查：")
        print("1. specs.txt中的规格信息是否正确")
        print("2. compare_data.xlsx中是否有符合条件的数据")
        print("3. 筛选条件是否过于严格")

    return all_filtered_df


def main():
    # 文件路径
    specs_file = 'specs.txt'
    input_file = 'compare_data.xlsx'
    output_file = 'filtered_result.xlsx'

    print("=" * 60)
    print("开始从specs.txt提取规格并筛选compare_data.xlsx")
    print("=" * 60)

    # 解析specs.txt
    print(f"\n[步骤1] 解析 {specs_file}...")
    specs_list = parse_specs_from_table(specs_file)

    if not specs_list:
        print("错误: 未能从specs.txt中提取到任何规格信息！")
        return

    print(f"\n共提取到 {len(specs_list)} 组规格信息")

    # 筛选数据
    print(f"\n[步骤2] 在 {input_file} 中筛选匹配数据...")
    filtered_df = filter_by_specs(specs_list, input_file, output_file)

    print("\n" + "=" * 60)
    print("处理完成!")
    print("=" * 60)


if __name__ == '__main__':
    main()

