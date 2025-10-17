import subprocess
import time
import os
import pandas as pd
import re
import sys
# 获取当前脚本所在目录

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 文件路径（全部默认在脚本同级目录）

recognize_script = os.path.join(BASE_DIR, "recognize_spec.py")
crawl_script = os.path.join(BASE_DIR, "crawl_price.py")
txt_path = os.path.join(BASE_DIR, "specs.txt")
excel_path = os.path.join(BASE_DIR, "compare_data.xlsx")
output_file = os.path.join(BASE_DIR, "matched_result.xlsx")


# 1️⃣ 启动识别脚本

print("👀 启动识别脚本，正在监听图片文件夹...")
subprocess.Popen(["python", recognize_script])  # 非阻塞启动识别脚本


# 2️⃣ 等待 specs.txt 文件生成

print("⏳ 等待 specs.txt 文件生成...")
while not os.path.exists(txt_path):
    time.sleep(1)
time.sleep(0.5)  # 确保文件写入完成
print(f"✅ {os.path.basename(txt_path)} 已生成，开始 Excel 筛选")


# 3️⃣ 启动爬取脚本生成 Excel 数据

print("📊 正在运行爬取脚本以获取最新 Excel 数据...")
subprocess.run(["python", crawl_script], check=True)

if not os.path.exists(excel_path):
    raise FileNotFoundError(f"❌ 找不到 Excel 文件：{excel_path}")

# 设置输出编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')


def parse_specs_txt(file_path):
    """
    解析specs.txt文件，提取公司、规格、厚度、长宽信息
    """
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    # 初始化提取的信息列表
    specs_info = {
        'companies': set(),
        'materials': set(),  # 材质/规格
        'thicknesses': set(),  # 厚度
        'dimensions': set()  # 长宽组合
    }

    # 提取供应商名称/公司（支持多种格式）
    company_patterns = [
        r'供应商名称[*\s]*[：:]\s*([^\n\\]+?)(?:\n|\\n)',
    ]
    for pattern in company_patterns:
        companies = re.findall(pattern, content)
        for company in companies:
            company = company.strip().replace('**', '').replace('*', '').replace('：', '').strip()
            # 过滤掉无效的公司名
            if (company and
                    company != '未指定' and
                    len(company) < 50 and  # 公司名不应该太长
                    not company.startswith('提供产品') and
                    not company.startswith('供应商')):
                specs_info['companies'].add(company)

    # 提取材质及表面贴膜要求（作为规格）
    material_patterns = [
        r'[*\s-]*材质及表面贴膜要求[*\s]*[：:]\s*([^\n]+)',
        r'材质[：:]\s*([^\n]+)'
    ]
    for pattern in material_patterns:
        materials = re.findall(pattern, content)
        for material in materials:
            material = material.strip().replace('**', '').replace('*', '')
            if material:
                # 从材质中提取304等材料代号
                material_codes = re.findall(r'(\d{3})[#\s]', material)
                for code in material_codes:
                    specs_info['materials'].add(code)

    # 提取厚度（支持Markdown格式）
    thickness_patterns = [
        r'[*\s-]*厚度[（(]mm[)）][*\s]*[：:]\s*([\d.]+)',
        r'厚度[：:]\s*([\d.]+)',
        r'(\d+\.?\d*)\s*mm'  # 匹配 "1.2mm" 这种格式
    ]
    for pattern in thickness_patterns:
        thicknesses = re.findall(pattern, content)
        for thickness in thicknesses:
            try:
                t = float(thickness)
                # 只添加合理范围内的厚度值（0.1-100mm）
                if 0.1 <= t <= 100:
                    specs_info['thicknesses'].add(t)
            except:
                pass

    # 提取实厚
    real_thickness_patterns = [
        r'[*\s-]*实厚[（(]mm[)）][*\s]*[：:]\s*[≥>=]?\s*([\d.]+)',
        r'实厚[：:]\s*[≥>=]?\s*([\d.]+)'
    ]
    for pattern in real_thickness_patterns:
        real_thicknesses = re.findall(pattern, content)
        for thickness in real_thicknesses:
            try:
                t = float(thickness)
                if 0.1 <= t <= 100:
                    specs_info['thicknesses'].add(t)
            except:
                pass

    # 提取长宽信息
    length_patterns = [
        r'[*\s-]*长[（(]mm[)）][*\s]*[：:]\s*([\d.]+)',
        r'长[：:]\s*([\d.]+)'
    ]
    width_patterns = [
        r'[*\s-]*宽[（(]mm[)）][*\s]*[：:]\s*([\d.]+)',
        r'宽[：:]\s*([\d.]+)'
    ]

    lengths = []
    widths = []

    for pattern in length_patterns:
        lengths.extend(re.findall(pattern, content))
    for pattern in width_patterns:
        widths.extend(re.findall(pattern, content))

    # 组合长宽为规格格式，例如: "2440*1220", "2*1.22"等
    # 取最小长度来配对
    for i in range(min(len(lengths), len(widths))):
        length = lengths[i]
        width = widths[i]
        try:
            # 原始mm格式
            specs_info['dimensions'].add(f"{length}*{width}")
            specs_info['dimensions'].add(f"{width}*{length}")  # 反向也加上

            # 转换为米格式
            length_m = float(length) / 1000
            width_m = float(width) / 1000
            specs_info['dimensions'].add(f"{length_m:.2f}*{width_m:.2f}")
            specs_info['dimensions'].add(f"{width_m:.2f}*{length_m:.2f}")
            specs_info['dimensions'].add(f"{length_m:.1f}*{width_m:.1f}")
            specs_info['dimensions'].add(f"{width_m:.1f}*{length_m:.1f}")

            # 简化格式（四舍五入到整数米）
            length_int = round(float(length) / 1000)
            width_int = round(float(width) / 1000)
            if length_int > 0 and width_int > 0:
                specs_info['dimensions'].add(f"{length_int}*{width_int}")
                specs_info['dimensions'].add(f"{width_int}*{length_int}")
        except:
            pass

    return specs_info


def filter_data(specs_info, input_file, output_file):
    """
    根据提取的specs信息筛选compare_data.xlsx中的数据
    只要满足任一条件即可
    """
    # 读取Excel文件
    print(f"正在读取 {input_file}...")
    df = pd.read_excel(input_file)

    print(f"原始数据行数: {len(df)}")
    print(f"\n数据列名: {df.columns.tolist()}")

    # 创建筛选条件
    mask = pd.Series([False] * len(df))

    # 筛选条件1: 公司匹配
    if specs_info['companies']:
        print(f"\n提取的公司: {specs_info['companies']}")
        company_col = None
        for col in df.columns:
            if '公司' in str(col) or 'company' in str(col).lower():
                company_col = col
                break

        if company_col:
            for company in specs_info['companies']:
                company_mask = df[company_col].astype(str).str.contains(company, na=False, case=False)
                mask = mask | company_mask
                print(f"  匹配公司 '{company}': {company_mask.sum()} 条")

    # 筛选条件2: 材质/规格匹配
    if specs_info['materials']:
        print(f"\n提取的材质/规格: {specs_info['materials']}")
        # 查找材质列
        material_col = None
        for col in df.columns:
            if '材' in str(col) or 'material' in str(col).lower():
                material_col = col
                break

        if material_col:
            for material in specs_info['materials']:
                material_mask = df[material_col].astype(str).str.contains(str(material), na=False, case=False)
                mask = mask | material_mask
                print(f"  匹配材质 '{material}': {material_mask.sum()} 条")

    # 筛选条件3: 厚度匹配
    if specs_info['thicknesses']:
        print(f"\n提取的厚度: {specs_info['thicknesses']}")
        thickness_col = None
        for col in df.columns:
            if '厚' in str(col) or 'thick' in str(col).lower():
                thickness_col = col
                break

        if thickness_col:
            for thickness in specs_info['thicknesses']:
                # 精确匹配或接近匹配（±0.1范围内）
                try:
                    thickness_mask = (
                            (pd.to_numeric(df[thickness_col], errors='coerce') >= thickness - 0.1) &
                            (pd.to_numeric(df[thickness_col], errors='coerce') <= thickness + 0.1)
                    )
                    mask = mask | thickness_mask
                    print(f"  匹配厚度 {thickness}: {thickness_mask.sum()} 条")
                except:
                    pass

    # 筛选条件4: 规格/长宽匹配
    if specs_info['dimensions']:
        print(f"\n提取的长宽规格: {specs_info['dimensions']}")
        spec_col = None
        for col in df.columns:
            if '规格' in str(col) or 'spec' in str(col).lower() or '尺寸' in str(col):
                spec_col = col
                break

        if spec_col:
            for dimension in specs_info['dimensions']:
                # 匹配多种格式: "2*1", "2.44*1.22", "2440*1220"等
                dimension_mask = df[spec_col].astype(str).str.contains(
                    dimension.replace('*', r'\*'),
                    na=False,
                    case=False,
                    regex=True
                )
                mask = mask | dimension_mask
                print(f"  匹配规格 '{dimension}': {dimension_mask.sum()} 条")

    # 应用筛选
    filtered_df = df[mask]

    print(f"\n筛选后数据行数: {len(filtered_df)}")

    # 保存结果
    if len(filtered_df) > 0:
        filtered_df.to_excel(output_file, index=False)
        print(f"\n结果已保存到: {output_file}")
    else:
        print("\n警告: 没有找到匹配的数据！")

    return filtered_df


def main():
    # 文件路径
    specs_file = 'specs.txt'
    input_file = 'compare_data.xlsx'
    output_file = 'filtered_specs_result.xlsx'

    print("=" * 60)
    print("开始处理specs.txt并筛选compare_data.xlsx")
    print("=" * 60)

    # 第一步: 解析specs.txt
    print(f"\n[步骤1] 解析 {specs_file}...")
    specs_info = parse_specs_txt(specs_file)

    print(f"\n提取的信息汇总:")
    print(f"  公司数: {len(specs_info['companies'])}")
    print(f"  材质/规格数: {len(specs_info['materials'])}")
    print(f"  厚度数: {len(specs_info['thicknesses'])}")
    print(f"  长宽规格数: {len(specs_info['dimensions'])}")

    # 第二步: 筛选数据
    print(f"\n[步骤2] 在 {input_file} 中筛选匹配数据...")
    filtered_df = filter_data(specs_info, input_file, output_file)

    print("\n" + "=" * 60)
    print("处理完成!")
    print("=" * 60)

    # 显示前几行结果
    if len(filtered_df) > 0:
        print("\n筛选结果预览 (前5行):")
        print(filtered_df.head().to_string())


if __name__ == '__main__':
    main()

