import os
import csv

def normalize_location(location):
    # 对地区名称进行归一化，将市、自治区、直辖市、特别行政区等归一化为省份
    normalized_location = location.replace("市", "").replace("自治区", "").replace("壮族", "").replace("回族", "").replace("维吾尔", "").replace("班", "").replace("联招", "").replace("侨", "").strip()
    return normalized_location

def normalize_province(province):
    # 对省份名称进行归一化，例如将 "安徽省" 归一化为 "安徽"
    normalized_province = province.replace("省", "").strip()
    return normalized_province

def merge_csv_files(input_dir, output_dir, years):
    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)

    # 创建一个字典来存储同一省份、同一年份的数据
    data_dict = {}

    # 定义表头
    header = ["录取年份", "录取专业名称", "省份", "录取院校名称", "文理科", "类型", "批次名称", "录取数量",
              "最低分", "最低分排名", "平均分", "平均分排名", "最高分", "最高分排名", "控制线", "备注",
              "数据来源", "计划数", "校区", "学制","专业组","民族","学院","授予学位","我校投档线","录取线"]

    # 遍历输入目录中的所有文件
    for root, dirs, files in os.walk(input_dir):
        for filename in files:
            # 检查文件名中是否包含在指定年份列表中
            file_year = filename.split('-')[0]
            if file_year not in years:
                continue
            
            # 读取CSV文件中的数据
            filepath = os.path.join(root, filename)
            with open(filepath, 'r', newline='', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    province = row['省份']
                    year = row['录取年份']
                    # 如果年份不在指定年份列表中，则跳过
                    if year not in years:
                        continue
                    # 对省份名称和地区名称进行归一化
                    normalized_province = normalize_province(province)
                    normalized_location = normalize_location(normalized_province)
                    # 构建省份目录和文件路径
                    province_dir = os.path.join(output_dir, year, normalized_location)
                    os.makedirs(province_dir, exist_ok=True)
                    output_filepath = os.path.join(province_dir, f"{year}-{normalized_location}.csv")
                    # 将数据添加到字典中的相应键的列表中
                    if output_filepath not in data_dict:
                        data_dict[output_filepath] = []
                    data_dict[output_filepath].append(row)

    # 遍历字典，将数据写入到合并后的CSV文件中
    for output_filepath, data_list in data_dict.items():
        with open(output_filepath, 'w', newline='', encoding='utf-8') as file:
            # 使用 DictWriter 写入数据，确保表头顺序一致
            writer = csv.DictWriter(file, fieldnames=header)
            writer.writeheader()
            for data in data_list:
                # 检查每个字段是否存在，如果不存在则填充为空值
                for field in header:
                    if field not in data:
                        data[field] = ''
                writer.writerow(data)

# 示例用法：只合并 2021、2022 和 2023 年的数据
input_dir = '学校信息第二版'
output_dir = '学校信息第三版'
years = ['2021', '2022', '2023']
merge_csv_files(input_dir, output_dir, years)
