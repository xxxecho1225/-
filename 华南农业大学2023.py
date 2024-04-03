import pandas as pd
import os

# CSV文件所在的目录路径
directory = '学校信息/华南农业大学/2021贵州/'

# 遍历目录下的所有文件
for filename in os.listdir(directory):
    # 检查文件是否为CSV文件
    if filename.startswith("2021") and filename.endswith(".csv"):
        parts = filename.split("-")
        year = parts[0]
        province = parts[1]
        school = parts[2].split(".")[0]  # 去掉文件扩展名
        wenjian = year + '-' + province + '-' + school
        
        # 读取CSV文件
        data = pd.read_excel(os.path.join(directory, filename), sheet_name=None,engine='openpyxl')
        combined_data = pd.concat(data, ignore_index=True)
        # 定义旧表头和新表头之间的映射关系
        header_mapping = {
            "录取专业": "录取专业名称",
            "科类": "文理科",
            "专业最低": "最低分",
            "省划线": "控制线",
            "我校投档线": "我校投档线",
        }
        
        # 定义新的表头
        new_headers = [
            "录取年份", "录取专业名称", "省份", "录取院校名称", "文理科", "类型", "批次名称", "录取数量",
            "最低分", "最低分排名","平均分", "平均分排名", "最高分", "最高分排名", "控制线", "备注", "数据来源",
            "计划数","校区","学制","专业组","授予学位","我校投档线"
        ]

        # 更改表头
        new_data = combined_data.rename(columns=header_mapping)

        # 添加缺失的列并赋予固定值
        new_data['录取年份'] = f'{year}'
        new_data['省份'] = f'{province}'
        new_data['录取院校名称'] = '华南农业大学'
        new_data['数据来源'] = '院校官网'
        new_data['计划数'] = ''
        new_data['批次名称'] = ''
        new_data['学制'] = ''
        new_data['备注'] = ''
        new_data['类型'] = ''
        new_data['控制线'] = ''
        new_data['校区'] = ''
        new_data['最高分排名'] = ''
        new_data['平均分排名'] = ''
        new_data['最低分排名'] = ''
        new_data['最高分'] = ''
        new_data['平均分'] = ''
        new_data['专业组'] = ''
        new_data['授予学位'] = ''
        new_data['录取数量'] = ''



        new_data = new_data[new_headers]
        

        # 存储合并后的数据到CSV文件
        new_directory = os.path.join(f'学校信息第二版/{school}/{year}/{province}/')
        os.makedirs(new_directory, exist_ok=True)

        # 生成文件路径
        new_filename = os.path.join(new_directory, f'{wenjian}.csv')

        # 存储合并后的数据到CSV文件
        new_data.to_csv(new_filename, index=False)
        print(f"数据已保存到 {new_filename}")
