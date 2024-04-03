import pandas as pd
import os

# CSV文件所在的目录路径
directory = '学校信息/海南师范大学/'

# 遍历目录下的所有文件
for filename in os.listdir(directory):
    # 检查文件是否为CSV文件
    if filename.endswith(".csv"):
        parts = filename.split("-")
        year = parts[0]
        province = parts[1]
        school = parts[2].split(".")[0]  # 去掉文件扩展名
        wenjian = year + '-' + province + '-' + school
        all_data = []
        # 读取CSV文件的每个工作表
        data_frames = pd.read_excel(os.path.join(directory, filename), sheet_name=None,engine='openpyxl')
        
        for sheet_name, df in data_frames.items():
            # 添加来源标识列
            df['工作表'] = sheet_name.split('_')[0]
            df['工作表2'] = sheet_name.split('_')[1]
            # 添加到当前文件的所有数据列表中
            all_data.append(df)
        # 合并每个工作表的数据
        combined_data = pd.concat(data_frames, ignore_index=True)
        
        # 定义旧表头和新表头之间的映射关系
        header_mapping = {
            "专业": "录取专业名称",
            "最低分": "最低分",
            "最高分": "最高分",
            "最低分排名": "最低分排名",
            "最高分排名": "最高分排名",
            "平均分": "平均分",
            "录取人数": "录取数量",
            "控制线": "控制线",
            "选考或备注": "备注",
        }
        
        # 定义新的表头
        new_headers = [
            "录取年份", "录取专业名称", "省份", "录取院校名称", "文理科", "类型", "批次名称", "录取数量",
            "最低分", "最低分排名", "平均分", "平均分排名", "最高分", "最高分排名", "控制线", "备注", "数据来源", "计划数"
        ]
        
        # 更改表头
        new_data = combined_data.rename(columns=header_mapping)
        
        # 添加缺失的列并赋予固定值
        for new_header in new_headers:
            if new_header not in new_data.columns:
                new_data[new_header] = ''
        new_data['录取院校名称'] = '海南师范大学'
        new_data['数据来源'] = '院校官网'
        new_data['录取年份'] = f'{year}'
        new_data['省份'] = f'{province}'
        # 获取工作表名字的一部分作为"校区"
        new_data['文理科'] = new_data['工作表']
        new_data['类型'] = new_data['工作表2']


        # 删除工作表列
        new_data.drop(columns=['工作表'], inplace=True)
        new_data.drop(columns=['工作表2'], inplace=True)
        # 重新排列列的顺序
        new_data = new_data[new_headers]
        
        # 存储合并后的数据到CSV文件
        new_directory = f'学校信息第二版/{school}/{year}/{province}/'
        os.makedirs(new_directory, exist_ok=True)
        
        # 存储合并后的数据到CSV文件
        new_filename = os.path.join(new_directory, f'{wenjian}.csv')
        new_data.to_csv(new_filename, index=False)
        print(f"数据已保存到 {new_filename}")

