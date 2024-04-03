import os

# CSV文件所在的目录路径
directory = '学校信息/北京语言大学/'

# 遍历目录下的所有文件
for filename in os.listdir(directory):
    # 检查文件是否为CSV文件
    if filename.endswith(".csv"):
        parts = filename.split("-")
        year = parts[0]
        province = parts[1]
        school = parts[2].split(".")[0]  # 去掉文件扩展名
        wenjian = year +'-'+province+'-'+school
        print(f"{wenjian}")
        # 输出结果
        # 输出结果
        #print(f"年份: {year}, 省份: {province}, 学校: {school}")


