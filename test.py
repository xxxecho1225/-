import os

def rename_files(directory):
    for filename in os.listdir(directory):
        if filename.endswith(".csv"):
            old_path = os.path.join(directory, filename)
            new_filename = filename.replace("中国矿业大学", "中国矿业大学(北京)")
            new_path = os.path.join(directory, new_filename)
            os.rename(old_path, new_path)
            print(f"Renamed: {old_path} -> {new_path}")

# 指定包含CSV文件的目录路径
directory_path = "D:/git_python/school/数据分析/学校分数信息/src/学校信息/中国矿业大学(北京)"

rename_files(directory_path)
