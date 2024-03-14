# Copyright (C) 2024 Nardo Kong


import pandas as pd
from tqdm import tqdm

print("© 2024 Nardo Kong. All rights reserved.")
print("请确认data与本程序在同一文件夹中, 且没有其他名为result.xlsx的文件, 否则将会被覆盖")
input("按回车键继续...")

# Read data.xlsx, must ensure that the file name correct
df = pd.read_excel('data.xlsx', sheet_name='Sheet1')

# 获取列名列表
columns = df.columns.tolist()

# 重命名第4至51列为1至48
new_columns = {columns[i]: i - 2 for i in range(3, 51)}  # 创建一个映射字典，将旧列名映射到新列名
df.rename(columns=new_columns, inplace=True)

# 假设您的数据框架已经准备好，并且每个题目的回答都标记为1（是）或0（否）
# 以下是每个量表的题目编号，其中包括正向记分和反向记分的题目
P_positive = [10, 14, 22, 31, 39]
P_negative = [2, 6, 18, 26, 28, 35, 43]
E_positive = [3, 7, 11, 15, 19, 23, 32, 36, 41, 44, 48]
E_negative = [27]
N_positive = [1, 5, 9, 13, 17, 21, 25, 30, 34, 38, 42, 46]
L_positive = [4, 16, 45]
L_negative = [8, 12, 20, 24, 29, 33, 37, 40, 47]

# 常模表数据
norms = {
    'male': {
        '16-19': {'P': {'M': 3.15, 'SD': 1.82}, 'E': {'M': 7.74, 'SD': 2.77}, 'N': {'M': 4.70, 'SD': 2.96}, 'L': {'M': 4.43, 'SD': 2.55}},
        '20-29': {'P': {'M': 3.00, 'SD': 2.00}, 'E': {'M': 8.05, 'SD': 2.67}, 'N': {'M': 4.57, 'SD': 3.06}, 'L': {'M': 4.90, 'SD': 2.66}},
        '30-39': {'P': {'M': 2.88, 'SD': 2.04}, 'E': {'M': 7.82, 'SD': 2.68}, 'N': {'M': 4.01, 'SD': 2.78}, 'L': {'M': 5.61, 'SD': 2.66}},
        '40-49': {'P': {'M': 2.91, 'SD': 2.34}, 'E': {'M': 7.34, 'SD': 2.88}, 'N': {'M': 4.34, 'SD': 2.95}, 'L': {'M': 6.55, 'SD': 2.78}},
        '50-59': {'P': {'M': 2.67, 'SD': 2.21}, 'E': {'M': 6.95, 'SD': 2.98}, 'N': {'M': 3.90, 'SD': 2.89}, 'L': {'M': 7.19, 'SD': 2.66}},
        '60-69': {'P': {'M': 2.68, 'SD': 2.31}, 'E': {'M': 7.08, 'SD': 3.01}, 'N': {'M': 3.70, 'SD': 3.00}, 'L': {'M': 7.73, 'SD': 3.08}},
        '70+'  : {'P': {'M': 2.92, 'SD': 2.79}, 'E': {'M': 6.89, 'SD': 3.08}, 'N': {'M': 4.38, 'SD': 3.39}, 'L': {'M': 8.00, 'SD': 3.13}},
    },
    'female': {
        '16-19': {'P': {'M': 2.63, 'SD': 1.81}, 'E': {'M': 8.13, 'SD': 2.58}, 'N': {'M': 4.93, 'SD': 2.75}, 'L': {'M': 4.86, 'SD': 2.43}},
        '20-29': {'P': {'M': 2.68, 'SD': 1.82}, 'E': {'M': 7.44, 'SD': 2.79}, 'N': {'M': 4.81, 'SD': 2.95}, 'L': {'M': 5.32, 'SD': 2.70}},
        '30-39': {'P': {'M': 2.44, 'SD': 1.82}, 'E': {'M': 7.50, 'SD': 2.87}, 'N': {'M': 4.49, 'SD': 2.89}, 'L': {'M': 6.64, 'SD': 2.76}},
        '40-49': {'P': {'M': 2.55, 'SD': 2.30}, 'E': {'M': 7.15, 'SD': 2.86}, 'N': {'M': 4.44, 'SD': 2.95}, 'L': {'M': 7.45, 'SD': 2.98}},
        '50-59': {'P': {'M': 2.36, 'SD': 1.82}, 'E': {'M': 6.92, 'SD': 2.90}, 'N': {'M': 4.48, 'SD': 2.88}, 'L': {'M': 7.73, 'SD': 2.68}},
        '60-69': {'P': {'M': 2.51, 'SD': 1.98}, 'E': {'M': 7.28, 'SD': 2.95}, 'N': {'M': 4.44, 'SD': 3.12}, 'L': {'M': 7.72, 'SD': 2.96}},
        '70+'  : {'P': {'M': 2.32, 'SD': 1.89}, 'E': {'M': 7.28, 'SD': 3.48}, 'N': {'M': 4.88, 'SD': 3.25}, 'L': {'M': 8.84, 'SD': 2.58}},
    }
}

# 将年龄转换为年龄组
def age_to_group(age):
    if 16 <= age <= 19:
        return '16-19'
    elif 20 <= age <= 29:
        return '20-29'
    elif 30 <= age <= 39:
        return '30-39'
    elif 40 <= age <= 49:
        return '40-49'
    elif 50 <= age <= 59:
        return '50-59'
    elif 60 <= age <= 69:
        return '60-69'
    elif age >= 70:
        return '70+'
    else:
        raise ValueError('Age does not match any group')


# 计算每个量表的原始分数
def calculate_raw_score(df, positive_items, negative_items):
    # 正向记分题目的分数
    positive_score = df[positive_items].applymap(lambda x: 1 if x == "是" else 0).sum(axis=1)
    # 反向记分题目的分数（反向题目的“是”为0分，“否”为1分）
    negative_score = df[negative_items].applymap(lambda x: 1 if x == "否" else 0).sum(axis=1)
    # 总分为正向题目和反向题目的分数之和
    return positive_score + negative_score

# 应用计分函数
df['P_raw'] = calculate_raw_score(df, P_positive, P_negative)
df['E_raw'] = calculate_raw_score(df, E_positive, E_negative)
df['N_raw'] = calculate_raw_score(df, N_positive, []).astype(int) # N量表没有反向记分题目
df['L_raw'] = calculate_raw_score(df, L_positive, L_negative)

# 创建一个进度条
pbar = tqdm(total=len(df) * len(['P', 'E', 'N', 'L']))

# 将原始分数转换为T分数的函数
def convert_to_t_score(row, factor):
    gender = 'male' if row['性别'] == '男' else 'female'
    age = row['您的年龄']
    age_group = age_to_group(age)
    norm = norms[gender][age_group][factor]
    raw_score = row[factor + '_raw']

    # 更新进度条
    pbar.update()
    return 50 + 10 * (raw_score - norm['M']) / norm['SD']

# 应用T分数转换函数
for factor in ['P', 'E', 'N', 'L']:
    df[factor + '_t_score'] = df.apply(lambda row: convert_to_t_score(row, factor), axis=1)

# 完成后关闭进度条
pbar.close()

# 保存结果
df.to_excel('result.xlsx', index=False)