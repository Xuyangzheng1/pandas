#此代码被用于提取明文身份证中的用户年龄并按年龄大小进行分段
import pandas as pd
import numpy as np
from openpyxl import load_workbook
def calculate_age(id_card, current_year=2023):
    birth_year_str = id_card[6:10]
    
    if birth_year_str.isdigit():
        birth_year = int(birth_year_str)
        age=current_year - birth_year
        if age>110 or age<0:
            return None
        else:
            return age
    else:
        return None

# def filter_and_add_age(input_df):
#     filtered_data = input_df[(input_df['统计年度'] == 2021) & (input_df['险类名称'] == '种植保险类')]
    
#     filtered_data['年龄'] = filtered_data['证件号'].apply(calculate_age)
    
#     return filtered_data
def filter_and_add_age(input_df):
    filtered_data = input_df[(input_df['统计年度'] == 2021) & (input_df['险类名称'] == '种植保险类')].copy()
    
    filtered_data['年龄'] = filtered_data['证件号'].apply(calculate_age)
    
    return filtered_data

def add_age_segmentation(dfYear):
    conditions = [
        (dfYear['年龄'] >= 11) & (dfYear['年龄'] <= 20),
        (dfYear['年龄'] >= 21) & (dfYear['年龄'] <= 30),
        (dfYear['年龄'] >= 31) & (dfYear['年龄'] <= 40),
        (dfYear['年龄'] >= 41) & (dfYear['年龄'] <= 50),
        (dfYear['年龄'] >= 51) & (dfYear['年龄'] <= 60),
        (dfYear['年龄'] >= 61) & (dfYear['年龄'] <= 70),
        (dfYear['年龄'] >= 71) & (dfYear['年龄']<= 80),
        (dfYear['年龄'] >= 81) & (dfYear['年龄']<= 90),
        (dfYear['年龄'] < 21) | (dfYear['年龄'] > 90)
    ]
    choices = [
        "11-20",
        "21-30",
        "31-40",
        "41-50",
        "51-60",
        "61-70",
        "71-80",
        "81-90",
        "超出范围"
    ]
    dfYear['年龄分段'] = np.select(conditions, choices)
    return dfYear['年龄分段']

import pandas as pd
import numpy as np

# 定义 calculate_age 函数，filter_and_add_age 函数，add_age_segmentation 函数

def main():
    excel_file = r'C:\Users\DELL\Desktop\pandas\种植险大户清单-0727(1).xlsx'
    
    df = pd.read_excel(excel_file, sheet_name="Sheet1")
    
    # 尝试读取带有年龄字段的工作表，如果不存在就创建一个新的
    try:
        dfYear = pd.read_excel(excel_file, sheet_name='-添加了年龄字段')
    except:
        dfYear = pd.DataFrame()
    
    # 使用条件筛选数据并添加年龄字段
    filtered_df = filter_and_add_age(df)
    
    # 调用 add_age_segmentation 函数并将结果保存到新的列
    filtered_df['年龄分段'] = add_age_segmentation(filtered_df)
    
    # 打开 Excel 写入器，将筛选和分段结果写入工作表
    new_sheet_name = '-添加了年龄分段'
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
        if new_sheet_name in pd.ExcelFile(excel_file).sheet_names:
            writer.book = load_workbook(excel_file)
            writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
            del writer.sheets[new_sheet_name]
            writer.book.save(excel_file)
        filtered_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
    
    print("筛选和分段结果写入完成")

if __name__ == "__main__":
    main()
