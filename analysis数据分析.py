import os
import pandas as pd
import matplotlib.pyplot as plt

#设置文件路径
INPUT_FILE="data/student_score_dirty.xlsx"
OUTPUT_DIR="output"
#创建输出文件夹
os.makedirs(OUTPUT_DIR,exist_ok=True)
#设置中文字体
plt.rcParams["font.sans-serif"]=["Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"]=False
#1.读取数据
def load_data():
    df=pd.read_excel(INPUT_FILE)
    print("原始数据读取成功！")
    print("原始数据行数：",len(df))
    return df
#2.清洗数据
def clean_data(df):
    print("\n开始清洗数据...")
    df=df.drop_duplicates()
    subjects=["数学","英语","python"]
    for subject in subjects:
        df[subject]=df[subject].fillna(df[subject].mean())
    df["班级"]=df["班级"].astype(str).fillna("未知")
    print("数据清洗完成！")
    return df
#3.计算学生成绩
def calculate_scores(df):
    df["总分"]=df[["数学","英语","python"]].sum(axis=1)
    df["平均分"]=df[["数学","英语","python"]].mean(axis=1)
    df=df.sort_values(by="总分",ascending=False)
    print("\n学生成绩计算完成！")
    return df
#4.统计班级成绩
def class_statistics(df):
    res=df.groupby("班级")[["数学","英语","python"]].mean()
    res=res.round(2)
    res.to_excel(os.path.join(OUTPUT_DIR,"班级成绩统计表.xlsx"),engine="openpyxl",index=False)
    print("班级成绩统计表已保存！")
    return res
#5.统计各科成绩
def subject_statistics(df):
    res=df[["数学","英语","python"]].agg(["mean","max","min"]).round(2)
    res.to_excel(os.path.join(OUTPUT_DIR,"各科成绩统计表.xlsx"),engine="openpyxl",index=False)
    print("各科成绩统计表已保存！")
    return res
#6.绘制班级数学平均成绩图
def draw_chart(df):
    res=df.groupby("班级")["数学"].mean()
    plt.figure(figsize=(8,5))
    plt.bar(res.index,res.values)
    plt.title("各班数学平均分")
    plt.xlabel("班级")
    plt.ylabel("平均成绩")
    plt.savefig(os.path.join(OUTPUT_DIR,"班级数学成绩图.png"),dpi=300,bbox_inches="tight")
    plt.close()
    print("成绩图表已保存！")
#7.主程序
def main():
    df=load_data()
    df=clean_data(df)
    df=calculate_scores(df)
    df.to_excel(os.path.join(OUTPUT_DIR,"学生成绩分析结果.xlsx"),engine="openpyxl",index=False)
    class_result=class_statistics(df)
    subject_result=subject_statistics(df)
    draw_chart(df)
    print("\n================成绩分析完成=================")
    print("\n学生成绩排名：")
    print(df[["姓名","班级","总分","平均分"]])
    print("\n班级成绩统计：")
    print(class_result)
    print("\n各科成绩统计：")
    print(subject_result)

if __name__=="__main__":
    main()

