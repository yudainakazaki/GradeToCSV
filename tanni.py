import tabula
import pandas as pd
from PyPDF2 import PdfFileReader

def count_page(file):
    file = file
    pdf = PdfFileReader(open(file, 'rb'))
    res = pdf.getNumPages()
    return res


def make_data(file):
    df = pd.DataFrame()
    for page in range (1, count_page(file)):
        temp = pd.DataFrame()
        if page == 1:
            frag = tabula.read_pdf(file, multiple_tables=False, pages=page, stream=True, area=[330, 30, 760, 580], columns=[190, 290, 340, 380, 435, 480, 530])
            temp = frag[0]
        else:
            frag = tabula.read_pdf(file, multiple_tables=False, pages=page, area=[25, 30, 760, 580], columns=[195, 290, 340, 380, 435, 480, 530])
            temp = frag[0]
        df = pd.concat([df, temp], axis='rows')
    df = df.reset_index(drop=True)
    df = df.fillna("nan")
    return df


def fixrow(df):
    for index, item in df.iterrows():
        if index > 0:
            try:
                if item["科目名称"] == "nan" and df.at[index-1, "担当者"] == "nan" and df.at[index+1, "担当者"] == "nan":
                    item["科目名称"] = df.at[index-1, "科目名称"] + df.at[index+1, "科目名称"]
                    df.at[index-1, "科目名称"] = "nan"
                    df.at[index+1, "科目名称"] = "nan"
                # print(index)
                # print(df.at[index-1, "担当者"])
            except:
                pass

    for index, item in df.iterrows():
        if item["科目名称"] == "nan" and item["担当者"] == "nan":
            df.drop(index, inplace=True)
            
    return df


def category(df):
    df["分野"] = ""
    curr_name = ""
    for index, item in df.iterrows():
        if str(item["科目名称"][:2]) == "分野":
            item["科目名称"] = item["科目名称"] + str(item["担当者"]) + str(item["評語"]) + str(item["単位"])
            item["科目名称"] = item["科目名称"].replace("nan", "")
            curr_name = item["科目名称"]
        else: 
            df.at[index, "分野"] = curr_name
    
    for index, item in df.iterrows():
        try:
            if item["科目名称"][:2] == "分野":
                df.drop(index, inplace=True)
        except:
            pass

    df = df.reset_index(drop=True) 
    df = df.replace("nan", "")
    return df 



file = "grade.pdf"

print("Executing...")
df = make_data(file)
df = fixrow(df)
df = category(df)
df.to_csv("grade.csv")
