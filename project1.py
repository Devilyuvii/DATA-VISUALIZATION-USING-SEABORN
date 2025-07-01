import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import os
from glob import glob
import openpyxl

# read and clean excel files:
def read_and_clean_files(path):
    files=glob(os.path.join(path,"*.xls"))+ glob(os.path.join(path,"*.xlsx"))
    combined_df=pd.DataFrame()

    for i in files:
        df=pd.read_excel(i)
        df.drop_duplicates(inplace=True)
        combined_df=pd.concat([combined_df,df],ignore_index=True)

    combined_df.drop_duplicates(inplace=True) 
    return combined_df

def export_to_excel(df,output_file):
    df.to_excel(output_file, index=False)
    print(f"cleaned and combined files saved to:{output_file}")

# coding for visualizaing the cleaned excel file:
def visualize_data(excel_file_path):
    sns.set(style="whitegrid")
    # step1: read the already cleaned file
    df=pd.read_excel(excel_file_path)
    
    # step2: lineplot :a vs b( a and b are the columns)
    plt.figure(figsize=(8, 5))   # 8=width and 5=height
    sns.lineplot(x="a",y="b",data=df,marker="o")
    plt.title("lineplot:column a vs column b")
    plt.xlabel(" column a")
    plt.ylabel(" column b")
    plt.tight_layout()
    plt.show()

    # step3: scatterplot :a vs b( a and b are the columns)
    plt.figure(figsize=(8, 5))   # 8=width and 5=height
    sns.scatterplot(x="a",y="b",data=df)
    plt.title("scatterplot:column a vs column b")
    plt.xlabel(" column a")
    plt.ylabel(" column b")
    plt.tight_layout()
    plt.show()

    # step4: heatmap :a vs b( a and b are the columns)
    plt.figure(figsize=(8, 5))   # 8=width and 5=height
    corr=df[["a","b"]].corr()
    sns.heatmap(corr,annot =True, cmap="coolwarm",fmt=".2f")
    plt.title("heatmap:correlation between column a and column b")
    plt.tight_layout()
    plt.show()


def main():
    path=input("enter the folder path containing excel files: ").strip()
    output_file=input("enter full output file path to save the cleaned excel: ").strip()

    if not os.path.isdir(path):
        print(" invalid path")
        return
    
    # call all the functons used above:
    # step1: call read and cleaned files

    cleaned_data= read_and_clean_files(path)

# step2:  call export to user specified file path
    
    export_to_excel(cleaned_data,output_file)

# step3: call visualize the data directly from Saved file

    visualize_data(output_file)

# step4: call main function :

if __name__ == '__main__': 
    main()