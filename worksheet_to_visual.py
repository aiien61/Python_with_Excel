import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

excel_file = "./data/SalesData.xlsx"


def automate_report(excel_file):
    df = pd.read_excel(excel_file)
    
    condition1 = (df['Delivered'] == 'No') & (df['Product'] == 'Iphone 7')
    condition2 = (df['Delivered'] == 'No') & (df['Product'] == 'Iphone 8')
    df_unreceived = df[condition1 | condition2]
    df_unreceived.to_excel('output/unreceived-iphones-iphone7-8.xlsx')

    df_commission = df.groupby(['Employee For Commission']).sum()['Commission']
    df_commission.to_excel('output/Commission.xlsx')

    df_total_sales = df.groupby(['Product']).sum()['Value']
    
    df_total_sales.plot(kind='bar')
    plt.title('Value of Iphone Sales')
    plt.xlabel('Iphone Series')
    plt.ylabel('Total Value of Sale')
    plt.xticks(rotation=0, ha='right')
    plt.show()
    return None


if __name__ == "__main__":
    automate_report(excel_file)