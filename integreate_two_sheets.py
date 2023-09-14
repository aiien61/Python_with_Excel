import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

df_1 = pd.read_excel('data/TimeseriesData.xlsx', sheet_name='Sheet1')
df_2 = pd.read_excel('data/TimeseriesData.xlsx', sheet_name='Sheet2')

plt.plot(df_2['Time'].astype(str), df_2['Amount Produced'], label='Amount Produced')
plt.show()