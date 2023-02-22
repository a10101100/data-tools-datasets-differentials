# %%
from os.path import dirname, join
import pandas as pd
import numpy as np
from datetime import datetime

project_path = dirname(__file__)
input_path = join(project_path,'input')
output_path = join(project_path,'output')

file_1 = join(input_path,'source_A.xlsx')
df1=pd.read_excel(file_1)
file_2 = join(input_path,'source_B.xlsx')
df2=pd.read_excel(file_2)

# %%    >>> comparing the 2 datasets, and getting the difference in a same dimensional dataframe
df1.equals(df2)
cmp_val = df1.values == df2.values
rows,cols=np.where(cmp_val==False)


#@       >>> preparing output dataframe to show the differences i.e. from old value to new
df_diff = df1.copy()
for item in zip(rows,cols):
    df_diff.iloc[item[0], item[1]] = '{} --> {}'.format(df1.iloc[item[0], item[1]],df2.iloc[item[0], item[1]])



# %%    >>> setting up excel to generate report 
output_file = 'dataset_differentials_' + datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p") + '.xlsx'
output_file = join(output_path,output_file)

writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
df_diff.to_excel(writer, sheet_name='diff', index=False)
df1.to_excel(writer, sheet_name='prev', index=False)
df2.to_excel(writer, sheet_name='curr', index=False)


workbook  = writer.book
worksheet = writer.sheets['diff']
# worksheet.hide_gridlines(2)

#@      >>> this highlights the changed cells / values
highlighter_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color':'#FFFF00'})
worksheet.conditional_format('A1:ZZ1000000', {'type': 'text',
    'criteria': 'containing',
    'value':'-->',
    'format': highlighter_fmt})
writer.save()
print('differential report generated!')
# %%
