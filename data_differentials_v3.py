# %%
from os.path import dirname, join
import pandas as pd
from datetime import datetime
import numpy as np

project_path = dirname(__file__)
input_path = join(project_path,'input')
output_path = join(project_path,'output')
file_curr = join(input_path,'source_B.xlsx')
df_curr=pd.read_excel(file_curr)


## You can specify your index here -- this sets the basis for comparison i.e. newly added rows, or removed.
index_col = df_curr.columns[1]
file_prev = join(input_path,'source_A.xlsx')
df_prev=pd.read_excel(file_prev, index_col=index_col).fillna(0)
df_curr=pd.read_excel(file_curr, index_col=index_col).fillna(0)



#@       >>> preparing output dataframe to show the differences i.e. from old value to new
df_diff = df_curr.copy()
removed_rows = []
added_rows = []
common_rows = []

cols_prev = df_prev.columns
cols_curr = df_curr.columns
cols_common = list(set(cols_prev).intersection(cols_curr))

for row in df_diff.index:
    if (row in df_prev.index) and (row in df_curr.index):
        print('existing row -- {}'.format(row))
        common_rows.append(row)
    else:
        print('added row -- {}'.format(row))
        added_rows.append(row)


for row in df_prev.index:
    if row not in df_curr.index:
        print('removed row -- {}'.format(row))
        removed_rows.append(row)
        df_diff = df_diff.append(df_prev.loc[row,:])
df_diff = df_diff.sort_index().fillna('')

df_added_rows = df_curr.loc[df_curr.index.isin(added_rows)]
df_removed_rows = df_prev.loc[df_prev.index.isin(removed_rows)]




df_prev_common = df_prev.loc[df_prev.index.isin(common_rows)]
df_curr_common = df_curr.loc[df_curr.index.isin(common_rows)]
df_prev_common.equals(df_curr_common)
cmp_val = df_prev_common.values == df_curr_common.values
rows,cols=np.where(cmp_val==False)
#@       >>> preparing output dataframe to show the differences i.e. from old value to new
df_diff_common = df_prev_common.copy()
for item in zip(rows,cols):
    df_diff_common.iloc[item[0], item[1]] = '{} --> {}'.format(df_prev_common.iloc[item[0], item[1]],df_curr_common.iloc[item[0], item[1]])




# %%    >>> setting up excel to generate report 
output_file = 'dataset_differentials_' + datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p") + '.xlsx'
output_file = join(output_path,output_file)

writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
df_diff_common.to_excel(writer, sheet_name='conmmon', index=True)
df_added_rows.to_excel(writer, sheet_name='added', index=True)
df_removed_rows.to_excel(writer, sheet_name='deleted', index=True)


workbook  = writer.book
worksheet = writer.sheets['diff']
worksheet.hide_gridlines(2)
worksheet.set_default_row(15)

#@      >>> this highlights the changed cells / values
highlighter_fmt = workbook.add_format({'font_color': '#FF0000'
    , 'italic':True
    , 'bg_color':'#FFFFCC'})

added_fmt = workbook.add_format({'font_color': '#09890F'
    , 'bold':True
    ,  'bg_color':'#70F676'})

removed_fmt = workbook.add_format({'font_color': '#600404'
    ,  'bg_color':'#FBABAB'})





worksheet.conditional_format('A1:ZZ1000000', {'type': 'text',
    'criteria': 'containing',
    'value':'-->',
    'format': highlighter_fmt})


# #     added / removed rows' highlight
# for rownum in range(df_diff.shape[0]):
#     row = df_diff.index[rownum]
#     for x in added_rows:
#         if x == row:
#             worksheet.set_row(rownum + 1, 15, added_fmt)
#     for y in removed_rows:
#         if y == row:
#             worksheet.set_row(rownum + 1, 15, removed_fmt)
    

writer.save()
print('differential report generated!')