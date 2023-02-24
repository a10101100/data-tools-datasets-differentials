# %%
from os.path import dirname, join
import pandas as pd
from datetime import datetime
import numpy as np

start_time = datetime.now() 


project_path = dirname(__file__)
input_path = join(project_path,'input')
output_path = join(project_path,'output')
file_curr = join(input_path,'source_B.xlsx')
df_curr=pd.read_excel(file_curr)

KEY_COLNAME = 'Country'


## You can specify your index here -- this sets the basis for comparison i.e. newly added rows, or removed.
file_prev = join(input_path,'source_A.xlsx')
df_prev=pd.read_excel(file_prev).fillna(0)
df_curr=pd.read_excel(file_curr).fillna(0)

df_added = df_curr[~df_curr[KEY_COLNAME].isin(df_prev[KEY_COLNAME])]
df_removed = df_prev[~df_prev[KEY_COLNAME].isin(df_curr[KEY_COLNAME])]

#@       >>> preparing output dataframe to show the differences i.e. from old value to new
df_prev.set_index(KEY_COLNAME, inplace=True)
df_curr.set_index(KEY_COLNAME, inplace=True)
df_prev_common = df_prev[df_prev.index.isin(df_curr.index)]
df_curr_common = df_curr[df_curr.index.isin(df_prev.index)]
df_prev_common.equals(df_curr_common)
cmp_val = df_prev_common.values == df_curr_common.values
rows,cols=np.where(cmp_val==False)

df_diff_common = df_prev_common.copy()
for item in zip(rows,cols):
    df_diff_common.iloc[item[0], item[1]] = '{} --> {}'.format(df_prev_common.iloc[item[0], item[1]],df_curr_common.iloc[item[0], item[1]])




# %%    >>> setting up excel to generate report 
output_file = 'dataset_differentials_' + datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p") + '.xlsx'
output_file = join(output_path,output_file)

writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
df_diff_common.to_excel(writer, sheet_name='common_changes', index=False)
df_added.to_excel(writer, sheet_name='added', index=False)
df_removed.to_excel(writer, sheet_name='deleted', index=False)


workbook  = writer.book
worksheet = writer.sheets['common_changes']
worksheet.set_tab_color('#FF9900')  # Orange
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

added_sht = writer.sheets['added']
added_sht.set_tab_color('green')

removed_sht = writer.sheets['deleted']
removed_sht.set_tab_color('red')

writer.save()
print('differential report generated!')
time_elapsed = datetime.now() - start_time 
print('Time elapsed (hh:mm:ss.ms) {}'.format(time_elapsed))