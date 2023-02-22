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

# %%
df1.equals(df2)
cmp_val = df1.values == df2.values
rows,cols=np.where(cmp_val==False)

for item in zip(rows,cols):
    df1.iloc[item[0], item[1]] = '{} --> {}'.format(df1.iloc[item[0], item[1]],df2.iloc[item[0], item[1]])

# %%
output_file = 'dataset_differentials_' + datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p") + '.xlsx'
output_file = join(output_path,output_file)
df1.to_excel(output_file,index=False,header=True)



# %%
