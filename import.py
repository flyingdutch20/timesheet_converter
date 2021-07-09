import pandas as pd
import os

my_dir = "reqlogic"
my_cols = "A,B,H,M,T,U,V,W,Y,AH,AI,AJ,AK,AL,AM,AN"

if not os.path.exists(my_dir):
    os.makedirs(my_dir)
os.chdir(my_dir)
files = os.listdir()

df = None

for my_file in files:
    try:
        my_df = pd.read_excel(my_file, usecols=my_cols)
        if df is not None:
            df = df.append(my_df)
        else:
            df = my_df
    except:
        pass

return df