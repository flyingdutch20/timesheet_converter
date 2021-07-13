import pandas as pd
import os
import configparser

def create_config_ini():
    if not os.path.exists("config.ini"):
        config = configparser.ConfigParser()
        config.add_section("OPTIONS")
        config.set("OPTIONS", "reqlogic_directory", "reqlogic")
        config.set("OPTIONS", "output_directory", "output")
        config.set("OPTIONS", "columns", "A,B,H,K,M,T,U,V,W,Y,AH,AI,AJ,AK,AL,AM,AN")
        with open("config.ini", 'w') as configfile:
            config.write(configfile)

create_config_ini()
config = configparser.ConfigParser()
config.read("config.ini")
options = config["OPTIONS"]
my_dir = options["reqlogic_directory"]
my_cols = options["columns"]
my_out = options["output_directory"]

if not os.path.exists(my_dir):
    os.makedirs(my_dir)
files = os.listdir(my_dir)

df = None

for my_file in files:
    try:
        my_df = pd.read_excel(my_dir + '/' + my_file, usecols=my_cols)
        if df is not None:
            df = df.append(my_df)
            df = df.drop_duplicates()
        else:
            df = my_df
    except:
        pass

if df is not None:
    if not os.path.exists(my_out):
        os.makedirs(my_out)
    output_name = my_out + "/" + "output.xlsx"
    df.to_excel(output_name)
