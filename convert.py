import pandas as pd
import os
import configparser
import datetime as dt

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
            df = pd.concat([df, my_df])
            df = df.drop_duplicates()
        else:
            df = my_df
    except:
        pass


my_lines = []
for index, my_line in df.iterrows():
    if my_line[3]=='Processed' or my_line[3]=='Approved':
        for idx, day in enumerate(['MON','TUE','WED','THU','FRI','SAT','SUN']):
            if my_line[10+idx] > 0:
                my_date = my_line[2] + dt.timedelta(days=idx)
                my_week = my_date.strftime("%U")
                my_year = my_date.strftime("%Y")
                my_month_no = my_date.strftime("%m")
                my_month = my_date.strftime("%b")
                my_project = f'{my_line[5]} - {my_line[6]}'
                my_days = my_line[10+idx] / 7.5
                my_row = [my_line[0],my_line[1],my_date,day,my_week,my_month_no,my_month,my_year,my_line[4],my_line[5],my_line[6],my_project,my_line[7],my_line[8],my_line[9],my_line[10+idx],my_days]
                my_lines.append(my_row)


df_out = pd.DataFrame(columns=['DOCNBR','NAME','DATE','DAY','WEEK','MONTHNUMBER','MONTH','YEAR','LINENBR','PROJECTID','PROJECTNAME','PROJECT','TASKID','TASKNAME','COMMENT','HOURS','DAYS'],data=my_lines)

if not os.path.exists(my_out):
    os.makedirs(my_out)
os.chdir(my_out)
output_name = "output.xlsx"
df_out.to_excel(output_name, index=False)
counter = 0
while os.path.exists(output_name):
    counter += 1
    output_name = f'output{counter}.xlsx'
df_out.to_excel(output_name, index=False)

