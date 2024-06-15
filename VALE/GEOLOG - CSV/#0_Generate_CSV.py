import pandas as pd
import glob
import numpy as np
import shutil
import os


def create_lith_csv():
    xlsx_files = glob.glob('*.xlsx')

    desired_columns = ['DEPTH', 'ROCK']
    df = pd.read_excel(xlsx_files[0], sheet_name='Interval Information', header=1)
    
    missing_columns = [col for col in desired_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"The following columns are missing in the 'Interval Information' sheet: {missing_columns}")

    df = df[desired_columns].rename(columns={'DEPTH': 'TOP', 'ROCK': 'LITH'})
    df['LITH'] = df['LITH'].shift(-1).str.strip()  # Removed leading and trailing spaces
    df['BOTTOM'] = df['TOP'].shift(-1)

    df = df.dropna(subset=['TOP'])
    df = df.dropna(subset=['BOTTOM'])

    df = df[['TOP', 'BOTTOM', 'LITH']]
    df.to_csv('Lith.csv', index=False)


def create_dip_csv():
    xlsx_files = glob.glob('*.xlsx')

    desired_columns = ['DEPTH', 'DIP']
    df = pd.read_excel(xlsx_files[0], sheet_name='Directional Information', header=0)
    
    missing_columns = [col for col in desired_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"The following columns are missing in the 'Directional Information' sheet: {missing_columns}")

    df = df[desired_columns]


    df.sort_values('DEPTH', inplace=True)


    df.set_index('DEPTH', inplace=True)


    new_index = pd.Index(np.arange(df.index.min(), df.index.max(), 0.1), name='DEPTH')
    df = df.reindex(new_index)


    df['DIP'] = df['DIP'].interpolate()

    df.reset_index(inplace=True)

    df.to_csv('Dip.csv', index=False)


def create_azi_csv():
    xlsx_files = glob.glob('*.xlsx')

    desired_columns = ['DEPTH', 'AZIMUTH']
    df = pd.read_excel(xlsx_files[0], sheet_name='Directional Information', header=0)
    
    missing_columns = [col for col in desired_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"The following columns are missing in the 'Directional Information' sheet: {missing_columns}")

    df = df[desired_columns]

    df.to_csv('Azi.csv', index=False)    


    
def create_assays_csv():
    xlsx_files = glob.glob('*.xlsx')

    desired_columns = ['DEPTH', 'EST', 'CU', 'NI', 'CO', 'AS', 'S', 'FE', 'PB', 'ZN', 'PT', 'PD', 'AU', 'TPM', 'RH', 'AG', 'SG']
    df = pd.read_excel(xlsx_files[0], sheet_name='Interval Information', header=1)

    missing_columns = [col for col in desired_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"The following columns are missing in the 'Interval Information' sheet: {missing_columns}")

    df = df[desired_columns]

   
    df.sort_values('DEPTH', inplace=True)

    df.set_index('DEPTH', inplace=True)


    new_index = pd.Index(np.arange(df.index.min(), df.index.max(), 0.1), name='DEPTH')
    df = df.reindex(new_index)


    for col in df.columns:
        if col != 'DEPTH':
            df[col] = df[col].interpolate(method='nearest')
            df[col] = df[col].bfill()  
            df[col] = df[col].ffill()  
            

    df.reset_index(inplace=True)

    df.to_csv('Assays.csv', index=False)



def move_files(file_name, target_folder):

    if not os.path.exists(file_name):
        print(f"File {file_name} does not exist. Skipping.")
        return

   
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)


    shutil.move(file_name, target_folder)



create_lith_csv()
create_dip_csv()
create_azi_csv()
create_assays_csv()


move_files('Azi.csv', 'Azimuth')
move_files('Dip.csv', 'Dip')
move_files('Lith.csv', 'Lith')
move_files('Assays.csv', 'Assays')


for file in glob.glob('*.xlsx'):
    move_files(file, 'GeoLog')