import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd
from tqdm import tqdm
from datetime import datetime

def compare_excel_files_by_keys(file1, file2, output_file):
    
    sheets = pd.ExcelFile(file1).sheet_names

    for sheet in tqdm(sheets, desc="Comparing sheets", unit="sheet", total=len(sheets)):
        
        df1 = pd.read_excel(file1, sheet_name=sheet).dropna(how='all').astype(str)
        df2 = pd.read_excel(file2, sheet_name=sheet).dropna(how='all').astype(str)

        df1['filename'] = file1
        df2['filename'] = file2

        df = pd.concat([df1, df2], ignore_index=True)

        df['same'] = df.duplicated(subset=df.columns[:-2], keep=False).astype(int)

        other_cols = ['MO', 'Atributo' if 'Atributo' in df.columns else 'Feature', 'filename', 'same']
        cols = df.columns.tolist()
        cols = other_cols + [col for col in cols if col not in other_cols]
        df = df[cols]

        df = df.sort_values(['Atributo' if 'Atributo' in df.columns else 'Feature', 'filename'])

        if sheet == sheets[0]:
            df = df[cols]
            df.to_excel(output_file, sheet_name=sheet, index=False)
        else:
            with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet, index=False)

        pass

    print("Comparison completed. Results saved to", output_file)

file1 = "MZI051_OPTI.xlsx"
file2 = "MZI051_RAW.xlsx"
output_file = file1.split(".")[0] + "-" + file2.split(".")[0] + "_comparison_" + datetime.now().strftime("%Y%m%d%H%M") + ".xlsx"

compare_excel_files_by_keys(file1, file2, output_file)