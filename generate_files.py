import pandas as pd
import numpy as np
import os
import shutil
from datetime import datetime

# === הגדרות נתיבים ===

# הקובץ הראשי (הטמפלט)
input_excel_path = r"C:\Users\dwmas\OneDrive\MYPRO4U\לקוחות או מקבלי שירותים\SC\TEMPLATE_AINVOICES_Yuval_LOCALE.xlsm"

# תיקיית פלט חדשה
output_folder = r"C:\Users\dwmas\OneDrive\MYPRO4U\לקוחות או מקבלי שירותים\SC\SC_Priority_Kaytanot_Interface\Exports\\"

# תיקיית גיבויים
backups_folder = r"C:\Users\dwmas\OneDrive\MYPRO4U\לקוחות או מקבלי שירותים\SC\SC_Priority_Kaytanot_Interface\Backups\\"

print("💡 תזכורת: ודא שהטמפלט מעודכן עם הנתונים מהרענון!")

def backup_file_if_exists(filepath):
    """גיבוי קובץ קיים"""
    if os.path.exists(filepath):
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        filename = os.path.basename(filepath)
        backup_name = f"{filename}.{timestamp}.bak"
        backup_path = os.path.join(backups_folder, backup_name)
        shutil.copy2(filepath, backup_path)
        print(f"קובץ קיים -> נוצר גיבוי: {backup_path}")

def create_cust_file(excel_path, output_dir):
    """מייצר את הקובץ CUST.txt"""
    try:
        df_cust = pd.read_excel(
            excel_path,
            sheet_name="CUST_TO_INT",
            usecols="P:V",
            header=None,
            skiprows=5 
        )

        # נקה תווים חריגים
        df_cust = df_cust.applymap(
            lambda x: str(x).replace('\u202a', '') if isinstance(x, str) else x
        )

        output_path = os.path.join(output_dir, "CUST.txt")
        backup_file_if_exists(output_path)

        df_cust.to_csv(output_path, sep='\t', index=False, header=False, encoding='utf-8')
        print(f"קובץ CUST.txt נוצר בהצלחה בנתיב: {output_path}")

    except Exception as e:
        print(f"שגיאה ביצירת קובץ CUST: {e}")

def create_eiv_file(excel_path, output_dir):
    """מייצר את הקובץ EIV.txt"""
    try:
        col_names = ['מסד', 'מזהה', 'ColC', 'נתונים', 'ColE', 'ColF', 'ColG']
        df_eiv = pd.read_excel(
            excel_path,
            sheet_name="EIV_INTERFACE",
            usecols="A:G",
            header=None,
            skiprows=1,
            names=col_names
        )

        df_eiv.dropna(subset=['מסד', 'מזהה'], inplace=True)
        df_eiv.sort_values(by=['מסד', 'מזהה'], ascending=[True, True], inplace=True)

        df_eiv['נתונים_כתאריך'] = pd.to_datetime(df_eiv['נתונים'], errors='coerce')

        df_eiv['נתונים'] = np.where(
            (df_eiv['מזהה'] == 1) & (df_eiv['נתונים_כתאריך'].notna()),
            df_eiv['נתונים_כתאריך'].dt.strftime('%d/%m/%y'),
            df_eiv['נתונים']
        )

        df_eiv['נתונים'] = np.where(
            (df_eiv['מזהה'] == 3),
            pd.to_numeric(df_eiv['נתונים'], errors='coerce').fillna(0).astype(int).astype(str),
            df_eiv['נתונים']
        )

        df_eiv.drop(columns=['מסד', 'נתונים_כתאריך'], inplace=True)
        df_eiv['מזהה'] = pd.to_numeric(df_eiv['מזהה'], errors='coerce').fillna(0).astype(int)
        df_eiv.fillna('', inplace=True)

        output_path = os.path.join(output_dir, "EIV.txt")
        backup_file_if_exists(output_path)

        df_eiv.to_csv(output_path, sep='\t', index=False, header=False, encoding='utf-8')
        print(f"קובץ EIV.txt נוצר בהצלחה בנתיב: {output_path}")

    except Exception as e:
        print(f"שגיאה ביצירת קובץ EIV: {e}")

# === הרצת כל התהליך ===

if __name__ == "__main__":
    create_cust_file(input_excel_path, output_folder)
    create_eiv_file(input_excel_path, output_folder)
    print("\n✅ כל התהליך הושלם בהצלחה! שני קבצי הטקסט מוכנים.")
