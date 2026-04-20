import pandas as pd
import re
import os
import json
import glob
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def select_folder(title="בחר תיקייה"):
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_path

def get_latest_report_data(reports_path):
    if not os.path.exists(reports_path): return {}
    
    # חיפוש כל קובצי האקסל בתיקייה
    files = glob.glob(os.path.join(reports_path, "*.xlsx"))
    
    # סינון קבצים זמניים של אקסל
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    
    if not files: return {}

    # מציאת הקובץ שהשתנה לאחרונה (הכי חדש לפי זמן שמירה)
    latest_file = max(files, key=os.path.getmtime)
    
    try:
        print(f"קורא נתונים מהדוח האחרון: {os.path.basename(latest_file)}")
        old_df = pd.read_excel(latest_file)
        cols_to_take = [c for c in ['הערות סוכן', 'תאריך עידכון הערת סוכן'] if c in old_df.columns]
        
        # מוודאים שיש T. File No. כדי למפות את הנתונים
        if 'T. File No.' not in old_df.columns:
            return {}
            
        return old_df.set_index('T. File No.')[cols_to_take].to_dict('index')
    except Exception as e:
        print(f"שגיאה בקריאת הקובץ האחרון: {e}")
        return {}

def run_daily_process(input_excel_path, reports_dir, json_db_path):
    # 1. חילוץ תאריך עוגן משם הקובץ המקורי
    file_name = os.path.basename(input_excel_path)
    date_match = re.search(r'(\d{1,2}\s\w{3}\s\d{4})', file_name)
    anchor_date = datetime.strptime(date_match.group(1), '%d %b %Y') if date_match else datetime.now()
    
    # 2. קריאת האקסל
    df = pd.read_excel(input_excel_path, sheet_name='Ext. T. File Report', skiprows=8)
    df = df[df['Branch'].notna()].copy()
    
    cols = ['T. File No.', 'User', 'Start Date', 'End Date', 'T. File Name', 'Agent/C. Client', 'Unconfirmed Refund', 'Balance']
    new_df = df[cols].copy()
    new_df['Start Date'] = pd.to_datetime(new_df['Start Date'])

    # 3. ניהול JSON
    state = {}
    if os.path.exists(json_db_path):
        with open(json_db_path, 'r', encoding='utf-8') as f:
            state = json.load(f)

    for _, row in new_df.iterrows():
        state[str(row['T. File No.'])] = {'Open Date': str(row['Start Date'])}

    # שליפת הערות מהקובץ האחרון (לפי זמן עריכה)
    history_comments = get_latest_report_data(reports_dir)

    # 4. לוגיקת העמודות
    def final_logic(row):
        fid_num = row['T. File No.']
        fid_str = str(fid_num)
        is_direct = str(row['Agent/C. Client']).strip() == 'Direct Sale'
        
        days_to_departure = (row['Start Date'] - anchor_date).days if is_direct else ""
        
        days_since_refund = 0
        if fid_str in state and 'Open Date' in state[fid_str]:
            open_date = pd.to_datetime(state[fid_str]['Open Date'])
            days_since_refund = (anchor_date - open_date).days
        
        refund_note = ""
        if row['Unconfirmed Refund'] > 0:
            refund_note = "החזר שטרם אושר מעל 60 יום" if days_since_refund > 60 else "החזר שטרם אושר"
            
        system_note = refund_note if is_direct else (f"{refund_note}; טרם הופקה חשבונית" if refund_note else "טרם הופקה חשבונית")
        
        # שימוש במפתח כמספר (Int) או כמחרוזת בהתאם למה שיש ב-JSON
        hist_entry = history_comments.get(fid_num, {})
        return pd.Series([days_to_departure, days_since_refund, system_note, hist_entry.get('הערות סוכן', ""), hist_entry.get('תאריך עידכון הערת סוכן', "")])

    # 5. עדכון הדוח
    new_cols = ["מס' הימים שנותרו עד ליציאה", "מס' ימי החזר שטרם אושר", "הערות", "הערות סוכן", "תאריך עידכון הערת סוכן"]
    new_df[new_cols] = new_df.apply(final_logic, axis=1)

    # 6. שמירה
    with open(json_db_path, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=4)
    
    if not os.path.exists(reports_dir): os.makedirs(reports_dir)
    
    timestamp = datetime.now().strftime("%d.%m.%y %H.%M")
    output_path = os.path.join(reports_dir, f"דוח בקרה וואוצרים {timestamp}.xlsx")
    
    new_df.to_excel(output_path, index=False)
    print(f"--- הקובץ נשמר בהצלחה: {output_path} ---")

if __name__ == "__main__":
    base_folder = select_folder(title="בחר את תיקיית הפרויקט הראשי")
    if not base_folder: exit()
    
    reports_dir = select_folder(title="בחר את תיקיית הדוחות")
    json_db_path = os.path.join(base_folder, "refunds_state.json")
    
    all_excel_files = glob.glob(os.path.join(base_folder, "*.xlsx"))
    valid_files = [f for f in all_excel_files if not os.path.basename(f).startswith("~$")]

    if not valid_files:
        print("לא נמצאו קבצי אקסל בתיקייה הראשית.")
    else:
        latest_input = max(valid_files, key=os.path.getctime)
        run_daily_process(latest_input, reports_dir, json_db_path)