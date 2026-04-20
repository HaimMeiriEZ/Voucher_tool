import pandas as pd
import os
import json
import glob
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def select_folder(title="בחר תיקייה"):
    """פתיחת חלון לבחירת תיקייה"""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_path

def get_latest_agent_comments(reports_path):
    """שליפת הערות סוכן מהדוח האחרון שנוצר (לפי זמן שינוי אחרון)"""
    if not os.path.exists(reports_path): return {}
    files = glob.glob(os.path.join(reports_path, "*.csv"))
    # סינון קבצים זמניים
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files: return {}
    
    # שימוש ב-getmtime כדי למצוא את הקובץ שעודכן לאחרונה
    latest_file = max(files, key=os.path.getmtime)
    
    try:
        old_df = pd.read_csv(latest_file, encoding='utf-8-sig')
        if 'Number' not in old_df.columns: return {}
        return old_df.set_index('Number')[['הערות סוכן', 'תאריך עידכון הערת סוכן']].to_dict('index')
    except: return {}

def run_daily_process_new_company(input_folder, reports_dir, json_db_path):
    # 1. מציאת קובץ ה-TXT האחרון (לפי זמן שינוי)
    all_txt_files = glob.glob(os.path.join(input_folder, "*.txt"))
    if not all_txt_files:
        print("לא נמצאו קבצי TXT בתיקייה שנבחרה.")
        return
    
    input_file_path = max(all_txt_files, key=os.path.getmtime)
    
    # 2. קריאת הקובץ
    df = pd.read_csv(input_file_path, sep=r'\s*\^\s*', engine='python', skipinitialspace=True, encoding='cp1255')
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.columns = df.columns.str.strip()
    
    # שימוש ב-getmtime לקביעת תאריך העוגן
    anchor_date = datetime.fromtimestamp(os.path.getmtime(input_file_path))
    
    # 3. ניהול JSON
    if os.path.exists(json_db_path):
        with open(json_db_path, 'r', encoding='utf-8') as f:
            state = json.load(f)
    else:
        state = {}

    for _, row in df.iterrows():
        fid = str(row['Number'])
        state[fid] = {'Date of first recognition': str(row['Open'])}
        
    with open(json_db_path, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=4)
    
    agent_history = get_latest_agent_comments(reports_dir)

    # 4. לוגיקה לעמודות הדוח
    def final_logic(row):
        try:
            start_date = pd.to_datetime(row['Start'], format='%d/%m/%Y')
            open_date = pd.to_datetime(row['Open'], format='%d/%m/%Y')
            days_to_departure = (start_date - anchor_date).days
            days_since_refund = (anchor_date - open_date).days
        except:
            days_to_departure, days_since_refund = "", 0
        
        raw_ref = str(row['Ref wl']).strip().replace(',', '')
        ref_val = float(raw_ref) if raw_ref != "" and raw_ref != "nan" else 0
        
        refund_note = ""
        if ref_val != 0:
            refund_note = "החזר שטרם אושר מעל 60 יום" if days_since_refund > 60 else "החזר שטרם אושר"
            
        is_empty_client = pd.isna(row['C.Client']) or str(row['C.Client']).strip() == ""
        system_note = f"{refund_note}; טרם הופקה חשבונית" if (not is_empty_client and refund_note) else ("טרם הופקה חשבונית" if not is_empty_client else refund_note)

        hist = agent_history.get(int(row['Number']), {})
        return pd.Series([days_to_departure, days_since_refund, system_note, hist.get('הערות סוכן', ""), hist.get('תאריך עידכון הערת סוכן', "")])

    df[["מס' ימים ליציאה", "מס' ימי החזר", "הערות", "הערות סוכן", "תאריך עידכון הערת סוכן"]] = df.apply(final_logic, axis=1)

    # 5. שמירת דוח CSV עם שם מותאם אישית (דוח בקרה וואוצרים + תאריך ושעה)
    if not os.path.exists(reports_dir): os.makedirs(reports_dir)
    
    timestamp = datetime.now().strftime("%d.%m.%y %H.%M")
    output_path = os.path.join(reports_dir, f"דוח בקרה וואוצרים {timestamp}.csv")
    
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"--- תהליך הושלם! הקובץ נשמר ב: {output_path} ---")

if __name__ == "__main__":
    base_folder = select_folder(title="בחר את תיקיית העבודה (בה נמצא ה-TXT וה-JSON)")
    if base_folder:
        reports_dir = select_folder(title="בחר את תיקיית הדוחות")
        json_db_path = os.path.join(base_folder, "refunds_state_new.json")
        
        run_daily_process_new_company(base_folder, reports_dir, json_db_path)