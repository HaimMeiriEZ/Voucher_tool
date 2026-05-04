from __future__ import annotations

import glob
import json
import logging
import os
import re
import sys
import tempfile
from datetime import datetime
from typing import Callable, Dict, Optional

import pandas as pd
from openpyxl import load_workbook
from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

APP_TITLE = "כלי ניהול ובקרת וואצ'רים"
STATE_FILE_NAME = "voucher_state.json"
AGENTS_CONFIG_FILE_NAME = "agents_config.json"
NOTE_COLUMNS = ["הערות", "הערות סוכן", "תאריך עידכון הערת סוכן"]
AGENT_REQUIRED_COLUMNS = ["משתמש בגלבוע", "סניף", "מייל", "מנהל סניף", "מנהל תחום", "מייל יפה", "מייל אילנית"]
_APP_DIR = os.path.dirname(os.path.abspath(__file__))

# File logger — writes to voucher_tool.log next to the script
_LOG_FILE = os.path.join(_APP_DIR, "voucher_tool.log")
logging.basicConfig(
    filename=_LOG_FILE,
    level=logging.DEBUG,
    encoding="utf-8",
    format="%(asctime)s %(levelname)-8s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
_logger = logging.getLogger("voucher_tool")


class ProcessingError(Exception):
    pass


def safe_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalize_identifier(value) -> str:
    text = safe_text(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text


def safe_datetime(value, *, dayfirst: bool = False) -> Optional[datetime]:
    text = safe_text(value)
    if not text:
        return None
    result = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
    if pd.isna(result):
        return None
    return result.to_pydatetime()


def build_alert_fields(days_to_departure: Optional[int], is_private: bool) -> tuple[str, str, str, str]:
    if not is_private or days_to_departure is None:
        return "", "", "", ""

    if 8 <= days_to_departure <= 14:
        category = "התראה שבועיים לפני היציאה"
        body = (
            "הלקוחות הבאים אמורים לצאת מהארץ בשבועיים הקרובים וטרם נגבה בגינם תשלום ההזמנה, "
            "במידה ותוך שבוע לא תתבצע גבייה היא תבוטל לאלתר בכדי למנוע הפסד לחברה"
        )
        reason = "חוסר גבייה כשבועיים לפני היציאה"
    elif 0 <= days_to_departure <= 7:
        category = "עלול להתבטל"
        body = (
            "הזמנת הלקוח טרם נגבתה, ואמורה לצאת לפועל תוך שבוע – לכן היא מיועדת לביטול ותבוטל עד סוף היום"
        )
        reason = "חוסר גבייה לפני היציאה"
    else:
        return "", "", "", ""

    return category, body, category, reason


def clean_numeric_series(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
    cleaned = cleaned.replace("", pd.NA)
    cleaned = cleaned.replace("nan", pd.NA)
    cleaned = cleaned.replace("None", pd.NA)
    return pd.to_numeric(cleaned, errors="coerce")


def load_agents_config(config_path: str) -> dict:
    if not os.path.exists(config_path):
        return {}
    with open(config_path, "r", encoding="utf-8") as file:
        return json.load(file)


def save_agents_config(config_path: str, data: dict) -> None:
    with open(config_path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)


def find_latest_output_file(output_dir: str, source_name: str) -> Optional[str]:
    if not output_dir or not os.path.exists(output_dir):
        return None
    candidates = [
        path
        for path in glob.glob(os.path.join(output_dir, "*.xlsx"))
        if not os.path.basename(path).startswith("~$")
        and source_name.lower() in os.path.basename(path).lower()
    ]
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)


def load_agent_table(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in AGENT_REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ProcessingError("עמודות חסרות בטבלת סוכנים: " + ", ".join(missing))
    df["משתמש בגלבוע"] = df["משתמש בגלבוע"].astype(str).str.strip()
    return df


def build_agent_rows_map(
    output_dir: str,
    agent_df: pd.DataFrame,
    logger: Callable[[str], None],
) -> tuple[Dict[str, pd.DataFrame], pd.DataFrame]:
    agent_lookup = agent_df[AGENT_REQUIRED_COLUMNS].copy()
    agent_lookup["_join_key"] = agent_lookup["משתמש בגלבוע"].str.strip()
    matched_keys = set(agent_lookup["_join_key"])
    frames = []
    unmatched_frames = []

    booster_file = find_latest_output_file(output_dir, "BOOSTER")
    if booster_file:
        logger(f"טוען BOOSTER: {os.path.basename(booster_file)}")
        try:
            b_df = pd.read_excel(booster_file, sheet_name="דוח מלא")
            if "User" in b_df.columns:
                b_df = b_df.copy()
                b_df["_join_key"] = b_df["User"].astype(str).str.strip()
                matched = b_df.merge(agent_lookup, on="_join_key", how="inner")
                matched = matched.drop(columns=["_join_key"])
                frames.append(matched)
                unmatched_b = b_df[~b_df["_join_key"].isin(matched_keys)].copy()
                unmatched_b = unmatched_b.rename(columns={"_join_key": "_agent_key"})
                if not unmatched_b.empty:
                    unmatched_frames.append(unmatched_b)
        except Exception as exc:
            logger(f"שגיאה בטעינת BOOSTER: {exc}")

    gilboa_file = find_latest_output_file(output_dir, "GILBOA")
    if gilboa_file:
        logger(f"טוען GILBOA: {os.path.basename(gilboa_file)}")
        try:
            g_df = pd.read_excel(gilboa_file, sheet_name="דוח מלא")
            if "Clerk" in g_df.columns:
                g_df = g_df.copy()
                g_df["_join_key"] = g_df["Clerk"].astype(str).str.strip()
                matched = g_df.merge(agent_lookup, on="_join_key", how="inner")
                matched = matched.drop(columns=["_join_key"])
                frames.append(matched)
                unmatched_g = g_df[~g_df["_join_key"].isin(matched_keys)].copy()
                unmatched_g = unmatched_g.rename(columns={"_join_key": "_agent_key"})
                if not unmatched_g.empty:
                    unmatched_frames.append(unmatched_g)
        except Exception as exc:
            logger(f"שגיאה בטעינת GILBOA: {exc}")

    if not frames and not unmatched_frames:
        raise ProcessingError("לא נמצאו נתונים לשיוך בקבצי BOOSTER/GILBOA בתיקיית הפלט")

    result: Dict[str, pd.DataFrame] = {}
    if frames:
        combined = pd.concat(frames, ignore_index=True)
        for email_addr, group in combined.groupby("מייל"):
            email_str = safe_text(email_addr)
            if email_str:
                result[email_str] = group.reset_index(drop=True)
    unmatched_combined = pd.concat(unmatched_frames, ignore_index=True) if unmatched_frames else pd.DataFrame()
    return result, unmatched_combined


def create_outlook_draft(
    to_email: str,
    cc_emails: list[str],
    subject: str,
    html_body: str,
    *,
    subfolder_name: str = "",
    save_as_path: str = "",
    attachments: list[str] | None = None,
) -> None:
    try:
        import win32com.client
    except ImportError as exc:
        raise ProcessingError("חבילת pywin32 אינה מותקנת. הרץ: pip install pywin32") from exc
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_email
    if cc_emails:
        mail.CC = "; ".join(cc_emails)
    mail.Subject = subject
    mail.HTMLBody = html_body
    if attachments:
        for att_path in attachments:
            mail.Attachments.Add(att_path)
    if save_as_path:
        mail.SaveAs(save_as_path, 3)  # 3 = olMSG
    else:
        mail.Save()
        if subfolder_name:
            namespace = outlook.GetNamespace("MAPI")
            drafts = namespace.GetDefaultFolder(16)
            try:
                target_folder = drafts.Folders[subfolder_name]
            except Exception:
                target_folder = drafts.Folders.Add(subfolder_name)
            mail.Move(target_folder)


def _save_df_excel(df: pd.DataFrame, path: str) -> None:
    """Write df to Excel. If multiple sources (מקור column) exist each gets its own sheet
    with only the columns that have at least one non-empty value in that group."""
    def _drop_empty_cols(frame: pd.DataFrame) -> pd.DataFrame:
        def has_value(col):
            return col.apply(lambda v: pd.notna(v) and str(v).strip() not in ("", "nan", "None", "NaT")).any()
        return frame.loc[:, frame.apply(has_value)]

    if "מקור" in df.columns and df["מקור"].nunique() > 1:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for source, group in df.groupby("מקור"):
                _drop_empty_cols(group.reset_index(drop=True)).to_excel(writer, sheet_name=str(source), index=False)
    else:
        _drop_empty_cols(df.reset_index(drop=True)).to_excel(path, index=False, engine="openpyxl")


def prepare_agent_emails(
    output_dir: str,
    agent_table_path: str,
    logger: Callable[[str], None],
) -> int:
    _ROUTING_COLS = set(AGENT_REQUIRED_COLUMNS) | {
        "email_target", "email_employee", "email_direct_manager", "email_department_manager",
        "גוף דוא\u05f4ל",
    }
    _UNMATCHED_SUBFOLDER = "ממתינות מייל - כתובת חסרה"
    agent_df = load_agent_table(agent_table_path)
    rows_map, unmatched_df = build_agent_rows_map(output_dir, agent_df, logger)

    today_str = datetime.now().strftime("%d/%m/%Y")
    cc_source_columns = ["מנהל סניף", "מנהל תחום", "מייל יפה", "מייל אילנית"]
    count = 0
    for to_email, rows_df in rows_map.items():
        agent_name = safe_text(rows_df["משתמש בגלבוע"].iloc[0]) if "משתמש בגלבוע" in rows_df.columns else to_email
        cc_emails = []
        for col in cc_source_columns:
            if col in rows_df.columns:
                val = safe_text(rows_df[col].iloc[0])
                if val and val != to_email and val not in cc_emails:
                    cc_emails.append(val)

        display_cols = [c for c in rows_df.columns if c not in _ROUTING_COLS]
        cat_col = "email_target" if "email_target" in rows_df.columns else None
        df_all = rows_df[display_cols]
        df_warning = rows_df[rows_df[cat_col] == "התראה שבועיים לפני היציאה"][display_cols] if cat_col else pd.DataFrame()
        df_cancel = rows_df[rows_df[cat_col] == "עלול להתבטל"][display_cols] if cat_col else pd.DataFrame()
        count_all = len(df_all)
        count_warning = len(df_warning)
        count_cancel = len(df_cancel)

        tmp_dir = tempfile.mkdtemp()
        tmp_files = []
        all_path = os.path.join(tmp_dir, f"הזמנות פתוחות - {agent_name}.xlsx")
        _save_df_excel(df_all, all_path)
        tmp_files.append(all_path)
        if not df_warning.empty:
            warn_path = os.path.join(tmp_dir, f"הזמנות שבועיים לפני יציאה - {agent_name}.xlsx")
            _save_df_excel(df_warning, warn_path)
            tmp_files.append(warn_path)
        if not df_cancel.empty:
            cancel_path = os.path.join(tmp_dir, f"הזמנות עלולות להתבטל - {agent_name}.xlsx")
            _save_df_excel(df_cancel, cancel_path)
            tmp_files.append(cancel_path)

        body_lines = [
            "<p>שלום,</p>",
            "<p>להלן סיכום ההזמנות הפתוחות עבורך (הפרטים המלאים מצורפים כקבצי אקסל):</p>",
            f"<p>&#128196; <strong>הזמנות פתוחות:</strong> {count_all} רשומות</p>",
        ]
        if count_warning:
            body_lines.append(f"<p>&#9888; <strong>הזמנות שבועיים לפני היציאה:</strong> {count_warning} רשומות</p>")
        if count_cancel:
            body_lines.append(f"<p>&#128308; <strong>הזמנות עלולות להתבטל עד סוף היום:</strong> {count_cancel} רשומות</p>")
        html_body = '<html><head></head><body dir="rtl">' + "".join(body_lines) + "</body></html>"

        subject = f"דוח בקרה וואוצרים \u2014 {agent_name} \u2014 {today_str}"
        create_outlook_draft(to_email, cc_emails, subject, html_body, attachments=tmp_files)
        logger(f"נוצרה טיוטה עבור: {agent_name} ({to_email})")
        count += 1
        for p in tmp_files:
            try:
                os.remove(p)
            except Exception:
                pass
        try:
            os.rmdir(tmp_dir)
        except Exception:
            pass

    if not unmatched_df.empty:
        unmatched_out_dir = os.path.join(output_dir, _UNMATCHED_SUBFOLDER)
        os.makedirs(unmatched_out_dir, exist_ok=True)
        logger(f"נמצאו {len(unmatched_df)} רשומות ללא כתובת מייל — שומר קבצי מייל בתיקייה '{_UNMATCHED_SUBFOLDER}'")
        display_routing_exclude = _ROUTING_COLS | {"_agent_key"}
        for agent_key, grp in unmatched_df.groupby("_agent_key"):
            agent_key_str = safe_text(agent_key)
            display_cols = [c for c in grp.columns if c not in display_routing_exclude]
            cat_col = "email_target" if "email_target" in grp.columns else None
            df_all = grp[display_cols]
            df_warning = grp[grp[cat_col] == "התראה שבועיים לפני היציאה"][display_cols] if cat_col else pd.DataFrame()
            df_cancel = grp[grp[cat_col] == "עלול להתבטל"][display_cols] if cat_col else pd.DataFrame()
            count_all = len(df_all)
            count_warning = len(df_warning)
            count_cancel = len(df_cancel)

            tmp_dir = tempfile.mkdtemp()
            tmp_files = []
            all_path = os.path.join(tmp_dir, f"הזמנות פתוחות - {agent_key_str}.xlsx")
            _save_df_excel(df_all, all_path)
            tmp_files.append(all_path)
            if not df_warning.empty:
                warn_path = os.path.join(tmp_dir, f"הזמנות שבועיים לפני יציאה - {agent_key_str}.xlsx")
                _save_df_excel(df_warning, warn_path)
                tmp_files.append(warn_path)
            if not df_cancel.empty:
                cancel_path = os.path.join(tmp_dir, f"הזמנות עלולות להתבטל - {agent_key_str}.xlsx")
                _save_df_excel(df_cancel, cancel_path)
                tmp_files.append(cancel_path)

            body_lines = [
                f"<p style=\"color:#e11d48;\"><strong>&#9888; לא נמצאה כתובת מייל עבור סוכן: {agent_key_str}</strong></p>",
                "<p>יש להוסיף את כתובת המייל ידנית לפני שליחה.</p>",
                f"<p>&#128196; <strong>הזמנות פתוחות:</strong> {count_all} רשומות</p>",
            ]
            if count_warning:
                body_lines.append(f"<p>&#9888; <strong>הזמנות שבועיים לפני היציאה:</strong> {count_warning} רשומות</p>")
            if count_cancel:
                body_lines.append(f"<p>&#128308; <strong>הזמנות עלולות להתבטל עד סוף היום:</strong> {count_cancel} רשומות</p>")
            html_body = '<html><head></head><body dir="rtl">' + "".join(body_lines) + "</body></html>"

            subject = f"דוח בקרה וואוצרים \u2014 {agent_key_str} \u2014 {today_str} [חסרה כתובת מייל]"
            safe_key = agent_key_str.replace("/", "-").replace("\\", "-")
            msg_path = os.path.join(unmatched_out_dir, f"דוח {safe_key} {today_str.replace('/', '.')}.msg")
            _UNMATCHED_CC = ["ilanit_b@ophirtours.co.il", "YWaksman@mycwt.co.il"]
            create_outlook_draft("", _UNMATCHED_CC, subject, html_body, save_as_path=msg_path, attachments=tmp_files)
            logger(f"קובץ מייל נשמר: {os.path.basename(msg_path)}")
            count += 1
            for p in tmp_files:
                try:
                    os.remove(p)
                except Exception:
                    pass
            try:
                os.rmdir(tmp_dir)
            except Exception:
                pass

    return count


class BaseVoucherProcessor:
    source_name = "BASE"
    history_key_columns = []

    def __init__(self, logger: Optional[Callable[[str], None]] = None):
        self.logger = logger or (lambda message: None)

    def log(self, message: str) -> None:
        self.logger(message)

    def ensure_columns(self, df: pd.DataFrame, required_columns: list[str]) -> None:
        missing = [column for column in required_columns if column not in df.columns]
        if missing:
            raise ProcessingError("עמודות חסרות בקובץ: " + ", ".join(missing))

    def load_state(self, state_path: str) -> Dict[str, Dict[str, str]]:
        if not os.path.exists(state_path):
            return {}
        with open(state_path, "r", encoding="utf-8") as file:
            return json.load(file)

    def save_state(self, state_path: str, state: Dict[str, Dict[str, str]]) -> None:
        with open(state_path, "w", encoding="utf-8") as file:
            json.dump(state, file, ensure_ascii=False, indent=2)

    def get_anchor_date(self, input_path: str) -> datetime:
        return datetime.now()

    def build_history_map(self, output_dir: str) -> Dict[str, Dict[str, str]]:
        if not output_dir or not os.path.exists(output_dir):
            return {}

        candidates = []
        for pattern in ("*.xlsx", "*.csv"):
            candidates.extend(
                [
                    path
                    for path in glob.glob(os.path.join(output_dir, pattern))
                    if not os.path.basename(path).startswith("~$")
                ]
            )

        if not candidates:
            return {}

        source_matches = [
            path for path in candidates if self.source_name.lower() in os.path.basename(path).lower()
        ]
        selected = source_matches if source_matches else candidates
        latest_file = max(selected, key=os.path.getmtime)
        self.log(f"טוען היסטוריית הערות: {os.path.basename(latest_file)}")

        try:
            if latest_file.lower().endswith(".xlsx"):
                old_df = pd.read_excel(latest_file)
            else:
                old_df = pd.read_csv(latest_file, encoding="utf-8-sig")
        except Exception:
            return {}

        id_column = None
        if "מקור" in old_df.columns and "מזהה רשומה" in old_df.columns:
            old_df = old_df[old_df["מקור"].astype(str).str.upper() == self.source_name.upper()].copy()
            id_column = "מזהה רשומה"
        else:
            for candidate in self.history_key_columns:
                if candidate in old_df.columns:
                    id_column = candidate
                    break

        if not id_column:
            return {}

        map_data: Dict[str, Dict[str, str]] = {}
        for _, row in old_df.iterrows():
            record_id = normalize_identifier(row.get(id_column, ""))
            if not record_id:
                continue
            map_data[record_id] = {
                "הערות סוכן": safe_text(row.get("הערות סוכן", "")),
                "תאריך עידכון הערת סוכן": safe_text(row.get("תאריך עידכון הערת סוכן", "")),
            }
        return map_data

    def history_lookup(self, history_map: Dict[str, Dict[str, str]], raw_identifier) -> Dict[str, str]:
        record_id = normalize_identifier(raw_identifier)
        return history_map.get(record_id, {})

    def export_report(self, df: pd.DataFrame, output_dir: str) -> str:
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%d.%m.%y %H.%M")
        output_path = os.path.join(output_dir, f"דוח בקרה וואוצרים {self.source_name} {timestamp}.xlsx")

        warning_14 = pd.DataFrame()
        threat_7 = pd.DataFrame()
        if "קטגוריית התראה" in df.columns:
            warning_14 = df[df["קטגוריית התראה"] == "התראה שבועיים לפני היציאה"].copy()
            threat_7 = df[df["קטגוריית התראה"] == "עלול להתבטל"].copy()

        drop_columns = ["קטגוריית התראה", "סיבת התראה"]
        if any(column in df.columns for column in drop_columns):
            df = df.drop(columns=drop_columns, errors="ignore")
            warning_14 = warning_14.drop(columns=drop_columns, errors="ignore")
            threat_7 = threat_7.drop(columns=drop_columns, errors="ignore")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="דוח מלא", index=False)
            if not warning_14.empty:
                warning_14.to_excel(writer, sheet_name="אזהרה שבועיים", index=False)
            if not threat_7.empty:
                threat_7.to_excel(writer, sheet_name="עלול להתבטל", index=False)

        workbook = load_workbook(output_path)
        date_columns = {"Open", "Start", "Start Date", "End Date", "תאריך עידכון הערת סוכן"}
        integer_columns = {"Number", "מס' ימים ליציאה", "מס' ימי החזר"}
        decimal_columns = {"Ref wl", "Attr", "Credit", "Balance", "Unconfirmed Refund"}

        for worksheet in workbook.worksheets:
            headers = {cell.value: index + 1 for index, cell in enumerate(worksheet[1])}
            for column_name in date_columns:
                col_idx = headers.get(column_name)
                if not col_idx:
                    continue
                for row in range(2, worksheet.max_row + 1):
                    worksheet.cell(row=row, column=col_idx).number_format = "DD/MM/YYYY"
            for column_name in integer_columns:
                col_idx = headers.get(column_name)
                if not col_idx:
                    continue
                for row in range(2, worksheet.max_row + 1):
                    worksheet.cell(row=row, column=col_idx).number_format = "0"
            for column_name in decimal_columns:
                col_idx = headers.get(column_name)
                if not col_idx:
                    continue
                for row in range(2, worksheet.max_row + 1):
                    worksheet.cell(row=row, column=col_idx).number_format = "#,##0.00"

        workbook.save(output_path)
        return output_path

    def process(self, input_path: str, output_dir: str) -> str:
        self.log(f"מתחיל עיבוד עבור {self.source_name}")
        self.log(f"קובץ קלט: {os.path.basename(input_path)}")

        state_path = os.path.join(output_dir, STATE_FILE_NAME)
        state = self.load_state(state_path)
        history = self.build_history_map(output_dir)
        anchor_date = self.get_anchor_date(input_path)

        df = self.load_input(input_path)
        self.log(f"נטענו {len(df)} רשומות")
        result_df = self.apply_business_logic(df, state, history, anchor_date)
        self.save_state(state_path, state)
        output_path = self.export_report(result_df, output_dir)
        self.log(f"הפלט נשמר אל: {output_path}")
        return output_path

    def load_input(self, input_path: str) -> pd.DataFrame:
        raise NotImplementedError

    def apply_business_logic(
        self,
        df: pd.DataFrame,
        state: Dict[str, Dict[str, str]],
        history: Dict[str, Dict[str, str]],
        anchor_date: datetime,
    ) -> pd.DataFrame:
        raise NotImplementedError


class BoosterProcessor(BaseVoucherProcessor):
    source_name = "BOOSTER"
    history_key_columns = ["מזהה רשומה", "T. File No."]

    def get_anchor_date(self, input_path: str) -> datetime:
        return super().get_anchor_date(input_path)

    def load_input(self, input_path: str) -> pd.DataFrame:
        self.log("קורא קובץ BOOSTER")
        df = pd.read_excel(input_path, sheet_name="Ext. T. File Report", skiprows=8)
        df.columns = [str(column).strip() for column in df.columns]
        self.ensure_columns(
            df,
            [
                "Branch",
                "T. File No.",
                "User",
                "Start Date",
                "End Date",
                "T. File Name",
                "Agent/C. Client",
                "Unconfirmed Refund",
                "Balance",
            ],
        )
        df = df[df["Branch"].notna()].copy()
        return df[
            [
                "T. File No.",
                "User",
                "Start Date",
                "End Date",
                "T. File Name",
                "Agent/C. Client",
                "Unconfirmed Refund",
                "Balance",
            ]
        ].copy()

    def apply_business_logic(
        self,
        df: pd.DataFrame,
        state: Dict[str, Dict[str, str]],
        history: Dict[str, Dict[str, str]],
        anchor_date: datetime,
    ) -> pd.DataFrame:
        df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
        df["User"] = df["User"].astype(str).str.strip()
        df["מקור"] = self.source_name
        df["מזהה רשומה"] = df["T. File No."].apply(normalize_identifier)

        run_date = anchor_date.date()
        calc_columns = []
        for _, row in df.iterrows():
            record_id = normalize_identifier(row["T. File No."])
            start_date = row["Start Date"] if pd.notna(row["Start Date"]) else None
            if start_date is not None:
                state[f"{self.source_name}:{record_id}"] = {"open_date": str(start_date)}

            start_date_only = start_date.date() if start_date is not None else None
            days_to_departure = (
                int((start_date_only - run_date).days) if start_date_only is not None else None
            )

            state_entry = state.get(f"{self.source_name}:{record_id}", {})
            open_date = safe_datetime(state_entry.get("open_date", ""))
            days_since_refund = (anchor_date - open_date).days if open_date else 0

            refund_value = pd.to_numeric(row.get("Unconfirmed Refund", 0), errors="coerce")
            refund_value = 0 if pd.isna(refund_value) else float(refund_value)
            refund_note = ""
            if refund_value > 0:
                refund_note = "החזר שטרם אושר מעל 60 יום" if days_since_refund > 60 else "החזר שטרם אושר"

            is_private = safe_text(row["Agent/C. Client"]) == "Direct Sale"
            alert_category, email_body, email_target, alert_reason = build_alert_fields(
                days_to_departure, is_private
            )

            if safe_text(row["Agent/C. Client"]) == "Direct Sale":
                system_note = refund_note
            else:
                system_note = f"{refund_note}; טרם הופקה חשבונית" if refund_note else "טרם הופקה חשבונית"

            hist = self.history_lookup(history, row["T. File No."])
            calc_columns.append(
                [
                    days_to_departure,
                    days_since_refund,
                    system_note,
                    hist.get("הערות סוכן", ""),
                    hist.get("תאריך עידכון הערת סוכן", ""),
                    alert_category,
                    email_body,
                    email_target,
                    alert_reason,
                    "",
                    "",
                ]
            )

        df[
            [
                "מס' ימים ליציאה",
                "מס' ימי החזר",
                "הערות",
                "הערות סוכן",
                "תאריך עידכון הערת סוכן",
                "קטגוריית התראה",
                "גוף דוא״ל",
                "email_target",
                "סיבת התראה",
                "email_employee",
                "email_direct_manager",
            ]
        ] = pd.DataFrame(calc_columns, index=df.index)
        df["email_department_manager"] = ""

        preferred = [
            "מקור",
            "מזהה רשומה",
            "T. File No.",
            "User",
            "Start Date",
            "End Date",
            "T. File Name",
            "Agent/C. Client",
            "Unconfirmed Refund",
            "Balance",
            "מס' ימים ליציאה",
            "מס' ימי החזר",
            "הערות",
            "הערות סוכן",
            "תאריך עידכון הערת סוכן",
            "קטגוריית התראה",
            "סיבת התראה",
            "גוף דוא״ל",
            "email_target",
            "email_employee",
            "email_direct_manager",
            "email_department_manager",
        ]
        return df[preferred]


class GilboaProcessor(BaseVoucherProcessor):
    source_name = "GILBOA"
    history_key_columns = ["מזהה רשומה", "Number"]

    def load_input(self, input_path: str) -> pd.DataFrame:
        self.log("קורא קובץ GILBOA")
        read_errors = []
        for encoding in ("cp1255", "utf-8-sig", "utf-8"):
            try:
                df = pd.read_csv(
                    input_path,
                    sep=r"\s*\^\s*",
                    engine="python",
                    skipinitialspace=True,
                    encoding=encoding,
                )
                df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
                df.columns = [str(column).strip() for column in df.columns]
                self.ensure_columns(df, ["Number", "Open", "Start", "Ref wl", "C.Client", "Clerk"])
                return df
            except Exception as error:
                read_errors.append(str(error))

        raise ProcessingError("לא ניתן לקרוא את קובץ GILBOA. " + " | ".join(read_errors))

    def apply_business_logic(
        self,
        df: pd.DataFrame,
        state: Dict[str, Dict[str, str]],
        history: Dict[str, Dict[str, str]],
        anchor_date: datetime,
    ) -> pd.DataFrame:
        df["Open"] = pd.to_datetime(df["Open"], format="%d/%m/%Y", errors="coerce")
        df["Start"] = pd.to_datetime(df["Start"], format="%d/%m/%Y", errors="coerce")
        df["Clerk"] = df["Clerk"].astype(str).str.strip()

        for numeric_column in ["Number", "Ref wl", "Attr", "Credit", "Balance"]:
            if numeric_column in df.columns:
                df[numeric_column] = clean_numeric_series(df[numeric_column])

        df["מקור"] = self.source_name
        df["מזהה רשומה"] = df["Number"].apply(normalize_identifier)

        run_date = anchor_date.date()
        calc_columns = []
        for _, row in df.iterrows():
            record_id = normalize_identifier(row["Number"])
            open_date = row["Open"].to_pydatetime() if pd.notna(row["Open"]) else None
            start_date = row["Start"].to_pydatetime() if pd.notna(row["Start"]) else None
            if open_date is not None:
                state[f"{self.source_name}:{record_id}"] = {"open_date": str(open_date)}

            start_date_only = start_date.date() if start_date is not None else None
            days_to_departure = (
                int((start_date_only - run_date).days) if start_date_only is not None else None
            )
            days_since_refund = (anchor_date - open_date).days if open_date is not None else 0

            refund_value = row.get("Ref wl", 0)
            refund_value = 0.0 if pd.isna(refund_value) else float(refund_value)

            refund_note = ""
            if refund_value != 0:
                refund_note = "החזר שטרם אושר מעל 60 יום" if days_since_refund > 60 else "החזר שטרם אושר"

            is_private = not safe_text(row.get("C.Client", ""))
            alert_category, email_body, email_target, alert_reason = build_alert_fields(
                days_to_departure, is_private
            )

            client_text = safe_text(row.get("C.Client", ""))
            has_invoice_gap = bool(client_text)
            if has_invoice_gap and refund_note:
                system_note = f"{refund_note}; טרם הופקה חשבונית"
            elif has_invoice_gap:
                system_note = "טרם הופקה חשבונית"
            else:
                system_note = refund_note

            hist = self.history_lookup(history, row["Number"])
            calc_columns.append(
                [
                    days_to_departure,
                    days_since_refund,
                    system_note,
                    hist.get("הערות סוכן", ""),
                    hist.get("תאריך עידכון הערת סוכן", ""),
                    alert_category,
                    email_body,
                    email_target,
                    alert_reason,
                    "",
                    "",
                ]
            )

        df[
            [
                "מס' ימים ליציאה",
                "מס' ימי החזר",
                "הערות",
                "הערות סוכן",
                "תאריך עידכון הערת סוכן",
                "קטגוריית התראה",
                "גוף דוא״ל",
                "email_target",
                "סיבת התראה",
                "email_employee",
                "email_direct_manager",
            ]
        ] = pd.DataFrame(calc_columns, index=df.index)
        df["email_department_manager"] = ""

        added_columns = [
            "מס' ימים ליציאה",
            "מס' ימי החזר",
            "הערות",
            "הערות סוכן",
            "תאריך עידכון הערת סוכן",
            "קטגוריית התראה",
            "סיבת התראה",
            "גוף דוא״ל",
            "email_target",
            "email_employee",
            "email_direct_manager",
            "email_department_manager",
        ]
        leading_columns = [
            "מקור",
            "מזהה רשומה",
            "Number",
            "Open",
            "Start",
            "Ref wl",
            "C.Client",
            "Clerk",
        ]
        middle_columns = [
            column for column in df.columns if column not in leading_columns and column not in added_columns
        ]
        return df[leading_columns + middle_columns + added_columns]


class StatCard(QFrame):
    """Small summary card showing a count and a label."""

    _STYLE_MATCHED = (
        "QFrame { background: #dcfce7; border: 1.5px solid #16a34a; border-radius: 10px; }"
    )
    _STYLE_UNMATCHED = (
        "QFrame { background: #fff7ed; border: 1.5px solid #ea580c; border-radius: 10px; }"
    )
    _STYLE_IDLE = (
        "QFrame { background: #f1f5f9; border: 1.5px solid #cbd5e1; border-radius: 10px; }"
    )

    def __init__(self, variant: str, parent=None):
        super().__init__(parent)
        self._variant = variant  # "matched" | "unmatched"
        self.setMinimumHeight(72)

        inner = QVBoxLayout(self)
        inner.setContentsMargins(14, 8, 14, 8)
        inner.setSpacing(2)

        self._number_label = QLabel("—")
        num_font = QFont("Arial", 26, QFont.Weight.Bold)
        self._number_label.setFont(num_font)
        self._number_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._title_label = QLabel()
        title_font = QFont("Arial", 9)
        self._title_label.setFont(title_font)
        self._title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._sub_label = QLabel()
        sub_font = QFont("Arial", 8)
        self._sub_label.setFont(sub_font)
        self._sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        inner.addWidget(self._number_label)
        inner.addWidget(self._title_label)
        inner.addWidget(self._sub_label)

        self._set_idle()

    def _set_idle(self) -> None:
        self.setStyleSheet(self._STYLE_IDLE)
        self._number_label.setStyleSheet("color: #94a3b8;")
        self._title_label.setStyleSheet("color: #94a3b8;")
        self._sub_label.setStyleSheet("color: #94a3b8;")
        self._number_label.setText("—")
        self._title_label.setText(
            "סוכנים ששויכו" if self._variant == "matched" else "ללא כתובת מייל"
        )
        self._sub_label.setText("")

    def set_value(self, count: int, sub: str = "") -> None:
        if self._variant == "matched":
            self.setStyleSheet(self._STYLE_MATCHED)
            self._number_label.setStyleSheet("color: #15803d;")
            self._title_label.setStyleSheet("color: #166534;")
            self._sub_label.setStyleSheet("color: #166534;")
            self._title_label.setText("סוכנים ששויכו")
        else:
            self.setStyleSheet(self._STYLE_UNMATCHED if count > 0 else self._STYLE_IDLE)
            color = "#c2410c" if count > 0 else "#94a3b8"
            self._number_label.setStyleSheet(f"color: {color};")
            self._title_label.setStyleSheet(f"color: {color};")
            self._sub_label.setStyleSheet(f"color: {color};")
            self._title_label.setText("ללא כתובת מייל")
        self._number_label.setText(str(count))
        self._sub_label.setText(sub)

    def reset(self) -> None:
        self._set_idle()


class ProcessWorker(QObject):
    finished = Signal(str)
    error = Signal(str)
    log_message = Signal(str)

    def __init__(self, processor_cls, input_path: str, output_dir: str):
        super().__init__()
        self.processor_cls = processor_cls
        self.input_path = input_path
        self.output_dir = output_dir

    @Slot()
    def run(self) -> None:
        try:
            processor = self.processor_cls(logger=self.log_message.emit)
            output_path = processor.process(self.input_path, self.output_dir)
            self.finished.emit(output_path)
        except Exception as error:
            self.error.emit(str(error))


class EmailWorker(QObject):
    finished = Signal(int)
    error = Signal(str)
    log_message = Signal(str)

    def __init__(self, output_dir: str, agent_table_path: str):
        super().__init__()
        self.output_dir = output_dir
        self.agent_table_path = agent_table_path

    @Slot()
    def run(self) -> None:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            count = prepare_agent_emails(self.output_dir, self.agent_table_path, self.log_message.emit)
            self.finished.emit(count)
        except Exception as error:
            self.error.emit(str(error))
        finally:
            pythoncom.CoUninitialize()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker_thread: Optional[QThread] = None
        self.worker: Optional[ProcessWorker] = None
        self.email_worker_thread: Optional[QThread] = None
        self.email_worker: Optional[EmailWorker] = None
        self.output_dir = ""
        self.agent_table_path = ""

        _config = load_agents_config(os.path.join(_APP_DIR, AGENTS_CONFIG_FILE_NAME))
        _saved = _config.get("agent_table_path", "")
        if _saved and os.path.exists(_saved):
            self.agent_table_path = _saved

        self.setWindowTitle(APP_TITLE)
        self.resize(950, 680)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setSpacing(12)

        title_label = QLabel(APP_TITLE)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        title_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter
        )
        layout.addWidget(
            title_label,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute,
        )

        subtitle = QLabel(
            "בחירת תיקיית פלט, ולאחר מכן טעינת קובץ \u2066BOOSTER\u2069 או \u2066GILBOA\u2069"
        )
        subtitle.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        subtitle.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter
        )
        layout.addWidget(
            subtitle,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute,
        )

        path_layout = QHBoxLayout()
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("בחר תיקיית פלט לשמירת הדוחות וההיסטוריה")
        self.output_path_edit.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        browse_button = QPushButton("בחר תיקיית פלט")
        browse_button.clicked.connect(self.choose_output_dir)
        path_layout.addWidget(browse_button)
        path_layout.addWidget(self.output_path_edit)
        layout.addLayout(path_layout)

        agent_layout = QHBoxLayout()
        self.agent_path_edit = QLineEdit()
        self.agent_path_edit.setReadOnly(True)
        self.agent_path_edit.setPlaceholderText("טבלת סוכנים לא נטענה")
        self.agent_path_edit.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.agent_path_edit.setStyleSheet(
            "background: #f0f4f8; color: #334155; border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px;"
        )
        if self.agent_table_path:
            self.agent_path_edit.setText(self.agent_table_path)
        agent_browse_button = QPushButton("קליטת טבלת סוכנים")
        agent_browse_button.clicked.connect(self.choose_agent_table)
        agent_layout.addWidget(agent_browse_button)
        agent_layout.addWidget(self.agent_path_edit, stretch=1)
        layout.addLayout(agent_layout)

        buttons_layout = QHBoxLayout()
        self.booster_button = QPushButton("קליטת קובץ בוסטר")
        self.gilboa_button = QPushButton("קליטת קובץ גילבוע")
        self.email_button = QPushButton("הכנת מיילים לסוכנים")
        self.booster_button.clicked.connect(lambda: self.select_and_run(BoosterProcessor, "קבצי Excel (*.xlsx *.xls)"))
        self.gilboa_button.clicked.connect(lambda: self.select_and_run(GilboaProcessor, "קבצי טקסט (*.txt)"))
        self.email_button.clicked.connect(self.start_email_worker)
        buttons_layout.addWidget(self.booster_button)
        buttons_layout.addWidget(self.gilboa_button)
        layout.addLayout(buttons_layout)

        stats_title = QLabel("סיכום שיוך סוכנים לכתובות מייל")
        stats_title.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        stats_title.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter
        )
        stats_title_font = QFont("Arial", 9, QFont.Weight.Bold)
        stats_title.setFont(stats_title_font)
        stats_title.setStyleSheet("color: #475569;")
        layout.addWidget(
            stats_title,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute,
        )

        stats_layout = QHBoxLayout()
        stats_layout.setSpacing(12)
        self.stat_unmatched = StatCard("unmatched")
        self.stat_matched = StatCard("matched")
        stats_layout.addWidget(self.stat_unmatched)
        stats_layout.addWidget(self.stat_matched)
        layout.addLayout(stats_layout)

        email_layout = QHBoxLayout()
        email_layout.addWidget(self.email_button)
        layout.addLayout(email_layout)

        self.status_label = QLabel("מצב: מוכן")
        self.status_label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.status_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter
        )
        self.status_label.setMinimumWidth(140)
        layout.addWidget(
            self.status_label,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute,
        )

        log_title = QLabel("לוג פעילות")
        log_title.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        log_title.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        log_title.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter
        )
        layout.addWidget(
            log_title,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute,
        )

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.log_view.setStyleSheet(
            "background-color: #10151c; color: #d8e2f0; border: 1px solid #2b3a4f;"
            "font-family: Consolas, 'Courier New', monospace; font-size: 12px;"
        )
        layout.addWidget(self.log_view)

        self.setStyleSheet(
            "QWidget { background: #f5f7fb; color: #16212f; }"
            "QPushButton { background: #1f6feb; color: white; border-radius: 6px; padding: 8px 14px; }"
            "QPushButton:disabled { background: #94a3b8; }"
            "QLineEdit { background: white; padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; }"
        )

        if self.agent_table_path:
            self.append_log(f"טבלת סוכנים נטענה אוטומטית: {self.agent_table_path}")
        self.append_log("המערכת מוכנה לעבודה")

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_view.append(f"[{timestamp}] {message}")
        _logger.info(message)

    def choose_output_dir(self) -> None:
        selected_dir = QFileDialog.getExistingDirectory(self, "בחר תיקיית פלט")
        if selected_dir:
            self.output_dir = selected_dir
            self.output_path_edit.setText(selected_dir)
            self.append_log(f"נבחרה תיקיית פלט: {selected_dir}")

    def select_and_run(self, processor_cls, file_filter: str) -> None:
        output_dir = self.output_path_edit.text().strip()
        if not output_dir:
            QMessageBox.warning(self, APP_TITLE, "יש לבחור תיקיית פלט לפני תחילת העבודה")
            return

        file_path, _ = QFileDialog.getOpenFileName(self, "בחר קובץ לעיבוד", "", file_filter)
        if not file_path:
            return

        self.append_log(f"נבחר קובץ: {file_path}")
        self.start_worker(processor_cls, file_path, output_dir)

    def set_busy(self, is_busy: bool) -> None:
        self.booster_button.setDisabled(is_busy)
        self.gilboa_button.setDisabled(is_busy)
        self.email_button.setDisabled(is_busy)
        self.status_label.setText("מצב: מעבד..." if is_busy else "מצב: מוכן")

    def start_worker(self, processor_cls, input_path: str, output_dir: str) -> None:
        if self.worker_thread is not None and self.worker_thread.isRunning():
            QMessageBox.information(self, APP_TITLE, "כבר מתבצע עיבוד, יש להמתין לסיום")
            return

        self.set_busy(True)
        self.worker_thread = QThread(self)
        self.worker = ProcessWorker(processor_cls, input_path, output_dir)
        self.worker.moveToThread(self.worker_thread)

        self.worker_thread.started.connect(self.worker.run)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.error.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.cleanup_thread)
        self.worker_thread.start()

    def cleanup_thread(self) -> None:
        self.set_busy(False)
        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None
        if self.worker_thread is not None:
            self.worker_thread.deleteLater()
            self.worker_thread = None

    def on_finished(self, output_path: str) -> None:
        self.append_log("העיבוד הסתיים בהצלחה")
        self.append_log(f"קובץ הפלט: {output_path}")
        self._refresh_stats()
        QMessageBox.information(self, APP_TITLE, f"הדוח נשמר בהצלחה:\n{output_path}")

    def on_error(self, error_message: str) -> None:
        self.append_log(f"שגיאה: {error_message}")
        _logger.error(error_message)
        QMessageBox.critical(self, APP_TITLE, error_message)

    def choose_agent_table(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ טבלת סוכנים", "", "קבצי Excel (*.xlsx *.xls)"
        )
        if not file_path:
            return
        try:
            load_agent_table(file_path)
        except Exception as exc:
            QMessageBox.critical(self, APP_TITLE, f"שגיאה בטעינת טבלת סוכנים:\n{exc}")
            return
        self.agent_table_path = file_path
        self.agent_path_edit.setText(file_path)
        save_agents_config(
            os.path.join(_APP_DIR, AGENTS_CONFIG_FILE_NAME),
            {"agent_table_path": file_path},
        )
        self.append_log(f"טבלת סוכנים נטענה: {os.path.basename(file_path)}")
        self._refresh_stats()

    def _refresh_stats(self) -> None:
        output_dir = self.output_path_edit.text().strip()
        if not output_dir or not self.agent_table_path:
            self.stat_matched.reset()
            self.stat_unmatched.reset()
            return
        try:
            agent_df = load_agent_table(self.agent_table_path)
            rows_map, unmatched_df = build_agent_rows_map(output_dir, agent_df, lambda _: None)
        except Exception:
            self.stat_matched.reset()
            self.stat_unmatched.reset()
            return
        matched_agents = len(rows_map)
        matched_records = sum(len(v) for v in rows_map.values())
        unmatched_agents = unmatched_df["_agent_key"].nunique() if not unmatched_df.empty else 0
        self.stat_matched.set_value(matched_agents, f"{matched_records} רשומות")
        self.stat_unmatched.set_value(unmatched_agents)

    def start_email_worker(self) -> None:
        output_dir = self.output_path_edit.text().strip()
        if not output_dir:
            QMessageBox.warning(self, APP_TITLE, "יש לבחור תיקיית פלט לפני הכנת המיילים")
            return
        if not self.agent_table_path:
            QMessageBox.warning(self, APP_TITLE, "יש לטעון טבלת סוכנים לפני הכנת המיילים")
            return
        if self.email_worker_thread is not None and self.email_worker_thread.isRunning():
            QMessageBox.information(self, APP_TITLE, "כבר מתבצעת הכנת מיילים, יש להמתין לסיום")
            return
        self.set_busy(True)
        self.email_worker_thread = QThread(self)
        self.email_worker = EmailWorker(output_dir, self.agent_table_path)
        self.email_worker.moveToThread(self.email_worker_thread)
        self.email_worker_thread.started.connect(self.email_worker.run)
        self.email_worker.log_message.connect(self.append_log)
        self.email_worker.finished.connect(self.on_emails_finished)
        self.email_worker.error.connect(self.on_email_error)
        self.email_worker.finished.connect(self.email_worker_thread.quit)
        self.email_worker.error.connect(self.email_worker_thread.quit)
        self.email_worker_thread.finished.connect(self.cleanup_email_thread)
        self.email_worker_thread.start()

    def cleanup_email_thread(self) -> None:
        self.set_busy(False)
        if self.email_worker is not None:
            self.email_worker.deleteLater()
            self.email_worker = None
        if self.email_worker_thread is not None:
            self.email_worker_thread.deleteLater()
            self.email_worker_thread = None

    def on_emails_finished(self, count: int) -> None:
        self.append_log(f"הכנת מיילים הסתיימה — נוצרו {count} טיוטות")
        QMessageBox.information(self, APP_TITLE, f"נוצרו {count} טיוטות מייל ב-Outlook")

    def on_email_error(self, error_message: str) -> None:
        self.append_log(f"שגיאה בהכנת מיילים: {error_message}")
        QMessageBox.critical(self, APP_TITLE, error_message)


def _handle_uncaught_exception(exc_type, exc_value, exc_tb):
    """Log unhandled exceptions to file before the process exits."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    _logger.critical("קריסה לא מטופלת", exc_info=(exc_type, exc_value, exc_tb))


def main() -> None:
    sys.excepthook = _handle_uncaught_exception
    _logger.info("=" * 60)
    _logger.info("הפעלת הכלי")
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
