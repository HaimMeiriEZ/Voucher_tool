from __future__ import annotations

import glob
import json
import os
import re
import sys
from datetime import datetime
from typing import Callable, Dict, Optional

import pandas as pd
from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
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
NOTE_COLUMNS = ["הערות", "הערות סוכן", "תאריך עידכון הערת סוכן"]


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
        return datetime.fromtimestamp(os.path.getmtime(input_path))

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
        df.to_excel(output_path, index=False)
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
        file_name = os.path.basename(input_path)
        date_match = re.search(r"(\d{1,2}\s\w{3}\s\d{4})", file_name)
        if date_match:
            try:
                return datetime.strptime(date_match.group(1), "%d %b %Y")
            except ValueError:
                pass
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
        df["מקור"] = self.source_name
        df["מזהה רשומה"] = df["T. File No."].apply(normalize_identifier)

        calc_columns = []
        for _, row in df.iterrows():
            record_id = normalize_identifier(row["T. File No."])
            start_date = row["Start Date"] if pd.notna(row["Start Date"]) else None
            if start_date is not None:
                state[f"{self.source_name}:{record_id}"] = {"open_date": str(start_date)}

            is_direct = safe_text(row["Agent/C. Client"]) == "Direct Sale"
            days_to_departure = (start_date - anchor_date).days if (is_direct and start_date is not None) else ""

            state_entry = state.get(f"{self.source_name}:{record_id}", {})
            open_date = safe_datetime(state_entry.get("open_date", ""))
            days_since_refund = (anchor_date - open_date).days if open_date else 0

            refund_value = pd.to_numeric(row.get("Unconfirmed Refund", 0), errors="coerce")
            refund_value = 0 if pd.isna(refund_value) else float(refund_value)
            refund_note = ""
            if refund_value > 0:
                refund_note = "החזר שטרם אושר מעל 60 יום" if days_since_refund > 60 else "החזר שטרם אושר"

            if is_direct:
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
                ]
            )

        df[
            [
                "מס' ימים ליציאה",
                "מס' ימי החזר",
                "הערות",
                "הערות סוכן",
                "תאריך עידכון הערת סוכן",
            ]
        ] = pd.DataFrame(calc_columns, index=df.index)

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
                self.ensure_columns(df, ["Number", "Open", "Start", "Ref wl", "C.Client"])
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
        df["מקור"] = self.source_name
        df["מזהה רשומה"] = df["Number"].apply(normalize_identifier)

        calc_columns = []
        for _, row in df.iterrows():
            record_id = normalize_identifier(row["Number"])
            open_date = safe_datetime(row.get("Open", ""), dayfirst=True)
            start_date = safe_datetime(row.get("Start", ""), dayfirst=True)
            if open_date is not None:
                state[f"{self.source_name}:{record_id}"] = {"open_date": str(open_date)}

            days_to_departure = (start_date - anchor_date).days if start_date else ""
            days_since_refund = (anchor_date - open_date).days if open_date else 0

            raw_refund = safe_text(row.get("Ref wl", "")).replace(",", "")
            try:
                refund_value = float(raw_refund) if raw_refund else 0.0
            except ValueError:
                refund_value = 0.0

            refund_note = ""
            if refund_value != 0:
                refund_note = "החזר שטרם אושר מעל 60 יום" if days_since_refund > 60 else "החזר שטרם אושר"

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
                ]
            )

        df[
            [
                "מס' ימים ליציאה",
                "מס' ימי החזר",
                "הערות",
                "הערות סוכן",
                "תאריך עידכון הערת סוכן",
            ]
        ] = pd.DataFrame(calc_columns, index=df.index)

        preferred = [
            "מקור",
            "מזהה רשומה",
            "Number",
            "Open",
            "Start",
            "Ref wl",
            "C.Client",
            "מס' ימים ליציאה",
            "מס' ימי החזר",
            "הערות",
            "הערות סוכן",
            "תאריך עידכון הערת סוכן",
        ]
        other_columns = [column for column in df.columns if column not in preferred]
        return df[preferred + other_columns]


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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.thread: Optional[QThread] = None
        self.worker: Optional[ProcessWorker] = None
        self.output_dir = ""

        self.setWindowTitle(APP_TITLE)
        self.resize(950, 680)
        self.setLayoutDirection(Qt.RightToLeft)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setSpacing(12)

        title_label = QLabel(APP_TITLE)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        layout.addWidget(title_label)

        subtitle = QLabel("בחירת תיקיית פלט, ולאחר מכן טעינת קובץ BOOSTER או GILBOA")
        layout.addWidget(subtitle)

        path_layout = QHBoxLayout()
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("בחר תיקיית פלט לשמירת הדוחות וההיסטוריה")
        browse_button = QPushButton("בחר תיקיית פלט")
        browse_button.clicked.connect(self.choose_output_dir)
        path_layout.addWidget(self.output_path_edit)
        path_layout.addWidget(browse_button)
        layout.addLayout(path_layout)

        buttons_layout = QHBoxLayout()
        self.booster_button = QPushButton("קליטת קובץ BOOSTER")
        self.gilboa_button = QPushButton("קליטת קובץ GILBOA")
        self.booster_button.clicked.connect(lambda: self.select_and_run(BoosterProcessor, "קבצי Excel (*.xlsx *.xls)"))
        self.gilboa_button.clicked.connect(lambda: self.select_and_run(GilboaProcessor, "קבצי טקסט (*.txt)"))
        buttons_layout.addWidget(self.booster_button)
        buttons_layout.addWidget(self.gilboa_button)
        layout.addLayout(buttons_layout)

        self.status_label = QLabel("מצב: מוכן")
        layout.addWidget(self.status_label)

        log_title = QLabel("לוג פעילות")
        log_title.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        layout.addWidget(log_title)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
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

        self.append_log("המערכת מוכנה לעבודה")

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_view.append(f"[{timestamp}] {message}")

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
        self.status_label.setText("מצב: מעבד..." if is_busy else "מצב: מוכן")

    def start_worker(self, processor_cls, input_path: str, output_dir: str) -> None:
        if self.thread is not None and self.thread.isRunning():
            QMessageBox.information(self, APP_TITLE, "כבר מתבצע עיבוד, יש להמתין לסיום")
            return

        self.set_busy(True)
        self.thread = QThread(self)
        self.worker = ProcessWorker(processor_cls, input_path, output_dir)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.thread.finished.connect(self.cleanup_thread)
        self.thread.start()

    def cleanup_thread(self) -> None:
        self.set_busy(False)
        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None
        if self.thread is not None:
            self.thread.deleteLater()
            self.thread = None

    def on_finished(self, output_path: str) -> None:
        self.append_log("העיבוד הסתיים בהצלחה")
        self.append_log(f"קובץ הפלט: {output_path}")
        QMessageBox.information(self, APP_TITLE, f"הדוח נשמר בהצלחה:\n{output_path}")

    def on_error(self, error_message: str) -> None:
        self.append_log(f"שגיאה: {error_message}")
        QMessageBox.critical(self, APP_TITLE, error_message)


def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
