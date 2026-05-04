# מדריך PySide6 עברית ו-RTL — טיפול מסודר

> מסמך זה מרכז את כל מה שנדרש להתאמת ממשק PySide6 לעברית ויישור מימין לשמאל (RTL).  
> מבוסס על ניסיון מעשי בפרויקט `voucher_app.py`.

---

## 1. הגדרת RTL גלובלית — חובה ראשונה

```python
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

def main():
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)  # ← חובה על QApplication
    ...
    app.exec()
```

**למה?**  
Qt מפיץ את `LayoutDirection` אוטומטית לכל ה-widgets שנוצרים אחרי ההגדרה.  
אם לא מגדירים על `QApplication`, כל widget יצטרך הגדרה נפרדת.

---

## 2. יישור תוויות (QLabel) — המלכודת הנפוצה ביותר

### הבעיה
```python
# ❌ לא עובד ב-RTL — Qt ממיר "Right" ל-"Left" לוגית
label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
```

### הפתרון
```python
# ✅ AlignAbsolute מונע את ה"מיררינג" הלוגי של Qt
from PySide6.QtCore import Qt

ALIGN_RTL = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter

label.setAlignment(ALIGN_RTL)
```

> **כלל אצבע:** בכל `QLabel` בממשק RTL — תמיד הוסף `AlignAbsolute`.

---

## 3. מיקום Widget בתוך Layout

גם כשמוסיפים widget ל-layout, יש לציין alignment:

```python
layout = QVBoxLayout()

# ✅ נכון — הwidget יישאר בצד ימין הפיזי
layout.addWidget(label, alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute)

# ❌ שגוי — Qt ימקם בצד שמאל כי "right" מתהפך
layout.addWidget(label, alignment=Qt.AlignmentFlag.AlignRight)
```

---

## 4. טקסט מעורב עברית + לטינית (BiDi)

### הבעיה
כשמשלבים עברית ואנגלית/מספרים בתוך מחרוזת אחת, מנגנון ה-BiDi (Bidirectional Text) של Unicode עלול להציג הכיוון הלא נכון.

### הפתרון — תוחמי LTR Isolate
```python
# \u2066 = LTR Isolate פתיחה
# \u2069 = Pop Directional Isolate (סגירה)

subtitle = f"בחירת תיקיית פלט, ולאחר מכן טעינת קובץ \u2066BOOSTER\u2069 או \u2066GILBOA\u2069"
label.setText(subtitle)
```

| תו Unicode | שם | תפקיד |
|---|---|---|
| `\u2066` | LEFT-TO-RIGHT ISOLATE | מתחיל קטע LTR בודד |
| `\u2069` | POP DIRECTIONAL ISOLATE | סוגר את הקטע |
| `\u200F` | RIGHT-TO-LEFT MARK | סמן RTL בודד (ללא בידוד) |
| `\u200E` | LEFT-TO-RIGHT MARK | סמן LTR בודד |

---

## 5. QTextEdit — יומן / שטח טקסט גדול

```python
from PySide6.QtCore import Qt

log_view = QTextEdit()
log_view.setReadOnly(True)
log_view.setLayoutDirection(Qt.LayoutDirection.RightToLeft)  # ← חשוב
```

גם בעת הוספת טקסט דינמי, הכיוון נשמר אוטומטית.

---

## 6. QLineEdit — שדות קלט

```python
from PySide6.QtCore import Qt

line_edit = QLineEdit()
line_edit.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
line_edit.setPlaceholderText("הכנס טקסט כאן...")
```

> Placeholder text גם הוא יוצג RTL אם `layoutDirection` מוגדר נכון.

---

## 7. פונטים מומלצים לעברית

```python
from PySide6.QtGui import QFont

# פונטים שתומכים בעברית בצורה טובה
font_hebrew = QFont("Arial", 11)       # Universal, נתמך בכל Windows
font_title  = QFont("Arial", 16, QFont.Weight.Bold)

label.setFont(font_hebrew)
```

**פונטים מומלצים לפי שימוש:**

| שימוש | פונט מומלץ |
|---|---|
| גוף טקסט | `Arial`, `Tahoma` |
| כותרות | `Arial Bold`, `David` |
| Monospace / יומן | `Courier New`, `Consolas` |

---

## 8. Stylesheets (CSS) — מה עובד ומה לא

Stylesheets ב-Qt **אינן מודעות ל-RTL**. הן עובדות בצורה ויזואלית בלבד.

```python
# ✅ עובד — צבעים, padding, border
widget.setStyleSheet("background-color: #f0f0f0; padding: 8px; border-radius: 6px;")

# ⚠️ לא מושפע מ-RTL — padding-right/left הם פיזיים
widget.setStyleSheet("padding-right: 12px;")  # תמיד ימין פיזי, לא לוגי
```

---

## 9. תבנית עוזר — פונקציה לבניית Label מיושר RTL

כדי לא לחזור על אותו קוד בכל מקום:

```python
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QLabel

_ALIGN_RTL = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter

def make_label(
    text: str,
    font_size: int = 11,
    bold: bool = False,
    color: str = "",
) -> QLabel:
    label = QLabel(text)
    font = QFont("Arial", font_size)
    if bold:
        font.setBold(True)
    label.setFont(font)
    label.setAlignment(_ALIGN_RTL)
    label.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    if color:
        label.setStyleSheet(f"color: {color};")
    return label
```

**שימוש:**
```python
title  = make_label("כותרת ראשית", font_size=16, bold=True)
status = make_label("מצב: מוכן", color="#555555")
```

---

## 10. QMessageBox בעברית

```python
from PySide6.QtWidgets import QMessageBox
from PySide6.QtCore import Qt

def show_rtl_message(parent, title: str, text: str, icon=QMessageBox.Icon.Information):
    msg = QMessageBox(parent)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setIcon(icon)
    msg.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    msg.exec()
```

---

## 11. QFileDialog בעברית

```python
from PySide6.QtWidgets import QFileDialog
from PySide6.QtCore import Qt

dialog = QFileDialog(self)
dialog.setWindowTitle("בחר קובץ")
dialog.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
dialog.exec()

# או בשיטה הקצרה (ללא RTL על הדיאלוג עצמו — Windows מטפל):
path, _ = QFileDialog.getOpenFileName(self, "בחר קובץ", "", "Excel Files (*.xlsx)")
```

---

## 12. רשימת בדיקה — Checklist לפרויקט חדש

```
[ ] app.setLayoutDirection(Qt.LayoutDirection.RightToLeft) ב-main()
[ ] כל QLabel עם AlignRight | AlignAbsolute | AlignVCenter
[ ] כל layout.addWidget עם alignment=AlignRight|AlignAbsolute כשנדרש
[ ] QTextEdit / QLineEdit — setLayoutDirection(RightToLeft)
[ ] טקסט מעורב עברית+אנגלית — עטוף עם \u2066...\u2069
[ ] פונט: Arial או David לתמיכה מלאה בעברית
[ ] QMessageBox — setLayoutDirection(RightToLeft)
```

---

## 13. בעיות נפוצות ופתרונות

| בעיה | סיבה | פתרון |
|---|---|---|
| Label מיושר שמאלה למרות AlignRight | Qt ממיר Right→Left ב-RTL | הוסף `AlignAbsolute` |
| אנגלית מוצגת בכיוון הפוך בתוך עברית | BiDi אוטומטי של Unicode | עטוף ב-`\u2066...\u2069` |
| Widget "קופץ" לשמאל ב-Layout | Layout לא מקבל alignment | `layout.addWidget(w, alignment=AlignRight|AlignAbsolute)` |
| Placeholder text בשמאל | LayoutDirection לא הוגדר על widget | `widget.setLayoutDirection(RightToLeft)` |
| כפתורים הפוכים (icon בצד הלא נכון) | ירושת LayoutDirection | הגדרה גלובלית על QApplication |

---

*נוצר: 04/05/2026 | פרויקט: Voucher Tool — Ophir Tours*
