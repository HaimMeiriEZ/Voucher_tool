# Voucher Tool

Desktop tool for managing and validating voucher reports for Booster and Gilboa.

## Features
- Hebrew RTL desktop UI (PySide6)
- Manual loading for Booster and Gilboa source files
- Output folder selection by the user
- Embedded activity log during processing
- Unified Excel export — split by source (מקור) into separate sheets, empty columns removed automatically
- Agent email preparation via Outlook (win32com):
  - **Matched agents** (email found in `agents_config.json`): Outlook draft created in the configured subfolder, with up to 3 Excel attachments (all open orders, upcoming departures, at-risk cancellations) and CC to branch/department managers
  - **Unmatched agents** (no email): `.msg` files saved to `ממתינות מייל - כתובת חסרה/` subfolder inside the output folder, with CC to the operations team; subfolder is created automatically if it does not exist
- Run state persisted in `voucher_state.json` (auto-created on first run)

## Dependencies
- `pandas>=2.2`
- `openpyxl>=3.1`
- `PySide6>=6.8`
- `pywin32>=306` — required for Outlook COM automation (`.msg` file creation and draft saving)

## Run
1. Create or activate a Python virtual environment
2. Install dependencies: `pip install -r requirements.txt`
3. Run `voucher_app.py` directly or use `start_voucher_tool.bat`

## Configuration
- `agents_config.json` — maps agent names/keys to email addresses and routing metadata (not committed to source control)
- `voucher_state.json` — persists last-used file paths between sessions (auto-created, not committed)
