import os
from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
EXCEL_FILE_PATH = os.getenv(
    "EXCEL_PATH",
    r"C:\Users\prath\Desktop\Student_Expense_Tracker.xlsx"
)
MODEL = "claude-sonnet-4-5"
MONTHLY_BUDGET = 650.0
