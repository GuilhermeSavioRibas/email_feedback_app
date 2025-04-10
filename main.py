# main.py
import os
import json
from email_feedback_app.ui import FeedbackApp
import tkinter as tk
import warnings

warnings.filterwarnings("ignore", category=UserWarning, message="Workbook contains no default style, apply openpyxl's default")

os.makedirs("data", exist_ok=True)
os.makedirs("logs", exist_ok=True)

CONFIG_PATH = "config/config.json"
DATA_DIR = "data"

def main():
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        config = json.load(f)

    root = tk.Tk()
    app = FeedbackApp(root, config)
    root.mainloop()

if __name__ == "__main__":
    main()