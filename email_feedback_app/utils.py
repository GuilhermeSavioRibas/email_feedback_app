import json
import os
import pandas as pd
from datetime import datetime
from tkinter import messagebox
from string import Template
from typing import Dict, List, Optional, Any
import html


def load_analysts_config() -> Dict[str, Any]:
    try:
        with open("config/analysts.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load analysts config: {e}")
        return {}

def get_email_config(account: str, analysts_config: Dict[str, Any]) -> Dict[str, Any]:
    return analysts_config.get(account, {})

def is_valid_feedback(message: Optional[str]) -> bool:
    return message and str(message).strip().lower() not in ["none", "", ".", "n/a"]

def load_existing_log_entries() -> pd.DataFrame:
    log_file = os.path.join("logs", "approved_feedbacks.xlsx")
    if os.path.exists(log_file):
        return pd.read_excel(log_file)
    return pd.DataFrame()

def filter_and_process_feedbacks(feedback_data: Dict[str, List[Dict[str, Any]]]) -> Dict[str, List[Dict[str, Any]]]:
    log_file = os.path.join("logs", "approved_feedbacks.xlsx")
    processed_keys = set()

    if os.path.exists(log_file):
        try:
            log_df = pd.read_excel(log_file)
            for _, row in log_df.iterrows():
                processed_keys.add((row["Account"], str(row["TicketID"])))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read log file: {e}")

    filtered_data = {}
    for account, entries in feedback_data.items():
        filtered_data[account] = []
        for entry in entries:
            ticket_id = str(entry.get("ticket_id"))
            if is_valid_feedback(entry.get("message")) and (account, ticket_id) not in processed_keys:
                filtered_data[account].append(entry)

    return filtered_data

def save_to_log(feedbacks: List[Dict[str, Any]], status: str = "Approved") -> bool:
    log_file = os.path.join("logs", "approved_feedbacks.xlsx")
    try:
        existing = pd.read_excel(log_file) if os.path.exists(log_file) else pd.DataFrame()

        new_data = pd.DataFrame([{
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Account": fb["account"],
            "TicketID": fb["ticket_id"],
            "UserName": fb["user_name"],
            "AnalystName": fb["analyst_name"],
            "Message": fb["message"],
            "Status": status
        } for fb in feedbacks])

        combined = pd.concat([existing, new_data], ignore_index=True)
        combined.drop_duplicates(subset=["Account", "TicketID"], inplace=True)

        combined.to_excel(log_file, index=False)
        return True

    except Exception as e:
        messagebox.showerror("Error", f"Failed to save to approved_feedbacks.xlsx: {e}")
        return False

def render_html_template(feedback: Dict[str, Any]) -> tuple[str, str]:
    try:
        with open("config/config.json", "r", encoding="utf-8") as f:
            config_data = json.load(f)

        template_language = config_data.get("template_language", "portuguese")
        if template_language == "english":
            template_file = "email_template_en.html"
        elif template_language == "spanish":
            template_file = "email_template_es.html"
        else:
            template_file = "email_template.html"

        template_path = os.path.join("templates", template_file)
        with open(template_path, "r", encoding="utf-8") as file:
            html_template = Template(file.read())

        raw_message = feedback.get("message", "")
        sanitized_message = html.escape(str(raw_message))

        filled_html = html_template.safe_substitute(
            analyst_name=feedback.get("analyst_name", ""),
            user_name=feedback.get("user_name", ""),
            message=sanitized_message,
            ticket_id=feedback.get("ticket_id", ""),
            header_img_path=os.path.abspath("templates/assets/header.png"),
            winner_img_path=os.path.abspath("templates/assets/Award-Winner.png"),
        )

        return filled_html, template_language

    except Exception as e:
        raise RuntimeError(f"Failed to render email template: {e}")

def generate_outlook_emails(feedbacks: List[Dict[str, Any]], analysts_config: Dict[str, Any],
                            default_sender: str) -> bool:
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")

        for feedback in feedbacks:
            account = feedback["account"]
            analyst_name = feedback["analyst_name"]
            email_config = analysts_config.get(account, {})
            groups = email_config.get("groups", {})

            analyst_email = None
            cc_emails = []

            for group_data in groups.values():
                analysts = group_data.get("analysts", {})
                if analyst_name in analysts:
                    analyst_email = analysts[analyst_name]
                    cc_emails = group_data.get("cc_emails", [])
                    break

            mail = outlook.CreateItem(0)
            mail.SentOnBehalfOfName = default_sender

            html_body, template_language = render_html_template(feedback)

            if template_language == "english":
                mail.Subject = f"[{account}] Recognition of Excellent Service"
            elif template_language == "spanish":
                mail.Subject = f"[{account}] Reconocimiento de Servicio Excelente"
            else:
                mail.Subject = f"[{account}] Reconhecimento de Excelente Atendimento"

            mail.To = analyst_email or ""
            mail.CC = "; ".join(cc_emails) if cc_emails else ""
            mail.HTMLBody = html_body

            mail.Save()

        return True

    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate Outlook drafts: {e}")
        return False

def load_config(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_config(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        