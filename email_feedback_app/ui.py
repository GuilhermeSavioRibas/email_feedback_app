import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from email_feedback_app.settings_window import SettingsWindow
import os
from datetime import datetime
from email_feedback_app.utils import (
    load_analysts_config,
    filter_and_process_feedbacks,
    get_email_config,
    save_to_log,    
    generate_outlook_emails
)
from email_feedback_app.processor import process_feedbacks

class FeedbackApp:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.selected_account = tk.StringVar()
        self.current_page = 1
        self.items_per_page = 20

        self.raw_feedbacks = self.load_all_feedbacks()
        self.all_feedbacks = filter_and_process_feedbacks(self.raw_feedbacks)

        self.analysts_config = load_analysts_config()

        self.setup_ui()
        self.setup_styles()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#003134")
        self.style.configure("TLabel", background="#003134", foreground="white")
        self.style.configure("TButton", background="#00E28B", foreground="black")
        self.style.map("TButton",
                       foreground=[('pressed', 'black'), ('active', 'black')],
                       background=[('pressed', '#00C47D'), ('active', '#00E28B')])

    def load_all_feedbacks(self):
        all_data = {}
        for file in os.listdir("data"):
            if file.endswith(".xlsx"):
                filepath = os.path.join("data", file)
                account_name = os.path.splitext(file)[0]
                feedbacks = process_feedbacks(filepath, self.config)
                if feedbacks:
                    all_data[account_name] = feedbacks
                else:
                    print(f"[!] No feedback loaded for: {account_name}")
        return all_data

    def setup_ui(self):
        self.root.title("Kudos Manager")
        self.root.geometry("1200x600")

        top_frame = ttk.Frame(self.root)
        top_frame.pack(pady=10, fill="x", padx=10)

        ttk.Label(top_frame, text="Select account:").pack(side="left", padx=5)
        account_list = list(self.all_feedbacks.keys())
        self.account_dropdown = ttk.Combobox(top_frame, textvariable=self.selected_account, values=account_list, state="readonly", width=30)
        self.account_dropdown.pack(side="left", padx=5)

        ttk.Button(top_frame, text="Load Feedbacks", command=self.load_feedbacks).pack(side="left", padx=5)
        ttk.Button(top_frame, text="Refresh Data", command=self.refresh_data).pack(side="left", padx=5)

        self.setup_table()
        self.setup_pagination()

        self.generate_btn = ttk.Button(self.root, text="Generate Emails", command=self.generate_emails, state="disabled")
        self.generate_btn.pack(pady=10)

        settings_btn = ttk.Button(top_frame, text="⚙️ Settings", command=self.open_settings)
        settings_btn.pack(side=tk.RIGHT, padx=10)

    def refresh_data(self):
        self.root.config(cursor="wait")
        self.root.update()
        try:
            self.raw_feedbacks = self.load_all_feedbacks()
            self.all_feedbacks = filter_and_process_feedbacks(self.raw_feedbacks)
            # Atualizar o dropdown de contas
            account_list = list(self.all_feedbacks.keys())
            self.account_dropdown.config(values=account_list)
            if account_list and not self.selected_account.get():
                self.selected_account.set(account_list[0])
            # Recarregar os feedbacks exibidos
            self.load_feedbacks()
        finally:
            self.root.config(cursor="")

    def open_settings(self):
        SettingsWindow(
            master=self.root,
            config_path="config/config.json",
            analysts_path="config/analysts.json"
        )

    def setup_table(self):
        self.tree = ttk.Treeview(self.root, columns=("Ticket", "User", "Message", "Analyst", "Action"), show="headings", height=20, selectmode="browse")
        columns_config = [
            ("Ticket", "Ticket ID", 120),
            ("User", "User Name", 180),
            ("Message", "Message (Double-click to edit)", 500),
            ("Analyst", "Analyst (Double-click to edit)", 200),
            ("Action", "Action", 120)
        ]
        for col, heading, width in columns_config:
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", self.on_delete_key)

    def setup_pagination(self):
        pagination_frame = ttk.Frame(self.root)
        pagination_frame.pack(pady=5)
        ttk.Button(pagination_frame, text="◄ Previous", command=self.prev_page).pack(side="left", padx=5)
        self.page_label = ttk.Label(pagination_frame, text="Page 1")
        self.page_label.pack(side="left", padx=5)
        ttk.Button(pagination_frame, text="Next ►", command=self.next_page).pack(side="left", padx=5)

    def load_feedbacks(self):
        account = self.selected_account.get()
        if not account:
            messagebox.showerror("Error", "Please select an account.")
            return
        if account not in self.raw_feedbacks:
            messagebox.showerror("Error", f"No data available for account: {account}")
            return

        filtered = filter_and_process_feedbacks({account: self.raw_feedbacks[account]})
        self.all_feedbacks[account] = filtered.get(account, [])

        self.current_page = 1
        self.display_feedbacks()
        self.generate_btn.config(state="normal")

    def display_feedbacks(self):
        account = self.selected_account.get()
        feedbacks = self.all_feedbacks.get(account, [])
        self.tree.delete(*self.tree.get_children())
        total_pages = max(1, (len(feedbacks) + self.items_per_page - 1) // self.items_per_page)
        if self.current_page > total_pages:
            self.current_page = total_pages
        start_idx = (self.current_page - 1) * self.items_per_page
        end_idx = start_idx + self.items_per_page
        paginated_feedbacks = feedbacks[start_idx:end_idx]

        for entry in paginated_feedbacks:
            original_name = entry.get("analyst_name", "")
            analyst_email = self.get_analyst_email(account, original_name)
            display_name = (original_name or "[Unknown Analyst]") + " ⚠️" if not analyst_email else (original_name or "[Unknown Analyst]")

            self.tree.insert("", "end", values=(
                entry.get("ticket_id", ""),
                entry.get("user_name", ""),
                entry.get("message", ""),
                display_name,
                "Delete"
            ))

        self.page_label.config(text=f"Page {self.current_page} of {total_pages}")

    def on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            item = self.tree.focus()
            col_name = self.get_column_name(column)
            if col_name == "Action":
                self.delete_row(item)
            else:
                self.edit_cell(item, column, col_name)

    def on_delete_key(self, event):
        item = self.tree.focus()
        if item:
            self.delete_row(item)

    def get_column_name(self, column_id):
        columns = ["Ticket", "User", "Message", "Analyst", "Action"]
        col_index = int(column_id[1:]) - 1
        return columns[col_index] if col_index < len(columns) else ""

    def edit_cell(self, item, column, col_name):
        current_values = list(self.tree.item(item, "values"))
        col_index = int(column[1:]) - 1
        current_text = current_values[col_index]

        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"Edit {col_name}")
        edit_window.geometry("600x300")
        edit_window.grab_set()

        self.root.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()

        w = 600
        h = 300
        x = root_x + (root_width // 2) - (w // 2)
        y = root_y + (root_height // 2) - (h // 2)
        edit_window.geometry(f"{w}x{h}+{x}+{y}")

        edit_window.rowconfigure(1, weight=1) 
        edit_window.columnconfigure(0, weight=1)

        label = ttk.Label(edit_window, text=f"{col_name}:", anchor="w")
        label.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))

        text_frame = ttk.Frame(edit_window)
        text_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 5))
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        text_widget = tk.Text(text_frame, wrap="word", font=("Arial", 12))
        text_widget.grid(row=0, column=0, sticky="nsew")
        text_widget.insert("1.0", current_text)

        save_button = ttk.Button(edit_window, text="Save", command=lambda: save_edit())
        save_button.grid(row=2, column=0, pady=(0, 10))

        def save_edit():
            new_value = text_widget.get("1.0", "end").strip()
            if new_value:
                current_values[col_index] = new_value
                self.tree.item(item, values=current_values)
            edit_window.destroy()

    def delete_row(self, item):
        if messagebox.askyesno("Confirm", "Delete this feedback?"):
            values = self.tree.item(item, "values")
            if len(values) < 4:
                return

            account = self.selected_account.get()
            analyst_name_clean = values[3].replace(" ⚠️", "").strip()

            rejected_feedback = {
                "ticket_id": values[0],
                "user_name": values[1],
                "message": values[2],
                "analyst_name": analyst_name_clean,
                "account": account,
                "status": "Rejected",
                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            save_to_log([rejected_feedback], status="Rejected")
            self.tree.delete(item) 

    def generate_emails(self):
        feedbacks_to_export = []
        account = self.selected_account.get()

        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if len(values) < 4:
                continue 
            feedback = {
                "ticket_id": values[0],
                "user_name": values[1],
                "message": values[2],
                "analyst_name": values[3].replace("⚠️", "").strip(), 
                "account": account
            }
            feedbacks_to_export.append(feedback)

        if not feedbacks_to_export:
            messagebox.showwarning("Warning", "No feedbacks to generate emails.")
            return

        if not save_to_log(feedbacks_to_export):
            return

        success = generate_outlook_emails(
            feedbacks_to_export,
            self.analysts_config,
            self.config.get("default_sender_email", "")
        )

        if success:
            self.load_feedbacks()
            messagebox.showinfo("Success", f"Generated {len(feedbacks_to_export)} email drafts in Outlook!")
        else:
            messagebox.showerror("Error", "Failed to generate emails. Check the logs for details.")
    
    def get_analyst_email(self, account, analyst_name):
        account_config = self.analysts_config.get(account, {})
        groups = account_config.get("groups", {})
        for group in groups.values():
            if analyst_name in group.get("analysts", {}):
                return group["analysts"][analyst_name]
        return None

    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.display_feedbacks()

    def next_page(self):
        account = self.selected_account.get()
        if not account:
            return
        feedbacks = self.all_feedbacks.get(account, [])
        total_pages = max(1, (len(feedbacks) + self.items_per_page - 1) // self.items_per_page)
        if self.current_page < total_pages:
            self.current_page += 1
            self.display_feedbacks()
