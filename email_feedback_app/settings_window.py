import json
import tkinter as tk
from tkinter import ttk, messagebox

class SettingsWindow(tk.Toplevel):
    def __init__(self, master, config_path, analysts_path):
        super().__init__(master)
        self.withdraw()
        self.title("Settings")
        self.config_path = config_path
        self.analysts_path = analysts_path
        self.geometry("800x600")

        warning_label = tk.Label(
            self,
            text="⚠️ Warning: Changing settings may affect the system's behavior.",
            bg="#fff8dc", fg="red", font=("Segoe UI", 10, "bold"), padx=10, pady=5, anchor="w"
        )
        warning_label.pack(fill=tk.X)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.init_config_tab()
        self.init_account_tab()
        self.init_analysts_tab()

        self.center_window()
        self.deiconify()

    def load_config(self):
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config file:\n{e}")
            return {}

        if "template_language" not in config_data:
            config_data["template_language"] = "portuguese"
            self.save_config(config_data)

        return config_data

    def save_config(self, data):
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)

    def init_config_tab(self):
        config_tab = ttk.Frame(self.notebook)
        self.notebook.add(config_tab, text="Config Settings")

        try:
            config_data = self.load_config()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config file:\n{e}")
            return

        main_frame = ttk.Frame(config_tab)
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        content_frame = ttk.Frame(main_frame)
        content_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)  
        content_frame.rowconfigure(1, weight=0)  
        content_frame.rowconfigure(2, weight=0)  
        content_frame.rowconfigure(3, weight=0)  
        content_frame.rowconfigure(4, weight=1)

        style = ttk.Style()
        style.configure("DarkFrame.TFrame", background="#003134")
        style.configure("DarkLabel.TLabel", background="#003134", foreground="white")
        style.configure("DarkCombobox.TCombobox", fieldbackground="#2E5E54", background="#2E5E54", foreground="white")

        def add_centered_field(label_text, var=None, widget_type="entry", values=None, row=None):
            container = ttk.Frame(content_frame)
            container.grid(row=row, column=0, pady=10, sticky="ew")
            container.configure(style="DarkFrame.TFrame")
            container.columnconfigure(0, weight=1)
            container.columnconfigure(1, weight=1)
            label = ttk.Label(container, text=label_text, style="DarkLabel.TLabel")
            label.grid(row=0, column=0, padx=5, sticky="e")
            if widget_type == "entry":
                entry = ttk.Entry(container, textvariable=var, width=40)
                entry.grid(row=0, column=1, padx=5, sticky="w")
                return entry
            elif widget_type == "combobox":
                combo = ttk.Combobox(container, textvariable=var, values=values, state="readonly", width=37, style="DarkCombobox.TCombobox")
                combo.grid(row=0, column=1, padx=5, sticky="w")
                current_value = var.get()
                if current_value in values:
                    combo.current(values.index(current_value))
                else:
                    combo.current(0)
                return combo

        self.default_email_var = tk.StringVar(value=config_data.get("default_sender_email", "teste@example.com"))  # Valor padrão para teste
        add_centered_field("Default Sender Email:", self.default_email_var, row=1)

        self.template_language_var = tk.StringVar(value=config_data.get("template_language", "portuguese"))
        add_centered_field("Template Language:", self.template_language_var, "combobox", ["portuguese", "english", "spanish"], row=2)

        def save_config_settings():
            config_data = self.load_config()
            config_data["default_sender_email"] = self.default_email_var.get().strip()
            config_data["template_language"] = self.template_language_var.get()
            self.save_config(config_data)

            if messagebox.askyesno("Success", "Config settings updated!\n\n⚠ Some changes may require a restart.\n\nDo you want to close the settings window now?"):
                self.destroy()
            else:
                self.lift()

        button_container = ttk.Frame(content_frame)
        button_container.grid(row=3, column=0, pady=10, sticky="ew")
        button_container.configure(style="DarkFrame.TFrame")
        button_container.columnconfigure(0, weight=1)

        ttk.Button(button_container, text="Save", command=save_config_settings).grid(row=0, column=0)
        
    def init_account_tab(self):
        account_tab = ttk.Frame(self.notebook)
        self.notebook.add(account_tab, text="Account Settings")

        try:
            config_data = self.load_config()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config file:\n{e}")
            return

        self.accounts_data = config_data.get("accounts", {})
        account_list = list(self.accounts_data.keys())

        self.account_dropdown = ttk.Combobox(account_tab, values=account_list, state="readonly")
        self.account_dropdown.set("Select an account")
        self.account_dropdown.pack(pady=10)

        self.account_fields_frame = ttk.Frame(account_tab)
        self.account_fields_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        button_frame = ttk.Frame(account_tab)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="Add Account", command=self.open_add_account_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remove Account", command=self.remove_selected_account).pack(side=tk.LEFT, padx=5)

        self.field_vars = {}

        def show_account_fields(event=None):
            for widget in self.account_fields_frame.winfo_children():
                widget.destroy()

            selected_account = self.account_dropdown.get()
            if not selected_account or selected_account == "Select an account":
                return

            account_config = self.accounts_data.get(selected_account, {})
            self.field_vars.clear()

            for key, value in account_config.items():
                ttk.Label(self.account_fields_frame, text=key).pack(anchor='w', pady=2)
                var = tk.StringVar(value=json.dumps(value) if isinstance(value, (dict, list)) else str(value))
                entry = ttk.Entry(self.account_fields_frame, textvariable=var, width=80)
                entry.pack(fill=tk.X, pady=2)
                self.field_vars[key] = var

            ttk.Button(
                self.account_fields_frame,
                text="Save",
                command=lambda: self.save_account_fields(selected_account)
            ).pack(pady=10)

        self.account_dropdown.bind("<<ComboboxSelected>>", show_account_fields)

    def open_add_account_window(self):
        add_win = tk.Toplevel(self)
        add_win.title("Add New Account")
        add_win.geometry("600x500")
        add_win.withdraw()

        add_win.update_idletasks()
        width = 600
        height = 500
        x = (add_win.winfo_screenwidth() // 2) - (width // 2)
        y = (add_win.winfo_screenheight() // 2) - (height // 2)
        add_win.geometry(f"{width}x{height}+{x}+{y}")

        add_win.transient(self)
        add_win.grab_set()

        canvas = tk.Canvas(add_win, highlightthickness=0, bg="#003134")
        scrollbar = ttk.Scrollbar(add_win, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.configure(style="DarkFrame.TFrame")

        style = ttk.Style()
        style.configure("DarkFrame.TFrame", background="#003134")
        style.configure("DarkLabel.TLabel", background="#003134", foreground="white")
        style.configure("DarkCombobox.TCombobox", fieldbackground="#2E5E54", background="#2E5E54", foreground="white")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set, bg="#003134")

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)  # Windows
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # Linux/Mac (scroll up)
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))   # Linux/Mac (scroll down)

        def on_closing():
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
            add_win.destroy()

        add_win.protocol("WM_DELETE_WINDOW", on_closing)

        def add_centered_field(label_text, var=None, widget_type="entry", values=None, parent=None):
            container = ttk.Frame(parent if parent else scrollable_frame)
            container.pack(pady=10, fill="x", expand=True)
            container.configure(style="DarkFrame.TFrame")
            label = ttk.Label(container, text=label_text, style="DarkLabel.TLabel")
            label.pack(expand=True)  
            if widget_type == "entry":
                entry = ttk.Entry(container, textvariable=var, width=40)
                entry.pack(expand=True)  
                return entry
            elif widget_type == "combobox":
                combo = ttk.Combobox(container, textvariable=var, values=values, state="readonly", width=37, style="DarkCombobox.TCombobox")
                combo.pack(expand=True)  
                return combo

        account_name_var = tk.StringVar()
        add_centered_field("Account Name:", account_name_var)

        sheet_name_var = tk.StringVar()
        add_centered_field("Sheet Name:", sheet_name_var)

        header_row_var = tk.StringVar()
        add_centered_field("Header Row:", header_row_var)

        ticket_id_var = tk.StringVar()
        add_centered_field("Ticket ID Column:", ticket_id_var)

        message_var = tk.StringVar()
        add_centered_field("Message Column:", message_var)

        analyst_name_var = tk.StringVar()
        add_centered_field("Analyst Name Column:", analyst_name_var)

        user_name_type_var = tk.StringVar(value="user_name")
        user_name_dropdown = add_centered_field("User Name Type:", user_name_type_var, "combobox", ["user_name", "user_name_parts"])

        user_name_frame = ttk.Frame(scrollable_frame)
        user_name_frame.pack(pady=10, fill="x", expand=True)
        user_name_frame.configure(style="DarkFrame.TFrame")
        user_name_vars = {}

        def update_user_name_fields(*args):
            for widget in user_name_frame.winfo_children():
                widget.destroy()
            user_name_type = user_name_type_var.get()
            if user_name_type == "user_name":
                add_centered_field("User Name Column:", user_name_vars.setdefault("user_name", tk.StringVar()), parent=user_name_frame)
            elif user_name_type == "user_name_parts":
                add_centered_field("User Name Parts (comma-separated columns):", user_name_vars.setdefault("user_name_parts", tk.StringVar()), parent=user_name_frame)

        user_name_type_var.trace_add("write", update_user_name_fields)
        update_user_name_fields()

        rating_type_var = tk.StringVar(value="rating")
        rating_dropdown = add_centered_field("Rating Type:", rating_type_var, "combobox", ["rating", "rating_text", "rating_inverted"])

        rating_frame = ttk.Frame(scrollable_frame)
        rating_frame.pack(pady=10, fill="x", expand=True)
        rating_frame.configure(style="DarkFrame.TFrame")
        rating_vars = {}

        def update_rating_fields(*args):
            for widget in rating_frame.winfo_children():
                widget.destroy()
            rating_type = rating_type_var.get()
            if rating_type == "rating":
                add_centered_field("Rating Column:", rating_vars.setdefault("rating", tk.StringVar()), parent=rating_frame)
            elif rating_type == "rating_text":
                add_centered_field("Positive Value:", rating_vars.setdefault("positive_value", tk.StringVar()), parent=rating_frame)
                add_centered_field("Column:", rating_vars.setdefault("column", tk.StringVar()), parent=rating_frame)
            elif rating_type == "rating_inverted":
                add_centered_field("Valid Values (comma-separated):", rating_vars.setdefault("valid_values", tk.StringVar()), parent=rating_frame)
                add_centered_field("Column:", rating_vars.setdefault("column", tk.StringVar()), parent=rating_frame)

        rating_type_var.trace_add("write", update_rating_fields)
        update_rating_fields()

        assignment_column_var = tk.StringVar()
        add_centered_field("Assignment Group Column (optional):", assignment_column_var)

        required_values_var = tk.StringVar()
        add_centered_field("Required Values (comma-separated, optional):", required_values_var)

        analyst_groups_var = tk.StringVar()
        add_centered_field("Analyst Groups (comma-separated):", analyst_groups_var)

        def confirm_add():
            account_name = account_name_var.get().strip()
            if not account_name:
                messagebox.showerror("Error", "Account name is required.")
                return
            if not ticket_id_var.get() or not message_var.get() or not analyst_name_var.get():
                messagebox.showerror("Error", "Ticket ID, Message, and Analyst Name columns are required.")
                return

            config_data = self.load_config()
            with open(self.analysts_path, 'r', encoding='utf-8') as f:
                analysts_data = json.load(f)

            new_account_config = {
                "sheet_name": sheet_name_var.get(),
                "header_row": int(header_row_var.get()) if header_row_var.get().isdigit() else 1,
                "ticket_id": ticket_id_var.get(),
                "message": message_var.get(),
                "analyst_name": analyst_name_var.get()
            }

            user_name_type = user_name_type_var.get()
            if user_name_type == "user_name" and user_name_vars["user_name"].get():
                new_account_config["user_name"] = user_name_vars["user_name"].get()
            elif user_name_type == "user_name_parts" and user_name_vars["user_name_parts"].get():
                new_account_config["user_name_parts"] = [v.strip() for v in user_name_vars["user_name_parts"].get().split(",") if v.strip()]

            rating_type = rating_type_var.get()
            if rating_type == "rating" and rating_vars.get("rating", tk.StringVar()).get():
                new_account_config["rating"] = rating_vars["rating"].get()
            elif rating_type == "rating_text" and rating_vars.get("positive_value", tk.StringVar()).get() and rating_vars.get("column", tk.StringVar()).get():
                new_account_config["rating_text"] = {
                    "positive_value": rating_vars["positive_value"].get(),
                    "column": rating_vars["column"].get()
                }
            elif rating_type == "rating_inverted" and rating_vars.get("valid_values", tk.StringVar()).get() and rating_vars.get("column", tk.StringVar()).get():
                valid_values = [v.strip() for v in rating_vars["valid_values"].get().split(",") if v.strip()]
                new_account_config["rating_inverted"] = {
                    "valid_values": [int(v) if v.isdigit() else v for v in valid_values],
                    "column": rating_vars["column"].get()
                }

            assignment_column = assignment_column_var.get().strip()
            required_values = [v.strip() for v in required_values_var.get().split(",") if v.strip()]
            if assignment_column and required_values:
                new_account_config["assignment_group"] = {
                    "column": assignment_column,
                    "required_value": required_values if len(required_values) > 1 else required_values[0]
                }

            config_data["accounts"][account_name] = new_account_config

            analyst_groups = [g.strip() for g in analyst_groups_var.get().split(",") if g.strip()]
            analysts_data[account_name] = {
                "groups": {group: {"analysts": {}, "cc_emails": []} for group in analyst_groups}
            } if analyst_groups else {"groups": {}}

            self.save_config(config_data)
            with open(self.analysts_path, 'w', encoding='utf-8') as f:
                json.dump(analysts_data, f, indent=4)

            self.account_dropdown["values"] = list(config_data["accounts"].keys())
            self.account_dropdown.set(account_name)

            self.accounts_data = config_data.get("accounts", {})
            self.account_dropdown.event_generate("<<ComboboxSelected>>")

            messagebox.showinfo("Success", f"Account '{account_name}' added successfully!")
            add_win.destroy()

        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10, fill="x", expand=True)
        button_frame.configure(style="DarkFrame.TFrame")
        ttk.Button(button_frame, text="Add", command=confirm_add).pack(expand=True)

        add_win.deiconify()

    def remove_selected_account(self):
        selected_account = self.account_dropdown.get()
        if not selected_account or selected_account == "Select an account":
            messagebox.showwarning("Warning", "Please select an account to remove.", parent=self)
            return

        if not messagebox.askyesno("Confirm", f"Are you sure you want to remove the account '{selected_account}'?\nThis will also remove its associated analysts from this account.", parent=self):
            return

        config_data = self.load_config()
        with open(self.analysts_path, 'r', encoding='utf-8') as f:
            analysts_data = json.load(f)

        if selected_account in config_data["accounts"]:
            del config_data["accounts"][selected_account]

        if selected_account in analysts_data:
            del analysts_data[selected_account]

        self.save_config(config_data)
        with open(self.analysts_path, 'w', encoding='utf-8') as f:
            json.dump(analysts_data, f, indent=4)

        self.account_dropdown["values"] = list(config_data["accounts"].keys())
        self.account_dropdown.set("Select an account")
        for widget in self.account_fields_frame.winfo_children():
            widget.destroy()

        messagebox.showinfo("Success", f"Account '{selected_account}' removed successfully!", parent=self)

    def save_account_fields(self, account):
        updated_data = {}
        for key, var in self.field_vars.items():
            value = var.get()
            try:
                updated_data[key] = json.loads(value)
            except:
                updated_data[key] = value

        config_data = self.load_config()
        config_data["accounts"][account] = updated_data
        self.save_config(config_data)

        if messagebox.askyesno("Success", f"Settings for '{account}' updated!\n\nDo you want to close the settings window now?"):
            self.destroy()
        else:
            self.lift()

    def init_analysts_tab(self):
        analysts_frame = ttk.Frame(self.notebook)
        self.notebook.add(analysts_frame, text="Analysts Settings")

        try:
            with open(self.analysts_path, 'r', encoding='utf-8') as f:
                self.analysts_data = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load analysts file:\n{e}")
            return

        account_list = list(self.analysts_data.keys())

        self.analysts_dropdown = ttk.Combobox(analysts_frame, values=account_list, state="readonly")
        self.analysts_dropdown.set("Select an account")
        self.analysts_dropdown.pack(pady=10)

        self.analysts_tree = ttk.Treeview(
            analysts_frame,
            columns=("Group", "Type", "Name/Email"),
            show="headings", height=15
        )
        self.analysts_tree.heading("Group", text="Group")
        self.analysts_tree.heading("Type", text="Type")
        self.analysts_tree.heading("Name/Email", text="Name / Email / CC")
        self.analysts_tree.column("Group", width=150, anchor="w")
        self.analysts_tree.column("Type", width=100, anchor="center")
        self.analysts_tree.column("Name/Email", width=400, anchor="w")
        self.analysts_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        bottom_frame = ttk.Frame(analysts_frame)
        bottom_frame.pack(fill=tk.X, pady=5)

        ttk.Button(bottom_frame, text="Save Changes", command=self.save_analysts_changes).pack(side=tk.LEFT, padx=5)

        button_frame = ttk.Frame(bottom_frame)
        button_frame.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="Add Analyst", command=self.open_add_analyst_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remove Selected", command=self.remove_selected_entry).pack(side=tk.LEFT, padx=5)

        self.analysts_dropdown.bind("<<ComboboxSelected>>", self.load_analysts_for_account)
        self.analysts_tree.bind("<Double-1>", self.on_tree_double_click)

    def load_analysts_for_account(self, event=None):
        selected_account = self.analysts_dropdown.get()
        if not selected_account or selected_account not in self.analysts_data:
            return

        self.analysts_tree.delete(*self.analysts_tree.get_children())
        account_data = self.analysts_data[selected_account]

        groups = account_data.get("groups", {})

        for group_name, group_info in groups.items():
            analysts = group_info.get("analysts", {})
            cc_emails = group_info.get("cc_emails", [])

            for name, email in analysts.items():
                display = f"{name} <{email}>" if email else f"{name} (no email)"
                self.analysts_tree.insert("", "end", values=(group_name, "Analyst", display))

            for cc_email in cc_emails:
                self.analysts_tree.insert("", "end", values=(group_name, "CC", cc_email))

    def on_tree_double_click(self, event):
        item_id = self.analysts_tree.identify_row(event.y)
        column = self.analysts_tree.identify_column(event.x)
        if not item_id or column != "#3":
            return

        col_index = int(column.replace("#", "")) - 1
        if col_index < 0:
            return

        item = self.analysts_tree.item(item_id)
        values = item["values"]
        if len(values) < 3:
            return

        group_name, entry_type, value = values

        x, y, width, height = self.analysts_tree.bbox(item_id, column)
        entry = tk.Entry(self.analysts_tree)
        entry.insert(0, value)
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus()

        def save_edit(event=None):
            new_value = entry.get()
            entry.destroy()
            self.analysts_tree.set(item_id, column=column, value=new_value)

            account = self.analysts_dropdown.get()
            if not account:
                return

            group_data = self.analysts_data[account]["groups"][group_name]

            if entry_type == "Analyst":
                if "<" in new_value and ">" in new_value:
                    name = new_value.split("<")[0].strip()
                    email = new_value.split("<")[1].replace(">", "").strip()
                else:
                    name = new_value.strip()
                    email = ""

                for old_name, old_email in list(group_data.get("analysts", {}).items()):
                    display_old = f"{old_name} <{old_email}>" if old_email else f"{old_name} (no email)"
                    if display_old == value:
                        del group_data["analysts"][old_name]
                        group_data["analysts"][name] = email
                        break

            elif entry_type == "CC":
                cc_list = group_data.get("cc_emails", [])
                for i, old_email in enumerate(cc_list):
                    if old_email == value:
                        cc_list[i] = new_value
                        break

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", lambda e: save_edit())

    def save_analysts_changes(self):
        try:
            with open(self.analysts_path, 'w', encoding='utf-8') as f:
                json.dump(self.analysts_data, f, indent=4)
            with open(self.analysts_path, 'r', encoding='utf-8') as f:
                self.analysts_data = json.load(f)
            if messagebox.askyesno("Success", "Analyst data saved successfully!\n\nDo you want to close the settings window?"):
                self.destroy()
            else:
                self.lift()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save analysts file:\n{e}")

    def open_add_analyst_window(self):
        account = self.analysts_dropdown.get()
        if not account or account not in self.analysts_data:
            messagebox.showwarning("Warning", "Please select a valid account first.", parent=self)
            self.account_dropdown.focus_set()
            return

        add_win = tk.Toplevel(self)
        add_win.title("Add Analyst / CC")
        add_win.geometry("400x300")
        add_win.withdraw() 

        add_win.update_idletasks()
        width = 400
        height = 300
        x = (add_win.winfo_screenwidth() // 2) - (width // 2)
        y = (add_win.winfo_screenheight() // 2) - (height // 2)
        add_win.geometry(f"{width}x{height}+{x}+{y}")

        add_win.transient(self)
        add_win.grab_set()

        main_frame = ttk.Frame(add_win)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        groups = list(self.analysts_data[account]["groups"].keys())

        ttk.Label(main_frame, text="Group:").pack(pady=5)
        group_var = tk.StringVar()
        group_dropdown = ttk.Combobox(main_frame, textvariable=group_var, values=groups, state="readonly")
        group_dropdown.pack(pady=5)

        ttk.Label(main_frame, text="Type:").pack(pady=5)
        type_var = tk.StringVar(value="Analyst")
        type_dropdown = ttk.Combobox(main_frame, textvariable=type_var, values=["Analyst", "CC"], state="readonly")
        type_dropdown.pack(pady=5)

        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill=tk.BOTH, expand=True)

        name_var = tk.StringVar()
        email_var = tk.StringVar()
        name_label = ttk.Label(fields_frame, text="Name:")
        name_entry = ttk.Entry(fields_frame, textvariable=name_var)
        email_label = ttk.Label(fields_frame, text="Email:")
        email_entry = ttk.Entry(fields_frame, textvariable=email_var)

        def update_fields(*args):
            name_label.pack_forget()
            name_entry.pack_forget()
            email_label.pack_forget()
            email_entry.pack_forget()

            if type_var.get() == "Analyst":
                name_label.pack(pady=5)
                name_entry.pack(pady=5)
                email_label.pack(pady=5)
                email_entry.pack(pady=5)
            else:
                email_label.pack(pady=5)
                email_entry.pack(pady=5)

        type_var.trace_add("write", update_fields)
        update_fields()

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        def confirm_add():
            group = group_var.get()
            entry_type = type_var.get()
            email = email_var.get().strip()
            name = name_var.get().strip()

            if not group or not email:
                messagebox.showerror("Error", "Group and email are required.")
                return

            group_data = self.analysts_data[account]["groups"][group]

            if entry_type == "Analyst":
                if not name:
                    messagebox.showerror("Error", "Name is required for analysts.")
                    return
                group_data.setdefault("analysts", {})[name] = email
                display = f"{name} <{email}>" if email else f"{name} (no email)"
                self.analysts_tree.insert("", "end", values=(group, "Analyst", display))
            else:
                group_data.setdefault("cc_emails", []).append(email)
                self.analysts_tree.insert("", "end", values=(group, "CC", email))

            add_win.destroy()

        ttk.Button(button_frame, text="Add", command=confirm_add).pack()

        add_win.deiconify()

    def remove_selected_entry(self):
        selected_item = self.analysts_tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "No item selected.", parent=self)
            return

        confirm = messagebox.askyesno("Confirm", "Are you sure you want to remove the selected entry?", parent=self)
        if not confirm:
            return

        item = self.analysts_tree.item(selected_item)
        group, entry_type, value = item["values"]
        account = self.analysts_dropdown.get()
        if not account:
            return

        group_data = self.analysts_data[account]["groups"].get(group, {})

        if entry_type == "Analyst":
            for name, email in list(group_data.get("analysts", {}).items()):
                display = f"{name} <{email}>" if email else f"{name} (no email)"
                if display == value:
                    del group_data["analysts"][name]
                    break
        elif entry_type == "CC":
            cc_list = group_data.get("cc_emails", [])
            if value in cc_list:
                cc_list.remove(value)

        self.analysts_tree.delete(selected_item)

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        