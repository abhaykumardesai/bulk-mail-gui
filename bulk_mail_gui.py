"""
Premium-looking Bulk Mail Sender - Responsive Tkinter UI
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, time, os
import pandas as pd
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ---------------- HARDCODED SENDER ----------------
SENDER_EMAIL = "avimuktaip@gmail.com"
SENDER_APP_PASSWORD = "hnzrtcauftoqwfun"
# --------------------------------------------------

class SafeDict(dict):
    def __missing__(self, key):
        return ""

def send_messages(gui_state, progress_callback, log_callback):
    try:
        df = pd.read_excel(gui_state["excel_path"],
                           sheet_name=gui_state.get("sheet_name", "Sheet1"),
                           engine="openpyxl").fillna("")
    except Exception as e:
        log_callback(f"ERROR loading Excel: {repr(e)}")
        return

    total = len(df)
    email_col = gui_state["col_email"]
    name_col = gui_state.get("col_name", "")

    log_callback(f"Starting sending. Total rows: {total}")
    log_callback(f"Mode: {'DRY RUN' if gui_state['dry_run'] else 'REAL SEND'}")

    context = ssl.create_default_context()
    server = None
    if not gui_state["dry_run"]:
        try:
            server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
            server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
            log_callback("SMTP login OK")
        except Exception as e:
            log_callback("SMTP login error: " + repr(e))
            return
    else:
        log_callback("Dry run: no SMTP login")

    sent = 0
    start = time.perf_counter()

    for idx, row in df.iterrows():
        recipient = str(row.get(email_col, "")).strip()
        if not recipient:
            log_callback(f"Skipping row {idx}: No email found")
            progress_callback(idx + 1, total)
            continue

        name_val = "" if pd.isna(row.get(name_col, "")) else str(row.get(name_col, ""))
        row_map = {"Name": name_val}

        try:
            body = gui_state["body_template"].format_map(SafeDict(row_map))
        except:
            body = gui_state["body_template"]

        subject = gui_state["subject_template"]

        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = recipient
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        for att in gui_state.get("attachments", []):
            if os.path.isfile(att):
                try:
                    with open(att, "rb") as f:
                        part = MIMEApplication(f.read(), Name=os.path.basename(att))
                        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(att)}"'
                        msg.attach(part)
                except Exception as e:
                    log_callback(f"Attachment error: {repr(e)}")

        if gui_state["dry_run"]:
            log_callback(f"[DRY] Would send to {recipient}")
        else:
            try:
                server.sendmail(SENDER_EMAIL, recipient, msg.as_string())
                sent += 1
                log_callback(f"Sent: {recipient}")
            except Exception as e:
                log_callback("Send error: " + repr(e))

        progress_callback(idx + 1, total)
        time.sleep(gui_state.get("delay", 0.1))

    if server:
        try: server.quit()
        except: pass

    log_callback(f"Finished sending: {sent}/{total}")

###############################################################
# ------------------ GUI CLASS START -------------------------
###############################################################

class BulkMailGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bulk Mail Sender — Abhay Premium")
        self.minsize(780, 520)

        try:
            self.attributes("-zoomed", True)
        except:
            self.state("zoomed")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 15, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=6)

        self._font_size = 10
        self.bind("<Configure>", self._resize)

        self.state_data = {
            "excel_path": "",
            "sheet_name": "Sheet1",
            "col_email": "",
            "col_name": "",
            "subject_template": "Hello from Abhay",
            "body_template": """Hi {Name},

Please find the attached document.

Regards,
Abhay
""",
            "attachments": [],
            "delay": 0.1,
            "dry_run": True
        }

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main = ttk.Frame(self, padding=12)
        main.grid(sticky="nsew")
        main.grid_rowconfigure(2, weight=1)
        main.grid_columnconfigure(0, weight=1)

        header = ttk.Label(main, text="Bulk Mail Sender", style="Title.TLabel")
        header.grid(row=0, column=0, sticky="w", pady=(0, 10))

        ##################################################
        # ---------------- TOP CONTROLS ------------------
        ##################################################

        controls = ttk.Frame(main)
        controls.grid(row=1, column=0, sticky="ew")
        controls.grid_columnconfigure(3, weight=1)

        ttk.Button(controls, text="Load Excel", style="Accent.TButton",
                   command=self.load_excel).grid(row=0, column=0, padx=6)

        self.lbl_excel = ttk.Label(controls, text="No file selected")
        self.lbl_excel.grid(row=0, column=1, sticky="w")

        ttk.Label(controls, text="Email column:").grid(row=0, column=2)
        self.combo_email = ttk.Combobox(controls, values=[], width=20)
        self.combo_email.grid(row=0, column=3)

        ttk.Label(controls, text="Name column:").grid(row=0, column=4)
        self.combo_name = ttk.Combobox(controls, values=[], width=20)
        self.combo_name.grid(row=0, column=5)

        ##################################################
        # ---------------- MAIN AREA ---------------------
        ##################################################

        mid = ttk.Frame(main)
        mid.grid(row=2, column=0, sticky="nsew")
        mid.grid_columnconfigure(0, weight=3)
        mid.grid_columnconfigure(1, weight=2)
        mid.grid_rowconfigure(0, weight=1)

        # Compose
        compose = ttk.LabelFrame(mid, text="Compose Message")
        compose.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        compose.grid_rowconfigure(3, weight=1)
        compose.grid_columnconfigure(0, weight=1)

        ttk.Label(compose, text="Subject:").grid(row=0, column=0, sticky="w", padx=6)
        self.entry_subject = ttk.Entry(compose)
        self.entry_subject.insert(0, self.state_data["subject_template"])
        self.entry_subject.grid(row=1, column=0, sticky="ew", padx=6)

        ttk.Label(compose, text="Body (uses {Name})").grid(row=2, column=0, sticky="w", padx=6)
        self.txt_body = tk.Text(compose, wrap="word")
        self.txt_body.insert("1.0", self.state_data["body_template"])
        self.txt_body.grid(row=3, column=0, sticky="nsew", padx=6, pady=6)

        ##################################################
        # -------- RIGHT SIDE → PREVIEW + ATTACHMENTS ----
        ##################################################

        right = ttk.Frame(mid)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        # Preview
        pv_frame = ttk.LabelFrame(right, text="Preview (first 10 rows)")
        pv_frame.grid(row=0, column=0, sticky="nsew")
        pv_frame.grid_rowconfigure(0, weight=1)

        self.tree_preview = ttk.Treeview(pv_frame, columns=("email", "name"),
                                         show="headings")
        self.tree_preview.heading("email", text="Email")
        self.tree_preview.heading("name", text="Name")
        self.tree_preview.grid(row=0, column=0, sticky="nsew")
        
        pv_scroll = ttk.Scrollbar(pv_frame, orient="vertical",
                                  command=self.tree_preview.yview)
        pv_scroll.grid(row=0, column=1, sticky="ns")
        self.tree_preview.configure(yscrollcommand=pv_scroll.set)

        ##################################################
        # ---------------- ATTACHMENTS -------------------
        ##################################################

        att_frame = ttk.LabelFrame(right, text="Attachments")
        att_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        att_frame.grid_columnconfigure(0, weight=1)
        att_frame.grid_rowconfigure(1, weight=1)

        ttk.Button(att_frame, text="Add", style="Accent.TButton",
                   command=self.add_attachments).grid(row=0, column=0, sticky="w", padx=6, pady=4)

        ttk.Button(att_frame, text="Remove Selected",
                   command=self.remove_selected_attachments).grid(row=0, column=1, sticky="w", padx=6)

        self.lbl_attach = ttk.Label(att_frame, text="0 files")
        self.lbl_attach.grid(row=0, column=2, sticky="e", padx=6)

        self.attach_tree = ttk.Treeview(att_frame, columns=("file",),
                                         show="headings")
        self.attach_tree.heading("file", text="Filename")
        self.attach_tree.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=6, pady=(0, 6))

        att_scroll = ttk.Scrollbar(att_frame, orient="vertical",
                                    command=self.attach_tree.yview)
        att_scroll.grid(row=1, column=3, sticky="ns")
        self.attach_tree.configure(yscrollcommand=att_scroll.set)

        ##################################################
        # ---------------- BOTTOM AREA --------------------
        ##################################################

        bottom = ttk.Frame(main)
        bottom.grid(row=3, column=0, sticky="ew", pady=10)
        bottom.grid_columnconfigure(1, weight=1)

        self.dry_run_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="Dry Run (Preview Only)",
                        variable=self.dry_run_var).grid(row=0, column=0, padx=6)

        ttk.Label(bottom, text="Delay (s):").grid(row=0, column=1, sticky="e")
        self.entry_delay = ttk.Entry(bottom, width=6)
        self.entry_delay.insert(0, "0.1")
        self.entry_delay.grid(row=0, column=2, sticky="w", padx=(6, 20))

        # Progress Bar
        self.progress = ttk.Progressbar(bottom, length=300)
        self.progress.grid(row=1, column=0, columnspan=3, sticky="ew", padx=6, pady=8)

        # ACTION BUTTONS
        btn_frame = ttk.Frame(bottom)
        btn_frame.grid(row=1, column=3, sticky="e")

        self.preview_button = ttk.Button(btn_frame, text="Preview (Dry Run)",
                                         style="Accent.TButton",
                                         command=lambda: self._start_send(True))
        self.preview_button.grid(row=0, column=0, padx=6)

        self.send_button = ttk.Button(btn_frame, text="Send (Actual)",
                                      style="Accent.TButton",
                                      command=lambda: self._start_send(False))
        self.send_button.grid(row=0, column=1, padx=6)

        ttk.Button(btn_frame, text="Quit",
                   command=self.quit).grid(row=0, column=2, padx=6)

        ##################################################
        # ---------------- LOG BOX ------------------------
        ##################################################

        log_frame = ttk.LabelFrame(main, text="Log")
        log_frame.grid(row=4, column=0, sticky="nsew")
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self.txt_log = tk.Text(log_frame, height=8, wrap="none")
        self.txt_log.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical",
                                   command=self.txt_log.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.txt_log.configure(yscrollcommand=log_scroll.set)

    ###############################################################
    # ---------------------- UI FUNCTIONS -------------------------
    ###############################################################

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xls")])
        if not path:
            return

        self.state_data["excel_path"] = path
        self.lbl_excel.config(text=os.path.basename(path))

        try:
            df = pd.read_excel(path, engine="openpyxl")
            cols = [str(c).strip() for c in df.columns.tolist()]
            self.combo_email.config(values=cols)
            self.combo_name.config(values=cols)

            # Prefill comboboxes if not set
            if cols:
                # if user hasn't chosen an email column yet, pick the first likely one
                if not self.combo_email.get().strip():
                    # try to auto-detect common email header names
                    lower = [c.lower() for c in cols]
                    email_idx = None
                    for key in ("email", "e-mail", "email address", "mail"):
                        if key in lower:
                            email_idx = lower.index(key)
                            break
                    if email_idx is None:
                        email_idx = 0
                    self.combo_email.set(cols[email_idx])

                # if name not set, pick a sensible second column if present
                if not self.combo_name.get().strip() and len(cols) > 1:
                    # prefer common name headers
                    lower = [c.lower() for c in cols]
                    name_idx = None
                    for key in ("name", "full name", "firstname", "first name"):
                        if key in lower:
                            name_idx = lower.index(key)
                            break
                    if name_idx is None:
                        # pick the next column (not the chosen email column)
                        name_idx = 1 if email_idx == 0 else 0
                    self.combo_name.set(cols[name_idx])

            # Resolve which columns we'll actually use for preview
            email_col = self.combo_email.get().strip()
            if email_col not in df.columns:
                email_col = cols[0] if cols else None

            name_col = self.combo_name.get().strip()
            if name_col not in df.columns:
                # try to find some other sensible column for name, otherwise leave empty
                name_col = None
                if len(cols) > 1:
                    # pick a column that's not the email column
                    for c in cols:
                        if c != email_col:
                            name_col = c
                            break

            # Clear and insert preview rows (first 10)
            for r in self.tree_preview.get_children():
                self.tree_preview.delete(r)

            for _, row in df.head(10).iterrows():
                email = row.get(email_col, "") if email_col else ""
                name = row.get(name_col, "") if name_col else ""
                self.tree_preview.insert("", "end", values=(str(email), str(name)))

            self.log(f"Loaded rows: {len(df)} — using Email column: '{email_col}' Name column: '{name_col or 'N/A'}'")

        except Exception as e:
            self.log("ERROR: " + repr(e))


    def add_attachments(self):
        paths = filedialog.askopenfilenames(filetypes=[("Files", "*.*")])
        if not paths:
            return

        for p in paths:
            if p not in self.state_data["attachments"]:
                self.state_data["attachments"].append(p)
                self.attach_tree.insert("", "end", values=(os.path.basename(p),))

        self.lbl_attach.config(text=f"{len(self.state_data['attachments'])} files")

    def remove_selected_attachments(self):
        sel = self.attach_tree.selection()
        if not sel:
            return

        items = self.attach_tree.get_children()
        indexes = [items.index(i) for i in sel]

        for item in sel:
            self.attach_tree.delete(item)

        for i in sorted(indexes, reverse=True):
            del self.state_data["attachments"][i]

        self.lbl_attach.config(text=f"{len(self.state_data['attachments'])} files")

    def log(self, text):
        time_str = time.strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{time_str}] {text}\n")
        self.txt_log.see("end")

    ###############################################################
    # ------------------- SEND MESSAGE WRAPPER --------------------
    ###############################################################

    def _start_send(self, dry_run):
        if not dry_run:
            if not messagebox.askyesno("Confirm", "Send actual emails?"):
                return

        if not self.state_data["excel_path"]:
            messagebox.showerror("Error", "Load Excel first.")
            return

        self.state_data["dry_run"] = dry_run
        self.state_data["col_email"] = self.combo_email.get().strip()
        self.state_data["col_name"] = self.combo_name.get().strip()
        self.state_data["subject_template"] = self.entry_subject.get().strip()
        self.state_data["body_template"] = self.txt_body.get("1.0", "end").strip()

        try:
            self.state_data["delay"] = float(self.entry_delay.get().strip())
        except:
            self.state_data["delay"] = 0.1

        self.preview_button.config(state="disabled")
        self.send_button.config(state="disabled")

        def worker():
            send_messages(self.state_data, self.progress_callback, self.log)
            self.after(100, self._finish_send)

        threading.Thread(target=worker, daemon=True).start()

        self.log("Started...")

    def _finish_send(self):
        self.preview_button.config(state="normal")
        self.send_button.config(state="normal")
        self.log("Done.")

    def progress_callback(self, idx, total):
        try:
            self.progress["value"] = (idx / total) * 100
        except:
            pass

    ###############################################################
    # ------------------- LAYOUT RESIZE FIX -----------------------
    ###############################################################

    def _resize(self, event):
        try:
            w = self.tree_preview.winfo_width()
            if w > 100:
                self.tree_preview.column("email", width=int(w * 0.6))
                self.tree_preview.column("name", width=int(w * 0.38))
        except:
            pass

        try:
            w2 = self.attach_tree.winfo_width()
            if w2 > 100:
                self.attach_tree.column("file", width=int(w2 * 0.95))
        except:
            pass

###############################################################
# ---------------------- RUN APP ------------------------------
###############################################################

if __name__ == "__main__":
    BulkMailGUI().mainloop()

