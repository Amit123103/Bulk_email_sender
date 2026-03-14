"""Bulk Email Pro — Desktop Application (Tkinter)"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import threading, time, queue, json, os, webbrowser
from datetime import datetime
from pathlib import Path
from account_manager import AccountManager # type: ignore
from excel_processor import ExcelProcessor # type: ignore
from smtp_engine import SMTPEngine # type: ignore
from ai_generator import AIEmailGenerator # type: ignore
from spam_checker import SpamChecker # type: ignore
from scheduler import CampaignScheduler # type: ignore
from ab_tester import ABTester # type: ignore
from contact_manager import ContactManager # type: ignore
import typing

DATA_DIR = Path(os.path.expanduser("~")) / ".bulk_email_pro"
DATA_DIR.mkdir(parents=True, exist_ok=True)

class AddAccountDialog(tk.Toplevel):
    PRESETS = {
        "Gmail": {"smtp_host":"smtp.gmail.com","smtp_port":587,"smtp_security":"TLS","daily_limit":500,
            "help":"Use Gmail App Password (not regular password).\nEnable: Google Account > Security > 2-Step Verification > App Passwords"},
        "Outlook / Hotmail": {"smtp_host":"smtp-mail.outlook.com","smtp_port":587,"smtp_security":"TLS","daily_limit":300,
            "help":"Use regular Outlook password or App Password if MFA enabled."},
        "Yahoo Mail": {"smtp_host":"smtp.mail.yahoo.com","smtp_port":465,"smtp_security":"SSL","daily_limit":500,
            "help":"Generate App Password: Yahoo Account Security > Generate App Password"},
        "Zoho Mail": {"smtp_host":"smtp.zoho.com","smtp_port":587,"smtp_security":"TLS","daily_limit":500,
            "help":"Use Zoho email and password. Enable SMTP in Zoho Mail settings."},
        "SendGrid": {"smtp_host":"smtp.sendgrid.net","smtp_port":587,"smtp_security":"TLS","daily_limit":100,
            "help":"Username: 'apikey' (literally). Password: your SendGrid API Key."},
        "Custom SMTP": {"smtp_host":"","smtp_port":587,"smtp_security":"TLS","daily_limit":500,
            "help":"Enter your hosting provider's SMTP details."},
    }
    def __init__(self, parent, account_manager, account_id=None):
        super().__init__(parent)
        self.am = account_manager
        self.account_id = account_id
        self.result: typing.Optional[bool] = None
        self.provider_var: tk.StringVar = None # type: ignore
        self.sec_var: tk.StringVar = None # type: ignore
        self.entries: dict[str, tk.Entry] = {}
        self.help_text: tk.Text = None # type: ignore
        self.status_label: tk.Label = None # type: ignore
        self.title("Edit Email Account" if account_id else "Add Email Account")
        self.geometry("520x620")
        self.resizable(False, False)
        self.configure(bg="#1a1a2e")
        self.transient(parent)
        self.grab_set()
        self.build_ui()
        if account_id:
            self.load_existing()
    def build_ui(self):
        C = {"bg":"#1a1a2e","fg":"#e0e0e0"}
        tk.Label(self, text="Email Provider:", **C, font=("Segoe UI",10,"bold")).pack(anchor="w",padx=15,pady=(15,2)) # type: ignore
        self.provider_var = tk.StringVar(value="Gmail")
        cb = ttk.Combobox(self, textvariable=self.provider_var, values=list(self.PRESETS.keys()), state="readonly", width=30)
        cb.pack(anchor="w",padx=15)
        cb.bind("<<ComboboxSelected>>", self.on_preset_change)
        fields = [("Nickname:","nick"),("Email Address:","email"),("Password:","pwd"),
                  ("SMTP Host:","host"),("SMTP Port:","port"),("Daily Limit:","limit")]
        self.entries = {}
        for label, key in fields:
            tk.Label(self, text=label, **C, font=("Segoe UI",9)).pack(anchor="w",padx=15,pady=(8,1)) # type: ignore
            show = "*" if key == "pwd" else ""
            e = tk.Entry(self, bg="#16213e", fg="#e0e0e0", insertbackground="#4ecca3", font=("Segoe UI",10), show=show)
            e.pack(fill="x",padx=15)
            self.entries[key] = e
        tk.Label(self, text="Security:", **C, font=("Segoe UI",9)).pack(anchor="w",padx=15,pady=(8,1)) # type: ignore
        self.sec_var = tk.StringVar(value="TLS")
        sf = tk.Frame(self, bg="#1a1a2e")
        sf.pack(anchor="w",padx=15)
        for v in ("TLS","SSL","None"):
            tk.Radiobutton(sf, text=v, variable=self.sec_var, value=v, bg="#1a1a2e", fg="#e0e0e0",
                          selectcolor="#16213e", activebackground="#1a1a2e", activeforeground="#4ecca3").pack(side="left",padx=5)
        self.help_text = tk.Text(self, height=3, bg="#0d0d1a", fg="#a0a0b0", font=("Segoe UI",8), wrap="word", state="disabled")
        self.help_text.pack(fill="x",padx=15,pady=(8,5))
        self.status_label = tk.Label(self, text="", bg="#1a1a2e", fg="#4ecca3", font=("Segoe UI",9))
        self.status_label.pack(padx=15)
        bf = tk.Frame(self, bg="#1a1a2e")
        bf.pack(fill="x",padx=15,pady=10)
        tk.Button(bf, text="🔌 Test Connection", bg="#16213e", fg="#4ecca3", font=("Segoe UI",9,"bold"),
                 command=self.test_connection, relief="flat", padx=10, pady=5).pack(side="left",padx=3)
        tk.Button(bf, text="💾 Save Account", bg="#4ecca3", fg="#0d0d1a", font=("Segoe UI",9,"bold"),
                 command=self.save_account, relief="flat", padx=10, pady=5).pack(side="left",padx=3)
        tk.Button(bf, text="Cancel", bg="#2a2a4a", fg="#e0e0e0", font=("Segoe UI",9),
                 command=self.destroy, relief="flat", padx=10, pady=5).pack(side="left",padx=3)
        self.on_preset_change(None)
    def on_preset_change(self, event):
        p = self.PRESETS.get(self.provider_var.get(), {})
        self.entries["host"].delete(0,"end"); self.entries["host"].insert(0, str(p.get("smtp_host","")))
        self.entries["port"].delete(0,"end"); self.entries["port"].insert(0, str(p.get("smtp_port",587)))
        self.entries["limit"].delete(0,"end"); self.entries["limit"].insert(0, str(p.get("daily_limit",500)))
        self.sec_var.set(str(p.get("smtp_security","TLS")))
        self.help_text.config(state="normal"); self.help_text.delete("1.0","end")
        self.help_text.insert("1.0", str(p.get("help",""))); self.help_text.config(state="disabled")
    def load_existing(self):
        acc = self.am.get_account(self.account_id)
        if not acc: return
        self.entries["nick"].insert(0, acc["nickname"])
        self.entries["email"].insert(0, acc["email"])
        self.entries["host"].insert(0, acc["smtp_host"])
        self.entries["port"].delete(0,"end"); self.entries["port"].insert(0, str(acc["smtp_port"]))
        self.entries["limit"].delete(0,"end"); self.entries["limit"].insert(0, str(acc["daily_limit"]))
        self.sec_var.set(acc["smtp_security"])
    def test_connection(self):
        self.status_label.config(text="Testing...", fg="#f39c12")
        self.update()
        def _test():
            import smtplib, ssl
            try:
                h,p,s = self.entries["host"].get(), int(self.entries["port"].get()), self.sec_var.get()
                e,pw = self.entries["email"].get(), self.entries["pwd"].get()
                if s=="SSL":
                    srv=smtplib.SMTP_SSL(h,p,context=ssl.create_default_context(),timeout=20)
                else:
                    srv=smtplib.SMTP(h,p,timeout=20); srv.ehlo()
                    if s=="TLS": srv.starttls(context=ssl.create_default_context()); srv.ehlo()
                srv.login(e,pw); srv.quit()
                self.status_label.config(text="✅ Connected successfully!", fg="#2ecc71")
            except Exception as ex:
                self.status_label.config(text=f"❌ Failed: {ex}", fg="#e74c3c")
        threading.Thread(target=_test, daemon=True).start()
    def save_account(self):
        nick = self.entries["nick"].get().strip()
        email = self.entries["email"].get().strip()
        pwd = self.entries["pwd"].get()
        host = self.entries["host"].get().strip()
        try: port = int(self.entries["port"].get())
        except: messagebox.showerror("Error","Invalid port number",parent=self); return
        try: limit = int(self.entries["limit"].get())
        except: limit = 500
        sec = self.sec_var.get()
        if not all([nick,email,host]): messagebox.showerror("Error","Fill all required fields",parent=self); return
        try:
            if self.account_id:
                updates = {"nickname":nick,"email":email,"smtp_host":host,"smtp_port":port,"smtp_security":sec,"daily_limit":limit}
                if pwd: updates["password"] = pwd
                self.am.update_account(self.account_id, updates)
            else:
                if not pwd: messagebox.showerror("Error","Password is required",parent=self); return
                self.am.add_account(nick, email, host, port, sec, pwd, limit)
            self.result = True
            self.destroy()
        except ValueError as ve:
            messagebox.showerror("Error", str(ve), parent=self)

class BulkEmailProApp:
    COLORS = {'bg':'#0d0d1a','bg2':'#1a1a2e','bg3':'#16213e','accent':'#4ecca3','accent2':'#e94560',
              'text':'#e0e0e0','text2':'#a0a0b0','success':'#2ecc71','warning':'#f39c12','danger':'#e74c3c','border':'#2a2a4a'}
    def __init__(self, root):
        self.root = root
        self.am = AccountManager(str(DATA_DIR))
        self.ep = ExcelProcessor()
        self.smtp = SMTPEngine()
        self.ai = AIEmailGenerator()
        self.spam = SpamChecker()
        self.scheduler = CampaignScheduler(str(DATA_DIR))
        self.ab_tester = ABTester(str(DATA_DIR))
        self.contacts = ContactManager(str(DATA_DIR))
        self.log_queue = queue.Queue()
        self.is_sending = False
        self.is_paused = False
        self.send_thread = None
        self.campaign_start_time = None
        self.attachment_paths = []
        self.scheduled_time = None
        self.schedule_timer = None
        
        # UI Attributes for typing constraints
        self.notebook: ttk.Notebook = None # type: ignore
        self.status_bar: tk.Label = None # type: ignore
        self.subject_entry: tk.Entry = None # type: ignore
        self.body_editor: scrolledtext.ScrolledText = None # type: ignore
        self.html_var: tk.BooleanVar = None # type: ignore
        self.rotate_var: tk.BooleanVar = None # type: ignore
        self.delay_var: tk.DoubleVar = None # type: ignore
        self.acc_stats_label: tk.Label = None # type: ignore
        self.acc_frame: tk.Frame = None # type: ignore
        self.file_label: tk.Label = None # type: ignore
        self.col_var: tk.StringVar = None # type: ignore
        self.col_combo: ttk.Combobox = None # type: ignore
        self.stats_frame: tk.Frame = None # type: ignore
        self.email_stats_label: tk.Label = None # type: ignore
        self.tree_scroll_y: ttk.Scrollbar = None # type: ignore
        self.tree_scroll_x: ttk.Scrollbar = None # type: ignore
        self.preview_tree: ttk.Treeview = None # type: ignore
        self.vars_label: tk.Label = None # type: ignore
        self.send_from_var: tk.StringVar = None # type: ignore
        self.send_from_combo: ttk.Combobox = None # type: ignore
        self.subj_count: tk.Label = None # type: ignore
        self.attach_label: tk.Label = None # type: ignore
        self.send_parent: tk.Frame = None # type: ignore
        self.send_summary: tk.Label = None # type: ignore
        self.delay_scale: tk.Scale = None # type: ignore
        self.delay_label: tk.Label = None # type: ignore
        self.start_btn: tk.Button = None # type: ignore
        self.pause_btn: tk.Button = None # type: ignore
        self.stop_btn: tk.Button = None # type: ignore
        self.progress_bar: ttk.Progressbar = None # type: ignore
        self.progress_label: tk.Label = None # type: ignore
        self.stats_label: tk.Label = None # type: ignore
        self.send_log: scrolledtext.ScrolledText = None # type: ignore
        self.send_results: list = []
        self.campaign_start_time: typing.Optional[datetime] = None
        self.report_summary: tk.Label = None # type: ignore
        self.report_tree: ttk.Treeview = None # type: ignore
        
        self.setup_window()
        self.build_ui()
        self.start_queue_processor()
        self.load_draft()
    def setup_window(self):
        self.root.title("📧 Bulk Email Pro v3.5 — Advanced Bulk Email Sender")
        self.root.geometry("1050x780")
        self.root.minsize(850,650)
        self.root.configure(bg=self.COLORS['bg'])
        try: self.root.state('zoomed')
        except: pass
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    def on_close(self):
        if self.is_sending:
            if not messagebox.askyesno("Confirm","A campaign is running. Stop and exit?"): return
            self.is_sending = False
        self.save_draft()
        self.root.destroy()
    def build_ui(self):
        self.build_menu()
        style = ttk.Style()
        style.theme_use('clam')
        C = self.COLORS
        style.configure('.', background=C['bg'], foreground=C['text'])
        style.configure('TNotebook', background=C['bg'], borderwidth=0)
        style.configure('TNotebook.Tab', background=C['bg2'], foreground=C['text'], padding=[12,4],
                        font=('Segoe UI',10,'bold'))
        style.map('TNotebook.Tab', background=[('selected',C['bg3'])], foreground=[('selected',C['accent'])])
        style.configure('TFrame', background=C['bg'])
        style.configure('TLabel', background=C['bg'], foreground=C['text'])
        style.configure('TButton', background=C['bg3'], foreground=C['text'])
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        tabs = [("🏠 Dashboard", self.build_dashboard_tab), ("📋 Accounts", self.build_accounts_tab),
                ("📂 Recipients", self.build_recipients_tab), ("✉️ Compose", self.build_compose_tab),
                ("📤 Send", self.build_send_tab), ("📊 Reports", self.build_reports_tab)]
        for title, builder in tabs:
            frame = tk.Frame(self.notebook, bg=C['bg'])
            self.notebook.add(frame, text=title)
            builder(frame)
        sb = tk.Frame(self.root, bg=C['bg2'], height=25)
        sb.pack(fill="x", side="bottom")
        self.status_bar = tk.Label(sb, text=f"  Accounts: {len(self.am.accounts)} | Ready", bg=C['bg2'],
                                   fg=C['text2'], font=("Segoe UI",8), anchor="w")
        self.status_bar.pack(fill="x", padx=5)

    # ─── DRAFT PERSISTENCE ────────────────────────────────────────
    def save_draft(self):
        """Auto-save compose state between sessions."""
        try:
            draft = {
                "subject": self.subject_entry.get() if hasattr(self, 'subject_entry') else "",
                "body": self.body_editor.get("1.0","end-1c") if hasattr(self, 'body_editor') else "",
                "is_html": self.html_var.get() if hasattr(self, 'html_var') else True,
                "rotate": self.rotate_var.get() if hasattr(self, 'rotate_var') else False,
                "delay": self.delay_var.get() if hasattr(self, 'delay_var') else 1.5,
            }
            with open(DATA_DIR / "draft.json", "w") as f:
                json.dump(draft, f)
        except Exception:
            pass
    def load_draft(self):
        """Restore last compose state on startup."""
        try:
            dp = DATA_DIR / "draft.json"
            if not dp.exists():
                return
            with open(dp, "r") as f:
                draft = json.load(f)
            if hasattr(self, 'subject_entry') and draft.get("subject"):
                self.subject_entry.insert(0, draft["subject"])
            if hasattr(self, 'body_editor') and draft.get("body"):
                self.body_editor.insert("1.0", draft["body"])
            if hasattr(self, 'html_var'):
                self.html_var.set(draft.get("is_html", True))
            if hasattr(self, 'rotate_var'):
                self.rotate_var.set(draft.get("rotate", False))
            if hasattr(self, 'delay_var'):
                self.delay_var.set(draft.get("delay", 1.5))
        except Exception:
            pass

    # ─── TAB 0: DASHBOARD ─────────────────────────────────────────
    def build_dashboard_tab(self, parent):
        C = self.COLORS
        header = tk.Frame(parent, bg=C['bg'])
        header.pack(fill="x", padx=15, pady=(15,5))
        tk.Label(header, text="📧 Bulk Email Pro v3.5", bg=C['bg'], fg=C['accent'],
                font=("Segoe UI",18,"bold")).pack(side="left")
        tk.Label(header, text="Advanced Bulk Email Sender", bg=C['bg'], fg=C['text2'],
                font=("Segoe UI",10)).pack(side="left",padx=15)
        # Stats cards row
        stats = self.am.get_stats()
        cards_frame = tk.Frame(parent, bg=C['bg'])
        cards_frame.pack(fill="x", padx=15, pady=10)
        card_data = [
            ("📋 Accounts", str(stats['total_accounts']), C['accent']),
            ("✅ Active", str(stats['active_accounts']), C['success']),
            ("📤 Sent Today", str(stats['total_sent_today']), C['warning']),
            ("📊 Remaining", str(stats['total_remaining_today']), C['accent']),
            ("❤️ Avg Health", f"{stats.get('avg_health_score',100)}%", C['success']),
            ("📈 Lifetime Sent", str(stats.get('lifetime_sent',0)), C['text']),
        ]
        for i, (label, value, color) in enumerate(card_data):
            card = tk.Frame(cards_frame, bg=C['bg2'], highlightbackground=C['border'],
                          highlightthickness=1, padx=20, pady=12)
            card.grid(row=0, column=i, padx=5, sticky="nsew")
            cards_frame.columnconfigure(i, weight=1)
            tk.Label(card, text=value, bg=C['bg2'], fg=color,
                    font=("Segoe UI",22,"bold")).pack()
            tk.Label(card, text=label, bg=C['bg2'], fg=C['text2'],
                    font=("Segoe UI",9)).pack()
        # Quick actions
        actions = tk.Frame(parent, bg=C['bg'])
        actions.pack(fill="x", padx=15, pady=10)
        tk.Label(actions, text="Quick Actions", bg=C['bg'], fg=C['text'],
                font=("Segoe UI",12,"bold")).pack(anchor="w",pady=(0,8))
        btn_row = tk.Frame(actions, bg=C['bg'])
        btn_row.pack(fill="x")
        btns = [
            ("+ Add Account", C['accent'], lambda: (self.notebook.select(1), self.open_add_account())),
            ("📂 Load Excel", C['bg3'], lambda: (self.notebook.select(2), self.browse_excel())),
            ("🚀 New Campaign", C['accent2'], lambda: self.notebook.select(4)),
            ("📊 View Reports", C['bg3'], lambda: self.notebook.select(5)),
            ("🔄 Test All", C['bg3'], self.test_all_accounts),
            ("📥 Export Accounts", C['bg3'], self.export_accounts_gui),
        ]
        for text, bg, cmd in btns:
            tk.Button(btn_row, text=text, bg=bg, fg=C['bg'] if bg==C['accent'] or bg==C['accent2'] else C['text'],
                     font=("Segoe UI",9,"bold"), relief="flat", padx=12, pady=6,
                     command=cmd).pack(side="left",padx=4)
        # Campaign history
        hist_frame = tk.Frame(parent, bg=C['bg'])
        hist_frame.pack(fill="both", expand=True, padx=15, pady=5)
        tk.Label(hist_frame, text="Recent Campaigns", bg=C['bg'], fg=C['text'],
                font=("Segoe UI",11,"bold")).pack(anchor="w",pady=(0,5))
        hist_path = DATA_DIR / "campaigns.json"
        if hist_path.exists():
            try:
                with open(hist_path,"r") as f:
                    history = json.load(f)
                for camp in reversed(history[-8:]):
                    row = tk.Frame(hist_frame, bg=C['bg2'])
                    row.pack(fill="x", pady=2)
                    d = camp.get("date","")[:16].replace("T"," ")
                    tk.Label(row, text=f"  📅 {d}  |  Sent: {camp.get('sent',0)}  |  Failed: {camp.get('failed',0)}  |  Rate: {camp.get('rate',0)}%",
                            bg=C['bg2'], fg=C['text2'], font=("Segoe UI",9), anchor="w").pack(fill="x",padx=10,pady=3)
            except Exception:
                pass
        else:
            tk.Label(hist_frame, text="No campaigns yet. Go to Send tab to launch one!",
                    bg=C['bg'], fg=C['text2'], font=("Segoe UI",10)).pack(pady=20)
    def export_accounts_gui(self):
        fp = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON","*.json")])
        if fp:
            ok = self.am.export_accounts(fp)
            messagebox.showinfo("Export", f"Exported {len(self.am.accounts)} accounts!" if ok else "Export failed")

    # ─── TAB 1: ACCOUNTS ──────────────────────────────────────────
    def build_accounts_tab(self, parent):
        C = self.COLORS
        top = tk.Frame(parent, bg=C['bg'])
        top.pack(fill="x", padx=10, pady=8)
        tk.Button(top, text="+ Add Account", bg=C['accent'], fg=C['bg'], font=("Segoe UI",10,"bold"),
                 relief="flat", padx=15, pady=4, command=self.open_add_account).pack(side="left",padx=3)
        tk.Button(top, text="🔄 Test All", bg=C['bg3'], fg=C['text'], font=("Segoe UI",9),
                 relief="flat", padx=10, pady=4, command=self.test_all_accounts).pack(side="left",padx=3)
        self.acc_stats_label = tk.Label(top, text="", bg=C['bg'], fg=C['text2'], font=("Segoe UI",9))
        self.acc_stats_label.pack(side="right",padx=10)
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.acc_frame = tk.Frame(canvas, bg=C['bg'])
        self.acc_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self.acc_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=10)
        scrollbar.pack(side="right", fill="y")
        self.refresh_accounts_list()
    def refresh_accounts_list(self):
        for w in self.acc_frame.winfo_children(): w.destroy()
        C = self.COLORS
        accounts = self.am.get_all_accounts()
        if not accounts:
            tk.Label(self.acc_frame, text="No email accounts added yet.\nClick '+ Add Account' to get started.",
                    bg=C['bg'], fg=C['text2'], font=("Segoe UI",12), justify="center").pack(pady=50)
            return
        for acc in accounts:
            card = tk.Frame(self.acc_frame, bg=C['bg2'], highlightbackground=C['border'], highlightthickness=1)
            card.pack(fill="x", pady=4, padx=5)
            top = tk.Frame(card, bg=C['bg2'])
            top.pack(fill="x", padx=10, pady=8)
            status_colors = {"connected":C['success'],"failed":C['danger'],"limit_reached":C['warning'],"unknown":C['text2']}
            status_icons = {"connected":"✅","failed":"❌","limit_reached":"⚠️","unknown":"❓"}
            st = acc.get("status","unknown")
            tk.Label(top, text=f"{status_icons.get(st,'❓')} {acc['nickname']}", bg=C['bg2'], fg=C['text'],
                    font=("Segoe UI",11,"bold")).pack(side="left")
            tk.Label(top, text=f"  {acc['email']}", bg=C['bg2'], fg=C['accent'], font=("Segoe UI",10)).pack(side="left",padx=5)
            tk.Label(top, text=st.upper(), bg=C['bg2'], fg=status_colors.get(st,C['text2']),
                    font=("Segoe UI",8,"bold")).pack(side="right",padx=5)
            if acc.get("is_default"):
                tk.Label(top, text="⭐ DEFAULT", bg=C['bg2'], fg=C['warning'], font=("Segoe UI",8,"bold")).pack(side="right",padx=5)
            mid = tk.Frame(card, bg=C['bg2'])
            mid.pack(fill="x", padx=10, pady=(0,5))
            sent = acc.get("sent_today",0); lim = acc.get("daily_limit",500)
            pct = min(sent/max(lim,1), 1.0)
            bar_frame = tk.Frame(mid, bg=C['bg3'], height=8)
            bar_frame.pack(fill="x", pady=2)
            bar_fill = tk.Frame(bar_frame, bg=C['accent'] if pct<0.9 else C['warning'], height=8,
                               width=max(1, int(pct * 400)))
            bar_fill.place(x=0,y=0)
            bar_frame.config(width=400, height=8)
            tk.Label(mid, text=f"Sent: {sent}/{lim} today  |  Remaining: {lim - sent}",
                    bg=C['bg2'], fg=C['text2'], font=("Segoe UI",8)).pack(anchor="w")
            if acc.get("last_error"):
                tk.Label(mid, text=f"Error: {acc['last_error'][:80]}", bg=C['bg2'], fg=C['danger'],
                        font=("Segoe UI",8)).pack(anchor="w")
            btns = tk.Frame(card, bg=C['bg2'])
            btns.pack(fill="x", padx=10, pady=(0,8))
            aid = acc["id"]
            tk.Button(btns, text="Test", bg=C['bg3'], fg=C['text'], font=("Segoe UI",8), relief="flat",
                     command=lambda a=aid: self.test_single_account(a)).pack(side="left",padx=2)
            tk.Button(btns, text="Edit", bg=C['bg3'], fg=C['text'], font=("Segoe UI",8), relief="flat",
                     command=lambda a=aid: self.open_add_account(a)).pack(side="left",padx=2)
            tk.Button(btns, text="Set Default", bg=C['bg3'], fg=C['warning'], font=("Segoe UI",8), relief="flat",
                     command=lambda a=aid: self.set_default(a)).pack(side="left",padx=2)
            tk.Button(btns, text="Delete", bg=C['bg3'], fg=C['danger'], font=("Segoe UI",8), relief="flat",
                     command=lambda a=aid: self.delete_account(a)).pack(side="left",padx=2)
        stats = self.am.get_stats()
        self.acc_stats_label.config(text=f"{stats['total_accounts']} accounts | Capacity: {stats['total_capacity_today']}/day | Remaining: {stats['total_remaining_today']}")
    def open_add_account(self, account_id=None):
        dlg = AddAccountDialog(self.root, self.am, account_id)
        self.root.wait_window(dlg)
        self.refresh_accounts_list()
    def test_single_account(self, aid):
        def _t():
            r = self.am.test_account(aid)
            self.root.after(0, self.refresh_accounts_list)
            self.root.after(0, lambda: messagebox.showinfo("Test Result", r["message"]))
        threading.Thread(target=_t, daemon=True).start()
    def test_all_accounts(self):
        def _t():
            results = self.am.test_all_accounts()
            ok = sum(1 for r in results if r["success"])
            self.root.after(0, self.refresh_accounts_list)
            self.root.after(0, lambda: messagebox.showinfo("Test Results", f"{ok}/{len(results)} accounts connected"))
        threading.Thread(target=_t, daemon=True).start()
    def set_default(self, aid):
        self.am.set_default_account(aid); self.refresh_accounts_list()
    def delete_account(self, aid):
        if messagebox.askyesno("Confirm","Delete this account? This cannot be undone."):
            self.am.delete_account(aid); self.refresh_accounts_list()

    # ─── TAB 2: RECIPIENTS ────────────────────────────────────────
    def build_recipients_tab(self, parent):
        C = self.COLORS
        top = tk.Frame(parent, bg=C['bg'])
        top.pack(fill="x", padx=10, pady=8)
        tk.Button(top, text="📂 Browse Excel/CSV", bg=C['accent'], fg=C['bg'], font=("Segoe UI",10,"bold"),
                 relief="flat", padx=15, pady=4, command=self.browse_excel).pack(side="left",padx=3)
        self.file_label = tk.Label(top, text="No file loaded", bg=C['bg'], fg=C['text2'], font=("Segoe UI",9))
        self.file_label.pack(side="left",padx=10)
        conf = tk.Frame(parent, bg=C['bg'])
        conf.pack(fill="x", padx=10, pady=4)
        tk.Label(conf, text="Email Column:", bg=C['bg'], fg=C['text'], font=("Segoe UI",9)).pack(side="left")
        self.col_var = tk.StringVar()
        self.col_combo = ttk.Combobox(conf, textvariable=self.col_var, state="readonly", width=25)
        self.col_combo.pack(side="left",padx=5)
        tk.Button(conf, text="🔍 Auto-Detect", bg=C['bg3'], fg=C['text'], font=("Segoe UI",8), relief="flat",
                 command=self.auto_detect_col).pack(side="left",padx=3)
        tk.Button(conf, text="✅ Validate", bg=C['bg3'], fg=C['accent'], font=("Segoe UI",8,"bold"), relief="flat",
                 command=self.validate_recipients).pack(side="left",padx=3)
        self.stats_frame = tk.Frame(parent, bg=C['bg'])
        self.stats_frame.pack(fill="x", padx=10, pady=4)
        self.email_stats_label = tk.Label(self.stats_frame, text="", bg=C['bg'], fg=C['text2'], font=("Segoe UI",9))
        self.email_stats_label.pack(anchor="w")
        tree_frame = tk.Frame(parent, bg=C['bg'])
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
        self.tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        self.preview_tree = ttk.Treeview(tree_frame, yscrollcommand=self.tree_scroll_y.set,
                                          xscrollcommand=self.tree_scroll_x.set, style="Custom.Treeview")
        self.tree_scroll_y.config(command=self.preview_tree.yview)
        self.tree_scroll_x.config(command=self.preview_tree.xview)
        self.tree_scroll_y.pack(side="right",fill="y")
        self.tree_scroll_x.pack(side="bottom",fill="x")
        self.preview_tree.pack(fill="both", expand=True)
        style = ttk.Style()
        style.configure("Custom.Treeview", background=C['bg2'], foreground=C['text'], fieldbackground=C['bg2'],
                        font=("Segoe UI",9))
        style.configure("Custom.Treeview.Heading", background=C['bg3'], foreground=C['accent'],
                        font=("Segoe UI",9,"bold"))
        self.vars_label = tk.Label(parent, text="", bg=C['bg'], fg=C['accent'], font=("Segoe UI",9), wraplength=900)
        self.vars_label.pack(anchor="w", padx=10, pady=5)
    def browse_excel(self):
        fp = filedialog.askopenfilename(filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not fp: return
        result = self.ep.load(fp)
        if not result["success"]:
            messagebox.showerror("Error", f"Failed to load file:\n{result['error']}"); return
        self.file_label.config(text=f"📄 {os.path.basename(fp)} ({result['rows']} rows, {result['file_size_kb']}KB)")
        self.col_combo['values'] = result['columns']
        self.auto_detect_col()
    def auto_detect_col(self):
        col = self.ep.auto_detect_email_column()
        if col:
            self.col_var.set(col)
            self.validate_recipients()
        else:
            messagebox.showwarning("Warning","Could not auto-detect email column. Please select manually.")
    def validate_recipients(self):
        col = self.col_var.get()
        if not col: messagebox.showwarning("Warning","Select an email column first."); return
        stats = self.ep.validate_and_load(col)
        if "error" in stats: messagebox.showerror("Error", stats["error"]); return
        C = self.COLORS
        self.email_stats_label.config(
            text=f"✅ Valid: {stats['valid']}  |  ❌ Invalid: {stats['invalid']}  |  "
                 f"🔄 Duplicates: {stats['duplicates']}  |  📝 Typos Fixed: {stats['typos_fixed']}  |  Total: {stats['total_rows']}")
        self.preview_tree.delete(*self.preview_tree.get_children())
        preview = self.ep.get_preview(100)
        if preview:
            cols = list(preview[0].keys())
            self.preview_tree['columns'] = cols
            self.preview_tree['show'] = 'headings'
            for c in cols:
                self.preview_tree.heading(c, text=str(c)) # type: ignore
                self.preview_tree.column(c, width=120, minwidth=60) # type: ignore
            for i, row in enumerate(preview):
                vals = [str(row.get(c,"")) for c in cols]
                tag = "even" if i%2==0 else "odd"
                self.preview_tree.insert("","end",values=vals, tags=(tag,))
            self.preview_tree.tag_configure("even", background=C['bg2'])
            self.preview_tree.tag_configure("odd", background=C['bg3'])
        pvars = self.ep.get_personalization_vars()
        self.vars_label.config(text="Variables: " + "  ".join(pvars))

    # ─── TAB 3: COMPOSE ───────────────────────────────────────────
    def build_compose_tab(self, parent):
        C = self.COLORS
        top = tk.Frame(parent, bg=C['bg'])
        top.pack(fill="x", padx=10, pady=8)
        tk.Label(top, text="Send From:", bg=C['bg'], fg=C['text'], font=("Segoe UI",10,"bold")).pack(side="left")
        self.send_from_var = tk.StringVar()
        self.send_from_combo = ttk.Combobox(top, textvariable=self.send_from_var, state="readonly", width=40)
        self.send_from_combo.pack(side="left", padx=5)
        self.refresh_send_from()
        self.rotate_var = tk.BooleanVar(value=False)
        tk.Checkbutton(top, text="🔄 Account Rotation", variable=self.rotate_var, bg=C['bg'], fg=C['accent'],
                      selectcolor=C['bg3'], activebackground=C['bg'], font=("Segoe UI",9)).pack(side="left",padx=10)
        subj_frame = tk.Frame(parent, bg=C['bg'])
        subj_frame.pack(fill="x", padx=10, pady=4)
        tk.Label(subj_frame, text="Subject:", bg=C['bg'], fg=C['text'], font=("Segoe UI",10)).pack(side="left")
        self.subject_entry = tk.Entry(subj_frame, bg=C['bg3'], fg=C['text'], insertbackground=C['accent'],
                                       font=("Segoe UI",11), width=60)
        self.subject_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.subj_count = tk.Label(subj_frame, text="0", bg=C['bg'], fg=C['text2'], font=("Segoe UI",8))
        self.subj_count.pack(side="left",padx=5)
        self.subject_entry.bind("<KeyRelease>", lambda e: self.subj_count.config(text=f"{len(self.subject_entry.get())} chars"))
        fmt = tk.Frame(parent, bg=C['bg'])
        fmt.pack(anchor="w", padx=10, pady=2)
        self.html_var = tk.BooleanVar(value=True)
        tk.Radiobutton(fmt, text="HTML Email", variable=self.html_var, value=True, bg=C['bg'], fg=C['text'],
                      selectcolor=C['bg3'], activebackground=C['bg']).pack(side="left",padx=5)
        tk.Radiobutton(fmt, text="Plain Text", variable=self.html_var, value=False, bg=C['bg'], fg=C['text'],
                      selectcolor=C['bg3'], activebackground=C['bg']).pack(side="left",padx=5)
        self.body_editor = scrolledtext.ScrolledText(parent, bg=C['bg3'], fg=C['text'], insertbackground=C['accent'],
                                                      font=("Consolas",10), wrap="word", height=14)
        self.body_editor.pack(fill="both", expand=True, padx=10, pady=5)
        vars_frame = tk.Frame(parent, bg=C['bg'])
        vars_frame.pack(fill="x", padx=10, pady=2)
        tk.Label(vars_frame, text="Insert:", bg=C['bg'], fg=C['text2'], font=("Segoe UI",8)).pack(side="left")
        for v in ["{{name}}","{{email}}","{{company}}","{{first_name}}","{{date}}"]:
            tk.Button(vars_frame, text=v, bg=C['bg2'], fg=C['accent'], font=("Segoe UI",8), relief="flat",
                     command=lambda var=v: self.insert_variable(var)).pack(side="left",padx=2)
        bottom = tk.Frame(parent, bg=C['bg'])
        bottom.pack(fill="x", padx=10, pady=5)
        tk.Button(bottom, text="📎 Attach Files", bg=C['bg3'], fg=C['text'], font=("Segoe UI",9), relief="flat",
                 command=self.attach_file).pack(side="left",padx=3)
        self.attach_label = tk.Label(bottom, text="0 attachments", bg=C['bg'], fg=C['text2'], font=("Segoe UI",8))
        self.attach_label.pack(side="left",padx=5)
        tk.Button(bottom, text="🛡 Spam Check", bg=C['accent2'], fg=C['bg'], font=("Segoe UI",9,"bold"), relief="flat",
                 command=self.check_spam_score).pack(side="right",padx=3)
        tk.Button(bottom, text="👁 Preview", bg=C['bg3'], fg=C['text'], font=("Segoe UI",9), relief="flat",
                 command=self.preview_email).pack(side="right",padx=3)
        tk.Button(bottom, text="🧪 Send Test", bg=C['bg3'], fg=C['warning'], font=("Segoe UI",9), relief="flat",
                 command=self.send_test_email).pack(side="right",padx=3)
        tk.Button(bottom, text="💾 Save Template", bg=C['bg3'], fg=C['text'], font=("Segoe UI",9), relief="flat",
                 command=self.save_template).pack(side="right",padx=3)
        tk.Button(bottom, text="📄 Load Template", bg=C['bg3'], fg=C['accent'], font=("Segoe UI",9), relief="flat",
                 command=self.load_template).pack(side="right",padx=3)
    def check_spam_score(self):
        """Analyze current email for spam score."""
        subj = self.subject_entry.get()
        body = self.body_editor.get("1.0","end-1c")
        if not subj and not body:
            messagebox.showinfo("Spam Check","Write a subject and body first."); return
        result = self.spam.analyze(subj, body, self.html_var.get())
        # Build report
        report = f"{result['summary']}\n\n"
        report += f"Passed: {result['pass_count']}  |  Warnings: {result['warn_count']}  |  Failed: {result['fail_count']}\n\n"
        for check in result['checks']:
            icon = "✅" if check['status']=='pass' else "⚠️" if check['status']=='warn' else "❌"
            report += f"{icon} {check['check']}: {check['message']}\n"
        tips = self.spam.get_tips()
        if tips and result['score'] > 30:
            report += "\n📝 Tips to improve:\n" + "\n".join(tips)
        # Show in dialog
        win = tk.Toplevel(self.root)
        win.title("🛡 Spam Score Analysis")
        win.geometry("550x450")
        win.configure(bg=self.COLORS['bg'])
        win.transient(self.root)
        score_color = self.COLORS['success'] if result['score']<=30 else self.COLORS['warning'] if result['score']<=60 else self.COLORS['danger']
        tk.Label(win, text=f"{result['emoji']} Score: {result['score']}/100", bg=self.COLORS['bg'], fg=score_color,
                font=("Segoe UI",20,"bold")).pack(pady=10)
        tk.Label(win, text=result['rating'], bg=self.COLORS['bg'], fg=score_color,
                font=("Segoe UI",12)).pack()
        text = scrolledtext.ScrolledText(win, bg=self.COLORS['bg3'], fg=self.COLORS['text'],
                                        font=("Consolas",9), wrap="word")
        text.pack(fill="both",expand=True,padx=15,pady=10)
        text.insert("1.0", report)
        text.configure(state="disabled")
    def save_contacts_group(self):
        """Save current valid recipients as a contact group."""
        if not self.ep.valid_emails:
            messagebox.showinfo("Contacts","Load and validate recipients first."); return
        name = simpledialog.askstring("Save Group","Group name:", parent=self.root)
        if not name: return
        self.contacts.save_recipients_as_group(name, self.ep.valid_emails)
        messagebox.showinfo("Saved",f"Saved {len(self.ep.valid_emails)} contacts to group '{name}'!")
    def refresh_send_from(self):
        accs = self.am.get_all_accounts()
        vals = [f"{a['nickname']} — {a['email']}" for a in accs]
        self.send_from_combo['values'] = vals
        if vals: self.send_from_combo.current(0)
    def insert_variable(self, var):
        self.body_editor.insert("insert", var)
        self.body_editor.focus_set()
    def attach_file(self):
        fps = filedialog.askopenfilenames()
        if fps:
            self.attachment_paths = list(fps)
            self.attach_label.config(text=f"📎 {len(self.attachment_paths)} files attached")
    def preview_email(self):
        subj = self.subject_entry.get()
        body = self.body_editor.get("1.0","end-1c")
        if not body.strip(): messagebox.showwarning("Warning","Email body is empty"); return
        sample = self.ep.valid_emails[0] if self.ep.valid_emails else {"_email":"test@example.com","name":"John Smith","company":"Acme Corp"}
        p_subj = self.smtp.personalize(subj, sample)
        p_body = self.smtp.personalize(body, sample)
        win = tk.Toplevel(self.root)
        win.title("📧 Email Preview")
        win.geometry("700x500")
        win.configure(bg=self.COLORS['bg'])
        tk.Label(win, text=f"Subject: {p_subj}", bg=self.COLORS['bg'], fg=self.COLORS['accent'],
                font=("Segoe UI",11,"bold"), wraplength=650).pack(padx=10,pady=10,anchor="w")
        txt = scrolledtext.ScrolledText(win, bg=self.COLORS['bg3'], fg=self.COLORS['text'], font=("Segoe UI",10),
                                         wrap="word")
        txt.pack(fill="both",expand=True,padx=10,pady=5)
        txt.insert("1.0", p_body)
        txt.config(state="disabled")
    def send_test_email(self):
        addr = simpledialog.askstring("Test Email","Enter email address to send test to:", parent=self.root)
        if not addr: return
        accs = self.am.get_all_accounts()
        if not accs: messagebox.showerror("Error","Add an email account first"); return
        def _send():
            acc = self.am.get_default_account() or self.am.get_account(accs[0]["id"])
            sample = {"_email":addr,"name":"Test User","company":"Test Company"}
            subj = self.subject_entry.get() or "Test Email from Bulk Email Pro"
            body = self.body_editor.get("1.0","end-1c") or "<p>This is a test email.</p>"
            result = self.smtp.send_one(acc["id"], acc, sample, subj, body, self.html_var.get(), self.attachment_paths)
            self.smtp.disconnect_all()
            msg = f"✅ Test email sent to {addr}!" if result["status"]=="sent" else f"❌ Failed: {result['error']}"
            self.root.after(0, lambda: messagebox.showinfo("Test Result", msg))
        threading.Thread(target=_send, daemon=True).start()
    def save_template(self):
        name = simpledialog.askstring("Save Template","Template name:", parent=self.root)
        if not name: return
        tpath = DATA_DIR / "templates.json"
        templates: dict[str, dict[str, str]] = {}
        if tpath.exists():
            try:
                with open(tpath,"r") as f: templates = json.load(f)
            except: pass
        templates[name] = {"subject":self.subject_entry.get(),"body":self.body_editor.get("1.0","end-1c")} # type: ignore
        with open(tpath,"w") as f: json.dump(templates, f, indent=2)
        messagebox.showinfo("Saved",f"Template '{name}' saved!")
    def load_template(self):
        """Load a saved template into the compose editor."""
        tpath = DATA_DIR / "templates.json"
        # Also check project root templates
        alt_path = Path(__file__).parent.parent / "templates.json"
        templates = {}
        for p in [tpath, alt_path]:
            if p.exists():
                try:
                    with open(p, "r") as f:
                        templates.update(json.load(f))
                except: pass
        if not templates:
            messagebox.showinfo("Templates","No templates saved yet. Use 'Save Template' first.")
            return
        # Show selection dialog
        win = tk.Toplevel(self.root)
        win.title("📄 Load Template")
        win.geometry("400x350")
        win.configure(bg=self.COLORS['bg'])
        win.transient(self.root)
        win.grab_set()
        tk.Label(win, text="Select a template:", bg=self.COLORS['bg'], fg=self.COLORS['text'],
                font=("Segoe UI",11,"bold")).pack(padx=15,pady=10,anchor="w")
        listbox = tk.Listbox(win, bg=self.COLORS['bg3'], fg=self.COLORS['text'],
                            font=("Segoe UI",10), selectbackground=self.COLORS['accent'])
        listbox.pack(fill="both",expand=True,padx=15,pady=5)
        names = list(templates.keys())
        for n in names:
            listbox.insert("end", n)
        def apply():
            sel = listbox.curselection()
            if not sel: return
            name = names[sel[0]]
            t = templates[name]
            self.subject_entry.delete(0,"end")
            self.subject_entry.insert(0, t.get("subject",""))
            self.body_editor.delete("1.0","end")
            self.body_editor.insert("1.0", t.get("body",""))
            win.destroy()
        tk.Button(win, text="✅ Apply Template", bg=self.COLORS['accent'], fg=self.COLORS['bg'],
                 font=("Segoe UI",10,"bold"), relief="flat", padx=15, pady=5,
                 command=apply).pack(pady=10)

    # ─── TAB 4: SEND ──────────────────────────────────────────────
    def build_send_tab(self, parent):
        C = self.COLORS
        self.send_parent = parent
        summary = tk.Frame(parent, bg=C['bg2'], highlightbackground=C['border'], highlightthickness=1)
        summary.pack(fill="x", padx=10, pady=8)
        self.send_summary = tk.Label(summary, text="Configure recipients, compose email, then start sending.",
                                      bg=C['bg2'], fg=C['text'], font=("Segoe UI",10), wraplength=900, justify="left")
        self.send_summary.pack(padx=15, pady=10, anchor="w")
        delay_frame = tk.Frame(parent, bg=C['bg'])
        delay_frame.pack(fill="x", padx=10, pady=4)
        tk.Label(delay_frame, text="Send Delay:", bg=C['bg'], fg=C['text'], font=("Segoe UI",9)).pack(side="left")
        self.delay_var = tk.DoubleVar(value=1.5)
        self.delay_scale = tk.Scale(delay_frame, from_=0.5, to=10.0, resolution=0.5, orient="horizontal",
                                     variable=self.delay_var, bg=C['bg'], fg=C['accent'], troughcolor=C['bg3'],
                                     highlightbackground=C['bg'], length=200, font=("Segoe UI",8))
        self.delay_scale.pack(side="left",padx=5)
        self.delay_label = tk.Label(delay_frame, text="1.5s", bg=C['bg'], fg=C['text2'], font=("Segoe UI",9))
        self.delay_label.pack(side="left",padx=5)
        self.delay_var.trace_add("write", lambda *a: self.delay_label.config(text=f"{self.delay_var.get():.1f}s")) # type: ignore
        btn_frame = tk.Frame(parent, bg=C['bg'])
        btn_frame.pack(fill="x", padx=10, pady=8)
        self.start_btn = tk.Button(btn_frame, text="▶ START SENDING", bg=C['accent'], fg=C['bg'],
                                    font=("Segoe UI",14,"bold"), relief="flat", padx=30, pady=8,
                                    command=self.start_sending)
        self.start_btn.pack(side="left", padx=5)
        self.pause_btn = tk.Button(btn_frame, text="⏸ PAUSE", bg=C['warning'], fg=C['bg'],
                                    font=("Segoe UI",10,"bold"), relief="flat", padx=15, pady=6,
                                    command=self.pause_campaign, state="disabled")
        self.pause_btn.pack(side="left",padx=5)
        self.stop_btn = tk.Button(btn_frame, text="⏹ STOP", bg=C['danger'], fg="white",
                                   font=("Segoe UI",10,"bold"), relief="flat", padx=15, pady=6,
                                   command=self.stop_campaign, state="disabled")
        self.stop_btn.pack(side="left",padx=5)
        prog_frame = tk.Frame(parent, bg=C['bg'])
        prog_frame.pack(fill="x", padx=10, pady=4)
        self.progress_bar = ttk.Progressbar(prog_frame, length=800, mode='determinate')
        self.progress_bar.pack(fill="x")
        self.progress_label = tk.Label(prog_frame, text="", bg=C['bg'], fg=C['text'], font=("Segoe UI",10))
        self.progress_label.pack(anchor="w",pady=2)
        self.stats_label = tk.Label(prog_frame, text="", bg=C['bg'], fg=C['text2'], font=("Segoe UI",9))
        self.stats_label.pack(anchor="w")
        self.send_log = scrolledtext.ScrolledText(parent, bg="#0a0a15", fg=C['text'], font=("Consolas",9),
                                                    wrap="word", height=12, state="disabled")
        self.send_log.pack(fill="both", expand=True, padx=10, pady=5)
        self.send_log.tag_configure("sent", foreground=C['success'])
        self.send_log.tag_configure("failed", foreground=C['danger'])
        self.send_log.tag_configure("event", foreground=C['warning'])
        self.send_log.tag_configure("system", foreground="#5dade2")
    def log_line(self, text, tag="sent"):
        self.send_log.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.send_log.insert("end", f"[{ts}] {text}\n", tag)
        self.send_log.see("end")
        self.send_log.config(state="disabled")
    def start_sending(self):
        emails = self.ep.valid_emails
        if not emails: messagebox.showerror("Error","Load and validate recipients first."); return
        accs = self.am.get_all_accounts()
        if not accs: messagebox.showerror("Error","Add at least one email account."); return
        subj = self.subject_entry.get()
        body = self.body_editor.get("1.0","end-1c")
        if not subj.strip(): messagebox.showerror("Error","Enter a subject line."); return
        if not body.strip(): messagebox.showerror("Error","Enter email body."); return
        total = len(emails)
        est_min = round(total * self.delay_var.get() / 60, 1)
        if not messagebox.askyesno("Confirm Send",
            f"Ready to send {total} emails\nEstimated time: ~{est_min} minutes\n\nContinue?"): return
        self.is_sending = True
        self.is_paused = False
        self.send_results = []
        self.campaign_start_time = datetime.now()
        self.start_btn.config(state="disabled")
        self.pause_btn.config(state="normal")
        self.stop_btn.config(state="normal")
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = 0
        self.log_line(f"🚀 Campaign started — {total} recipients", "system")
        self.send_thread = threading.Thread(target=self.run_send_campaign, daemon=True)
        self.send_thread.start()
    def run_send_campaign(self):
        emails = self.ep.valid_emails
        subj = self.subject_entry.get()
        body = self.body_editor.get("1.0","end-1c")
        is_html = self.html_var.get()
        delay = self.delay_var.get()
        use_rot = self.rotate_var.get()
        results = self.smtp.send_bulk(
            email_list=emails, account_manager=self.am,
            subject_template=subj, body_template=body,
            is_html=is_html, delay=delay, use_rotation=use_rot,
            attachment_paths=self.attachment_paths,
            progress_callback=self.on_email_result,
            stop_flag=lambda: not self.is_sending,
            pause_flag=lambda: self.is_paused)
        self.send_results = results
        self.log_queue.put({"type":"complete"})
    def on_email_result(self, index, total, result):
        self.log_queue.put({"type":"progress","index":index,"total":total,"result":result})
    def process_log_queue(self):
        try:
            while not self.log_queue.empty():
                item = self.log_queue.get_nowait()
                if item["type"] == "progress":
                    r = item["result"]
                    idx, total = item["index"], item["total"]
                    self.progress_bar['value'] = idx
                    pct = round(idx/max(total,1)*100,1)
                    self.progress_label.config(text=f"{idx} of {total} ({pct}%)")
                    sent = sum(1 for x in self.send_results if x.get("status")=="sent")
                    failed = sum(1 for x in self.send_results if x.get("status")=="failed")
                    elapsed = (datetime.now()-self.campaign_start_time).total_seconds() if self.campaign_start_time else 1
                    speed = round(idx/(elapsed/60),1) if elapsed>0 else 0
                    remaining = total - idx
                    eta_sec = int(remaining / (idx/elapsed)) if idx>0 and elapsed>0 else 0
                    eta_str = f"{eta_sec//60}:{eta_sec%60:02d}"
                    self.stats_label.config(text=f"✅ Sent: {sent}  ❌ Failed: {failed}  ⚡ {speed}/min  ⏱ ETA: {eta_str}")
                    st = r.get("status","")
                    if st == "sent":
                        self.log_line(f"✅ {r.get('account_used','')} → {r.get('_email','')}", "sent")
                    elif st == "failed":
                        self.log_line(f"❌ FAILED → {r.get('_email','')} — {r.get('error','')}", "failed")
                    elif st == "rotation":
                        self.log_line(f"🔄 {r.get('message','Account rotation')}", "event")
                    elif st == "warning":
                        self.log_line(f"⚠️ {r.get('error','')}", "event")
                    if r.get("status") in ("sent","failed"):
                        self.send_results.append(r)
                elif item["type"] == "complete":
                    self.on_campaign_complete()
        except Exception:
            pass
        self.root.after(100, self.process_log_queue)
    def on_campaign_complete(self):
        self.is_sending = False
        self.start_btn.config(state="normal")
        self.pause_btn.config(state="disabled")
        self.stop_btn.config(state="disabled")
        duration = (datetime.now()-self.campaign_start_time).total_seconds() if self.campaign_start_time else 0
        sent = sum(1 for r in self.send_results if r.get("status")=="sent")
        failed = sum(1 for r in self.send_results if r.get("status")=="failed")
        total = sent + failed
        rate = round(sent/max(total,1)*100,1)
        dur_str = f"{int(duration//60)}m {int(duration%60)}s"
        self.log_line(f"🏁 Campaign complete! Sent: {sent}, Failed: {failed}, Duration: {dur_str}", "system")
        messagebox.showinfo("Campaign Complete",
            f"🎉 Campaign Complete!\n\n✅ Sent: {sent}\n❌ Failed: {failed}\n"
            f"📊 Success Rate: {rate}%\n⏱ Duration: {dur_str}")
        self.save_campaign_history()
    def pause_campaign(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_btn.config(text="▶ RESUME")
            self.log_line("⏸ Campaign paused", "event")
        else:
            self.pause_btn.config(text="⏸ PAUSE")
            self.log_line("▶ Campaign resumed", "event")
    def stop_campaign(self):
        if messagebox.askyesno("Stop","Stop the campaign? Remaining emails will be skipped."):
            self.is_sending = False
            self.log_line("⏹ Campaign stopped by user", "failed")

    # ─── TAB 5: REPORTS ───────────────────────────────────────────
    def build_reports_tab(self, parent):
        C = self.COLORS
        top = tk.Frame(parent, bg=C['bg'])
        top.pack(fill="x", padx=10, pady=8)
        tk.Label(top, text="📊 Campaign Reports", bg=C['bg'], fg=C['text'], font=("Segoe UI",12,"bold")).pack(side="left")
        tk.Button(top, text="📥 Export Sent (CSV)", bg=C['bg3'], fg=C['success'], font=("Segoe UI",9), relief="flat",
                 command=lambda: self.export_report("sent")).pack(side="right",padx=3)
        tk.Button(top, text="📥 Export Failed (CSV)", bg=C['bg3'], fg=C['danger'], font=("Segoe UI",9), relief="flat",
                 command=lambda: self.export_report("failed")).pack(side="right",padx=3)
        tk.Button(top, text="📥 Export All (CSV)", bg=C['accent'], fg=C['bg'], font=("Segoe UI",9,"bold"), relief="flat",
                 command=lambda: self.export_report("all")).pack(side="right",padx=3)
        self.report_summary = tk.Label(parent, text="No campaign data yet. Send emails to generate a report.",
                                        bg=C['bg'], fg=C['text2'], font=("Segoe UI",10))
        self.report_summary.pack(padx=10, pady=5, anchor="w")
        tree_frame = tk.Frame(parent, bg=C['bg'])
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        cols = ("Email","Status","Account Used","Time","Error")
        self.report_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", style="Custom.Treeview")
        for c in cols:
            self.report_tree.heading(c, text=c)
            self.report_tree.column(c, width=150)
        rsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.report_tree.yview)
        self.report_tree.configure(yscrollcommand=rsb.set)
        rsb.pack(side="right",fill="y")
        self.report_tree.pack(fill="both",expand=True)
    def refresh_report(self):
        self.report_tree.delete(*self.report_tree.get_children())
        C = self.COLORS
        sent = failed = 0
        for r in self.send_results:
            st = r.get("status","")
            if st == "sent": sent += 1
            elif st == "failed": failed += 1
            tag = "sent" if st=="sent" else "failed"
            ts = r.get("timestamp","")
            if "T" in ts: ts = ts.split("T")[1][:8]
            self.report_tree.insert("","end", values=(r.get("_email",""),st,r.get("account_used",""),ts,r.get("error","")), tags=(tag,))
        self.report_tree.tag_configure("sent", foreground=C['success'])
        self.report_tree.tag_configure("failed", foreground=C['danger'])
        total = sent+failed
        rate = round(sent/max(total,1)*100,1)
        self.report_summary.config(text=f"Total: {total}  |  ✅ Sent: {sent}  |  ❌ Failed: {failed}  |  Success Rate: {rate}%")
    def export_report(self, filter_status="all"):
        if not self.send_results: messagebox.showinfo("Info","No results to export."); return
        fp = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not fp: return
        import csv
        rows = self.send_results
        if filter_status == "sent": rows = [r for r in rows if r.get("status")=="sent"]
        elif filter_status == "failed": rows = [r for r in rows if r.get("status")=="failed"]
        try:
            with open(fp, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=["_email","status","account_used","account_id","timestamp","error"])
                w.writeheader(); w.writerows(rows)
            messagebox.showinfo("Exported",f"Exported {len(rows)} rows to {fp}")
        except Exception as e:
            messagebox.showerror("Error",f"Export failed: {e}")
    def save_campaign_history(self):
        self.refresh_report()
        hist_path = DATA_DIR / "campaigns.json"
        history: list = []
        if hist_path.exists():
            try:
                with open(hist_path,"r") as f: history = json.load(f)
            except: pass
        sent = sum(1 for r in self.send_results if r.get("status")=="sent")
        failed = sum(1 for r in self.send_results if r.get("status")=="failed")
        campaign = {"date":datetime.now().isoformat(),"total":sent+failed,"sent":sent,"failed":failed,
                    "rate":round(sent/max(sent+failed,1)*100,1)}
        history.append(campaign)
        if len(history)>50: history = history[-50:] # type: ignore
        with open(hist_path,"w") as f: json.dump(history,f,indent=2)

    # ─── MENU ─────────────────────────────────────────────────────
    def build_menu(self):
        menu = tk.Menu(self.root, bg="#1a1a2e", fg="#e0e0e0", activebackground="#4ecca3", activeforeground="#0d0d1a")
        self.root.config(menu=menu)
        file_menu = tk.Menu(menu, tearoff=0, bg="#1a1a2e", fg="#e0e0e0")
        menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Excel File  (Ctrl+O)", command=self.browse_excel)
        file_menu.add_command(label="Export Report  (Ctrl+S)", command=lambda: self.export_report("all"))
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        acc_menu = tk.Menu(menu, tearoff=0, bg="#1a1a2e", fg="#e0e0e0")
        menu.add_cascade(label="Accounts", menu=acc_menu)
        acc_menu.add_command(label="Add New Account  (Ctrl+A)", command=self.open_add_account)
        acc_menu.add_command(label="Test All Accounts", command=self.test_all_accounts)
        camp_menu = tk.Menu(menu, tearoff=0, bg="#1a1a2e", fg="#e0e0e0")
        menu.add_cascade(label="Campaign", menu=camp_menu)
        camp_menu.add_command(label="Start Sending  (Ctrl+Enter)", command=self.start_sending)
        camp_menu.add_command(label="Pause/Resume  (Ctrl+P)", command=self.pause_campaign)
        camp_menu.add_command(label="Stop  (Escape)", command=self.stop_campaign)
        help_menu = tk.Menu(menu, tearoff=0, bg="#1a1a2e", fg="#e0e0e0")
        menu.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Gmail Setup Guide", command=lambda: messagebox.showinfo("Gmail Setup",
            "1. Go to myaccount.google.com\n2. Security > 2-Step Verification (enable)\n3. Security > App Passwords\n"
            "4. Select Mail, Generate\n5. Use the 16-char password in Bulk Email Pro"))
        help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About",
            "Bulk Email Pro v3.0\nSend bulk emails from your own accounts.\nDark Theme • Account Rotation • Excel Upload"))
        self.root.bind("<Control-o>", lambda e: self.browse_excel())
        self.root.bind("<Control-a>", lambda e: self.open_add_account())
        self.root.bind("<Control-Return>", lambda e: self.start_sending())
        self.root.bind("<Control-p>", lambda e: self.pause_campaign())
        self.root.bind("<Escape>", lambda e: self.stop_campaign() if self.is_sending else None)
        self.root.bind("<Control-s>", lambda e: self.export_report("all"))
    def start_queue_processor(self):
        self.root.after(100, self.process_log_queue)

def main():
    root = tk.Tk()
    app = BulkEmailProApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
