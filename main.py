import tkinter as tk  # GUI library for building the app interface
from tkinter import filedialog, messagebox, ttk  # GUI library for building the app interface
import csv  # Used for reading and writing CSV files
import os  # Used for file system path operations
import logging  # Used to log debug and error messages
import re  # Regex module for pattern matching in strings
import ctypes  # Windows system calls for DPI settings
import subprocess  # Used to open files/folders externally
import json  # To read/write configuration files in JSON format
from datetime import datetime  # To get current date/time for audit logs

# Enable High DPI scaling for crisp text on Windows
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Ensure proper scaling on high-DPI displays (Windows)
except Exception:
    pass

# Setup logging
logging.basicConfig(  # Setup logging configuration
    filename='converter_debug.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


class ExamSoftToBlackboardApp:  # Define the main application class
    def __init__(self, root):  # Define a function
        self.root = root
        self.root.title("ExamSoft to Blackboard Converter")
        self.root.geometry("680x750")
        self.root.minsize(680, 500)
        self.center_window()
        self.root.configure(bg="#f3f3f3")
        self.default_font = ("Segoe UI", 10)
        self.root.option_add("*Font", self.default_font)

        self.config_file = os.path.join(os.path.expanduser("~"), ".examsoft_converter_config")
        config = self.load_config()
        self.last_dir = config.get("last_dir", os.path.expanduser("~"))
        self.mapping_history = config.get("mapping_history", {})

        self.examsoft_file_path = ""
        self.blackboard_file_path = ""
        self.examsoft_score_col = ""
        self.bb_usernames = set()

        self.setup_scrollable_frame()
        self.setup_ui()
        self.setup_shortcuts()
        self.es_btn.focus_set()

    def center_window(self):  # Function to center the window on the screen
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x, y = (self.root.winfo_screenwidth() // 2) - (w // 2), (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f'{w}x{h}+{x}+{y}')

    def load_config(self):  # Load saved configuration (e.g., file paths)
        try:
            path = os.path.join(os.path.expanduser("~"), ".examsoft_converter_config")
            if os.path.exists(path):
                with open(path, 'r') as f: return json.load(f)
        except:
            pass
        return {}

    def save_config(self):  # Save current config to file
        try:
            with open(self.config_file, 'w') as f:
                json.dump({"last_dir": self.last_dir, "mapping_history": self.mapping_history}, f)
        except:
            pass

    def setup_shortcuts(self):  # Bind keyboard shortcuts to actions
        self.root.bind("<Control-o>", lambda e: self.select_examsoft_file())
        self.root.bind("<Control-b>", lambda e: self.select_blackboard_file())
        self.root.bind("<Control-r>", lambda e: self.reset_app())
        self.root.bind("<F1>", lambda e: self.show_help())

    def setup_scrollable_frame(self):  # Create scrollable area for main interface
        self.main_canvas = tk.Canvas(self.root, bg="#f3f3f3", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.main_canvas.yview)
        self.scroll_content = tk.Frame(self.main_canvas, bg="#f3f3f3")
        self.scroll_content.bind("<Configure>",
                                 lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.window_item = self.main_canvas.create_window((0, 0), window=self.scroll_content, anchor="nw")
        self.main_canvas.bind("<Configure>", lambda e: self.main_canvas.itemconfig(self.window_item, width=e.width))
        self.main_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.root.bind_all("<MouseWheel>", lambda e: self.main_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def _read_csv(self, path):  # Define a function
        self.root.config(cursor="wait");
        self.root.update()
        try:
            encoding = 'utf-8-sig'
            try:
                with open(path, mode='r', newline='', encoding=encoding) as f:
                    f.read(1024)
            except UnicodeDecodeError:
                encoding = 'latin-1'
            with open(path, mode='r', newline='', encoding=encoding) as f:
                sample = f.read(2048);
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=',;\t')
                except:
                    dialect = csv.excel
                reader = csv.DictReader(f, dialect=dialect);
                data = list(reader);
                headers = reader.fieldnames
            if headers: headers = [col.strip() for col in headers]
            self.root.config(cursor="");
            return data, headers
        except Exception as e:
            self.root.config(cursor="");
            logging.error(f"Read Error: {e}");
            raise

    def _find_header(self, headers, keywords, default=""):  # Define a function
        for h in headers:
            if any(k.lower() in h.lower() for k in keywords): return h
        return default

    def _truncate_filename(self, filename, max_length=45):  # Define a function
        if len(filename) <= max_length: return filename
        keep = (max_length - 3) // 2
        return f"{filename[:keep]}...{filename[-keep:]}"

    def setup_ui(self):  # Build the main user interface layout
        style = ttk.Style();
        style.theme_use('vista')
        style.configure("Highlight.TLabelframe", background="#e7f3ff")
        style.configure("Success.TLabelframe", background="#f0fff0")

        h_outer = tk.Frame(self.scroll_content, bg="#ffffff", highlightthickness=1, highlightbackground="#e0e0e0");
        h_outer.pack(fill=tk.X)
        h_cont = tk.Frame(h_outer, bg="#ffffff", padx=30, pady=15);
        h_cont.pack(fill=tk.X)
        tk.Label(h_cont, text="ExamSoft to Blackboard Converter", font=("Segoe UI", 16, "bold"), bg="#ffffff",
                 fg="#0078d4").pack(side=tk.LEFT, anchor="w")
        util_frame = tk.Frame(h_cont, bg="#ffffff");
        util_frame.pack(side=tk.RIGHT, anchor="e")
        ttk.Button(util_frame, text="Reset (Ctrl+R)", command=self.reset_app).pack(side=tk.LEFT, padx=5)
        ttk.Button(util_frame, text="Help (F1)", command=self.show_help).pack(side=tk.LEFT)

        self.es_section = ttk.LabelFrame(self.scroll_content, text=" STEP 1: ExamSoft Scores ");
        self.es_section.pack(pady=10, fill=tk.X, padx=30)
        es_c = tk.Frame(self.es_section, cursor="hand2");
        es_c.pack(fill=tk.X, padx=10, pady=10);
        es_c.bind("<Button-1>", lambda e: self.select_examsoft_file())
        self.es_btn = ttk.Button(es_c, text="Browse...", command=self.select_examsoft_file, width=12);
        self.es_btn.pack(side=tk.LEFT)
        self.es_label = tk.Label(es_c, text="Select source file", fg="#666666", font=("Segoe UI", 9, "italic"),
                                 anchor="w");
        self.es_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.es_label.bind("<Button-1>", lambda e: self.select_examsoft_file())
        self.es_drop_frame = tk.Frame(self.es_section);
        tk.Label(self.es_drop_frame, text="Score column:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.es_col_var = tk.StringVar();
        self.es_combo = ttk.Combobox(self.es_drop_frame, textvariable=self.es_col_var, state="readonly", width=25);
        self.es_combo.bind("<<ComboboxSelected>>", self.on_es_combo_select);
        self.es_combo.pack(side=tk.LEFT, pady=5)

        self.sep1 = ttk.Separator(self.scroll_content, orient='horizontal')
        self.bb_section = ttk.LabelFrame(self.scroll_content, text=" STEP 2: Blackboard Gradebook CSV ");
        bb_c = tk.Frame(self.bb_section, cursor="hand2");
        bb_c.pack(fill=tk.X, padx=10, pady=10);
        bb_c.bind("<Button-1>", lambda e: self.select_blackboard_file())
        self.bb_btn = ttk.Button(bb_c, text="Browse...", command=self.select_blackboard_file, width=12);
        self.bb_btn.pack(side=tk.LEFT)
        self.bb_label = tk.Label(bb_c, text="Select your Blackboard gradebook CSV", fg="#666666",
                                 font=("Segoe UI", 9, "italic"), anchor="w");
        self.bb_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.bb_label.bind("<Button-1>", lambda e: self.select_blackboard_file())
        self.bb_drop_frame = tk.Frame(self.bb_section);
        tk.Label(self.bb_drop_frame, text="Target column:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.bb_col_var = tk.StringVar();
        self.bb_combo = ttk.Combobox(self.bb_drop_frame, textvariable=self.bb_col_var, state="readonly", width=30);
        self.bb_combo.bind("<<ComboboxSelected>>", lambda e: self.update_preview());
        self.bb_combo.pack(side=tk.LEFT, pady=5)
        self.audit_status_label = tk.Label(self.bb_section, font=("Segoe UI", 9, "bold"))

        self.sep2 = ttk.Separator(self.scroll_content, orient='horizontal')
        self.preview_group = tk.Frame(self.scroll_content, bg="#f3f3f3");
        tk.Label(self.preview_group, text="Mapping Preview", font=("Segoe UI", 10, "bold"), bg="#f3f3f3").pack(
            anchor="w")
        self.tree = ttk.Treeview(self.preview_group, columns=("User", "ID", "Score"), show='headings', height=8);
        self.tree.heading("User", text="User");
        self.tree.heading("ID", text="ID");
        self.tree.heading("Score", text="Score")
        self.tree.column("User", width=150);
        self.tree.column("ID", width=120);
        self.tree.column("Score", width=100);
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.tag_configure("missing", foreground="#d83b01");
        self.tree.tag_configure("modified", foreground="#0078d4")
        self.generate_btn = ttk.Button(self.preview_group, text="Generate Blackboard Import File",
                                       command=self.process_files);
        self.generate_btn.pack(pady=10, anchor="e")

        self.success_panel = tk.Frame(self.scroll_content, bg="#ffffff", highlightthickness=1,
                                      highlightbackground="#107c10")
        self.success_label = tk.Label(self.success_panel, font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#107c10")
        self.success_path_label = tk.Label(self.success_panel, text="", bg="#ffffff", font=("Segoe UI", 9),
                                           wraplength=500)

    def reset_app(self):  # Reset app to initial state (clears selections)
        self.examsoft_file_path = "";
        self.blackboard_file_path = "";
        self.examsoft_score_col = "";
        self.bb_usernames = set()
        self.es_label.config(text="Select source file", fg="#666666");
        self.bb_label.config(text="Select target template", fg="#666666")
        self.es_btn.config(text="Browse...");
        self.bb_btn.config(text="Browse...")
        self.es_drop_frame.pack_forget();
        self.bb_drop_frame.pack_forget();
        self.bb_section.pack_forget();
        self.preview_group.pack_forget();
        self.success_panel.pack_forget()
        self.sep1.pack_forget();
        self.sep2.pack_forget();
        self.audit_status_label.pack_forget()
        self.es_section.configure(style="TLabelframe");
        self.bb_section.configure(style="TLabelframe")
        self.es_btn.focus_set();
        [self.tree.delete(i) for i in self.tree.get_children()]

    def show_help(self):  # Show help/instruction window
        help_window = tk.Toplevel(self.root)
        help_window.title("Instructions & Help")
        help_window.geometry("550x700")
        help_window.configure(bg="white")
        help_window.transient(self.root)

        canvas = tk.Canvas(help_window, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        sf = tk.Frame(canvas, bg="white", padx=25, pady=20)
        sf.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=sf, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y");
        canvas.pack(side="left", fill="both", expand=True)

        def _on_mousewheel(event):  # Define a function
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _on_help_close():  # Define a function
            canvas.unbind_all("<MouseWheel>")
            help_window.destroy()

        help_window.protocol("WM_DELETE_WINDOW", _on_help_close)

        sections = [
            ("Exporting Assessment Scores as a CSV in ExamSoft",
             "Open the Assessment in ExamSoft.\n\nSelect Reporting / Scoring, then choose Exam Taker Results.\n\nIn the Exam Taker Results options, select:\n• Exam Taker Name\n• Email\n• Desired Score field (e.g., Percentage Score)\n\nClick View Report.\n\nSet the display to Show 250 to ensure all exam takers are visible.\n\nClick the Export CSV icon on the right side of the screen."),
            ("Exporting Assessment Scores as a CSV in Blackboard Ultra",
             "Open the course in Blackboard Ultra.\n\nNavigate to the Gradebook and select Download Grades.\n\nUnder Select Records, choose Full Gradebook.\n\nUnder Record Details, select the assessment name you want included.\n\nSet File Type to Comma-Separated Values (CSV).\n\nClick Download to export the file."),
            ("Uploading a CSV File of Grades into Blackboard Ultra",
             "Open the course in Blackboard Ultra.\n\nNavigate to the Gradebook.\n\nSelect Upload Grades.\n\nChoose Upload Local File and select the CSV file created by this program.\n\nReview the column-mapping screen to confirm the correct assessment and score column.\n\nClick Upload to import the grades.")
        ]
        for title, body in sections:
            tk.Label(sf, text=title, font=("Segoe UI", 12, "bold"), bg="white", wraplength=460, justify=tk.LEFT).pack(
                anchor="w", pady=(10, 5))
            tk.Label(sf, text=body, bg="white", wraplength=460, justify=tk.LEFT, font=("Segoe UI", 10)).pack(anchor="w",
                                                                                                             pady=(0,
                                                                                                                   15))
        ttk.Button(sf, text="Got it", command=_on_help_close).pack(pady=10)

    def on_es_combo_select(self, event):  # Define a function
        self.examsoft_score_col = self.es_col_var.get();
        self.update_preview();
        self.perform_instant_audit()

    def select_examsoft_file(self):  # Prompt user to select ExamSoft CSV
        path = filedialog.askopenfilename(initialdir=self.last_dir, title="Select ExamSoft CSV",
                                          filetypes=[("CSV files", "*.csv")])
        if path:
            try:
                self.last_dir = os.path.dirname(path);
                self.save_config()
                d, h = self._read_csv(path)
                if not self._find_header(h, ["email"]): messagebox.showerror("Error", "Missing 'Email' column."); return
                self.examsoft_file_path = path;
                self.es_label.config(text=f"📄 {self._truncate_filename(os.path.basename(path))} ({len(d)} rows)",
                                     fg="#0078d4", font=("Segoe UI", 10, "bold"));
                self.es_btn.config(text="Change File")
                self.es_section.configure(style="Success.TLabelframe");
                self.identify_examsoft_score_column(h)
                self.sep1.pack(fill=tk.X, padx=50, pady=5);
                self.bb_section.pack(pady=10, fill=tk.X, padx=30);
                self.update_preview()
                self.root.after(100, lambda: self.main_canvas.yview_moveto(0.3));
                self.perform_instant_audit()
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def select_blackboard_file(self):  # Prompt user to select Blackboard CSV
        path = filedialog.askopenfilename(initialdir=self.last_dir, title="Select Blackboard CSV",
                                          filetypes=[("CSV files", "*.csv")])
        if path:
            try:
                self.last_dir = os.path.dirname(path);
                self.save_config()
                d, h = self._read_csv(path);
                self.blackboard_file_path = path
                u_col = self._find_header(h, ["username"])
                self.bb_usernames = {row.get(u_col, "").lower().strip() for row in d if row.get(u_col)}
                self.bb_label.config(text=f"📄 {self._truncate_filename(os.path.basename(path))} ({len(d)} rows)",
                                     fg="#0078d4", font=("Segoe UI", 10, "bold"));
                self.bb_btn.config(text="Change File")
                ex = {"last name", "first name", "username", "student id", "last access", "availability"}
                filtered = [col for col in h if col.lower() not in ex]
                if filtered:
                    self.bb_section.configure(style="Success.TLabelframe");
                    self.bb_drop_frame.pack(side=tk.TOP, anchor="w", padx=10, pady=5);
                    self.bb_combo['values'] = filtered
                    hist = self.mapping_history.get(self.examsoft_score_col)
                    self.bb_col_var.set(hist if hist in filtered else next(
                        (c for c in filtered if self.examsoft_score_col.lower() in c.lower()), filtered[0]))
                    self.sep2.pack(fill=tk.X, padx=50, pady=5);
                    self.preview_group.pack(pady=15, fill=tk.BOTH, expand=True, padx=30)
                self.update_preview();
                self.root.after(100, lambda: self.main_canvas.yview_moveto(1.0));
                self.perform_instant_audit()
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def perform_instant_audit(self):  # Compare usernames between both files and show match %
        if not self.examsoft_file_path or not self.blackboard_file_path: return
        try:
            es_d, es_h = self._read_csv(self.examsoft_file_path);
            bb_d, bb_h = self._read_csv(self.blackboard_file_path)
            e_c = self._find_header(es_h, ["email"]);
            u_c = self._find_header(bb_h, ["username"])
            es_u = {r.get(e_c, "").split('@')[0].split('+')[0].lower().strip() for r in es_d if r.get(e_c)}
            bb_u = {r.get(u_c, "").lower().strip() for r in bb_d if r.get(u_c)}
            matches = len(es_u.intersection(bb_u));
            total = len(es_u)
            percent = round((matches / total) * 100) if total > 0 else 0
            self.audit_status_label.config(text=f"📊 Roster Sync: {matches} of {total} matched ({percent}%)",
                                           fg="#107c10" if percent > 95 else "#d83b01")
            self.audit_status_label.pack(pady=5, anchor="w", padx=5)
        except:
            pass

    def identify_examsoft_score_column(self, h):  # Try to detect which column has the ExamSoft scores
        cands = ["%", "pts", "raw", "score", "percentage"];
        found = [col for col in h if any(c.lower() in col.lower() for c in cands)]
        if found:
            self.es_section.configure(style="Highlight.TLabelframe");
            self.es_drop_frame.pack(side=tk.TOP, anchor="w", padx=10, pady=5);
            self.es_combo['values'] = found
            choice = next((col for col in found if "%" in col), found[0]);
            self.es_col_var.set(choice);
            self.examsoft_score_col = choice

    def show_success_state(self, p, c, stats=None):  # Display success message and export options
        self.preview_group.pack_forget();
        self.success_panel.pack(pady=20, fill=tk.X, padx=30)
        self.success_label.config(text=f"✅ {c} Scores Mapped Successfully!");
        self.success_label.pack(pady=(10, 0))
        if stats: tk.Label(self.success_panel,
                           text=f"Avg: {stats['avg']}% | High: {stats['high']} | Low: {stats['low']}",
                           font=("Segoe UI", 10, "bold"), bg="#ffffff").pack(pady=5)
        self.success_path_label.config(text=p);
        self.success_path_label.pack(pady=10)
        ttk.Button(self.success_panel, text="Open Exported CSV", command=lambda: os.startfile(p)).pack(pady=5)
        ttk.Button(self.success_panel, text="Open Folder",
                   command=lambda: subprocess.Popen(f'explorer /select,"{os.path.normpath(p)}"')).pack(pady=5)
        ttk.Button(self.success_panel, text="Exit Application", command=self.root.destroy).pack(side=tk.BOTTOM, pady=10)

    def clean_score(self, v):  # Remove extra characters and convert score to number
        if not v: return "0"
        n = "".join(re.findall(r'[0-9.]+', str(v)))
        try:
            return str(round(float(n or 0), 2))
        except:
            return "0"

    def update_preview(self):  # Show preview of mapped scores
        [self.tree.delete(i) for i in self.tree.get_children()]
        if not self.examsoft_file_path or not self.examsoft_score_col: return
        try:
            d, h = self._read_csv(self.examsoft_file_path);
            e_col = self._find_header(h, ["email"])
            for i, row in enumerate(d):
                if i >= 12: break
                raw_e = row.get(e_col, "").strip()
                if not raw_e: continue
                u = raw_e.split('@')[0].split('+')[0].lower().strip()
                r = row.get(self.examsoft_score_col, "");
                c = self.clean_score(r)
                tags = []
                if c != r.strip(): tags.append("modified")
                if self.bb_usernames and u not in self.bb_usernames: tags.append("missing")
                self.tree.insert("", tk.END, values=(u, row.get("StudentID", row.get("Student ID", "")), c), tags=tags)
            self.tree.tag_configure("missing", foreground="#d83b01");
            self.tree.tag_configure("modified", foreground="#0078d4")
        except Exception:
            pass

    def process_files(self):  # Main logic to generate Blackboard import file
        target = self.bb_col_var.get();
        out = filedialog.asksaveasfilename(initialdir=self.last_dir, title="Save File", defaultextension=".csv",
                                           initialfile=f"BB_Import_{os.path.basename(self.examsoft_file_path)}",
                                           filetypes=[("CSV files", "*.csv")])
        if not out: return
        self.mapping_history[self.examsoft_score_col] = target;
        self.save_config()
        self.root.config(cursor="wait");
        self.root.update()
        try:
            bb_u = {};
            bb_d, bb_h = self._read_csv(self.blackboard_file_path)
            u_c, f_c, l_c = self._find_header(bb_h, ["username"]), self._find_header(bb_h,
                                                                                     ["first"]), self._find_header(bb_h,
                                                                                                                   ["last"])
            for r in bb_d:
                uname = r.get(u_c, "").lower().strip()
                if uname: bb_u[uname] = f"{r.get(f_c)} {r.get(l_c)}"
            unique_s = {};
            not_in_bb = [];
            zero_count = 0;
            scores = [];
            es_d, es_h = self._read_csv(self.examsoft_file_path)
            e_c, sid_c = self._find_header(es_h, ["email"]), self._find_header(es_h,
                                                                               ["student", "id"]) or self._find_header(
                es_h, ["email"])
            for idx, row in enumerate(es_d, start=2):
                raw_e = row.get(e_c, "").strip()
                if not raw_e: continue
                username = raw_e.split('@')[0].split('+')[0].lower().strip();
                s_val = self.clean_score(row.get(self.examsoft_score_col, "0"))
                sid = row.get(sid_c, "").strip() or username;
                s_num = float(s_val or 0);
                scores.append(s_num);
                zero_count += (1 if s_num == 0 else 0)
                record = {"Last Name": row.get(self._find_header(es_h, ["last"])),
                          "First Name": row.get(self._find_header(es_h, ["first"])), "Username": username,
                          "Student ID": sid, target: s_val}
                if sid in unique_s:
                    if s_num > float(unique_s[sid][target]): unique_s[sid] = record
                else:
                    unique_s[sid] = record
                if bb_u and username not in bb_u: not_in_bb.append(
                    f"Row {idx}: {record['First Name']} {record['Last Name']} ({username})")
            rows = list(unique_s.values())
            if rows and (zero_count / len(rows)) > 0.2:
                self.root.config(cursor="")
                if not messagebox.askyesno("Confirm", f"{zero_count} students have a score of 0. Continue?"):
                    return
                self.root.config(cursor="wait");
                self.root.update()
            stats = {"avg": round(sum(scores) / len(scores), 1) if scores else 0, "high": max(scores) if scores else 0,
                     "low": min(scores) if scores else 0}
            not_in_es = [f"- {name} ({u})" for u, name in bb_u.items() if u not in {r['Username'] for r in rows}]
            with open(out, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["Last Name", "First Name", "Username", "Student ID", target]);
                writer.writeheader();
                writer.writerows(rows)
            self.root.config(cursor="");
            self.show_success_state(out, len(rows), stats)
            if not_in_bb or not_in_es:
                a_path = os.path.normpath(os.path.join(os.path.dirname(out), "Audit_Report.txt"))
                with open(a_path, 'w', encoding='utf-8') as f:
                    f.write("EXAMSOFT TO BLACKBOARD AUDIT REPORT\n");
                    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n");
                    f.write("=" * 35 + "\n\n")
                    f.write(f"STATS: Avg {stats['avg']}% | High {stats['high']} | Low {stats['low']}\n\n")
                    if not_in_bb: f.write(
                        f"⚠️ IN EXAMSOFT ONLY ({len(not_in_bb)}):\n" + "\n".join([f"- {s}" for s in not_in_bb]) + "\n")
                    if not_in_es: f.write(f"\n⚠️ MISSING SCORES ({len(not_in_es)}):\n" + "\n".join(not_in_es) + "\n")
                if messagebox.askyesno("Audit Warning", "Roster mismatches detected. Open Audit Report?"): os.startfile(
                    a_path)
        except Exception as e:
            self.root.config(cursor=""); messagebox.showerror("Error", str(e))


if __name__ == '__main__':  # Launch the application
    root = tk.Tk();
    app = ExamSoftToBlackboardApp(root);
    root.mainloop()



def select_examsoft_file(self):  # Open a file dialog to choose the ExamSoft CSV file
    path = filedialog.askopenfilename(initialdir=self.last_dir, title="Select ExamSoft CSV", filetypes=[
        ("CSV files", "*.csv")])  # Show a dialog for the user to select a file to open
    if path:
        try:
            self.last_dir = os.path.dirname(path);
            self.save_config()  # Remember the last directory the user accessed
            d, h = self._read_csv(path)  # Call internal method to read and parse a CSV file
            if not self._find_header(h, ["email"]): messagebox.showerror("Error",
                                                                         "Missing 'Email' column."); return  # Helper function to locate a column by matching its name to expected keywords
            self.examsoft_file_path = path;
            self.es_label.config(text=f"📄 {self._truncate_filename(os.path.basename(path))} ({len(d)} rows)",
                                 fg="#0078d4", font=("Segoe UI", 10, "bold"));
            self.es_btn.config(text="Change File")  # Update the UI label to show selected file name and row count
            self.es_section.configure(style="Success.TLabelframe");
            self.identify_examsoft_score_column(h)
            self.sep1.pack(fill=tk.X, padx=50, pady=5);
            self.bb_section.pack(pady=10, fill=tk.X, padx=30);
            self.update_preview()
            self.root.after(100, lambda: self.main_canvas.yview_moveto(0.3));
            self.perform_instant_audit()  # Schedule a scroll or update to the UI shortly after the current event loop
        except Exception as e:
            messagebox.showerror("Error", str(e))  # Show an error message dialog if something went wrong


def select_blackboard_file(self):  # Open a file dialog to choose the Blackboard CSV file
    path = filedialog.askopenfilename(initialdir=self.last_dir, title="Select Blackboard CSV", filetypes=[
        ("CSV files", "*.csv")])  # Show a dialog for the user to select a file to open
    if path:
        try:
            self.last_dir = os.path.dirname(path);
            self.save_config()  # Remember the last directory the user accessed
            d, h = self._read_csv(path);
            self.blackboard_file_path = path  # Call internal method to read and parse a CSV file
            u_col = self._find_header(h,
                                      ["username"])  # Helper function to locate a column by matching its name to expected keywords
            self.bb_usernames = {row.get(u_col, "").lower().strip() for row in d if row.get(u_col)}
            self.bb_label.config(text=f"📄 {self._truncate_filename(os.path.basename(path))} ({len(d)} rows)",
                                 fg="#0078d4", font=("Segoe UI", 10, "bold"));
            self.bb_btn.config(text="Change File")  # Update the UI label to show selected file name and row count
            ex = {"last name", "first name", "username", "student id", "last access", "availability"}
            filtered = [col for col in h if col.lower() not in ex]
            if filtered:
                self.bb_section.configure(style="Success.TLabelframe");
                self.bb_drop_frame.pack(side=tk.TOP, anchor="w", padx=10, pady=5);
                self.bb_combo['values'] = filtered  # Populate the dropdown menu with column names
                hist = self.mapping_history.get(self.examsoft_score_col)
                self.bb_col_var.set(hist if hist in filtered else next(
                    (c for c in filtered if self.examsoft_score_col.lower() in c.lower()), filtered[0]))
                self.sep2.pack(fill=tk.X, padx=50, pady=5);
                self.preview_group.pack(pady=15, fill=tk.BOTH, expand=True, padx=30)  # Show the preview table section
            self.update_preview();
            self.root.after(100, lambda: self.main_canvas.yview_moveto(1.0));
            self.perform_instant_audit()  # Schedule a scroll or update to the UI shortly after the current event loop
        except Exception as e:
            messagebox.showerror("Error", str(e))  # Show an error message dialog if something went wrong


def perform_instant_audit(self):  # Function definition
    if not self.examsoft_file_path or not self.blackboard_file_path: return
    try:
        es_d, es_h = self._read_csv(self.examsoft_file_path);
        bb_d, bb_h = self._read_csv(self.blackboard_file_path)  # Call internal method to read and parse a CSV file
        e_c = self._find_header(es_h, ["email"]);
        u_c = self._find_header(bb_h,
                                ["username"])  # Helper function to locate a column by matching its name to expected keywords
        es_u = {r.get(e_c, "").split('@')[0].split('+')[0].lower().strip() for r in es_d if r.get(e_c)}
        bb_u = {r.get(u_c, "").lower().strip() for r in bb_d if r.get(u_c)}
        matches = len(es_u.intersection(bb_u));
        total = len(es_u)
        percent = round((matches / total) * 100) if total > 0 else 0
        self.audit_status_label.config(text=f"📊 Roster Sync: {matches} of {total} matched ({percent}%)",
                                       fg="#107c10" if percent > 95 else "#d83b01")
        self.audit_status_label.pack(pady=5, anchor="w", padx=5)
    except:
        pass


def identify_examsoft_score_column(self, h):  # Function definition
    cands = ["%", "pts", "raw", "score", "percentage"];
    found = [col for col in h if any(c.lower() in col.lower() for c in cands)]
    if found:
        self.es_section.configure(style="Highlight.TLabelframe");
        self.es_drop_frame.pack(side=tk.TOP, anchor="w", padx=10, pady=5);
        self.es_combo['values'] = found
        choice = next((col for col in found if "%" in col), found[0]);
        self.es_col_var.set(choice);
        self.examsoft_score_col = choice


def show_success_state(self, p, c, stats=None):  # Function definition
    self.preview_group.pack_forget();
    self.success_panel.pack(pady=20, fill=tk.X, padx=30)  # Show the preview table section
    self.success_label.config(text=f"✅ {c} Scores Mapped Successfully!");
    self.success_label.pack(pady=(10, 0))
    if stats: tk.Label(self.success_panel, text=f"Avg: {stats['avg']}% | High: {stats['high']} | Low: {stats['low']}",
                       font=("Segoe UI", 10, "bold"), bg="#ffffff").pack(pady=5)
    self.success_path_label.config(text=p);
    self.success_path_label.pack(pady=10)
    ttk.Button(self.success_panel, text="Open Exported CSV", command=lambda: os.startfile(p)).pack(pady=5)
    ttk.Button(self.success_panel, text="Open Folder",
               command=lambda: subprocess.Popen(f'explorer /select,"{os.path.normpath(p)}"')).pack(pady=5)
    ttk.Button(self.success_panel, text="Exit Application", command=self.root.destroy).pack(side=tk.BOTTOM, pady=10)


def clean_score(self, v):  # Function definition
    if not v: return "0"
    n = "".join(re.findall(r'[0-9.]+', str(v)))
    try:
        return str(round(float(n or 0), 2))
    except:
        return "0"


def update_preview(self):  # Function definition
    [self.tree.delete(i) for i in self.tree.get_children()]
    if not self.examsoft_file_path or not self.examsoft_score_col: return
    try:
        d, h = self._read_csv(self.examsoft_file_path);
        e_col = self._find_header(h, ["email"])  # Call internal method to read and parse a CSV file
        for i, row in enumerate(d):
            if i >= 12: break
            raw_e = row.get(e_col, "").strip()
            if not raw_e: continue
            u = raw_e.split('@')[0].split('+')[0].lower().strip()
            r = row.get(self.examsoft_score_col, "");
            c = self.clean_score(r)
            tags = []
            if c != r.strip(): tags.append("modified")
            if self.bb_usernames and u not in self.bb_usernames: tags.append("missing")
            self.tree.insert("", tk.END, values=(u, row.get("StudentID", row.get("Student ID", "")), c),
                             tags=tags)  # Add a new row to the preview TreeView
        self.tree.tag_configure("missing", foreground="#d83b01");
        self.tree.tag_configure("modified", foreground="#0078d4")
    except Exception:
        pass


def process_files(self):  # Function definition
    target = self.bb_col_var.get();
    out = filedialog.asksaveasfilename(initialdir=self.last_dir, title="Save File", defaultextension=".csv",
                                       initialfile=f"BB_Import_{os.path.basename(self.examsoft_file_path)}", filetypes=[
            ("CSV files", "*.csv")])  # Remember the last directory the user accessed
    if not out: return
    self.mapping_history[self.examsoft_score_col] = target;
    self.save_config()
    self.root.config(cursor="wait");
    self.root.update()
    try:
        bb_u = {};
        bb_d, bb_h = self._read_csv(self.blackboard_file_path)  # Call internal method to read and parse a CSV file
        u_c, f_c, l_c = self._find_header(bb_h, ["username"]), self._find_header(bb_h, ["first"]), self._find_header(
            bb_h, ["last"])  # Helper function to locate a column by matching its name to expected keywords
        for r in bb_d:
            uname = r.get(u_c, "").lower().strip()
            if uname: bb_u[uname] = f"{r.get(f_c)} {r.get(l_c)}"
        unique_s = {};
        not_in_bb = [];
        zero_count = 0;
        scores = [];
        es_d, es_h = self._read_csv(self.examsoft_file_path)  # Call internal method to read and parse a CSV file
        e_c, sid_c = self._find_header(es_h, ["email"]), self._find_header(es_h,
                                                                           ["student", "id"]) or self._find_header(es_h,
                                                                                                                   ["email"])  # Helper function to locate a column by matching its name to expected keywords
        for idx, row in enumerate(es_d, start=2):
            raw_e = row.get(e_c, "").strip()
            if not raw_e: continue
            username = raw_e.split('@')[0].split('+')[0].lower().strip();
            s_val = self.clean_score(row.get(self.examsoft_score_col, "0"))
            sid = row.get(sid_c, "").strip() or username;
            s_num = float(s_val or 0);
            scores.append(s_num);
            zero_count += (1 if s_num == 0 else 0)
            record = {"Last Name": row.get(self._find_header(es_h, ["last"])),
                      "First Name": row.get(self._find_header(es_h, ["first"])), "Username": username,
                      "Student ID": sid,
                      target: s_val}  # Helper function to locate a column by matching its name to expected keywords
            if sid in unique_s:
                if s_num > float(unique_s[sid][target]): unique_s[sid] = record
            else:
                unique_s[sid] = record
            if bb_u and username not in bb_u: not_in_bb.append(
                f"Row {idx}: {record['First Name']} {record['Last Name']} ({username})")
        rows = list(unique_s.values())
        if rows and (zero_count / len(rows)) > 0.2:
            self.root.config(cursor="")
            if not messagebox.askyesno("Confirm", f"{zero_count} students have a score of 0. Continue?"):
                return
            self.root.config(cursor="wait");
            self.root.update()
        stats = {"avg": round(sum(scores) / len(scores), 1) if scores else 0, "high": max(scores) if scores else 0,
                 "low": min(scores) if scores else 0}
        not_in_es = [f"- {name} ({u})" for u, name in bb_u.items() if u not in {r['Username'] for r in rows}]
        with open(out, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=["Last Name", "First Name", "Username", "Student ID", target]);
            writer.writeheader();
            writer.writerows(rows)
        self.root.config(cursor="");
        self.show_success_state(out, len(rows), stats)
        if not_in_bb or not_in_es:
            a_path = os.path.normpath(os.path.join(os.path.dirname(out), "Audit_Report.txt"))
            with open(a_path, 'w', encoding='utf-8') as f:
                f.write("EXAMSOFT TO BLACKBOARD AUDIT REPORT\n");
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n");
                f.write("=" * 35 + "\n\n")
                f.write(f"STATS: Avg {stats['avg']}% | High {stats['high']} | Low {stats['low']}\n\n")
                if not_in_bb: f.write(
                    f"⚠️ IN EXAMSOFT ONLY ({len(not_in_bb)}):\n" + "\n".join([f"- {s}" for s in not_in_bb]) + "\n")
                if not_in_es: f.write(f"\n⚠️ MISSING SCORES ({len(not_in_es)}):\n" + "\n".join(not_in_es) + "\n")
            if messagebox.askyesno("Audit Warning", "Roster mismatches detected. Open Audit Report?"): os.startfile(
                a_path)
    except Exception as e:
        self.root.config(cursor=""); messagebox.showerror("Error",
                                                          str(e))  # Show an error message dialog if something went wrong


if __name__ == '__main__':
    root = tk.Tk();
    app = ExamSoftToBlackboardApp(root);
    root.mainloop()
