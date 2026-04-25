import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
from datetime import datetime

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

FILE_NAME = "exam_backend.xlsx"
REQUIRED_COLUMNS = ["Subject", "Topic", "Status", "Exam Date", "Priority", "Notes"]


# ------------------ Data File Initialization ------------------
def initialize_file():
    if not os.path.exists(FILE_NAME):
        df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        df.to_excel(FILE_NAME, index=False)
    else:
        try:
            df = pd.read_excel(FILE_NAME)
        except Exception:
            df = pd.DataFrame(columns=REQUIRED_COLUMNS)

        df.fillna("", inplace=True)

        for col in REQUIRED_COLUMNS:
            if col not in df.columns:
                if col == "Priority":
                    df[col] = "Medium"
                else:
                    df[col] = ""

        df = df[REQUIRED_COLUMNS]
        df.to_excel(FILE_NAME, index=False)


def read_data():
    try:
        df = pd.read_excel(FILE_NAME)
    except Exception:
        df = pd.DataFrame(columns=REQUIRED_COLUMNS)

    df.fillna("", inplace=True)

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            if col == "Priority":
                df[col] = "Medium"
            else:
                df[col] = ""

    df = df[REQUIRED_COLUMNS]
    return df


def write_data(df):
    df.fillna("", inplace=True)

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            if col == "Priority":
                df[col] = "Medium"
            else:
                df[col] = ""

    df = df[REQUIRED_COLUMNS]
    df.to_excel(FILE_NAME, index=False)


# ------------------ Main App ------------------
class SmartExamPlanner:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Exam Planner, Revision Tracker & Weak Area Analyzer")
        self.root.geometry("1320x800")
        self.root.configure(bg="#f5f7fa")
        self.root.minsize(1180, 720)

        self.ai_file_path = ""
        self.ai_busy = False

        initialize_file()
        self.create_widgets()
        self.load_table()
        self.update_dashboard()
        self.load_weak_areas()
        self.refresh_ai_subjects()
        self.show_recommendation()
        self.check_deadline_alerts_on_start()

    # ------------------ Widgets ------------------
    def create_widgets(self):
        title = tk.Label(
            self.root,
            text="Smart Exam Planner, Revision Tracker & Weak Area Analyzer",
            font=("Arial", 20, "bold"),
            bg="#f5f7fa",
            fg="#1e3d59"
        )
        title.pack(pady=10)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=10)

        self.planner_tab = tk.Frame(self.notebook, bg="#f5f7fa")
        self.analytics_tab = tk.Frame(self.notebook, bg="#f5f7fa")
        self.weak_tab = tk.Frame(self.notebook, bg="#f5f7fa")
        self.charts_tab = tk.Frame(self.notebook, bg="#f5f7fa")
        self.ai_tab = tk.Frame(self.notebook, bg="#f5f7fa")

        self.notebook.add(self.planner_tab, text="Planner")
        self.notebook.add(self.analytics_tab, text="Analytics")
        self.notebook.add(self.weak_tab, text="Weak Areas")
        self.notebook.add(self.charts_tab, text="Charts")
        self.notebook.add(self.ai_tab, text="AI Assistant")

        self.create_planner_tab()
        self.create_analytics_tab()
        self.create_weak_tab()
        self.create_charts_tab()
        self.create_ai_tab()

    # ------------------ Planner Tab ------------------
    def create_planner_tab(self):
        form_frame = tk.LabelFrame(
            self.planner_tab,
            text="Topic Entry",
            font=("Arial", 11, "bold"),
            bg="white",
            fg="#1e3d59",
            padx=10,
            pady=10
        )
        form_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(form_frame, text="Subject", bg="white", font=("Arial", 10, "bold")).grid(
            row=0, column=0, padx=10, pady=6, sticky="w"
        )
        self.subject_entry = tk.Entry(form_frame, width=25)
        self.subject_entry.grid(row=0, column=1, padx=10, pady=6)

        tk.Label(form_frame, text="Topic", bg="white", font=("Arial", 10, "bold")).grid(
            row=0, column=2, padx=10, pady=6, sticky="w"
        )
        self.topic_entry = tk.Entry(form_frame, width=25)
        self.topic_entry.grid(row=0, column=3, padx=10, pady=6)

        tk.Label(form_frame, text="Status", bg="white", font=("Arial", 10, "bold")).grid(
            row=1, column=0, padx=10, pady=6, sticky="w"
        )
        self.status_combo = ttk.Combobox(
            form_frame,
            values=["Pending", "Revising", "Done"],
            state="readonly",
            width=22
        )
        self.status_combo.grid(row=1, column=1, padx=10, pady=6)

        tk.Label(form_frame, text="Exam Date (YYYY-MM-DD)", bg="white", font=("Arial", 10, "bold")).grid(
            row=1, column=2, padx=10, pady=6, sticky="w"
        )
        self.date_entry = tk.Entry(form_frame, width=25)
        self.date_entry.grid(row=1, column=3, padx=10, pady=6)

        tk.Label(form_frame, text="Priority", bg="white", font=("Arial", 10, "bold")).grid(
            row=2, column=0, padx=10, pady=6, sticky="w"
        )
        self.priority_combo = ttk.Combobox(
            form_frame,
            values=["High", "Medium", "Low"],
            state="readonly",
            width=22
        )
        self.priority_combo.grid(row=2, column=1, padx=10, pady=6)

        tk.Label(form_frame, text="Notes", bg="white", font=("Arial", 10, "bold")).grid(
            row=2, column=2, padx=10, pady=6, sticky="w"
        )
        self.notes_entry = tk.Entry(form_frame, width=25)
        self.notes_entry.grid(row=2, column=3, padx=10, pady=6)

        button_frame = tk.Frame(form_frame, bg="white")
        button_frame.grid(row=3, column=0, columnspan=4, pady=12)

        tk.Button(button_frame, text="Add Topic", width=15, command=self.add_record, bg="#1e90ff", fg="white").grid(
            row=0, column=0, padx=5, pady=5
        )
        tk.Button(button_frame, text="Update Selected", width=15, command=self.update_selected, bg="#ff9800", fg="white").grid(
            row=0, column=1, padx=5, pady=5
        )
        tk.Button(button_frame, text="Delete Selected", width=15, command=self.delete_selected, bg="#e74c3c", fg="white").grid(
            row=0, column=2, padx=5, pady=5
        )
        tk.Button(button_frame, text="Clear Fields", width=15, command=self.clear_fields, bg="#7f8c8d", fg="white").grid(
            row=0, column=3, padx=5, pady=5
        )
        tk.Button(button_frame, text="Sort by Exam Date", width=15, command=self.sort_by_exam_date, bg="#16a085", fg="white").grid(
            row=0, column=4, padx=5, pady=5
        )

        filter_frame = tk.LabelFrame(
            self.planner_tab,
            text="Filters",
            font=("Arial", 11, "bold"),
            bg="white",
            fg="#1e3d59",
            padx=10,
            pady=10
        )
        filter_frame.pack(fill="x", padx=10, pady=6)

        tk.Label(filter_frame, text="Filter Status", bg="white", font=("Arial", 10, "bold")).grid(
            row=0, column=0, padx=10, pady=5
        )
        self.filter_status = ttk.Combobox(
            filter_frame,
            values=["All", "Pending", "Revising", "Done"],
            state="readonly",
            width=18
        )
        self.filter_status.grid(row=0, column=1, padx=10, pady=5)
        self.filter_status.set("All")

        tk.Label(filter_frame, text="Filter Priority", bg="white", font=("Arial", 10, "bold")).grid(
            row=0, column=2, padx=10, pady=5
        )
        self.filter_priority = ttk.Combobox(
            filter_frame,
            values=["All", "High", "Medium", "Low"],
            state="readonly",
            width=18
        )
        self.filter_priority.grid(row=0, column=3, padx=10, pady=5)
        self.filter_priority.set("All")

        tk.Button(filter_frame, text="Apply Filter", width=15, command=self.apply_filters, bg="#34495e", fg="white").grid(
            row=0, column=4, padx=10, pady=5
        )
        tk.Button(filter_frame, text="Show All", width=15, command=self.load_table, bg="#2ecc71", fg="white").grid(
            row=0, column=5, padx=10, pady=5
        )

        table_frame = tk.Frame(self.planner_tab, bg="#f5f7fa")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        cols = ("Subject", "Topic", "Status", "Exam Date", "Priority", "Notes", "Days Left", "Condition")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=18)

        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=120)

        self.tree.column("Topic", width=170)
        self.tree.column("Notes", width=220)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

    # ------------------ Analytics Tab ------------------
    def create_analytics_tab(self):
        self.analytics_frame = tk.Frame(self.analytics_tab, bg="#f5f7fa")
        self.analytics_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.total_label = tk.Label(
            self.analytics_frame, text="Total Topics: 0",
            font=("Arial", 14, "bold"), bg="#f5f7fa"
        )
        self.total_label.pack(pady=8)

        self.done_label = tk.Label(
            self.analytics_frame, text="Completed Topics: 0",
            font=("Arial", 14, "bold"), fg="green", bg="#f5f7fa"
        )
        self.done_label.pack(pady=8)

        self.pending_label = tk.Label(
            self.analytics_frame, text="Pending Topics: 0",
            font=("Arial", 14, "bold"), fg="red", bg="#f5f7fa"
        )
        self.pending_label.pack(pady=8)

        self.revising_label = tk.Label(
            self.analytics_frame, text="Revising Topics: 0",
            font=("Arial", 14, "bold"), fg="orange", bg="#f5f7fa"
        )
        self.revising_label.pack(pady=8)

        self.overdue_label = tk.Label(
            self.analytics_frame, text="Overdue Topics: 0",
            font=("Arial", 14, "bold"), fg="maroon", bg="#f5f7fa"
        )
        self.overdue_label.pack(pady=8)

        self.high_priority_label = tk.Label(
            self.analytics_frame, text="High Priority Topics: 0",
            font=("Arial", 14, "bold"), fg="#6a1b9a", bg="#f5f7fa"
        )
        self.high_priority_label.pack(pady=8)

        self.progress_label = tk.Label(
            self.analytics_frame, text="Overall Completion: 0%",
            font=("Arial", 14, "bold"), fg="#1e3d59", bg="#f5f7fa"
        )
        self.progress_label.pack(pady=12)

        self.recommendation_label = tk.Label(
            self.analytics_frame,
            text="Recommendation: -",
            font=("Arial", 13, "bold"),
            fg="#003366",
            bg="#f5f7fa",
            wraplength=900,
            justify="center"
        )
        self.recommendation_label.pack(pady=15)

    # ------------------ Weak Areas Tab ------------------
    def create_weak_tab(self):
        tk.Label(
            self.weak_tab,
            text="Weak Areas (Pending + High Priority + Overdue + Urgent)",
            font=("Arial", 16, "bold"),
            bg="#f5f7fa",
            fg="red"
        ).pack(pady=10)

        self.weak_text = tk.Text(self.weak_tab, width=140, height=30, font=("Consolas", 11))
        self.weak_text.pack(padx=20, pady=10)

    # ------------------ Charts Tab ------------------
    def create_charts_tab(self):
        top_frame = tk.Frame(self.charts_tab, bg="#f5f7fa")
        top_frame.pack(fill="x", pady=10)

        tk.Label(
            top_frame,
            text="Visual Analytics Dashboard",
            font=("Arial", 16, "bold"),
            bg="#f5f7fa",
            fg="#1e3d59"
        ).pack(pady=5)

        button_frame = tk.Frame(self.charts_tab, bg="#f5f7fa")
        button_frame.pack(pady=10)

        tk.Button(
            button_frame,
            text="Show Status Pie Chart",
            width=22,
            command=self.show_status_pie_chart,
            bg="#3498db",
            fg="white"
        ).grid(row=0, column=0, padx=10, pady=5)

        tk.Button(
            button_frame,
            text="Show Priority Pie Chart",
            width=22,
            command=self.show_priority_pie_chart,
            bg="#9b59b6",
            fg="white"
        ).grid(row=0, column=1, padx=10, pady=5)

        tk.Button(
            button_frame,
            text="Show Subject Bar Chart",
            width=22,
            command=self.show_subject_bar_chart,
            bg="#27ae60",
            fg="white"
        ).grid(row=0, column=2, padx=10, pady=5)

        self.chart_display_frame = tk.Frame(self.charts_tab, bg="white", bd=2, relief="groove")
        self.chart_display_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # ------------------ AI Tab ------------------
    def create_ai_tab(self):
        top_frame = tk.LabelFrame(
            self.ai_tab,
            text="AI Study Assistant",
            font=("Arial", 11, "bold"),
            bg="white",
            fg="#1e3d59",
            padx=10,
            pady=10
        )
        top_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(top_frame, text="Weak Subject", bg="white", font=("Arial", 10, "bold")).grid(
            row=0, column=0, padx=10, pady=6, sticky="w"
        )
        self.ai_subject_combo = ttk.Combobox(top_frame, state="readonly", width=28)
        self.ai_subject_combo.grid(row=0, column=1, padx=10, pady=6)

        tk.Button(
            top_frame,
            text="Refresh Weak Subjects",
            command=self.refresh_ai_subjects,
            bg="#2ecc71",
            fg="white",
            width=18
        ).grid(row=0, column=2, padx=10, pady=6)

        tk.Button(
            top_frame,
            text="Upload Study File",
            command=self.select_ai_file,
            bg="#3498db",
            fg="white",
            width=18
        ).grid(row=0, column=3, padx=10, pady=6)

        self.ai_file_label = tk.Label(
            top_frame,
            text="No file selected",
            bg="white",
            fg="#555555",
            font=("Arial", 10),
            anchor="w",
            width=60
        )
        self.ai_file_label.grid(row=1, column=0, columnspan=4, padx=10, pady=6, sticky="w")

        tk.Label(
            top_frame,
            text="Important Topics / Extra Instructions",
            bg="white",
            font=("Arial", 10, "bold")
        ).grid(row=2, column=0, padx=10, pady=6, sticky="nw")

        self.ai_topics_text = tk.Text(top_frame, height=5, width=92, font=("Arial", 10))
        self.ai_topics_text.grid(row=2, column=1, columnspan=3, padx=10, pady=6, sticky="w")

        tk.Label(
            top_frame,
            text="Custom Question",
            bg="white",
            font=("Arial", 10, "bold")
        ).grid(row=3, column=0, padx=10, pady=6, sticky="w")

        self.ai_question_entry = tk.Entry(top_frame, width=80, font=("Arial", 10))
        self.ai_question_entry.grid(row=3, column=1, columnspan=3, padx=10, pady=6, sticky="w")

        button_frame = tk.Frame(top_frame, bg="white")
        button_frame.grid(row=4, column=0, columnspan=4, pady=10)

        tk.Button(
            button_frame,
            text="Generate Short Notes",
            command=lambda: self.generate_ai_notes("short_notes"),
            bg="#8e44ad",
            fg="white",
            width=20
        ).grid(row=0, column=0, padx=6, pady=5)

        tk.Button(
            button_frame,
            text="Important Points",
            command=lambda: self.generate_ai_notes("important_points"),
            bg="#e67e22",
            fg="white",
            width=20
        ).grid(row=0, column=1, padx=6, pady=5)

        tk.Button(
            button_frame,
            text="Quick Revision",
            command=lambda: self.generate_ai_notes("quick_revision"),
            bg="#16a085",
            fg="white",
            width=20
        ).grid(row=0, column=2, padx=6, pady=5)

        tk.Button(
            button_frame,
            text="Ask AI",
            command=lambda: self.generate_ai_notes("custom"),
            bg="#34495e",
            fg="white",
            width=20
        ).grid(row=0, column=3, padx=6, pady=5)

        self.ai_status_label = tk.Label(
            self.ai_tab,
            text="AI Status: Ready",
            bg="#f5f7fa",
            fg="#1e3d59",
            font=("Arial", 10, "bold")
        )
        self.ai_status_label.pack(pady=(0, 6))

        self.ai_output = tk.Text(self.ai_tab, wrap="word", font=("Arial", 11), bg="white")
        self.ai_output.pack(fill="both", expand=True, padx=12, pady=8)

    # ------------------ Validation ------------------
    def validate_inputs(self, subject, topic, status, exam_date, priority):
        if not subject or not topic or not status or not exam_date or not priority:
            messagebox.showerror("Error", "Please fill all required fields.")
            return False

        try:
            datetime.strptime(exam_date, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Error", "Exam date must be in YYYY-MM-DD format.")
            return False

        return True

    # ------------------ Helpers ------------------
    def get_condition(self, status, days_left):
        if pd.isna(days_left):
            return "Unknown"
        if days_left < 0 and status != "Done":
            return "Overdue"
        if days_left <= 3 and status != "Done":
            return "Urgent"
        return "Normal"

    def get_priority_rank(self, priority):
        order = {"High": 1, "Medium": 2, "Low": 3}
        return order.get(priority, 4)

    def clear_chart_frame(self):
        for widget in self.chart_display_frame.winfo_children():
            widget.destroy()

    def get_weak_df(self):
        df = read_data()
        if df.empty:
            return df

        df["Exam Date"] = pd.to_datetime(df["Exam Date"], errors="coerce")
        today = pd.Timestamp.today().normalize()
        df["Days Left"] = (df["Exam Date"] - today).dt.days
        df["Condition"] = df.apply(lambda row: self.get_condition(row["Status"], row["Days Left"]), axis=1)

        weak_df = df[
            ((df["Status"] == "Pending") & (df["Priority"] == "High")) |
            ((df["Exam Date"] < today) & (df["Status"] != "Done")) |
            ((df["Status"] == "Pending") & (df["Days Left"] <= 3))
        ].copy()

        return weak_df

    # ------------------ CRUD ------------------
    def add_record(self):
        subject = self.subject_entry.get().strip()
        topic = self.topic_entry.get().strip()
        status = self.status_combo.get().strip()
        exam_date = self.date_entry.get().strip()
        priority = self.priority_combo.get().strip()
        notes = self.notes_entry.get().strip()

        if not self.validate_inputs(subject, topic, status, exam_date, priority):
            return

        df = read_data()

        new_row = {
            "Subject": subject,
            "Topic": topic,
            "Status": status,
            "Exam Date": exam_date,
            "Priority": priority,
            "Notes": notes
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        write_data(df)

        self.clear_fields()
        self.load_table()
        self.update_dashboard()
        self.load_weak_areas()
        self.refresh_ai_subjects()
        self.show_recommendation()
        self.check_deadline_alerts()

        messagebox.showinfo("Success", "Topic added successfully.")

    def update_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a row to update.")
            return

        tree_id = selected[0]
        if not tree_id.isdigit():
            messagebox.showerror("Error", "Invalid row selected.")
            return

        row_index = int(tree_id)

        subject = self.subject_entry.get().strip()
        topic = self.topic_entry.get().strip()
        status = self.status_combo.get().strip()
        exam_date = self.date_entry.get().strip()
        priority = self.priority_combo.get().strip()
        notes = self.notes_entry.get().strip()

        if not self.validate_inputs(subject, topic, status, exam_date, priority):
            return

        df = read_data()

        if row_index >= len(df):
            messagebox.showerror("Error", "Selected record no longer exists.")
            return

        df.at[row_index, "Subject"] = subject
        df.at[row_index, "Topic"] = topic
        df.at[row_index, "Status"] = status
        df.at[row_index, "Exam Date"] = exam_date
        df.at[row_index, "Priority"] = priority
        df.at[row_index, "Notes"] = notes

        write_data(df)

        self.clear_fields()
        self.load_table()
        self.update_dashboard()
        self.load_weak_areas()
        self.refresh_ai_subjects()
        self.show_recommendation()

        messagebox.showinfo("Updated", "Selected record updated successfully.")

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a row to delete.")
            return

        confirm = messagebox.askyesno("Confirm Delete", "Do you want to delete the selected row?")
        if not confirm:
            return

        df = read_data()
        indexes = sorted([int(item) for item in selected if item.isdigit()], reverse=True)

        for idx in indexes:
            if 0 <= idx < len(df):
                df = df.drop(index=idx)

        df = df.reset_index(drop=True)
        write_data(df)

        self.clear_fields()
        self.load_table()
        self.update_dashboard()
        self.load_weak_areas()
        self.refresh_ai_subjects()
        self.show_recommendation()

        messagebox.showinfo("Deleted", "Selected row deleted successfully.")

    # ------------------ Display ------------------
    def load_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        df = read_data()
        if df.empty:
            return

        df["Exam Date"] = pd.to_datetime(df["Exam Date"], errors="coerce")
        today = pd.Timestamp.today().normalize()
        df["Days Left"] = (df["Exam Date"] - today).dt.days
        df["Condition"] = df.apply(lambda row: self.get_condition(row["Status"], row["Days Left"]), axis=1)

        for i, row in df.iterrows():
            exam_date_str = ""
            if pd.notna(row["Exam Date"]):
                exam_date_str = row["Exam Date"].strftime("%Y-%m-%d")

            self.tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    row["Subject"],
                    row["Topic"],
                    row["Status"],
                    exam_date_str,
                    row["Priority"],
                    row["Notes"],
                    row["Days Left"],
                    row["Condition"]
                )
            )

    def apply_filters(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        df = read_data()
        if df.empty:
            return

        status_filter = self.filter_status.get()
        priority_filter = self.filter_priority.get()

        if status_filter != "All":
            df = df[df["Status"] == status_filter]

        if priority_filter != "All":
            df = df[df["Priority"] == priority_filter]

        if df.empty:
            messagebox.showinfo("No Results", "No records match the selected filters.")
            return

        df["Exam Date"] = pd.to_datetime(df["Exam Date"], errors="coerce")
        today = pd.Timestamp.today().normalize()
        df["Days Left"] = (df["Exam Date"] - today).dt.days
        df["Condition"] = df.apply(lambda row: self.get_condition(row["Status"], row["Days Left"]), axis=1)

        for i, row in df.iterrows():
            exam_date_str = ""
            if pd.notna(row["Exam Date"]):
                exam_date_str = row["Exam Date"].strftime("%Y-%m-%d")

            self.tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    row["Subject"],
                    row["Topic"],
                    row["Status"],
                    exam_date_str,
                    row["Priority"],
                    row["Notes"],
                    row["Days Left"],
                    row["Condition"]
                )
            )

    def sort_by_exam_date(self):
        df = read_data()
        if df.empty:
            messagebox.showinfo("Info", "No data available to sort.")
            return

        df["Exam Date"] = pd.to_datetime(df["Exam Date"], errors="coerce")
        df = df.sort_values(by="Exam Date", na_position="last").reset_index(drop=True)
        df["Exam Date"] = df["Exam Date"].dt.strftime("%Y-%m-%d")
        write_data(df)
        self.load_table()

        messagebox.showinfo("Sorted", "Records sorted by exam date.")

    def on_row_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return

        values = self.tree.item(selected[0], "values")
        if not values:
            return

        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, values[0])

        self.topic_entry.delete(0, tk.END)
        self.topic_entry.insert(0, values[1])

        self.status_combo.set(values[2])

        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, values[3])

        self.priority_combo.set(values[4])

        self.notes_entry.delete(0, tk.END)
        self.notes_entry.insert(0, values[5])

    # ------------------ Dashboard ------------------
    def update_dashboard(self):
        df = read_data()

        total = len(df)
        done = len(df[df["Status"] == "Done"])
        pending = len(df[df["Status"] == "Pending"])
        revising = len(df[df["Status"] == "Revising"])
        high_priority = len(df[df["Priority"] == "High"])

        overdue = 0
        if not df.empty:
            df["Exam Date"] = pd.to_datetime(df["Exam Date"], errors="coerce")
            today = pd.Timestamp.today().normalize()
            overdue = len(df[(df["Exam Date"] < today) & (df["Status"] != "Done")])

        completion = 0 if total == 0 else (done / total) * 100

        self.total_label.config(text=f"Total Topics: {total}")
        self.done_label.config(text=f"Completed Topics: {done}")
        self.pending_label.config(text=f"Pending Topics: {pending}")
        self.revising_label.config(text=f"Revising Topics: {revising}")
        self.overdue_label.config(text=f"Overdue Topics: {overdue}")
        self.high_priority_label.config(text=f"High Priority Topics: {high_priority}")
        self.progress_label.config(text=f"Overall Completion: {completion:.2f}%")

    # ------------------ Weak Areas ------------------
    def load_weak_areas(self):
        weak_df = self.get_weak_df()
        self.weak_text.delete("1.0", tk.END)

        if weak_df.empty:
            self.weak_text.insert(tk.END, "No major weak areas detected.")
            return

        weak_df["Exam Date"] = weak_df["Exam Date"].dt.strftime("%Y-%m-%d")
        display_columns = ["Subject", "Topic", "Status", "Exam Date", "Priority", "Notes", "Days Left", "Condition"]
        self.weak_text.insert(tk.END, weak_df[display_columns].to_string(index=False))

    # ------------------ AI Subject Refresh ------------------
    def refresh_ai_subjects(self):
        weak_df = self.get_weak_df()
        subjects = []

        if not weak_df.empty and "Subject" in weak_df.columns:
            subjects = sorted([str(s) for s in weak_df["Subject"].dropna().unique() if str(s).strip()])

        self.ai_subject_combo["values"] = subjects

        if subjects:
            current = self.ai_subject_combo.get().strip()
            if current not in subjects:
                self.ai_subject_combo.set(subjects[0])
        else:
            self.ai_subject_combo.set("")

    # ------------------ Recommendation Engine ------------------
    def show_recommendation(self):
        df = read_data()

        if df.empty:
            self.recommendation_label.config(
                text="Recommendation: Add topics to get smart study suggestions."
            )
            return

        df["Exam Date"] = pd.to_datetime(df["Exam Date"], errors="coerce")
        today = pd.Timestamp.today().normalize()

        pending_df = df[df["Status"] != "Done"].copy()
        if pending_df.empty:
            self.recommendation_label.config(
                text="Recommendation: Great job. All topics are completed."
            )
            return

        pending_df["Days Left"] = (pending_df["Exam Date"] - today).dt.days
        pending_df["Priority Rank"] = pending_df["Priority"].apply(self.get_priority_rank)

        pending_df = pending_df.sort_values(
            by=["Days Left", "Priority Rank"],
            ascending=[True, True],
            na_position="last"
        )

        first = pending_df.iloc[0]

        exam_date_text = "Unknown"
        if pd.notna(first["Exam Date"]):
            exam_date_text = first["Exam Date"].strftime("%Y-%m-%d")

        msg = (
            f"Recommendation: Focus first on '{first['Topic']}' from {first['Subject']} "
            f"because it is {first['Priority']} priority and the exam is on {exam_date_text}."
        )
        self.recommendation_label.config(text=msg)

    # ------------------ Alerts ------------------
    def check_deadline_alerts(self):
        df = read_data()
        if df.empty:
            return

        df["Exam Date"] = pd.to_datetime(df["Exam Date"], errors="coerce")
        today = pd.Timestamp.today().normalize()
        df["Days Left"] = (df["Exam Date"] - today).dt.days

        urgent = df[
            (df["Days Left"] <= 3) &
            (df["Days Left"] >= 0) &
            (df["Status"] != "Done")
        ]

        if not urgent.empty:
            topics = "\n".join([
                f"{row['Subject']} - {row['Topic']} ({row['Days Left']} days left)"
                for _, row in urgent.iterrows()
            ])
            messagebox.showwarning(
                "Deadline Alert",
                f"These topics need immediate attention:\n\n{topics}"
            )

    def check_deadline_alerts_on_start(self):
        self.root.after(700, self.check_deadline_alerts)

    # ------------------ Charts ------------------
    def show_status_pie_chart(self):
        df = read_data()

        if df.empty:
            messagebox.showinfo("No Data", "No data available to generate chart.")
            return

        self.clear_chart_frame()
        status_counts = df["Status"].value_counts()

        fig, ax = plt.subplots(figsize=(6, 5))
        ax.pie(
            status_counts.values,
            labels=status_counts.index,
            autopct="%1.1f%%",
            startangle=90
        )
        ax.set_title("Topic Status Distribution")

        canvas = FigureCanvasTkAgg(fig, master=self.chart_display_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def show_priority_pie_chart(self):
        df = read_data()

        if df.empty:
            messagebox.showinfo("No Data", "No data available to generate chart.")
            return

        self.clear_chart_frame()
        priority_counts = df["Priority"].value_counts()

        fig, ax = plt.subplots(figsize=(6, 5))
        ax.pie(
            priority_counts.values,
            labels=priority_counts.index,
            autopct="%1.1f%%",
            startangle=90
        )
        ax.set_title("Priority Distribution")

        canvas = FigureCanvasTkAgg(fig, master=self.chart_display_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def show_subject_bar_chart(self):
        df = read_data()

        if df.empty:
            messagebox.showinfo("No Data", "No data available to generate chart.")
            return

        self.clear_chart_frame()
        subject_counts = df["Subject"].value_counts()

        fig, ax = plt.subplots(figsize=(8, 5))
        ax.bar(subject_counts.index, subject_counts.values)
        ax.set_title("Subject-wise Topic Count")
        ax.set_xlabel("Subject")
        ax.set_ylabel("Number of Topics")
        plt.xticks(rotation=30)

        canvas = FigureCanvasTkAgg(fig, master=self.chart_display_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # ------------------ AI Assistant ------------------
    def select_ai_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Study Material",
            filetypes=[
                ("All Supported Files", "*.pdf *.txt *.docx *.md *.csv"),
                ("PDF Files", "*.pdf"),
                ("Text Files", "*.txt"),
                ("Word Files", "*.docx"),
                ("Markdown Files", "*.md"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ]
        )

        if file_path:
            self.ai_file_path = file_path
            self.ai_file_label.config(text=f"Selected File: {os.path.basename(file_path)}")

    def set_ai_status(self, text):
        self.ai_status_label.config(text=f"AI Status: {text}")

    def append_ai_output(self, text):
        self.ai_output.delete("1.0", tk.END)
        self.ai_output.insert(tk.END, text)

    def build_ai_prompt(self, mode, subject, extra_topics, custom_question):
        base = (
            f"You are an exam-focused study assistant. The user's weak subject is '{subject}'. "
            f"Read the uploaded material and any pasted topics carefully. "
            f"Generate concise, accurate, easy-to-revise content in simple student-friendly language. "
            f"Do not write long essays. Use headings and bullet points."
        )

        if extra_topics:
            base += f"\n\nImportant topics or instructions from the user:\n{extra_topics}"

        if mode == "short_notes":
            base += (
                "\n\nTask: Create short notes for this weak subject."
                "\nFormat strictly as:"
                "\n1. Very short overview"
                "\n2. Key concepts"
                "\n3. Important formulas / definitions"
                "\n4. Exam-important points"
                "\n5. 5 quick revision bullets"
            )
        elif mode == "important_points":
            base += (
                "\n\nTask: Extract only the most important exam points."
                "\nFocus on definitions, formulas, repeated concepts, and must-study subtopics."
                "\nKeep it crisp."
            )
        elif mode == "quick_revision":
            base += (
                "\n\nTask: Create a last-minute quick revision sheet."
                "\nInclude only what a student should revise one day before the exam."
            )
        elif mode == "custom":
            base += f"\n\nCustom user question:\n{custom_question}"

        return base

    def generate_ai_notes(self, mode):
        if self.ai_busy:
            messagebox.showinfo("Please Wait", "AI is already generating a response.")
            return

        subject = self.ai_subject_combo.get().strip()
        extra_topics = self.ai_topics_text.get("1.0", tk.END).strip()
        custom_question = self.ai_question_entry.get().strip()

        if not subject:
            messagebox.showerror("Error", "Please select a weak subject.")
            return

        if mode == "custom" and not custom_question:
            messagebox.showerror("Error", "Please enter a custom question.")
            return

        if not self.ai_file_path and not extra_topics:
            messagebox.showerror(
                "Error",
                "Please upload a study file or paste important topics/instructions."
            )
            return

        self.ai_busy = True
        self.set_ai_status("Generating...")
        self.append_ai_output("Generating response... Please wait.\n")

        worker = threading.Thread(
            target=self._run_ai_request,
            args=(mode, subject, extra_topics, custom_question),
            daemon=True
        )
        worker.start()

    def _run_ai_request(self, mode, subject, extra_topics, custom_question):
        try:
            try:
                from openai import OpenAI
            except ImportError:
                raise Exception("OpenAI package is not installed. Run: pip install openai")

            api_key = os.getenv("OPENAI_API_KEY", "").strip()
            if not api_key:
                raise Exception(
                    "OPENAI_API_KEY is not set. Set it in your system environment and restart the app."
                )

            client = OpenAI()
            prompt = self.build_ai_prompt(mode, subject, extra_topics, custom_question)

            content_items = [{"type": "input_text", "text": prompt}]

            if self.ai_file_path:
                with open(self.ai_file_path, "rb") as f:
                    uploaded_file = client.files.create(file=f, purpose="user_data")
                content_items.append({"type": "input_file", "file_id": uploaded_file.id})

            response = client.responses.create(
                model="gpt-5.4",
                input=[
                    {
                        "role": "user",
                        "content": content_items
                    }
                ]
            )

            output_text = getattr(response, "output_text", "").strip()
            if not output_text:
                output_text = "No response text returned."

            self.root.after(0, lambda: self._finish_ai_success(output_text))

        except Exception as e:
            self.root.after(0, lambda: self._finish_ai_error(str(e)))

    def _finish_ai_success(self, text):
        self.ai_busy = False
        self.set_ai_status("Done")
        self.append_ai_output(text)

    def _finish_ai_error(self, error_text):
        self.ai_busy = False
        self.set_ai_status("Error")
        self.append_ai_output(f"AI Error:\n{error_text}")

    # ------------------ Clear ------------------
    def clear_fields(self):
        self.subject_entry.delete(0, tk.END)
        self.topic_entry.delete(0, tk.END)
        self.status_combo.set("")
        self.date_entry.delete(0, tk.END)
        self.priority_combo.set("")
        self.notes_entry.delete(0, tk.END)
        self.tree.selection_remove(self.tree.selection())


# ------------------ Run ------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = SmartExamPlanner(root)
    root.mainloop()
