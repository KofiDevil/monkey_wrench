#
# Author: Levi Meltabarger
# Description: Duplicates and adds random columns to a google sheet.
#

# Native
import csv
import random
import time
import os
import sys
from typing import Optional, List

import tkinter as tk
from tkinter import filedialog, messagebox

# Third-Party
from PIL import Image, ImageTk
import pygame

import gspread
from google.oauth2.service_account import Credentials
from urllib.parse import urlparse
import re

# Extra columns used for random insertion
EXTRA_COLUMNS_POOL = [
    "Source", "Date", "Obsolescence Date", "Physical Address Number", "Physical Pre Direction", 
    "Physical Address Name", "Physical Address Suffix", "First Name", "Last Name", 
    "Mailing Address Number", "Mailing Address Name", "Mailing Address Suffix", 
    "Mailing Post Direction", "Number of PCs", "Stock Exchange", "2021 % Sales Growth", 
    "2020 % Sales Growth", "2019 % Sales Growth", "2022 Employees", "2021 Employees", 
    "2020 Employees", "2019 Employees", "2021 % Employee Growth", "2020 % Employee Growth", 
    "2019 % Employee Growth", "Executive Gender 3", "Executive Gender 2", "Executive Department 3", 
    "Executive Gender 4", "Executive Gender 5", "Executive Gender 6", "Executive Gender 7", 
    "Executive Gender 8", "Executive First Name 9", "Executive Last Name 9", 
    "Executive Title 9", "Executive Gender 9", "Executive Department 9", "Executive First Name 10", 
    "Executive Last Name 10", "Executive Title 10", "Executive Gender 10", 
    "Executive Department 10", "Executive First Name 11", "Executive Last Name 11", 
    "Executive Title 11", "Executive Gender 11", "Executive Department 11", 
    "Executive First Name 12", "Executive Last Name 12", "Executive Title 12", 
    "Executive Gender 12", "Executive Department 12", "Executive First Name 13", 
    "Executive Last Name 13", "Executive Title 13", "Executive Gender 13", 
    "Executive Department 13", "Executive First Name 14", "Executive Last Name 14", 
    "Executive Title 14", "Executive Gender 14", "Executive Department 14", 
    "Executive First Name 15", "Executive Last Name 15", "Executive Title 15", 
    "Executive Gender 15", "Executive Department 15", "Executive First Name 16", 
    "Executive Last Name 16", "Executive Title 16", "Executive Gender 16", 
    "Executive Department 16", "Executive First Name 17", "Executive Last Name 17", 
    "Executive Title 17", "Executive Gender 17", "Executive Department 17", 
    "Executive First Name 18", "Executive Last Name 18", "Executive Title 18", 
    "Executive Gender 18", "Executive Department 18", "Executive First Name 19", 
    "Executive Last Name 19", "Executive Title 19", "Executive Gender 19", 
    "Executive Department 19", "Executive First Name 20", "Executive Last Name 20", 
    "Executive Title 20", "Executive Gender 20", "Executive Department 20", 
    "Est. Accounting Annual Expense", "Est. Advertising Annual Expense", 
    "Est. Business Insurance Annual Expense", "Est. Legal Annual Expense", 
    "Est. Office Equipment Annual Expense", "Est. Rent Annual Expense", 
    "Est. Technology Annual Expense", "Est. Utilities Annual Expense", "AtoZ ID", 
    "Home Based Business"
]

# Get absolute resource path
def resource_path(relative_path: str) -> str:
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)  # type: ignore
    return os.path.join(os.path.abspath("."), relative_path)


class MonkeyWrenchApp:
    def __init__(self, root: tk.Tk) -> None:
        """Initialize the GUI window and load assets"""
        self.root = root
        self.root.title("Monkey Wrench 🐒🔧")
        self.root.geometry("800x900")
        self.root.resizable(False, False)

        self.canvas = tk.Canvas(root, width=800, height=900, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        try:
            bg_img = Image.open(resource_path("grid_texture.png")).resize((800, 900))
            self.bg_photo = ImageTk.PhotoImage(bg_img)
            self.canvas.create_image(0, 0, anchor="nw", image=self.bg_photo)
        except Exception as e:
            print("Could not load background texture:", e)
            self.canvas.configure(bg="#1a1a1a")

        self.file_path: Optional[str] = None

        try:
            img = Image.open(resource_path("monkey_logo.png")).resize((160, 160))
            photo = ImageTk.PhotoImage(img)
            self.logo = self.canvas.create_image(400, 100, image=photo)
            self.canvas.image = photo  # type: ignore
        except Exception:
            pass

        self.add_label("Google Sheet URL:", 200)
        self.sheet_url_entry = self.add_entry("", 250)

        self.add_label("Number of duplicates:", 310)
        self.dup_entry = self.add_entry("5", 360)

        self.add_label("Number of columns to add:", 420)
        self.col_entry = self.add_entry("2", 470)

        self.run_button = self.add_button("Throw Wrench!", self.start_process, 530)

        self.status_text = tk.StringVar()
        self.status_label = tk.Label(self.root, textvariable=self.status_text, fg="#f0c674", bg="#1a1a1a", bd=0, highlightthickness=0)
        self.status_window = self.canvas.create_window(400, 590, window=self.status_label)

        self.animation_label = tk.Label(self.root, text="", font=("Courier", 22), fg="#f0c674", bg="#1a1a1a", bd=0, highlightthickness=0)
        self.animation_window = self.canvas.create_window(400, 640, window=self.animation_label)

    def add_label(self, text: str, y: int) -> None:
        """Add a styled label to the GUI"""
        label = tk.Label(self.root, text=text, fg="#f0c674", bg="#1a1a1a", bd=0, highlightthickness=0, font=("Segoe UI", 14))
        self.canvas.create_window(400, y, window=label)

    def add_entry(self, default: str, y: int) -> tk.Entry:
        """Add a styled entry box to the GUI"""
        entry = tk.Entry(self.root, fg="white", insertbackground="white", bg="#1a1a1a", bd=0, highlightthickness=0, font=("Segoe UI", 12))
        entry.insert(0, default)
        self.canvas.create_window(400, y, window=entry, width=280)
        return entry

    def add_button(self, text: str, command, y: int) -> tk.Button:
        """Add a button to the GUI"""
        button = tk.Button(self.root, text=text, command=command, fg="#f0c674", bg="#1a1a1a", activebackground="#333", activeforeground="white", bd=0, highlightthickness=0, font=("Segoe UI", 12, "bold"))
        self.canvas.create_window(400, y, window=button, width=180)
        return button

    def start_process(self) -> None:
        """Parse form input and start processing the sheet"""
        try:
            count: int = int(self.dup_entry.get())
            cols: int = int(self.col_entry.get())
            if count < 1 or cols < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Enter valid numbers for duplicates and columns.")
            return

        sheet_url = self.sheet_url_entry.get().strip()
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match:
            messagebox.showerror("Error", "Invalid Google Sheets URL.")
            return

        sheet_id = match.group(1)

        self.run_button.config(state=tk.DISABLED)
        self.status_text.set("Monkey throwing wrench... 🐒")
        self.root.after(100, lambda: self.process_csv(count, cols, use_google=True, sheet_id=sheet_id))

    def monkey_animation(self) -> None:
        """Play simple ASCII animation"""
        frames = [
            r"  (o_o)    ",
            r"  (o_o)    ==>> ",
            r"  (o_o)    ===>>> ",
            r"  (o_o)   ===>>>> 🔧",
            r"  (o_o)   ===>>>>     🔧",
            r"  (o_o)   ===>>>>        🔧",
        ]
        for frame in frames:
            self.animation_label.config(text=frame)
            self.root.update()
            time.sleep(0.2)

    def play_monkey_sound(self) -> None:
        """Play a monkey sound clip using pygame"""
        try:
            pygame.mixer.init()
            pygame.mixer.music.load(resource_path("monkey.wav"))
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy():
                time.sleep(0.1)
        except Exception as e:
            print("Sound error:", e)

    def load_google_sheet_by_id(self, sheet_id: str, worksheet_title: str = "Sheet1") -> List[List[str]]:
        """Load worksheet data from Google Sheets by sheet ID"""
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(resource_path("service_account.json"), scopes=SCOPES)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id)
        worksheet = sheet.worksheet(worksheet_title)
        return worksheet.get_all_values()

    def save_to_existing_sheet_by_id(self, sheet_id: str, worksheet_title: str, data: List[List[str]]) -> None:
        """Overwrite the contents of an existing worksheet"""
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(resource_path("service_account.json"), scopes=SCOPES)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id)
        worksheet = sheet.worksheet(worksheet_title)

        worksheet.resize(rows=len(data), cols=len(data[0]))
        worksheet.update(values=data, range_name="A1")

    def process_csv(self, count: int, col_count: int, use_google: bool = False, sheet_id: Optional[str] = None) -> None:
        """Main logic to load, mutate, and save spreadsheet"""
        self.monkey_animation()
        self.play_monkey_sound()

        try:
            if use_google and sheet_id:
                reader = self.load_google_sheet_by_id(sheet_id)
            else:
                with open(self.file_path, newline='', encoding='utf-8') as f:
                    reader = list(csv.reader(f))
            headers = reader[0]
            rows = reader[1:]
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data:\n{e}")
            self.status_text.set("❌ Error loading data.")
            self.run_button.config(state=tk.NORMAL)
            return

        for i in range(col_count):
            new_col_name = random.choice(EXTRA_COLUMNS_POOL)
            tries = 0
            while new_col_name in headers and tries < 10:
                new_col_name = random.choice(EXTRA_COLUMNS_POOL)
                tries += 1
            if new_col_name in headers:
                new_col_name += f"_{random.randint(1000,9999)}"
            insert_index = random.randint(0, len(headers))
            headers.insert(insert_index, new_col_name)
            for row in rows:
                row.insert(insert_index, "")

        duplicates = [random.choice(rows) for _ in range(count)]
        all_rows = rows + duplicates
        random.shuffle(all_rows)
        output_data = [headers] + all_rows

        if use_google and sheet_id:
            self.save_to_existing_sheet_by_id(sheet_id, "Sheet1", output_data)
            self.status_text.set("✔ Sheet updated!")
        else:
            output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[["CSV files", "*.csv"]], initialfile="output.csv")
            if output_path:
                with open(output_path, "w", newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerows(output_data)
                self.status_text.set(f"✔ Saved to {os.path.basename(output_path)}")
            else:
                self.status_text.set("❌ Save cancelled.")

        self.run_button.config(state=tk.NORMAL)


if __name__ == "__main__":
    if sys.platform == "win32":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u"monkey.wrench.tool")
    pygame.init()
    root = tk.Tk()
    if sys.platform.startswith("win"):
        root.iconbitmap(resource_path("monkey.ico"))
    else:
        root.iconbitmap("@" + resource_path("monkey-2.xbm"))
    app = MonkeyWrenchApp(root)
    root.mainloop()