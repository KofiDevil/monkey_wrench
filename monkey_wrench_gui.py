# Author: Levi Meltabarger
# Description: Messes with CSV files.
# Core
import csv
import random
import time
import os
import sys
from typing import Optional

# Third-Party
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pygame

def resource_path(relative_path: str) -> str:
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path) # type: ignore
    return os.path.join(os.path.abspath("."), relative_path)


class MonkeyWrenchApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root: tk.Tk = root
        self.root.title("Monkey Wrench 🐒🔨")
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
            self.canvas.image = photo # type: ignore
        except Exception:
            pass

        self.add_label("Select a CSV file:", 200)
        self.browse_button = self.add_button("Browse CSV", self.browse_file, 250)

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
        label = tk.Label(self.root, text=text, fg="#f0c674", bg="#1a1a1a", bd=0, highlightthickness=0, font=("Segoe UI", 14))
        self.canvas.create_window(400, y, window=label)

    def add_entry(self, default: str, y: int) -> tk.Entry:
        entry = tk.Entry(self.root, fg="white", insertbackground="white", bg="#1a1a1a", bd=0, highlightthickness=0, font=("Segoe UI", 12))
        entry.insert(0, default)
        self.canvas.create_window(400, y, window=entry, width=280)
        return entry

    def add_button(self, text: str, command, y: int) -> tk.Button:
        button = tk.Button(self.root, text=text, command=command, fg="#f0c674", bg="#1a1a1a", activebackground="#333", activeforeground="white", bd=0, highlightthickness=0, font=("Segoe UI", 12, "bold"))
        self.canvas.create_window(400, y, window=button, width=180)
        return button

    def browse_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[["CSV files", "*.csv"]]) # type: ignore
        if path:
            self.file_path = path
            self.status_text.set(f"Loaded: {os.path.basename(path)}")

    def start_process(self) -> None:
        try:
            count: int = int(self.dup_entry.get())
            cols: int = int(self.col_entry.get())
            if count < 1 or cols < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Enter valid numbers for duplicates and columns.")
            return

        self.run_button.config(state=tk.DISABLED)
        self.status_text.set("Monkey throwing wrench...")
        self.root.after(100, lambda: self.process_csv(count, cols))

    def monkey_animation(self) -> None:
        frames = [
            r"  (o_o)    ",
            r"  (o_o)    ==>> ",
            r"  (o_o)    ===>>> ",
            r"  (o_o)   ===>>>> 🔨 ",
            r"  (o_o)   ===>>>>      🔨 ",
            r"  (o_o)   ===>>>>         🔨",
        ]
        for frame in frames:
            self.animation_label.config(text=frame)
            self.root.update()
            time.sleep(0.2)

    def play_monkey_sound(self) -> None:
        try:
            pygame.mixer.init()
            pygame.mixer.music.load(resource_path("monkey.wav"))
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy():
                time.sleep(0.1)
        except Exception as e:
            print("Sound error:", e)

    def process_csv(self, count: int, col_count: int) -> None:
        self.monkey_animation()
        self.play_monkey_sound()

        try:
            with open(self.file_path, newline='', encoding='utf-8') as f: # type: ignore
                reader = list(csv.reader(f))
                headers = reader[0]
                rows = reader[1:]
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data:\n{e}")
            self.status_text.set("❌ Error loading data.")
            self.run_button.config(state=tk.NORMAL)
            return

        for i in range(col_count):
            headers.append(f"ExtraCol{i+1}")

        duplicates = [random.choice(rows) for _ in range(count)]
        all_rows = rows + duplicates
        random.shuffle(all_rows)
        output_data = [headers] + all_rows

        output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[["CSV files", "*.csv"]], initialfile="output.csv") # type: ignore
        if output_path:
            with open(output_path, "w", newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerows(output_data)
            self.status_text.set(f"✔ Saved to {os.path.basename(output_path)}")
        else:
            self.status_text.set("❌ Save cancelled.")

        self.run_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u"monkey.wrench.tool")
    pygame.init()
    root = tk.Tk()
    root.iconbitmap(resource_path("monkey.ico"))
    app = MonkeyWrenchApp(root)
    root.mainloop()
