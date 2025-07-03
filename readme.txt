# 🐒 Monkey Wrench

**Monkey Wrench** is a fun and quirky Python GUI tool that duplicates and mutates CSV files — perfect for testing deduplication tools. It features a cartoon UI, monkey sound effects, and even a little animation of a monkey throwing a wrench!

![screenshot](preview.png) ./image.png

---

## 🛠 Features

- 🗂 Load any `.csv` file
- 🔁 Duplicate random rows
- ➕ Add extra columns with generated values
- 🎨 Cartoon-themed UI with custom tile background
- 🔊 Monkey sound and animation when activated
- 💾 Save mutated CSV to any location

---

## 🧰 Requirements

- Python 3.10–3.13

## ▶️ To Run
pyinstaller --noconfirm --onefile --windowed --icon=monkey.ico `
--add-data "monkey.wav;." `
--add-data "monkey_logo.png;." `
--add-data "grid_texture.png;." `
--add-data "monkey.ico;." `
monkey_wrench_gui.py
