# 🌈 xl-pq-handler

> 🧩 A Pythonic Power Query (.pq) File Manager for Excel & Power BI Automation

[![PyPI Version](https://img.shields.io/pypi/v/xl-pq-handler.svg?color=4CAF50&logo=python&logoColor=white)](https://pypi.org/project/xl-pq-handler/)
[![Python Versions](https://img.shields.io/pypi/pyversions/xl-pq-handler.svg?color=blue)](https://pypi.org/project/xl-pq-handler/)
[![License](https://img.shields.io/github/license/tks18/xl-pq-handler.svg?color=orange)](LICENSE)

---

### 🧠 What is `xl-pq-handler`?

`xl-pq-handler` is a **Python UI App + library** built for developers, data analysts, and automation engineers who work with **Power Query (.pq)** files in Excel or Power BI.

It lets you:

- 🔍 Parse, search, and index `.pq` scripts
- 📋 Copy Power Query code to clipboard
- 🪄 Insert queries directly into Excel workbooks
- 🧾 Maintain YAML-based metadata (name, category, tags, description, version)
- 🔁 Export, validate, and refresh PQ indexes
- ⚡ Batch-insert queries for rapid Excel automation

All from Python. No manual clicks. No clutter. 🚀

---

> Stop the cap. Managing Power Query `.pq` files is low-key a nightmare.
>
> **This tool is the ultimate glow-up for your M-code.** 💅
>
> It's not just a library; it's your new **Power Query IDE**.

---

### 🫠 The Struggle is Real (Before)

- Copy-pasting M-code between 8 different Excel files.
- Forgetting which queries depend on `fn_Helper_v2_final`.
- Your `Downloads/PQ` folder looks like a bomb went off.
- Manually updating a query in 3 different workbooks. 💀

### 😎 The Vibe (After)

- Launch a sleek UI that shows your _entire_ `.pq` library.
- Right-click a query -\> "Insert" -\> pick your open Excel file from a **dropdown**. Done.
- Right-click -\> "Edit Metadata" -\> change `category` -\> the app **auto-moves the file** to the new folder. 🤯
- Click the "Graph" tab to see all the dependencies. No more guessing.

This is that **main character energy** for your data workflow.

---

## ✨ The Features _Actually_ Slap

This tool is a whole mood. It's a UI app _and_ a Python library.

### 🖥️ **The UI App (Your New BFF)**

Forget scripts. Just launch the app from your terminal. This is your mission control.
`python -m xl_pq_handler "path/to/your/repo"`

### 📥 **Smart Extract (The "Yoink")**

- **Yank from File:** Pick any `.xlsx` and rip all its queries.
- **Yank from Open WB:** Don't even know where the file is? Just pick from a **dropdown of all your open workbooks**. Bet.

### 🪄 **Dependency-Aware Insert (The "Yeet")**

- Select a query. This tool auto-finds **all its dependencies**.
- It inserts them _in the correct order_.
- **Pick your target:** Don't just spray and pray into your active workbook. A **dropdown shows all open workbooks** so you can snipe the exact one you want.

### ✏️ **Edit & Sync (The "Glow-Up")**

- Right-click any query to **edit its metadata** (name, category, tags, deps).
- **The best part:** You change the `category` from "Staging" to "Production"? The app **auto-moves the `.pq` file** from the `Staging/` folder to the `Production/` folder. IYKYK. 🤯

### 🔗 **See the Receipts (Dependency Graph)**

- Tired of guessing what a query needs?
- Click a query -\> click the **"Graph" tab**.
- See a beautiful tree of all its dependencies, right there. No cap.

### 💻 **"I'm Out" (External Editor)**

- Need to edit the _actual_ M-code?
- Right-click -\> "Open in Editor."
- This instantly opens the file in **VS Code** (or Notepad, if you're basic) for you to edit. Save, go back to the app, hit refresh. ✨

### 🧠 **The Brain (For Scripting)**

- Under the hood is the `PQManager`, a sick Python library.
- Use it in your own automation scripts for all the features above, but headless. 🤖

---

## 📦 Get it Already (Installation)

```bash
pip install xl-pq-handler
```

_(^ above installs all of the dependencies - `customtkinter`, `xlwings`, `pydantic`, `pyyaml`, `pandas`, & `filelock` too\!)_

---

## 🚀 How to Vibe

### 1\. The Main Way (The UI) 💅

This is what you want. Open your terminal and run this.

```bash
# Launch the UI
# Point it at the folder where you store your .pq files
python -m xl_pq_handler "D:\My-Power-Query-Repo"
```

_(If you set up the script, you can just do `pq-magic "..."`)_

Then just... click buttons. It's that easy.

### 2\. The Automation Way (Python Script) 🤓

For your `main.py` automation scripts, use the `PQManager`.

```python
from xl_pq_handler import PQManager

# Point it at your repo
manager = PQManager(r"D:\My-Power-Query-Repo")

# Rebuild index (good practice)
manager.build_index()

# ---- SCRIPTING EXAMPLE ----
# Insert "FinalReport" + all its dependencies
# into a *specific* open workbook named "Dashboard.xlsm"

queries_to_add = ["FinalReport"]

try:
    manager.insert_into_excel(
        names=queries_to_add,
        workbook_name="Dashboard.xlsm"  # <-- So clean!
    )
    print("🚀 Queries sent! Go be a hero.")
except Exception as e:
    print(f"😬 Bruh, it failed: {e}")
```

---

## 📁 The Drip (File Structure)

This is how you organize your repo. The app does the rest.

```
My-Power-Query-Repo/
│
├── index.json          <-- The app makes this. Don't touch.
│
├── API/                <-- "API" Category
│   ├── Get_API_Data.pq
│   └── fn_Get_Credentials.pq
│
├── Helpers/            <-- "Helpers" Category
│   ├── fn_Format_Date.pq
│   └── fn_Safe_Divide.pq
│
└── Reports/            <-- "Reports" Category
    └── Final_Sales_Report.pq
```

Each `.pq` file is just M-code with a **YAML "frontmatter"** block at the top. This is the metadata.

```yaml
---
name: Final_Sales_Report
category: Reports
tags: [sales, final, public]
dependencies:
  - Get_API_Data
  - fn_Format_Date
description: The main query for the monthly sales dashboard.
version: 1.5
---

(let
    Source = Get_API_Data(),
    #"Formatted Date" = fn_Format_Date(Source, "DateColumn")
in
    #"Formatted Date")

```

---

## 📜 License

This project is licensed under the **GNU-GPL 3.0 License**. Go wild.

---

## 💚 Credits

Made by **Sudharshan TK** (tks18)

If this tool just saved your workflow, give it a ⭐ **Star on [GitHub](https://github.com/tks18/xl-pq-handler)\!**

---

> ⚡ _“Automate the boring Power Query stuff — one `.pq` at a time.”_

---
