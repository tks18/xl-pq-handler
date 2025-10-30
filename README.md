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

## 💅 The Vibe Check: Before vs. After

_(The PQ IDE You Didn't Know You Needed ✨)_

| **Before xl-pq-handler 🫠**                   | **After xl-pq-handler 😎**                              |
| :-------------------------------------------- | :------------------------------------------------------ |
| Endless copy-pasting M-code                   | One-click insert into **any** open Excel workbook       |
| Forgetting `fn_Helper_v3` needs `fn_Util_v1`  | **Dependency graph** shows you the whole family tree 🌳 |
| decentralized file organization               | Auto-organized folders based on `category`              |
| Editing metadata = Manual YAML torture        | Right-click -\> **Edit Metadata** -\> Save -\> Done ✅  |
| "Which file uses that API?" -\> 🤷‍♂️            | **Data Sources tab** spills the tea ☕                  |
| Updating one function in 5 workbooks manually | Edit once -\> Refresh UI -\> Insert where needed        |

This is that **main character energy** for your data workflow.

---

## ✨ Features That Absolutely Slap

This ain't your grandpa's script library. We got a whole ecosystem:

### 🖥️ **The UI App (Your New Dashboard)**

- Launch a **dedicated desktop app** straight from your terminal. No more sad script outputs.
- Visually browse, search, and filter your _entire_ `.pq` library like a pro.
- It's got that dark mode aesthetic. You know the vibes. ✨

### 📥 **Smart Extract ("Yoink\! Button")**

- **From File:** Point it at _any_ `.xlsx` / `.xlsm` / `.xlsb` and instantly rip out all the Power Queries.
- **From Open Workbook:** Got 5 Excels open? No stress. A **dropdown lists all open workbooks**. Pick one, hit extract. Easy.
- **Preview Before Saving:** See the code (with syntax highlighting\!), parameters, and data sources _before_ you commit to saving the `.pq` file. No more blind extraction\!

### 🪄 **Dependency-Aware Insert ("Yeet Button")**

- Select a query (e.g., `FinalReport`). The app automatically knows it needs `GetSalesData` and `fn_FormatDate`.
- It yeets **all required queries** into Excel _in the correct order_. 🤯
- **Target Practice:** Don't just spray into the active workbook. Use the **dropdown to select _exactly_ which open workbook** gets the queries. Precision\!

### ✏️ **Edit Metadata + Auto-Sync ("The Organizer")**

- Right-click a query -\> "Edit Metadata."
- Change the `name`, `tags`, `dependencies`, `description`, `version`.
- **The Magic ✨:** Change the `category` from `Staging` to `Production`? The app **automatically moves the `.pq` file** to the `Production/` folder. Chef's kiss\! 🤌

### 💅 **Syntax Highlighting ("Make it Pretty")**

- See your M-code in the **Preview tabs** (Library, Edit, Extract) with **VS Code-style syntax highlighting**. Keywords, functions, strings, comments – all colored up. ✨

### 🧐 **Code Intelligence ("The Brain")**

- **Parameter Peek:** Select a function query, and the **"Parameters" tab** shows its inputs, types (`any`, `text`, etc.), and if they're `optional`.
- **Data Source Detective:** The **"Data Sources" tab** scans the code and lists out _all_ the external connections (`Sql.Database`, `Web.Contents`, `File.Contents`, etc.) and whether the source is a literal string or an input parameter. Big for security audits\! 🕵️‍♀️
- **Dependency Deets:**
  - **Auto-Detect:** Click the button in the Edit dialog to automatically scan the code and suggest the `dependencies`. Saves _so_ much typing.
  - **Visual Graph:** The **"Graph" tab** shows a slick tree view of a query's entire dependency chain. No more surprises. 🌳

### 💻 **External Editor Escape Hatch ("Send It")**

- Need to tweak the _actual_ M-code logic?
- Right-click -\> "Open in Editor."
- Instantly opens the `.pq` file in **VS Code** (if it's in your PATH) or falls back to Notepad. Edit, save, hit refresh in the UI. Seamless.

### 🤖 **Python Backend (`PQManager`)**

- All the power, none of the clicks. Import `PQManager` into your own Python automation scripts.
- Headless extraction, insertion, index building – you name it. Perfect for CI/CD or scheduled tasks.

---

## 📦 Get it Already (Installation)

```bash
pip install xl-pq-handler
```

_(This single command grabs everything you need: `customtkinter`, `xlwings`, `pydantic`, `pyyaml`, `pandas`, `filelock` – the whole squad.)_

---

## 🚀 How to Vibe

### 1\. The Main Way (The UI) 💅

This is the main event. Open your terminal:

```bash

# Better launch - point it at your actual PQ repo folder
python -m xl_pq_handler "D:\Path\To\Your\PowerQuery_Repo"

# Or even better way
pqmagic "D:\Path\To\Your\PowerQuery_Repo"
```

Now just... use the app. Click around. It's built different. 😎

Then just... click buttons. It's that easy.

### 2\. 🤓 Script Kiddie Corner (Python Usage)

For your `main.py` automation scripts, use the `PQManager`.

```python
from xl_pq_handler import PQManager

# Point manager at your repo
manager = PQManager(r"D:\Path\To\Your\PowerQuery_Repo")

# Rebuild index (always a good move)
manager.build_index()

# ---- EXAMPLE: INSERT INTO SPECIFIC WORKBOOK ----
target_workbook = "Monthly_Report_WIP.xlsx" # Must be open!
queries_needed = ["Calculate_KPIs", "Generate_Summary"]

try:
    manager.insert_into_excel(
        names=queries_needed,
        workbook_name=target_workbook # <-- Target acquired 🎯
    )
    print(f"🚀 Sent queries to {target_workbook}. Mission accomplished.")
except Exception as e:
    print(f"😭 Insert failed: {e}")

# ---- EXAMPLE: EXTRACT FROM FILE ----
source_file = r"C:\Downloads\NewDataSource.xlsx"
try:
    manager.extract_from_excel(category="Downloaded", file_path=source_file)
    print(f"✅ Successfully yoinked queries from {source_file}!")
except Exception as e:
    print(f"💀 Extraction failed: {e}")
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
name: Clean_RawSales          # The query's name in Excel/PBI
category: Staging             # Matches the folder name (keep it sync'd!)
tags: [cleaning, sales, raw]  # Searchable tags
dependencies:                 # List other queries *this one* calls
  - fn_FormatDate
description: Cleans and transforms the raw monthly sales data dump. # What it does
version: 2.1                  # Your version number
---

let                           # Start of your actual M-code
    Source = Csv.Document(File.Contents("path/to/raw.csv"), ...),
    #"Formatted Date" = fn_FormatDate(Source, "OrderDate")
in
    #"Formatted Date"
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
