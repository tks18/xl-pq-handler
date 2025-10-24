# 🌈 xl-pq-handler  
> 🧩 A Pythonic Power Query (.pq) File Manager for Excel & Power BI Automation

[![PyPI Version](https://img.shields.io/pypi/v/xl-pq-handler.svg?color=4CAF50&logo=python&logoColor=white)](https://pypi.org/project/xl-pq-handler/)
[![Python Versions](https://img.shields.io/pypi/pyversions/xl-pq-handler.svg?color=blue)](https://pypi.org/project/xl-pq-handler/)
[![License](https://img.shields.io/github/license/tks18/xl-pq-handler.svg?color=orange)](LICENSE)

---

### 🧠 What is `xl-pq-handler`?

`xl-pq-handler` is a **single-class Python library** built for developers, data analysts, and automation engineers who work with **Power Query (.pq)** files in Excel or Power BI.  

It lets you:
- 🔍 Parse, search, and index `.pq` scripts  
- 📋 Copy Power Query code to clipboard  
- 🪄 Insert queries directly into Excel workbooks  
- 🧾 Maintain YAML-based metadata (name, category, tags, description, version)  
- 🔁 Export, validate, and refresh PQ indexes  
- ⚡ Batch-insert queries for rapid Excel automation  

All from Python. No manual clicks. No clutter. 🚀  

---

## 📦 Installation

```bash
pip install xl-pq-handler
````

**Dependencies**
This package uses:

* `xlwings` – for Excel integration
* `pyperclip` – for clipboard operations
* `pandas`, `yaml` – for data wrangling
* `csv`, `json`, `os`, `logging` – built-in utilities

---

## 🧩 Quick Start

```python
from xl_pq_handler import XLPowerQueryHandler

# Initialize handler (your PQ root folder)
handler = XLPowerQueryHandler(r"C:\MyPQFiles")

# 🔧 Build or refresh index
handler.build_index()

# 📚 Read all indexed queries
entries = handler.read_index()

# 🔍 Search for queries
matches = handler.search_pq("sales")

# 🧠 Get full PQ entry (including code)
pq = handler.get_pq_by_name("TransformSalesData")

# 🪄 Insert PQ into active Excel
handler.insert_pq_into_active_excel("TransformSalesData")

# 💾 Export the index to JSON
handler.export_index_to_json("pq_index.json")
```

---

## 📁 Folder Structure Example

```
MyPQFiles/
│
├── Sales_Transform.pq
├── Merge_Customers.pq
├── Region_Filter.pq
└── index.csv  ← Generated automatically
```

Each `.pq` file may optionally contain **YAML frontmatter**:

```yaml
---
name: TransformSalesData
category: Sales
tags: [cleaning, merge, sales]
description: Cleans and merges monthly sales data
version: 1.2
---
let
    Source = Excel.CurrentWorkbook(){[Name="SalesData"]}[Content],
    Clean = Table.TransformColumnTypes(Source,{{"Amount", type number}})
in
    Clean
```

---

## ⚙️ Key Features

| Feature                      | Description                                              |
| ---------------------------- | -------------------------------------------------------- |
| 🧾 **YAML Metadata Parsing** | Reads and updates `.pq` frontmatter (name, tags, etc.)   |
| 📊 **CSV Index Builder**     | Auto-builds and maintains a full PQ index (`index.csv`)  |
| 🔍 **Search Engine**         | Keyword search across name, description, and tags        |
| 🧠 **Metadata Updater**      | Modify YAML frontmatter fields dynamically               |
| ✂️ **Clipboard Copier**      | Copy Power Query code directly to clipboard              |
| 📤 **Excel Integration**     | Insert queries into Excel (active or specified workbook) |
| ⚡ **Batch Insert**           | Add multiple queries to Excel in one go                  |
| 🧱 **DataFrame Export**      | Get PQ index as a Pandas DataFrame                       |
| ✅ **Validation Tools**       | Detect missing or invalid PQ paths                       |

---

## 💻 Excel Integration Demo

```python
# Insert a PQ into specific workbook
handler.insert_pq_into_excel(
    file_path=r"C:\Reports\SalesReport.xlsx",
    name="TransformSalesData"
)

# Insert multiple PQs at once
handler.insert_pqs_batch(
    names=["TransformSalesData", "Merge_Customers"],
    file_path=r"C:\Reports\MonthlyDashboard.xlsx"
)
```

---

## 🧾 Export and Analysis

```python
# Export index to JSON
handler.export_index_to_json("pq_index.json")

# Convert to DataFrame
df = handler.index_to_dataframe()
print(df.head())

# Validate missing PQ paths
missing = handler.validate_index()
```

---

## 🧩 Advanced Use: Metadata Update

```python
handler.update_pq_metadata("TransformSalesData", {
    "version": "2.0",
    "description": "Now includes quarterly adjustments"
})
```

---

## 🧰 CLI-Style Usage (Example Workflow)

```python
# Step 1: Build your PQ index
handler.build_index()

# Step 2: Quickly search and copy to clipboard
handler.copy_pq_function("Region_Filter")

# Step 3: Paste directly into Power BI Advanced Editor 🚀
```

---

## 📘 Output Example (`index.csv`)

| name               | category  | tags                  | description       | version | path                            |
| ------------------ | --------- | --------------------- | ----------------- | ------- | ------------------------------- |
| TransformSalesData | Sales     | ["merge", "cleaning"] | Cleans sales data | 1.2     | C:\MyPQFiles\Sales_Transform.pq |
| Region_Filter      | Geography | ["filter", "region"]  | Filters by region | 1.0     | C:\MyPQFiles\Region_Filter.pq   |

---

## 🧩 Typing Support

✅ This package is **fully type-hinted** and includes a `py.typed` marker, enabling rich IDE autocompletion and static analysis via `mypy` or `pyright`.

---

## 📜 License

This project is licensed under the **GNU-GPL 3.0 License** — free for personal and commercial use.
Feel free to extend, fork, or integrate into your analytics stack!

---

## 💚 Credits

Created with 💚 by **Sudharshan TK**

If you like this project, ⭐ Star it on [GitHub](https://github.com/tks18/xl-pq-handler)!

---

> ⚡ *“Automate the boring Power Query stuff — one `.pq` at a time.”*

---