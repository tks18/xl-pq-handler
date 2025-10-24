from __future__ import annotations
import csv
import json
import os
import logging
from typing import Any, Dict, List, Optional, TypedDict
import pyperclip
import xlwings as xw
import yaml
import pandas as pd


logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s"
)


class PowerQueryEntry(TypedDict):
    """Typed structure representing one Power Query (.pq) file entry."""
    name: str
    category: str
    tags: List[str]
    description: str
    version: str
    path: str
    body: str


class XLPowerQueryHandler:
    """
    A full-featured Power Query (.pq) handler for Excel / Power BI automation.

    Features:
    ----------
    - Parse and build YAML frontmatter-based index of .pq files.
    - Insert PQs into Excel workbooks (active or specified path).
    - Copy functions to clipboard for Power BI.
    - Manage index consistency (add / update / delete PQs).
    - Export / import index to JSON or pandas DataFrame.

    Parameters
    ----------
    root : str
        Root directory containing Power Query (.pq) files.
    index_name : str, default = "index.csv"
        File name of the generated index CSV within `root`.
    """

    def __init__(self, root: str, index_name: str = "index.csv") -> None:
        self.root = os.path.abspath(root)
        self.index_name = index_name

    # ============================================================
    # Internal Utilities
    # ============================================================

    def __safe_str(self, val: Any) -> str:
        """Return safe single-line string."""
        return "" if val is None else str(val).replace("\r", " ").replace("\n", " ").strip()

    def __index_path(self) -> str:
        return os.path.join(self.root, self.index_name)

    # ============================================================
    # Power Query File Operations
    # ============================================================

    def parse_pq_file(self, path: str) -> PowerQueryEntry:
        """
        Parse a .pq file with optional YAML frontmatter.

        Returns
        -------
        PowerQueryEntry
        """
        with open(path, "r", encoding="utf-8") as fh:
            text = fh.read()

        fm: Dict[str, Any] = {}
        body = text
        if text.lstrip().startswith("---"):
            parts = text.split("---", 2)
            if len(parts) >= 3:
                try:
                    fm = yaml.safe_load(parts[1]) or {}
                except Exception:
                    fm = {}
                body = parts[2].lstrip("\n")

        return PowerQueryEntry(
            name=self.__safe_str(fm.get("name") or os.path.splitext(
                os.path.basename(path))[0]),
            category=self.__safe_str(fm.get("category") or "Uncategorized"),
            tags=fm.get("tags") or [],
            description=self.__safe_str(fm.get("description") or ""),
            version=self.__safe_str(fm.get("version") or ""),
            path=os.path.abspath(path),
            body=body,
        )

    def build_index(self) -> str:
        """
        Build or refresh the index of all .pq files in the root folder.

        Returns
        -------
        str : Absolute path of generated index file.
        """
        rows: List[PowerQueryEntry] = []
        for dirpath, _, files in os.walk(self.root):
            for fn in files:
                if fn.lower().endswith(".pq"):
                    path = os.path.join(dirpath, fn)
                    try:
                        rows.append(self.parse_pq_file(path))
                    except Exception as e:
                        logging.warning(f"Failed to parse {path}: {e}")

        rows.sort(key=lambda r: (r["category"].lower(), r["name"].lower()))

        with open(self.__index_path(), "w", newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
            writer.writerow(["name", "category", "tags",
                            "description", "version", "path"])
            for r in rows:
                writer.writerow([
                    r["name"],
                    r["category"],
                    json.dumps(r["tags"], ensure_ascii=False),
                    r["description"],
                    r["version"],
                    r["path"],
                ])
        return self.__index_path()

    def read_index(self) -> List[PowerQueryEntry]:
        """Read the CSV index into a list of PowerQueryEntry."""
        path = self.__index_path()
        if not os.path.exists(path):
            return []

        entries: List[PowerQueryEntry] = []
        with open(path, "r", encoding="utf-8") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                try:
                    tags = json.loads(row.get("tags", "[]")) or []
                except Exception:
                    tags = [row["tags"]] if row.get("tags") else []
                entries.append(PowerQueryEntry(
                    name=row.get("name", ""),
                    category=row.get("category", ""),
                    tags=tags,
                    description=row.get("description", ""),
                    version=row.get("version", ""),
                    path=row.get("path", ""),
                    body="",
                ))
        return entries

    # ============================================================
    # Index Maintenance
    # ============================================================

    def get_pq_by_name(self, name: str) -> Optional[PowerQueryEntry]:
        """Fetch parsed PQ by name."""
        for e in self.read_index():
            if e["name"].lower() == name.lower():
                return self.parse_pq_file(e["path"])
        return None

    def update_pq_metadata(self, name: str, new_meta: Dict[str, Any]) -> bool:
        """
        Update YAML frontmatter metadata for a PQ file.

        Parameters
        ----------
        name : str
            PQ name.
        new_meta : Dict[str, Any]
            Dict of fields to update.

        Returns
        -------
        bool : True if updated successfully.
        """
        entry = self.get_pq_by_name(name)
        if not entry:
            raise FileNotFoundError(f"{name} not found.")
        path = entry["path"]

        with open(path, "r", encoding="utf-8") as fh:
            text = fh.read()

        if not text.lstrip().startswith("---"):
            logging.warning(
                f"{name} has no YAML frontmatter, skipping update.")
            return False

        parts = text.split("---", 2)
        if len(parts) < 3:
            return False

        fm = yaml.safe_load(parts[1]) or {}
        fm.update(new_meta)
        new_text = f"---\n{yaml.safe_dump(fm, sort_keys=False)}---{parts[2]}"
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(new_text)
        logging.info(f"Metadata updated for {name}")
        return True

    def delete_pq(self, name: str) -> bool:
        """Delete a PQ file by name and refresh index."""
        entry = self.get_pq_by_name(name)
        if not entry:
            return False
        os.remove(entry["path"])
        self.build_index()
        logging.info(f"Deleted PQ: {name}")
        return True

    def refresh_index(self) -> None:
        """Rebuild index from current root."""
        self.build_index()

    # ============================================================
    # Search / Export / Validation
    # ============================================================

    def search_pq(self, keyword: str) -> List[PowerQueryEntry]:
        """Search PQ files by name, description, or tags."""
        keyword = keyword.lower()
        return [
            e for e in self.read_index()
            if keyword in e["name"].lower()
            or keyword in e["description"].lower()
            or any(keyword in t.lower() for t in e["tags"])
        ]

    def list_categories(self) -> List[str]:
        """List all unique categories."""
        return sorted(set(e["category"] for e in self.read_index()))

    def validate_index(self) -> List[str]:
        """Return list of missing file paths from index."""
        return [e["path"] for e in self.read_index() if not os.path.exists(e["path"])]

    def export_index_to_json(self, output_path: str) -> str:
        """Export index to JSON."""
        with open(output_path, "w", encoding="utf-8") as fh:
            json.dump(self.read_index(), fh, indent=2, ensure_ascii=False)
        return output_path

    def index_to_dataframe(self) -> pd.DataFrame:
        """Return index as pandas DataFrame."""
        return pd.DataFrame(self.read_index())

    # ============================================================
    # Excel Integration
    # ============================================================

    def _insert_into_workbook(self, wb: Any, entry: PowerQueryEntry) -> None:
        """Internal helper to insert PQ into given Excel workbook COM object."""
        queries = wb.Queries
        for i in range(queries.Count, 0, -1):
            q = queries.Item(i)
            if q.Name == entry["name"]:
                q.Delete()
        wb.Queries.Add(
            Name=entry["name"], Formula=entry["body"], Description=entry["description"])

    def insert_pq_into_active_excel(self, name: str) -> Dict[str, str]:
        """Insert PQ into currently active Excel workbook."""
        entry = self.get_pq_by_name(name)
        if not entry:
            raise FileNotFoundError(f"{name} not found.")
        app = xw.apps.active
        if not app:
            raise RuntimeError("No active Excel instance found.")
        self._insert_into_workbook(app.api.ActiveWorkbook, entry)
        return {"status": "ok", "name": name}

    def insert_pq_into_excel(self, file_path: str, name: str) -> Dict[str, str]:
        """
        Insert PQ into a specific Excel file (opens if needed).

        Parameters
        ----------
        file_path : str
            Path to Excel workbook.
        name : str
            PQ name.
        """
        entry = self.get_pq_by_name(name)
        if not entry:
            raise FileNotFoundError(f"{name} not found.")

        app = xw.App(visible=False)
        try:
            wb = app.books.open(file_path)
            self._insert_into_workbook(wb.api, entry)
            wb.save()
            wb.close()
        finally:
            app.quit()
        return {"status": "ok", "file": file_path, "name": name}

    def insert_pqs_batch(self, names: List[str], file_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Insert multiple PQs into Excel (active or specified workbook).
        """
        results = []
        app = None
        try:
            if file_path:
                app = xw.App(visible=False)
                wb = app.books.open(file_path)
            else:
                app = xw.apps.active
                if not app:
                    raise RuntimeError("No active Excel instance found.")
                wb = app.api.ActiveWorkbook

            for name in names:
                entry = self.get_pq_by_name(name)
                if entry:
                    self._insert_into_workbook(wb.api, entry)
                    results.append(name)
                else:
                    logging.warning(f"{name} not found in index.")

            if file_path:
                wb.save()
                wb.close()

        finally:
            if file_path and app:
                app.quit()
        return {"status": "ok", "inserted": results}

    def copy_pq_function(self, name: str) -> Dict[str, str]:
        """Copy PQ body to clipboard."""
        entry = self.get_pq_by_name(name)
        if not entry:
            raise FileNotFoundError(f"{name} not found.")
        text = f"// {entry['name']}\n// {entry['description']}\n{entry['body']}".strip()
        pyperclip.copy(text)
        return {"status": "ok", "name": name}
