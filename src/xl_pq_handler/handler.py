from __future__ import annotations
import csv
import json
import os
import logging
import time
from typing import Any, Dict, List, Optional, TypedDict
import pyperclip
import xlwings as xw
import yaml
import pandas as pd

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s"
)


class PowerQueryEntry(TypedDict):
    """Typed structure representing one Power Query (.pq) file entry."""
    name: str
    category: str
    tags: List[str]
    dependencies: List[str]  # New
    description: str
    version: str
    path: str
    body: str


class XLPowerQueryHandler:
    """
    A full-featured Power Query (.pq) handler for Excel / Power BI automation.

    Features:
    - In-memory caching of the index for high performance.
    - Parse/build JSON-based index of .pq files with YAML frontmatter.
    - Recursive insertion to manage PQ dependencies.
    - Extract queries from Excel to .pq files.
    - Create new .pq files programmatically.
    """

    def __init__(self, root: str, index_file: str = "index.json") -> None:
        self.root = os.path.abspath(root)
        self.index_file = index_file
        self.index_path = os.path.join(self.root, self.index_file)
        self._index: List[PowerQueryEntry] = []
        self._index_load_time: float = 0.0

        # Load the index into cache on initialization
        self.load_index()

    # ============================================================
    # Internal Utilities & Caching
    # ============================================================

    def __safe_str(self, val: Any) -> str:
        return "" if val is None else str(val).replace("\r", " ").replace("\n", " ").strip()

    def is_index_stale(self) -> bool:
        """Check if the index file on disk is newer than the cached index."""
        if not self._index:
            return True
        if not os.path.exists(self.index_path):
            return True
        try:
            return os.path.getmtime(self.index_path) > self._index_load_time
        except OSError:
            return True

    def load_index(self, force_reload: bool = False) -> None:
        """Load or reload the index into memory if it's stale or forced."""
        if force_reload or self.is_index_stale():
            logging.info("Loading or refreshing index from disk...")
            self._index = self.read_index()
            self._index_load_time = time.time()
        else:
            logging.debug("Index is up-to-date, using cache.")

    # ============================================================
    # Power Query File & Index Operations (Now JSON)
    # ============================================================

    def parse_pq_file(self, path: str) -> PowerQueryEntry:
        """Parse a .pq file with optional YAML frontmatter."""
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
            dependencies=fm.get("dependencies") or [],  # Read dependencies
            description=self.__safe_str(fm.get("description") or ""),
            version=self.__safe_str(fm.get("version") or ""),
            path=os.path.abspath(path),
            body=body,
        )

    def build_index(self) -> str:
        """Build or refresh the JSON index of all .pq files."""
        rows: List[Dict[str, Any]] = []
        for dirpath, _, files in os.walk(self.root):
            for fn in files:
                if fn.lower().endswith(".pq"):
                    path = os.path.join(dirpath, fn)
                    try:
                        # Parse the file to get metadata
                        entry = self.parse_pq_file(path)
                        # We only want to store metadata in the index, not the body
                        metadata = {k: v for k, v in entry.items()
                                    if k != "body"}
                        rows.append(metadata)
                    except Exception as e:
                        logging.warning(f"Failed to parse {path}: {e}")

        rows.sort(key=lambda r: (r["category"].lower(), r["name"].lower()))

        with open(self.index_path, "w", encoding="utf-8") as fh:
            json.dump(rows, fh, indent=2, ensure_ascii=False)

        # Force the cache to reload
        self.load_index(force_reload=True)
        return self.index_path

    def read_index(self) -> List[PowerQueryEntry]:
        """Read the JSON index into a list of PowerQueryEntry."""
        if not os.path.exists(self.index_path):
            return []

        try:
            with open(self.index_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
                # Ensure all entries conform to the TypedDict
                entries = []
                for row in data:
                    entries.append(PowerQueryEntry(
                        name=row.get("name", ""),
                        category=row.get("category", ""),
                        tags=row.get("tags", []),
                        dependencies=row.get("dependencies", []),
                        description=row.get("description", ""),
                        version=row.get("version", ""),
                        path=row.get("path", ""),
                        body="",  # Body is not stored in the index
                    ))
                return entries
        except Exception as e:
            logging.error(
                f"Failed to read or parse index {self.index_path}: {e}")
            return []

    # ============================================================
    # Index Maintenance
    # ============================================================

    def get_pq_metadata_by_name(self, name: str) -> Optional[PowerQueryEntry]:
        """Fetch PQ metadata from the cache by name."""
        self.load_index()  # Ensure cache is fresh
        name_lower = name.lower()
        for e in self._index:
            if e["name"].lower() == name_lower:
                return e
        return None

    def get_pq_by_name(self, name: str) -> Optional[PowerQueryEntry]:
        """Fetch *full* parsed PQ by name (reads file for body)."""
        metadata = self.get_pq_metadata_by_name(name)
        if metadata:
            try:
                # Re-parse the file to get the full body and latest metadata
                return self.parse_pq_file(metadata["path"])
            except FileNotFoundError:
                logging.warning(
                    f"File not found for {name} at {metadata['path']}. Index may be stale.")
                return None
        return None

    def delete_pq(self, name: str) -> bool:
        """Delete a PQ file by name and refresh index."""
        metadata = self.get_pq_metadata_by_name(name)
        if not metadata:
            return False
        try:
            os.remove(metadata["path"])
            self.build_index()  # Rebuilds index and reloads cache
            logging.info(f"Deleted PQ: {name}")
            return True
        except Exception as e:
            logging.error(f"Failed to delete {name}: {e}")
            return False

    def refresh_index(self) -> None:
        """Rebuild index and force-reload cache."""
        self.build_index()

    # ============================================================
    # Search / Export / Validation (Uses Cache)
    # ============================================================

    def search_pq(self, keyword: str) -> List[PowerQueryEntry]:
        """Search PQ files by name, description, or tags from cache."""
        self.load_index()
        keyword = keyword.lower()
        return [
            e for e in self._index
            if keyword in e["name"].lower()
            or keyword in e["description"].lower()
            or any(keyword in t.lower() for t in e["tags"])
        ]

    def list_categories(self) -> List[str]:
        """List all unique categories from cache."""
        self.load_index()
        return sorted(set(e["category"] for e in self._index))

    def index_to_dataframe(self) -> pd.DataFrame:
        """Return index cache as pandas DataFrame."""
        self.load_index()
        return pd.DataFrame(self._index)

    # ... other export methods like export_to_json ...

    # ============================================================
    # NEW: Create & Extract
    # ============================================================

    def create_new_pq(
        self,
        name: str,
        body: str,
        category: str = "Uncategorized",
        description: str = "",
        tags: Optional[List[str]] = None,
        dependencies: Optional[List[str]] = None,
        version: str = "1.0",
        overwrite: bool = False
    ) -> str:
        """Create a new .pq file in the root directory and refresh index."""
        meta = {
            "name": name,
            "category": category,
            "description": description,
            "tags": tags or [],
            "dependencies": dependencies or [],
            "version": version
        }

        safe_name = "".join(c for c in name if c.isalnum()
                            or c in (' ', '_', '-')).rstrip()
        out_path = os.path.join(self.root, f"{safe_name}.pq")

        if os.path.exists(out_path) and not overwrite:
            raise FileExistsError(
                f"{out_path} already exists. Set overwrite=True.")

        fm = f"---\n{yaml.safe_dump(meta, sort_keys=False)}---\n\n"

        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write(fm + body)

        logging.info(f"Created new PQ file: {out_path}")
        self.build_index()  # Refresh the index
        return out_path

    # --- NEW: Internal helper for extraction ---
    def _extract_queries_from_workbook(self, wb_api: Any, output_root: Optional[str] = None) -> List[str]:
        """Internal helper to extract queries from a workbook COM object."""
        target_root = output_root or self.root
        if not os.path.exists(target_root):
            os.makedirs(target_root)

        created_files = []
        queries = wb_api.Queries

        if queries.Count == 0:
            logging.warning("No queries found in workbook.")
            return []

        for q in queries:
            try:
                body = q.Formula
                # Create a new .pq file for each
                out_path = self.create_new_pq(
                    name=q.Name,
                    body=body,
                    category="Extracted",
                    description=q.Description,
                    tags=["extracted"],
                    overwrite=True  # Always overwrite when extracting
                )
                created_files.append(out_path)
                logging.info(f"Saved query: {q.Name} -> {out_path}")
            except Exception as e:
                logging.error(f"Failed to extract query {q.Name}: {e}")

        return created_files

    # --- REFACTORED: extract_from_excel (by file path) ---
    def extract_from_excel(self, file_path: str, output_root: Optional[str] = None) -> List[str]:
        """Extract all Power Queries from a specific Excel workbook file."""
        logging.info(f"Extracting queries from {file_path}...")
        created_files = []
        app = xw.App(visible=False)
        try:
            wb = app.books.open(file_path)
            created_files = self._extract_queries_from_workbook(
                wb.api, output_root)
        finally:
            app.quit()

        # Rebuild the index *once* after all files are created
        if created_files:
            self.build_index()
        return created_files

    # --- NEW: extract_from_active_excel ---
    def extract_from_active_excel(self, output_root: Optional[str] = None) -> List[str]:
        """Extract all Power Queries from the active Excel workbook."""
        logging.info("Extracting queries from active Excel instance...")

        app = xw.apps.active
        if not app:
            raise RuntimeError("No active Excel instance found.")

        wb_api = app.api.ActiveWorkbook
        if not wb_api:
            raise RuntimeError("No active workbook found.")

        logging.info(f"Found active workbook: {wb_api.Name}")
        created_files = self._extract_queries_from_workbook(
            wb_api, output_root)

        # Rebuild the index *once*
        if created_files:
            self.build_index()
        return created_files

    # ============================================================
    # Excel Integration (with Dependency Management)
    # ============================================================

    def _insert_into_workbook(
        self,
        wb_api: Any,
        name: str,
        inserted_cache: set[str]
    ) -> bool:
        """
        Internal helper to recursively insert a PQ and its dependencies
        into a given Excel workbook COM object.
        """
        if name in inserted_cache:
            return True  # Already processed

        # Get the full entry, including body and dependencies
        entry = self.get_pq_by_name(name)
        if not entry:
            logging.warning(f"Query '{name}' not found in index, skipping.")
            return False

        # 1. Insert dependencies first
        for dep_name in entry.get("dependencies", []):
            if dep_name not in inserted_cache:
                logging.info(
                    f"Inserting dependency for '{name}': '{dep_name}'")
                if not self._insert_into_workbook(wb_api, dep_name, inserted_cache):
                    logging.error(
                        f"Failed to insert dependency '{dep_name}' for '{name}'. Aborting insert for '{name}'.")
                    return False

        # 2. Insert the query itself
        queries = wb_api.Queries
        try:
            # Delete if exists
            for i in range(queries.Count, 0, -1):
                q = queries.Item(i)
                if q.Name == entry["name"]:
                    q.Delete()

            # Add new one
            logging.info(f"Inserting query: {entry['name']}")
            wb_api.Queries.Add(
                Name=entry["name"], Formula=entry["body"], Description=entry["description"])

            inserted_cache.add(name)  # Mark as successful
            return True

        except Exception as e:
            logging.error(f"Failed to insert query '{name}': {e}")
            return False

    def insert_pq_into_active_excel(self, name: str) -> Dict[str, Any]:
        """Insert PQ (and dependencies) into currently active Excel workbook."""
        app = xw.apps.active
        if not app:
            raise RuntimeError("No active Excel instance found.")

        inserted_cache = set()
        if self._insert_into_workbook(app.api.ActiveWorkbook, name, inserted_cache):
            return {"status": "ok", "name": name, "inserted": list(inserted_cache)}
        else:
            raise RuntimeError(f"Failed to insert query '{name}'.")

    def insert_pqs_batch(self, names: List[str], file_path: Optional[str] = None) -> Dict[str, Any]:
        """Insert multiple PQs (and dependencies) into Excel."""
        inserted_cache = set()
        results = {"success": [], "failed": []}

        app = None
        wb = None
        wb_api = None

        try:
            if file_path:
                app = xw.App(visible=False)
                wb = app.books.open(file_path)
                wb_api = wb.api
            else:
                app = xw.apps.active
                if not app:
                    raise RuntimeError("No active Excel instance found.")
                wb_api = app.api.ActiveWorkbook

            for name in names:
                if name not in inserted_cache:
                    if self._insert_into_workbook(wb_api, name, inserted_cache):
                        results["success"].append(name)
                    else:
                        results["failed"].append(name)

            if file_path and wb:
                wb.save()
                wb.close()

            return {"status": "ok", "inserted": list(inserted_cache), "results": results}

        finally:
            if file_path and app:
                app.quit()
