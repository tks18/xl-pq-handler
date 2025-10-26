# storage.py
import os
import json
import time
from typing import List, Optional, Dict
import pandas as pd
from .models import PowerQueryMetadata, PowerQueryScript
from .utils import get_logger

logger = get_logger(__name__)


class PQFileStore:
    """
    Manages the storage and indexing of .pq files on the file system.
    Handles all reading, writing, and index caching.
    """

    def __init__(self, root: str, index_file: str = "index.json"):
        self.root = os.path.abspath(root)
        self.index_path = os.path.join(self.root, index_file)
        self._index: Dict[str, PowerQueryMetadata] = {}  # Name -> Metadata
        self._index_load_time: float = 0.0
        self.load_index()

    # --- Caching ---

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
            logger.info("Loading or refreshing index from disk...")
            self._index = {}
            if not os.path.exists(self.index_path):
                logger.warning(
                    f"Index file not found at {self.index_path}. Building new one.")
                self.build_index()
                return

            try:
                with open(self.index_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for item in data:
                        meta = PowerQueryMetadata(**item)
                        self._index[meta.name.lower()] = meta
                self._index_load_time = time.time()
                logger.info(
                    f"Loaded {len(self._index)} items into index cache.")
            except Exception as e:
                logger.error(
                    f"Failed to read or parse index {self.index_path}: {e}")

    # --- Index & File Operations ---

    def build_index(self) -> str:
        """Build or refresh the JSON index of all .pq files."""
        logger.info(f"Building index from file system root: {self.root}")
        metadata_list: List[Dict] = []
        for dirpath, _, files in os.walk(self.root):
            for fn in files:
                if fn.lower().endswith(".pq"):
                    path = os.path.join(dirpath, fn)
                    try:
                        script = PowerQueryScript.from_file(path)
                        metadata_list.append(script.meta.model_dump())
                    except Exception as e:
                        logger.warning(f"Failed to parse {path}: {e}")

        metadata_list.sort(key=lambda r: (
            r["category"].lower(), r["name"].lower()))

        with open(self.index_path, "w", encoding="utf-8") as f:
            json.dump(metadata_list, f, indent=2, ensure_ascii=False)

        self.load_index(force_reload=True)
        return self.index_path

    def get_metadata_by_name(self, name: str) -> Optional[PowerQueryMetadata]:
        """Fetch PQ metadata from the cache by name (case-insensitive)."""
        self.load_index()  # Ensure cache is fresh
        return self._index.get(name.lower())

    def get_script_by_name(self, name: str) -> Optional[PowerQueryScript]:
        """Fetch *full* parsed PQ by name (reads file for body)."""
        metadata = self.get_metadata_by_name(name)
        if metadata:
            try:
                return PowerQueryScript.from_file(metadata.path)
            except FileNotFoundError:
                logger.warning(
                    f"File not found for {name} at {metadata.path}. Index may be stale."
                )
                return None
        return None

    def save_script(self, script: PowerQueryScript, overwrite: bool = False) -> None:
        """Saves a PowerQueryScript object to its path."""
        script.save(overwrite=overwrite)
        # Update cache immediately
        self._index[script.meta.name.lower()] = script.meta
        # Note: This doesn't rebuild the whole index, just updates the cache.
        # You might still want to call build_index() periodically.

    def delete_script(self, name: str) -> bool:
        """Delete a PQ file by name and refresh index."""
        metadata = self.get_metadata_by_name(name)
        if not metadata:
            logger.warning(
                f"Cannot delete: Query '{name}' not found in index.")
            return False
        try:
            os.remove(metadata.path)
            logger.info(f"Deleted file: {metadata.path}")
            # Easiest way to ensure consistency is to rebuild
            self.build_index()
            return True
        except Exception as e:
            logger.error(f"Failed to delete {metadata.path}: {e}")
            return False

    # --- Search & Access ---

    def search_pq(self, keyword: str) -> List[PowerQueryMetadata]:
        """Search PQ files by name, description, or tags from cache."""
        self.load_index()
        keyword_lower = keyword.lower()
        return [
            meta for meta in self._index.values()
            if keyword_lower in meta.name.lower()
            or keyword_lower in meta.description.lower()
            or any(keyword_lower in t.lower() for t in meta.tags)
        ]

    def list_categories(self) -> List[str]:
        """List all unique categories from cache."""
        self.load_index()
        return sorted(set(e.category for e in self._index.values()))

    def index_to_dataframe(self) -> pd.DataFrame:
        """Return index cache as pandas DataFrame."""
        self.load_index()
        return pd.DataFrame([meta.model_dump() for meta in self._index.values()])
