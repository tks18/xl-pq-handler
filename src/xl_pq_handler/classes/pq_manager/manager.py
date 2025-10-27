# manager.py
import os
from typing import List, Optional
import subprocess
import sys

from .models import PowerQueryMetadata, PowerQueryScript
from .storage import PQFileStore
from .excel_service import ExcelQueryService
from .dependencies import DependencyResolver
from .utils import get_logger

logger = get_logger(__name__)


class PQManager:
    """
    A full-featured Power Query (.pq) manager.

    This class acts as a 'facade', coordinating the storage,
    excel, and dependency modules.
    """

    def __init__(self, root: str):
        self.store = PQFileStore(root=root)
        self.excel = ExcelQueryService()
        self.resolver = DependencyResolver(self.store)

    def build_index(self) -> None:
        """Rebuilds the entire .pq file index."""
        self.store.build_index()
        logger.info("Index rebuild complete.")

    def search(self, keyword: str) -> List[dict]:
        """Searches the index for a keyword."""
        return [meta.model_dump() for meta in self.store.search_pq(keyword)]

    def list_categories(self) -> List[str]:
        """Lists all unique categories."""
        return self.store.list_categories()

    def get_script(self, name: str) -> Optional[PowerQueryScript]:
        """Gets a full script object by name."""
        return self.store.get_script_by_name(name)

    def extract_from_excel(
        self,
        category: str = "Extracted",
        file_path: Optional[str] = None
    ) -> List[str]:
        """
        Extracts all queries from an Excel file (or active)
        and saves them as .pq files.
        """
        logger.info(f"Starting extraction to category: {category}")
        query_dicts = self.excel.get_queries_from_workbook(file_path)

        if not query_dicts:
            logger.info("No queries found to extract.")
            return []

        created_files = []
        for q in query_dicts:
            try:
                # Sanitize name for file
                safe_name = "".join(c for c in q["name"] if c.isalnum()
                                    or c in (' ', '_', '-')).rstrip()
                # Sanitize category for folder
                safe_category = "".join(c for c in category if c.isalnum()
                                        or c in (' ', '_', '-')).rstrip()
                if not safe_category:
                    safe_category = "Uncategorized"

                target_dir = os.path.join(self.store.root, safe_category)
                target_path = os.path.join(target_dir, f"{safe_name}.pq")

                script = PowerQueryScript(
                    meta=PowerQueryMetadata(
                        name=q["name"],
                        category=category,
                        description=q["description"],
                        tags=["extracted"],
                        dependencies=[],
                        path=target_path
                    ),
                    body=q["formula"]
                )

                self.store.save_script(script, overwrite=True)
                created_files.append(target_path)

            except Exception as e:
                logger.error(f"Failed to save script for {q['name']}: {e}")

        # Rebuild index *once* after all files are saved
        if created_files:
            self.build_index()

        logger.info(
            f"Extraction complete. {len(created_files)} files created.")
        return created_files

    def insert_into_excel(
        self,
        names: List[str],
        file_path: Optional[str] = None,
        workbook_name: Optional[str] = None
    ) -> None:
        """
        Inserts one or more queries (and their dependencies)
        into an Excel file (or active/specified open workbook).
        """
        logger.info(f"Starting insertion for: {names}")

        try:
            # 1. Resolve dependencies
            names_in_order = self.resolver.get_insertion_order(names)
            logger.info(
                f"Full insertion list (with dependencies): {names_in_order}")

            # 2. Get full script objects
            scripts_to_insert: List[PowerQueryScript] = []
            for name in names_in_order:
                script = self.store.get_script_by_name(name)
                if script:
                    scripts_to_insert.append(script)
                else:
                    logger.error(f"Cannot insert: Script '{name}' not found.")
                    raise FileNotFoundError(f"Script '{name}' not found.")

            # 3. Pass to Excel service
            self.excel.insert_queries_into_workbook(
                scripts=scripts_to_insert,
                file_path=file_path,
                workbook_name=workbook_name,
                delete_existing=True
            )

            logger.info("Insertion complete.")

        except Exception as e:
            logger.error(f"Insertion failed: {e}")
            raise  # Re-raise for the caller

    def open_in_editor(self, name: str):
        """
        Opens the .pq file for the given query name in an external editor.
        Tries VS Code first, then falls back to Notepad (Windows) or default OS handler.
        """
        logger.info(f"Attempting to open '{name}' in external editor...")
        script = self.store.get_script_by_name(name)
        if not script:
            raise FileNotFoundError(f"Script '{name}' not found.")

        path = script.meta.path

        try:
            # Try launching VS Code
            subprocess.Popen(["code", path])
            logger.info(f"Launched VS Code for {path}")
        except FileNotFoundError:
            logger.warning(
                "VS Code not found in PATH. Trying default editor...")
            try:
                # Fallback for Windows: notepad or default .pq handler
                if sys.platform == "win32":
                    os.startfile(path)
                # Fallback for macOS/Linux
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
            except Exception as e:
                logger.error(f"Failed to open file in default editor: {e}")
                raise IOError(
                    f"Could not open file in VS Code or default editor.")
        except Exception as e:
            logger.error(f"Error opening in VS Code: {e}")
            raise

    def update_query_metadata(self, old_name: str, new_meta: PowerQueryMetadata):
        """
        Updates a query's metadata. Handles file moves if name or category changes.
        """
        logger.info(f"Updating metadata for '{old_name}'...")

        # 1. Get the original script's metadata and body
        old_meta = self.store.get_metadata_by_name(old_name)
        if not old_meta:
            raise FileNotFoundError(f"Query '{old_name}' not found in index.")

        old_path = old_meta.path
        script = self.store.get_script_by_name(old_name)

        if script is None:
            raise FileNotFoundError(f"Query '{old_name}' not found in store.")

        body = script.body

        # 2. The new_meta object (from the UI) already has the new path
        new_script = PowerQueryScript(meta=new_meta, body=body)

        # 3. Save the script to its new path
        # This creates the new file.
        try:
            self.store.save_script(new_script, overwrite=True)
        except Exception as e:
            raise IOError(
                f"Failed to save updated file at {new_meta.path}: {e}")

        # 4. If path changed, delete the old file
        if old_path != new_meta.path:
            try:
                os.remove(old_path)
                logger.info(f"Removed old file at {old_path}")
            except OSError as e:
                logger.warning(f"Failed to remove old file at {old_path}: {e}")
                # Don't fail the whole operation, just warn.

        # 5. Rebuild index to make all changes live
        self.build_index()
        logger.info(f"Successfully updated '{old_name}' to '{new_meta.name}'.")
