# manager.py
import os
from typing import List, Optional
from .models import PowerQueryMetadata, PowerQueryScript
from .storage import PQFileStore
from .excel_service import ExcelQueryService
from .dependencies import DependencyResolver
from .utils import get_logger
from xl_pq_handler.classes.pq_manager import dependencies

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
        file_path: Optional[str] = None
    ) -> None:
        """
        Inserts one or more queries (and their dependencies)
        into an Excel file (or active).
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
                delete_existing=True
            )

            logger.info("Insertion complete.")

        except Exception as e:
            logger.error(f"Insertion failed: {e}")
            raise  # Re-raise for the caller
