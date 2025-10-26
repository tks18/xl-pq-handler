# excel_service.py
from typing import List, Dict, Any, Optional
import xlwings as xw
from .utils import get_logger
from .models import PowerQueryScript  # Only for type hinting

logger = get_logger(__name__)


class ExcelQueryService:
    """
    Handles all direct interaction with the Excel COM API via xlwings.
    This class is 'dumb' and just executes Excel commands.
    """

    def _get_wb_api_from_path(self, file_path: str) -> Any:
        """(Helper) Opens a workbook and returns its API object."""
        app = xw.App(visible=False)
        try:
            wb = app.books.open(file_path)
            return wb.api, app, wb  # Return all to be managed
        except Exception as e:
            app.quit()
            raise RuntimeError(f"Failed to open Excel file {file_path}: {e}")

    def _get_active_wb_api(self) -> Any:
        """(Helper) Gets the active workbook's API object."""
        try:
            app = xw.apps.active
            if not app:
                raise RuntimeError("No active Excel instance found.")

            wb_api = app.api.ActiveWorkbook
            if not wb_api:
                raise RuntimeError("No active workbook found.")
            return wb_api
        except Exception as e:
            raise RuntimeError(f"Failed to connect to active Excel: {e}")

    def get_queries_from_workbook(
        self, file_path: Optional[str] = None
    ) -> List[Dict[str, str]]:
        """
        Reads all Power Queries from an Excel workbook.
        If file_path is None, uses the active workbook.
        """
        app_to_quit = None
        wb_to_close = None

        try:
            if file_path:
                wb_api, app_to_quit, wb_to_close = self._get_wb_api_from_path(
                    file_path)
                logger.info(f"Reading queries from {file_path}...")
            else:
                wb_api = self._get_active_wb_api()
                logger.info(
                    f"Reading queries from active workbook: {wb_api.Name}")

            queries_found = []
            queries = wb_api.Queries
            if queries.Count == 0:
                logger.warning("No queries found in workbook.")
                return []

            for q in queries:
                try:
                    queries_found.append({
                        "name": q.Name,
                        "formula": q.Formula,
                        "description": q.Description or ""
                    })
                except Exception as e:
                    logger.error(f"Failed to read query {q.Name}: {e}")

            return queries_found

        finally:
            if wb_to_close:
                wb_to_close.close()
            if app_to_quit:
                app_to_quit.quit()

    def insert_queries_into_workbook(
        self,
        scripts: List[PowerQueryScript],
        file_path: Optional[str] = None,
        delete_existing: bool = True
    ) -> None:
        """
        Inserts a list of PowerQueryScript objects into an Excel workbook.
        If file_path is None, uses the active workbook.
        """
        app_to_quit = None
        wb_to_close = None

        try:
            if file_path:
                wb_api, app_to_quit, wb_to_close = self._get_wb_api_from_path(
                    file_path)
                logger.info(f"Inserting queries into {file_path}...")
            else:
                wb_api = self._get_active_wb_api()
                logger.info(
                    f"Inserting queries into active workbook: {wb_api.Name}")

            queries = wb_api.Queries

            # Create a map of existing queries for fast lookup
            existing_queries = {q.Name: q for q in queries}

            for script in scripts:
                try:
                    name = script.meta.name
                    if name in existing_queries:
                        if delete_existing:
                            logger.debug(f"Deleting existing query: {name}")
                            existing_queries[name].Delete()
                        else:
                            logger.warning(
                                f"Query '{name}' already exists. Skipping.")
                            continue

                    logger.info(f"Inserting query: {name}")
                    wb_api.Queries.Add(
                        Name=name,
                        Formula=script.body,
                        Description=script.meta.description
                    )
                except Exception as e:
                    logger.error(
                        f"Failed to insert query '{script.meta.name}': {e}")
                    # Continue with other queries

            if wb_to_close:
                wb_to_close.save()

        finally:
            if wb_to_close:
                wb_to_close.close()
            if app_to_quit:
                app_to_quit.quit()
