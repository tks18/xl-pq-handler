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

    def list_open_workbooks(self) -> List[str]:
        """
        Lists the names of all currently open Excel workbooks.
        Returns a list of strings.
        """
        try:
            app = xw.apps.active
            if not app:
                logger.warning("No active Excel instance found.")
                return []

            return [book.name for book in app.books]
        except Exception as e:
            logger.error(f"Failed to list open workbooks: {e}")
            return []

    def get_queries_from_open_workbook(self, workbook_name: str) -> List[Dict[str, str]]:
        """
        Connects to an already-open workbook by name and gets its queries.
        This *must* be called from the thread that will use the object.
        """
        logger.info(f"Connecting to open workbook: {workbook_name}")
        try:
            app = xw.apps.active
            if not app:
                raise RuntimeError("No active Excel instance found.")

            # Find the book by name
            wb = app.books[workbook_name]
            if not wb:
                raise FileNotFoundError(
                    f"Open workbook '{workbook_name}' not found.")

            # Now we call our existing logic using the wb.api object
            # that was created ON THIS THREAD.
            return self.get_queries_from_workbook(wb_api=wb.api)

        except Exception as e:
            logger.error(
                f"Failed to get queries from open book {workbook_name}: {e}")
            raise  # Re-raise the exception for the UI to catch

    def get_queries_from_workbook(
        self,
        file_path: Optional[str] = None,
        wb_api: Optional[Any] = None
    ) -> List[Dict[str, str]]:
        """
        Reads all Power Queries from an Excel workbook.
        Priority: wb_api > file_path > active_workbook
        """
        app_to_quit = None
        wb_to_close = None

        try:
            if wb_api:
                # Use the provided workbook API object
                logger.info(
                    f"Reading queries from provided workbook: {wb_api.Name}")
            elif file_path:
                wb_api, app_to_quit, wb_to_close = self._get_wb_api_from_path(
                    file_path)
                logger.info(f"Reading queries from {file_path}...")
            else:
                # Fallback to active workbook
                wb_api = self._get_active_wb_api()
                if not wb_api:
                    logger.warning("No active workbook found.")
                    return []
                logger.info(
                    f"Reading queries from active workbook: {wb_api.Name}")

            if not wb_api:
                logger.warning("No active workbook found.")
                return []
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
        workbook_name: Optional[str] = None,
        delete_existing: bool = True
    ) -> None:
        """
        Inserts a list of PowerQueryScript objects into an Excel workbook.
        If file_path is None, uses the active workbook.
        Priority: workbook_name > file_path > active_workbook
        """
        app_to_quit = None
        wb_to_close = None

        try:
            if workbook_name:
                logger.info(f"Connecting to open workbook: {workbook_name}")
                app = xw.apps.active
                if not app:
                    raise RuntimeError(
                        "No active Excel instance found to connect to.")
                wb_api = app.books[workbook_name].api
                logger.info(f"Inserting queries into {workbook_name}...")
            elif file_path:
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
