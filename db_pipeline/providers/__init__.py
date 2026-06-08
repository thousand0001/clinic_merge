"""FolderDataProvider 與 PostgresDataProvider 的共用資料契約。"""

from db_pipeline.providers.contracts import DataProvider
from db_pipeline.providers.folder import FolderDataProvider
from db_pipeline.providers.postgres import PostgresDataProvider

__all__ = ["DataProvider", "FolderDataProvider", "PostgresDataProvider"]
