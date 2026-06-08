"""各來源系統解析器。

解析器只產生 DatasetBundle，不直接寫資料庫或 Excel。
"""

from .contracts import ParseCoverage, ParseResult, SourceParser
from .sm import SmParser

__all__ = ["ParseCoverage", "ParseResult", "SmParser", "SourceParser"]
