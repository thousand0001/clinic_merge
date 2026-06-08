"""各來源系統解析器。

解析器只產生 DatasetBundle，不直接寫資料庫或 Excel。
"""

from .解析器介面 import ParseCoverage, ParseResult, SourceParser
from .耀聖 import SmParser
from .新耀聖 import NewSmParser
from .展望 import ProspectParser
from .醫聖 import MedicalSaintParser
from .調和 import TiaoheParser
from .宏誠 import HongchengParser
from .自行系統 import CustomParser

# source_system → parser 實例的對照表
PARSER_REGISTRY: dict = {
    "sm":            SmParser(),
    "new_sm":        NewSmParser(),
    "prospect":      ProspectParser(),
    "medical_saint": MedicalSaintParser(),
    "tiaohe":        TiaoheParser(),
    "hongcheng":     HongchengParser(),
    "custom":        CustomParser(),
}


def get_parser(source_system: str) -> SourceParser:
    """依 source_system 名稱取得對應解析器；找不到時拋 ValueError。"""
    parser = PARSER_REGISTRY.get(source_system)
    if parser is None:
        raise ValueError(
            f"尚未實作解析器：{source_system}；"
            f"可用系統：{sorted(PARSER_REGISTRY)}"
        )
    return parser


__all__ = [
    "ParseCoverage", "ParseResult", "SourceParser",
    "SmParser", "NewSmParser", "ProspectParser", "MedicalSaintParser",
    "TiaoheParser", "HongchengParser", "CustomParser",
    "PARSER_REGISTRY", "get_parser",
]
