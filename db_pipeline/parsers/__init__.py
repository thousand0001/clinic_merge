"""各來源系統解析器。

解析器只產生 DatasetBundle，不直接寫資料庫或 Excel。
"""

from .contracts     import ParseCoverage, ParseResult, SourceParser
from .sm            import SmParser
from .new_sm        import NewSmParser
from .prospect      import ProspectParser
from .medical_saint import MedicalSaintParser
from .tiaohe        import TiaoheParser
from .hongcheng     import HongchengParser
from .custom        import CustomParser

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
