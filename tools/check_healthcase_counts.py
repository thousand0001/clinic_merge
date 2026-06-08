# 注意：此工具為舊版除錯用途，依賴的 run_merge_0326_V4_2.py 已不存在，目前無法執行。
# 路徑寫死僅供歷史參考，移植時請勿直接使用。
import importlib.util
import sys
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parent.parent / "run_merge_0326_V4_2.py"  # 此檔已不存在
SOURCE_PATH = ""   # 請改用命令列參數或資料夾選擇視窗
TEMPLATE_PATH = ""  # 請改用命令列參數或資料夾選擇視窗


def load_module():
    spec = importlib.util.spec_from_file_location("merge_mod", SCRIPT_PATH)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def main():
    mod = load_module()
    now = mod.datetime.datetime.now(mod._TZ_TW).date()

    source_ctx = mod.load_source(SOURCE_PATH)
    template_ctx = mod.load_template(TEMPLATE_PATH)
    runtime_ctx = mod.fill_basic_data(source_ctx, template_ctx, now)
    mod.fill_external_data(source_ctx, template_ctx, runtime_ctx)

    ws = template_ctx.ws
    cols = template_ctx.cols
    id_col = cols["id"]

    def main_filled(date_key):
        out = set()
        date_col = cols.get(date_key)
        for row in range(template_ctx.data_start, runtime_ctx.last_row + 1):
            pid = mod.normalize_id(ws.cell(row, id_col).value)
            dt = mod.parse_date(ws.cell(row, date_col).value) if date_col else None
            if pid and dt:
                out.add(pid)
        return out

    hmap = mod.build_header_map(source_ctx.sh_health, 1)
    fc = {
        "id": mod.find_column_exact(hmap, ["家醫收案會員ID", "ID"]),
        "hba": mod.find_column_exact(hmap, ["最近一次HbA1c檢查結果(%)"]),
        "hba_dt": mod.find_column_exact(hmap, ["最近一次HbA1c檢查日期"]),
        "ldl": mod.find_column_exact(hmap, ["最近一次LDL檢查結果(mg/dL)"]),
        "ldl_dt": mod.find_column_exact(hmap, ["最近一次LDL檢查日期"]),
        "uacr": mod.find_column_exact(hmap, ["最近一次UACR檢查結果(mg/gm)"]),
        "uacr_dt": mod.find_column_exact(hmap, ["最近一次UACR檢查日期"]),
    }

    source_sets = {"hba_dt": set(), "ldl_dt": set(), "uacr_dt": set()}
    for row in range(2, source_ctx.sh_health.max_row + 1):
        pid = mod.normalize_id(source_ctx.sh_health.cell(row, fc["id"]).value)
        if not pid:
            continue
        for val_key, dt_key, set_key in [
            ("hba", "hba_dt", "hba_dt"),
            ("ldl", "ldl_dt", "ldl_dt"),
            ("uacr", "uacr_dt", "uacr_dt"),
        ]:
            value = mod.parse_float(source_ctx.sh_health.cell(row, fc[val_key]).value)
            dt = mod.parse_date(source_ctx.sh_health.cell(row, fc[dt_key]).value)
            if dt and value is not None and value != 0:
                source_sets[set_key].add(pid)

    for key, label in [("hba_dt", "HbA1c"), ("ldl_dt", "LDL"), ("uacr_dt", "UACR")]:
        src = source_sets[key]
        main_ids = main_filled(key)
        missing = sorted(src - main_ids)
        extra = sorted(main_ids - src)
        print(label)
        print(
            f"source={len(src)} main={len(main_ids)} "
            f"missing={len(missing)} extra={len(extra)}"
        )
        if missing:
            print("missing_sample=", missing[:20])
        if extra:
            print("extra_sample=", extra[:20])


if __name__ == "__main__":
    main()
