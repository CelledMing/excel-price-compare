"""先统一不同格式的 Excel，再按日合并比价。

依赖：pip install pandas openpyxl xlrd
运行：python 自动统计_统一格式.py
也可指定目录：python 自动统计_统一格式.py --input-dir 路径 --month 9
源 Excel 只读，结果默认写入本脚本旁的 output/output_统一格式.xlsx。
"""

import argparse
import re
import unicodedata
from contextlib import ExitStack
from pathlib import Path

import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter


STANDARD_COLUMNS = ["商品名称", "规格", "单位", "单价", "优惠90%单价"]
MERGE_KEYS = ["商品名称", "规格", "单位"]
BASE_PATH = Path(__file__).resolve().parent
DEFAULT_INPUT = BASE_PATH / "input"
ALIASES = {
    "商品名称": ["商品名称", "名称", "品名", "货品名称", "产品名称", "物品名称"],
    "规格": ["规格", "规格型号", "型号规格", "商品规格"],
    "单位": ["单位", "计量单位"],
    "单价": ["单价", "原单价", "原价", "单价(元)", "单价（元）"],
    "优惠90%单价": ["优惠90%单价", "优惠90%价格", "优惠90%价", "折后单价", "优惠单价", "九折单价", "9折单价"],
    "序号": ["序号", "编号"],
    "数量": ["数量", "商品数量", "采购数量", "配送数量"],
}


def clean_text(value):
    if pd.isna(value):
        return ""
    return unicodedata.normalize("NFKC", str(value)).strip()


def header_key(value):
    return re.sub(r"\s+", "", clean_text(value))


ALIAS_LOOKUP = {header_key(alias): key for key, aliases in ALIASES.items() for alias in aliases}


def numeric(series):
    text = series.astype("string").str.strip()
    text = text.str.replace(r"[,，¥￥\s]", "", regex=True)
    return pd.to_numeric(text, errors="coerce").astype("float64")


def canonical_column(column):
    key = header_key(column)
    # pandas 在读取重复表头时可能生成“商品名称.1”等列名。
    return ALIAS_LOOKUP.get(key) or ALIAS_LOOKUP.get(re.sub(r"\.\d+$", "", key))


def normalize_dataframe(df, include_quantity=False):
    """单个商品名称列直接清洗，多个商品名称列分块后纵向拼接。

    每遇到新的商品名称列开始下一块，每块同类列只取第一次。
    序号在商品名称左侧时随对应块一起处理，可按需保留数量。
    左侧只有名称、序号且右侧存在价格表时，跳过左侧辅助列表。
    """
    names = [canonical_column(column) for column in df.columns]
    product_positions = [i for i, name in enumerate(names) if name == "商品名称"]
    if not product_positions:
        raise ValueError("未识别到商品名称列，请在 ALIASES 中添加对应列名")
    if len(product_positions) == 1:
        return normalize_block(df, include_quantity)

    boundaries = [0]
    for previous, position in zip(product_positions, product_positions[1:]):
        serial_positions = [i for i in range(previous + 1, position) if names[i] == "序号"]
        boundaries.append(serial_positions[-1] if serial_positions else position)
    boundaries.append(len(df.columns))

    ranges = list(zip(boundaries, boundaries[1:]))
    block_columns = [set(names[start:end]) - {None} for start, end in ranges]
    price_columns = {"单价", "优惠90%单价"}
    blocks = []
    for number, (start, end) in enumerate(ranges, start=1):
        # 名称辅助列表不单独构成订单，以右侧完整表的名称和价格为准。
        name_only = block_columns[number - 1] <= {"序号", "商品名称"}
        priced_table_on_right = any(columns & price_columns for columns in block_columns[number:])
        if name_only and priced_table_on_right:
            continue
        try:
            blocks.append(normalize_block(df.iloc[:, start:end], include_quantity))
        except ValueError as exc:
            raise ValueError(f"第 {number} 个商品表块（第 {start + 1} 至 {end} 列）：{exc}") from exc
    return pd.concat(blocks, ignore_index=True)


def normalize_block(df, include_quantity=False):
    """整理一个表块，每种标准列仅使用从左向右找到的第一列。"""
    result = pd.DataFrame(index=df.index)
    for position, column in enumerate(df.columns):
        canonical = canonical_column(column)
        if canonical is None or canonical in result:
            continue
        values = df.iloc[:, position].map(lambda v: pd.NA if clean_text(v) == "" else v)
        result[canonical] = values

    if "商品名称" not in result:
        raise ValueError("未识别到商品名称列，请在 ALIASES 中添加对应列名")
    if not any(column in result for column in ["单价", "优惠90%单价"]):
        raise ValueError("未识别到单价或优惠价列，请检查表头或补充 ALIASES")
    if "序号" in result:
        result = result.loc[numeric(result["序号"]) > 0].copy()

    for column in MERGE_KEYS:
        if column not in result:
            result[column] = ""
        result[column] = result[column].map(clean_text).astype("string")
    result = result.loc[result["商品名称"].ne("")].copy()
    result = result.loc[~result["商品名称"].isin(["合计", "总计", "小计", *ALIASES["商品名称"]])].copy()
    for column in ["单价", "优惠90%单价"]:
        if column not in result:
            result[column] = float("nan")
        result[column] = numeric(result[column])
    result["优惠90%单价"] = result["优惠90%单价"].fillna(result["单价"] * 0.9)
    columns = STANDARD_COLUMNS.copy()
    if include_quantity:
        if "数量" not in result:
            result["数量"] = float("nan")
        result["数量"] = numeric(result["数量"])
        columns.append("数量")
    return result[columns].reset_index(drop=True)


def read_standard_dataframe(opener, sheet_name):
    """扫描前 20 行定位单行表头，兼容标题行数量不同、列顺序不同。"""
    raw = pd.read_excel(opener, sheet_name=sheet_name, header=None)
    for row_number in range(min(20, len(raw))):
        names = [canonical_column(value) for value in raw.iloc[row_number]]
        if "商品名称" in names and any(name in names for name in ["单价", "优惠90%单价"]):
            data = raw.iloc[row_number + 1:].copy()
            data.columns = raw.iloc[row_number].tolist()
            return normalize_dataframe(data, include_quantity=True)
    raise ValueError("前 20 行未找到包含商品名称和价格的表头")


def merge_standard_frames(frames, source_sheets=None):
    """完全相同的记录去重，数量或价格不同的重复商品暂存后追加。

    重复组缺少有效数量时停止，避免误删。
    每个来源只取首条参与一对一合并，其余暂存，当天比价后追加。
    追加记录只含自身来源的价格和数量，不复制其他来源的数据，
    也不再次执行至少两个来源的价格差异筛选；有暂存项的首条也保留。
    """
    source_sheets = source_sheets or {}
    result = None
    duplicate_frames = []
    for source, frame in frames.items():
        location = f"文件={source}，sheet={source_sheets.get(source, '未知（未提供来源sheet）')}"
        current = frame.copy()
        if "数量" not in current:
            current["数量"] = float("nan")
        current["数量"] = numeric(current["数量"])
        duplicate = current.duplicated(MERGE_KEYS, keep=False)
        if current.loc[duplicate, "数量"].isna().any():
            raise ValueError(f"{location}：重复商品缺少有效数量，无法判断是否去重")
        quantity_keys = MERGE_KEYS + ["数量"]
        current = current.drop_duplicates(quantity_keys + ["单价", "优惠90%单价"])
        current = current[MERGE_KEYS + ["优惠90%单价", "数量"]]
        current = current.rename(columns={"优惠90%单价": source, "数量": f"{source} 数量"})
        extra = current.duplicated(MERGE_KEYS, keep="first")
        if extra.any():
            duplicate_frames.append(current.loc[extra].copy())
        current = current.loc[~extra].copy()
        result = current if result is None else result.merge(
            current, on=MERGE_KEYS, how="outer", validate="one_to_one"
        )
    if result is None:
        return pd.DataFrame(columns=MERGE_KEYS)
    prices = result[list(frames)].apply(pd.to_numeric, errors="coerce")
    different = (prices.count(axis=1) >= 2) & ((prices.max(axis=1) - prices.min(axis=1)) > 1e-9)
    if duplicate_frames:
        # 首条也必须保留，否则只有一个来源或首条同价时会丢失原记录。
        duplicate_keys = pd.concat(duplicate_frames, ignore_index=True)[MERGE_KEYS].drop_duplicates()
        has_extra = pd.MultiIndex.from_frame(result[MERGE_KEYS]).isin(
            pd.MultiIndex.from_frame(duplicate_keys)
        )
        different = different | has_extra
    result = result.loc[different].copy()
    if duplicate_frames:
        result = pd.concat([result, *duplicate_frames], ignore_index=True, sort=False)
    return result.reset_index(drop=True)


def sheets_for_month(opener, month):
    sheets = {}
    for name in opener.sheet_names:
        match = re.fullmatch(r"\s*(\d{1,2})\.(\d{1,2})\s*", name)
        if match and int(match[1]) == month and 1 <= int(match[2]) <= 31:
            day = int(match[2])
            if day in sheets:
                raise ValueError(f"同一天有多个工作表：{sheets[day]}、{name}")
            sheets[day] = name
    return sheets


def excel_reader(input_dir=DEFAULT_INPUT, output_path=None, month=None):
    input_dir = Path(input_dir)
    output_path = Path(output_path) if output_path else Path(__file__).resolve().parent / "output" / "output_统一格式.xlsx"
    month = int(input("请用数字输入筛选表格月份：")) if month is None else month
    if not 1 <= month <= 12:
        raise ValueError("月份必须是 1 到 12")
    if not input_dir.is_dir():
        raise ValueError(f"输入目录不存在：{input_dir}")
    files = sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"} and not p.name.startswith("~$") and p.resolve() != output_path.resolve())
    if not files:
        raise ValueError(f"未找到 Excel 文件：{input_dir}")
    if output_path.exists() and any(output_path.resolve() == p.resolve() for p in input_dir.iterdir()):
        raise ValueError("输出文件不能是输入目录中的已有文件")
    sources = [p.name for p in files]
    outputs = {}
    with ExitStack() as stack:
        openers = {p: stack.enter_context(pd.ExcelFile(p)) for p in files}
        sheet_maps = {p: sheets_for_month(opener, month) for p, opener in openers.items()}
        days = sorted({day for sheets in sheet_maps.values() for day in sheets})
        if not days:
            raise ValueError(f"未找到 {month}.日 格式的工作表")
        for day in days:
            frames = {}
            source_sheets = {}
            for path, opener in openers.items():
                sheet = sheet_maps[path].get(day)
                if sheet is None:
                    print(f"跳过：{path.name} 缺少 {month}.{day} 工作表")
                    continue
                try:
                    frames[path.name] = read_standard_dataframe(opener, sheet)
                    source_sheets[path.name] = sheet
                except ValueError as exc:
                    raise ValueError(f"{path.name} / {sheet}：{exc}") from exc
            columns = MERGE_KEYS + [column for source in sources for column in (source, f"{source} 数量")]
            outputs[f"{month}.{day}"] = merge_standard_frames(frames, source_sheets).reindex(columns=columns)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    quantity_fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet, result in outputs.items():
            result.to_excel(writer, sheet_name=sheet, index=False)
            worksheet = writer.sheets[sheet]
            for column_number, column_name in enumerate(result.columns, start=1):
                worksheet.column_dimensions[get_column_letter(column_number)].width = 20
                if column_name.endswith(" 数量"):
                    for row_number in range(1, len(result) + 2):
                        worksheet.cell(row=row_number, column=column_number).fill = quantity_fill
            print(f"{sheet} 已写入，共 {len(result)} 行")
    print(f"完成：{output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--month", type=int)
    args = parser.parse_args()
    try:
        excel_reader(args.input_dir, args.output, args.month)
    except (ValueError, PermissionError, ImportError) as exc:
        parser.exit(1, f"处理失败：{exc}\n")


if __name__ == "__main__":
    main()
