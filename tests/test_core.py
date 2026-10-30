import contextlib
import importlib
import io
from pathlib import Path
import tempfile
import unittest

import pandas as pd

core = importlib.import_module("自动统计_统一格式")


class CoreTests(unittest.TestCase):
    def test_side_by_side_and_name_only_prefix(self):
        data = pd.DataFrame([[1, "旧名称", 1, "苹果", 10, 2, "梨", 20]],
                            columns=["序号", "名称", "序号", "商品名称", "单价", "序号", "商品名称", "单价"])
        result = core.normalize_dataframe(data)
        self.assertEqual(result["商品名称"].tolist(), ["苹果", "梨"])
        self.assertEqual(result["优惠90%单价"].tolist(), [9, 18])

    def test_duplicates_keep_prices_without_cartesian_product(self):
        def frame(prices, quantities):
            return core.normalize_dataframe(pd.DataFrame({
                "商品名称": ["雪糕"] * len(prices), "单价": prices, "数量": quantities,
            }), include_quantity=True)
        a = frame([160, 150, 150], [2, 2, 2])
        b = frame([160, 140], [1, 3])
        result = core.merge_standard_frames({"A": a, "B": b})
        self.assertEqual(len(result), 3)
        self.assertEqual(result["A"].dropna().tolist(), [144, 135])
        self.assertEqual(result["B"].dropna().tolist(), [144, 126])

    def test_export_and_date_names(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1]) as folder:
            root = Path(folder)
            inputs = root / "input"
            inputs.mkdir()
            for name, price, sheet in [("A", 10, "8.09"), ("B", 12, "8.9")]:
                pd.DataFrame({"名称": ["苹果"], "单价": [price], "数量": [2]}).to_excel(
                    inputs / f"{name}.xlsx", sheet_name=sheet, index=False, startrow=3)
            with contextlib.redirect_stdout(io.StringIO()):
                output = core.excel_reader(inputs, root / "result.xlsx", 8)
            with pd.ExcelFile(output) as book:
                self.assertEqual(book.sheet_names, ["8.9"])
                data = pd.read_excel(book)
            self.assertEqual(data.columns.tolist(), core.MERGE_KEYS + ["A.xlsx", "B.xlsx"])
            self.assertEqual(len(data), 1)
            self.assertAlmostEqual(data.loc[0, "A.xlsx"], 9)
            self.assertAlmostEqual(data.loc[0, "B.xlsx"], 10.8)


if __name__ == "__main__":
    unittest.main()
