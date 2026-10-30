"""便携版入口，确保双击时能看到中文提示和错误。"""
from pathlib import Path
import sys


def main():
    root = Path(__file__).resolve().parent
    for folder in ("input", "output"):
        (root / folder).mkdir(exist_ok=True)
    try:
        from 自动统计_统一格式 import main as run
        print("Excel 商品比价工具")
        print(f"输入目录：{root / 'input'}")
        print(f"默认输出：{root / 'output' / 'output_统一格式.xlsx'}")
        run()
    except KeyboardInterrupt:
        print("\n已取消。")
        return 130
    except EOFError:
        print("未读取到月份，请双击 run.bat 后输入月份数字。")
        return 1
    except ImportError as exc:
        print(f"运行依赖缺失：{exc}。请完整解压便携版 ZIP，保留 runtime 文件夹。")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
