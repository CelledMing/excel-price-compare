"""在 Windows x64 Python 3.12 开发环境构建可迁移的运行包，不下载依赖。"""
from importlib import metadata
import hashlib
import json
from pathlib import Path
import shutil
import struct
import sys
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent
SOURCE_FILES = [
    "自动统计_统一格式.py", "launcher.py", "run.bat", "requirements.txt",
    "README.md", ".gitignore", "build_portable.py",
    "input/.gitkeep", "output/.gitkeep", "tests/test_core.py",
]


def build():
    if sys.platform != "win32" or sys.version_info[:2] != (3, 12) or struct.calcsize("P") != 8:
        raise SystemExit("构建需要 Windows x64 Python 3.12。")
    runtime = ROOT / "runtime"
    dist = ROOT / "dist"
    if runtime.exists() or dist.exists():
        raise SystemExit("runtime 或 dist 已存在；请在新的源码副本中构建。")
    packages = []
    for line in (ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        name, version = line.split("==")
        package = metadata.distribution(name)
        if package.version != version:
            raise SystemExit(f"版本不匹配：{name} 需要 {version}，已安装 {package.version}")
        packages.append(package)
    for name in SOURCE_FILES:
        if not (ROOT / name).is_file():
            raise SystemExit(f"缺少源码文件：{name}")

    base = Path(sys.base_prefix)
    runtime.mkdir()
    for name in ["python.exe", "python3.dll", "python312.dll", "vcruntime140.dll", "vcruntime140_1.dll", "LICENSE.txt"]:
        shutil.copy2(base / name, runtime / name)
    shutil.copytree(base / "DLLs", runtime / "DLLs", ignore=shutil.ignore_patterns("*_test*", "*.ico", "*.cat", "_tkinter.pyd", "tcl*.dll", "tk*.dll"))
    shutil.copytree(base / "Lib", runtime / "Lib", ignore=shutil.ignore_patterns("site-packages", "__pycache__", "*.pyc", "test", "tests", "idlelib", "tkinter", "ensurepip"))

    site = runtime / "Lib" / "site-packages"
    site.mkdir()
    for package in packages:
        for relative in package.files or []:
            if relative.is_absolute() or ".." in relative.parts or "__pycache__" in relative.parts or relative.suffix == ".pyc":
                continue
            source = Path(package.locate_file(relative))
            if not source.is_file():
                continue
            target = site / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)
    # 不导入 site，不读取注册表或外部 PYTHONPATH；只加载包内的标准库和依赖。
    (runtime / "python312._pth").write_text("Lib\nDLLs\nLib/site-packages\n..\n", encoding="utf-8")
    (runtime / "runtime-manifest.json").write_text(json.dumps({
        "python": sys.version.split()[0], "platform": "windows-x64",
        "packages": {p.metadata["Name"]: p.version for p in packages},
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    write_archives(runtime, dist)


def write_archives(runtime, dist):
    """打包已构建的运行环境；可在修改核心代码后重新生成 ZIP。"""
    dist.mkdir(exist_ok=True)
    for suffix, portable in [("source", False), ("windows-x64", True)]:
        destination = dist / f"excel-price-compare-{suffix}.zip"
        with ZipFile(destination, "w", ZIP_DEFLATED, compresslevel=6) as archive:
            for name in SOURCE_FILES:
                archive.write(ROOT / name, f"excel-price-compare/{name}")
            if portable:
                for path in sorted(runtime.rglob("*")):
                    if path.is_file():
                        archive.write(path, "excel-price-compare/" + path.relative_to(ROOT).as_posix())
        print(f"生成：{destination.name} ({destination.stat().st_size / 1024 / 1024:.1f} MiB)")
    lines = [f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.name}" for path in sorted(dist.glob("*.zip"))]
    (dist / "SHA256SUMS.txt").write_text("\n".join(lines) + "\n", encoding="ascii")


if __name__ == "__main__":
    if sys.argv[1:] == ["--archives-only"]:
        if not (ROOT / "runtime" / "python.exe").is_file():
            raise SystemExit("请先构建 runtime。")
        write_archives(ROOT / "runtime", ROOT / "dist")
    elif sys.argv[1:]:
        raise SystemExit("用法：python build_portable.py [--archives-only]")
    else:
        build()
