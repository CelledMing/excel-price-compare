# Excel 商品比价工具

读取多个 `.xls` / `.xlsx` 文件，按月选择每日工作表，统一商品列后合并比价。

## Windows 双击运行（不需要安装 Python）

1. 在 GitHub **Releases** 下载 `excel-price-compare-windows-x64.zip`。
2. **完整解压**到可写文件夹，不要在压缩包内直接启动。
3. 把待处理 Excel 放入项目的 `input` 文件夹。
4. 双击 `run.bat`，输入月份，例如 `8`，按回车。
5. 结果在 `output/output_统一格式.xlsx`，按日生成工作表。

便携版面向 Windows 10/11 x64，内置 Python 3.12.10 和所有运行依赖；运行时无需联网、管理员权限或 Excel 软件。请保留整个 `runtime` 目录。运行前关闭已打开的输出文件；再次运行会覆盖同名结果。原始输入文件只读。

`input`、`output` 初始仅有空的 `.gitkeep` 占位文件，不包含实际订单。

## 输入格式和处理规则

- 工作表名为 `月.日`，兼容 `8.9`、`8.09`、`08.09`；只处理所选月份实际存在的日期。
- 在前 20 行查找单行表头，须识别到商品名称和单价或优惠单价。
- 标准列为商品名称、规格、单位、单价、优惠90%单价。数量供内部重复判断，不输出。
- 列名别名在核心脚本的 `ALIASES` 中维护。缺少优惠价时按单价 × 0.9 补齐；若源表单价已是折后价，请先给它正确的优惠价列名，避免再次打折。
- 单个商品表正常清洗；多个完整商品表横向并排时，分别清洗后纵向拼接。左侧只有名称、序号且右侧有价格表时，跳过左侧辅助列表。
- 按商品名称、规格、单位匹配。完全相同的记录去重；数量或价格不同的额外记录暂存，当天合并后追加，不产生多对多组合。
- 普通记录保留至少两个来源有价格且存在差异的商品；有暂存重复项的首条及追加项也保留。追加行只有自身来源价格，其他来源为空。
- 重复商品缺少有效数量会报错，并显示文件和原始 sheet 名称。
- 不自动换算单位、不推断无标题列含义、不反推原价。请核对不符合上述规则的源表。

## 目录

```text
excel-price-compare/
  自动统计_统一格式.py    核心代码
  launcher.py            运行入口
  run.bat                Windows 双击启动
  requirements.txt       固定版本的依赖清单
  build_portable.py      生成源码包及 Windows 便携包
  tests/                 合并和导出的回归检查
  input/                 输入文件（初始为空）
  output/                输出文件（初始为空）
  runtime/               内置隔离运行环境（仅便携包包含）
```

## 上传 GitHub

将源码包 `excel-price-compare-source.zip` 解压，提交目录内的文件，包括 `.gitignore` 和两个 `.gitkeep`。将 `excel-price-compare-windows-x64.zip` 上传为 GitHub Release 附件。用户无需安装 Python 时，应下载 Release 的便携包；GitHub 的 **Code → Download ZIP** 只提供源码。

`.gitignore` 排除了订单、输出、运行环境、虚拟环境和生成的压缩包。发布文件不包含开发者的订单、旧脚本或机器上的其他资料。第三方许可证保留在便携包 `runtime/LICENSE.txt` 和各依赖的 `.dist-info` 中；核心代码的开源许可证由仓库维护者自行确定。

## 开发与重新打包

源码开发使用 Windows x64 Python 3.12，创建可重建的虚拟环境：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe launcher.py
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe build_portable.py
```

无需激活虚拟环境。首次安装依赖需要联网；打包脚本从已安装的解释器和依赖构建便携环境，运行阶段不需要 pip。

已有便携版也可以检查功能：

```powershell
.\runtime\python.exe -B -X utf8 -m unittest discover -s tests -v
.\runtime\python.exe -B -X utf8 launcher.py --month 8
```

支持 `--input-dir`、`--output` 自定义路径。完整构建拒绝覆盖已有 `runtime` / `dist`，重建运行环境请在新的源码副本中执行。仅修改代码后，可运行 `python build_portable.py --archives-only` 更新两个 ZIP 和校验值。

便携环境使用 Python 的相对路径 `._pth` 配置，隔离系统 Python 和用户安装的包，参见 [Python 官方 Windows 路径配置说明](https://docs.python.org/3.12/using/windows.html#finding-modules)。
