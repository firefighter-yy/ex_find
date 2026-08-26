# Excel Information Search 0.1.0

发布日期：2026-08-26

## 支持范围

- Windows 10 或 Windows 11，64 位。
- Microsoft Excel 桌面版，支持 COM 自动化。
- `.xlsx`、`.xlsm` 和 `.xls` 源文件与导出文件。
- 应用产物内置 Python 运行时，目标电脑无需安装 Python。

## 安装与卸载

1. 从发布包中取得 `ExcelInformationSearch.exe`，复制到用户可写目录后运行。
2. 首次运行前确认 Microsoft Excel 可以正常启动，且安全策略允许本地 COM 自动化。
3. 卸载时退出应用并删除程序文件；会话索引在正常退出时自动删除。

## 已知限制

- 不支持无 Microsoft Excel 的电脑。
- 不支持受密码保护、损坏或被独占锁定的工作簿。
- 导出期间请勿修改源文件；检测到变化时必须重新准备文件。
- 应用不会终止用户自行启动的 Excel 进程，也不会上传遥测或崩溃内容。
- 不提供自动更新；升级时替换程序文件即可。

## 可追溯构建

- 项目版本和直接依赖记录在 `pyproject.toml`。
- `scripts/build-windows.ps1` 运行测试并通过 PyInstaller 生成单文件产物。
- 构建同时生成 `dependencies.txt` 和 `SHA256.txt`，用于依赖与产物追溯。
