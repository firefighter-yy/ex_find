# Technical Issues（TISSUE）

本目录将 [PRD](../../PRD.md) 和 [ADR-0001](../adr/0001-local-desktop-search-and-excel-automation.md) 拆分为可实施、可验收的技术任务。

## 状态定义

- 待办：尚未开始。
- 进行中：已经开始，尚未满足完成定义。
- 阻塞：存在明确的外部依赖，当前无法继续。
- 完成：验收标准全部满足，相关测试通过。

## 执行顺序

| 编号 | 名称 | 类型 | 优先级 | 依赖 | 状态 |
| --- | --- | --- | --- | --- | --- |
| [TISSUE-001](./TISSUE-001-project-foundation.md) | 建立工程基础与环境诊断 | 工程 | P0 | 无 | 完成 |
| [TISSUE-002](./TISSUE-002-excel-com-ingestion.md) | Excel COM 只读分块读取 | 功能 | P0 | 001 | 完成 |
| [TISSUE-003](./TISSUE-003-session-index.md) | 会话级 SQLite 索引与文件准备 | 功能 | P0 | 001、002 的数据契约 | 待办 |
| [TISSUE-004](./TISSUE-004-search-engine.md) | 值标准化与行级搜索引擎 | 功能 | P0 | 003 | 待办 |
| [TISSUE-005](./TISSUE-005-minimal-ui.md) | 极简搜索与结果预览界面 | 功能 | P0 | 001、003、004 | 待办 |
| [TISSUE-006](./TISSUE-006-export-fidelity-spike.md) | 导出格式保真原型与 ADR-0002 | 技术验证 | P0 | 002 | 待办 |
| [TISSUE-007](./TISSUE-007-export-workbook.md) | 新工作簿导出 | 功能 | P0 | 005、006、ADR-0002 | 待办 |
| [TISSUE-008](./TISSUE-008-hardening-and-packaging.md) | 取消、恢复、集成验收与打包 | 工程 | P1 | 002 至 007 | 待办 |

TISSUE-003、TISSUE-004 和 TISSUE-006 在接口稳定后可以并行推进。TISSUE-007 必须等待 ADR-0002 确定默认导出模式。

## MVP 完成定义

以下条件全部满足时，MVP 才可视为完成：

1. TISSUE-001 至 TISSUE-008 均为完成状态。
2. PRD 第 11 节的验收标准全部通过。
3. ADR-0001 的 COM 生命周期约束已经通过集成测试。
4. ADR-0002 已接受，并明确格式保真的支持边界。
5. 可以在目标 Windows 电脑上安装并运行，不需要开发环境。
6. 测试过程中未修改任何源 Excel 文件，未残留由应用创建的 Excel 进程。
