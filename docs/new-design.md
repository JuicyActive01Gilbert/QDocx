# QDocx 新结构设计说明

## 目标

QDocx 的目标是提供一个清晰、可维护的 Qt 6 文档自动化库，用于学习和演示 Office Word / WPS COM 自动化。

新的结构将原有直接暴露 COM 细节的接口整理为高层文档模型：

- `QDocxDocument`：文档生命周期和文档级内容。
- `QDocxTable`：表格级操作。
- `QDocxCell`：单元格级操作。
- `QDocxResult`：统一返回状态。
- `QDocxOpenOptions`：打开参数，包括是否显示窗口和使用 Word/WPS。

## 目录结构

```text
include/QDocx/              公共 API
src/                        高层 API 实现
src/word/                   COM 后端实现
examples/basic/             最小示例
tests/manual-report/        本地测试项目
apps/report-studio/         中文演示软件
docs/                       文档
cmake/                      CMake 包配置模板
```

## 设计原则

1. 应用层不直接操作 `QAxObject`。
2. COM 调用集中在 `src/word/qdocxwordbackend.*` 内部。
3. API 使用链式调用组织常见报告生成流程。
4. 错误通过 `QDocxResult` 返回，不依赖异常。
5. 支持 Office Word 和 WPS 两种引擎。
6. 演示软件负责处理用户体验问题，例如生成超时、Word/WPS 切换和环境状态展示。

## Office/WPS 引擎选择

库层通过 `QDocxOpenOptions::engine` 指定引擎：

```cpp
doc.open({false, QDocxOfficeEngine::Word});
doc.open({false, QDocxOfficeEngine::Wps});
```

当前支持的 COM ProgID：

- `Word.Application`
- `KWPS.Application`

## 演示软件策略

`qdocx_report_studio` 提供三种办公套件选择：

- 自动：优先 Office，失败后使用 WPS
- 仅使用 Office Word
- 仅使用 WPS

生成过程放在独立 worker 进程中执行，避免 Office/WPS COM 阻塞导致主界面卡死。
