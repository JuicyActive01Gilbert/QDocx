# QDocx

QDocx 是一个基于 Qt 6 / ActiveQt 的 `.docx` 自动化生成库。它通过 COM 调用 Office Word 或 WPS，面向 Windows 桌面应用中的报告、检测单、表格文档和自动化文档生成场景。

> 本项目仅供学习、研究和技术交流使用，不可用于商业用途。商业使用、二次分发、集成到商业产品或提供商业服务前，需要获得项目作者授权。

## 功能特性

- 使用 `QDocxDocument` 管理文档生命周期。
- 支持 Office Word 与 WPS 两种 COM 引擎。
- 支持标题、段落、页眉、页码、目录、图片和分页。
- 支持表格创建、单元格文本、字体、颜色、背景色、对齐和合并。
- 使用 `QDocxResult` 返回错误状态，便于上层应用处理失败场景。
- 提供 Qt Widgets 中文演示程序 `qdocx_report_studio`。
- 提供基础测试项目 `qdocx_manual_report_test`。

## 运行要求

- Windows
- Qt 6，需包含 `Core`、`Gui`、`Widgets`、`AxContainer`
- CMake 3.21+
- MSVC C++ 工具链
- Office Word 或 WPS，至少安装并注册其中一个 COM 自动化接口

常见 COM ProgID：

- Office Word：`Word.Application`
- WPS：`KWPS.Application`

## 项目结构

```text
include/QDocx/              公共头文件
src/                        QDocx 高层 API 实现
src/word/                   Office/WPS COM 后端
examples/basic/             最小使用示例
tests/manual-report/        本地功能测试项目
apps/report-studio/         中文 Qt Widgets 演示软件
docs/zh-cn-usage.md         中文使用说明文档
docs/new-design.md          设计说明
cmake/                      CMake 包配置模板
```

## 构建

请将 `CMAKE_PREFIX_PATH` 替换为本机 Qt 6 安装目录。

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64 `
  -DCMAKE_PREFIX_PATH="D:/Qt/6.x.x/msvcxxxx_64"

cmake --build build --config Release
```

运行测试：

```powershell
ctest --test-dir build -C Release --output-on-failure
```

## 基本用法

```cpp
#include <QDocx/QDocx>

int main()
{
    QDocxDocument doc;

    QDocxResult opened = doc.open({false, QDocxOfficeEngine::Word});
    if (!opened) {
        return 1;
    }

    doc.setDefaultFont({QStringLiteral("Microsoft YaHei"), 10.0f})
       .insertHeader(QStringLiteral("QDocx 示例报告"))
       .insertPageNumbers()
       .addHeading(QStringLiteral("检测报告"), QDocxHeadingLevel::Level1)
       .addParagraph(QStringLiteral("这是一份由 QDocx 自动生成的文档。"));

    QDocxTable table = doc.addTable(2, 2);
    table.cell(1, 1).setText(QStringLiteral("项目"));
    table.cell(1, 2).setText(QStringLiteral("结果"));
    table.cell(2, 1).setText(QStringLiteral("电压"));
    table.cell(2, 2).setText(QStringLiteral("合格")).setTextColor(QColor(0, 128, 0));
    table.moveCursorAfter();

    doc.saveAs(QStringLiteral("report.docx"));
    doc.close();
    return 0;
}
```

使用 WPS：

```cpp
doc.open({false, QDocxOfficeEngine::Wps});
```

## 演示软件

构建后可运行：

```text
build/apps/report-studio/Release/qdocx_report_studio.exe
```

演示软件提供中文界面，包含：

- 当前环境检测：Office Word / WPS 是否可用。
- 办公套件选择：自动、仅 Office Word、仅 WPS。
- 报告信息编辑。
- 检测表格编辑。
- 页眉、页码、目录、行距等选项。
- 一键生成 `.docx` 报告。

## 注意事项

- 本项目依赖 COM 自动化，仅支持 Windows。
- Office Word 或 WPS 首次启动时可能出现登录、激活、隐私确认、默认应用或修复提示，这些窗口可能导致自动化阻塞。
- 演示软件已将生成过程放入独立进程，并提供超时恢复逻辑。
- 不建议在服务端、无桌面会话或无人值守环境中使用 Office/WPS COM 自动化。

## 授权声明

本项目使用现有项目授权声明，并补充以下限制：

- 仅供学习、研究和技术交流使用。
- 禁止未经授权的商业使用。
- 禁止未经授权将本项目集成到商业软件、商业硬件、商业服务或闭源收费产品中。
- 如需商业授权，请联系项目作者获得书面许可。
