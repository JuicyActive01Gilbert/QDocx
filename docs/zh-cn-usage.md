# QDocx 中文使用说明

## 1. 项目简介

QDocx 是一个 Qt 6 文档自动化库，用于在 Windows 桌面程序中生成 `.docx` 文档。它封装了 Office Word / WPS COM 自动化细节，对外提供 `QDocxDocument`、`QDocxTable`、`QDocxCell` 等高层接口。

适合学习以下内容：

- Qt 6 动态库项目组织。
- ActiveQt / `QAxObject` 调用 COM 自动化。
- Word/WPS 文档自动生成。
- CMake 管理库、示例、测试和演示软件。

## 2. 环境要求

- 操作系统：Windows
- Qt：Qt 6，需包含 ActiveQt / AxContainer
- 编译器：MSVC
- 构建工具：CMake 3.21+
- 文档软件：Office Word 或 WPS

支持的 COM 引擎：

| 引擎 | ProgID | 说明 |
| --- | --- | --- |
| Office Word | `Word.Application` | Microsoft Office Word |
| WPS | `KWPS.Application` | WPS 文字 |

## 3. 构建项目

将 `CMAKE_PREFIX_PATH` 改成本机 Qt 6 路径：

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64 `
  -DCMAKE_PREFIX_PATH="D:/Qt/6.x.x/msvcxxxx_64"

cmake --build build --config Release
```

运行基础测试：

```powershell
ctest --test-dir build -C Release --output-on-failure
```

## 4. 引入头文件

推荐只包含总头文件：

```cpp
#include <QDocx/QDocx>
```

核心类型：

| 类型 | 用途 |
| --- | --- |
| `QDocxDocument` | 文档打开、写入、保存、关闭 |
| `QDocxTable` | 表格级操作 |
| `QDocxCell` | 单元格级操作 |
| `QDocxResult` | 操作结果和错误状态 |
| `QDocxFont` | 字体设置 |
| `QDocxOpenOptions` | 打开文档选项 |
| `QDocxImageOptions` | 图片插入选项 |

## 5. 创建文档

```cpp
QDocxDocument doc;
QDocxResult result = doc.open({false, QDocxOfficeEngine::Word});
if (!result) {
    qWarning() << result.message();
    return;
}

doc.addHeading(QStringLiteral("报告标题"), QDocxHeadingLevel::Level1)
   .addParagraph(QStringLiteral("第一段正文。"));

result = doc.saveAs(QStringLiteral("report.docx"));
doc.close();
```

## 6. 选择 Office Word 或 WPS

使用 Office Word：

```cpp
doc.open({false, QDocxOfficeEngine::Word});
```

使用 WPS：

```cpp
doc.open({false, QDocxOfficeEngine::Wps});
```

第一个参数表示是否显示办公软件窗口：

```cpp
doc.open({true, QDocxOfficeEngine::Word});  // 显示 Word
doc.open({false, QDocxOfficeEngine::Wps});  // 后台使用 WPS
```

## 7. 设置字体和行距

```cpp
QDocxFont font;
font.family = QStringLiteral("Microsoft YaHei");
font.pointSize = 10.0f;
font.bold = false;
font.color = QColor(30, 30, 30);

doc.setDefaultFont(font);
doc.setLineSpacing(QDocxLineSpacing::OnePointFive);
```

## 8. 页眉、页码和目录

```cpp
doc.insertHeader(QStringLiteral("公司名称"))
   .insertPageNumbers()
   .insertTableOfContents();

doc.addHeading(QStringLiteral("第一章"), QDocxHeadingLevel::Level1)
   .addParagraph(QStringLiteral("章节内容。"));

doc.updateTableOfContents();
```

目录需要在文档内容写完后调用 `updateTableOfContents()`。

## 9. 插入表格

```cpp
QDocxTable table = doc.addTable(3, 3);
if (!table.isValid()) {
    qWarning() << doc.lastResult().message();
    return;
}

QDocxFont headerFont;
headerFont.family = QStringLiteral("Microsoft YaHei");
headerFont.pointSize = 10.0f;
headerFont.bold = true;
headerFont.color = QColor(255, 255, 255);

table.cell(1, 1).setText(QStringLiteral("项目")).setFont(headerFont).setBackgroundColor(QColor(79, 129, 189));
table.cell(1, 2).setText(QStringLiteral("标准")).setFont(headerFont).setBackgroundColor(QColor(79, 129, 189));
table.cell(1, 3).setText(QStringLiteral("结果")).setFont(headerFont).setBackgroundColor(QColor(79, 129, 189));

table.cell(2, 1).setText(QStringLiteral("电压"));
table.cell(2, 2).setText(QStringLiteral("220V"));
table.cell(2, 3).setText(QStringLiteral("合格")).setTextColor(QColor(0, 128, 0));

table.moveCursorAfter();
```

## 10. 合并单元格

```cpp
table.mergeCells(1, 1, 1, 3);
table.cell(1, 1).setText(QStringLiteral("检测结果汇总"));
```

## 11. 插入图片

```cpp
doc.addHeading(QStringLiteral("图片"), QDocxHeadingLevel::Level2)
   .addImage(QStringLiteral("image.png"), {QDocxAlignment::Center});
```

## 12. 错误处理

`QDocxResult` 可用于判断操作是否成功：

```cpp
QDocxResult result = doc.saveAs(QStringLiteral("report.docx"));
if (!result) {
    qWarning() << result.code() << result.message();
}
```

常见失败原因：

- 未安装 Office Word / WPS。
- COM ProgID 未注册。
- Office/WPS 首次启动弹出登录、激活、隐私或修复窗口。
- 输出文件正在被占用。
- 输出目录不存在或无写入权限。

## 13. 演示软件

`apps/report-studio` 是一个中文 Qt Widgets 演示程序。它展示了：

- 当前环境是否检测到 Office Word 和 WPS。
- 用户自主选择办公套件：
  - 自动：优先 Office，失败后使用 WPS
  - 仅使用 Office Word
  - 仅使用 WPS
- 编辑报告信息。
- 编辑检测表格。
- 生成 `.docx` 报告。

构建后运行：

```text
build/apps/report-studio/Release/qdocx_report_studio.exe
```

## 14. 授权限制

本项目仅供学习、研究和技术交流使用，不可商用。商业使用、二次分发、商业产品集成、商业服务集成均需要获得项目作者授权。
