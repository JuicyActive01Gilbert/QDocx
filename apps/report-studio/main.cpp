#include <QDocx/QDocx>

#include <QApplication>
#include <QBasicTimer>
#include <QCheckBox>
#include <QComboBox>
#include <QCoreApplication>
#include <QDateTime>
#include <QDesktopServices>
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QFormLayout>
#include <QGroupBox>
#include <QHeaderView>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QLabel>
#include <QLineEdit>
#include <QListWidget>
#include <QMainWindow>
#include <QMessageBox>
#include <QPlainTextEdit>
#include <QProcess>
#include <QRegularExpression>
#include <QPushButton>
#include <QSplitter>
#include <QSysInfo>
#include <QTableWidget>
#include <QTextStream>
#include <QTimerEvent>
#include <QTime>
#include <QUrl>
#include <QUuid>
#include <QVBoxLayout>

#include <cstdio>
#include <functional>
#include <utility>

#ifdef Q_OS_WIN
#include <objbase.h>
#include <windows.h>
#endif

namespace {

enum class OfficeStrategy {
    Auto = 0,
    WordOnly,
    WpsOnly
};

struct ReportOptions {
    bool visibleWord = false;
    bool enableHeader = true;
    bool enablePageNumbers = true;
    bool enableTableOfContents = false;
    QDocxLineSpacing lineSpacing = QDocxLineSpacing::OnePointFive;
    QDocxOfficeEngine engine = QDocxOfficeEngine::Word;
    OfficeStrategy strategy = OfficeStrategy::Auto;
};

struct ReportData {
    QString title;
    QString subtitle;
    QString company;
    QString author;
    QString reportNo;
    QString imagePath;
    QString outputPath;
    ReportOptions options;
};

using LogSink = std::function<void(const QString &)>;

QString lineSpacingText(QDocxLineSpacing spacing)
{
    switch (spacing) {
    case QDocxLineSpacing::Single:
        return QStringLiteral("单倍");
    case QDocxLineSpacing::OnePointFive:
        return QStringLiteral("1.5 倍");
    case QDocxLineSpacing::Double:
        return QStringLiteral("双倍");
    case QDocxLineSpacing::AtLeast:
        return QStringLiteral("最小值");
    case QDocxLineSpacing::Exactly:
        return QStringLiteral("固定值");
    case QDocxLineSpacing::Multiple:
        return QStringLiteral("多倍");
    }
    return QStringLiteral("未知");
}

QString officeEngineName(QDocxOfficeEngine engine)
{
    return engine == QDocxOfficeEngine::Wps ? QStringLiteral("WPS") : QStringLiteral("Word");
}

QString officeStrategyText(OfficeStrategy strategy)
{
    switch (strategy) {
    case OfficeStrategy::Auto:
        return QStringLiteral("自动：优先 Office，失败后使用 WPS");
    case OfficeStrategy::WordOnly:
        return QStringLiteral("仅使用 Office Word");
    case OfficeStrategy::WpsOnly:
        return QStringLiteral("仅使用 WPS");
    }
    return QStringLiteral("自动：优先 Office，失败后使用 WPS");
}

OfficeStrategy officeStrategyFromIndex(int index)
{
    switch (index) {
    case 1:
        return OfficeStrategy::WordOnly;
    case 2:
        return OfficeStrategy::WpsOnly;
    default:
        return OfficeStrategy::Auto;
    }
}

int indexFromOfficeStrategy(OfficeStrategy strategy)
{
    return static_cast<int>(strategy);
}

QDocxOfficeEngine officeEngineFromString(const QString &value)
{
    return value.compare(QStringLiteral("wps"), Qt::CaseInsensitive) == 0
        ? QDocxOfficeEngine::Wps
        : QDocxOfficeEngine::Word;
}

QString officeEngineToString(QDocxOfficeEngine engine)
{
    return engine == QDocxOfficeEngine::Wps ? QStringLiteral("wps") : QStringLiteral("word");
}

int officeEngineTimeoutMs(QDocxOfficeEngine engine)
{
    return engine == QDocxOfficeEngine::Word ? 15000 : 60000;
}

QString resultCodeText(QDocxErrorCode code)
{
    switch (code) {
    case QDocxErrorCode::None:
        return QStringLiteral("无错误");
    case QDocxErrorCode::InvalidState:
        return QStringLiteral("状态无效");
    case QDocxErrorCode::WordStartupFailed:
        return QStringLiteral("办公软件启动失败");
    case QDocxErrorCode::SaveFailed:
        return QStringLiteral("保存失败");
    case QDocxErrorCode::InvalidArgument:
        return QStringLiteral("参数无效");
    }
    return QStringLiteral("未知错误");
}

QDocxLineSpacing spacingFromIndex(int index)
{
    switch (index) {
    case 0:
        return QDocxLineSpacing::Single;
    case 1:
        return QDocxLineSpacing::OnePointFive;
    case 2:
        return QDocxLineSpacing::Double;
    default:
        return QDocxLineSpacing::OnePointFive;
    }
}

int indexFromSpacing(QDocxLineSpacing spacing)
{
    switch (spacing) {
    case QDocxLineSpacing::Single:
        return 0;
    case QDocxLineSpacing::OnePointFive:
        return 1;
    case QDocxLineSpacing::Double:
        return 2;
    case QDocxLineSpacing::AtLeast:
    case QDocxLineSpacing::Exactly:
    case QDocxLineSpacing::Multiple:
        return 1;
    }
    return 1;
}

QJsonObject reportDataToJson(const ReportData &data)
{
    QJsonObject options;
    options.insert(QStringLiteral("visibleWord"), data.options.visibleWord);
    options.insert(QStringLiteral("enableHeader"), data.options.enableHeader);
    options.insert(QStringLiteral("enablePageNumbers"), data.options.enablePageNumbers);
    options.insert(QStringLiteral("enableTableOfContents"), data.options.enableTableOfContents);
    options.insert(QStringLiteral("lineSpacing"), indexFromSpacing(data.options.lineSpacing));
    options.insert(QStringLiteral("engine"), officeEngineToString(data.options.engine));
    options.insert(QStringLiteral("strategy"), indexFromOfficeStrategy(data.options.strategy));

    QJsonObject object;
    object.insert(QStringLiteral("title"), data.title);
    object.insert(QStringLiteral("subtitle"), data.subtitle);
    object.insert(QStringLiteral("company"), data.company);
    object.insert(QStringLiteral("author"), data.author);
    object.insert(QStringLiteral("reportNo"), data.reportNo);
    object.insert(QStringLiteral("imagePath"), data.imagePath);
    object.insert(QStringLiteral("outputPath"), data.outputPath);
    object.insert(QStringLiteral("options"), options);
    return object;
}

ReportData reportDataFromJson(const QJsonObject &object)
{
    ReportData data;
    data.title = object.value(QStringLiteral("title")).toString();
    data.subtitle = object.value(QStringLiteral("subtitle")).toString();
    data.company = object.value(QStringLiteral("company")).toString();
    data.author = object.value(QStringLiteral("author")).toString();
    data.reportNo = object.value(QStringLiteral("reportNo")).toString();
    data.imagePath = object.value(QStringLiteral("imagePath")).toString();
    data.outputPath = object.value(QStringLiteral("outputPath")).toString();

    const QJsonObject options = object.value(QStringLiteral("options")).toObject();
    data.options.visibleWord = options.value(QStringLiteral("visibleWord")).toBool(false);
    data.options.enableHeader = options.value(QStringLiteral("enableHeader")).toBool(true);
    data.options.enablePageNumbers = options.value(QStringLiteral("enablePageNumbers")).toBool(true);
    data.options.enableTableOfContents = options.value(QStringLiteral("enableTableOfContents")).toBool(false);
    data.options.lineSpacing = spacingFromIndex(options.value(QStringLiteral("lineSpacing")).toInt(1));
    data.options.engine = officeEngineFromString(options.value(QStringLiteral("engine")).toString(QStringLiteral("word")));
    data.options.strategy = officeStrategyFromIndex(options.value(QStringLiteral("strategy")).toInt(0));
    return data;
}

QJsonArray rowsToJson(const QList<QStringList> &rows)
{
    QJsonArray array;
    for (const QStringList &row : rows) {
        QJsonArray rowArray;
        for (const QString &value : row) {
            rowArray.append(value);
        }
        array.append(rowArray);
    }
    return array;
}

QList<QStringList> rowsFromJson(const QJsonArray &array)
{
    QList<QStringList> rows;
    for (const QJsonValue &rowValue : array) {
        QStringList row;
        for (const QJsonValue &cellValue : rowValue.toArray()) {
            row.append(cellValue.toString());
        }
        rows.append(row);
    }
    return rows;
}

QJsonObject requestToJson(const ReportData &data, const QList<QStringList> &rows)
{
    QJsonObject object;
    object.insert(QStringLiteral("data"), reportDataToJson(data));
    object.insert(QStringLiteral("rows"), rowsToJson(rows));
    return object;
}

bool writeJsonFile(const QString &path, const QJsonObject &object, QString *error)
{
    QFile file(path);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        if (error) {
            *error = file.errorString();
        }
        return false;
    }

    file.write(QJsonDocument(object).toJson(QJsonDocument::Compact));
    return true;
}

bool readJsonFile(const QString &path, QJsonObject *object, QString *error)
{
    QFile file(path);
    if (!file.open(QIODevice::ReadOnly)) {
        if (error) {
            *error = file.errorString();
        }
        return false;
    }

    QJsonParseError parseError;
    const QJsonDocument document = QJsonDocument::fromJson(file.readAll(), &parseError);
    if (parseError.error != QJsonParseError::NoError || !document.isObject()) {
        if (error) {
            *error = parseError.errorString();
        }
        return false;
    }

    *object = document.object();
    return true;
}

void appendLog(QPlainTextEdit *log, const QString &message)
{
    if (!log) {
        return;
    }

    log->appendPlainText(QStringLiteral("[%1] %2")
                             .arg(QTime::currentTime().toString(QStringLiteral("HH:mm:ss")), message));
}

void appendLog(const LogSink &log, const QString &message)
{
    if (log) {
        log(message);
    }
}

QString defaultOutputPath()
{
    return QDir::toNativeSeparators(
        QDir::home().absoluteFilePath(QStringLiteral("Documents/qdocx-report-studio-demo.docx")));
}

bool isWordRegistered()
{
#ifdef Q_OS_WIN
    CLSID clsid;
    const HRESULT result = CLSIDFromProgID(L"Word.Application", &clsid);
    return SUCCEEDED(result);
#else
    return false;
#endif
}

bool isWpsRegistered()
{
#ifdef Q_OS_WIN
    CLSID clsid;
    const HRESULT result = CLSIDFromProgID(L"KWPS.Application", &clsid);
    return SUCCEEDED(result);
#else
    return false;
#endif
}

class ReportWriter
{
public:
    explicit ReportWriter(LogSink log)
        : m_log(std::move(log))
    {
    }

    QDocxResult write(const ReportData &data, const QList<QStringList> &rows)
    {
        QDocxDocument doc;
        appendLog(m_log, QStringLiteral("正在打开 %1 文档").arg(officeEngineName(data.options.engine)));
        QDocxResult result = doc.open({data.options.visibleWord, data.options.engine});
        if (!result) {
            appendLog(m_log, QStringLiteral("打开失败：%1").arg(result.message()));
            return result;
        }

        QDocxFont bodyFont;
        bodyFont.family = QStringLiteral("Microsoft YaHei");
        bodyFont.pointSize = 10.0f;
        bodyFont.color = QColor(30, 30, 30);

        QDocxFont headingFont = bodyFont;
        headingFont.pointSize = 11.0f;
        headingFont.bold = true;

        QDocxFont headerFont = bodyFont;
        headerFont.bold = true;
        headerFont.color = QColor(255, 255, 255);

        doc.setDefaultFont(bodyFont);
        doc.setLineSpacing(data.options.lineSpacing);
        appendLog(m_log, QStringLiteral("已设置默认字体和行距"));

        if (data.options.enableHeader) {
            doc.insertHeader(data.company);
            appendLog(m_log, QStringLiteral("已插入页眉"));
        }
        if (data.options.enablePageNumbers) {
            doc.insertPageNumbers();
            appendLog(m_log, QStringLiteral("已插入页码"));
        }
        if (data.options.enableTableOfContents) {
            doc.insertTableOfContents();
            appendLog(m_log, QStringLiteral("已插入目录"));
        }

        doc.addHeading(data.title, QDocxHeadingLevel::Level1)
            .addParagraph(data.subtitle)
            .addParagraph(QStringLiteral("单位：%1").arg(data.company))
            .addParagraph(QStringLiteral("报告编号：%1").arg(data.reportNo))
            .addParagraph(QStringLiteral("作者：%1").arg(data.author))
            .addParagraph(QStringLiteral("生成时间：%1").arg(QDateTime::currentDateTime().toString(Qt::ISODate)))
            .addHeading(QStringLiteral("概述"), QDocxHeadingLevel::Level2)
            .addParagraph(QStringLiteral("本报告由 QDocx 报告工作台自动生成，用于展示文档生命周期、标题段落、表格格式、页眉页码以及图片插入能力。"))
            .addHeading(QStringLiteral("检测数据"), QDocxHeadingLevel::Level2);
        appendLog(m_log, QStringLiteral("已写入文档章节"));

        QDocxTable table = doc.addTable(rows.size() + 1, 4);
        if (!table.isValid()) {
            result = doc.lastResult();
            doc.close();
            appendLog(m_log, QStringLiteral("表格创建失败：%1").arg(result.message()));
            return result;
        }

        table.setFont(bodyFont).setAlignment(QDocxAlignment::Center);
        const QStringList headers{
            QStringLiteral("检测项"),
            QStringLiteral("标准值"),
            QStringLiteral("实测值"),
            QStringLiteral("结论")
        };
        for (int column = 0; column < headers.size(); ++column) {
            table.cell(1, column + 1)
                .setText(headers.at(column))
                .setFont(headerFont)
                .setBackgroundColor(QColor(79, 129, 189));
        }

        for (int row = 0; row < rows.size(); ++row) {
            const QStringList values = rows.at(row);
            for (int column = 0; column < headers.size(); ++column) {
                const QString value = column < values.size() ? values.at(column) : QString();
                QDocxCell cell = table.cell(row + 2, column + 1);
                cell.setText(value);
                if (column == 3) {
                    cell.setTextColor((value.contains(QStringLiteral("合格"))
                                           || value.contains(QStringLiteral("Pass"), Qt::CaseInsensitive))
                                          ? QColor(0, 128, 0)
                                          : QColor(180, 0, 0));
                }
            }
        }
        table.moveCursorAfter();
        appendLog(m_log, QStringLiteral("已写入检测表格"));

        if (!data.imagePath.trimmed().isEmpty()) {
            doc.addParagraph()
                .addHeading(QStringLiteral("图片"), QDocxHeadingLevel::Level2)
                .addImage(data.imagePath, {QDocxAlignment::Center});
            appendLog(m_log, QStringLiteral("已插入图片：%1").arg(data.imagePath));
        }

        doc.addParagraph()
            .addHeading(QStringLiteral("结论"), QDocxHeadingLevel::Level2)
            .addParagraph(QStringLiteral("报告生成流程已完成。"))
            .addParagraph(QStringLiteral("行距：%1").arg(lineSpacingText(data.options.lineSpacing)));

        if (data.options.enableTableOfContents) {
            doc.updateTableOfContents();
            appendLog(m_log, QStringLiteral("已更新目录"));
        }

        appendLog(m_log, QStringLiteral("正在保存文档：%1").arg(data.outputPath));
        result = doc.saveAs(data.outputPath);
        doc.close();

        if (result) {
            appendLog(m_log, QStringLiteral("保存成功"));
        } else {
            appendLog(m_log, QStringLiteral("保存失败：%1").arg(result.message()));
        }
        return result;
    }

private:
    LogSink m_log;
};

int runWorkerMode(const QString &requestPath)
{
    QTextStream out(stdout);
    QTextStream err(stderr);

    QJsonObject request;
    QString error;
    if (!readJsonFile(requestPath, &request, &error)) {
        err << QStringLiteral("读取生成请求失败：%1").arg(error) << Qt::endl;
        return static_cast<int>(QDocxErrorCode::InvalidArgument);
    }

    const ReportData data = reportDataFromJson(request.value(QStringLiteral("data")).toObject());
    const QList<QStringList> rows = rowsFromJson(request.value(QStringLiteral("rows")).toArray());
    ReportWriter writer([&out](const QString &message) {
        out << message << Qt::endl;
    });

    const QDocxResult result = writer.write(data, rows);
    if (!result) {
        err << result.message() << Qt::endl;
        return static_cast<int>(result.code());
    }

    out << QStringLiteral("WORKER_RESULT_OK:%1").arg(data.outputPath) << Qt::endl;
    return 0;
}

class MainWindow : public QMainWindow
{
public:
    MainWindow()
    {
        setWindowTitle(QStringLiteral("QDocx 报告工作台"));
        resize(1280, 820);

        buildUi();
        loadDefaults();
        updateSummary();
        updateEnvironment();
    }

private:
    void timerEvent(QTimerEvent *event) override
    {
        if (event->timerId() == m_reportTimeoutTimer.timerId() && m_reportProcess) {
            const QDocxOfficeEngine timedOutEngine = m_currentEngine;
            appendLog(m_log, QStringLiteral("%1 生成超时，正在终止生成进程").arg(officeEngineName(timedOutEngine)));
            QObject::disconnect(m_reportProcess, nullptr, this, nullptr);
            m_reportProcess->kill();
            m_reportProcess->waitForFinished(3000);
            finishReportProcess(QDocxResult::fail(
                QDocxErrorCode::WordStartupFailed,
                QStringLiteral("%1 自动化超过 %2 秒未完成。")
                    .arg(officeEngineName(timedOutEngine))
                    .arg(officeEngineTimeoutMs(timedOutEngine) / 1000)));
            return;
        }

        QMainWindow::timerEvent(event);
    }

    void buildUi()
    {
        auto *central = new QWidget;
        auto *root = new QVBoxLayout(central);

        auto *splitter = new QSplitter;
        splitter->addWidget(createTemplatePanel());
        splitter->addWidget(createEditorPanel());
        splitter->addWidget(createSidePanel());
        splitter->setStretchFactor(0, 0);
        splitter->setStretchFactor(1, 1);
        splitter->setStretchFactor(2, 0);
        splitter->setSizes({220, 720, 320});
        root->addWidget(splitter, 1);

        auto *actions = new QHBoxLayout;
        m_outputPath = new QLineEdit;
        auto *browseOutputButton = new QPushButton(QStringLiteral("浏览..."));
        m_generateButton = new QPushButton(QStringLiteral("生成文档"));
        auto *openFolderButton = new QPushButton(QStringLiteral("打开输出目录"));
        auto *clearLogButton = new QPushButton(QStringLiteral("清空日志"));
        auto *apiTestButton = new QPushButton(QStringLiteral("API 验证"));

        actions->addWidget(new QLabel(QStringLiteral("输出文件")));
        actions->addWidget(m_outputPath, 1);
        actions->addWidget(browseOutputButton);
        actions->addWidget(m_generateButton);
        actions->addWidget(openFolderButton);
        actions->addWidget(apiTestButton);
        actions->addWidget(clearLogButton);
        root->addLayout(actions);

        connect(browseOutputButton, &QPushButton::clicked, this, [this] { chooseOutputPath(); });
        connect(m_generateButton, &QPushButton::clicked, this, [this] { generateReport(); });
        connect(openFolderButton, &QPushButton::clicked, this, [this] { openOutputFolder(); });
        connect(clearLogButton, &QPushButton::clicked, this, [this] { m_log->clear(); });
        connect(apiTestButton, &QPushButton::clicked, this, [this] { runApiValidation(); });
        connect(m_outputPath, &QLineEdit::textChanged, this, [this] { updateSummary(); });

        setCentralWidget(central);
    }

    QWidget *createTemplatePanel()
    {
        auto *group = new QGroupBox(QStringLiteral("报告模板"));
        auto *layout = new QVBoxLayout(group);

        m_templateList = new QListWidget;
        m_templateList->addItems({
            QStringLiteral("检测报告"),
            QStringLiteral("表格报告"),
            QStringLiteral("文档结构"),
            QStringLiteral("错误处理")
        });
        m_templateList->setCurrentRow(0);

        auto *description = new QLabel(QStringLiteral("选择一个演示场景。检测报告会展示完整的 QDocx 报告生成流程。"));
        description->setWordWrap(true);

        layout->addWidget(m_templateList);
        layout->addWidget(description);

        connect(m_templateList, &QListWidget::currentTextChanged, this, [this](const QString &text) {
            appendLog(m_log, QStringLiteral("已选择模板：%1").arg(text));
            updateSummary();
        });

        return group;
    }

    QWidget *createEditorPanel()
    {
        auto *panel = new QWidget;
        auto *layout = new QVBoxLayout(panel);

        auto *infoGroup = new QGroupBox(QStringLiteral("报告信息"));
        auto *form = new QFormLayout(infoGroup);
        m_title = new QLineEdit;
        m_subtitle = new QLineEdit;
        m_company = new QLineEdit;
        m_author = new QLineEdit;
        m_reportNo = new QLineEdit;
        m_imagePath = new QLineEdit;
        auto *browseImageButton = new QPushButton(QStringLiteral("浏览..."));
        auto *imageLayout = new QHBoxLayout;
        imageLayout->addWidget(m_imagePath, 1);
        imageLayout->addWidget(browseImageButton);

        form->addRow(QStringLiteral("标题"), m_title);
        form->addRow(QStringLiteral("副标题"), m_subtitle);
        form->addRow(QStringLiteral("单位"), m_company);
        form->addRow(QStringLiteral("作者"), m_author);
        form->addRow(QStringLiteral("报告编号"), m_reportNo);
        form->addRow(QStringLiteral("图片"), imageLayout);

        auto *optionsGroup = new QGroupBox(QStringLiteral("生成选项"));
        auto *options = new QHBoxLayout(optionsGroup);
        m_visibleWord = new QCheckBox(QStringLiteral("显示办公软件"));
        m_enableHeader = new QCheckBox(QStringLiteral("页眉"));
        m_enablePageNumbers = new QCheckBox(QStringLiteral("页码"));
        m_enableToc = new QCheckBox(QStringLiteral("目录"));
        m_officeStrategy = new QComboBox;
        m_officeStrategy->addItems({
            officeStrategyText(OfficeStrategy::Auto),
            officeStrategyText(OfficeStrategy::WordOnly),
            officeStrategyText(OfficeStrategy::WpsOnly)
        });
        m_lineSpacing = new QComboBox;
        m_lineSpacing->addItems({
            QStringLiteral("单倍"),
            QStringLiteral("1.5 倍"),
            QStringLiteral("双倍")
        });
        options->addWidget(m_visibleWord);
        options->addWidget(m_enableHeader);
        options->addWidget(m_enablePageNumbers);
        options->addWidget(m_enableToc);
        options->addWidget(new QLabel(QStringLiteral("办公套件")));
        options->addWidget(m_officeStrategy);
        options->addWidget(new QLabel(QStringLiteral("行距")));
        options->addWidget(m_lineSpacing);
        options->addStretch();

        auto *tableGroup = new QGroupBox(QStringLiteral("检测数据"));
        auto *tableLayout = new QVBoxLayout(tableGroup);
        m_table = new QTableWidget(3, 4);
        m_table->setHorizontalHeaderLabels({
            QStringLiteral("检测项"),
            QStringLiteral("标准值"),
            QStringLiteral("实测值"),
            QStringLiteral("结论")
        });
        m_table->horizontalHeader()->setStretchLastSection(true);
        m_table->verticalHeader()->setVisible(false);
        m_table->setSelectionBehavior(QAbstractItemView::SelectRows);

        auto *tableActions = new QHBoxLayout;
        auto *addRowButton = new QPushButton(QStringLiteral("添加行"));
        auto *removeRowButton = new QPushButton(QStringLiteral("删除行"));
        tableActions->addWidget(addRowButton);
        tableActions->addWidget(removeRowButton);
        tableActions->addStretch();

        tableLayout->addWidget(m_table);
        tableLayout->addLayout(tableActions);

        layout->addWidget(infoGroup);
        layout->addWidget(optionsGroup);
        layout->addWidget(tableGroup, 1);

        connect(browseImageButton, &QPushButton::clicked, this, [this] { chooseImagePath(); });
        connect(addRowButton, &QPushButton::clicked, this, [this] { addTableRow(); });
        connect(removeRowButton, &QPushButton::clicked, this, [this] { removeSelectedTableRow(); });

        for (auto *lineEdit : {m_title, m_subtitle, m_company, m_author, m_reportNo, m_imagePath}) {
            connect(lineEdit, &QLineEdit::textChanged, this, [this] { updateSummary(); });
        }
        for (auto *check : {m_visibleWord, m_enableHeader, m_enablePageNumbers, m_enableToc}) {
            connect(check, &QCheckBox::toggled, this, [this] { updateSummary(); });
        }
        connect(m_officeStrategy, &QComboBox::currentIndexChanged, this, [this] { updateSummary(); });
        connect(m_lineSpacing, &QComboBox::currentIndexChanged, this, [this] { updateSummary(); });

        return panel;
    }

    QWidget *createSidePanel()
    {
        auto *panel = new QWidget;
        auto *layout = new QVBoxLayout(panel);

        auto *envGroup = new QGroupBox(QStringLiteral("环境状态"));
        auto *envLayout = new QVBoxLayout(envGroup);
        m_environment = new QLabel;
        m_environment->setWordWrap(true);
        auto *refreshEnvButton = new QPushButton(QStringLiteral("刷新"));
        envLayout->addWidget(m_environment);
        envLayout->addWidget(refreshEnvButton);

        auto *summaryGroup = new QGroupBox(QStringLiteral("报告摘要"));
        auto *summaryLayout = new QVBoxLayout(summaryGroup);
        m_summary = new QLabel;
        m_summary->setWordWrap(true);
        summaryLayout->addWidget(m_summary);

        auto *logGroup = new QGroupBox(QStringLiteral("操作日志"));
        auto *logLayout = new QVBoxLayout(logGroup);
        m_log = new QPlainTextEdit;
        m_log->setReadOnly(true);
        logLayout->addWidget(m_log);

        layout->addWidget(envGroup);
        layout->addWidget(summaryGroup);
        layout->addWidget(logGroup, 1);

        connect(refreshEnvButton, &QPushButton::clicked, this, [this] { updateEnvironment(); });
        return panel;
    }

    void loadDefaults()
    {
        m_title->setText(QStringLiteral("QDocx 自动化报告"));
        m_subtitle->setText(QStringLiteral("由 QDocx 报告工作台生成"));
        m_company->setText(QStringLiteral("QDocx 演示实验室"));
        m_author->setText(qEnvironmentVariable("USERNAME", QStringLiteral("操作员")));
        m_reportNo->setText(QStringLiteral("QD-%1").arg(QDate::currentDate().toString(QStringLiteral("yyyyMMdd"))));
        m_outputPath->setText(defaultOutputPath());
        m_enableHeader->setChecked(true);
        m_enablePageNumbers->setChecked(true);
        m_enableToc->setChecked(false);
        m_officeStrategy->setCurrentIndex(indexFromOfficeStrategy(OfficeStrategy::Auto));
        m_lineSpacing->setCurrentIndex(1);

        const QList<QStringList> defaults{
            {QStringLiteral("电压"), QStringLiteral("220V"), QStringLiteral("219.8V"), QStringLiteral("合格")},
            {QStringLiteral("电流"), QStringLiteral("10A"), QStringLiteral("9.7A"), QStringLiteral("合格")},
            {QStringLiteral("温度"), QStringLiteral("<80C"), QStringLiteral("76C"), QStringLiteral("合格")}
        };
        for (int row = 0; row < defaults.size(); ++row) {
            for (int column = 0; column < defaults.at(row).size(); ++column) {
                m_table->setItem(row, column, new QTableWidgetItem(defaults.at(row).at(column)));
            }
        }
    }

    ReportData collectReportData() const
    {
        ReportData data;
        data.title = m_title->text().trimmed();
        data.subtitle = m_subtitle->text().trimmed();
        data.company = m_company->text().trimmed();
        data.author = m_author->text().trimmed();
        data.reportNo = m_reportNo->text().trimmed();
        data.imagePath = m_imagePath->text().trimmed();
        data.outputPath = m_outputPath->text().trimmed();
        data.options.visibleWord = m_visibleWord->isChecked();
        data.options.enableHeader = m_enableHeader->isChecked();
        data.options.enablePageNumbers = m_enablePageNumbers->isChecked();
        data.options.enableTableOfContents = m_enableToc->isChecked();
        data.options.lineSpacing = spacingFromIndex(m_lineSpacing->currentIndex());
        data.options.strategy = officeStrategyFromIndex(m_officeStrategy->currentIndex());
        return data;
    }

    QList<QStringList> collectRows() const
    {
        QList<QStringList> rows;
        for (int row = 0; row < m_table->rowCount(); ++row) {
            QStringList values;
            bool hasContent = false;
            for (int column = 0; column < m_table->columnCount(); ++column) {
                QTableWidgetItem *item = m_table->item(row, column);
                const QString value = item ? item->text().trimmed() : QString();
                hasContent = hasContent || !value.isEmpty();
                values.append(value);
            }
            if (hasContent) {
                rows.append(values);
            }
        }
        return rows;
    }

    void updateEnvironment()
    {
        const bool wordRegistered = isWordRegistered();
        const bool wpsRegistered = isWpsRegistered();
        m_environment->setText(QStringLiteral("系统：%1\nOffice Word：%2\nWPS：%3\nQt Widgets：可用\nQt ActiveQt：已由 QDocx 链接")
                                   .arg(QSysInfo::prettyProductName(),
                                        wordRegistered ? QStringLiteral("已检测到") : QStringLiteral("未检测到"),
                                        wpsRegistered ? QStringLiteral("已检测到") : QStringLiteral("未检测到")));
        appendLog(m_log, QStringLiteral("已刷新环境状态"));
    }

    void updateSummary()
    {
        if (!m_summary) {
            return;
        }

        const ReportData data = collectReportData();
        m_summary->setText(QStringLiteral("模板：%1\n标题：%2\n办公套件：%3\n数据行：%4\n页眉：%5\n页码：%6\n目录：%7\n输出：%8")
                               .arg(m_templateList ? m_templateList->currentItem()->text() : QStringLiteral("检测报告"),
                                    data.title,
                                    officeStrategyText(data.options.strategy),
                                    QString::number(collectRows().size()),
                                    data.options.enableHeader ? QStringLiteral("开") : QStringLiteral("关"),
                                    data.options.enablePageNumbers ? QStringLiteral("开") : QStringLiteral("关"),
                                    data.options.enableTableOfContents ? QStringLiteral("开") : QStringLiteral("关"),
                                    data.outputPath));
    }

    void chooseOutputPath()
    {
        const QString path = QFileDialog::getSaveFileName(
            this,
            QStringLiteral("保存 Word 报告"),
            m_outputPath->text().isEmpty() ? defaultOutputPath() : m_outputPath->text(),
            QStringLiteral("Word 文档 (*.docx)"));
        if (!path.isEmpty()) {
            m_outputPath->setText(QDir::toNativeSeparators(path));
        }
    }

    void chooseImagePath()
    {
        const QString path = QFileDialog::getOpenFileName(
            this,
            QStringLiteral("选择图片"),
            QDir::homePath(),
            QStringLiteral("图片 (*.png *.jpg *.jpeg *.bmp)"));
        if (!path.isEmpty()) {
            m_imagePath->setText(QDir::toNativeSeparators(path));
        }
    }

    void addTableRow()
    {
        const int row = m_table->rowCount();
        m_table->insertRow(row);
        m_table->setItem(row, 0, new QTableWidgetItem(QStringLiteral("新检测项")));
        m_table->setItem(row, 1, new QTableWidgetItem(QStringLiteral("标准值")));
        m_table->setItem(row, 2, new QTableWidgetItem(QStringLiteral("实测值")));
        m_table->setItem(row, 3, new QTableWidgetItem(QStringLiteral("合格")));
        updateSummary();
    }

    void removeSelectedTableRow()
    {
        const int row = m_table->currentRow();
        if (row >= 0) {
            m_table->removeRow(row);
            updateSummary();
        }
    }

    void generateReport()
    {
        if (m_reportProcess) {
            return;
        }

        const ReportData data = collectReportData();
        if (data.outputPath.isEmpty()) {
            QMessageBox::warning(this, QStringLiteral("缺少输出路径"), QStringLiteral("请选择输出 .docx 文件路径。"));
            return;
        }

        m_pendingData = data;
        m_pendingRows = collectRows();
        m_pendingWpsFallback = data.options.strategy == OfficeStrategy::Auto;
        m_currentEngine = initialEngineForStrategy(data.options.strategy);
        m_workerOutputPath = data.outputPath;

        appendLog(m_log, makeSeparatorText());
        appendLog(m_log, QStringLiteral("用户选择：%1").arg(officeStrategyText(data.options.strategy)));
        startReportWorker(m_currentEngine);
    }

    QDocxOfficeEngine initialEngineForStrategy(OfficeStrategy strategy) const
    {
        return strategy == OfficeStrategy::WpsOnly ? QDocxOfficeEngine::Wps : QDocxOfficeEngine::Word;
    }

    void startReportWorker(QDocxOfficeEngine engine)
    {
        ReportData data = m_pendingData;
        data.options.engine = engine;
        const QString requestPath = QDir::temp().absoluteFilePath(
            QStringLiteral("qdocx-report-studio-%1.json").arg(QUuid::createUuid().toString(QUuid::WithoutBraces)));
        QString error;
        if (!writeJsonFile(requestPath, requestToJson(data, m_pendingRows), &error)) {
            QMessageBox::warning(this,
                                 QStringLiteral("生成失败"),
                                 QStringLiteral("无法写入临时生成请求：%1").arg(error));
            resetReportState();
            return;
        }

        m_workerRequestPath = requestPath;
        m_currentEngine = engine;
        m_generateButton->setEnabled(false);
        m_generateButton->setText(QStringLiteral("生成中..."));

        appendLog(m_log, QStringLiteral("正在启动 %1 生成进程").arg(officeEngineName(engine)));

        auto *process = new QProcess(this);
        m_reportProcess = process;
        m_workerStdout.clear();
        m_workerStderr.clear();

        process->setProgram(QCoreApplication::applicationFilePath());
        process->setArguments({QStringLiteral("--qdocx-worker"), requestPath});
        process->setProcessChannelMode(QProcess::SeparateChannels);

        connect(process, &QProcess::readyReadStandardOutput, this, [this, process] {
            const QString text = QString::fromUtf8(process->readAllStandardOutput());
            m_workerStdout += text;
            for (const QString &line : text.split(QRegularExpression(QStringLiteral("\\r?\\n")), Qt::SkipEmptyParts)) {
                if (!line.startsWith(QStringLiteral("WORKER_RESULT_OK:"))) {
                    appendLog(m_log, line);
                }
            }
        });
        connect(process, &QProcess::readyReadStandardError, this, [this, process] {
            const QString text = QString::fromUtf8(process->readAllStandardError());
            m_workerStderr += text;
            for (const QString &line : text.split(QRegularExpression(QStringLiteral("\\r?\\n")), Qt::SkipEmptyParts)) {
                appendLog(m_log, QStringLiteral("错误：%1").arg(line));
            }
        });
        connect(process, &QProcess::errorOccurred, this, [this](QProcess::ProcessError error) {
            if (error == QProcess::FailedToStart) {
                finishReportProcess(QDocxResult::fail(QDocxErrorCode::WordStartupFailed,
                                                      QStringLiteral("生成进程启动失败。")));
            }
        });
        connect(process,
                QOverload<int, QProcess::ExitStatus>::of(&QProcess::finished),
                this,
                [this](int exitCode, QProcess::ExitStatus exitStatus) {
                    if (!m_reportProcess) {
                        return;
                    }

                    if (exitStatus == QProcess::CrashExit) {
                        finishReportProcess(QDocxResult::fail(QDocxErrorCode::WordStartupFailed,
                                                              QStringLiteral("生成进程异常结束。")));
                        return;
                    }

                    if (exitCode == 0) {
                        finishReportProcess(QDocxResult::ok());
                        return;
                    }

                    const QString message = m_workerStderr.trimmed().isEmpty()
                        ? QStringLiteral("生成进程失败，退出码：%1").arg(exitCode)
                        : m_workerStderr.trimmed();
                    finishReportProcess(QDocxResult::fail(static_cast<QDocxErrorCode>(exitCode), message));
                });

        m_reportTimeoutTimer.start(officeEngineTimeoutMs(engine), this);
        process->start();
    }

    void finishReportProcess(const QDocxResult &result)
    {
        if (!m_reportProcess) {
            return;
        }

        QProcess *process = m_reportProcess;
        m_reportProcess = nullptr;
        m_reportTimeoutTimer.stop();

        process->deleteLater();
        QFile::remove(m_workerRequestPath);
        m_workerRequestPath.clear();

        if (!result && m_currentEngine == QDocxOfficeEngine::Word && m_pendingWpsFallback) {
            m_pendingWpsFallback = false;
            if (isWpsRegistered()) {
                appendLog(m_log, QStringLiteral("Word 生成失败，自动切换到 WPS 重试"));
                startReportWorker(QDocxOfficeEngine::Wps);
                return;
            }
            appendLog(m_log, QStringLiteral("Word 生成失败，且未检测到 WPS COM 注册"));
        }

        m_generateButton->setEnabled(true);
        m_generateButton->setText(QStringLiteral("生成文档"));

        if (result) {
            QMessageBox::information(
                this,
                QStringLiteral("报告生成完成"),
                QStringLiteral("已使用 %1 保存报告：\n%2").arg(officeEngineName(m_currentEngine), m_workerOutputPath));
        } else {
            const QString engineHint = failureHintForStrategy(m_pendingData.options.strategy);
            QMessageBox::critical(this,
                                  QStringLiteral("生成失败"),
                                  QStringLiteral("%1\n\n%2\n\n错误码：%3")
                                      .arg(result.message(), engineHint, resultCodeText(result.code())));
        }

        resetReportState();
    }

    QString failureHintForStrategy(OfficeStrategy strategy) const
    {
        switch (strategy) {
        case OfficeStrategy::WordOnly:
            return QStringLiteral("当前选择为仅使用 Office Word。请手动打开 Word，处理登录、激活、隐私确认或修复提示后再试。");
        case OfficeStrategy::WpsOnly:
            return QStringLiteral("当前选择为仅使用 WPS。请手动打开 WPS，处理登录、隐私确认、默认应用或修复提示后再试。");
        case OfficeStrategy::Auto:
            return QStringLiteral("自动模式下 Office Word 与 WPS 都未能完成自动化。请分别手动打开并处理首次启动提示后再试。");
        }
        return QStringLiteral("办公软件自动化未能完成。");
    }

    void resetReportState()
    {
        if (m_generateButton) {
            m_generateButton->setEnabled(true);
            m_generateButton->setText(QStringLiteral("生成文档"));
        }
        m_reportTimeoutTimer.stop();
        m_workerOutputPath.clear();
        m_workerStdout.clear();
        m_workerStderr.clear();
        m_pendingRows.clear();
        m_pendingData = {};
        m_pendingWpsFallback = false;
        m_currentEngine = QDocxOfficeEngine::Word;
    }

    QString makeSeparatorText() const
    {
        return QStringLiteral("------------------------------");
    }

    void openOutputFolder()
    {
        const QString path = m_outputPath->text().trimmed();
        if (path.isEmpty()) {
            return;
        }
        QDesktopServices::openUrl(QUrl::fromLocalFile(QFileInfo(path).absolutePath()));
    }

    void runApiValidation()
    {
        QDocxDocument doc;
        QDocxTable table = doc.addTable(1, 1);
        if (!table.isValid() && doc.lastResult().code() == QDocxErrorCode::InvalidState) {
            appendLog(m_log, QStringLiteral("API 验证通过：无效状态已正确返回"));
            QMessageBox::information(this, QStringLiteral("API 验证"), QStringLiteral("API 验证通过。"));
        } else {
            appendLog(m_log, QStringLiteral("API 验证失败"));
            QMessageBox::warning(this, QStringLiteral("API 验证"), QStringLiteral("API 验证失败。"));
        }
    }

    QListWidget *m_templateList = nullptr;
    QLineEdit *m_title = nullptr;
    QLineEdit *m_subtitle = nullptr;
    QLineEdit *m_company = nullptr;
    QLineEdit *m_author = nullptr;
    QLineEdit *m_reportNo = nullptr;
    QLineEdit *m_imagePath = nullptr;
    QLineEdit *m_outputPath = nullptr;
    QCheckBox *m_visibleWord = nullptr;
    QCheckBox *m_enableHeader = nullptr;
    QCheckBox *m_enablePageNumbers = nullptr;
    QCheckBox *m_enableToc = nullptr;
    QComboBox *m_officeStrategy = nullptr;
    QComboBox *m_lineSpacing = nullptr;
    QTableWidget *m_table = nullptr;
    QLabel *m_environment = nullptr;
    QLabel *m_summary = nullptr;
    QPushButton *m_generateButton = nullptr;
    QPlainTextEdit *m_log = nullptr;
    QProcess *m_reportProcess = nullptr;
    QBasicTimer m_reportTimeoutTimer;
    QString m_workerRequestPath;
    QString m_workerOutputPath;
    QString m_workerStdout;
    QString m_workerStderr;
    ReportData m_pendingData;
    QList<QStringList> m_pendingRows;
    QDocxOfficeEngine m_currentEngine = QDocxOfficeEngine::Word;
    bool m_pendingWpsFallback = false;
};

} // namespace

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    QApplication::setApplicationName(QStringLiteral("QDocx 报告工作台"));
    QApplication::setOrganizationName(QStringLiteral("QDocx"));

    const QStringList arguments = app.arguments();
    const int workerIndex = arguments.indexOf(QStringLiteral("--qdocx-worker"));
    if (workerIndex >= 0) {
        if (workerIndex + 1 >= arguments.size()) {
            QTextStream(stderr) << QStringLiteral("缺少生成请求文件路径。") << Qt::endl;
            return static_cast<int>(QDocxErrorCode::InvalidArgument);
        }
        return runWorkerMode(arguments.at(workerIndex + 1));
    }

    MainWindow window;
    window.show();
    return app.exec();
}
