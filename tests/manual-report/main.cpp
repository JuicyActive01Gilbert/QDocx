#include <QCoreApplication>
#include <QDateTime>
#include <QDir>
#include <QDocx/QDocx>
#include <QStringList>
#include <QTextStream>

namespace {

int fail(const QDocxResult &result)
{
    QTextStream(stderr) << "QDocx test failed: " << result.message() << Qt::endl;
    return static_cast<int>(result.code()) == 0 ? 1 : static_cast<int>(result.code());
}

} // namespace

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);
    const bool runWord = app.arguments().contains(QStringLiteral("--run-word"));

    QDocxDocument doc;
    QDocxTable invalidTable = doc.addTable(1, 1);
    if (invalidTable.isValid() || doc.lastResult().code() != QDocxErrorCode::InvalidState) {
        QTextStream(stderr) << "QDocx state validation failed." << Qt::endl;
        return 1;
    }

    if (!runWord) {
        QTextStream(stdout)
            << "QDocx API validation passed. Run with --run-word to generate a Word document."
            << Qt::endl;
        return 0;
    }

    QDocxResult result = doc.open({false});
    if (!result) {
        return fail(result);
    }

    const QDocxFont bodyFont{
        QStringLiteral("Microsoft YaHei"),
        10.0f,
        false,
        false,
        false,
        QColor(30, 30, 30)
    };
    const QDocxFont headerFont{
        QStringLiteral("Microsoft YaHei"),
        10.0f,
        true,
        false,
        false,
        QColor(255, 255, 255)
    };

    doc.setDefaultFont(bodyFont)
        .insertHeader(QStringLiteral("QDocx local function test"))
        .insertPageNumbers()
        .addHeading(QStringLiteral("QDocx Function Test Report"), QDocxHeadingLevel::Level1)
        .addParagraph(QStringLiteral("Generated at: %1").arg(QDateTime::currentDateTime().toString(Qt::ISODate)))
        .addParagraph(QStringLiteral("This document exercises the high-level QDocxDocument API."))
        .addHeading(QStringLiteral("Summary Table"), QDocxHeadingLevel::Level2);

    QDocxTable table = doc.addTable(4, 3);
    if (!table.isValid()) {
        return fail(doc.lastResult());
    }

    table.setFont(bodyFont).setAlignment(QDocxAlignment::Center);
    table.cell(1, 1).setText(QStringLiteral("Feature")).setFont(headerFont).setBackgroundColor(QColor(79, 129, 189));
    table.cell(1, 2).setText(QStringLiteral("Status")).setFont(headerFont).setBackgroundColor(QColor(79, 129, 189));
    table.cell(1, 3).setText(QStringLiteral("Notes")).setFont(headerFont).setBackgroundColor(QColor(79, 129, 189));

    table.cell(2, 1).setText(QStringLiteral("Document lifecycle"));
    table.cell(2, 2).setText(QStringLiteral("OK")).setTextColor(QColor(0, 128, 0));
    table.cell(2, 3).setText(QStringLiteral("Open, write, save, close"));

    table.cell(3, 1).setText(QStringLiteral("Table wrapper"));
    table.cell(3, 2).setText(QStringLiteral("OK")).setTextColor(QColor(0, 128, 0));
    table.cell(3, 3).setText(QStringLiteral("QDocxTable and QDocxCell"));

    table.cell(4, 1).setText(QStringLiteral("Formatting"));
    table.cell(4, 2).setText(QStringLiteral("OK")).setTextColor(QColor(0, 128, 0));
    table.cell(4, 3).setText(QStringLiteral("Font, color and alignment"));
    table.moveCursorAfter();

    doc.addParagraph()
        .addHeading(QStringLiteral("Result"), QDocxHeadingLevel::Level2)
        .addParagraph(QStringLiteral("If this file opens in Word with headings, header/footer, page numbers and a table, the local function test passed."));

    const QString outputPath = QDir::toNativeSeparators(
        QDir::current().absoluteFilePath(QStringLiteral("qdocx-manual-report-test.docx")));
    result = doc.saveAs(outputPath);
    doc.close();

    if (!result) {
        return fail(result);
    }

    QTextStream(stdout) << "Generated: " << outputPath << Qt::endl;
    return 0;
}
