#include <QCoreApplication>
#include <QDir>
#include <QDocx/QDocx>

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    QDocxDocument doc;
    QDocxResult opened = doc.open({false});
    if (!opened) {
        return 1;
    }

    doc.setDefaultFont({QStringLiteral("Microsoft YaHei"), 10.0f});
    doc.addHeading(QStringLiteral("QDocx Qt6 example"));
    doc.addParagraph(QStringLiteral("Generated with Microsoft Word COM automation."));

    const QString outputPath = QDir::toNativeSeparators(
        QDir::current().absoluteFilePath(QStringLiteral("qdocx-example.docx")));
    QDocxResult saved = doc.saveAs(outputPath);
    doc.close();

    return saved ? 0 : 2;
}
