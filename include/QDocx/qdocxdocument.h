#ifndef QDOCXDOCUMENT_H
#define QDOCXDOCUMENT_H

#include <QDocx/qdocxglobal.h>
#include <QDocx/qdocxresult.h>
#include <QDocx/qdocxstyle.h>
#include <QDocx/qdocxtable.h>

#include <QString>

class QDOCX_EXPORT QDocxDocument
{
public:
    QDocxDocument();
    ~QDocxDocument();

    QDocxDocument(const QDocxDocument &) = delete;
    QDocxDocument &operator=(const QDocxDocument &) = delete;

    QDocxResult open(const QDocxOpenOptions &options = {});
    QDocxResult saveAs(const QString &path);
    void close();

    bool isOpen() const;
    QDocxResult lastResult() const;

    QDocxDocument &setDefaultFont(const QDocxFont &font);
    QDocxDocument &setLineSpacing(QDocxLineSpacing spacing);

    QDocxDocument &addText(const QString &text);
    QDocxDocument &addParagraph(const QString &text = QString());
    QDocxDocument &addHeading(const QString &text, QDocxHeadingLevel level = QDocxHeadingLevel::Level1);
    QDocxDocument &addPageBreak();
    QDocxDocument &addImage(const QString &path, const QDocxImageOptions &options = {});

    QDocxDocument &insertHeader(const QString &text);
    QDocxDocument &insertPageNumbers();
    QDocxDocument &insertTableOfContents();
    QDocxDocument &updateTableOfContents();

    QDocxTable addTable(int rows, int columns);

private:
    friend class QDocxTable;
    friend class QDocxCell;

    QDocxResult requireOpen(const char *operation);
    void setLastResult(const QDocxResult &result);
    static int toBackendAlignment(QDocxAlignment alignment);

    void *m_backend = nullptr;
    bool m_isOpen = false;
    int m_tableCount = 0;
    QDocxResult m_lastResult;
};

#endif // QDOCXDOCUMENT_H
