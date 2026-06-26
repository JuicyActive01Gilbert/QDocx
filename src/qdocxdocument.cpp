#include <QDocx/qdocxdocument.h>

#include "word/qdocxwordbackend.h"

namespace {

QDocxWordBackend *asBackend(void *backend)
{
    return static_cast<QDocxWordBackend *>(backend);
}

const QDocxWordBackend *asBackend(const void *backend)
{
    return static_cast<const QDocxWordBackend *>(backend);
}

} // namespace

QDocxDocument::QDocxDocument()
    : m_lastResult(QDocxResult::ok())
{
    m_backend = new QDocxWordBackend();
}

QDocxDocument::~QDocxDocument()
{
    close();
    delete asBackend(m_backend);
    m_backend = nullptr;
}

QDocxResult QDocxDocument::open(const QDocxOpenOptions &options)
{
    if (m_isOpen) {
        return m_lastResult = QDocxResult::ok();
    }

    if (!asBackend(m_backend)->openNewWord(options.visible, options.engine)) {
        const QString engineName = options.engine == QDocxOfficeEngine::Wps
            ? QStringLiteral("WPS")
            : QStringLiteral("Word");
        m_lastResult = QDocxResult::fail(QDocxErrorCode::WordStartupFailed,
                                         QStringLiteral("Failed to open a new %1 document.").arg(engineName));
        return m_lastResult;
    }

    m_isOpen = true;
    m_tableCount = 0;
    m_lastResult = QDocxResult::ok();
    return m_lastResult;
}

QDocxResult QDocxDocument::saveAs(const QString &path)
{
    QDocxResult state = requireOpen("saveAs");
    if (!state) {
        return state;
    }
    if (path.trimmed().isEmpty()) {
        m_lastResult = QDocxResult::fail(QDocxErrorCode::InvalidArgument,
                                         QStringLiteral("saveAs requires a non-empty path."));
        return m_lastResult;
    }

    if (!asBackend(m_backend)->saveWord(path)) {
        m_lastResult = QDocxResult::fail(QDocxErrorCode::SaveFailed,
                                         QStringLiteral("Failed to save Word document."));
        return m_lastResult;
    }

    m_lastResult = QDocxResult::ok();
    return m_lastResult;
}

void QDocxDocument::close()
{
    if (m_isOpen) {
        asBackend(m_backend)->quitWord();
        m_isOpen = false;
        m_tableCount = 0;
    }
}

bool QDocxDocument::isOpen() const
{
    return m_isOpen;
}

QDocxResult QDocxDocument::lastResult() const
{
    return m_lastResult;
}

QDocxDocument &QDocxDocument::setDefaultFont(const QDocxFont &font)
{
    if (requireOpen("setDefaultFont")) {
        asBackend(m_backend)->setFontName(font.family);
        asBackend(m_backend)->setFontStyle(font.pointSize, font.bold, font.italic, font.underline);
        asBackend(m_backend)->setTextColor(font.color);
    }
    return *this;
}

QDocxDocument &QDocxDocument::setLineSpacing(QDocxLineSpacing spacing)
{
    if (requireOpen("setLineSpacing")) {
        asBackend(m_backend)->setLineSpace(static_cast<QDocxWordBackend::LineSpacing>(spacing));
    }
    return *this;
}

QDocxDocument &QDocxDocument::addText(const QString &text)
{
    if (requireOpen("addText")) {
        asBackend(m_backend)->addText(text);
    }
    return *this;
}

QDocxDocument &QDocxDocument::addParagraph(const QString &text)
{
    if (requireOpen("addParagraph")) {
        if (!text.isEmpty()) {
            asBackend(m_backend)->addText(text);
        }
        asBackend(m_backend)->newLine();
    }
    return *this;
}

QDocxDocument &QDocxDocument::addHeading(const QString &text, QDocxHeadingLevel level)
{
    if (requireOpen("addHeading")) {
        asBackend(m_backend)->setTitleText(text, static_cast<QDocxWordBackend::TitleLevel>(level));
    }
    return *this;
}

QDocxDocument &QDocxDocument::addPageBreak()
{
    if (requireOpen("addPageBreak")) {
        asBackend(m_backend)->changePage();
    }
    return *this;
}

QDocxDocument &QDocxDocument::addImage(const QString &path, const QDocxImageOptions &options)
{
    if (requireOpen("addImage")) {
        asBackend(m_backend)->setTextAlign(static_cast<QDocxWordBackend::TextAlign>(toBackendAlignment(options.alignment)));
        asBackend(m_backend)->addPic(path);
    }
    return *this;
}

QDocxDocument &QDocxDocument::insertHeader(const QString &text)
{
    if (requireOpen("insertHeader")) {
        asBackend(m_backend)->insertPageHead(text);
    }
    return *this;
}

QDocxDocument &QDocxDocument::insertPageNumbers()
{
    if (requireOpen("insertPageNumbers")) {
        asBackend(m_backend)->insertPageNumber();
    }
    return *this;
}

QDocxDocument &QDocxDocument::insertTableOfContents()
{
    if (requireOpen("insertTableOfContents")) {
        asBackend(m_backend)->insertCatalogue();
    }
    return *this;
}

QDocxDocument &QDocxDocument::updateTableOfContents()
{
    if (requireOpen("updateTableOfContents")) {
        asBackend(m_backend)->updateCatalogue();
    }
    return *this;
}

QDocxTable QDocxDocument::addTable(int rows, int columns)
{
    QDocxResult state = requireOpen("addTable");
    if (!state) {
        return {};
    }

    if (rows <= 0 || columns <= 0) {
        setLastResult(QDocxResult::fail(QDocxErrorCode::InvalidArgument,
                                        QStringLiteral("addTable requires positive row and column counts.")));
        return {};
    }

    asBackend(m_backend)->addTable(rows, columns, QDocxWordBackend::TableFitWindow);
    ++m_tableCount;
    setLastResult(QDocxResult::ok());
    return {m_backend, m_tableCount, rows, columns};
}

QDocxResult QDocxDocument::requireOpen(const char *operation)
{
    if (!m_isOpen) {
        m_lastResult = QDocxResult::fail(
            QDocxErrorCode::InvalidState,
            QStringLiteral("%1 requires an open document.").arg(QString::fromLatin1(operation)));
        return m_lastResult;
    }
    m_lastResult = QDocxResult::ok();
    return m_lastResult;
}

void QDocxDocument::setLastResult(const QDocxResult &result)
{
    m_lastResult = result;
}

int QDocxDocument::toBackendAlignment(QDocxAlignment alignment)
{
    return static_cast<int>(alignment);
}
