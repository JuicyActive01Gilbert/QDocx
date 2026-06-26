#include "word/qdocxwordbackend.h"

#include <windows.h>

#include <QDebug>
#include <QFile>
#include <QVariant>
#include <QtAxContainer/QAxObject>

QDocxWordBackend::QDocxWordBackend()
{
}

QDocxWordBackend::~QDocxWordBackend()
{
    deleteObject();
    if (m_isOleInitialized) {
        OleUninitialize();
        m_isOleInitialized = false;
    }
}

bool QDocxWordBackend::openNewWord(const bool &isShow, QDocxOfficeEngine engine)
{
    const HRESULT result = OleInitialize(nullptr);
    if (result != S_OK && result != S_FALSE) {
        qCritical() << "OLE initialization failed";
        return false;
    }
    m_isOleInitialized = true;

    const QString progId = engine == QDocxOfficeEngine::Wps
        ? QStringLiteral("KWPS.Application")
        : QStringLiteral("Word.Application");
    m_pWordApp = new QAxObject(progId);
    if (m_pWordApp->isNull()) {
        qCritical() << "Failed to start office automation engine:" << progId;
        deleteObject();
        return false;
    }
    m_pWordApp->setProperty("Visible", isShow);
    m_pWordApp->setProperty("DisplayAlerts", false);

    QAxObject *pDocuments = m_pWordApp->querySubObject("Documents");
    if (pDocuments->isNull()) {
        deleteObject();
        return false;
    }
    pDocuments->dynamicCall("Add(QString)", "");
    m_pSaveDocument = m_pWordApp->querySubObject("ActiveDocument");
    if (m_pSaveDocument->isNull()) {
        deleteObject();
        return false;
    }

    m_pSelection = m_pWordApp->querySubObject("Selection");
    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetLineSpacingRule(QVariant)", QVariant(LineSpace1pt5));

    return true;
}

bool QDocxWordBackend::saveWord(const QString &path)
{
    if (QFile::exists(path) && !QFile::remove(path)) {
        qWarning() << "Failed to remove existing Word document:" << path;
        return false;
    }

    m_pSaveDocument->dynamicCall("SaveAs(QString)", path);
    return QFile::exists(path);
}

void QDocxWordBackend::quitWord()
{
    if (m_pWordApp) {
        m_pWordApp->setProperty("DisplayAlerts", false);
    }
    if (m_pSaveDocument) {
        m_pSaveDocument->dynamicCall("Close(bool)", false);
    }
    if (m_pWordApp) {
        m_pWordApp->dynamicCall("Quit()");
    }

    deleteObject();
    if (m_isOleInitialized) {
        OleUninitialize();
        m_isOleInitialized = false;
    }
}

void QDocxWordBackend::newLine(int nLine)
{
    if (nLine <= 0) {
        return;
    }

    for (int i = 0; i < nLine; ++i) {
        m_pSelection->dynamicCall("TypeParagraph(void)");
    }
}

void QDocxWordBackend::changePage()
{
    m_pSelection->dynamicCall("InsertBreak(QVariant)", QVariant(7));
}

void QDocxWordBackend::setTitleText(const QString &title, QDocxWordBackend::TitleLevel level)
{
    QAxObject *pSelection = m_pWordApp->querySubObject("Selection");
    pSelection->dynamicCall("SetStyle(QVariant)", QVariant(static_cast<int>(level)));
    pSelection->dynamicCall("TypeText(QString)", title);

    newLine();

    iniParagraphText();
}

void QDocxWordBackend::setCenterTitleText(const QString &title)
{
    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)", QVariant(QDocxWordBackend::AliginCenter));
    m_pSelection->dynamicCall("TypeText(QString)", title);
}

void QDocxWordBackend::iniParagraphText()
{
    QAxObject *pSelection = m_pWordApp->querySubObject("Selection");
    QAxObject *pParagraphFormat = pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetStyle(QVariant)", QVariant(-67));
}

void QDocxWordBackend::setLineSpace(QDocxWordBackend::LineSpacing space)
{
    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetLineSpacingRule(QVariant)", QVariant(space));
}

void QDocxWordBackend::insertPageHead(const QString &text)
{
    m_pWindowActive = m_pWordApp->querySubObject("ActiveWindow");
    m_pPane = m_pWindowActive->querySubObject("ActivePane");
    m_pViewActive = m_pPane->querySubObject("View");
    m_pViewActive->dynamicCall("SetSeekView(QVariant)", QVariant(9));

    addText(text);

    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)", QVariant(QDocxWordBackend::AliginCenter));

    m_pViewActive->dynamicCall("SetSeekView(QVariant)", QVariant(0));
}

void QDocxWordBackend::insertPageNumber()
{
    QAxObject *pFields = m_pSelection->querySubObject("Fields");
    m_pWindowActive = m_pWordApp->querySubObject("ActiveWindow");
    m_pPane = m_pWindowActive->querySubObject("ActivePane");
    m_pViewActive = m_pPane->querySubObject("View");
    m_pViewActive->dynamicCall("SetSeekView(QVariant)", QVariant(10));

    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)", QVariant(QDocxWordBackend::AliginCenter));
    addText("- ");
    pFields->dynamicCall("Add(QVariant,QVariant,QVariant,QVariant)",
                         m_pSelection->querySubObject("Range")->asVariant(),
                         33,
                         "PAGE  ",
                         1);
    addText(" -");
    m_pSelection->dynamicCall("TypeParagraph()");
    m_pViewActive->dynamicCall("SetSeekView(QVariant)", QVariant(0));
}

void QDocxWordBackend::insertCatalogue()
{
    setFontStyle(18, true);
    addText(QStringLiteral("                    Contents\n\n"));
    m_pTablesOfContents = m_pSaveDocument->querySubObject("TablesOfContents");

    QAxObject *pRange = m_pSelection->querySubObject("Range");
    m_pTablesOfContents->dynamicCall("Add(QVariant)", pRange->asVariant());
}

void QDocxWordBackend::updateCatalogue()
{
    int count = m_pTablesOfContents->property("Count").toInt();
    if (count > 0) {
        QAxObject *pTableOfContent = m_pTablesOfContents->querySubObject("Item(int)", count);
        pTableOfContent->dynamicCall("UpdatePageNumbers(void)");
    }
}

void QDocxWordBackend::addText(const QString &text)
{
    m_pSelection->dynamicCall("TypeText(QString)", text);
}

void QDocxWordBackend::addNumber_E(const double &dVal)
{
    QString strText = QString::number(dVal, 'e');
    m_pSelection->dynamicCall("TypeText(QString)", strText);
}

void QDocxWordBackend::addNumber_Int(const int &nNum)
{
    QString strText = QString::number(nNum);
    m_pSelection->dynamicCall("TypeText(QString)", strText);
}

void QDocxWordBackend::addNumber_Float(const float &fVal)
{
    QString strText = QString::number(fVal, 'f');
    m_pSelection->dynamicCall("TypeText(QString)", strText);
}

void QDocxWordBackend::addPic(const QString &path)
{
    QAxObject *pSelection = m_pWordApp->querySubObject("Selection");
    QAxObject *pParagraphFormat = pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetLineSpacingRule(QVariant)", QVariant(0));
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)", QVariant(QDocxWordBackend::AliginCenter));
    QAxObject *pInlineShapes = pSelection->querySubObject("InlineShapes");
    pInlineShapes->dynamicCall("AddPicture(const &QString)", path);

    newLine();
    pParagraphFormat = pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)", QVariant(QDocxWordBackend::AliginCenter));
}

void QDocxWordBackend::setFontStyle(const float &fSize, const bool &isBold, bool isItalic, bool isUnderLine)
{
    if (m_pSelection->isNull()) {
        qWarning() << "Selection is null. Font setup failed.";
        return;
    }

    m_pFont = m_pSelection->querySubObject("Font");
    m_pFont->dynamicCall("SetSize(float)", fSize);
    m_pFont->dynamicCall("SetBold(bool)", isBold);
    m_pFont->dynamicCall("SetItalic(bool)", isItalic);
    m_pFont->dynamicCall("SetUnderline(bool)", isUnderLine);

    m_pSelection->dynamicCall("SetFont(QVariant)", m_pFont->asVariant());
}

void QDocxWordBackend::setTextAlign(QDocxWordBackend::TextAlign textAlign)
{
    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(int)", static_cast<int>(textAlign));
}

void QDocxWordBackend::setTextColor(const QColor &color)
{
    QAxObject *pFont = m_pSelection->querySubObject("Font");
    pFont->dynamicCall("SetColor(QVariant)", color);
}

void QDocxWordBackend::setFontName(const QString &name)
{
    QAxObject *pFont = m_pSelection->querySubObject("Font");
    pFont->dynamicCall("SetName(QString)", name);
}

void QDocxWordBackend::addTable(const int &nRow, const int &nCol, QDocxWordBackend::TableFitBehavior autoFit)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pRange = m_pSelection->querySubObject("Range");

    pTables->dynamicCall("Add(QVariant,int,int,QVariant,QVariant)",
                         pRange->asVariant(),
                         nRow,
                         nCol,
                         1,
                         static_cast<int>(autoFit));
}

void QDocxWordBackend::setTableWidth(const int &nTableIndex, const float &fWidth)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    pTable->setProperty("PreferredWidthType", 2);
    pTable->setProperty("PreferredWidth", fWidth);
}

void QDocxWordBackend::setTableColWidth(const int &nTableIndex, const int &col, const float &fWidth)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCols = pTable->querySubObject("Columns");
    QAxObject *pCol = pCols->querySubObject("Item(int)", col);
    pCol->setProperty("PreferredWidthType", 2);
    pCol->setProperty("PreferredWidth", fWidth);
}

void QDocxWordBackend::setTableRowHeight(const int nTableIndex, const int &row, const float &fHeight)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pRows = pTable->querySubObject("Rows");
    QAxObject *pRow = pRows->querySubObject("Item(int)", row);
    pRow->setProperty("Height", fHeight);
}

void QDocxWordBackend::setCellsBorderStyle(const int &nTableIndex,
                                const int &nStartRow,
                                const int &nStartCol,
                                const int &nEndRow,
                                const int &nEndCol,
                                const QDocxWordBackend::LineStyle &top,
                                const QDocxWordBackend::LineStyle &bottom,
                                const QDocxWordBackend::LineStyle &left,
                                const QDocxWordBackend::LineStyle &right)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pRows = pTable->querySubObject("Rows");
    QAxObject *pCols = pTable->querySubObject("Columns");

    int nRowCount = pRows->property("Count").toInt();
    int nColCount = pCols->property("Count").toInt();

    pTable->dynamicCall("Select(void)");
    QAxObject *pRange = m_pSelection->querySubObject("Range");

    pRange->dynamicCall("MoveStart(QVariant,QVariant)", 10, nStartRow - 1);
    pRange->dynamicCall("MoveEnd(QVariant,QVariant)", 10, nEndRow - nRowCount);
    pRange->dynamicCall("MoveStart(QVariant,QVariant)", 12, nStartCol - 1);
    pRange->dynamicCall("MoveEnd(QVariant,QVariant)", 12, nEndCol - nColCount);

    pRange->dynamicCall("Select(void)");

    QAxObject *pBorders = m_pSelection->querySubObject("Borders");
    pBorders->querySubObject("Item(int)", 1)->dynamicCall("SetLineStyle(int)", static_cast<int>(top));
    pBorders->querySubObject("Item(int)", 2)->dynamicCall("SetLineStyle(int)", static_cast<int>(left));
    pBorders->querySubObject("Item(int)", 3)->dynamicCall("SetLineStyle(int)", static_cast<int>(bottom));
    pBorders->querySubObject("Item(int)", 4)->dynamicCall("SetLineStyle(int)", static_cast<int>(right));
}

void QDocxWordBackend::setCellBorderStyle(const int &nTableIndex,
                               const int &nRow,
                               const int &nCol,
                               const QDocxWordBackend::LineStyle &top,
                               const QDocxWordBackend::LineStyle &bottom,
                               const QDocxWordBackend::LineStyle &left,
                               const QDocxWordBackend::LineStyle &right)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)", nRow, nCol);
    pCell->dynamicCall("Select(void)");

    QAxObject *pBorders = pCell->querySubObject("Borders");
    pBorders->querySubObject("Item(int)", 1)->dynamicCall("SetLineStyle(int)", static_cast<int>(top));
    pBorders->querySubObject("Item(int)", 2)->dynamicCall("SetLineStyle(int)", static_cast<int>(left));
    pBorders->querySubObject("Item(int)", 3)->dynamicCall("SetLineStyle(int)", static_cast<int>(bottom));
    pBorders->querySubObject("Item(int)", 4)->dynamicCall("SetLineStyle(int)", static_cast<int>(right));
}

void QDocxWordBackend::setCellsColor(const int &nTableIndex,
                          const int &nStartRow,
                          const int &nStartCol,
                          const int &nEndRow,
                          const int &nEndCol,
                          const QColor &color)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pRows = pTable->querySubObject("Rows");
    QAxObject *pCols = pTable->querySubObject("Columns");

    int nRowCount = pRows->property("Count").toInt();
    int nColCount = pCols->property("Count").toInt();

    pTable->dynamicCall("Select(void)");
    QAxObject *pRange = m_pSelection->querySubObject("Range");

    pRange->dynamicCall("MoveStart(QVariant,QVariant)", 10, nStartRow - 1);
    pRange->dynamicCall("MoveEnd(QVariant,QVariant)", 10, nEndRow - nRowCount);
    pRange->dynamicCall("MoveStart(QVariant,QVariant)", 12, nStartCol - 1);
    pRange->dynamicCall("MoveEnd(QVariant,QVariant)", 12, nEndCol - nColCount);

    pRange->dynamicCall("Select(void)");

    QAxObject *pShading = m_pSelection->querySubObject("Shading");
    pShading->dynamicCall("SetTexture(int)", 1000);
    pShading->dynamicCall("SetBackgroundPatternColor(QVariant)", QColor(255, 255, 255));
    pShading->dynamicCall("SetForegroundPatternColor(QVariant)", color);
}

void QDocxWordBackend::setCellColor(const int &nTableIndex, const int &nRow, const int &nCol, const QColor &color)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)", nRow, nCol);
    pCell->dynamicCall("Select(void)");

    QAxObject *pShading = m_pSelection->querySubObject("Shading");
    pShading->dynamicCall("SetTexture(int)", 1000);
    pShading->dynamicCall("SetBackgroundPatternColor(QVariant)", QColor(255, 255, 255));
    pShading->dynamicCall("SetForegroundPatternColor(QVariant)", color);
}

void QDocxWordBackend::selectTable(const int &nTableIndex)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    pTable->dynamicCall("Select(void)");
}

void QDocxWordBackend::moveToTableEnd(const int &nTableIndex)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    pTable->dynamicCall("Select(void)");

    m_pSelection->dynamicCall("MoveRight(QVariant,QVariant,QVariant)", 1, 1, 0);
}

void QDocxWordBackend::spanCells(const int &nTableIndex,
                      const int &nStartRow,
                      const int &nStartCol,
                      const int &nEndRow,
                      const int &nEndCol)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCell_1 = pTable->querySubObject("Cell(int,int)", nStartRow, nStartCol);
    QAxObject *pCell_2 = pTable->querySubObject("Cell(int,int)", nEndRow, nEndCol);

    pCell_1->dynamicCall("Merge(QVariant)", pCell_2->asVariant());
    pCell_1->dynamicCall("Select(void)");
    pCell_1->dynamicCall("SetVerticalAlignment(QVariant)", QVariant(QDocxWordBackend::AliginCenter));

    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)", QVariant(QDocxWordBackend::AliginLeft));
}

void QDocxWordBackend::setCellText(const int &nTableIndex, const int &nRow, const int &nCol, const QString &text)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)", nRow, nCol);

    pCell->dynamicCall("Select(void)");
    addText(text);
}

void QDocxWordBackend::setCellTextAlign(const int &nTableIndex, const int &nRow, const int &nCol, QDocxWordBackend::TextAlign align)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)", nRow, nCol);

    pCell->dynamicCall("Select(void)");
    setTextAlign(align);
}

void QDocxWordBackend::setCellTextColor(const int &nTableIndex, const int &nRow, const int &nCol, const QColor &color)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)", nRow, nCol);
    pCell->dynamicCall("Select(void)");

    setTextColor(color);
}

void QDocxWordBackend::setCellFont(const int &nTableIndex,
                        const int &nRow,
                        const int &nCol,
                        QString fontName,
                        float fontSize,
                        bool isBold,
                        bool isItalic,
                        bool isUnderLine)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)", nRow, nCol);
    pCell->dynamicCall("Select(void)");

    QAxObject *pFont = m_pSelection->querySubObject("Font");
    pFont->dynamicCall("SetName(QString)", fontName);
    pFont->dynamicCall("SetSize(float)", fontSize);
    pFont->dynamicCall("SetBold(bool)", isBold);
    pFont->dynamicCall("SetItalic(bool)", isItalic);
    pFont->dynamicCall("SetUnderline(bool)", isUnderLine);

    m_pSelection->dynamicCall("SetFont(QVariant)", pFont->asVariant());
}

void QDocxWordBackend::setTableFont(const int &nTableIndex,
                         QString fontName,
                         float fontSize,
                         bool isBold,
                         bool isItalic,
                         bool isUnderLine)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    pTable->dynamicCall("Select(void)");

    QAxObject *pFont = m_pSelection->querySubObject("Font");
    pFont->dynamicCall("SetName(QString)", fontName);
    pFont->dynamicCall("SetSize(float)", fontSize);
    pFont->dynamicCall("SetBold(bool)", isBold);
    pFont->dynamicCall("SetItalic(bool)", isItalic);
    pFont->dynamicCall("SetUnderline(bool)", isUnderLine);

    m_pSelection->dynamicCall("SetFont(QVariant)", pFont->asVariant());
}

void QDocxWordBackend::setTableTextAlign(const int &nTableIndex, QDocxWordBackend::TextAlign aligin)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);

    pTable->dynamicCall("Select(void)");

    QAxObject *pCells = m_pSelection->querySubObject("Cells");
    pCells->dynamicCall("SetVerticalAlignment(QVariant)", QVariant(aligin));
}

void QDocxWordBackend::setCellPicture(const int &nTableIndex, const int &nRow, const int &nCol, const QString &path)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)", nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)", nRow, nCol);

    pCell->dynamicCall("Select(void)");
    pCell->dynamicCall("SetVerticalAlignment(QVariant)", QVariant(QDocxWordBackend::AliginCenter));
    pCell->dynamicCall("Select(void)");

    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)", QVariant(QDocxWordBackend::AliginCenter));

    QAxObject *pRange = pCell->querySubObject("Range");
    QAxObject *pInlineShapes = pRange->querySubObject("InlineShapes");
    pInlineShapes->dynamicCall("AddPicture(QString)", path);
}

void QDocxWordBackend::releaseDispatch(QAxObject *pObject)
{
    pObject->dynamicCall("ReleaseDispatch()");
}

QAxObject *QDocxWordBackend::getTable(const int &nTableIndex)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    return pTables->querySubObject("Item(int)", nTableIndex);
}

void QDocxWordBackend::deleteObject()
{
    if (m_pTablesOfContents) {
        delete m_pTablesOfContents;
        m_pTablesOfContents = nullptr;
    }
    if (m_pViewActive) {
        delete m_pViewActive;
        m_pViewActive = nullptr;
    }
    if (m_pPane) {
        delete m_pPane;
        m_pPane = nullptr;
    }
    if (m_pWindowActive) {
        delete m_pWindowActive;
        m_pWindowActive = nullptr;
    }
    if (m_pFont) {
        delete m_pFont;
        m_pFont = nullptr;
    }
    if (m_pSelection) {
        delete m_pSelection;
        m_pSelection = nullptr;
    }
    if (m_pSaveDocument) {
        delete m_pSaveDocument;
        m_pSaveDocument = nullptr;
    }
    if (m_pWordApp) {
        delete m_pWordApp;
        m_pWordApp = nullptr;
    }
}
