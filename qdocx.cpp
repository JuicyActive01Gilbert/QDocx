#include "qdocx.h"
#include <windows.h>
#include <QDebug>


QDocx::QDocx()
{

}

QDocx::~QDocx()
{
    OleUninitialize();
}

bool QDocx::openNewWord(const bool &isShow)
{
    HRESULT r = OleInitialize(0);
    if (r != S_OK && r != S_FALSE)
    {
        qCritical() << QStringLiteral("初始化失败");
        return false;
    }

    m_pWordApp = new QAxObject("Word.Application");
    if(m_pWordApp->isNull()){
        qCritical() << QStringLiteral("Word 启动失败，请检查是否已经安装Office Word!");
        return false;
    }
    m_pWordApp->setProperty("Visible",isShow);

    QAxObject * pDocuments = m_pWordApp->querySubObject("Documents");
    if(pDocuments->isNull()){
        return false;
    }
    pDocuments->dynamicCall("Add(QString)","");
    m_pSaveDocument = m_pWordApp->querySubObject("ActiveDocument");
    if(m_pSaveDocument->isNull()){
        return false;
    }

    m_pSelection = m_pWordApp->querySubObject("Selection");
    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetLineSpacingRule(QVariant)",LineSpace1pt5);//1代表1.5倍行间距

    return true;
}

void QDocx::saveWord(const QString &path)
{
    m_pSaveDocument->dynamicCall("SaveAs(QString)",path);
}

void QDocx::quitWord()
{
    if(m_pWordApp){
        m_pWordApp->setProperty("DisplayAlerts", true);
    }
    if(m_pSaveDocument){
        m_pSaveDocument->dynamicCall("Close(bool)", true);
    }
    if(m_pWordApp){
        m_pWordApp->dynamicCall("Quit()");
    }

    deleteObject();
}

void QDocx::newLine(int nLine)
{
    if(nLine <= 0){
        return;
    }

    for(int i = 0; i < nLine;++i){
        m_pSelection->dynamicCall("TypeParagraph(void)");
    }

}

void QDocx::changePage()
{
    m_pSelection->dynamicCall("InsertBreak(QVariant)",7);
}

void QDocx::setTitleText(const QString &title,
                         QDocx::TitleLevel level)
{
    QAxObject *pSelection = m_pWordApp->querySubObject("Selection");
    pSelection->dynamicCall("SetStyle(QVariant)",static_cast<long>(level));
    pSelection->dynamicCall("TypeText(QString)",title);

    newLine();

    iniParagraphText();
}

void QDocx::setCenterTitleText(const QString &title)
{
    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)",
                                  QDocx::AliginCenter);
    m_pSelection->dynamicCall("TypeText(QString)",title);
}

void QDocx::iniParagraphText()
{
    //设置段落格式为正文
    QAxObject *pSelection = m_pWordApp->querySubObject("Selection");
    QAxObject *pParagraphFormat = pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetStyle(QVariant)",-67);
}

void QDocx::setLineSpace(QDocx::LineSpacing space)
{
    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetLineSpacingRule(QVariant)",space);//1代表1.5倍行间距
}

void QDocx::insertPageHead(const QString &text)
{
    m_pWindowActive = m_pWordApp->querySubObject("ActiveWindow");
    m_pPane = m_pWindowActive->querySubObject("ActivePane");
    m_pViewActive = m_pPane->querySubObject("View");
    m_pViewActive->dynamicCall("SetSeekView(QVariant)",9);

    addText(text);

    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)",
                                  QDocx::AliginCenter);

    m_pViewActive->dynamicCall("SetSeekView(QVariant)",0);
}

void QDocx::insertPageNumber()
{
    QAxObject *pFields = m_pSelection->querySubObject("Fields");
    m_pWindowActive = m_pWordApp->querySubObject("ActiveWindow");
    m_pPane = m_pWindowActive->querySubObject("ActivePane");
    m_pViewActive = m_pPane->querySubObject("View");
    m_pViewActive->dynamicCall("SetSeekView(QVariant)",10);

    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)",
                                  QDocx::AliginCenter);
    addText("- ");
    pFields->dynamicCall("Add(QVariant,QVariant,QVariant,QVariant)",
                         m_pSelection->querySubObject("Range")->asVariant(),
                         33,"PAGE  ",1);
    addText(" -");
    m_pSelection->dynamicCall("TypeParagraph()");
    m_pViewActive->dynamicCall("SetSeekView(QVariant)",0);

}

void QDocx::insertCatalogue()
{
    setFontStyle(18,true); //正文字体
    addText(QStringLiteral("                    目 录\n\n"));
    m_pTablesOfContents = m_pSaveDocument->querySubObject("TablesOfContents");

    QAxObject *pRange=m_pSelection->querySubObject("Range");
    m_pTablesOfContents->dynamicCall("Add(QVariant)",pRange->asVariant());
}

void QDocx::updateCatalogue()
{
    int count = m_pTablesOfContents->property("Count").toInt();
    if (count > 0){
        QAxObject *pTableOfContent = m_pTablesOfContents->querySubObject("Item(int)"
                                                                         ,count);
        pTableOfContent->dynamicCall("UpdatePageNumbers(void)");
    }
}

void QDocx::addText(const QString &text)
{
    m_pSelection->dynamicCall("TypeText(QString)",text);
}

void QDocx::addNumber_E(const double &dVal)
{
    QString strText = QString::number(dVal,'e');
    m_pSelection->dynamicCall("TypeText(QString)",strText);
}

void QDocx::addNumber_Int(const int &nNum)
{
    QString strText = QString::number(nNum);
    m_pSelection->dynamicCall("TypeText(QString)",strText);
}

void QDocx::addNumber_Float(const float &fVal)
{
    QString strText = QString::number(fVal,'f');
    m_pSelection->dynamicCall("TypeText(QString)",strText);
}

void QDocx::addPic(const QString &path)
{
    QAxObject *pSelection = m_pWordApp->querySubObject("Selection");
    QAxObject *pParagraphFormat = pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetLineSpacingRule(QVariant)",0);
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)",
                                  QDocx::AliginCenter);
    QAxObject *pInlineShapes = pSelection->querySubObject("InlineShapes");
    pInlineShapes->dynamicCall("AddPicture(const &QString)",path);

    newLine();
    pParagraphFormat = pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)",
                                  QDocx::AliginCenter);

}

void QDocx::setFontStyle(const float &fSize,
                         const bool &isBold,
                         bool isItalic,
                         bool isUnderLine)
{
    if(m_pSelection->isNull()){
        qWarning() << QStringLiteral("Selection为空，字体设置失败!");
        return;
    }
    m_pSelection->dynamicCall("SetText(QString)","Font");

    m_pFont = m_pSelection->querySubObject("Font");
    m_pFont->dynamicCall("SetSize(float)",fSize);
    m_pFont->dynamicCall("SetBold(bool)",isBold);
    m_pFont->dynamicCall("SetItalic(bool)",isItalic);
    m_pFont->dynamicCall("SetUnderline(bool)",isUnderLine);

    m_pSelection->dynamicCall("SetFont(QVariant)",m_pFont->asVariant());
}

void QDocx::setTextAlign(QDocx::TextAlign textAlign)
{
    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(int)",static_cast<int>(textAlign));
}

void QDocx::setTextColor(const QColor &color)
{
    QAxObject *pFont = m_pSelection->querySubObject("Font");
    pFont->dynamicCall("SetColor(QVariant)",color);//设置颜色参数使用QVariant即可
}

void QDocx::setFontName(const QString &name)
{
    QAxObject *pFont = m_pSelection->querySubObject("Font");
    pFont->dynamicCall("SetName(QString)",name);
}

void QDocx::addTable(const int &nRow,
                     const int &nCol,
                     QDocx::TableFitBehavior autoFit)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pRange = m_pSelection->querySubObject("Range");

    pTables->dynamicCall("Add(QVariant,int,int,QVariant,QVariant)",
                         pRange->asVariant(),nRow,nCol,
                         1,static_cast<int>(autoFit));

}

void QDocx::setTableWidth(const int &nTableIndex, const float &fWidth)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    pTable->setProperty("PreferredWidthType",2);
    pTable->setProperty("PreferredWidth",fWidth);
}

void QDocx::setTableColWidth(const int &nTableIndex,
                             const int &col,
                             const float &fWidth)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCols = pTable->querySubObject("Columns");
    QAxObject *pCol = pCols->querySubObject("Item(int)",col);
    pCol->setProperty("PreferredWidthType",2);
    pCol->setProperty("PreferredWidth",fWidth);
}

void QDocx::setTableRowHeight(const int nTableIndex,
                              const int &row,
                              const float &fHeight)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pRows = pTable->querySubObject("Rows");
    QAxObject *pRow = pRows->querySubObject("Item(int)",row);
    pRow->setProperty("Height",fHeight);
}

void QDocx::setCellsBorderStyle(const int &nTableIndex,
                                const int &nStartRow,
                                const int &nStartCol,
                                const int &nEndRow,
                                const int &nEndCol,
                                const QDocx::LineStyle &top,
                                const QDocx::LineStyle &bottom,
                                const QDocx::LineStyle &left,
                                const QDocx::LineStyle &right)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pRows = pTable->querySubObject("Rows");
    QAxObject *pCols = pTable->querySubObject("Columns");

    int nRowCount = pRows->property("Count").toInt();
    int nColCount = pCols->property("Count").toInt();

    pTable->dynamicCall("Select(void)");
    QAxObject *pRange = m_pSelection->querySubObject("Range");

    //行
    pRange->dynamicCall("MoveStart(QVariant,QVariant)",10,nStartRow - 1);
    pRange->dynamicCall("MoveEnd(QVariant,QVariant)",10,nEndRow - nRowCount);

    //列
    pRange->dynamicCall("MoveStart(QVariant,QVariant)",12,nStartCol - 1);
    pRange->dynamicCall("MoveEnd(QVariant,QVariant)",12,nEndCol - nColCount);

    pRange->dynamicCall("Select(void)");

    QAxObject *pBorders = m_pSelection->querySubObject("Borders");
    pBorders->querySubObject("Item(int)",1)
            ->dynamicCall("SetLineStyle(int)",static_cast<int>(top));
    pBorders->querySubObject("Item(int)",2)
            ->dynamicCall("SetLineStyle(int)",static_cast<int>(left));
    pBorders->querySubObject("Item(int)",3)
            ->dynamicCall("SetLineStyle(int)",static_cast<int>(bottom));
    pBorders->querySubObject("Item(int)",4)
            ->dynamicCall("SetLineStyle(int)",static_cast<int>(right));

}

void QDocx::setCellBorderStyle(const int &nTableIndex,
                               const int &nRow,
                               const int &nCol,
                               const QDocx::LineStyle &top,
                               const QDocx::LineStyle &bottom,
                               const QDocx::LineStyle &left,
                               const QDocx::LineStyle &right)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)",nRow,nCol);
    pCell->dynamicCall("Select(void)");

    QAxObject *pBorders = pCell->querySubObject("Borders");
    pBorders->querySubObject("Item(int)",1)
            ->dynamicCall("SetLineStyle(int)",static_cast<int>(top));
    pBorders->querySubObject("Item(int)",2)
            ->dynamicCall("SetLineStyle(int)",static_cast<int>(left));
    pBorders->querySubObject("Item(int)",3)
            ->dynamicCall("SetLineStyle(int)",static_cast<int>(bottom));
    pBorders->querySubObject("Item(int)",4)
            ->dynamicCall("SetLineStyle(int)",static_cast<int>(right));
}

void QDocx::setCellsColor(const int &nTableIndex,
                          const int &nStartRow,
                          const int &nStartCol,
                          const int &nEndRow,
                          const int &nEndCol,
                          const QColor &color)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pRows = pTable->querySubObject("Rows");
    QAxObject *pCols = pTable->querySubObject("Columns");

    int nRowCount = pRows->property("Count").toInt();
    int nColCount = pCols->property("Count").toInt();

    pTable->dynamicCall("Select(void)");
    QAxObject *pRange = m_pSelection->querySubObject("Range");

    //行
    pRange->dynamicCall("MoveStart(QVariant,QVariant)",10,nStartRow - 1);
    pRange->dynamicCall("MoveEnd(QVariant,QVariant)",10,nEndRow - nRowCount);

    //列
    pRange->dynamicCall("MoveStart(QVariant,QVariant)",12,nStartCol - 1);
    pRange->dynamicCall("MoveEnd(QVariant,QVariant)",12,nEndCol - nColCount);

    pRange->dynamicCall("Select(void)");

    QAxObject *pShading = m_pSelection->querySubObject("Shading");
    pShading->dynamicCall("SetTexture(int)",1000);
    pShading->dynamicCall("SetBackgroundPatternColor(QVariant)",QColor(255,255,255));
    pShading->dynamicCall("SetForegroundPatternColor(QVariant)",color);
}

void QDocx::setCellColor(const int &nTableIndex,
                         const int &nRow,
                         const int &nCol,
                         const QColor &color)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)",nRow,nCol);
    pCell->dynamicCall("Select(void)");

    QAxObject *pShading = m_pSelection->querySubObject("Shading");
    pShading->dynamicCall("SetTexture(int)",1000);
    pShading->dynamicCall("SetBackgroundPatternColor(QVariant)",QColor(255,255,255));
    pShading->dynamicCall("SetForegroundPatternColor(QVariant)",color);
}

void QDocx::selectTable(const int &nTableIndex)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    pTable->dynamicCall("Select(void)");
}

void QDocx::moveToTableEnd(const int &nTableIndex)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    pTable->dynamicCall("Select(void)");

    m_pSelection->dynamicCall("MoveRight(QVariant,QVariant,QVariant)",1,1,0);
}

void QDocx::spanCells(const int &nTableIndex,
                      const int &nStartRow,
                      const int &nStartCol,
                      const int &nEndRow,
                      const int &nEndCol)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCell_1 = pTable->querySubObject("Cell(int,int)",nStartRow,nStartCol);
    QAxObject *pCell_2 = pTable->querySubObject("Cell(int,int)",nEndRow,nEndCol);

    pCell_1->dynamicCall("Merge(QVariant)",pCell_2->asVariant());
    pCell_1->dynamicCall("Select(void)");
    pCell_1->dynamicCall("SetVerticalAlignment(QVariant)",QDocx::AliginCenter);

    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)",QDocx::AliginLeft);
}

void QDocx::setCellText(const int &nTableIndex,
                        const int &nRow,
                        const int &nCol,
                        const QString &text)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)",nRow,nCol);

    pCell->dynamicCall("Select(void)");
    addText(text);
}

void QDocx::setCellTextAlign(const int &nTableIndex,
                             const int &nRow,
                             const int &nCol,
                             QDocx::TextAlign align)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)",nRow,nCol);

    pCell->dynamicCall("Select(void)");
    setTextAlign(align);
}

void QDocx::setCellTextColor(const int &nTableIndex,
                             const int &nRow,
                             const int &nCol,
                             const QColor &color)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)",nRow,nCol);
    pCell->dynamicCall("Select(void)");

    setTextColor(color);
}

void QDocx::setCellFont(const int &nTableIndex,
                        const int &nRow,
                        const int &nCol,
                        QString fontName,
                        float fontSize,
                        bool isBold,
                        bool isItalic,
                        bool isUnderLine)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)",nRow,nCol);
    pCell->dynamicCall("Select(void)");

    QAxObject *pFont = m_pSelection->querySubObject("Font");
    pFont->dynamicCall("SetName(QString)",fontName);
    pFont->dynamicCall("SetSize(float)",fontSize);
    pFont->dynamicCall("SetBold(bool)",isBold);
    pFont->dynamicCall("SetItalic(bool)",isItalic);
    pFont->dynamicCall("SetUnderline(bool)",isUnderLine);

    m_pSelection->dynamicCall("SetFont(QVariant)",pFont->asVariant());
}

void QDocx::setTableFont(const int &nTableIndex,
                         QString fontName,
                         float fontSize,
                         bool isBold,
                         bool isItalic,
                         bool isUnderLine)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    pTable->dynamicCall("Select(void)");

    QAxObject *pFont = m_pSelection->querySubObject("Font");
    pFont->dynamicCall("SetName(QString)",fontName);
    pFont->dynamicCall("SetSize(float)",fontSize);
    pFont->dynamicCall("SetBold(bool)",isBold);
    pFont->dynamicCall("SetItalic(bool)",isItalic);
    pFont->dynamicCall("SetUnderline(bool)",isUnderLine);

    m_pSelection->dynamicCall("SetFont(QVariant)",pFont->asVariant());
}

void QDocx::setTableTextAlign(const int &nTableIndex,
                              QDocx::TextAlign aligin)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);

    pTable->dynamicCall("Select(void)");

    QAxObject *pCells = m_pSelection->querySubObject("Cells");
    pCells->dynamicCall("SetVerticalAlignment(QVariant)",aligin);
}

void QDocx::setCellPicture(const int &nTableIndex,
                           const int &nRow,
                           const int &nCol,
                           const QString &path)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    QAxObject *pTable = pTables->querySubObject("Item(int)",nTableIndex);
    QAxObject *pCell = pTable->querySubObject("Cell(int,int)",nRow,nCol);

    pCell->dynamicCall("Select(void)");
    pCell->dynamicCall("SetVerticalAlignment(QVariant)",QDocx::AliginCenter);
    pCell->dynamicCall("Select(void)");

    QAxObject *pParagraphFormat = m_pSelection->querySubObject("ParagraphFormat");
    pParagraphFormat->dynamicCall("SetAlignment(QVariant)",QDocx::AliginCenter);

    QAxObject *pRange = pCell->querySubObject("Range");
    QAxObject *pInlineShapes = pRange->querySubObject("InlineShapes");
    pInlineShapes->dynamicCall("AddPicture(QString)",path);
}

void QDocx::releaseDispatch(QAxObject *pObject)
{
    pObject->dynamicCall("ReleaseDispatch()");
}

QAxObject *QDocx::getTable(const int &nTableIndex)
{
    QAxObject *pTables = m_pSaveDocument->querySubObject("Tables");
    return pTables->querySubObject("Item(int)",nTableIndex);
}

void QDocx::deleteObject()
{
    if(m_pWordApp){
        m_pWordApp->deleteLater();
        m_pWordApp = nullptr;
    }
    if(m_pSelection){
        m_pSelection->deleteLater();
        m_pSelection = nullptr;
    }
    if(m_pFont){
        m_pFont->deleteLater();
        m_pFont = nullptr;
    }
    if(m_pSaveDocument){
        m_pSaveDocument->deleteLater();
        m_pFont = nullptr;
    }

    if(m_pWindowActive){
        m_pWindowActive->deleteLater();
        m_pWindowActive = nullptr;
    }
    if(m_pViewActive){
        m_pViewActive->deleteLater();
        m_pViewActive = nullptr;
    }
    if(m_pPane){
        m_pPane->deleteLater();
        m_pPane = nullptr;
    }
}
