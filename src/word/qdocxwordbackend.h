#ifndef QDOCXWORDBACKEND_H
#define QDOCXWORDBACKEND_H

#include <QColor>
#include <QDocx/qdocxstyle.h>
#include <QString>

class QAxObject;

class QDocxWordBackend
{
public:
    enum TextAlign { AliginLeft = 0, AliginCenter = 1, AliginRight = 2 };
    enum TitleLevel {
        TitleNine = -10,
        TitleEight = -9,
        TitleSeven = -8,
        TitleSix = -7,
        TitleFive = -6,
        TitleFour = -5,
        TitleThree = -4,
        TitleTwo = -3,
        TitleOne = -2
    };
    enum TableFitBehavior { TableFitFixed = 0, TableFitContent = 1, TableFitWindow = 2 };
    enum LineStyle { LineStyleSingle = 1, LineStyleDouble = 7, LineStyleDot = 2 };
    enum LineSpacing {
        LineSpaceSingle = 0,
        LineSpace1pt5 = 1,
        LineSpaceDouble = 2,
        LineSpaceAtLeast = 3,
        LineSpaceExactly = 4,
        LineSpaceMultiple = 5
    };

    QDocxWordBackend();
    ~QDocxWordBackend();

    QDocxWordBackend(const QDocxWordBackend &) = delete;
    QDocxWordBackend &operator=(const QDocxWordBackend &) = delete;

    bool openNewWord(const bool &isShow = bool(true), QDocxOfficeEngine engine = QDocxOfficeEngine::Word);
    bool saveWord(const QString &path);
    void quitWord();

    void newLine(int nLine = 1);
    void changePage();
    void setTitleText(const QString &title, QDocxWordBackend::TitleLevel level);
    void setCenterTitleText(const QString &title);
    void iniParagraphText();
    void setLineSpace(QDocxWordBackend::LineSpacing space);

    void insertPageHead(const QString &text);
    void insertPageNumber();
    void insertCatalogue();
    void updateCatalogue();

    void addText(const QString &text);
    void addNumber_E(const double &dVal);
    void addNumber_Int(const int &nNum);
    void addNumber_Float(const float &fVal);
    void addPic(const QString &path);

    void setFontStyle(const float &fSize = float(12),
                      const bool &isBold = bool(false),
                      bool isItalic = false,
                      bool isUnderLine = false);
    void setTextAlign(QDocxWordBackend::TextAlign textAlign);
    void setTextColor(const QColor &color = QColor(0, 0, 0));
    void setFontName(const QString &name);

    void addTable(const int &nRow,
                  const int &nCol,
                  QDocxWordBackend::TableFitBehavior autoFit = QDocxWordBackend::TableFitFixed);
    void setTableWidth(const int &nTableIndex, const float &fWidth);
    void setTableColWidth(const int &nTableIndex, const int &col, const float &fWidth);
    void setTableRowHeight(const int nTableIndex, const int &row, const float &fHeight);
    void setCellsBorderStyle(const int &nTableIndex,
                             const int &nStartRow,
                             const int &nStartCol,
                             const int &nEndRow,
                             const int &nEndCol,
                             const QDocxWordBackend::LineStyle &top,
                             const QDocxWordBackend::LineStyle &bottom,
                             const QDocxWordBackend::LineStyle &left,
                             const QDocxWordBackend::LineStyle &right);
    void setCellBorderStyle(const int &nTableIndex,
                            const int &nRow,
                            const int &nCol,
                            const QDocxWordBackend::LineStyle &top,
                            const QDocxWordBackend::LineStyle &bottom,
                            const QDocxWordBackend::LineStyle &left,
                            const QDocxWordBackend::LineStyle &right);
    void setCellsColor(const int &nTableIndex,
                       const int &nStartRow,
                       const int &nStartCol,
                       const int &nEndRow,
                       const int &nEndCol,
                       const QColor &color);
    void setCellColor(const int &nTableIndex, const int &nRow, const int &nCol, const QColor &color);
    void selectTable(const int &nTableIndex);
    void moveToTableEnd(const int &nTableIndex);
    void spanCells(const int &nTableIndex,
                   const int &nStartRow,
                   const int &nStartCol,
                   const int &nEndRow,
                   const int &nEndCol);
    void setCellText(const int &nTableIndex, const int &nRow, const int &nCol, const QString &text);
    void setCellTextAlign(const int &nTableIndex,
                          const int &nRow,
                          const int &nCol,
                          QDocxWordBackend::TextAlign align);
    void setCellTextColor(const int &nTableIndex, const int &nRow, const int &nCol, const QColor &color);
    void setCellFont(const int &nTableIndex,
                     const int &nRow,
                     const int &nCol,
                     QString fontName = QStringLiteral("SimSun"),
                     float fontSize = 9,
                     bool isBold = false,
                     bool isItalic = false,
                     bool isUnderLine = false);
    void setTableFont(const int &nTableIndex,
                      QString fontName = QStringLiteral("SimSun"),
                      float fontSize = 9,
                      bool isBold = false,
                      bool isItalic = false,
                      bool isUnderLine = false);
    void setTableTextAlign(const int &nTableIndex, QDocxWordBackend::TextAlign aligin);
    void setCellPicture(const int &nTableIndex, const int &nRow, const int &nCol, const QString &path);

private:
    void releaseDispatch(QAxObject *pObject);
    QAxObject *getTable(const int &nTableIndex);
    void deleteObject();

    QAxObject *m_pWordApp = nullptr;
    QAxObject *m_pSelection = nullptr;
    QAxObject *m_pFont = nullptr;
    QAxObject *m_pSaveDocument = nullptr;

    QAxObject *m_pWindowActive = nullptr;
    QAxObject *m_pViewActive = nullptr;
    QAxObject *m_pPane = nullptr;
    QAxObject *m_pTablesOfContents = nullptr;
    bool m_isOleInitialized = false;
};

#endif // QDOCXWORDBACKEND_H
