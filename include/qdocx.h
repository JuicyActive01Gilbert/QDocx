/*****************************************************************************
*  @file     qdocx.h                                                         *
*  @brief    Qt operates Office Word documents                               *
*  Details.                                                                  *
*                                                                            *
*  @author   uuMinds-Gilbert                                                 *
*  @email                                                                    *
*  @version  1.0.0                                                           *
*  @date     2018-8-14                                                       *
*  @license  GNU General Public License (GPL)                                *
*                                                                            *
* ****************************************************************************/
#ifndef QDOCX_H
#define QDOCX_H

#include <QAxObject>
#include <QAxWidget>
#include <QObject>
#include <QColor>

class QDocx
{
public:
    enum TextAlign{AliginLeft = 0,AliginCenter = 1,AliginRight = 2};    //字体对齐方式
    enum TitleLevel{TitleNine = -10,TitleEight = -9,TitleSeven = -8,
                    TitleSix = -7,TitleFive = -6,TitleFour = -5,
                    TitleThree = -4,TitleTwo = -3,TitleOne = -2};       //标题分级
    enum TableFitBehavior{TableFitFixed = 0,TableFitContent = 1,
                         TableFitWindow = 2};                           //表格拟合方式
    enum LineStyle{LineStyleSingle = 1,LineStyleDouble = 7,
                   LineStyleDot = 2};                                   //线风格

public:
    QDocx();
    ~QDocx();

    /************功能相关************/
    bool openNewWord(const bool &isShow = bool(true));      //打开一个新文档
    void saveWord(const QString &path);                     //保存文档
    void quitWord();                                        //退出word

    void newLine(int nLine = 1);                            //回车换行，默认为1行
    void changePage();                                      //换页
    void setTitleText(const QString &title,
                      QDocx::TitleLevel level);             //设置标题及大小
    void setCenterTitleText(const QString &title);          //设置居中标题
    void iniParagraphText();                                //初始化段落文字格式为正文

    /************页眉页脚************/
    void insertPageHead(const QString &text);               //插入页眉
    void insertPageNumber();                                //插入页码
    /************目录相关************/
    void insertCatalogue();                                 //插入目录
    void updateCatalogue();                                 //更新目录

    /************添加内容************/
    void addText(const QString &text);                      //添加一段文字
    void addNumber_E(const double &dVal);                   //以科学计数法方式添加一个数字
    void addNumber_Int(const int &nNum);                    //添加整数
    void addNumber_Float(const float &fVal);                //添加小数
    void addPic(const QString &path);                       //添加图片

    /************字体相关************/
    void setFontStyle(const float &fSize = float(12),
                      const bool &isBold = bool(false),
                      bool isItalic = false,
                      bool isUnderLine = false);
    void setTextAlign(QDocx::TextAlign textAlign);
    void setTextColor(const QColor &color = QColor(0,0,0));

    /************表格相关************/
    void addTable(const int &nRow,
                  const int &nCol,
                  QDocx::TableFitBehavior autoFit = QDocx::TableFitFixed);//新增表格
    void setTableColWidth(const int &nTableIndex,
                          const int &col,
                          const float &fWidth);             //设置表格列宽
    void setTableRowHeight(const int nTableIndex,
                           const int &row,
                           const float &fHeight);           //设置表格行高
    void setCellsBorderStyle(const int &nTableIndex,
                             const int &nStartRow,
                             const int &nStartCol,
                             const int &nEndRow,
                             const int &nEndCol,
                             const QDocx::LineStyle &top,
                             const QDocx::LineStyle &bottom,
                             const QDocx::LineStyle &left,
                             const QDocx::LineStyle &right);//设置批量单元格风格
    void setCellBorderStyle(const int &nTableIndex,
                            const int &nRow,
                            const int &nCol,
                            const QDocx::LineStyle &top,
                            const QDocx::LineStyle &bottom,
                            const QDocx::LineStyle &left,
                            const QDocx::LineStyle &right); //设置单个单元格风格
    void setCellsColor(const int &nTableIndex,
                       const int &nStartRow,
                       const int &nStartCol,
                       const int &nEndRow,
                       const int &nEndCol,
                       const QColor &color);                //批量设置单元格颜色
    void setCellColor(const int &nTableIndex,
                       const int &nRow,
                       const int &nCol,
                       const QColor &color);                //设置单个单元格颜色
    void selectTable(const int &nTableIndex);               //选择一个表格
    void moveToTableEnd(const int &nTableIndex);            //光标移动到表格下面
    void spanCells(const int &nTableIndex,
                   const int &nStartRow,
                   const int &nStartCol,
                   const int &nEndRow,
                   const int &nEndCol);                     //合并单元格
    void setCellText(const int &nTableIndex,
                     const int &nRow,
                     const int &nCol,
                     const QString &text);                  //设置单元格字符串内容
    void setCellTextColor(const int &nTableIndex,
                          const int &nRow,
                          const int &nCol,
                          const QColor &color);             //设置单元格中字符串的颜色
    void setCellFont(const int &nTableIndex,
                     const int &nRow,
                     const int &nCol,
                     QString fontName = QStringLiteral("仿宋"),
                     float fontSize = 9,
                     bool isBold = false,
                     bool isItalic = false,
                     bool isUnderLine = false);             //设置单元格字体
    void setTableFont(const int &nTableIndex,
                      QString fontName = QStringLiteral("仿宋"),
                      float fontSize = 9,
                      bool isBold = false,
                      bool isItalic = false,
                      bool isUnderLine = false);            //设置表格字体
    void setTableTextAlign(const int &nTableIndex,
                           QDocx::TextAlign aligin);        //设置表格内容对齐方式
    void setCellPicture(const int &nTableIndex,
                        const int &nRow,
                        const int &nCol,
                        const QString &path);               //向单元格内插入图片


protected:
    void releaseDispatch(QAxObject *pObject);
    QAxObject* getTable(const int &nTableIndex);
    void deleteObject();                                    //释放所有对象
private:
    QAxObject *m_pWordApp = nullptr;
    QAxObject *m_pSelection = nullptr;
    QAxObject *m_pFont = nullptr;
    QAxObject *m_pSaveDocument = nullptr;

    QAxObject *m_pWindowActive = nullptr;
    QAxObject *m_pViewActive = nullptr;
    QAxObject *m_pPane = nullptr;
    QAxObject *m_pTablesOfContents = nullptr;
};

#endif // QDOCX_H
