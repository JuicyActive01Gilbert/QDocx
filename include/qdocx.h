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
    enum TextAlign{AliginLeft = 0,AliginCenter = 1,AliginRight = 2};    //������뷽ʽ
    enum TitleLevel{TitleNine = -10,TitleEight = -9,TitleSeven = -8,
                    TitleSix = -7,TitleFive = -6,TitleFour = -5,
                    TitleThree = -4,TitleTwo = -3,TitleOne = -2};       //����ּ�
    enum TableFitBehavior{TableFitFixed = 0,TableFitContent = 1,
                         TableFitWindow = 2};                           //�����Ϸ�ʽ
    enum LineStyle{LineStyleSingle = 1,LineStyleDouble = 7,
                   LineStyleDot = 2};                                   //�߷��

public:
    QDocx();
    ~QDocx();

    /************�������************/
    bool openNewWord(const bool &isShow = bool(true));      //��һ�����ĵ�
    void saveWord(const QString &path);                     //�����ĵ�
    void quitWord();                                        //�˳�word

    void newLine(int nLine = 1);                            //�س����У�Ĭ��Ϊ1��
    void changePage();                                      //��ҳ
    void setTitleText(const QString &title,
                      QDocx::TitleLevel level);             //���ñ��⼰��С
    void setCenterTitleText(const QString &title);          //���þ��б���
    void iniParagraphText();                                //��ʼ���������ָ�ʽΪ����

    /************ҳüҳ��************/
    void insertPageHead(const QString &text);               //����ҳü
    void insertPageNumber();                                //����ҳ��
    /************Ŀ¼���************/
    void insertCatalogue();                                 //����Ŀ¼
    void updateCatalogue();                                 //����Ŀ¼

    /************�������************/
    void addText(const QString &text);                      //���һ������
    void addNumber_E(const double &dVal);                   //�Կ�ѧ��������ʽ���һ������
    void addNumber_Int(const int &nNum);                    //�������
    void addNumber_Float(const float &fVal);                //���С��
    void addPic(const QString &path);                       //���ͼƬ

    /************�������************/
    void setFontStyle(const float &fSize = float(12),
                      const bool &isBold = bool(false),
                      bool isItalic = false,
                      bool isUnderLine = false);
    void setTextAlign(QDocx::TextAlign textAlign);
    void setTextColor(const QColor &color = QColor(0,0,0));

    /************������************/
    void addTable(const int &nRow,
                  const int &nCol,
                  QDocx::TableFitBehavior autoFit = QDocx::TableFitFixed);//�������
    void setTableColWidth(const int &nTableIndex,
                          const int &col,
                          const float &fWidth);             //���ñ���п�
    void setTableRowHeight(const int nTableIndex,
                           const int &row,
                           const float &fHeight);           //���ñ���и�
    void setCellsBorderStyle(const int &nTableIndex,
                             const int &nStartRow,
                             const int &nStartCol,
                             const int &nEndRow,
                             const int &nEndCol,
                             const QDocx::LineStyle &top,
                             const QDocx::LineStyle &bottom,
                             const QDocx::LineStyle &left,
                             const QDocx::LineStyle &right);//����������Ԫ����
    void setCellBorderStyle(const int &nTableIndex,
                            const int &nRow,
                            const int &nCol,
                            const QDocx::LineStyle &top,
                            const QDocx::LineStyle &bottom,
                            const QDocx::LineStyle &left,
                            const QDocx::LineStyle &right); //���õ�����Ԫ����
    void setCellsColor(const int &nTableIndex,
                       const int &nStartRow,
                       const int &nStartCol,
                       const int &nEndRow,
                       const int &nEndCol,
                       const QColor &color);                //�������õ�Ԫ����ɫ
    void setCellColor(const int &nTableIndex,
                       const int &nRow,
                       const int &nCol,
                       const QColor &color);                //���õ�����Ԫ����ɫ
    void selectTable(const int &nTableIndex);               //ѡ��һ�����
    void moveToTableEnd(const int &nTableIndex);            //����ƶ����������
    void spanCells(const int &nTableIndex,
                   const int &nStartRow,
                   const int &nStartCol,
                   const int &nEndRow,
                   const int &nEndCol);                     //�ϲ���Ԫ��
    void setCellText(const int &nTableIndex,
                     const int &nRow,
                     const int &nCol,
                     const QString &text);                  //���õ�Ԫ���ַ�������
    void setCellTextColor(const int &nTableIndex,
                          const int &nRow,
                          const int &nCol,
                          const QColor &color);             //���õ�Ԫ�����ַ�������ɫ
    void setCellFont(const int &nTableIndex,
                     const int &nRow,
                     const int &nCol,
                     QString fontName = QStringLiteral("����"),
                     float fontSize = 9,
                     bool isBold = false,
                     bool isItalic = false,
                     bool isUnderLine = false);             //���õ�Ԫ������
    void setTableFont(const int &nTableIndex,
                      QString fontName = QStringLiteral("����"),
                      float fontSize = 9,
                      bool isBold = false,
                      bool isItalic = false,
                      bool isUnderLine = false);            //���ñ������
    void setTableTextAlign(const int &nTableIndex,
                           QDocx::TextAlign aligin);        //���ñ�����ݶ��뷽ʽ
    void setCellPicture(const int &nTableIndex,
                        const int &nRow,
                        const int &nCol,
                        const QString &path);               //��Ԫ���ڲ���ͼƬ


protected:
    void releaseDispatch(QAxObject *pObject);
    QAxObject* getTable(const int &nTableIndex);
    void deleteObject();                                    //�ͷ����ж���
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
