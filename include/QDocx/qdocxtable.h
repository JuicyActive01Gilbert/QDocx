#ifndef QDOCXTABLE_H
#define QDOCXTABLE_H

#include <QDocx/qdocxglobal.h>
#include <QDocx/qdocxstyle.h>

#include <QColor>
#include <QString>

class QDocxDocument;

class QDOCX_EXPORT QDocxCell
{
public:
    QDocxCell();

    bool isValid() const;
    int tableIndex() const;
    int row() const;
    int column() const;

    QDocxCell &setText(const QString &text);
    QDocxCell &setTextColor(const QColor &color);
    QDocxCell &setBackgroundColor(const QColor &color);
    QDocxCell &setAlignment(QDocxAlignment alignment);
    QDocxCell &setFont(const QDocxFont &font);
    QDocxCell &addImage(const QString &path);

private:
    friend class QDocxTable;

    QDocxCell(void *backend, int tableIndex, int row, int column);

    void *m_backend = nullptr;
    int m_tableIndex = 0;
    int m_row = 0;
    int m_column = 0;
};

class QDOCX_EXPORT QDocxTable
{
public:
    QDocxTable();

    bool isValid() const;
    int index() const;
    int rows() const;
    int columns() const;

    QDocxCell cell(int row, int column) const;

    QDocxTable &setWidth(float width);
    QDocxTable &setColumnWidth(int column, float width);
    QDocxTable &setRowHeight(int row, float height);
    QDocxTable &setFont(const QDocxFont &font);
    QDocxTable &setAlignment(QDocxAlignment alignment);
    QDocxTable &mergeCells(int startRow, int startColumn, int endRow, int endColumn);
    QDocxTable &moveCursorAfter();

private:
    friend class QDocxDocument;

    QDocxTable(void *backend, int tableIndex, int rows, int columns);

    void *m_backend = nullptr;
    int m_index = 0;
    int m_rows = 0;
    int m_columns = 0;
};

#endif // QDOCXTABLE_H
