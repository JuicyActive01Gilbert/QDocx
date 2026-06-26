#include <QDocx/qdocxtable.h>

#include "word/qdocxwordbackend.h"

QDocxCell::QDocxCell() = default;

namespace {

QDocxWordBackend *asBackend(void *backend)
{
    return static_cast<QDocxWordBackend *>(backend);
}

} // namespace

QDocxCell::QDocxCell(void *backend, int tableIndex, int row, int column)
    : m_backend(backend),
      m_tableIndex(tableIndex),
      m_row(row),
      m_column(column)
{
}

bool QDocxCell::isValid() const
{
    return m_backend && m_tableIndex > 0 && m_row > 0 && m_column > 0;
}

int QDocxCell::tableIndex() const
{
    return m_tableIndex;
}

int QDocxCell::row() const
{
    return m_row;
}

int QDocxCell::column() const
{
    return m_column;
}

QDocxCell &QDocxCell::setText(const QString &text)
{
    if (isValid()) {
        asBackend(m_backend)->setCellText(m_tableIndex, m_row, m_column, text);
    }
    return *this;
}

QDocxCell &QDocxCell::setTextColor(const QColor &color)
{
    if (isValid()) {
        asBackend(m_backend)->setCellTextColor(m_tableIndex, m_row, m_column, color);
    }
    return *this;
}

QDocxCell &QDocxCell::setBackgroundColor(const QColor &color)
{
    if (isValid()) {
        asBackend(m_backend)->setCellColor(m_tableIndex, m_row, m_column, color);
    }
    return *this;
}

QDocxCell &QDocxCell::setAlignment(QDocxAlignment alignment)
{
    if (isValid()) {
        asBackend(m_backend)->setCellTextAlign(m_tableIndex,
                                               m_row,
                                               m_column,
                                               static_cast<QDocxWordBackend::TextAlign>(alignment));
    }
    return *this;
}

QDocxCell &QDocxCell::setFont(const QDocxFont &font)
{
    if (isValid()) {
        asBackend(m_backend)->setCellFont(m_tableIndex,
                                          m_row,
                                          m_column,
                                          font.family,
                                          font.pointSize,
                                          font.bold,
                                          font.italic,
                                          font.underline);
        asBackend(m_backend)->setCellTextColor(m_tableIndex, m_row, m_column, font.color);
    }
    return *this;
}

QDocxCell &QDocxCell::addImage(const QString &path)
{
    if (isValid()) {
        asBackend(m_backend)->setCellPicture(m_tableIndex, m_row, m_column, path);
    }
    return *this;
}

QDocxTable::QDocxTable() = default;

QDocxTable::QDocxTable(void *backend, int tableIndex, int rows, int columns)
    : m_backend(backend),
      m_index(tableIndex),
      m_rows(rows),
      m_columns(columns)
{
}

bool QDocxTable::isValid() const
{
    return m_backend && m_index > 0;
}

int QDocxTable::index() const
{
    return m_index;
}

int QDocxTable::rows() const
{
    return m_rows;
}

int QDocxTable::columns() const
{
    return m_columns;
}

QDocxCell QDocxTable::cell(int row, int column) const
{
    if (!isValid() || row <= 0 || column <= 0 || row > m_rows || column > m_columns) {
        return {};
    }
    return {m_backend, m_index, row, column};
}

QDocxTable &QDocxTable::setWidth(float width)
{
    if (isValid()) {
        asBackend(m_backend)->setTableWidth(m_index, width);
    }
    return *this;
}

QDocxTable &QDocxTable::setColumnWidth(int column, float width)
{
    if (isValid()) {
        asBackend(m_backend)->setTableColWidth(m_index, column, width);
    }
    return *this;
}

QDocxTable &QDocxTable::setRowHeight(int row, float height)
{
    if (isValid()) {
        asBackend(m_backend)->setTableRowHeight(m_index, row, height);
    }
    return *this;
}

QDocxTable &QDocxTable::setFont(const QDocxFont &font)
{
    if (isValid()) {
        asBackend(m_backend)->setTableFont(m_index, font.family, font.pointSize, font.bold, font.italic, font.underline);
    }
    return *this;
}

QDocxTable &QDocxTable::setAlignment(QDocxAlignment alignment)
{
    if (isValid()) {
        asBackend(m_backend)->setTableTextAlign(m_index, static_cast<QDocxWordBackend::TextAlign>(alignment));
    }
    return *this;
}

QDocxTable &QDocxTable::mergeCells(int startRow, int startColumn, int endRow, int endColumn)
{
    if (isValid()) {
        asBackend(m_backend)->spanCells(m_index, startRow, startColumn, endRow, endColumn);
    }
    return *this;
}

QDocxTable &QDocxTable::moveCursorAfter()
{
    if (isValid()) {
        asBackend(m_backend)->moveToTableEnd(m_index);
    }
    return *this;
}
