#ifndef QDOCXSTYLE_H
#define QDOCXSTYLE_H

#include <QDocx/qdocxglobal.h>

#include <QColor>
#include <QString>

enum class QDocxAlignment {
    Left = 0,
    Center = 1,
    Right = 2
};

enum class QDocxHeadingLevel {
    Level1 = -2,
    Level2 = -3,
    Level3 = -4,
    Level4 = -5,
    Level5 = -6,
    Level6 = -7,
    Level7 = -8,
    Level8 = -9,
    Level9 = -10
};

enum class QDocxLineSpacing {
    Single = 0,
    OnePointFive = 1,
    Double = 2,
    AtLeast = 3,
    Exactly = 4,
    Multiple = 5
};

enum class QDocxOfficeEngine {
    Word = 0,
    Wps = 1
};

struct QDOCX_EXPORT QDocxOpenOptions
{
    bool visible = true;
    QDocxOfficeEngine engine = QDocxOfficeEngine::Word;
};

struct QDOCX_EXPORT QDocxFont
{
    QString family = QStringLiteral("Microsoft YaHei");
    float pointSize = 10.0f;
    bool bold = false;
    bool italic = false;
    bool underline = false;
    QColor color = QColor(0, 0, 0);
};

struct QDOCX_EXPORT QDocxImageOptions
{
    QDocxAlignment alignment = QDocxAlignment::Center;
};

#endif // QDOCXSTYLE_H
