#ifndef QDOCXRESULT_H
#define QDOCXRESULT_H

#include <QDocx/qdocxglobal.h>

#include <QString>

enum class QDocxErrorCode {
    None = 0,
    InvalidState,
    WordStartupFailed,
    SaveFailed,
    InvalidArgument
};

class QDOCX_EXPORT QDocxResult
{
public:
    QDocxResult();
    QDocxResult(QDocxErrorCode code, QString message);

    static QDocxResult ok();
    static QDocxResult fail(QDocxErrorCode code, const QString &message);

    explicit operator bool() const;

    bool isOk() const;
    QDocxErrorCode code() const;
    QString message() const;

private:
    QDocxErrorCode m_code = QDocxErrorCode::None;
    QString m_message;
};

#endif // QDOCXRESULT_H
