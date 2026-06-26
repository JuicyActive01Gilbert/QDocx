#include <QDocx/qdocxresult.h>

#include <utility>

QDocxResult::QDocxResult() = default;

QDocxResult::QDocxResult(QDocxErrorCode code, QString message)
    : m_code(code),
      m_message(std::move(message))
{
}

QDocxResult QDocxResult::ok()
{
    return {};
}

QDocxResult QDocxResult::fail(QDocxErrorCode code, const QString &message)
{
    return {code, message};
}

QDocxResult::operator bool() const
{
    return isOk();
}

bool QDocxResult::isOk() const
{
    return m_code == QDocxErrorCode::None;
}

QDocxErrorCode QDocxResult::code() const
{
    return m_code;
}

QString QDocxResult::message() const
{
    return m_message;
}
