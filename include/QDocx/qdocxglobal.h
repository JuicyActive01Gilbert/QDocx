#ifndef QDOCXGLOBAL_H
#define QDOCXGLOBAL_H

#include <QtCore/qglobal.h>

#ifdef QDOCX_LIBRARY
#  define QDOCX_EXPORT Q_DECL_EXPORT
#else
#  define QDOCX_EXPORT Q_DECL_IMPORT
#endif

#endif // QDOCXGLOBAL_H
