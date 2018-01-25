//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTableCellEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTableCellEvents interface</em>
///
/// \if UNICODE
///   \sa TableCell, RichTextBoxCtlLibU::_IRichTableCellEvents
/// \else
///   \sa TableCell, RichTextBoxCtlLibA::_IRichTableCellEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTableCellEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTableCellEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTableCellEvents