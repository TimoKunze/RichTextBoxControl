//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTableRowEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTableRowEvents interface</em>
///
/// \if UNICODE
///   \sa TableRow, RichTextBoxCtlLibU::_IRichTableRowEvents
/// \else
///   \sa TableRow, RichTextBoxCtlLibA::_IRichTableRowEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTableRowEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTableRowEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTableRowEvents