//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTableRowsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTableRowsEvents interface</em>
///
/// \if UNICODE
///   \sa TableRows, RichTextBoxCtlLibU::_IRichTableRowsEvents
/// \else
///   \sa TableRows, RichTextBoxCtlLibA::_IRichTableRowsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTableRowsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTableRowsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTableRowsEvents