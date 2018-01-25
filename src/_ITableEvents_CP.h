//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTableEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTableEvents interface</em>
///
/// \if UNICODE
///   \sa Table, RichTextBoxCtlLibU::_IRichTableEvents
/// \else
///   \sa Table, RichTextBoxCtlLibA::_IRichTableEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTableEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTableEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTableEvents