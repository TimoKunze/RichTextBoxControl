//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTextRangeEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTextRangeEvents interface</em>
///
/// \if UNICODE
///   \sa TextRange, RichTextBoxCtlLibU::_IRichTextRangeEvents
/// \else
///   \sa TextRange, RichTextBoxCtlLibA::_IRichTextRangeEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTextRangeEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTextRangeEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTextRangeEvents