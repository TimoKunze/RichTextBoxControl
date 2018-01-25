//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTextSubRangesEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTextSubRangesEvents interface</em>
///
/// \if UNICODE
///   \sa TextSubRanges, RichTextBoxCtlLibU::_IRichTextSubRangesEvents
/// \else
///   \sa TextSubRanges, RichTextBoxCtlLibA::_IRichTextSubRangesEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTextSubRangesEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTextSubRangesEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTextSubRangesEvents