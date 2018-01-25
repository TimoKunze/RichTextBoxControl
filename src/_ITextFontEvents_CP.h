//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTextFontEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTextFontEvents interface</em>
///
/// \if UNICODE
///   \sa TextFont, RichTextBoxCtlLibU::_IRichTextFontEvents
/// \else
///   \sa TextFont, RichTextBoxCtlLibA::_IRichTextFontEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTextFontEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTextFontEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTextFontEvents