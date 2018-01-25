//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTextParagraphTabsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTextParagraphTabsEvents interface</em>
///
/// \if UNICODE
///   \sa TextParagraphTabs, RichTextBoxCtlLibU::_IRichTextParagraphTabsEvents
/// \else
///   \sa TextParagraphTabs, RichTextBoxCtlLibA::_IRichTextParagraphTabsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTextParagraphTabsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTextParagraphTabsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTextParagraphTabsEvents