//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTextParagraphEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTextParagraphEvents interface</em>
///
/// \if UNICODE
///   \sa TextParagraph, RichTextBoxCtlLibU::_IRichTextParagraphEvents
/// \else
///   \sa TextParagraph, RichTextBoxCtlLibA::_IRichTextParagraphEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTextParagraphEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTextParagraphEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTextParagraphEvents