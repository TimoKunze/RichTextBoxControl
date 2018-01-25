//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTextParagraphTabEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTextParagraphTabEvents interface</em>
///
/// \if UNICODE
///   \sa TextParagraphTab, RichTextBoxCtlLibU::_IRichTextParagraphTabEvents
/// \else
///   \sa TextParagraphTab, RichTextBoxCtlLibA::_IRichTextParagraphTabEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTextParagraphTabEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTextParagraphTabEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTextParagraphTabEvents