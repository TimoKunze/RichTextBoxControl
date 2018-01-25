//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichOLEObjectEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichOLEObjectEvents interface</em>
///
/// \if UNICODE
///   \sa OLEObject, RichTextBoxCtlLibU::_IRichOLEObjectEvents
/// \else
///   \sa OLEObject, RichTextBoxCtlLibA::_IRichOLEObjectEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichOLEObjectEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichOLEObjectEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IOLEObjectEvents