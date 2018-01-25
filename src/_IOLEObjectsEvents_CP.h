//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichOLEObjectsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichOLEObjectsEvents interface</em>
///
/// \if UNICODE
///   \sa OLEObjects, RichTextBoxCtlLibU::_IRichOLEObjectsEvents
/// \else
///   \sa OLEObjects, RichTextBoxCtlLibA::_IRichOLEObjectsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichOLEObjectsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichOLEObjectsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IOLEObjectsEvents