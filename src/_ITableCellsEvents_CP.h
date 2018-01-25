//////////////////////////////////////////////////////////////////////
/// \class Proxy_IRichTableCellsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IRichTableCellsEvents interface</em>
///
/// \if UNICODE
///   \sa TableCells, RichTextBoxCtlLibU::_IRichTableCellsEvents
/// \else
///   \sa TableCells, RichTextBoxCtlLibA::_IRichTableCellsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IRichTableCellsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IRichTableCellsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IRichTableCellsEvents