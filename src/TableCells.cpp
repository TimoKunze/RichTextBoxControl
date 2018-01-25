// TableCells.cpp: Manages a collection of TableCell objects

#include "stdafx.h"
#include "TableCells.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TableCells::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTableCells, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TableCells::QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRange), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextRange2), queriedInterface)) {
		TableCells* pTableCells = reinterpret_cast<TableCells*>(pThis);
		return pTableCells->properties.pRowTextRange->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}

HRESULT CALLBACK TableCells::QueryITextRowInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRow), queriedInterface)) {
		TableCells* pTableCells = reinterpret_cast<TableCells*>(pThis);
		RichEditAPIVersionConstants apiVersion = reavUnknown;
		pTableCells->properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);
		if(apiVersion == reav41 || apiVersion >= reav75) {
			CComPtr<ITextRange> pRowTextRange = NULL;
			if(SUCCEEDED(pTableCells->properties.pRowTextRange->GetDuplicate(&pRowTextRange))) {
				CComQIPtr<ITextRange2> pTextRange2 = pRowTextRange;
				if(pTextRange2) {
					LONG rowStart = 0;
					pTextRange2->GetStart(&rowStart);
					pTextRange2->SetEnd(rowStart);
					ITextRow* pTextRow = NULL;
					HRESULT hr = pTextRange2->GetRow(&pTextRow);
					if(SUCCEEDED(hr) && pTextRow) {
						hr = pTextRow->QueryInterface(queriedInterface, ppImplementation);
					}
					if(pTextRow) {
						pTextRow->Release();
					}
					return hr;
				}
			}
		}
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP TableCells::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<TableCells>* pTableCellsObj = NULL;
	CComObject<TableCells>::CreateInstance(&pTableCellsObj);
	pTableCellsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTableCellsObj->properties);

	pTableCellsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTableCellsObj->Release();
	return S_OK;
}

STDMETHODIMP TableCells::Next(ULONG numberOfMaxCells, VARIANT* pCells, ULONG* pNumberOfCellsReturned)
{
	ATLASSERT_POINTER(pCells, VARIANT);
	if(!pCells) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTableTextRange);
	ATLASSERT(properties.pRowTextRange);

	// check each cell in the text range
	ULONG i = 0;
	for(i = 0; i < numberOfMaxCells; ++i) {
		VariantInit(&pCells[i]);

		do {
			if(properties.pLastEnumeratedCellTextRange) {
				ITextRange* pTmp = GetNextCellToProcess(properties.pLastEnumeratedCellTextRange);
				properties.pLastEnumeratedCellTextRange->Release();
				properties.pLastEnumeratedCellTextRange = pTmp;
			} else {
				properties.pLastEnumeratedCellTextRange = GetFirstCellToProcess();
			}
		} while((properties.pLastEnumeratedCellTextRange != NULL) && (!IsPartOfCollection(properties.pLastEnumeratedCellTextRange)));

		if(properties.pLastEnumeratedCellTextRange) {
			CComPtr<ITextRange> pCellTextRange = NULL;
			if(SUCCEEDED(properties.pLastEnumeratedCellTextRange->GetDuplicate(&pCellTextRange))) {
				ClassFactory::InitTableCell(properties.pTableTextRange, properties.pRowTextRange, pCellTextRange, properties.pOwnerRTB, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pCells[i].pdispVal));
				pCells[i].vt = VT_DISPATCH;
			} else {
				i--;
				break;
			}
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfCellsReturned) {
		*pNumberOfCellsReturned = i;
	}

	return (i == numberOfMaxCells ? S_OK : S_FALSE);
}

STDMETHODIMP TableCells::Reset(void)
{
	if(properties.pLastEnumeratedCellTextRange) {
		properties.pLastEnumeratedCellTextRange->Release();
		properties.pLastEnumeratedCellTextRange = NULL;
	}
	return S_OK;
}

STDMETHODIMP TableCells::Skip(ULONG numberOfCellsToSkip)
{
	VARIANT dummy;
	ULONG numCellsReturned = 0;
	// we could skip all cells at once, but it's easier to skip them one by one
	for(ULONG i = 1; i <= numberOfCellsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numCellsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numCellsReturned == 0) {
			// there're no more cells to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////

ITextRange* TableCells::GetFirstCellToProcess(void)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return NULL;
	}

	CComPtr<ITextRange> pCellTextRange = NULL;
	if(SUCCEEDED(properties.pRowTextRange->GetDuplicate(&pCellTextRange))) {
		LONG rowStart = 0;
		LONG rowEnd = 0;
		properties.pRowTextRange->GetStart(&rowStart);
		properties.pRowTextRange->GetEnd(&rowEnd);

		// NOTE: StartOf(tomCell, tomExtend) and EndOf(tomCell, tomExtend) do not really work on Windows 7.

		// start with the first cell
		LONG cellStart = rowStart + 2;
		LONG cellEnd = cellStart;
		LONG startOfNextCellCandidate = 0;
		HRESULT hr = S_OK;
		do {
			// set the cursor to the current cell
			pCellTextRange->SetStart(cellStart);
			pCellTextRange->SetEnd(cellEnd);
			// move one cell ahead
			hr = pCellTextRange->Move(tomCell, 1, NULL);
			if(hr == S_OK) {
				// the start of the next cell minus one is the end of the current cell
				pCellTextRange->GetStart(&startOfNextCellCandidate);
				cellEnd = startOfNextCellCandidate - 1;
				if(startOfNextCellCandidate < rowEnd) {
					// we're still inside the row, so use this cell
					pCellTextRange->SetStart(cellStart);
					pCellTextRange->SetEnd(cellEnd);
					return pCellTextRange.Detach();
				}
				if(startOfNextCellCandidate >= rowEnd) {
					// we reached the end of the row
					break;
				}
			} else {
				// this has not been a cell, so don't count it and stop
				break;
			}
		} while(TRUE);
	}
	return NULL;
}

ITextRange* TableCells::GetNextCellToProcess(ITextRange* pPreviousCell)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return NULL;
	}

	CComPtr<ITextRange> pCellTextRange = NULL;
	if(SUCCEEDED(properties.pRowTextRange->GetDuplicate(&pCellTextRange))) {
		LONG rowEnd = 0;
		properties.pRowTextRange->GetEnd(&rowEnd);
		LONG startOfNextCellCandidate = 0;
		pPreviousCell->GetEnd(&startOfNextCellCandidate);
		if(startOfNextCellCandidate >= rowEnd) {
			// we already reached the end of the row
			return NULL;
		}

		// NOTE: StartOf(tomCell, tomExtend) and EndOf(tomCell, tomExtend) do not really work on Windows 7.

		// start with the first cell
		LONG cellStart = startOfNextCellCandidate + 1;
		LONG cellEnd = cellStart;
		HRESULT hr = S_OK;
		do {
			// set the cursor to the current cell
			pCellTextRange->SetStart(cellStart);
			pCellTextRange->SetEnd(cellEnd);
			// move one cell ahead
			hr = pCellTextRange->Move(tomCell, 1, NULL);
			if(hr == S_OK) {
				// the start of the next cell minus one is the end of the current cell
				pCellTextRange->GetStart(&startOfNextCellCandidate);
				cellEnd = startOfNextCellCandidate - 1;
				// it is important to use "<" here, otherwise a ghost cell will be found in the last row
				if(startOfNextCellCandidate < rowEnd) {
					// we're still inside the row, so use this cell
					pCellTextRange->SetStart(cellStart);
					pCellTextRange->SetEnd(cellEnd);
					return pCellTextRange.Detach();
				}
				if(startOfNextCellCandidate >= rowEnd) {
					// we reached the end of the row
					break;
				}
			} else {
				// this has not been a cell, so don't count it and stop
				break;
			}
		} while(TRUE);
	}
	return NULL;
}

BOOL TableCells::IsPartOfCollection(ITextRange* pCellTextRange)
{
	LONG rowStart = 0;
	LONG rowEnd = 0;
	if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowStart)) && SUCCEEDED(properties.pRowTextRange->GetEnd(&rowEnd))) {
		LONG cellStart = 0;
		LONG cellEnd = 0;
		if(SUCCEEDED(pCellTextRange->GetStart(&cellStart)) && SUCCEEDED(pCellTextRange->GetEnd(&cellEnd))) {
			if(cellStart <= cellEnd && cellStart >= rowStart && cellEnd <= rowEnd) {
				// StartOf and EndOf do not really work for tomCell. The range start is set to 0.
				/*CComPtr<ITextRange> pTextRange = NULL;
				if(SUCCEEDED(pCellTextRange->GetDuplicate(&pTextRange))) {
					HRESULT hr = pTextRange->StartOf(tomCell, tomExtend, NULL);
					hr = pTextRange->EndOf(tomCell, tomExtend, NULL);
					LONG cellStart2 = 0;
					LONG cellEnd2 = 0;
					if(SUCCEEDED(pTextRange->GetStart(&cellStart2)) && SUCCEEDED(pTextRange->GetEnd(&cellEnd2))) {
						if(cellStart == cellStart2 && cellEnd == cellEnd2) {*/
							return TRUE;
						/*}
					}
				}*/
			}
		}
	}
	return FALSE;
}


TableCells::Properties::~Properties()
{
	if(pOwnerRTB) {
		pOwnerRTB->Release();
		pOwnerRTB = NULL;
	}
	if(pTableTextRange) {
		pTableTextRange->Release();
		pTableTextRange = NULL;
	}
	if(pRowTextRange) {
		pRowTextRange->Release();
		pRowTextRange = NULL;
	}
	if(pLastEnumeratedCellTextRange) {
		pLastEnumeratedCellTextRange->Release();
		pLastEnumeratedCellTextRange = NULL;
	}
}

void TableCells::Properties::CopyTo(TableCells::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerRTB = this->pOwnerRTB;
		if(pTarget->pOwnerRTB) {
			pTarget->pOwnerRTB->AddRef();
		}
		pTarget->pTableTextRange = this->pTableTextRange;
		if(pTarget->pTableTextRange) {
			pTarget->pTableTextRange->AddRef();
		}
		pTarget->pRowTextRange = this->pRowTextRange;
		if(pTarget->pRowTextRange) {
			pTarget->pRowTextRange->AddRef();
		}
		pTarget->pLastEnumeratedCellTextRange = NULL;
		if(this->pLastEnumeratedCellTextRange) {
			this->pLastEnumeratedCellTextRange->GetDuplicate(&pTarget->pLastEnumeratedCellTextRange);
		}
	}
}

HWND TableCells::Properties::GetRTBHWnd(void)
{
	ATLASSUME(pOwnerRTB);

	OLE_HANDLE handle = NULL;
	pOwnerRTB->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TableCells::Attach(ITextRange* pTableTextRange, ITextRange* pRowTextRange)
{
	if(properties.pTableTextRange) {
		properties.pTableTextRange->Release();
		properties.pTableTextRange = NULL;
	}
	if(properties.pRowTextRange) {
		properties.pRowTextRange->Release();
		properties.pRowTextRange = NULL;
	}
	if(pTableTextRange) {
		//pTableTextRange->QueryInterface(IID_PPV_ARGS(&properties.pTableTextRange));
		properties.pTableTextRange = pTableTextRange;
		properties.pTableTextRange->AddRef();
	}
	if(pRowTextRange) {
		//pRowTextRange->QueryInterface(IID_PPV_ARGS(&properties.pRowTextRange));
		properties.pRowTextRange = pRowTextRange;
		properties.pRowTextRange->AddRef();
	}
}

void TableCells::Detach(void)
{
	if(properties.pTableTextRange) {
		properties.pTableTextRange->Release();
		properties.pTableTextRange = NULL;
	}
	if(properties.pRowTextRange) {
		properties.pRowTextRange->Release();
		properties.pRowTextRange = NULL;
	}
	if(properties.pLastEnumeratedCellTextRange) {
		properties.pLastEnumeratedCellTextRange->Release();
		properties.pLastEnumeratedCellTextRange = NULL;
	}
}

void TableCells::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP TableCells::get_Item(LONG cellIdentifier, CellIdentifierTypeConstants cellIdentifierType/* = citIndex*/, IRichTableCell** ppCell/* = NULL*/)
{
	ATLASSERT_POINTER(ppCell, IRichTableCell*);
	if(!ppCell) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}

	// retrieve the cell's text range
	CComPtr<ITextRange> pCellTextRange = NULL;
	HRESULT hr = properties.pRowTextRange->GetDuplicate(&pCellTextRange);
	if(FAILED(hr)) {
		return hr;
	}
	LONG rowStart = 0;
	LONG rowEnd = 0;
	properties.pRowTextRange->GetStart(&rowStart);
	properties.pRowTextRange->GetEnd(&rowEnd);

	// NOTE: StartOf(tomCell, tomExtend) and EndOf(tomCell, tomExtend) do not really work on Windows 7.

	// start with the first cell
	int cellIndex = 0;
	BOOL foundCell = FALSE;
	LONG cellStart = rowStart + 2;
	LONG cellEnd = cellStart;
	LONG startOfNextCellCandidate = 0;
	do {
		// set the cursor to the current cell
		pCellTextRange->SetStart(cellStart);
		pCellTextRange->SetEnd(cellEnd);
		// move one cell ahead
		hr = pCellTextRange->Move(tomCell, 1, NULL);
		if(hr == S_OK) {
			// the start of the next cell minus one is the end of the current cell
			pCellTextRange->GetStart(&startOfNextCellCandidate);
			cellEnd = startOfNextCellCandidate - 1;
			if(startOfNextCellCandidate < rowEnd) {
				// we're still inside the row, so count this cell
				cellIndex++;
				switch(cellIdentifierType) {
					case citIndex:
						if(cellIdentifier == cellIndex - 1) {
							foundCell = TRUE;
						}
						break;
					case citCharacterIndex:
					{
						if(cellStart <= cellIdentifier && cellIdentifier <= cellEnd) {
							foundCell = TRUE;
						}
						break;
					}
				}
			}
			if(startOfNextCellCandidate >= rowEnd) {
				// we reached the end of the row
				break;
			}
		} else {
			// this has not been a cell, so don't count it and stop
			break;
		}
		if(!foundCell) {
			// prepare for next cell
			cellStart = startOfNextCellCandidate;
			cellEnd = cellStart;
		}
	} while(!foundCell);

	if(foundCell) {
		pCellTextRange->SetStart(cellStart);
		pCellTextRange->SetEnd(cellEnd);
		if(IsPartOfCollection(pCellTextRange)) {
			ClassFactory::InitTableCell(properties.pTableTextRange, properties.pRowTextRange, pCellTextRange, properties.pOwnerRTB, IID_IRichTableCell, reinterpret_cast<LPUNKNOWN*>(ppCell));
			if(*ppCell) {
				return S_OK;
			}
		}
	}

	// cell not found
	if(cellIdentifierType == citIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP TableCells::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}

STDMETHODIMP TableCells::get_ParentRow(IRichTableRow** ppRow)
{
	ATLASSERT_POINTER(ppRow, IRichTableRow*);
	if(!ppRow) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}

	ClassFactory::InitTableRow(properties.pTableTextRange, properties.pRowTextRange, properties.pOwnerRTB, IID_IRichTableRow, reinterpret_cast<LPUNKNOWN*>(ppRow));
	return S_OK;
}

STDMETHODIMP TableCells::get_ParentTable(IRichTable** ppTable)
{
	ATLASSERT_POINTER(ppTable, IRichTable*);
	if(!ppTable) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}

	ClassFactory::InitTable(properties.pTableTextRange, properties.pOwnerRTB, IID_IRichTable, reinterpret_cast<LPUNKNOWN*>(ppTable));
	return S_OK;
}


STDMETHODIMP TableCells::Add(LONG width, LONG insertAt/* = -1*/, SHORT borderSize/* = -1*/, VAlignmentConstants verticalAlignment/* = valCenter*/, IRichTableCell** ppAddedCell/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedCell, IRichTableCell*);
	if(!ppAddedCell) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	if(insertAt < -1 || borderSize < -1 || width < -1) {
		return E_INVALIDARG;
	}
	if(verticalAlignment < valTop || verticalAlignment > valBottom) {
		return E_INVALIDARG;
	}

	HWND hWndRTB = properties.GetRTBHWnd();
	ATLASSERT(IsWindow(hWndRTB));
	LONG twipsPerPixelX = 15;
	LONG twipsPerPixelY = 15;
	HDC hDCRTB = GetDC(hWndRTB);
	if(hDCRTB) {
		twipsPerPixelX = 1440 / GetDeviceCaps(hDCRTB, LOGPIXELSX);
		twipsPerPixelY = 1440 / GetDeviceCaps(hDCRTB, LOGPIXELSY);
		ReleaseDC(hWndRTB, hDCRTB);
	}
	if(borderSize == -1) {
		borderSize = static_cast<SHORT>(twipsPerPixelX);
	}

	*ppAddedCell = NULL;

	LONG cellCount = 0;
	HRESULT hr = Count(&cellCount);
	if(FAILED(hr)) {
		return hr;
	}
	if(insertAt == -1) {
		insertAt = cellCount;
	}
	LONG cellIndex = insertAt;
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	BOOL useFallback = TRUE;

	CComPtr<ITextRange> pRowTextRange = NULL;
	if(SUCCEEDED(properties.pRowTextRange->GetDuplicate(&pRowTextRange))) {
		CComQIPtr<ITextRange2> pTextRange2 = pRowTextRange;
		if(pTextRange2) {
			LONG rowStart = 0;
			pTextRange2->GetStart(&rowStart);
			pTextRange2->SetEnd(rowStart);
			if(apiVersion == reav41 || apiVersion >= reav75) {
				CComPtr<ITextRow> pTextRow = NULL;
				hr = pTextRange2->GetRow(&pTextRow);
				if(SUCCEEDED(hr) && pTextRow) {
					hr = pTextRow->SetCellIndex(cellIndex);
					if(SUCCEEDED(hr)) {
						hr = pTextRow->SetCellCount(cellCount + 1);
						if(SUCCEEDED(hr)) {
							pTextRow->SetCellIndex(insertAt);
							pTextRow->SetCellWidth(width);
							pTextRow->SetCellAlignment(verticalAlignment);
							pTextRow->SetCellBorderWidths(borderSize, borderSize, borderSize, borderSize);
							LONG backColor1 = RGB(255, 255, 255);
							CComPtr<ITextFont> pTextFont = NULL;
							if(SUCCEEDED(pTextRange2->GetFont(&pTextFont))) {
								pTextFont->GetBackColor(&backColor1);
							}
							pTextRow->SetCellColorBack(backColor1);
							if(SUCCEEDED(pTextRow->Apply(1, tomRowApplyDefault))) {
								useFallback = FALSE;
							}
						}
					}
				}
			}
		}
	}

	if(useFallback) {
		// NOTE: On Windows 7 this does not fail, but also does not have any effect.
		if(IsWindow(hWndRTB)) {
			TABLEROWPARMS rowParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			rowParams.cCell = static_cast<BYTE>(cellCount);
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, (rowParams.cCell + 1) * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						for(LONG cell = cellCount; cell > cellIndex; --cell) {
							pCellParams[cell] = pCellParams[cell - 1];
						}
						ZeroMemory(&pCellParams[cellIndex], sizeof(TABLECELLPARMS));
						pCellParams[cellIndex].dxWidth = width;
						pCellParams[cellIndex].nVertAlign = static_cast<WORD>(verticalAlignment);
						pCellParams[cellIndex].dxBrdrLeft = borderSize;
						pCellParams[cellIndex].dxBrdrRight = borderSize;
						pCellParams[cellIndex].dyBrdrTop = borderSize;
						pCellParams[cellIndex].dyBrdrBottom = borderSize;
						CComPtr<ITextFont> pTextFont = NULL;
						if(SUCCEEDED(properties.pRowTextRange->GetFont(&pTextFont))) {
							pTextFont->GetBackColor(reinterpret_cast<LONG*>(&pCellParams[cellIndex].crBackPat));
						} else {
							pCellParams[cellIndex].crBackPat = RGB(255, 255, 255);
						}

						rowParams.iCell = static_cast<BYTE>(cellIndex);
						rowParams.cCell = static_cast<BYTE>(cellCount + 1);
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}

	if(SUCCEEDED(hr)) {
		get_Item(cellIndex, citIndex, ppAddedCell);
	}
	return hr;
}

STDMETHODIMP TableCells::Contains(LONG cellIdentifier, CellIdentifierTypeConstants cellIdentifierType/* = citIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}

	*pValue = VARIANT_FALSE;

	// retrieve the cell's text range
	CComPtr<ITextRange> pCellTextRange = NULL;
	HRESULT hr = properties.pRowTextRange->GetDuplicate(&pCellTextRange);
	if(FAILED(hr)) {
		return hr;
	}
	LONG rowStart = 0;
	LONG rowEnd = 0;
	properties.pRowTextRange->GetStart(&rowStart);
	properties.pRowTextRange->GetEnd(&rowEnd);

	// NOTE: StartOf(tomCell, tomExtend) and EndOf(tomCell, tomExtend) do not really work on Windows 7.

	// start with the first cell
	int cellIndex = 0;
	BOOL foundCell = FALSE;
	LONG cellStart = rowStart + 2;
	LONG cellEnd = cellStart;
	LONG startOfNextCellCandidate = 0;
	do {
		// set the cursor to the current cell
		pCellTextRange->SetStart(cellStart);
		pCellTextRange->SetEnd(cellEnd);
		// move one cell ahead
		hr = pCellTextRange->Move(tomCell, 1, NULL);
		if(hr == S_OK) {
			// the start of the next cell minus one is the end of the current cell
			pCellTextRange->GetStart(&startOfNextCellCandidate);
			cellEnd = startOfNextCellCandidate - 1;
			if(startOfNextCellCandidate < rowEnd) {
				// we're still inside the row, so count this cell
				cellIndex++;
				switch(cellIdentifierType) {
					case citIndex:
						if(cellIdentifier == cellIndex - 1) {
							foundCell = TRUE;
						}
						break;
					case citCharacterIndex:
					{
						if(cellStart <= cellIdentifier && cellIdentifier <= cellEnd) {
							foundCell = TRUE;
						}
						break;
					}
				}
			}
			if(startOfNextCellCandidate >= rowEnd) {
				// we reached the end of the row
				break;
			}
		} else {
			// this has not been a cell, so don't count it and stop
			break;
		}
		if(!foundCell) {
			// prepare for next cell
			cellStart = startOfNextCellCandidate;
			cellEnd = cellStart;
		}
	} while(!foundCell);

	*pValue = BOOL2VARIANTBOOL(foundCell);
	return S_OK;
}

STDMETHODIMP TableCells::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;
	*pValue = 0;

	CComPtr<ITextRange> pRowTextRange = NULL;
	if(SUCCEEDED(properties.pRowTextRange->GetDuplicate(&pRowTextRange))) {
		CComQIPtr<ITextRange2> pTextRange2 = pRowTextRange;
		if(pTextRange2) {
			LONG rowStart = 0;
			pTextRange2->GetStart(&rowStart);
			pTextRange2->SetEnd(rowStart);
			if(apiVersion == reav41 || apiVersion >= reav75) {
				CComPtr<ITextRow> pTextRow = NULL;
				hr = pTextRange2->GetRow(&pTextRow);
				if(SUCCEEDED(hr) && pTextRow) {
					hr = pTextRow->GetCellCount(pValue);
					if(SUCCEEDED(hr)) {
						useFallback = FALSE;
					}
				}
			}
		}
	}

	if(useFallback) {
		HWND hWndRTB = properties.GetRTBHWnd();
		ATLASSERT(IsWindow(hWndRTB));
		if(IsWindow(hWndRTB)) {
			TABLEROWPARMS rowParams = {0};
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
				if(SUCCEEDED(hr)) {
					*pValue = rowParams.cCell;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCells::Remove(LONG cellIdentifier, CellIdentifierTypeConstants cellIdentifierType/* = citIndex*/)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}

	LONG cellCount = 0;
	HRESULT hr = Count(&cellCount);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	BOOL useFallback = TRUE;

	CComPtr<IRichTableCell> pCell = NULL;
	hr = get_Item(cellIdentifier, cellIdentifierType, &pCell);
	if(SUCCEEDED(hr) && pCell) {
		LONG cellIndex = 0;
		hr = pCell->get_Index(&cellIndex);
		if(SUCCEEDED(hr)) {
			CComPtr<ITextRange> pRowTextRange = NULL;
			if(SUCCEEDED(properties.pRowTextRange->GetDuplicate(&pRowTextRange))) {
				CComQIPtr<ITextRange2> pTextRange2 = pRowTextRange;
				if(pTextRange2) {
					LONG rowStart = 0;
					pTextRange2->GetStart(&rowStart);
					pTextRange2->SetEnd(rowStart);
					if(apiVersion == reav41 || apiVersion >= reav75) {
						CComPtr<ITextRow> pTextRow = NULL;
						hr = pTextRange2->GetRow(&pTextRow);
						if(SUCCEEDED(hr) && pTextRow) {
							hr = pTextRow->SetCellIndex(cellIndex);
							if(SUCCEEDED(hr)) {
								hr = pTextRow->SetCellCount(cellCount - 1);
								if(SUCCEEDED(hr)) {
									hr = pTextRow->Apply(1, tomCellStructureChangeOnly);
									if(SUCCEEDED(hr)) {
										useFallback = FALSE;
									}
								}
							}
						}
					}
				}
			}

			if(useFallback) {
				// NOTE: On Windows 7 this does not fail, but also does not have any effect.
				HWND hWndRTB = properties.GetRTBHWnd();
				ATLASSERT(IsWindow(hWndRTB));
				if(IsWindow(hWndRTB)) {
					TABLEROWPARMS rowParams = {0};
					rowParams.cbRow = sizeof(TABLEROWPARMS);
					rowParams.cbCell = sizeof(TABLECELLPARMS);
					rowParams.cRow = 1;
					rowParams.cCell = static_cast<BYTE>(cellCount);
					if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
						rowParams.iCell = static_cast<BYTE>(cellIndex);
						TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
						if(pCellParams) {
							hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
							if(SUCCEEDED(hr)) {
								for(LONG cell = cellIndex + 1; cell < cellCount; ++cell) {
									pCellParams[cell - 1] = pCellParams[cell];
								}
								rowParams.cCell = static_cast<BYTE>(cellCount - 1);
								hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
							}
							HeapFree(GetProcessHeap(), 0, pCellParams);
						} else {
							hr = E_OUTOFMEMORY;
						}
					}
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCells::RemoveAll(void)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}

	// as we don't support filtering, we can use a shortcut
	return properties.pRowTextRange->Delete(tomCharacter, 0, NULL);
}