// TableRows.cpp: Manages a collection of TableRow objects

#include "stdafx.h"
#include "TableRows.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TableRows::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTableRows, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TableRows::QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRange), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextRange2), queriedInterface)) {
		TableRows* pTableRows = reinterpret_cast<TableRows*>(pThis);
		return pTableRows->properties.pTableTextRange->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}

HRESULT CALLBACK TableRows::QueryITextRowInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRow), queriedInterface)) {
		TableRows* pTableRows = reinterpret_cast<TableRows*>(pThis);
		RichEditAPIVersionConstants apiVersion = reavUnknown;
		pTableRows->properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);
		if(apiVersion == reav41 || apiVersion >= reav75) {
			CComPtr<ITextRange> pRowTextRange = NULL;
			if(SUCCEEDED(pTableRows->properties.pTableTextRange->GetDuplicate(&pRowTextRange))) {
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
STDMETHODIMP TableRows::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<TableRows>* pTableRowsObj = NULL;
	CComObject<TableRows>::CreateInstance(&pTableRowsObj);
	pTableRowsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTableRowsObj->properties);

	pTableRowsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTableRowsObj->Release();
	return S_OK;
}

STDMETHODIMP TableRows::Next(ULONG numberOfMaxRows, VARIANT* pRows, ULONG* pNumberOfRowsReturned)
{
	ATLASSERT_POINTER(pRows, VARIANT);
	if(!pRows) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTableTextRange);

	// check each row in the text range
	ULONG i = 0;
	for(i = 0; i < numberOfMaxRows; ++i) {
		VariantInit(&pRows[i]);

		do {
			if(properties.pLastEnumeratedRowTextRange) {
				ITextRange* pTmp = GetNextRowToProcess(properties.pLastEnumeratedRowTextRange);
				properties.pLastEnumeratedRowTextRange->Release();
				properties.pLastEnumeratedRowTextRange = pTmp;
			} else {
				properties.pLastEnumeratedRowTextRange = GetFirstRowToProcess();
			}
		} while((properties.pLastEnumeratedRowTextRange != NULL) && (!IsPartOfCollection(properties.pLastEnumeratedRowTextRange)));

		if(properties.pLastEnumeratedRowTextRange) {
			CComPtr<ITextRange> pRowTextRange = NULL;
			if(SUCCEEDED(properties.pLastEnumeratedRowTextRange->GetDuplicate(&pRowTextRange))) {
				ClassFactory::InitTableRow(properties.pTableTextRange, pRowTextRange, properties.pOwnerRTB, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pRows[i].pdispVal));
				pRows[i].vt = VT_DISPATCH;
			} else {
				i--;
				break;
			}
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfRowsReturned) {
		*pNumberOfRowsReturned = i;
	}

	return (i == numberOfMaxRows ? S_OK : S_FALSE);
}

STDMETHODIMP TableRows::Reset(void)
{
	if(properties.pLastEnumeratedRowTextRange) {
		properties.pLastEnumeratedRowTextRange->Release();
		properties.pLastEnumeratedRowTextRange = NULL;
	}
	return S_OK;
}

STDMETHODIMP TableRows::Skip(ULONG numberOfRowsToSkip)
{
	VARIANT dummy;
	ULONG numRowsReturned = 0;
	// we could skip all rows at once, but it's easier to skip them one by one
	for(ULONG i = 1; i <= numberOfRowsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numRowsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numRowsReturned == 0) {
			// there're no more rows to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////

ITextRange* TableRows::GetFirstRowToProcess(void)
{
	if(!properties.pTableTextRange) {
		return NULL;
	}

	CComPtr<ITextRange> pRowTextRange = NULL;
	if(SUCCEEDED(properties.pTableTextRange->GetDuplicate(&pRowTextRange))) {
		LONG tableStart = 0;
		LONG tableEnd = 0;
		properties.pTableTextRange->GetStart(&tableStart);
		properties.pTableTextRange->GetEnd(&tableEnd);

		// start with the first row
		LONG startOfNextRowCandidate = tableStart;
		HRESULT hr = S_OK;
		do {
			// set the cursor to the current row
			pRowTextRange->SetStart(startOfNextRowCandidate);
			pRowTextRange->SetEnd(startOfNextRowCandidate);
			hr = pRowTextRange->EndOf(tomRow, tomExtend, NULL);
			if(hr == S_OK) {
				// this has been a row, retrieve the position at which the next row would start
				pRowTextRange->GetEnd(&startOfNextRowCandidate);
				if(startOfNextRowCandidate <= tableEnd) {
					// we're still inside the table, so use this row
					return pRowTextRange.Detach();
				}
				if(startOfNextRowCandidate >= tableEnd) {
					// we reached the end of the table
					break;
				}
			} else {
				// this has not been a row, so don't use it and stop
				break;
			}
		} while(TRUE);
	}
	return NULL;
}

ITextRange* TableRows::GetNextRowToProcess(ITextRange* pPreviousRow)
{
	if(!properties.pTableTextRange) {
		return NULL;
	}

	CComPtr<ITextRange> pRowTextRange = NULL;
	if(SUCCEEDED(properties.pTableTextRange->GetDuplicate(&pRowTextRange))) {
		LONG tableEnd = 0;
		properties.pTableTextRange->GetEnd(&tableEnd);
		LONG startOfNextRowCandidate = 0;
		pPreviousRow->GetEnd(&startOfNextRowCandidate);
		if(startOfNextRowCandidate >= tableEnd) {
			// we already reached the end of the table
			return NULL;
		}

		HRESULT hr = S_OK;
		do {
			// set the cursor to the current row
			pRowTextRange->SetStart(startOfNextRowCandidate);
			pRowTextRange->SetEnd(startOfNextRowCandidate);
			hr = pRowTextRange->EndOf(tomRow, tomExtend, NULL);
			if(hr == S_OK) {
				// this has been a row, retrieve the position at which the next row would start
				pRowTextRange->GetEnd(&startOfNextRowCandidate);
				if(startOfNextRowCandidate <= tableEnd) {
					// we're still inside the table, so use this row
					return pRowTextRange.Detach();
				}
				if(startOfNextRowCandidate >= tableEnd) {
					// we reached the end of the table
					break;
				}
			} else {
				// this has not been a row, so don't use it and stop
				break;
			}
		} while(TRUE);
	}
	return NULL;
}

BOOL TableRows::IsPartOfCollection(ITextRange* pRowTextRange)
{
	LONG tableStart = 0;
	LONG tableEnd = 0;
	if(SUCCEEDED(properties.pTableTextRange->GetStart(&tableStart)) && SUCCEEDED(properties.pTableTextRange->GetEnd(&tableEnd))) {
		LONG rowStart = 0;
		LONG rowEnd = 0;
		if(SUCCEEDED(pRowTextRange->GetStart(&rowStart)) && SUCCEEDED(pRowTextRange->GetEnd(&rowEnd))) {
			if(rowStart <= rowEnd && rowStart >= tableStart && rowEnd <= tableEnd) {
				CComPtr<ITextRange> pTextRange = NULL;
				if(SUCCEEDED(pRowTextRange->GetDuplicate(&pTextRange))) {
					pTextRange->StartOf(tomRow, tomExtend, NULL);
					pTextRange->EndOf(tomRow, tomExtend, NULL);
					LONG rowStart2 = 0;
					LONG rowEnd2 = 0;
					if(SUCCEEDED(pTextRange->GetStart(&rowStart2)) && SUCCEEDED(pTextRange->GetEnd(&rowEnd2))) {
						if(rowStart == rowStart2 && rowEnd == rowEnd2) {
							return TRUE;
						}
					}
				}
			}
		}
	}
	return FALSE;
}


TableRows::Properties::~Properties()
{
	if(pOwnerRTB) {
		pOwnerRTB->Release();
		pOwnerRTB = NULL;
	}
	if(pTableTextRange) {
		pTableTextRange->Release();
		pTableTextRange = NULL;
	}
	if(pLastEnumeratedRowTextRange) {
		pLastEnumeratedRowTextRange->Release();
		pLastEnumeratedRowTextRange = NULL;
	}
}

void TableRows::Properties::CopyTo(TableRows::Properties* pTarget)
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
		pTarget->pLastEnumeratedRowTextRange = NULL;
		if(this->pLastEnumeratedRowTextRange) {
			this->pLastEnumeratedRowTextRange->GetDuplicate(&pTarget->pLastEnumeratedRowTextRange);
		}
	}
}

HWND TableRows::Properties::GetRTBHWnd(void)
{
	ATLASSUME(pOwnerRTB);

	OLE_HANDLE handle = NULL;
	pOwnerRTB->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TableRows::Attach(ITextRange* pTableTextRange)
{
	if(properties.pTableTextRange) {
		properties.pTableTextRange->Release();
		properties.pTableTextRange = NULL;
	}
	if(pTableTextRange) {
		//pTableTextRange->QueryInterface(IID_PPV_ARGS(&properties.pTableTextRange));
		properties.pTableTextRange = pTableTextRange;
		properties.pTableTextRange->AddRef();
	}
}

void TableRows::Detach(void)
{
	if(properties.pTableTextRange) {
		properties.pTableTextRange->Release();
		properties.pTableTextRange = NULL;
	}
	if(properties.pLastEnumeratedRowTextRange) {
		properties.pLastEnumeratedRowTextRange->Release();
		properties.pLastEnumeratedRowTextRange = NULL;
	}
}

void TableRows::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP TableRows::get_Item(LONG rowIdentifier, RowIdentifierTypeConstants rowIdentifierType/* = ritIndex*/, IRichTableRow** ppRow/* = NULL*/)
{
	ATLASSERT_POINTER(ppRow, IRichTableRow*);
	if(!ppRow) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}

	// retrieve the row's text range
	CComPtr<ITextRange> pRowTextRange = NULL;
	HRESULT hr = properties.pTableTextRange->GetDuplicate(&pRowTextRange);
	if(FAILED(hr)) {
		return hr;
	}
	LONG tableStart = 0;
	LONG tableEnd = 0;
	properties.pTableTextRange->GetStart(&tableStart);
	properties.pTableTextRange->GetEnd(&tableEnd);

	// start with the first row
	LONG startOfNextRowCandidate = tableStart;
	int rowIndex = 0;
	BOOL foundRow = FALSE;
	do {
		// set the cursor to the current row
		pRowTextRange->SetStart(startOfNextRowCandidate);
		pRowTextRange->SetEnd(startOfNextRowCandidate);
		hr = pRowTextRange->EndOf(tomRow, tomExtend, NULL);
		if(hr == S_OK) {
			// this has been a row, retrieve the position at which the next row would start
			pRowTextRange->GetEnd(&startOfNextRowCandidate);
			if(startOfNextRowCandidate <= tableEnd) {
				// we're still inside the table, so count this row
				rowIndex++;
				switch(rowIdentifierType) {
					case ritIndex:
						if(rowIdentifier == rowIndex - 1) {
							foundRow = TRUE;
						}
						break;
					case ritCharacterIndex:
					{
						LONG rowStart = 0;
						LONG rowEnd = 0;
						pRowTextRange->GetStart(&rowStart);
						pRowTextRange->GetEnd(&rowEnd);
						if(rowStart <= rowIdentifier && rowIdentifier <= rowEnd) {
							foundRow = TRUE;
						}
						break;
					}
				}
			}
			if(startOfNextRowCandidate >= tableEnd) {
				// we reached the end of the table
				break;
			}
		} else {
			// this has not been a row, so don't count it and stop
			break;
		}
	} while(!foundRow);

	if(foundRow) {
		if(IsPartOfCollection(pRowTextRange)) {
			ClassFactory::InitTableRow(properties.pTableTextRange, pRowTextRange, properties.pOwnerRTB, IID_IRichTableRow, reinterpret_cast<LPUNKNOWN*>(ppRow));
			if(*ppRow) {
				return S_OK;
			}
		}
	}

	// row not found
	if(rowIdentifierType == ritIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP TableRows::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}

STDMETHODIMP TableRows::get_ParentTable(IRichTable** ppTable)
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


// TODO: columnCount, allowIndividualCellStyles
STDMETHODIMP TableRows::Add(LONG insertAt/* = -1*/, HAlignmentConstants horizontalAlignment/* = halLeft*/, LONG indent/* = -1*/, LONG horizontalCellMargin/* = -1*/, LONG height/* = -1*/, IRichTableRow** ppAddedRow/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedRow, IRichTableRow*);
	if(!ppAddedRow) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}
	if(indent < -1 || horizontalCellMargin < -1 || height < -1) {
		return E_INVALIDARG;
	}
	if(horizontalAlignment < halLeft || horizontalAlignment > halRight) {
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
	if(indent == -1) {
		indent = twipsPerPixelX * 35 / 10;
	}
	if(horizontalCellMargin == -1) {
		horizontalCellMargin = twipsPerPixelX * 48 / 10;
	}

	*ppAddedRow = NULL;

	LONG rowCount = 0;
	HRESULT hr = Count(&rowCount);
	if(FAILED(hr)) {
		return hr;
	}
	if(insertAt == -1) {
		insertAt = rowCount;
	}

	BOOL useFallback = TRUE;

	CComPtr<IRichTableRow> pRow = NULL;
	if(insertAt == rowCount) {
		hr = get_Item(insertAt - 1, ritIndex, &pRow);
	} else {
		hr = get_Item(insertAt, ritIndex, &pRow);
	}
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	CComQIPtr<ITextRange> pRowTextRange = pRow;
	if(pRowTextRange) {
		CComPtr<ITextRange> pTextRange = NULL;
		if(SUCCEEDED(pRowTextRange->GetDuplicate(&pTextRange))) {
			hr = E_FAIL;
			if(insertAt == rowCount) {
				// even if setting the TextRange to the row's end TOM cannot append a row
			} else {
				LONG rowStart = 0;
				pTextRange->GetStart(&rowStart);
				pTextRange->SetEnd(rowStart);
				if(apiVersion == reav41 || apiVersion >= reav75) {
					CComQIPtr<ITextRange2> pTextRange2 = pTextRange;
					if(pTextRange2) {
						CComPtr<ITextRow> pTextRow = NULL;
						hr = pTextRange2->GetRow(&pTextRow);
						if(SUCCEEDED(hr) && pTextRow) {
							hr = pTextRow->Insert(1);
						}
					}
				}
			}
			if(SUCCEEDED(hr)) {
				useFallback = FALSE;
				// pTextRow still references the old row
				pRow = NULL;
				if(SUCCEEDED(get_Item(insertAt, ritIndex, &pRow))) {
					CComQIPtr<ITextRange2> pRowTextRange2 = pRow;
					if(pRowTextRange2) {
						CComPtr<ITextRange2> pTextRange2 = NULL;
						if(SUCCEEDED(pRowTextRange2->GetDuplicate2(&pTextRange2))) {
							LONG rowStart = 0;
							pTextRange2->GetStart(&rowStart);
							pTextRange2->SetEnd(rowStart);
							CComPtr<ITextRow> pTextRow = NULL;
							if(SUCCEEDED(pTextRange2->GetRow(&pTextRow))) {
								pTextRow->SetAlignment(static_cast<LONG>(horizontalAlignment));
								pTextRow->SetCellMargin(horizontalCellMargin);
								pTextRow->SetIndent(indent);
								pTextRow->SetHeight(height);
								pTextRow->Apply(1, tomRowApplyDefault);
							}
						}
					}
				}
			}

			if(useFallback && IsWindow(hWndRTB)) {
				if(insertAt == 0) {
					// no idea how to do this...
					ATLASSERT(FALSE && "If using the fallback, you cannot insert a row at position 0");
					return E_NOTIMPL;
				}

				// this will destroy the current selection, but saving & restoring the selection seems quite difficult
				LONG rowEnd = 0;
				pTextRange->GetEnd(&rowEnd);
				pTextRange->SetStart(rowEnd - 2);
				pTextRange->SetEnd(rowEnd - 2);
				pTextRange->Select();

				SendMessage(hWndRTB, WM_KEYDOWN, VK_RETURN, 0);
				SendMessage(hWndRTB, WM_KEYUP, VK_RETURN, 0);

				hr = S_OK;
				// update end of properties.pTableTextRange
				pRow = NULL;
				if(SUCCEEDED(get_Item(rowCount - 1, ritIndex, &pRow))) {
					CComQIPtr<ITextRange> pRowTextRange = pRow;
					if(pRowTextRange) {
						if(SUCCEEDED(pRowTextRange->EndOf(tomRow, tomExtend, NULL))) {
							LONG startOfNextRowCandidate = 0;
							pRowTextRange->GetEnd(&startOfNextRowCandidate);
							pRowTextRange->SetStart(startOfNextRowCandidate);
							pRowTextRange->SetEnd(startOfNextRowCandidate);
							pRowTextRange->EndOf(tomRow, tomExtend, NULL);
							pRowTextRange->GetEnd(&startOfNextRowCandidate);
							properties.pTableTextRange->SetEnd(startOfNextRowCandidate);
						}
					}
				}
				pRow = NULL;
				if(SUCCEEDED(get_Item(insertAt, ritIndex, &pRow))) {
					// NOTE: This won't work with RichEdit prior to 8.0.
					pRow->put_HAlignment(horizontalAlignment);
					pRow->put_HorizontalCellMargin(horizontalCellMargin);
					pRow->put_Indent(indent);
					pRow->put_Height(height);
				}
			}
		}
	}

	if(SUCCEEDED(hr)) {
		get_Item(insertAt, ritIndex, ppAddedRow);
	}
	return hr;
}

STDMETHODIMP TableRows::Contains(LONG rowIdentifier, RowIdentifierTypeConstants rowIdentifierType/* = ritIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}

	*pValue = VARIANT_FALSE;

	// retrieve the row's text range
	CComPtr<ITextRange> pRowTextRange = NULL;
	HRESULT hr = properties.pTableTextRange->GetDuplicate(&pRowTextRange);
	if(FAILED(hr)) {
		return hr;
	}
	LONG tableStart = 0;
	LONG tableEnd = 0;
	properties.pTableTextRange->GetStart(&tableStart);
	properties.pTableTextRange->GetEnd(&tableEnd);

	// start with the first row
	LONG startOfNextRowCandidate = tableStart;
	int rowIndex = 0;
	BOOL foundRow = FALSE;
	do {
		// set the cursor to the current row
		pRowTextRange->SetStart(startOfNextRowCandidate);
		pRowTextRange->SetEnd(startOfNextRowCandidate);
		hr = pRowTextRange->EndOf(tomRow, tomExtend, NULL);
		if(hr == S_OK) {
			// this has been a row, retrieve the position at which the next row would start
			pRowTextRange->GetEnd(&startOfNextRowCandidate);
			if(startOfNextRowCandidate <= tableEnd) {
				// we're still inside the table, so count this row
				rowIndex++;
				switch(rowIdentifierType) {
					case ritIndex:
						if(rowIdentifier == rowIndex - 1) {
							foundRow = TRUE;
						}
						break;
					case ritCharacterIndex:
					{
						LONG rowStart = 0;
						LONG rowEnd = 0;
						pRowTextRange->GetStart(&rowStart);
						pRowTextRange->GetEnd(&rowEnd);
						if(rowStart <= rowIdentifier && rowIdentifier <= rowEnd) {
							foundRow = TRUE;
						}
						break;
					}
				}
			}
			if(startOfNextRowCandidate >= tableEnd) {
				// we reached the end of the table
				break;
			}
		} else {
			// this has not been a row, so don't count it and stop
			break;
		}
	} while(!foundRow);

	*pValue = BOOL2VARIANTBOOL(foundRow);
	return S_OK;
}

STDMETHODIMP TableRows::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = 0;
	// we'll change the range start and end, so work with a clone
	CComPtr<ITextRange> pRowTextRange = NULL;
	if(SUCCEEDED(properties.pTableTextRange->GetDuplicate(&pRowTextRange))) {
		LONG tableStart = 0;
		LONG tableEnd = 0;
		properties.pTableTextRange->GetStart(&tableStart);
		properties.pTableTextRange->GetEnd(&tableEnd);

		// start with the first row
		LONG startOfNextRowCandidate = tableStart;
		HRESULT hr2 = S_OK;
		do {
			// set the cursor to the current row
			pRowTextRange->SetStart(startOfNextRowCandidate);
			pRowTextRange->SetEnd(startOfNextRowCandidate);
			hr2 = pRowTextRange->EndOf(tomRow, tomExtend, NULL);
			if(hr2 == S_OK) {
				// this has been a row, retrieve the position at which the next row would start
				pRowTextRange->GetEnd(&startOfNextRowCandidate);
				if(startOfNextRowCandidate <= tableEnd) {
					// we're still inside the table, so count this row
					if(IsPartOfCollection(pRowTextRange)) {
						(*pValue)++;
					}
				}
				if(startOfNextRowCandidate >= tableEnd) {
					// we reached the end of the table
					break;
				}
			} else {
				// this has not been a row, so don't count it and stop
				break;
			}
		} while(TRUE);
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP TableRows::Remove(LONG rowIdentifier, RowIdentifierTypeConstants rowIdentifierType/* = ritIndex*/)
{
	ATLASSERT(properties.pTableTextRange);

	// retrieve the row's text range
	CComPtr<ITextRange> pRowTextRange = NULL;
	HRESULT hr = properties.pTableTextRange->GetDuplicate(&pRowTextRange);
	if(FAILED(hr)) {
		return hr;
	}
	LONG tableStart = 0;
	LONG tableEnd = 0;
	properties.pTableTextRange->GetStart(&tableStart);
	properties.pTableTextRange->GetEnd(&tableEnd);

	// start with the first row
	LONG startOfNextRowCandidate = tableStart;
	int rowIndex = 0;
	BOOL foundRow = FALSE;
	do {
		// set the cursor to the current row
		pRowTextRange->SetStart(startOfNextRowCandidate);
		pRowTextRange->SetEnd(startOfNextRowCandidate);
		hr = pRowTextRange->EndOf(tomRow, tomExtend, NULL);
		if(hr == S_OK) {
			// this has been a row, retrieve the position at which the next row would start
			pRowTextRange->GetEnd(&startOfNextRowCandidate);
			if(startOfNextRowCandidate <= tableEnd) {
				// we're still inside the table, so count this row
				rowIndex++;
				switch(rowIdentifierType) {
					case ritIndex:
						if(rowIdentifier == rowIndex - 1) {
							foundRow = TRUE;
						}
						break;
					case ritCharacterIndex:
					{
						LONG rowStart = 0;
						LONG rowEnd = 0;
						pRowTextRange->GetStart(&rowStart);
						pRowTextRange->GetEnd(&rowEnd);
						if(rowStart <= rowIdentifier && rowIdentifier <= rowEnd) {
							foundRow = TRUE;
						}
						break;
					}
				}
			}
			if(startOfNextRowCandidate >= tableEnd) {
				// we reached the end of the table
				break;
			}
		} else {
			// this has not been a row, so don't count it and stop
			break;
		}
	} while(!foundRow);

	if(foundRow) {
		if(IsPartOfCollection(pRowTextRange)) {
			return pRowTextRange->Delete(tomCharacter, 0, NULL);
		}
	}

	// row not found
	if(rowIdentifierType == ritIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP TableRows::RemoveAll(void)
{
	ATLASSERT(properties.pTableTextRange);

	// as we don't support filtering, we can use a shortcut
	return properties.pTableTextRange->Delete(tomCharacter, 0, NULL);
}