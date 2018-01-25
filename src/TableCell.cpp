// TableCell.cpp: A wrapper for the styling of a table cell.

#include "stdafx.h"
#include "TableCell.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TableCell::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTableCell, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TableCell::QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRange), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextRange2), queriedInterface)) {
		TableCell* pTableCell = reinterpret_cast<TableCell*>(pThis);
		return pTableCell->properties.pCellTextRange->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}

HRESULT CALLBACK TableCell::QueryITextRowInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRow), queriedInterface)) {
		TableCell* pTableCell = reinterpret_cast<TableCell*>(pThis);
		RichEditAPIVersionConstants apiVersion = reavUnknown;
		pTableCell->properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);
		if(apiVersion == reav41 || apiVersion >= reav75) {
			LONG cellIndex = 0;
			HRESULT hr = pTableCell->get_Index(&cellIndex);
			if(FAILED(hr)) {
				return hr;
			}
			CComPtr<ITextRange> pRowTextRange = NULL;
			if(SUCCEEDED(pTableCell->properties.pRowTextRange->GetDuplicate(&pRowTextRange))) {
				CComQIPtr<ITextRange2> pTextRange2 = pRowTextRange;
				if(pTextRange2) {
					LONG rowStart = 0;
					pTextRange2->GetStart(&rowStart);
					pTextRange2->SetEnd(rowStart);
					ITextRow* pTextRow = NULL;
					hr = pTextRange2->GetRow(&pTextRow);
					if(SUCCEEDED(hr) && pTextRow) {
						hr = pTextRow->SetCellIndex(cellIndex);
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


TableCell::Properties::~Properties()
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
	if(pCellTextRange) {
		pCellTextRange->Release();
		pCellTextRange = NULL;
	}
}

HWND TableCell::Properties::GetRTBHWnd(void)
{
	ATLASSUME(pOwnerRTB);

	OLE_HANDLE handle = NULL;
	pOwnerRTB->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TableCell::Attach(ITextRange* pTableTextRange, ITextRange* pRowTextRange, ITextRange* pCellTextRange)
{
	if(properties.pTableTextRange) {
		properties.pTableTextRange->Release();
		properties.pTableTextRange = NULL;
	}
	if(properties.pRowTextRange) {
		properties.pRowTextRange->Release();
		properties.pRowTextRange = NULL;
	}
	if(properties.pCellTextRange) {
		properties.pCellTextRange->Release();
		properties.pCellTextRange = NULL;
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
	if(pCellTextRange) {
		//pCellTextRange->QueryInterface(IID_PPV_ARGS(&properties.pCellTextRange));
		properties.pCellTextRange = pCellTextRange;
		properties.pCellTextRange->AddRef();
	}
}

void TableCell::Detach(void)
{
	if(properties.pTableTextRange) {
		properties.pTableTextRange->Release();
		properties.pTableTextRange = NULL;
	}
	if(properties.pRowTextRange) {
		properties.pRowTextRange->Release();
		properties.pRowTextRange = NULL;
	}
	if(properties.pCellTextRange) {
		properties.pCellTextRange->Release();
		properties.pCellTextRange = NULL;
	}
}

void TableCell::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP TableCell::get_BackColor1(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = 0;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						LONG clr = 0;
						hr = pTextRow->GetCellColorBack(&clr);
						if(SUCCEEDED(hr)) {
							*pValue = clr;
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						*pValue = pCellParams[cellIndex].crBackPat;
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	if(SUCCEEDED(hr)) {
		if(static_cast<int>(*pValue) < 0) {
			// e.g. tomAutoColor
			*pValue = static_cast<OLE_COLOR>(-1);
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_BackColor1(OLE_COLOR newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	LONG color = tomAutoColor;
	if(static_cast<int>(newValue) >= 0) {
		color = OLECOLOR2COLORREF(newValue);
	}

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						hr = pTextRow->SetCellColorBack(color);
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
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						pCellParams[cellIndex].crBackPat = color;
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::get_BackColor2(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = 0;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						LONG clr = 0;
						hr = pTextRow->GetCellColorFore(&clr);
						if(SUCCEEDED(hr)) {
							*pValue = clr;
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						*pValue = pCellParams[cellIndex].crForePat;
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	if(SUCCEEDED(hr)) {
		if(static_cast<int>(*pValue) < 0) {
			// e.g. tomAutoColor
			*pValue = static_cast<OLE_COLOR>(-1);
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_BackColor2(OLE_COLOR newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	LONG color = tomAutoColor;
	if(static_cast<int>(newValue) >= 0) {
		color = OLECOLOR2COLORREF(newValue);
	}

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						hr = pTextRow->SetCellColorFore(color);
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
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						pCellParams[cellIndex].crForePat = color;
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::get_BackColorShading(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = 0;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						hr = pTextRow->GetCellShading(pValue);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						*pValue = static_cast<LONG>(pCellParams[cellIndex].wShading);
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_BackColorShading(LONG newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						hr = pTextRow->SetCellShading(newValue);
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
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						pCellParams[cellIndex].wShading = static_cast<SHORT>(newValue);
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::get_BorderColor(CellBorderConstants border, OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = 0;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						LONG left = 0;
						LONG top = 0;
						LONG right = 0;
						LONG bottom = 0;
						hr = pTextRow->GetCellBorderColors(&left, &top, &right, &bottom);
						if(SUCCEEDED(hr)) {
							switch(border) {
								case cbLeft:
									*pValue = left;
									break;
								case cbTop:
									*pValue = top;
									break;
								case cbRight:
									*pValue = right;
									break;
								case cbBottom:
									*pValue = bottom;
									break;
							}
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						switch(border) {
							case cbLeft:
								*pValue = pCellParams[cellIndex].crBrdrLeft;
								break;
							case cbTop:
								*pValue = pCellParams[cellIndex].crBrdrTop;
								break;
							case cbRight:
								*pValue = pCellParams[cellIndex].crBrdrRight;
								break;
							case cbBottom:
								*pValue = pCellParams[cellIndex].crBrdrBottom;
								break;
						}
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_BorderColor(CellBorderConstants border, OLE_COLOR newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	COLORREF color = OLECOLOR2COLORREF(newValue);

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						LONG left = 0;
						LONG top = 0;
						LONG right = 0;
						LONG bottom = 0;
						hr = pTextRow->GetCellBorderColors(&left, &top, &right, &bottom);
						if(SUCCEEDED(hr)) {
							switch(border) {
								case cbLeft:
									left = color;
									break;
								case cbTop:
									top = color;
									break;
								case cbRight:
									right = color;
									break;
								case cbBottom:
									bottom = color;
									break;
							}
							hr = pTextRow->SetCellBorderColors(left, top, right, bottom);
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
	}

	if(useFallback) {
		// NOTE: On Windows 7 this does not fail, but also does not have any effect.
		HWND hWndRTB = properties.GetRTBHWnd();
		ATLASSERT(IsWindow(hWndRTB));
		if(IsWindow(hWndRTB)) {
			TABLEROWPARMS rowParams = {0};
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						switch(border) {
							case cbLeft:
								pCellParams[cellIndex].crBrdrLeft = color;
								break;
							case cbTop:
								pCellParams[cellIndex].crBrdrTop = color;
								break;
							case cbRight:
								pCellParams[cellIndex].crBrdrRight = color;
								break;
							case cbBottom:
								pCellParams[cellIndex].crBrdrBottom = color;
								break;
						}
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::get_BorderSize(CellBorderConstants border, SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = 0;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						LONG left = 0;
						LONG top = 0;
						LONG right = 0;
						LONG bottom = 0;
						hr = pTextRow->GetCellBorderWidths(&left, &top, &right, &bottom);
						if(SUCCEEDED(hr)) {
							switch(border) {
								case cbLeft:
									*pValue = static_cast<SHORT>(left);
									break;
								case cbTop:
									*pValue = static_cast<SHORT>(top);
									break;
								case cbRight:
									*pValue = static_cast<SHORT>(right);
									break;
								case cbBottom:
									*pValue = static_cast<SHORT>(bottom);
									break;
							}
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						switch(border) {
							case cbLeft:
								*pValue = pCellParams[cellIndex].dxBrdrLeft;
								break;
							case cbTop:
								*pValue = pCellParams[cellIndex].dyBrdrTop;
								break;
							case cbRight:
								*pValue = pCellParams[cellIndex].dxBrdrRight;
								break;
							case cbBottom:
								*pValue = pCellParams[cellIndex].dyBrdrBottom;
								break;
						}
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_BorderSize(CellBorderConstants border, SHORT newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
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
	if(newValue == -1) {
		newValue = static_cast<SHORT>(twipsPerPixelX);
	}

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						LONG left = 0;
						LONG top = 0;
						LONG right = 0;
						LONG bottom = 0;
						hr = pTextRow->GetCellBorderWidths(&left, &top, &right, &bottom);
						if(SUCCEEDED(hr)) {
							switch(border) {
								case cbLeft:
									left = newValue;
									break;
								case cbTop:
									top = newValue;
									break;
								case cbRight:
									right = newValue;
									break;
								case cbBottom:
									bottom = newValue;
									break;
							}
							hr = pTextRow->SetCellBorderWidths(left, top, right, bottom);
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
	}

	if(useFallback) {
		// NOTE: On Windows 7 this does not fail, but also does not have any effect.
		HWND hWndRTB = properties.GetRTBHWnd();
		ATLASSERT(IsWindow(hWndRTB));
		if(IsWindow(hWndRTB)) {
			TABLEROWPARMS rowParams = {0};
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						switch(border) {
							case cbLeft:
								pCellParams[cellIndex].dxBrdrLeft = newValue;
								break;
							case cbTop:
								pCellParams[cellIndex].dyBrdrTop = newValue;
								break;
							case cbRight:
								pCellParams[cellIndex].dxBrdrRight = newValue;
								break;
							case cbBottom:
								pCellParams[cellIndex].dyBrdrBottom = newValue;
								break;
						}
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::get_CellMergeFlags(CellMergeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, CellMergeConstants);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = cmNotMerged;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						LONG flags = 0;
						hr = pTextRow->GetCellMergeFlags(&flags);
						if(SUCCEEDED(hr)) {
							*pValue = static_cast<CellMergeConstants>(flags);
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						*pValue = static_cast<CellMergeConstants>(*pValue | (pCellParams[cellIndex].fMergeTop ? cmTopCellInVerticallyMergedSet : 0));
						*pValue = static_cast<CellMergeConstants>(*pValue | (pCellParams[cellIndex].fMergePrev ? cmContinueVerticallyMergedSet : 0));
						*pValue = static_cast<CellMergeConstants>(*pValue | (pCellParams[cellIndex].fMergeStart ? cmStartCellInHorizontallyMergedSet : 0));
						*pValue = static_cast<CellMergeConstants>(*pValue | (pCellParams[cellIndex].fMergeCont ? cmContinueHorizontallyMergedSet : 0));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_CellMergeFlags(CellMergeConstants newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	// NOTE: Horizontal merging does not work well for existing tables. The generated RTF seems to be buggy.
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
						hr = pTextRow->SetCellMergeFlags(newValue);
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
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						pCellParams[cellIndex].fMergeTop = ((newValue & cmTopCellInVerticallyMergedSet) == cmTopCellInVerticallyMergedSet);
						pCellParams[cellIndex].fMergePrev = ((newValue & cmContinueVerticallyMergedSet) == cmContinueVerticallyMergedSet);
						pCellParams[cellIndex].fMergeStart = ((newValue & cmStartCellInHorizontallyMergedSet) == cmStartCellInHorizontallyMergedSet);
						pCellParams[cellIndex].fMergeCont = ((newValue & cmContinueHorizontallyMergedSet) == cmContinueHorizontallyMergedSet);
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::get_DefinesStyleForSubsequentCells(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = VARIANT_FALSE;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						LONG cacheCount = 0;
						hr = pTextRow->GetCellCountCache(&cacheCount);
						if(SUCCEEDED(hr)) {
							*pValue = BOOL2VARIANTBOOL(cellIndex == cacheCount - 1);
							useFallback = FALSE;
						}
					}
				}
			}
		}
	}

	if(useFallback) {
		return E_NOTIMPL;
	}
	return hr;
}

STDMETHODIMP TableCell::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}

	// find the cell within the row
	LONG thisCellStart = 0;
	LONG thisCellEnd = 0;
	HRESULT hr = properties.pCellTextRange->GetStart(&thisCellStart);
	if(FAILED(hr)) {
		return hr;
	}
	hr = properties.pCellTextRange->GetEnd(&thisCellEnd);
	if(FAILED(hr)) {
		return hr;
	}
	CComPtr<ITextRange> pCellTextRange = NULL;
	hr = properties.pRowTextRange->GetDuplicate(&pCellTextRange);
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
				if(cellStart == thisCellStart && cellEnd == thisCellEnd) {
					foundCell = TRUE;
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
		*pValue = cellIndex - 1;
		return S_OK;
	} else {
		*pValue = 0;
		return E_FAIL;
	}
}

STDMETHODIMP TableCell::get_ParentRow(IRichTableRow** ppRow)
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

STDMETHODIMP TableCell::get_ParentTable(IRichTable** ppTable)
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

STDMETHODIMP TableCell::get_TextFlow(TextFlowConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TextFlowConstants);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = tfLeftToRightTopToBottom;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						// actually it seems to be a BOOL, but fortunately TRUE and FALSE map perfectly to the corresponding tf* values
						LONG textFlow = 0;
						hr = pTextRow->GetCellVerticalText(&textFlow);
						if(SUCCEEDED(hr)) {
							*pValue = static_cast<TextFlowConstants>(textFlow);
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						// actually it seems to be a BOOL, but fortunately TRUE and FALSE map perfectly to the corresponding tf* values
						*pValue = static_cast<TextFlowConstants>(pCellParams[cellIndex].fVertical);
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_TextFlow(TextFlowConstants newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						// actually it seems to be a BOOL, but fortunately TRUE and FALSE map perfectly to the corresponding tf* values
						hr = pTextRow->SetCellVerticalText(newValue);
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
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						// actually it seems to be a BOOL, but fortunately TRUE and FALSE map perfectly to the corresponding tf* values
						pCellParams[cellIndex].fVertical = newValue;
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::get_TextRange(IRichTextRange** ppTextRange)
{
	ATLASSERT_POINTER(ppTextRange, IRichTextRange*);
	if(!ppTextRange) {
		return E_POINTER;
	}

	CComPtr<ITextRange> pTextRange = NULL;
	HRESULT hr = properties.pCellTextRange->GetDuplicate(&pTextRange);
	if(SUCCEEDED(hr)) {
		ClassFactory::InitTextRange(pTextRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppTextRange));
	}
	return hr;
}

STDMETHODIMP TableCell::get_VAlignment(VAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, VAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = valCenter;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						hr = pTextRow->GetCellAlignment(reinterpret_cast<LONG*>(pValue));
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						*pValue = static_cast<VAlignmentConstants>(pCellParams[cellIndex].nVertAlign);
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_VAlignment(VAlignmentConstants newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						hr = pTextRow->SetCellAlignment(newValue);
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
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						pCellParams[cellIndex].nVertAlign = static_cast<WORD>(newValue);
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::get_Width(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	*pValue = 0;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						hr = pTextRow->GetCellWidth(pValue);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
						}
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
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						*pValue = pCellParams[cellIndex].dxWidth;
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableCell::put_Width(LONG newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange || !properties.pCellTextRange) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = 0;
	hr = get_Index(&cellIndex);
	if(FAILED(hr)) {
		return hr;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

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
						hr = pTextRow->SetCellWidth(newValue);
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
			TABLECELLPARMS cellParams = {0};
			rowParams.cbRow = sizeof(TABLEROWPARMS);
			rowParams.cbCell = sizeof(TABLECELLPARMS);
			rowParams.cRow = 1;
			// get the cell count at first
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				rowParams.iCell = static_cast<BYTE>(cellIndex);
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						pCellParams[cellIndex].dxWidth = newValue;
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_SETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					}
					HeapFree(GetProcessHeap(), 0, pCellParams);
				} else {
					hr = E_OUTOFMEMORY;
				}
			}
		}
	}
	return hr;
}