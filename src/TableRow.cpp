// TableRow.cpp: A wrapper for the styling of a table row.

#include "stdafx.h"
#include "TableRow.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TableRow::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTableRow, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TableRow::QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRange), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextRange2), queriedInterface)) {
		TableRow* pTableRow = reinterpret_cast<TableRow*>(pThis);
		return pTableRow->properties.pRowTextRange->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}

HRESULT CALLBACK TableRow::QueryITextRowInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRow), queriedInterface)) {
		TableRow* pTableRow = reinterpret_cast<TableRow*>(pThis);
		RichEditAPIVersionConstants apiVersion = reavUnknown;
		pTableRow->properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);
		if(apiVersion == reav41 || apiVersion >= reav75) {
			CComPtr<ITextRange> pRowTextRange = NULL;
			if(SUCCEEDED(pTableRow->properties.pRowTextRange->GetDuplicate(&pRowTextRange))) {
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


TableRow::Properties::~Properties()
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
}

HWND TableRow::Properties::GetRTBHWnd(void)
{
	ATLASSUME(pOwnerRTB);

	OLE_HANDLE handle = NULL;
	pOwnerRTB->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TableRow::Attach(ITextRange* pTableTextRange, ITextRange* pRowTextRange)
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

void TableRow::Detach(void)
{
	if(properties.pTableTextRange) {
		properties.pTableTextRange->Release();
		properties.pTableTextRange = NULL;
	}
	if(properties.pRowTextRange) {
		properties.pRowTextRange->Release();
		properties.pRowTextRange = NULL;
	}
}

void TableRow::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP TableRow::get_AllowPageBreakWithinRow(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
	*pValue = VARIANT_FALSE;
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
					LONG keepTogether = 0;
					hr = pTextRow->GetKeepTogether(&keepTogether);
					if(SUCCEEDED(hr)) {
						*pValue = BOOL2VARIANTBOOL(!keepTogether);
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
			rowParams.cRow = 1;
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				*pValue = BOOL2VARIANTBOOL(!rowParams.fKeep);
			}
		}
	}
	return hr;
}

STDMETHODIMP TableRow::put_AllowPageBreakWithinRow(VARIANT_BOOL newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
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
					hr = pTextRow->SetKeepTogether(!VARIANTBOOL2BOOL(newValue));
					if(SUCCEEDED(hr)) {
						hr = pTextRow->Apply(1, tomRowApplyDefault);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
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
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						rowParams.fKeep = !VARIANTBOOL2BOOL(newValue);
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

STDMETHODIMP TableRow::get_CanChange(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = S_OK;
	*pValue = bpvTrue;

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
					LONG canChange = tomFalse;
					hr = pTextRow->CanChange(&canChange);
					*pValue = static_cast<BooleanPropertyValueConstants>(canChange);
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableRow::get_Cells(IRichTableCells** ppCells)
{
	ATLASSERT_POINTER(ppCells, IRichTableCells*);
	if(!ppCells) {
		return E_POINTER;
	}

	ClassFactory::InitTableCells(properties.pTableTextRange, properties.pRowTextRange, properties.pOwnerRTB, IID_IRichTableCells, reinterpret_cast<LPUNKNOWN*>(ppCells));
	return S_OK;
}

STDMETHODIMP TableRow::get_HAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
	*pValue = halLeft;
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
					hr = pTextRow->GetAlignment(reinterpret_cast<LONG*>(pValue));
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
			rowParams.cRow = 1;
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				switch(rowParams.nAlignment) {
					case PFA_LEFT:
						*pValue = halLeft;
						break;
					case PFA_CENTER:
						*pValue = halCenter;
						break;
					case PFA_RIGHT:
						*pValue = halRight;
						break;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TableRow::put_HAlignment(HAlignmentConstants newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
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
					hr = pTextRow->SetAlignment(static_cast<LONG>(newValue));
					if(SUCCEEDED(hr)) {
						hr = pTextRow->Apply(1, tomRowApplyDefault);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
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
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						switch(newValue) {
							case halLeft:
								rowParams.nAlignment = PFA_LEFT;
								break;
							case halCenter:
								rowParams.nAlignment = PFA_CENTER;
								break;
							case halRight:
								rowParams.nAlignment = PFA_RIGHT;
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

STDMETHODIMP TableRow::get_Height(LONG* pValue)
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
	*pValue = 0;
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
					hr = pTextRow->GetHeight(pValue);
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
			rowParams.cRow = 1;
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				*pValue = rowParams.dyHeight;
			}
		}
	}
	return hr;
}

STDMETHODIMP TableRow::put_Height(LONG newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
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
					hr = pTextRow->SetHeight(newValue);
					if(SUCCEEDED(hr)) {
						hr = pTextRow->Apply(1, tomRowApplyDefault);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
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
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						rowParams.dyHeight = newValue;
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

STDMETHODIMP TableRow::get_HorizontalCellMargin(LONG* pValue)
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
	*pValue = 0;
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
					hr = pTextRow->GetCellMargin(pValue);
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
			rowParams.cRow = 1;
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				*pValue = rowParams.dxCellMargin;
			}
		}
	}
	return hr;
}

STDMETHODIMP TableRow::put_HorizontalCellMargin(LONG newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
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
					hr = pTextRow->SetCellMargin(newValue);
					if(SUCCEEDED(hr)) {
						hr = pTextRow->Apply(1, tomRowApplyDefault);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
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
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						rowParams.dxCellMargin = newValue;
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

STDMETHODIMP TableRow::get_Indent(LONG* pValue)
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
	*pValue = 0;
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
					hr = pTextRow->GetIndent(pValue);
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
			rowParams.cRow = 1;
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				*pValue = rowParams.dxIndent;
			}
		}
	}
	return hr;
}

STDMETHODIMP TableRow::put_Indent(LONG newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
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
					hr = pTextRow->SetIndent(newValue);
					if(SUCCEEDED(hr)) {
						hr = pTextRow->Apply(1, tomRowApplyDefault);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
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
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						rowParams.dxIndent = newValue;
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

STDMETHODIMP TableRow::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}

	// find the row within the table
	LONG thisRowStart = 0;
	LONG thisRowEnd = 0;
	HRESULT hr = properties.pRowTextRange->GetStart(&thisRowStart);
	if(FAILED(hr)) {
		return hr;
	}
	hr = properties.pRowTextRange->GetEnd(&thisRowEnd);
	if(FAILED(hr)) {
		return hr;
	}
	CComPtr<ITextRange> pRowTextRange = NULL;
	hr = properties.pTableTextRange->GetDuplicate(&pRowTextRange);
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
				LONG rowStart = 0;
				LONG rowEnd = 0;
				pRowTextRange->GetStart(&rowStart);
				pRowTextRange->GetEnd(&rowEnd);
				if(rowStart == thisRowStart && rowEnd == thisRowEnd) {
					foundRow = TRUE;
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
		*pValue = rowIndex - 1;
		return S_OK;
	} else {
		*pValue = 0;
		return E_FAIL;
	}
}

STDMETHODIMP TableRow::get_KeepTogetherWithNextRow(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
	*pValue = VARIANT_FALSE;
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
					LONG keepTogether = 0;
					hr = pTextRow->GetKeepWithNext(&keepTogether);
					if(SUCCEEDED(hr)) {
						*pValue = BOOL2VARIANTBOOL(keepTogether);
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
			rowParams.cRow = 1;
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				*pValue = BOOL2VARIANTBOOL(rowParams.fKeepFollow);
			}
		}
	}
	return hr;
}

STDMETHODIMP TableRow::put_KeepTogetherWithNextRow(VARIANT_BOOL newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
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
					hr = pTextRow->SetKeepWithNext(VARIANTBOOL2BOOL(newValue));
					if(SUCCEEDED(hr)) {
						hr = pTextRow->Apply(1, tomRowApplyDefault);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
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
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						rowParams.fKeepFollow = VARIANTBOOL2BOOL(newValue);
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

STDMETHODIMP TableRow::get_MasterCell(IRichTableCell** ppMasterCell)
{
	ATLASSERT_POINTER(ppMasterCell, IRichTableCell*);
	if(!ppMasterCell) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
	*ppMasterCell = NULL;
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
					LONG cacheCount = 0;
					hr = pTextRow->GetCellCountCache(&cacheCount);
					/*CAtlString s;
					s.Format(_T("%i 0x%X"), cacheCount, hr);
					MessageBox(NULL, s, _T("get_MasterCell"), MB_OK);*/
					if(SUCCEEDED(hr)) {
						useFallback = FALSE;
						if(cacheCount > 0) {
							CComPtr<IRichTableCells> pCells = NULL;
							hr = get_Cells(&pCells);
							if(SUCCEEDED(hr) && pCells) {
								hr = pCells->get_Item(cacheCount - 1, citIndex, ppMasterCell);
							}
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

STDMETHODIMP TableRow::putref_MasterCell(IRichTableCell* pNewMasterCell)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	LONG cellIndex = -1;
	if(pNewMasterCell) {
		pNewMasterCell->get_Index(&cellIndex);
		// TODO: Shouldn't we AddRef' pNewMasterCell?
	}

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
					LONG cacheCount = cellIndex + 1;
					if(cellIndex == -1) {
						CComPtr<IRichTableCells> pCells = NULL;
						hr = get_Cells(&pCells);
						if(SUCCEEDED(hr) && pCells) {
							hr = pCells->Count(&cacheCount);
						}
					}
					if(cacheCount > 0) {
						// TODO: Currently this does not fail, but also does not have any effect, not even on Windows 10.
						hr = pTextRow->SetCellCountCache(cacheCount);
						/*CAtlString s;
						s.Format(_T("%i 0x%X"), cacheCount, hr);
						MessageBox(NULL, s, _T("putref_MasterCell"), MB_OK);*/
						if(SUCCEEDED(hr)) {
							hr = pTextRow->Apply(1, tomRowApplyDefault);
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
		return E_NOTIMPL;
	}
	return hr;
}

STDMETHODIMP TableRow::get_ParentTable(IRichTable** ppTable)
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

STDMETHODIMP TableRow::get_RightToLeft(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
	*pValue = VARIANT_FALSE;
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
					LONG rtl = 0;
					hr = pTextRow->GetRTL(&rtl);
					if(SUCCEEDED(hr)) {
						*pValue = BOOL2VARIANTBOOL(rtl);
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
			rowParams.cRow = 1;
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				*pValue = BOOL2VARIANTBOOL(rowParams.fRTL);
			}
		}
	}
	return hr;
}

STDMETHODIMP TableRow::put_RightToLeft(VARIANT_BOOL newValue)
{
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;
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
					hr = pTextRow->SetRTL(VARIANTBOOL2BOOL(newValue));
					if(SUCCEEDED(hr)) {
						hr = pTextRow->Apply(1, tomRowApplyDefault);
						if(SUCCEEDED(hr)) {
							useFallback = FALSE;
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
			rowParams.cCell = 1;
			if(SUCCEEDED(properties.pRowTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, rowParams.cCell * sizeof(TABLECELLPARMS)));
				if(pCellParams) {
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
					if(SUCCEEDED(hr)) {
						rowParams.fRTL = VARIANTBOOL2BOOL(newValue);
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

STDMETHODIMP TableRow::get_TextRange(IRichTextRange** ppTextRange)
{
	ATLASSERT_POINTER(ppTextRange, IRichTextRange*);
	if(!ppTextRange) {
		return E_POINTER;
	}

	CComPtr<ITextRange> pTextRange = NULL;
	HRESULT hr = properties.pRowTextRange->GetDuplicate(&pTextRange);
	if(SUCCEEDED(hr)) {
		ClassFactory::InitTextRange(pTextRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppTextRange));
	}
	return hr;
}


STDMETHODIMP TableRow::Equals(IRichTableRow* pCompareAgainst, VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pCompareAgainst, IRichTableRow);
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pCompareAgainst || !pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange || !properties.pRowTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	HRESULT hr = E_FAIL;

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
				CComQIPtr<ITextRow> pCompareRow = pCompareAgainst;
				if(SUCCEEDED(hr) && pTextRow && pCompareRow) {
					LONG equal = tomFalse;
					hr = pTextRow->IsEqual(pCompareRow, &equal);
					*pValue = BOOL2VARIANTBOOL(equal == tomTrue);
				}
			}
		}
	}
	return hr;
}