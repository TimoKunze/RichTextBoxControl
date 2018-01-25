// Table.cpp: A wrapper for the styling of a table.

#include "stdafx.h"
#include "Table.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP Table::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTable, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK Table::QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRange), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextRange2), queriedInterface)) {
		Table* pTable = reinterpret_cast<Table*>(pThis);
		return pTable->properties.pTableTextRange->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}

HRESULT CALLBACK Table::QueryITextRowInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRow), queriedInterface)) {
		Table* pTable = reinterpret_cast<Table*>(pThis);
		RichEditAPIVersionConstants apiVersion = reavUnknown;
		pTable->properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);
		if(apiVersion == reav41 || apiVersion >= reav75) {
			CComPtr<ITextRange> pRowTextRange = NULL;
			if(SUCCEEDED(pTable->properties.pTableTextRange->GetDuplicate(&pRowTextRange))) {
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


Table::Properties::~Properties()
{
	if(pOwnerRTB) {
		pOwnerRTB->Release();
		pOwnerRTB = NULL;
	}
	if(pTableTextRange) {
		pTableTextRange->Release();
		pTableTextRange = NULL;
	}
}

HWND Table::Properties::GetRTBHWnd(void)
{
	ATLASSUME(pOwnerRTB);

	OLE_HANDLE handle = NULL;
	pOwnerRTB->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void Table::Attach(ITextRange* pTableTextRange)
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

void Table::Detach(void)
{
	if(properties.pTableTextRange) {
		properties.pTableTextRange->Release();
		properties.pTableTextRange = NULL;
	}
}

void Table::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP Table::get_NestingLevel(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}
	RichEditAPIVersionConstants apiVersion = reavUnknown;
	properties.pOwnerRTB->get_RichEditAPIVersion(&apiVersion);

	// use the nesting level of the first row

	HRESULT hr = E_FAIL;
	*pValue = 1;
	BOOL useFallback = TRUE;

	CComPtr<ITextRange> pRowTextRange = NULL;
	if(SUCCEEDED(properties.pTableTextRange->GetDuplicate(&pRowTextRange))) {
		CComQIPtr<ITextRange2> pTextRange2 = pRowTextRange;
		if(pTextRange2) {
			LONG rowStart = 0;
			pTextRange2->GetStart(&rowStart);
			pTextRange2->SetEnd(rowStart);
			if(apiVersion == reav41 || apiVersion >= reav75) {
				CComPtr<ITextRow> pTextRow = NULL;
				hr = pTextRange2->GetRow(&pTextRow);
				if(SUCCEEDED(hr) && pTextRow) {
					hr = pTextRow->GetNestLevel(pValue);
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
			if(SUCCEEDED(properties.pTableTextRange->GetStart(&rowParams.cpStartRow))) {
				hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_GETTABLEPARMS, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
			}
			if(SUCCEEDED(hr)) {
				*pValue = rowParams.bTableLevel;
			}
		}
	}
	return hr;
}

STDMETHODIMP Table::get_Rows(IRichTableRows** ppRows)
{
	ATLASSERT_POINTER(ppRows, IRichTableRows*);
	if(!ppRows) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}

	ClassFactory::InitTableRows(properties.pTableTextRange, properties.pOwnerRTB, IID_IRichTableRows, reinterpret_cast<LPUNKNOWN*>(ppRows));
	return S_OK;
}

STDMETHODIMP Table::get_TextRange(IRichTextRange** ppTextRange)
{
	ATLASSERT_POINTER(ppTextRange, IRichTextRange*);
	if(!ppTextRange) {
		return E_POINTER;
	}
	if(!properties.pTableTextRange) {
		return CO_E_RELEASED;
	}

	CComPtr<ITextRange> pTextRange = NULL;
	HRESULT hr = properties.pTableTextRange->GetDuplicate(&pTextRange);
	if(SUCCEEDED(hr)) {
		ClassFactory::InitTextRange(pTextRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppTextRange));
	}
	return hr;
}