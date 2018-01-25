// TextRange.cpp: A wrapper for a range of text.

#include "stdafx.h"
#include "TextRange.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TextRange::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTextRange, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TextRange::QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRange), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextRange2), queriedInterface)) {
		TextRange* pTextRange = reinterpret_cast<TextRange*>(pThis);
		return pTextRange->properties.pTextRange->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}


TextRange::Properties::~Properties()
{
	if(pOwnerRTB) {
		pOwnerRTB->Release();
		pOwnerRTB = NULL;
	}
	if(pTextRange) {
		pTextRange->Release();
		pTextRange = NULL;
	}
}

HWND TextRange::Properties::GetRTBHWnd(void)
{
	ATLASSUME(pOwnerRTB);

	OLE_HANDLE handle = NULL;
	pOwnerRTB->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void TextRange::Attach(ITextRange* pTextRange)
{
	if(properties.pTextRange) {
		properties.pTextRange->Release();
		properties.pTextRange = NULL;
	}
	if(pTextRange) {
		//pTextRange->QueryInterface(IID_PPV_ARGS(&properties.pTextRange));
		properties.pTextRange = pTextRange;
		properties.pTextRange->AddRef();
	}
}

void TextRange::Detach(void)
{
	if(properties.pTextRange) {
		properties.pTextRange->Release();
		properties.pTextRange = NULL;
	}
}

void TextRange::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP TextRange::get_CanEdit(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		LONG canEdit = tomFalse;
		hr = properties.pTextRange->CanEdit(&canEdit);
		*pValue = static_cast<BooleanPropertyValueConstants>(canEdit);
	}
	return hr;
}

STDMETHODIMP TextRange::get_EmbeddedObject(IRichOLEObject** ppOLEObject)
{
	ATLASSERT_POINTER(ppOLEObject, IRichOLEObject*);
	if(!ppOLEObject) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComPtr<IUnknown> pObject = NULL;
		hr = properties.pTextRange->GetEmbeddedObject(&pObject);
		if(pObject) {
			CComQIPtr<IOleObject> pOleObject = pObject;
			CComPtr<ITextRange> pTextRangeClone = NULL;
			properties.pTextRange->GetDuplicate(&pTextRangeClone);
			ClassFactory::InitOLEObject(pTextRangeClone, pOleObject, properties.pOwnerRTB, IID_IRichOLEObject, reinterpret_cast<LPUNKNOWN*>(ppOLEObject));
		} else {
			*ppOLEObject = NULL;
		}
	}
	return hr;
}

STDMETHODIMP TextRange::get_FirstChar(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->GetChar(pValue);
	}
	return hr;
}

STDMETHODIMP TextRange::put_FirstChar(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->SetChar(newValue);
	}
	return hr;
}

STDMETHODIMP TextRange::get_Font(IRichTextFont** ppTextFont)
{
	ATLASSERT_POINTER(ppTextFont, IRichTextFont*);
	if(!ppTextFont) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComPtr<ITextFont> pFont = NULL;
		hr = properties.pTextRange->GetFont(&pFont);
		ClassFactory::InitTextFont(pFont, IID_IRichTextFont, reinterpret_cast<LPUNKNOWN*>(ppTextFont));
	}
	return hr;
}

STDMETHODIMP TextRange::put_Font(IRichTextFont* pNewTextFont)
{
	if(pNewTextFont) {
		CComPtr<ITextFont> pFont = NULL;
		if(SUCCEEDED(pNewTextFont->QueryInterface(__uuidof(ITextFont), reinterpret_cast<LPVOID*>(&pFont)))) {
			if(properties.pTextRange) {
				return properties.pTextRange->SetFont(pFont);
			} else {
				return E_FAIL;
			}
		}
	}
	// invalid value - raise VB runtime error 380
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
}

STDMETHODIMP TextRange::get_Paragraph(IRichTextParagraph** ppTextParagraph)
{
	ATLASSERT_POINTER(ppTextParagraph, IRichTextParagraph*);
	if(!ppTextParagraph) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComPtr<ITextPara> pParagraph = NULL;
		hr = properties.pTextRange->GetPara(&pParagraph);
		ClassFactory::InitTextParagraph(pParagraph, IID_IRichTextParagraph, reinterpret_cast<LPUNKNOWN*>(ppTextParagraph));
	}
	return hr;
}

STDMETHODIMP TextRange::put_Paragraph(IRichTextParagraph* pNewTextParagraph)
{
	if(pNewTextParagraph) {
		CComPtr<ITextPara> pParagraph = NULL;
		if(SUCCEEDED(pNewTextParagraph->QueryInterface(__uuidof(ITextPara), reinterpret_cast<LPVOID*>(&pParagraph)))) {
			if(properties.pTextRange) {
				return properties.pTextRange->SetPara(pParagraph);
			} else {
				return E_FAIL;
			}
		}
	}
	// invalid value - raise VB runtime error 380
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
}

STDMETHODIMP TextRange::get_RangeEnd(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->GetEnd(pValue);
	}
	return hr;
}

STDMETHODIMP TextRange::put_RangeEnd(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->SetEnd(newValue);
	}
	return hr;
}

STDMETHODIMP TextRange::get_RangeLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		LONG start = 0;
		LONG end = 0;
		hr = properties.pTextRange->GetStart(&start);
		if(SUCCEEDED(hr)) {
			hr = properties.pTextRange->GetEnd(&end);
			if(SUCCEEDED(hr)) {
				*pValue = end - start;
			}
		}
	}
	return hr;
}

STDMETHODIMP TextRange::get_RangeStart(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->GetStart(pValue);
	}
	return hr;
}

STDMETHODIMP TextRange::put_RangeStart(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->SetStart(newValue);
	}
	return hr;
}

STDMETHODIMP TextRange::get_RichText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		VARIANT v;
		VariantClear(&v);
		IUnknown* pDataObjectUnknown = NULL;
		v.ppunkVal = &pDataObjectUnknown;
		v.vt = VT_UNKNOWN | VT_BYREF;
		hr = properties.pTextRange->Copy(&v);
		if(v.vt == (VT_UNKNOWN | VT_BYREF) && pDataObjectUnknown) {
			CComQIPtr<IDataObject> pDataObject = pDataObjectUnknown;
			if(pDataObject) {
				FORMATETC format = {static_cast<CLIPFORMAT>(RegisterClipboardFormat(CF_RTF)), NULL, DVASPECT_CONTENT, -1, 0};
				format.tymed = TYMED_HGLOBAL;
				STGMEDIUM storageMedium = {0};
				hr = pDataObject->GetData(&format, &storageMedium);
				if(storageMedium.tymed & TYMED_HGLOBAL) {
					LPSTR pText = reinterpret_cast<LPSTR>(GlobalLock(storageMedium.hGlobal));
					*pValue = _bstr_t(pText).Detach();
					GlobalUnlock(storageMedium.hGlobal);
				}
				ReleaseStgMedium(&storageMedium);
			}
		} else {
			hr = E_FAIL;
		}
		VariantClear(&v);
	}
	return hr;
}

STDMETHODIMP TextRange::put_RichText(BSTR newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		//pOLEDataObjectObj->SetOwner(NULL);
		CComQIPtr<IDataObject> pDataObject = pOLEDataObjectObj;
		pOLEDataObjectObj->Release();
		if(pDataObject) {
			FORMATETC format = {static_cast<CLIPFORMAT>(RegisterClipboardFormat(CF_RTF)), NULL, DVASPECT_CONTENT, -1, 0};
			STGMEDIUM storageMedium = {0};
			format.tymed = storageMedium.tymed = TYMED_HGLOBAL;
			int textSize = _bstr_t(newValue).length() * sizeof(CHAR);
			storageMedium.hGlobal = GlobalAlloc(GPTR, textSize + sizeof(CHAR));
			if(storageMedium.hGlobal) {
				LPSTR pText = reinterpret_cast<LPSTR>(GlobalLock(storageMedium.hGlobal));
				CopyMemory(pText, CW2A(newValue), textSize);
				GlobalUnlock(storageMedium.hGlobal);
				hr = pDataObject->SetData(&format, &storageMedium, TRUE);
				if(SUCCEEDED(hr)) {
					VARIANT v;
					VariantClear(&v);
					v.punkVal = pDataObject.Detach();
					v.vt = VT_UNKNOWN;
					hr = properties.pTextRange->Paste(&v, format.cfFormat);
					VariantClear(&v);
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP TextRange::get_StoryLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->GetStoryLength(pValue);
	}
	return hr;
}

STDMETHODIMP TextRange::get_StoryType(StoryTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, StoryTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		long type = tomUnknownStory;
		hr = properties.pTextRange->GetStoryType(&type);
		*pValue = static_cast<StoryTypeConstants>(type);
	}
	return hr;
}

STDMETHODIMP TextRange::get_SubRanges(IRichTextSubRanges** ppSubRanges)
{
	ATLASSERT_POINTER(ppSubRanges, IRichTextSubRanges*);
	if(!ppSubRanges) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComPtr<ITextRange> pTextRangeClone = NULL;
		properties.pTextRange->GetDuplicate(&pTextRangeClone);
		ClassFactory::InitTextSubRanges(pTextRangeClone, properties.pOwnerRTB, IID_IRichTextSubRanges, reinterpret_cast<LPUNKNOWN*>(ppSubRanges));
	}
	return hr;
}

STDMETHODIMP TextRange::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->GetText(pValue);
	}
	return hr;
}

STDMETHODIMP TextRange::put_Text(BSTR newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->SetText(newValue);
	}
	return hr;
}

STDMETHODIMP TextRange::get_UnitIndex(UnitConstants unit, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->GetIndex(unit, pValue);
	}
	return hr;
}

STDMETHODIMP TextRange::get_URL(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
		if(pTextRange2) {
			hr = pTextRange2->GetURL(pValue);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextRange::put_URL(BSTR newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
		if(pTextRange2) {
			if(newValue && SysStringLen(newValue) > 0) {
				LPTSTR pURL;
				#ifdef UNICODE
					pURL = OLE2W(newValue);
				#else
					COLE2T converter(newValue);
					pURL = converter;
				#endif
				CString tmp = pURL;
				if(tmp.GetAt(0) != TEXT('\"')) {
					tmp = TEXT('\"') + tmp;
				}
				if(tmp.GetAt(tmp.GetLength() - 1) != TEXT('\"')) {
					tmp = tmp + TEXT('\"');
				}
				BSTR v = tmp.AllocSysString();
				hr = pTextRange2->SetURL(v);
				SysFreeString(v);
			} else {
				hr = pTextRange2->SetURL(NULL);
			}
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}


STDMETHODIMP TextRange::BuildDownMath(BuildUpMathConstants flags, VARIANT_BOOL* pDidAnyChanges)
{
	ATLASSERT_POINTER(pDidAnyChanges, VARIANT_BOOL);
	if(!pDidAnyChanges) {
		return E_POINTER;
	}
	ATLASSERT(properties.pTextRange);
	if(!properties.pTextRange) {
		return CO_E_RELEASED;
	}

	*pDidAnyChanges = VARIANT_FALSE;
	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(pTextRange2) {
		HRESULT hr = pTextRange2->Linearize(flags);
		if(hr == NOERROR) {
			*pDidAnyChanges = VARIANT_TRUE;
			hr = S_OK;
		} else if(hr == S_FALSE) {
			hr = S_OK;
		}
		return hr;
	} else {
		return E_NOINTERFACE;
	}
}

STDMETHODIMP TextRange::BuildUpMath(BuildUpMathConstants flags, VARIANT_BOOL* pDidAnyChanges)
{
	ATLASSERT_POINTER(pDidAnyChanges, VARIANT_BOOL);
	if(!pDidAnyChanges) {
		return E_POINTER;
	}
	ATLASSERT(properties.pTextRange);
	if(!properties.pTextRange) {
		return CO_E_RELEASED;
	}

	*pDidAnyChanges = VARIANT_FALSE;
	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(pTextRange2) {
		HRESULT hr = pTextRange2->BuildUpMath(flags);
		if(hr == NOERROR) {
			*pDidAnyChanges = VARIANT_TRUE;
			hr = S_OK;
		} else if(hr == S_FALSE) {
			hr = S_OK;
		}
		return hr;
	} else {
		return E_NOINTERFACE;
	}
}

STDMETHODIMP TextRange::CanPaste(IOLEDataObject* pOLEDataObject/* = NULL*/, LONG formatID/* = 0*/, VARIANT_BOOL* pCanPaste/* = NULL*/)
{
	ATLASSERT_POINTER(pCanPaste, VARIANT_BOOL);
	if(!pCanPaste) {
		return E_POINTER;
	}
	if(formatID == 0xffffbf01) {
		formatID = RegisterClipboardFormat(CF_RTF);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		LONG canPaste = tomFalse;
		BOOL useFallback = TRUE;
		if(pOLEDataObject) {
			CComQIPtr<IDataObject> pDataObject = pOLEDataObject;
			if(pDataObject) {
				useFallback = FALSE;
				VARIANT v;
				VariantClear(&v);
				pDataObject->QueryInterface(IID_PPV_ARGS(&v.punkVal));
				v.vt = VT_UNKNOWN;
				hr = properties.pTextRange->CanPaste(&v, formatID, &canPaste);
				VariantClear(&v);
			}
		}
		if(useFallback) {
			hr = properties.pTextRange->CanPaste(NULL, formatID, &canPaste);
		}
		if(SUCCEEDED(hr)) {
			// there does not seem to be a difference between hr and canPaste
			*pCanPaste = BOOL2VARIANTBOOL(hr == S_OK);
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP TextRange::ChangeCase(CaseConstants newCase)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->ChangeCase(newCase);
	}
	return hr;
}

STDMETHODIMP TextRange::Clone(IRichTextRange** ppClone)
{
	ATLASSERT_POINTER(ppClone, IRichTextRange*);
	if(!ppClone) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	*ppClone = NULL;
	if(properties.pTextRange) {
		CComPtr<ITextRange> pRange = NULL;
		hr = properties.pTextRange->GetDuplicate(&pRange);
		ClassFactory::InitTextRange(pRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppClone));
	}
	return hr;
}

STDMETHODIMP TextRange::Collapse(VARIANT_BOOL collapseToStart/* = VARIANT_TRUE*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->Collapse(VARIANTBOOL2BOOL(collapseToStart) ? tomStart : tomEnd);
		*pSucceeded = BOOL2VARIANTBOOL(SUCCEEDED(hr));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP TextRange::ContainsRange(IRichTextRange* pCompareAgainst, VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pCompareAgainst, IRichTextRange);
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pCompareAgainst || !pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComQIPtr<ITextRange> pRange = pCompareAgainst;
		LONG contains = tomFalse;
		hr = pRange->InRange(properties.pTextRange, &contains);
		*pValue = BOOL2VARIANTBOOL(contains == tomTrue);
	}
	return hr;
}

STDMETHODIMP TextRange::Copy(VARIANT* pOLEDataObject/* = NULL*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		if(pOLEDataObject && pOLEDataObject->vt == (VT_DISPATCH | VT_BYREF)) {
			VARIANT v;
			VariantClear(&v);
			IUnknown* pDataObjectUnknown = NULL;
			v.ppunkVal = &pDataObjectUnknown;
			v.vt = VT_UNKNOWN | VT_BYREF;
			hr = properties.pTextRange->Copy(&v);
			if(v.vt == (VT_UNKNOWN | VT_BYREF) && pDataObjectUnknown) {
				CComQIPtr<IDataObject> pDataObject = pDataObjectUnknown;
				ClassFactory::InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(pOLEDataObject->ppdispVal));
			}
			VariantClear(&v);
		} else {
			hr = properties.pTextRange->Copy(NULL);
		}
		*pSucceeded = BOOL2VARIANTBOOL(SUCCEEDED(hr));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP TextRange::CopyRichTextFromTextRange(IRichTextRange* pSourceObject, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSourceObject, IRichTextRange);
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSourceObject || !pSucceeded) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComQIPtr<ITextRange> pSourceRange = pSourceObject;
		if(pSourceRange) {
			CComPtr<ITextRange> pSourceFormattedText = NULL;
			hr = pSourceRange->GetFormattedText(&pSourceFormattedText);
			if(SUCCEEDED(hr)) {
				hr = properties.pTextRange->SetFormattedText(pSourceFormattedText);
			}
		}
		*pSucceeded = BOOL2VARIANTBOOL(SUCCEEDED(hr));
	}
	return hr;
}

STDMETHODIMP TextRange::Cut(VARIANT* pOLEDataObject/* = NULL*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		if(pOLEDataObject && pOLEDataObject->vt == (VT_DISPATCH | VT_BYREF)) {
			VARIANT v;
			VariantClear(&v);
			IUnknown* pDataObjectUnknown = NULL;
			v.ppunkVal = &pDataObjectUnknown;
			v.vt = VT_UNKNOWN | VT_BYREF;
			hr = properties.pTextRange->Cut(&v);
			if(v.vt == (VT_UNKNOWN | VT_BYREF) && pDataObjectUnknown) {
				CComQIPtr<IDataObject> pDataObject = pDataObjectUnknown;
				ClassFactory::InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(pOLEDataObject->ppdispVal));
			}
			VariantClear(&v);
		} else {
			hr = properties.pTextRange->Copy(NULL);
		}
		*pSucceeded = BOOL2VARIANTBOOL(SUCCEEDED(hr));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP TextRange::Delete(UnitConstants unit/* = uCharacter*/, LONG count/* = 0*/, LONG* pDelta/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->Delete(unit, count, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::Equals(IRichTextRange* pCompareAgainst, VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pCompareAgainst, IRichTextRange);
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pCompareAgainst || !pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComQIPtr<ITextRange> pRange = pCompareAgainst;
		LONG equal = tomFalse;
		hr = properties.pTextRange->IsEqual(pRange, &equal);
		*pValue = BOOL2VARIANTBOOL(equal == tomTrue);
	}
	return hr;
}

STDMETHODIMP TextRange::EqualsStory(IRichTextRange* pCompareAgainst, VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pCompareAgainst, IRichTextRange);
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pCompareAgainst || !pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComQIPtr<ITextRange> pRange = pCompareAgainst;
		LONG equal = tomFalse;
		hr = properties.pTextRange->InStory(pRange, &equal);
		*pValue = BOOL2VARIANTBOOL(equal == tomTrue);
	}
	return hr;
}

STDMETHODIMP TextRange::ExpandToUnit(UnitConstants unit, LONG* pDelta)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->Expand(unit, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::FindText(BSTR searchFor, SearchDirectionConstants searchDirectionAndMaxCharsToPass/* = sdForward*/, SearchModeConstants searchMode/* = smDefault*/, MoveRangeBoundariesConstants moveRangeBoundaries/* = mrbMoveBoth*/, LONG* pLengthMatched/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		switch(moveRangeBoundaries) {
			case mrbMoveBoth:
				hr = properties.pTextRange->FindText(searchFor, searchDirectionAndMaxCharsToPass, searchMode, pLengthMatched);
				break;
			case mrbMoveStartOnly:
				hr = properties.pTextRange->FindTextStart(searchFor, searchDirectionAndMaxCharsToPass, searchMode, pLengthMatched);
				break;
			case mrbMoveEndOnly:
				hr = properties.pTextRange->FindTextEnd(searchFor, searchDirectionAndMaxCharsToPass, searchMode, pLengthMatched);
				break;
		}
	}
	return hr;
}

STDMETHODIMP TextRange::GetEndPosition(HorizontalPositionConstants horizontalPosition, VerticalPositionConstants verticalPosition, RangePositionConstants flags, OLE_XPOS_PIXELS* pX, OLE_YPOS_PIXELS* pY, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		long type = tomEnd | horizontalPosition | verticalPosition | flags;
		hr = properties.pTextRange->GetPoint(type, pX, pY);
		// NOTE: hr will be S_FALSE if the coordinates are not within the client rectangle and tomAllowOffClient has not been specified.
		*pSucceeded = BOOL2VARIANTBOOL(hr == S_OK);
		// Don't set hr to S_OK, so that S_FALSE can be distinguished from a real error.
	}
	return hr;
}

STDMETHODIMP TextRange::GetStartPosition(HorizontalPositionConstants horizontalPosition, VerticalPositionConstants verticalPosition, RangePositionConstants flags, OLE_XPOS_PIXELS* pX, OLE_YPOS_PIXELS* pY, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		long type = tomStart | horizontalPosition | verticalPosition | flags;
		hr = properties.pTextRange->GetPoint(type, pX, pY);
		// NOTE: hr will be S_FALSE if the coordinates are not within the client rectangle and tomAllowOffClient has not been specified.
		*pSucceeded = BOOL2VARIANTBOOL(hr == S_OK);
		// Don't set hr to S_OK, so that S_FALSE can be distinguished from a real error.
	}
	return hr;
}

STDMETHODIMP TextRange::IsWithinTable(VARIANT* pTable/* = NULL*/, VARIANT* pTableRow/* = NULL*/, VARIANT* pTableCell/* = NULL*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	ATLASSERT(properties.pTextRange);
	if(!properties.pTextRange) {
		return CO_E_RELEASED;
	}

	*pValue = VARIANT_FALSE;

	LONG rangeStart = -1;
	LONG rangeEnd = -1;
	properties.pTextRange->GetStart(&rangeStart);
	properties.pTextRange->GetEnd(&rangeEnd);

	CComPtr<ITextRange> pTableTextRange = NULL;
	HRESULT hr = properties.pTextRange->GetDuplicate(&pTableTextRange);
	if(FAILED(hr)) {
		return hr;
	}

	LONG startOfTable = -1;
	LONG endOfTable = -1;

	LONG start = -1;
	pTableTextRange->GetStart(&start);
	LONG previousStart = start;
	BOOL mayTryToGetBoundary = FALSE;
	// unfortunately this will walk into nested tables
	while(pTableTextRange->Move(tomCell, -1, NULL) == S_OK) {
		pTableTextRange->GetStart(&start);
		if(start == previousStart) {
			break;
		}
		mayTryToGetBoundary = TRUE;
		previousStart = start;
	}
	if(!mayTryToGetBoundary) {
		// maybe the range start is within the first cell, so go one forward and then back again to check this
		if(pTableTextRange->Move(tomCell, 1, NULL) == S_OK) {
			pTableTextRange->Move(tomCell, -1, NULL);
			mayTryToGetBoundary = TRUE;
		} else {
			// maybe the range start is equal to the table start
			// TODO: Probably this won't work if the first cell contains a nested table.
			pTableTextRange->SetStart(rangeStart + 1);
			// it might be better to just set mayTryToGetBoundary to TRUE here
			if(pTableTextRange->Move(tomCell, 1, NULL) == S_OK) {
				pTableTextRange->Move(tomCell, -1, NULL);
				mayTryToGetBoundary = TRUE;
			}
		}
	}
	if(mayTryToGetBoundary) {
		// Here we should at least be in the table's first row.
		HRESULT hr2 = pTableTextRange->StartOf(tomRow, tomExtend, NULL);
		if(hr2 == S_OK) {
			// everything looks fine
			pTableTextRange->GetStart(&startOfTable);
		} else if(hr2 == S_FALSE) {
			// this happens for instance if we are in cell 2 and there is a nested table in the first cell
			pTableTextRange->GetStart(&start);
			// NOTE: This trick seems to work even if there are multiple levels of nested tables in the first cell.
			pTableTextRange->SetStart(start - 1);
			if(pTableTextRange->StartOf(tomRow, tomExtend, NULL) == S_OK) {
				pTableTextRange->GetStart(&startOfTable);
			}
		}
	}

	if(startOfTable >= 0) {
		// We have the start of the table now and it seems quite reliable. Find the table's end now.
		pTableTextRange->SetStart(rangeEnd);
		start = rangeEnd;
		previousStart = start - 1;
		mayTryToGetBoundary = FALSE;
		// unfortunately this will walk into nested tables
		while(pTableTextRange->Move(tomCell, 1, NULL) == S_OK) {
			pTableTextRange->GetStart(&start);
			if(start == previousStart) {
				break;
			}
			// EndOf will fail if we already are at the end
			pTableTextRange->SetEnd(previousStart);
			mayTryToGetBoundary = TRUE;
			previousStart = start;
		}
		if(!mayTryToGetBoundary) {
			// maybe the range end is within the last cell
			if(pTableTextRange->Move(tomCell, -1, NULL) == S_OK) {
				pTableTextRange->Move(tomCell, 1, NULL);
				mayTryToGetBoundary = TRUE;
			} else {
				// maybe the range end is equal to the table end
				pTableTextRange->SetStart(rangeEnd - 1);
				if(pTableTextRange->Move(tomCell, -1, NULL) == S_OK) {
					pTableTextRange->Move(tomCell, 1, NULL);
					mayTryToGetBoundary = TRUE;
				}
			}
		}
		if(mayTryToGetBoundary) {
			// unfortunately the loop would walk out of the nested table, so check the nesting levels
			LONG ourNestingLevel = -1;
			CComPtr<ITextRange> pDuplicate = NULL;
			if(SUCCEEDED(pTableTextRange->GetDuplicate(&pDuplicate)) && SUCCEEDED(pDuplicate->SetStart(startOfTable))) {
				CComPtr<IRichTable> pOurTable = ClassFactory::InitTable(pDuplicate, properties.pOwnerRTB);
				pOurTable->get_NestingLevel(&ourNestingLevel);
				do {
					hr = pTableTextRange->EndOf(tomRow, tomExtend, NULL);
					if(hr == S_OK) {
						CComPtr<ITextRange> pDuplicate = NULL;
						if(SUCCEEDED(pTableTextRange->GetDuplicate(&pDuplicate)) && SUCCEEDED(pDuplicate->Collapse(tomEnd))) {
							CComPtr<IRichTable> pTableOfEnd = ClassFactory::InitTable(pDuplicate, properties.pOwnerRTB);
							LONG endNestingLevel = -1;
							pTableOfEnd->get_NestingLevel(&endNestingLevel);
							if(endNestingLevel >= ourNestingLevel) {
								pTableTextRange->GetEnd(&endOfTable);
								pTableTextRange->SetEnd(endOfTable + 1);
							} else {
								// we're in a nested table and we've already been in the last row
								pTableTextRange->GetEnd(&endOfTable);
								hr = S_FALSE;
							}
						} else {
							hr = E_FAIL;
							endOfTable = -1;
						}
					}
				} while(hr == S_OK);
			}
		}
	}

	if(startOfTable >= 0 && endOfTable >= 0) {
		*pValue = VARIANT_TRUE;
		pTableTextRange->SetStart(startOfTable);
		pTableTextRange->SetEnd(endOfTable);
	} else {
		pTableTextRange = NULL;
	}

	CComPtr<ITextRange> pRowTextRange = NULL;
	if(pTableTextRange && ((pTableRow && pTableRow->vt == (VT_DISPATCH | VT_BYREF)) || (pTableCell && pTableCell->vt == (VT_DISPATCH | VT_BYREF)))) {
		// now search for the table row
		hr = pTableTextRange->GetDuplicate(&pRowTextRange);
		if(FAILED(hr)) {
			return hr;
		}

		// start with the first row
		LONG startOfNextRowCandidate = startOfTable;
		BOOL foundRow = FALSE;
		do {
			// set the cursor to the current row
			pRowTextRange->SetStart(startOfNextRowCandidate);
			pRowTextRange->SetEnd(startOfNextRowCandidate);
			hr = pRowTextRange->EndOf(tomRow, tomExtend, NULL);
			if(hr == S_OK) {
				// this has been a row, retrieve the position at which the next row would start
				pRowTextRange->GetEnd(&startOfNextRowCandidate);
				if(startOfNextRowCandidate <= endOfTable) {
					// we're still inside the table, so count this row
					LONG rowStart = 0;
					LONG rowEnd = 0;
					pRowTextRange->GetStart(&rowStart);
					pRowTextRange->GetEnd(&rowEnd);
					if(rowStart <= rangeStart && rangeStart <= rowEnd) {
						foundRow = (rowStart <= rangeEnd && rangeEnd <= rowEnd);
						if(foundRow) {
							pRowTextRange->SetStart(rowStart);
							pRowTextRange->SetEnd(rowEnd);
						}
						break;
					}
				}
				if(startOfNextRowCandidate >= endOfTable) {
					// we reached the end of the table
					break;
				}
			} else {
				// this has not been a row, so don't count it and stop
				break;
			}
		} while(!foundRow);
		if(!foundRow) {
			pRowTextRange = NULL;
		}
	}

	CComPtr<ITextRange> pCellTextRange = NULL;
	if(pTableTextRange && pRowTextRange && pTableCell && pTableCell->vt == (VT_DISPATCH | VT_BYREF)) {
		// now search for the table cell
		hr = pRowTextRange->GetDuplicate(&pCellTextRange);
		if(FAILED(hr)) {
			return hr;
		}
		LONG rowStart = 0;
		LONG rowEnd = 0;
		pRowTextRange->GetStart(&rowStart);
		pRowTextRange->GetEnd(&rowEnd);

		// NOTE: StartOf(tomCell, tomExtend) and EndOf(tomCell, tomExtend) do not really work on Windows 7.

		// start with the first cell
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
					if(cellStart <= rangeStart && rangeStart <= cellEnd) {
						foundCell = (cellStart <= rangeEnd && rangeEnd <= cellEnd);
						if(foundCell) {
							pCellTextRange->SetStart(cellStart);
							pCellTextRange->SetEnd(cellEnd);
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
		if(!foundCell) {
			pCellTextRange = NULL;
		}
	}

	if(pTableTextRange && pTable && pTable->vt == (VT_DISPATCH | VT_BYREF)) {
		ClassFactory::InitTable(pTableTextRange, properties.pOwnerRTB, IID_IRichTable, reinterpret_cast<LPUNKNOWN*>(pTable->ppdispVal));
	}
	if(pTableTextRange && pRowTextRange && pTableRow && pTableRow->vt == (VT_DISPATCH | VT_BYREF)) {
		ClassFactory::InitTableRow(pTableTextRange, pRowTextRange, properties.pOwnerRTB, IID_IRichTableRow, reinterpret_cast<LPUNKNOWN*>(pTableRow->ppdispVal));
	}
	if(pTableTextRange && pRowTextRange && pCellTextRange && pTableCell && pTableCell->vt == (VT_DISPATCH | VT_BYREF)) {
		ClassFactory::InitTableCell(pTableTextRange, pRowTextRange, pCellTextRange, properties.pOwnerRTB, IID_IRichTableCell, reinterpret_cast<LPUNKNOWN*>(pTableCell->ppdispVal));
	}
	return S_OK;
}

STDMETHODIMP TextRange::MoveEndByUnit(UnitConstants unit, LONG count, LONG* pDelta)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->MoveEnd(unit, count, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveEndToPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL doNotMoveStart)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->SetPoint(x, y, tomEnd, (VARIANTBOOL2BOOL(doNotMoveStart) ? tomExtend : tomMove));
	}
	return hr;
}

STDMETHODIMP TextRange::MoveEndToEndOfUnit(UnitConstants unit, VARIANT_BOOL doNotMoveStart, LONG* pDelta)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->EndOf(unit, (VARIANTBOOL2BOOL(doNotMoveStart) ? tomExtend : tomMove), pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveEndUntil(VARIANT* pCharacterSet, LONG maximumCharactersToPass/* = tomForward*/, LONG* pDelta/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->MoveEndUntil(pCharacterSet, maximumCharactersToPass, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveEndWhile(VARIANT* pCharacterSet, LONG maximumCharactersToPass/* = tomForward*/, LONG* pDelta/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->MoveEndWhile(pCharacterSet, maximumCharactersToPass, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveInsertionPointByUnit(UnitConstants unit, LONG count, LONG* pDelta)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->Move(unit, count, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveInsertionPointUntil(VARIANT* pCharacterSet, LONG maximumCharactersToPass/* = tomForward*/, LONG* pDelta/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->MoveUntil(pCharacterSet, maximumCharactersToPass, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveInsertionPointWhile(VARIANT* pCharacterSet, LONG maximumCharactersToPass/* = tomForward*/, LONG* pDelta/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->MoveWhile(pCharacterSet, maximumCharactersToPass, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveStartByUnit(UnitConstants unit, LONG count, LONG* pDelta)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->MoveStart(unit, count, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveStartToPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL doNotMoveEnd)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->SetPoint(x, y, tomStart, (VARIANTBOOL2BOOL(doNotMoveEnd) ? tomExtend : tomMove));
	}
	return hr;
}

STDMETHODIMP TextRange::MoveStartToStartOfUnit(UnitConstants unit, VARIANT_BOOL doNotMoveEnd, LONG* pDelta)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->StartOf(unit, (VARIANTBOOL2BOOL(doNotMoveEnd) ? tomExtend : tomMove), pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveStartUntil(VARIANT* pCharacterSet, LONG maximumCharactersToPass/* = tomForward*/, LONG* pDelta/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->MoveStartUntil(pCharacterSet, maximumCharactersToPass, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveStartWhile(VARIANT* pCharacterSet, LONG maximumCharactersToPass/* = tomForward*/, LONG* pDelta/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->MoveStartWhile(pCharacterSet, maximumCharactersToPass, pDelta);
	}
	return hr;
}

STDMETHODIMP TextRange::MoveToUnitIndex(UnitConstants unit, LONG index, VARIANT_BOOL setRangeToEntireUnit/* = VARIANT_FALSE*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->SetIndex(unit, index, (VARIANTBOOL2BOOL(setRangeToEntireUnit) ? tomExtend : tomMove));
		*pSucceeded = BOOL2VARIANTBOOL(hr == S_OK);
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP TextRange::Paste(IOLEDataObject* pOLEDataObject/* = NULL*/, LONG formatID/* = 0*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	if(formatID == 0xffffbf01) {
		formatID = RegisterClipboardFormat(CF_RTF);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		BOOL useFallback = TRUE;
		if(pOLEDataObject) {
			CComQIPtr<IDataObject> pDataObject = pOLEDataObject;
			if(pDataObject) {
				useFallback = FALSE;
				VARIANT v;
				VariantInit(&v);
				pDataObject->QueryInterface(IID_PPV_ARGS(&v.punkVal));
				v.vt = VT_UNKNOWN;
				hr = properties.pTextRange->Paste(&v, formatID);
				VariantClear(&v);
			}
		}
		if(useFallback) {
			hr = properties.pTextRange->Paste(NULL, formatID);
		}
		*pSucceeded = BOOL2VARIANTBOOL(SUCCEEDED(hr));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP TextRange::ReplaceWithTable(LONG columnCount, LONG rowCount, VARIANT_BOOL allowIndividualCellStyles/* = VARIANT_TRUE*/, SHORT borderSize/* = 0*/, HAlignmentConstants horizontalRowAlignment/* = halLeft*/, VAlignmentConstants verticalCellAlignment/* = valCenter*/, LONG rowIndent/* = -1*/, LONG columnWidth/* = -1*/, LONG horizontalCellMargin/* = -1*/, LONG rowHeight/* = -1*/, IRichTable** ppAddedTable/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedTable, IRichTable*);
	if(!ppAddedTable) {
		return E_POINTER;
	}

	if(columnCount <= 0 || rowCount <= 0 || borderSize < -1 || rowIndent < -1 || columnWidth < -1 || horizontalCellMargin < -1 || rowHeight < -1) {
		return E_INVALIDARG;
	}
	if(horizontalRowAlignment < halLeft || horizontalRowAlignment > halRight) {
		return E_INVALIDARG;
	}
	if(verticalCellAlignment < valTop || verticalCellAlignment > valBottom) {
		return E_INVALIDARG;
	}

	*ppAddedTable = NULL;
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
	if(rowIndent == -1) {
		rowIndent = twipsPerPixelX * 35 / 10;
	}
	if(horizontalCellMargin == -1) {
		horizontalCellMargin = twipsPerPixelX * 48 / 10;
	}
	if(borderSize == -1) {
		borderSize = static_cast<SHORT>(twipsPerPixelX);
	}

	BOOL useFallback = TRUE;
	LONG rangeStart = 0;

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		properties.pTextRange->GetStart(&rangeStart);
		CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
		if(pTextRange2) {
			// auto-fit can be:
			// 0 - disabled
			// 1 - enabled, may be overridden by fixed width specifications
			#if FALSE
				// Does not work well, e.g. the column width is too wide for autoFit=1 and the cell properties are not applied.
				// Also it always returns E_INVALIDARG if the MDIWordpad sample is maximized (both windows - doc and frame).
				long autoFit = (columnWidth == -1 ? 1 : 0);
				hr = pTextRange2->InsertTable(columnCount, rowCount, autoFit);
				/*CAtlString s;
				s.Format(_T("0x%X"), hr);
				MessageBox(NULL, s, _T("InsertTable"), MB_OK);*/
				if(SUCCEEDED(hr)) {
					useFallback = FALSE;
					ATLASSERT(FALSE);
					CComPtr<ITextRow> pTextRow = NULL;
					if(SUCCEEDED(pTextRange2->GetRow(&pTextRow))) {
						if(VARIANTBOOL2BOOL(allowIndividualCellStyles)) {
							// does not seem to work anyway, it has no effect
							pTextRow->SetCellCountCache(columnCount);
						}
						pTextRow->SetAlignment(static_cast<LONG>(horizontalRowAlignment));
						pTextRow->SetCellMargin(horizontalCellMargin);
						pTextRow->SetIndent(rowIndent);
						pTextRow->SetHeight(rowHeight);
						pTextRow->SetCellIndex(0);
						pTextRow->SetCellWidth(columnWidth);
						pTextRow->SetCellAlignment(verticalCellAlignment);
						pTextRow->SetCellBorderWidths(borderSize, borderSize, borderSize, borderSize);
						LONG backColor1 = RGB(255, 255, 255);
						CComPtr<ITextFont> pTextFont = NULL;
						if(SUCCEEDED(pTextRange2->GetFont(&pTextFont))) {
							pTextFont->GetBackColor(&backColor1);
						}
						pTextRow->SetCellColorBack(backColor1);
						if(FAILED(pTextRow->Apply(rowCount, tomRowApplyDefault))) {
							ATLASSERT(FALSE && "ITextRow::Apply failed");
						}
					} else {
						ATLASSERT(FALSE && "ITextRange2::GetRow failed");
					}
				}
			#endif
		}

		if(useFallback) {
			// try EM_INSERTTABLE
			hr = E_NOINTERFACE;
			if(hWndRTB) {
				if(columnWidth == -1) {
					// emulate auto-fit
					SCROLLINFO scrollInfo = {0};
					scrollInfo.cbSize = sizeof(SCROLLINFO);
					scrollInfo.fMask = SIF_ALL;
					GetScrollInfo(hWndRTB, SB_HORZ, &scrollInfo);
					int width = scrollInfo.nMax - scrollInfo.nMin;
					columnWidth = (twipsPerPixelX * width - rowIndent) / columnCount;
				}

				TABLEROWPARMS rowParams = {0};
				rowParams.cbRow = sizeof(TABLEROWPARMS);
				rowParams.cbCell = sizeof(TABLECELLPARMS);
				properties.pTextRange->GetStart(&rowParams.cpStartRow);
				rowParams.cCell = static_cast<BYTE>(columnCount);
				rowParams.cRow = static_cast<BYTE>(rowCount);
				switch(horizontalRowAlignment) {
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
				rowParams.dxCellMargin = horizontalCellMargin;
				rowParams.dxIndent = rowIndent;
				rowParams.dyHeight = (rowHeight <= 0 ? 0 : rowHeight);
				rowParams.fIdentCells = !VARIANTBOOL2BOOL(allowIndividualCellStyles);
				if(rowParams.fIdentCells) {
					TABLECELLPARMS cellParams = {0};
					cellParams.dxWidth = columnWidth;
					cellParams.nVertAlign = static_cast<WORD>(verticalCellAlignment);
					cellParams.dxBrdrLeft = borderSize;
					cellParams.dxBrdrRight = borderSize;
					cellParams.dyBrdrTop = borderSize;
					cellParams.dyBrdrBottom = borderSize;
					CComPtr<ITextFont> pTextFont = NULL;
					if(SUCCEEDED(properties.pTextRange->GetFont(&pTextFont))) {
						pTextFont->GetBackColor(reinterpret_cast<LONG*>(&cellParams.crBackPat));
					} else {
						cellParams.crBackPat = RGB(255, 255, 255);
					}
					hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_INSERTTABLE, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));
					if(hr == E_INVALIDARG) {
						// might be Windows XP
						/*rowParams.cbRow -= 2 * sizeof(long);
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_INSERTTABLE, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(&cellParams)));*/
					}
				} else {
					TABLECELLPARMS* pCellParams = reinterpret_cast<TABLECELLPARMS*>(HeapAlloc(GetProcessHeap(), 0, columnCount * sizeof(TABLECELLPARMS)));
					if(pCellParams) {
						ZeroMemory(pCellParams, columnCount * sizeof(TABLECELLPARMS));
						for(int col = 0; col < columnCount; ++col) {
							pCellParams[col].dxWidth = columnWidth;
							pCellParams[col].nVertAlign = static_cast<WORD>(verticalCellAlignment);
							pCellParams[col].dxBrdrLeft = borderSize;
							pCellParams[col].dxBrdrRight = borderSize;
							pCellParams[col].dyBrdrTop = borderSize;
							pCellParams[col].dyBrdrBottom = borderSize;
							CComPtr<ITextFont> pTextFont = NULL;
							if(SUCCEEDED(properties.pTextRange->GetFont(&pTextFont))) {
								pTextFont->GetBackColor(reinterpret_cast<LONG*>(&pCellParams[col].crBackPat));
							} else {
								pCellParams[col].crBackPat = RGB(255, 255, 255);
							}
						}
						hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_INSERTTABLE, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));
						if(hr == E_INVALIDARG) {
							// might be Windows XP
							/*rowParams.cbRow -= 2 * sizeof(long);
							hr = static_cast<HRESULT>(SendMessage(hWndRTB, EM_INSERTTABLE, reinterpret_cast<WPARAM>(&rowParams), reinterpret_cast<LPARAM>(pCellParams)));*/
						}
						HeapFree(GetProcessHeap(), 0, pCellParams);
					} else {
						hr = E_OUTOFMEMORY;
					}
				}
				/*if(SUCCEEDED(hr)) {
					useFallback = FALSE;
				}*/
			}
		}
	}

	/*if(useFallback && hWndRTB) {
		// NOTE: This fallback probably won't work for nested tables.
		LPSTR pTmp = new CHAR[70];
		if(columnWidth == -1) {
			// \trautofit control word does not seem to work, so emulate it
			SCROLLINFO scrollInfo = {0};
			scrollInfo.cbSize = sizeof(SCROLLINFO);
			scrollInfo.fMask = SIF_ALL;
			GetScrollInfo(hWndRTB, SB_HORZ, &scrollInfo);
			int width = scrollInfo.nMax - scrollInfo.nMin;
			columnWidth = (twipsPerPixelX * width - rowIndent) / columnCount;
		}

		CAtlStringA rawRichText = "{\\rtf1\\ansi\r\n\\uc1";
		for(LONG row = 1; row <= rowCount; ++row) {
			// row properties
			rawRichText += "\\trowd";
			if(pTmp && _itoa_s(horizontalCellMargin, pTmp, 70, 10) == 0) {
				rawRichText += "\\trgaph";
				rawRichText += pTmp;
			} else {
				rawRichText += "\\trgaph72";
			}
			switch(horizontalRowAlignment) {
				case halLeft:
					rawRichText += "\\trql";
					break;
				case halCenter:
					rawRichText += "\\trqc";
					break;
				case halRight:
					rawRichText += "\\trqr";
					break;
			}
			if(pTmp && _itoa_s(rowIndent, pTmp, 70, 10) == 0) {
				rawRichText += "\\trleft";
				rawRichText += pTmp;
			} else {
				rawRichText += "\\trleft0";
			}
			if(rowHeight > 0 && pTmp) {
				if(_itoa_s(rowHeight, pTmp, 70, 10) == 0) {
					rawRichText += "\\trrh";
					rawRichText += pTmp;
				}
			}
			if(pTmp && _itoa_s(horizontalCellMargin, pTmp, 70, 10) == 0) {
				rawRichText += "\\trpaddl";
				rawRichText += pTmp;
				rawRichText += "\\trpaddr";
				rawRichText += pTmp;
			} else {
				rawRichText += "\\trpaddl72\\trpaddr72";
			}
			rawRichText += "\\trpaddfl3\\trpaddfr3\r\n";

			for(LONG column = 1; column <= columnCount; ++column) {
				switch(verticalCellAlignment) {
					case valTop:
						rawRichText += "\\clvertalt";
						break;
					case valCenter:
						rawRichText += "\\clvertalc";
						break;
					case valBottom:
						rawRichText += "\\clvertalb";
						break;
				}
				// cell borders and cell width
				if(borderSize > 0 && pTmp && _itoa_s(borderSize, pTmp, 70, 10) == 0) {
					rawRichText += "\\clbrdrl\\brdrw";
					rawRichText += pTmp;
					rawRichText += "\\brdrs\\clbrdrt\\brdrw";
					rawRichText += pTmp;
					rawRichText += "\\brdrs\\clbrdrr\\brdrw";
					rawRichText += pTmp;
					rawRichText += "\\brdrs\\clbrdrb\\brdrw";
					rawRichText += pTmp;
					rawRichText += "\\brdrs";
				} else {
					rawRichText += "\\clbrdrl\\brdrw0\\brdrs\\clbrdrt\\brdrw0\\brdrs\\clbrdrr\\brdrw0\\brdrs\\clbrdrb\\brdrw0\\brdrs";
				}
				if(pTmp && _itoa_s(rowIndent + (column * columnWidth), pTmp, 70, 10) == 0) {
					rawRichText += " \\cellx";
					rawRichText += pTmp;
				}
			}
			rawRichText += "\\pard\\intbl";
			for(LONG column = 1; column <= columnCount; ++column) {
				// end of cell
				rawRichText += "\\cell";
			}
			// end of row
			rawRichText += "\\row";
			if(row == rowCount) {
				rawRichText += "\r\n";
			}
		}
		rawRichText += "}\r\n";
		if(pTmp) {
			SECUREFREE(pTmp);
		}
		SETTEXTEX options = {0};
		options.flags = ST_SELECTION | ST_KEEPUNDO;
		options.codepage = 0xFFFFFFFF;
		LPCSTR pBuffer = rawRichText;
		if(SendMessage(hWndRTB, EM_SETTEXTEX, reinterpret_cast<WPARAM>(&options), reinterpret_cast<LPARAM>(pBuffer))) {
			hr = S_OK;
		} else {
			hr = E_FAIL;
		}
		// keep the object alive until after EM_SETTEXTEX
		rawRichText = "";
	}*/

	if(SUCCEEDED(hr)) {
		CComPtr<ITextRange> pTextRange = NULL;
		properties.pTextRange->GetDuplicate(&pTextRange);
		pTextRange->SetStart(rangeStart);
		ClassFactory::InitTable(pTextRange, properties.pOwnerRTB, IID_IRichTable, reinterpret_cast<LPUNKNOWN*>(ppAddedTable));
	}
	return hr;
}

STDMETHODIMP TextRange::ScrollIntoView(ScrollIntoViewConstants flags/* = sivScrollRangeStartToTop*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->ScrollIntoView(flags);
	}
	return hr;
}

STDMETHODIMP TextRange::Select(void)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->Select();
	}
	return hr;
}

STDMETHODIMP TextRange::SetStartAndEnd(LONG anchorEnd, LONG activeEnd)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		hr = properties.pTextRange->SetRange(anchorEnd, activeEnd);
	}
	return hr;
}