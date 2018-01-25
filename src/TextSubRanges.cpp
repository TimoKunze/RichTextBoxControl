// TextSubRanges.cpp: Manages a collection of TextRange objects

#include "stdafx.h"
#include "TextSubRanges.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TextSubRanges::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTextSubRanges, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TextSubRanges::QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRange), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextRange2), queriedInterface)) {
		TextSubRanges* pTextSubRanges = reinterpret_cast<TextSubRanges*>(pThis);
		return pTextSubRanges->properties.pTextRange->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP TextSubRanges::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<TextSubRanges>* pTextSubRangesObj = NULL;
	CComObject<TextSubRanges>::CreateInstance(&pTextSubRangesObj);
	pTextSubRangesObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTextSubRangesObj->properties);

	pTextSubRangesObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTextSubRangesObj->Release();
	return S_OK;
}

STDMETHODIMP TextSubRanges::Next(ULONG numberOfMaxRanges, VARIANT* pTextRanges, ULONG* pNumberOfRangesReturned)
{
	ATLASSERT_POINTER(pTextRanges, VARIANT);
	if(!pTextRanges) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextRange);
	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}

	// check each sub-range in the range
	LONG numberOfSubRanges = 0;
	HRESULT hr = pTextRange2->GetCount(&numberOfSubRanges);
	if(FAILED(hr)) {
		return hr;
	}
	ULONG i = 0;
	for(i = 0; i < numberOfMaxRanges; ++i) {
		VariantInit(&pTextRanges[i]);

		do {
			if(properties.lastEnumeratedSubRange <= 0) {
				if(-properties.lastEnumeratedSubRange <= numberOfSubRanges) {
					properties.lastEnumeratedSubRange--;//= GetNextSubRangeToProcess(properties.lastEnumeratedSubRange, numberOfSubRanges);
				}
			} else {
				properties.lastEnumeratedSubRange = 0;//GetFirstSubRangeToProcess(numberOfSubRanges);
			}
			if(-properties.lastEnumeratedSubRange > numberOfSubRanges) {
				properties.lastEnumeratedSubRange = 1;
			}
		} while((properties.lastEnumeratedSubRange != 1) && !TRUE/*(!IsPartOfCollection(properties.lastEnumeratedSubRange))*/);

		if(properties.lastEnumeratedSubRange != 1) {
			LONG start = 0;
			LONG end = 0;
			hr = pTextRange2->GetSubrange(properties.lastEnumeratedSubRange, &start, &end);
			if(SUCCEEDED(hr)) {
				CComPtr<ITextRange> pSubRange = NULL;
				if(SUCCEEDED(properties.pTextRange->GetDuplicate(&pSubRange))) {
					pSubRange->SetStart(start);
					pSubRange->SetEnd(end);
				}
				if(pSubRange) {
					ClassFactory::InitTextRange(pSubRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(&pTextRanges[i].pdispVal));
					pTextRanges[i].vt = VT_DISPATCH;
				} else {
					i--;
					break;
				}
			} else {
				i--;
				break;
			}
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfRangesReturned) {
		*pNumberOfRangesReturned = i;
	}

	return (i == numberOfMaxRanges ? S_OK : S_FALSE);
}

STDMETHODIMP TextSubRanges::Reset(void)
{
	properties.lastEnumeratedSubRange = 1;
	return S_OK;
}

STDMETHODIMP TextSubRanges::Skip(ULONG numberOfRangesToSkip)
{
	VARIANT dummy;
	ULONG numSubRangesReturned = 0;
	// we could skip all sub-ranges at once, but it's easier to skip them one by one
	for(ULONG i = 1; i <= numberOfRangesToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numSubRangesReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numSubRangesReturned == 0) {
			// there're no more sub-ranges to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


TextSubRanges::Properties::~Properties()
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

void TextSubRanges::Properties::CopyTo(TextSubRanges::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pTextRange = this->pTextRange;
		if(pTarget->pTextRange) {
			pTarget->pTextRange->AddRef();
		}
		pTarget->lastEnumeratedSubRange = this->lastEnumeratedSubRange;
	}
}


void TextSubRanges::Attach(ITextRange* pTextRange)
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

void TextSubRanges::Detach(void)
{
	if(properties.pTextRange) {
		properties.pTextRange->Release();
		properties.pTextRange = NULL;
	}
}

void TextSubRanges::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP TextSubRanges::get_ActiveSubRange(IRichTextRange** ppActiveSubRange)
{
	ATLASSERT_POINTER(ppActiveSubRange, IRichTextRange*);
	if(!ppActiveSubRange) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextRange);

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}

	*ppActiveSubRange = NULL;

	LONG start = 0;
	LONG end = 0;
	HRESULT hr = pTextRange2->GetSubrange(0, &start, &end);
	if(SUCCEEDED(hr)) {
		CComPtr<ITextRange> pSubRange = NULL;
		if(SUCCEEDED(properties.pTextRange->GetDuplicate(&pSubRange))) {
			pSubRange->SetStart(start);
			pSubRange->SetEnd(end);
			ClassFactory::InitTextRange(pSubRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppActiveSubRange));
			if(*ppActiveSubRange) {
				hr = S_OK;
			}
		}
	}

	return hr;
}

/*STDMETHODIMP TextSubRanges::putref_ActiveSubRange(IRichTextRange* pNewActiveSubRange)
{
	ATLASSERT(properties.pTextRange);

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}

	if(pNewActiveSubRange) {
		LONG start = 0;
		LONG end = 0;
		pNewActiveSubRange->get_RangeStart(&start);
		pNewActiveSubRange->get_RangeEnd(&end);
		// TODO: Shouldn't we AddRef' pNewActiveSubRange?

		// currently fails with E_NOTIMPL, even on Windows 10
		return pTextRange2->SetActiveSubrange(end, start);
	}
	return E_INVALIDARG;
}*/

STDMETHODIMP TextSubRanges::get_Item(VARIANT subRangeIdentifier, SubRangeIdentifierTypeConstants subRangeIdentifierType/* = sritIndex*/, IRichTextRange** ppSubRange/* = NULL*/)
{
	ATLASSERT_POINTER(ppSubRange, IRichTextRange*);
	if(!ppSubRange) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextRange);

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}

	// retrieve the sub-range's start and end
	LONG start = 0;
	LONG end = 0;
	HRESULT hr = E_FAIL;
	VARIANT v;
	VariantInit(&v);
	switch(subRangeIdentifierType) {
		case sritIndex:
			if(SUCCEEDED(VariantChangeType(&v, &subRangeIdentifier, 0, VT_I4))) {
				hr = pTextRange2->GetSubrange(v.intVal, &start, &end);
			}
			break;
	}
	VariantClear(&v);

	CComPtr<ITextRange> pSubRange = NULL;
	if(SUCCEEDED(properties.pTextRange->GetDuplicate(&pSubRange))) {
		pSubRange->SetStart(start);
		pSubRange->SetEnd(end);
	}

	if(SUCCEEDED(hr) && pSubRange) {
		ClassFactory::InitTextRange(pSubRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppSubRange));
		if(*ppSubRange) {
			return S_OK;
		}
	}
	// sub-range not found
	if(subRangeIdentifierType == sritIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP TextSubRanges::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP TextSubRanges::Add(LONG activeSubRangeEnd, LONG anchorSubRangeEnd, VARIANT_BOOL activateSubRange/* = VARIANT_TRUE*/, IRichTextRange** ppAddedSubRange/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedSubRange, IRichTextRange*);
	if(!ppAddedSubRange) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextRange);
	*ppAddedSubRange = NULL;

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}
	HRESULT hr = E_FAIL;

	if(SUCCEEDED(pTextRange2->AddSubrange(activeSubRangeEnd, anchorSubRangeEnd, (activateSubRange == VARIANT_FALSE ? tomFalse : tomTrue)))) {
		CComPtr<ITextRange> pSubRange = NULL;
		if(SUCCEEDED(properties.pTextRange->GetDuplicate(&pSubRange))) {
			pSubRange->SetStart(activeSubRangeEnd);
			pSubRange->SetEnd(anchorSubRangeEnd);
		}
		if(pSubRange) {
			ClassFactory::InitTextRange(pSubRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppAddedSubRange));
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP TextSubRanges::Contains(VARIANT subRangeIdentifier, SubRangeIdentifierTypeConstants subRangeIdentifierType/* = sritIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextRange);
	*pValue = VARIANT_FALSE;

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}

	// retrieve the sub-range's start and end
	LONG start = 0;
	LONG end = 0;
	HRESULT hr = E_FAIL;
	VARIANT v;
	VariantInit(&v);
	switch(subRangeIdentifierType) {
		case sritIndex:
			if(SUCCEEDED(VariantChangeType(&v, &subRangeIdentifier, 0, VT_I4))) {
				hr = pTextRange2->GetSubrange(v.intVal, &start, &end);
			}
			break;
	}
	VariantClear(&v);

	if(SUCCEEDED(hr)) {
		*pValue = VARIANT_TRUE;
	} else {
		// sub-range not found
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP TextSubRanges::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextRange);

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}
	return pTextRange2->GetCount(pValue);
}

STDMETHODIMP TextSubRanges::Remove(VARIANT subRangeIdentifier, SubRangeIdentifierTypeConstants subRangeIdentifierType/* = sritIndex*/)
{
	ATLASSERT(properties.pTextRange);

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}

	// retrieve the sub-range's start and end
	LONG start = 0;
	LONG end = 0;
	HRESULT hr = E_FAIL;
	VARIANT v;
	VariantInit(&v);
	switch(subRangeIdentifierType) {
		case sritIndex:
			if(SUCCEEDED(VariantChangeType(&v, &subRangeIdentifier, 0, VT_I4))) {
				hr = pTextRange2->GetSubrange(v.intVal, &start, &end);
			}
			break;
	}
	VariantClear(&v);

	if(SUCCEEDED(hr)) {
		return pTextRange2->DeleteSubrange(start, end);
	} else {
		// sub-range not found
		if(subRangeIdentifierType == sritIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}
}

STDMETHODIMP TextSubRanges::RemoveAll(void)
{
	ATLASSERT(properties.pTextRange);

	CComQIPtr<ITextRange2> pTextRange2 = properties.pTextRange;
	if(!pTextRange2) {
		return E_NOINTERFACE;
	}

	LONG count = 0;
	HRESULT hr = pTextRange2->GetCount(&count);
	while(SUCCEEDED(hr) && count > 0) {
		LONG start = 0;
		LONG end = 0;
		hr = pTextRange2->GetSubrange(count - 1, &start, &end);
		if(SUCCEEDED(hr)) {
			hr = pTextRange2->DeleteSubrange(start, end);
		}

		if(SUCCEEDED(hr)) {
			LONG newCount = 0;
			pTextRange2->GetCount(&newCount);
			if(newCount == count) {
				hr = S_OK;
				break;
			}
			count = newCount;
		}
	}
	return hr;
}