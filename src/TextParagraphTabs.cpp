// TextParagraphTabs.cpp: Manages a collection of TextParagraphTab objects

#include "stdafx.h"
#include "TextParagraphTabs.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP TextParagraphTabs::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTextParagraphTabs, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP TextParagraphTabs::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<TextParagraphTabs>* pTextParagraphTabsObj = NULL;
	CComObject<TextParagraphTabs>::CreateInstance(&pTextParagraphTabsObj);
	pTextParagraphTabsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTextParagraphTabsObj->properties);

	pTextParagraphTabsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTextParagraphTabsObj->Release();
	return S_OK;
}

STDMETHODIMP TextParagraphTabs::Next(ULONG numberOfMaxTabs, VARIANT* pTabs, ULONG* pNumberOfTabsReturned)
{
	ATLASSERT_POINTER(pTabs, VARIANT);
	if(!pTabs) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextParagraph);

	// check each tab in the paragraph
	long numberOfTabs = 0;
	HRESULT hr = properties.pTextParagraph->GetTabCount(&numberOfTabs);
	if(FAILED(hr)) {
		return hr;
	}
	ULONG i = 0;
	for(i = 0; i < numberOfMaxTabs; ++i) {
		VariantInit(&pTabs[i]);

		do {
			if(properties.lastEnumeratedTab >= 0) {
				if(properties.lastEnumeratedTab < numberOfTabs) {
					properties.lastEnumeratedTab = GetNextTabToProcess(properties.lastEnumeratedTab, numberOfTabs);
				}
			} else {
				properties.lastEnumeratedTab = GetFirstTabToProcess(numberOfTabs);
			}
			if(properties.lastEnumeratedTab >= numberOfTabs) {
				properties.lastEnumeratedTab = -1;
			}
		} while((properties.lastEnumeratedTab != -1) && (!IsPartOfCollection(properties.lastEnumeratedTab)));

		if(properties.lastEnumeratedTab != -1) {
			float tabPosition = 0.0f;
			long tabAlign = tomAlignLeft;
			long leaderChar = tomSpaces;
			hr = properties.pTextParagraph->GetTab(properties.lastEnumeratedTab, &tabPosition, &tabAlign, &leaderChar);
			if(SUCCEEDED(hr) && hr != S_FALSE) {
				ClassFactory::InitTextParagraphTab(tabPosition, properties.pTextParagraph, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pTabs[i].pdispVal));
				pTabs[i].vt = VT_DISPATCH;
			} else {
				i--;
				break;
			}
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfTabsReturned) {
		*pNumberOfTabsReturned = i;
	}

	return (i == numberOfMaxTabs ? S_OK : S_FALSE);
}

STDMETHODIMP TextParagraphTabs::Reset(void)
{
	properties.lastEnumeratedTab = -1;
	return S_OK;
}

STDMETHODIMP TextParagraphTabs::Skip(ULONG numberOfTabsToSkip)
{
	VARIANT dummy;
	ULONG numTabsReturned = 0;
	// we could skip all tabs at once, but it's easier to skip them one by one
	for(ULONG i = 1; i <= numberOfTabsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numTabsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numTabsReturned == 0) {
			// there're no more tabs to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////

int TextParagraphTabs::GetFirstTabToProcess(int numberOfTabs)
{
	if(numberOfTabs == 0) {
		return -1;
	}
	return 0;
}

int TextParagraphTabs::GetNextTabToProcess(int previousTab, int numberOfTabs)
{
	if(previousTab < numberOfTabs - 1) {
		return previousTab + 1;
	} else {
		return -1;
	}
}

BOOL TextParagraphTabs::IsPartOfCollection(int tabIndex)
{
	ATLASSERT(properties.pTextParagraph);

	return IsValidTextParagraphTabIndex(tabIndex, properties.pTextParagraph);
}


TextParagraphTabs::Properties::~Properties()
{
	if(pTextParagraph) {
		pTextParagraph->Release();
		pTextParagraph = NULL;
	}
}

void TextParagraphTabs::Properties::CopyTo(TextParagraphTabs::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pTextParagraph = this->pTextParagraph;
		if(pTarget->pTextParagraph) {
			pTarget->pTextParagraph->AddRef();
		}
		pTarget->lastEnumeratedTab = this->lastEnumeratedTab;
	}
}


void TextParagraphTabs::Attach(ITextPara* pTextParagraph)
{
	if(properties.pTextParagraph) {
		properties.pTextParagraph->Release();
		properties.pTextParagraph = NULL;
	}
	if(pTextParagraph) {
		//pTextParagraph->QueryInterface(IID_PPV_ARGS(&properties.pTextParagraph));
		properties.pTextParagraph = pTextParagraph;
		properties.pTextParagraph->AddRef();
	}
}

void TextParagraphTabs::Detach(void)
{
	if(properties.pTextParagraph) {
		properties.pTextParagraph->Release();
		properties.pTextParagraph = NULL;
	}
}


STDMETHODIMP TextParagraphTabs::get_Item(VARIANT tabIdentifier, TabIdentifierTypeConstants tabIdentifierType/* = titPosition*/, IRichTextParagraphTab** ppTab/* = NULL*/)
{
	ATLASSERT_POINTER(ppTab, IRichTextParagraphTab*);
	if(!ppTab) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextParagraph);

	// retrieve the tab's position
	float tabPosition = 0.0f;
	HRESULT hr = E_FAIL;
	VARIANT v;
	VariantInit(&v);
	switch(tabIdentifierType) {
		case titIndex:
			if(SUCCEEDED(VariantChangeType(&v, &tabIdentifier, 0, VT_I4))) {
				long alignment = tomUndefined;
				long leaderChar = tomSpaces;
				hr = properties.pTextParagraph->GetTab(v.intVal, &tabPosition, &alignment, &leaderChar);
			}
			break;
		case titPosition:
			if(SUCCEEDED(VariantChangeType(&v, &tabIdentifier, 0, VT_R4))) {
				tabPosition = v.fltVal;
				long alignment = tomUndefined;
				long leaderChar = tomSpaces;
				hr = properties.pTextParagraph->GetTab(tomTabHere, &tabPosition, &alignment, &leaderChar);
			}
			break;
	}
	VariantClear(&v);

	if(SUCCEEDED(hr) && hr != S_FALSE) {
		ClassFactory::InitTextParagraphTab(tabPosition, properties.pTextParagraph, IID_IRichTextParagraphTab, reinterpret_cast<LPUNKNOWN*>(ppTab));
		if(*ppTab) {
			return S_OK;
		}
	}

	// tab not found
	if(tabIdentifierType == titIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP TextParagraphTabs::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP TextParagraphTabs::Add(FLOAT tabPosition, TabAlignmentConstants tabAlignment/* = taLeft*/, TabLeaderCharacterConstants leaderCharacter/* = tlcSpaces*/, IRichTextParagraphTab** ppAddedTab/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedTab, IRichTextParagraphTab*);
	if(!ppAddedTab) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextParagraph);
	*ppAddedTab = NULL;
	HRESULT hr = E_FAIL;

	if(SUCCEEDED(properties.pTextParagraph->AddTab(tabPosition, tabAlignment, leaderCharacter))) {
		long tabAlign = tomAlignLeft;
		long leaderChar = tomSpaces;
		hr = properties.pTextParagraph->GetTab(tomTabHere, &tabPosition, &tabAlign, &leaderChar);
		if(SUCCEEDED(hr) && hr != S_FALSE) {
			if(tabAlign == tabAlignment && leaderChar == leaderCharacter) {
				ClassFactory::InitTextParagraphTab(tabPosition, properties.pTextParagraph, IID_IRichTextParagraphTab, reinterpret_cast<LPUNKNOWN*>(ppAddedTab), FALSE);
				hr = S_OK;
			}
		} else {
			hr = E_FAIL;
		}
	}

	return hr;
}

STDMETHODIMP TextParagraphTabs::Contains(VARIANT tabIdentifier, TabIdentifierTypeConstants tabIdentifierType/* = titPosition*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextParagraph);
	*pValue = VARIANT_FALSE;

	// retrieve the tab's position
	VARIANT v;
	VariantInit(&v);
	switch(tabIdentifierType) {
		case titIndex:
			if(SUCCEEDED(VariantChangeType(&v, &tabIdentifier, 0, VT_I4))) {
				float tabPosition = 0.0f;
				long alignment = tomUndefined;
				long leaderChar = tomSpaces;
				HRESULT hr = properties.pTextParagraph->GetTab(v.intVal, &tabPosition, &alignment, &leaderChar);
				*pValue = BOOL2VARIANTBOOL(SUCCEEDED(hr) && hr != S_FALSE);
			}
			break;
		case titPosition:
			if(SUCCEEDED(VariantChangeType(&v, &tabIdentifier, 0, VT_R4))) {
				float tabPosition = v.fltVal;
				long alignment = tomUndefined;
				long leaderChar = tomSpaces;
				HRESULT hr = properties.pTextParagraph->GetTab(tomTabHere, &tabPosition, &alignment, &leaderChar);
				*pValue = BOOL2VARIANTBOOL(SUCCEEDED(hr) && hr != S_FALSE);
			}
			break;
	}
	VariantClear(&v);
	return S_OK;
}

STDMETHODIMP TextParagraphTabs::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextParagraph);
	return properties.pTextParagraph->GetTabCount(pValue);
}

STDMETHODIMP TextParagraphTabs::Remove(VARIANT tabIdentifier, TabIdentifierTypeConstants tabIdentifierType/* = titPosition*/)
{
	ATLASSERT(properties.pTextParagraph);

	// retrieve the tab's position
	float tabPosition = 0.0f;
	HRESULT hr = E_FAIL;
	VARIANT v;
	VariantInit(&v);
	switch(tabIdentifierType) {
		case titIndex:
			if(SUCCEEDED(VariantChangeType(&v, &tabIdentifier, 0, VT_I4))) {
				long alignment = tomUndefined;
				long leaderChar = tomSpaces;
				hr = properties.pTextParagraph->GetTab(v.intVal, &tabPosition, &alignment, &leaderChar);
			}
			break;
		case titPosition:
			if(SUCCEEDED(VariantChangeType(&v, &tabIdentifier, 0, VT_R4))) {
				tabPosition = v.fltVal;
				long alignment = tomUndefined;
				long leaderChar = tomSpaces;
				hr = properties.pTextParagraph->GetTab(tomTabHere, &tabPosition, &alignment, &leaderChar);
			}
			break;
	}
	VariantClear(&v);

	if(SUCCEEDED(hr) && hr != S_FALSE) {
		return properties.pTextParagraph->DeleteTab(tabPosition);
	} else {
		// tab not found
		if(tabIdentifierType == titIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}
}

STDMETHODIMP TextParagraphTabs::RemoveAll(void)
{
	ATLASSERT(properties.pTextParagraph);
	return properties.pTextParagraph->ClearAllTabs();
}