// TextParagraphTab.cpp: A wrapper for a specific tab of a text range's paragraph.

#include "stdafx.h"
#include "TextParagraphTab.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP TextParagraphTab::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTextParagraphTab, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


TextParagraphTab::Properties::~Properties()
{
	if(pTextParagraph) {
		pTextParagraph->Release();
		pTextParagraph = NULL;
	}
}


void TextParagraphTab::Attach(ITextPara* pTextParagraph, float tabPosition)
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
	properties.tabPosition = tabPosition;
}

void TextParagraphTab::Detach(void)
{
	if(properties.pTextParagraph) {
		properties.pTextParagraph->Release();
		properties.pTextParagraph = NULL;
	}
	properties.tabPosition = 0.0f;
}


STDMETHODIMP TextParagraphTab::get_Alignment(TabAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TabAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		float position = properties.tabPosition;
		long alignment = tomUndefined;
		long leaderChar = tomSpaces;
		hr = properties.pTextParagraph->GetTab(tomTabHere, &position, &alignment, &leaderChar);
		if(SUCCEEDED(hr) && hr != S_FALSE) {
			*pValue = static_cast<TabAlignmentConstants>(alignment);
		} else if(SUCCEEDED(hr)) {
			hr = E_FAIL;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraphTab::put_Alignment(TabAlignmentConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		float position = properties.tabPosition;
		long alignment = tomUndefined;
		long leaderChar = tomSpaces;
		hr = properties.pTextParagraph->GetTab(tomTabHere, &position, &alignment, &leaderChar);
		if(SUCCEEDED(hr) && hr != S_FALSE) {
			hr = properties.pTextParagraph->DeleteTab(position);
			hr = properties.pTextParagraph->AddTab(position, newValue, leaderChar);
		} else if(SUCCEEDED(hr)) {
			hr = E_FAIL;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraphTab::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ATLASSERT(properties.pTextParagraph);
	long tabCount = 0;
	if(SUCCEEDED(properties.pTextParagraph->GetTabCount(&tabCount))) {
		for(long tabIndex = 0; tabIndex < tabCount; ++tabIndex) {
			float position = 0.0f;
			long alignment = tomUndefined;
			long leaderChar = tomSpaces;
			HRESULT hr = properties.pTextParagraph->GetTab(tabIndex, &position, &alignment, &leaderChar);
			if(SUCCEEDED(hr) && hr != S_FALSE) {
				if(position == properties.tabPosition) {
					*pValue = tabIndex;
					return S_OK;
				}
			}
		}
	}
	return E_FAIL;
}

STDMETHODIMP TextParagraphTab::get_LeaderCharacter(TabLeaderCharacterConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TabLeaderCharacterConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		float position = properties.tabPosition;
		long alignment = tomUndefined;
		long leaderChar = tomSpaces;
		hr = properties.pTextParagraph->GetTab(tomTabHere, &position, &alignment, &leaderChar);
		if(SUCCEEDED(hr) && hr != S_FALSE) {
			*pValue = static_cast<TabLeaderCharacterConstants>(leaderChar);
		} else if(SUCCEEDED(hr)) {
			hr = E_FAIL;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraphTab::put_LeaderCharacter(TabLeaderCharacterConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		float position = properties.tabPosition;
		long alignment = tomUndefined;
		long leaderChar = tomSpaces;
		hr = properties.pTextParagraph->GetTab(tomTabHere, &position, &alignment, &leaderChar);
		if(SUCCEEDED(hr) && hr != S_FALSE) {
			hr = properties.pTextParagraph->DeleteTab(position);
			hr = properties.pTextParagraph->AddTab(position, alignment, newValue);
		} else if(SUCCEEDED(hr)) {
			hr = E_FAIL;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraphTab::get_NextTab(IRichTextParagraphTab** ppTab)
{
	ATLASSERT_POINTER(ppTab, IRichTextParagraphTab*);
	if(!ppTab) {
		return E_POINTER;
	}

	if(properties.pTextParagraph) {
		float position = properties.tabPosition;
		long alignment = tomUndefined;
		long leaderChar = tomSpaces;
		HRESULT hr = properties.pTextParagraph->GetTab(tomTabNext, &position, &alignment, &leaderChar);
		if(SUCCEEDED(hr) && hr != S_FALSE) {
			ClassFactory::InitTextParagraphTab(position, properties.pTextParagraph, IID_IRichTextParagraphTab, reinterpret_cast<LPUNKNOWN*>(ppTab), FALSE);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP TextParagraphTab::get_Position(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.tabPosition;
	return S_OK;
}

STDMETHODIMP TextParagraphTab::put_Position(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		float position = properties.tabPosition;
		long alignment = tomUndefined;
		long leaderChar = tomSpaces;
		hr = properties.pTextParagraph->GetTab(tomTabHere, &position, &alignment, &leaderChar);
		if(SUCCEEDED(hr) && hr != S_FALSE) {
			hr = properties.pTextParagraph->DeleteTab(position);
			hr = properties.pTextParagraph->AddTab(newValue, alignment, leaderChar);
			properties.tabPosition = newValue;
		} else if(SUCCEEDED(hr)) {
			hr = E_FAIL;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraphTab::get_PreviousTab(IRichTextParagraphTab** ppTab)
{
	ATLASSERT_POINTER(ppTab, IRichTextParagraphTab*);
	if(!ppTab) {
		return E_POINTER;
	}

	if(properties.pTextParagraph) {
		float position = properties.tabPosition;
		long alignment = tomUndefined;
		long leaderChar = tomSpaces;
		HRESULT hr = properties.pTextParagraph->GetTab(tomTabBack, &position, &alignment, &leaderChar);
		if(SUCCEEDED(hr) && hr != S_FALSE) {
			ClassFactory::InitTextParagraphTab(position, properties.pTextParagraph, IID_IRichTextParagraphTab, reinterpret_cast<LPUNKNOWN*>(ppTab), FALSE);
		}
		return S_OK;
	}
	return E_FAIL;
}