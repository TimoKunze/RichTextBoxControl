// TextParagraph.cpp: A wrapper for the styling of a text range's paragraph.

#include "stdafx.h"
#include "TextParagraph.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TextParagraph::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTextParagraph, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TextParagraph::QueryITextParaInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextPara), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextPara2), queriedInterface)) {
		TextParagraph* pTextParagraph = reinterpret_cast<TextParagraph*>(pThis);
		return pTextParagraph->properties.pTextParagraph->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}


TextParagraph::Properties::~Properties()
{
	if(pTextParagraph) {
		pTextParagraph->Release();
		pTextParagraph = NULL;
	}
}


void TextParagraph::Attach(ITextPara* pTextParagraph)
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

void TextParagraph::Detach(void)
{
	if(properties.pTextParagraph) {
		properties.pTextParagraph->Release();
		properties.pTextParagraph = NULL;
	}
}


STDMETHODIMP TextParagraph::get_CanChange(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG canChange = tomFalse;
		hr = properties.pTextParagraph->CanChange(&canChange);
		*pValue = static_cast<BooleanPropertyValueConstants>(canChange);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_FirstLineIndent(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetFirstLineIndent(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_FirstLineIndent(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		FLOAT leftIndent = 0.0f;
		properties.pTextParagraph->GetLeftIndent(&leftIndent);
		FLOAT rightIndent = 0.0f;
		properties.pTextParagraph->GetRightIndent(&rightIndent);
		hr = properties.pTextParagraph->SetIndents(newValue, leftIndent, rightIndent);
		// TODO: Call ITextDocument::GetSelection()::Select() to ensure the cursor is displayed at the correct position
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_HAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG alignment = tomUndefined;
		hr = properties.pTextParagraph->GetAlignment(&alignment);
		*pValue = static_cast<HAlignmentConstants>(alignment);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_HAlignment(HAlignmentConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetAlignment(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_Hyphenation(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG hyphenation = tomFalse;
		hr = properties.pTextParagraph->GetHyphenation(&hyphenation);
		*pValue = static_cast<BooleanPropertyValueConstants>(hyphenation);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_Hyphenation(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetHyphenation(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_InsertPageBreakBeforeParagraph(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG pageBreakBefore = tomFalse;
		hr = properties.pTextParagraph->GetPageBreakBefore(&pageBreakBefore);
		*pValue = static_cast<BooleanPropertyValueConstants>(pageBreakBefore);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_InsertPageBreakBeforeParagraph(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetPageBreakBefore(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_KeepTogether(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG keepTogether = tomFalse;
		hr = properties.pTextParagraph->GetKeepTogether(&keepTogether);
		*pValue = static_cast<BooleanPropertyValueConstants>(keepTogether);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_KeepTogether(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetKeepTogether(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_KeepWithNext(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG keepWithNext = tomFalse;
		hr = properties.pTextParagraph->GetKeepWithNext(&keepWithNext);
		*pValue = static_cast<BooleanPropertyValueConstants>(keepWithNext);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_KeepWithNext(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetKeepWithNext(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_LeftIndent(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetLeftIndent(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_LeftIndent(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		FLOAT firstLineIndent = 0.0f;
		properties.pTextParagraph->GetFirstLineIndent(&firstLineIndent);
		FLOAT rightIndent = 0.0f;
		properties.pTextParagraph->GetRightIndent(&rightIndent);
		hr = properties.pTextParagraph->SetIndents(firstLineIndent, newValue, rightIndent);
		// TODO: Call ITextDocument::GetSelection()::Select() to ensure the cursor is displayed at the correct position
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_LineSpacing(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetLineSpacing(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_LineSpacing(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG lineSpacingRule = tomLineSpaceSingle;
		properties.pTextParagraph->GetLineSpacingRule(&lineSpacingRule);
		hr = properties.pTextParagraph->SetLineSpacing(lineSpacingRule, newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_LineSpacingRule(LineSpacingRuleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, LineSpacingRuleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG lineSpacingRule = tomLineSpaceSingle;
		hr = properties.pTextParagraph->GetLineSpacingRule(&lineSpacingRule);
		*pValue = static_cast<LineSpacingRuleConstants>(lineSpacingRule);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_LineSpacingRule(LineSpacingRuleConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		FLOAT lineSpacing = 0;
		properties.pTextParagraph->GetLineSpacing(&lineSpacing);
		hr = properties.pTextParagraph->SetLineSpacing(newValue, lineSpacing);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_ListLevelIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetListLevelIndex(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_ListLevelIndex(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetListLevelIndex(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_ListNumberingHAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG alignment = tomUndefined;
		hr = properties.pTextParagraph->GetListAlignment(&alignment);
		*pValue = static_cast<HAlignmentConstants>(alignment);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_ListNumberingHAlignment(HAlignmentConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetListAlignment(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_ListNumberingStart(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetListStart(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_ListNumberingStart(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetListStart(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_ListNumberingStyle(ListNumberingStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ListNumberingStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG listNumberingType = tomListNone;
		hr = properties.pTextParagraph->GetListType(&listNumberingType);
		*pValue = static_cast<ListNumberingStyleConstants>(listNumberingType);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_ListNumberingStyle(ListNumberingStyleConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetListType(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_ListNumberingTabStop(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetListTab(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_ListNumberingTabStop(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetListTab(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_Numbered(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG notNumbered = tomFalse;
		hr = properties.pTextParagraph->GetNoLineNumber(&notNumbered);
		switch(notNumbered) {
			case tomFalse:
				*pValue = bpvTrue;
				break;
			case tomTrue:
				*pValue = bpvFalse;
				break;
			default:
				*pValue = bpvUndefined;
				break;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_Numbered(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		switch(newValue) {
			case bpvToggle:
				hr = properties.pTextParagraph->SetNoLineNumber(tomToggle);
				break;
			case bpvFalse:
				hr = properties.pTextParagraph->SetNoLineNumber(tomTrue);
				break;
			case bpvTrue:
				hr = properties.pTextParagraph->SetNoLineNumber(tomFalse);
				break;
			default:
				hr = properties.pTextParagraph->SetNoLineNumber(tomUndefined);
				break;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_PreventWidowsAndOrphans(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		LONG preventWidowsAndOrphans = tomFalse;
		hr = properties.pTextParagraph->GetWidowControl(&preventWidowsAndOrphans);
		*pValue = static_cast<BooleanPropertyValueConstants>(preventWidowsAndOrphans);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_PreventWidowsAndOrphans(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetWidowControl(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_RightIndent(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetRightIndent(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_RightIndent(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		FLOAT firstLineIndent = 0.0f;
		properties.pTextParagraph->GetFirstLineIndent(&firstLineIndent);
		FLOAT leftIndent = 0.0f;
		properties.pTextParagraph->GetLeftIndent(&leftIndent);
		hr = properties.pTextParagraph->SetIndents(firstLineIndent, leftIndent, newValue);
		// TODO: Call ITextDocument::GetSelection()::Select() to ensure the cursor is displayed at the correct position
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_RightToLeft(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTextParagraph) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		CComQIPtr<ITextPara2> pTextParagraph2 = properties.pTextParagraph;
		if(pTextParagraph2) {
			long effects = 0;
			long mask = 0;
			hr = pTextParagraph2->GetEffects(&effects, &mask);
			if(SUCCEEDED(hr)) {
				if(!(mask & tomParaEffectRTL)) {
					*pValue = bpvUndefined;
				} else {
					if(effects & tomParaEffectRTL) {
						*pValue = bpvTrue;
					} else {
						*pValue = bpvFalse;
					}
				}
			}
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_RightToLeft(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && /*newValue != bpvUndefined && */newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(!properties.pTextParagraph) {
		return CO_E_RELEASED;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		CComQIPtr<ITextPara2> pTextParagraph2 = properties.pTextParagraph;
		if(pTextParagraph2) {
			LONG effects = 0;
			switch(newValue) {
				case bpvTrue:
					effects = tomParaEffectRTL;
					break;
				case bpvToggle:
				{
					LONG mask = tomParaEffectRTL;
					LONG curVal = 0;
					pTextParagraph2->GetEffects(&curVal, &mask);
					switch(curVal & tomParaEffectRTL) {
						case 0:
							effects = tomParaEffectRTL;
							break;
						/*default:
							effects = 0;
							break;*/
					}
					break;
				}
			}
			hr = pTextParagraph2->SetEffects(effects, tomParaEffectRTL);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_SpaceAfter(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetSpaceAfter(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_SpaceAfter(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetSpaceAfter(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_SpaceBefore(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetSpaceBefore(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_SpaceBefore(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetSpaceBefore(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_StyleID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->GetStyle(pValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::put_StyleID(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->SetStyle(newValue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::get_Tabs(IRichTextParagraphTabs** ppTabs)
{
	ATLASSERT_POINTER(ppTabs, IRichTextParagraphTabs*);
	if(!ppTabs) {
		return E_POINTER;
	}

	ClassFactory::InitTextParagraphTabs(properties.pTextParagraph, IID_IRichTextParagraphTabs, reinterpret_cast<LPUNKNOWN*>(ppTabs));
	return S_OK;
}


STDMETHODIMP TextParagraph::Clone(IRichTextParagraph** ppClone)
{
	ATLASSERT_POINTER(ppClone, IRichTextParagraph*);
	if(!ppClone) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	*ppClone = NULL;
	if(properties.pTextParagraph) {
		CComPtr<ITextPara> pParagraph = NULL;
		hr = properties.pTextParagraph->GetDuplicate(&pParagraph);
		ClassFactory::InitTextParagraph(pParagraph, IID_IRichTextParagraph, reinterpret_cast<LPUNKNOWN*>(ppClone));
	}
	return hr;
}

STDMETHODIMP TextParagraph::CopySettings(IRichTextParagraph* pSourceObject)
{
	if(!pSourceObject) {
		// invalid parameter - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		CComQIPtr<ITextPara> pParagraph = pSourceObject;
		hr = properties.pTextParagraph->SetDuplicate(pParagraph);
	}
	return hr;
}

STDMETHODIMP TextParagraph::Equals(IRichTextParagraph* pCompareAgainst, VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pCompareAgainst, IRichTextParagraph);
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pCompareAgainst || !pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		CComQIPtr<ITextPara> pParagraph = pCompareAgainst;
		LONG equal = tomFalse;
		hr = properties.pTextParagraph->IsEqual(pParagraph, &equal);
		*pValue = BOOL2VARIANTBOOL(equal == tomTrue);
	}
	return hr;
}

STDMETHODIMP TextParagraph::Reset(ResetRulesConstants rules)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextParagraph) {
		hr = properties.pTextParagraph->Reset(rules);
	}
	return hr;
}