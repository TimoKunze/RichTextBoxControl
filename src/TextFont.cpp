// TextFont.cpp: A wrapper for the styling of a text range.

#include "stdafx.h"
#include "TextFont.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP TextFont::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTextFont, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK TextFont::QueryITextFontInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextFont), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextFont2), queriedInterface)) {
		TextFont* pTextFont = reinterpret_cast<TextFont*>(pThis);
		return pTextFont->properties.pTextFont->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}


TextFont::Properties::~Properties()
{
	if(pTextFont) {
		pTextFont->Release();
		pTextFont = NULL;
	}
}


void TextFont::Attach(ITextFont* pTextFont)
{
	if(properties.pTextFont) {
		properties.pTextFont->Release();
		properties.pTextFont = NULL;
	}
	if(pTextFont) {
		//pTextFont->QueryInterface(IID_PPV_ARGS(&properties.pTextFont));
		properties.pTextFont = pTextFont;
		properties.pTextFont->AddRef();
	}
}

void TextFont::Detach(void)
{
	if(properties.pTextFont) {
		properties.pTextFont->Release();
		properties.pTextFont = NULL;
	}
}


/*STDMETHODIMP TextFont::get_AlignEquationsAtOperator(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG alignment = 0;
			hr = pTextFont2->GetProperty(tomFontPropAlign, &alignment);
			*pValue = static_cast<SHORT>(alignment);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::put_AlignEquationsAtOperator(SHORT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			hr = pTextFont2->SetProperty(tomFontPropAlign, newValue);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}*/

STDMETHODIMP TextFont::get_AllCaps(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG allCaps = tomFalse;
		hr = properties.pTextFont->GetAllCaps(&allCaps);
		*pValue = static_cast<BooleanPropertyValueConstants>(allCaps);
	}
	return hr;
}

STDMETHODIMP TextFont::put_AllCaps(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetAllCaps(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_AnimationType(AnimationTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AnimationTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG animationType = tomNone;
		hr = properties.pTextFont->GetAnimation(&animationType);
		*pValue = static_cast<AnimationTypeConstants>(animationType);
	}
	return hr;
}

STDMETHODIMP TextFont::put_AnimationType(AnimationTypeConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetAnimation(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_AutoLigatures(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG autoLigatures = tomFalse;
			hr = pTextFont2->GetAutoLigatures(&autoLigatures);
			*pValue = static_cast<BooleanPropertyValueConstants>(autoLigatures);
		} else {
			*pValue = bpvFalse;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::put_AutoLigatures(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			hr = pTextFont2->SetAutoLigatures(newValue);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG color = tomAutoColor;
		hr = properties.pTextFont->GetBackColor(&color);
		if(color == tomAutoColor) {
			*pValue = static_cast<OLE_COLOR>(-1);
		} else {
			*pValue = color;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::put_BackColor(OLE_COLOR newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		if(newValue == -1) {
			hr = properties.pTextFont->SetBackColor(tomAutoColor);
		} else {
			hr = properties.pTextFont->SetBackColor(OLECOLOR2COLORREF(newValue));
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_Bold(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG bold = tomFalse;
		hr = properties.pTextFont->GetBold(&bold);
		*pValue = static_cast<BooleanPropertyValueConstants>(bold);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Bold(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetBold(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_CanChange(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG canChange = tomFalse;
		hr = properties.pTextFont->CanChange(&canChange);
		*pValue = static_cast<BooleanPropertyValueConstants>(canChange);
	}
	return hr;
}

STDMETHODIMP TextFont::get_DisplayAsOrdinaryTextWithMathZone(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			long effects = 0;
			long mask = 0;
			hr = pTextFont2->GetEffects(&effects, &mask);
			if(SUCCEEDED(hr)) {
				if(!(mask & tomMathZoneOrdinary)) {
					*pValue = bpvUndefined;
				} else {
					if(effects & tomMathZoneOrdinary) {
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

STDMETHODIMP TextFont::put_DisplayAsOrdinaryTextWithMathZone(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG effects = 0;
			switch(newValue) {
				case bpvTrue:
					effects = tomMathZoneOrdinary;
					break;
				case bpvToggle:
				{
					LONG mask = tomMathZoneOrdinary;
					LONG curVal = 0;
					pTextFont2->GetEffects(&curVal, &mask);
					switch(curVal & tomMathZoneOrdinary) {
						case 0:
							effects = tomMathZoneOrdinary;
							break;
						/*default:
							effects = 0;
							break;*/
					}
					break;
				}
			}
			hr = pTextFont2->SetEffects(effects, tomMathZoneOrdinary);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_Emboss(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG emboss = tomFalse;
		hr = properties.pTextFont->GetEmboss(&emboss);
		*pValue = static_cast<BooleanPropertyValueConstants>(emboss);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Emboss(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetEmboss(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Engrave(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG engrave = tomFalse;
		hr = properties.pTextFont->GetEngrave(&engrave);
		*pValue = static_cast<BooleanPropertyValueConstants>(engrave);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Engrave(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetEngrave(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG color = tomAutoColor;
		hr = properties.pTextFont->GetForeColor(&color);
		if(color == tomAutoColor) {
			*pValue = static_cast<OLE_COLOR>(-1);
		} else {
			*pValue = color;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::put_ForeColor(OLE_COLOR newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		if(newValue == -1) {
			hr = properties.pTextFont->SetForeColor(tomAutoColor);
		} else {
			hr = properties.pTextFont->SetForeColor(OLECOLOR2COLORREF(newValue));
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_Hidden(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG hidden = tomFalse;
		hr = properties.pTextFont->GetHidden(&hidden);
		*pValue = static_cast<BooleanPropertyValueConstants>(hidden);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Hidden(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetHidden(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_HorizontalScaling(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG scaling = 100;
			hr = pTextFont2->GetScaling(&scaling);
			*pValue = scaling;
		} else {
			*pValue = 100;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::put_HorizontalScaling(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			hr = pTextFont2->SetScaling(newValue);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_IsAutomaticLink(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG effects = tomLink | tomLinkProtected;
			LONG mask = tomLink | tomLinkProtected;
			hr = pTextFont2->GetEffects(&effects, &mask);
			if(SUCCEEDED(hr)) {
				if(!(mask & tomLinkProtected)) {
					*pValue = bpvUndefined;
				} else {
					// "Protected" means "Don't touch this link, because it has been created manually."
					if(effects & tomLinkProtected) {
						*pValue = bpvFalse;
					} else if(effects & tomLink) {
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

STDMETHODIMP TextFont::get_IsInlineMathZone(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			long effects = 0;
			long mask = 0;
			hr = pTextFont2->GetEffects(&effects, &mask);
			if(SUCCEEDED(hr)) {
				if(!(mask & tomMathZoneDisplay)) {
					*pValue = bpvUndefined;
				} else {
					if(effects & tomMathZoneDisplay) {
						*pValue = bpvFalse;
					} else {
						*pValue = bpvTrue;
					}
				}
			}
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_IsLink(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG effects = tomLink;
			LONG mask = tomLink;
			hr = pTextFont2->GetEffects(&effects, &mask);
			if(SUCCEEDED(hr)) {
				if(!(mask & tomLink)) {
					*pValue = bpvUndefined;
				} else {
					if(effects & tomLink) {
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

STDMETHODIMP TextFont::get_IsMathZone(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG isMathZone = tomFalse;
			hr = pTextFont2->GetMathZone(&isMathZone);
			*pValue = static_cast<BooleanPropertyValueConstants>(isMathZone);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::put_IsMathZone(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			hr = pTextFont2->SetMathZone(newValue);
			if(FAILED(hr) && newValue == bpvTrue) {
				hr = pTextFont2->SetEffects(tomMathZone, tomMathZone);
			}
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_Italic(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG italic = tomFalse;
		hr = properties.pTextFont->GetItalic(&italic);
		*pValue = static_cast<BooleanPropertyValueConstants>(italic);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Italic(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetItalic(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_KerningSize(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->GetKerning(pValue);
	}
	return hr;
}

STDMETHODIMP TextFont::put_KerningSize(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetKerning(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Locale(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->GetLanguageID(pValue);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Locale(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetLanguageID(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Name(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->GetName(pValue);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Name(BSTR newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetName(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Outline(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG outline = tomFalse;
		hr = properties.pTextFont->GetOutline(&outline);
		*pValue = static_cast<BooleanPropertyValueConstants>(outline);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Outline(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetOutline(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Protected(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG prot = tomFalse;
		hr = properties.pTextFont->GetProtected(&prot);
		*pValue = static_cast<BooleanPropertyValueConstants>(prot);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Protected(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetProtected(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Revised(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG effects = tomRevised;
			LONG mask = tomRevised;
			hr = pTextFont2->GetEffects(&effects, &mask);
			if(SUCCEEDED(hr)) {
				if(!(mask & tomRevised)) {
					*pValue = bpvUndefined;
				} else {
					if(effects & tomRevised) {
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

STDMETHODIMP TextFont::put_Revised(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && /*newValue != bpvUndefined && */newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG effects = 0;
			switch(newValue) {
				case bpvTrue:
					effects = tomRevised;
					break;
				case bpvToggle:
				{
					LONG mask = tomRevised;
					LONG curVal = 0;
					pTextFont2->GetEffects(&curVal, &mask);
					switch(curVal & tomRevised) {
						case 0:
							effects = tomRevised;
							break;
						/*default:
							effects = 0;
							break;*/
					}
					break;
				}
			}
			hr = pTextFont2->SetEffects(effects, tomRevised);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_Shadow(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG shadow = tomFalse;
		hr = properties.pTextFont->GetShadow(&shadow);
		*pValue = static_cast<BooleanPropertyValueConstants>(shadow);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Shadow(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetShadow(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Size(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->GetSize(pValue);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Size(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetSize(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_SmallCaps(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG smallCaps = tomFalse;
		hr = properties.pTextFont->GetSmallCaps(&smallCaps);
		*pValue = static_cast<BooleanPropertyValueConstants>(smallCaps);
	}
	return hr;
}

STDMETHODIMP TextFont::put_SmallCaps(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetSmallCaps(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Spacing(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->GetSpacing(pValue);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Spacing(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetSpacing(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_StrikeThrough(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG stroke = tomFalse;
		hr = properties.pTextFont->GetStrikeThrough(&stroke);
		*pValue = static_cast<BooleanPropertyValueConstants>(stroke);
	}
	return hr;
}

STDMETHODIMP TextFont::put_StrikeThrough(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvToggle && newValue != bpvUndefined && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetStrikeThrough(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_StyleID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->GetStyle(pValue);
	}
	return hr;
}

STDMETHODIMP TextFont::put_StyleID(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetStyle(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Subscript(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG subscript = tomFalse;
		hr = properties.pTextFont->GetSubscript(&subscript);
		*pValue = static_cast<BooleanPropertyValueConstants>(subscript);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Subscript(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvUndefined && newValue != tomToggle && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetSubscript(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Superscript(BooleanPropertyValueConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BooleanPropertyValueConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG superscript = tomFalse;
		hr = properties.pTextFont->GetSuperscript(&superscript);
		*pValue = static_cast<BooleanPropertyValueConstants>(superscript);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Superscript(BooleanPropertyValueConstants newValue)
{
	if(newValue != bpvUndefined && newValue != tomToggle && newValue != bpvFalse && newValue != bpvTrue) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetSuperscript(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_TeXStyle(TeXFontStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TeXFontStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG texStyle = tomStyleDefault;
			hr = pTextFont2->GetProperty(tomFontPropTeXStyle, &texStyle);
			*pValue = static_cast<TeXFontStyleConstants>(texStyle);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::put_TeXStyle(TeXFontStyleConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			hr = pTextFont2->SetProperty(tomFontPropTeXStyle, newValue);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_UnderlineColorIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		// TODO: currently (Office 2013) this does not retrieve the color index
		LONG underline = tomNone;
		hr = properties.pTextFont->GetUnderline(&underline);
		*pValue = underline & 0xFFFFFF00;
	}
	return hr;
}

STDMETHODIMP TextFont::put_UnderlineColorIndex(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		// TODO: currently (Office 2013) this does not retrieve the color index
		LONG underline = tomNone;
		hr = properties.pTextFont->GetUnderline(&underline);
		underline &= 0xFF;
		underline |= (newValue << 8);
		hr = properties.pTextFont->SetUnderline(underline);
	}
	return hr;
}

STDMETHODIMP TextFont::get_UnderlinePosition(UnderlinePositionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UnderlinePositionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			LONG underlinePosition = tomUnderlinePositionAuto;
			hr = pTextFont2->GetUnderlinePositionMode(&underlinePosition);
			*pValue = static_cast<UnderlinePositionConstants>(underlinePosition);
		} else {
			*pValue = upAuto;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::put_UnderlinePosition(UnderlinePositionConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont2> pTextFont2 = properties.pTextFont;
		if(pTextFont2) {
			hr = pTextFont2->SetUnderlinePositionMode(newValue);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP TextFont::get_UnderlineType(UnderlineTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UnderlineTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		// TODO: currently (Office 2013) this does not retrieve the color index
		LONG underline = tomNone;
		hr = properties.pTextFont->GetUnderline(&underline);
		*pValue = static_cast<UnderlineTypeConstants>(underline & 0xFF);
	}
	return hr;
}

STDMETHODIMP TextFont::put_UnderlineType(UnderlineTypeConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		// TODO: currently (Office 2013) this does not retrieve the color index
		LONG underline = tomNone;
		hr = properties.pTextFont->GetUnderline(&underline);
		underline &= 0xFFFFFF00;
		underline |= newValue;
		hr = properties.pTextFont->SetUnderline(underline);
	}
	return hr;
}

STDMETHODIMP TextFont::get_VerticalOffset(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->GetPosition(pValue);
	}
	return hr;
}

STDMETHODIMP TextFont::put_VerticalOffset(FLOAT newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetPosition(newValue);
	}
	return hr;
}

STDMETHODIMP TextFont::get_Weight(FontWeightConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FontWeightConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		LONG weight = tomUndefined;
		hr = properties.pTextFont->GetWeight(&weight);
		*pValue = static_cast<FontWeightConstants>(weight);
	}
	return hr;
}

STDMETHODIMP TextFont::put_Weight(FontWeightConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->SetWeight(newValue);
	}
	return hr;
}


STDMETHODIMP TextFont::Clone(IRichTextFont** ppClone)
{
	ATLASSERT_POINTER(ppClone, IRichTextFont*);
	if(!ppClone) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	*ppClone = NULL;
	if(properties.pTextFont) {
		CComPtr<ITextFont> pFont = NULL;
		hr = properties.pTextFont->GetDuplicate(&pFont);
		ClassFactory::InitTextFont(pFont, IID_IRichTextFont, reinterpret_cast<LPUNKNOWN*>(ppClone));
	}
	return hr;
}

STDMETHODIMP TextFont::CopySettings(IRichTextFont* pSourceObject)
{
	if(!pSourceObject) {
		// invalid parameter - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont> pFont = pSourceObject;
		hr = properties.pTextFont->SetDuplicate(pFont);
	}
	return hr;
}

STDMETHODIMP TextFont::Equals(IRichTextFont* pCompareAgainst, VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pCompareAgainst, IRichTextFont);
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pCompareAgainst || !pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		CComQIPtr<ITextFont> pFont = pCompareAgainst;
		LONG equal = tomFalse;
		hr = properties.pTextFont->IsEqual(pFont, &equal);
		*pValue = BOOL2VARIANTBOOL(equal == tomTrue);
	}
	return hr;
}

STDMETHODIMP TextFont::Reset(ResetRulesConstants rules)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextFont) {
		hr = properties.pTextFont->Reset(rules);
	}
	return hr;
}