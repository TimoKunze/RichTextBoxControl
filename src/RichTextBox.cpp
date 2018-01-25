// RichTextBox.cpp: Superclasses RichEdit50W.

#include "stdafx.h"
#include "RichTextBox.h"
#include "ClassFactory.h"

#ifdef UNICODE
	LPCTSTR WNDCLASS_RICHEDIT50 = MSFTEDIT_CLASS;
	LPCTSTR WNDCLASS_RICHEDIT60 = L"RICHEDIT60W";
#else
	// I'm not sure that there are ANSI versions available, so better use the Unicode window classes
	LPCTSTR WNDCLASS_RICHEDIT50 = "RICHEDIT50W";
	LPCTSTR WNDCLASS_RICHEDIT60 = "RICHEDIT60W";
#endif

#ifndef PFNHYPHENATEPROC
	typedef void (WINAPI* PFNHYPHENATEPROC)(WCHAR*, LANGID, long, HYPHRESULT*);
#endif


// initialize complex constants
IMEModeConstants RichTextBox::IMEFlags::chineseIMETable[10] = {
    imeOff,
    imeOff,
    imeOff,
    imeOff,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOff
};

IMEModeConstants RichTextBox::IMEFlags::japaneseIMETable[10] = {
    imeDisable,
    imeDisable,
    imeOff,
    imeOff,
    imeHiragana,
    imeHiragana,
    imeKatakana,
    imeKatakanaHalf,
    imeAlphaFull,
    imeAlpha
};

IMEModeConstants RichTextBox::IMEFlags::koreanIMETable[10] = {
    imeDisable,
    imeDisable,
    imeAlpha,
    imeAlpha,
    imeHangulFull,
    imeHangul,
    imeHangulFull,
    imeHangul,
    imeAlphaFull,
    imeAlpha
};

FONTDESC RichTextBox::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


#pragma warning(push)
#pragma warning(disable: 4355)     // passing "this" to a member constructor
RichTextBox::RichTextBox()
{
	gdiplusToken = NULL;

	properties.font.InitializePropertyWatcher(this, DISPID_RTB_FONT);
	properties.linkMouseIcon.InitializePropertyWatcher(this, DISPID_RTB_LINKMOUSEICON);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_RTB_MOUSEICON);

	SIZEL size = {200, 200};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	// initialize
	pDocumentStorage = NULL;
	if(RunTimeHelper::IsVista()) {
		CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
		ATLASSUME(pWICImagingFactory);
	}
	
	Gdiplus::GdiplusStartupInput gdiplusStartupInput;
	Gdiplus::GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);
}
#pragma warning(pop)

RichTextBox::~RichTextBox()
{
	/* Don't unload msftedit.dll, because our window class, which superclasses RICHEDIT50W,
	 * seems to be registered only once per process. Unloading msftedit.dll destroys any
	 * reference to RICHEDIT50W, so our window class won't work anymore. */
	/*if(hMsftEdit) {
		FreeLibrary(hMsftEdit);
	}*/
	if(gdiplusToken) {
		Gdiplus::GdiplusShutdown(gdiplusToken);
		gdiplusToken = NULL;
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP RichTextBox::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichTextBox, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK RichTextBox::QueryITextDocumentInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextDocument), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextDocument2), queriedInterface)) {
		RichTextBox* pRTB = reinterpret_cast<RichTextBox*>(pThis);
		return pRTB->properties.pTextDocument->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}


STDMETHODIMP RichTextBox::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<RichTextBox>::Load(pPropertyBag, pErrorLog);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);

		CComBSTR bstr;
		hr = pPropertyBag->Read(OLESTR("AutoDetectedURLSchemes"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_AutoDetectedURLSchemes(bstr);
		VariantClear(&propertyValue);

		hr = pPropertyBag->Read(OLESTR("Text"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_Text(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP RichTextBox::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<RichTextBox>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.autoDetectedURLSchemes.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("AutoDetectedURLSchemes"), &propertyValue);
		VariantClear(&propertyValue);

		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.text.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("Text"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP RichTextBox::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 1 VT_I2 property...
	pSize->QuadPart += 1 * (sizeof(VARTYPE) + sizeof(SHORT));
	// ...and 34 VT_I4 properties...
	pSize->QuadPart += 34 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 1 VT_R4 property...
	pSize->QuadPart += 1 * (sizeof(VARTYPE) + sizeof(FLOAT));
	// ...and 1 VT_R8 property...
	pSize->QuadPart += 1 * (sizeof(VARTYPE) + sizeof(DOUBLE));
	// ...and 44 VT_BOOL properties...
	pSize->QuadPart += 44 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 2 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.text.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.autoDetectedURLSchemes.ByteLength() + sizeof(OLECHAR);

	// ...and 3 VT_DISPATCH properties
	pSize->QuadPart += 3 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.linkMouseIcon.pPictureDisp) {
		if(FAILED(properties.linkMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.linkMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP RichTextBox::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x13131313/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version > 0x0100) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL acceptTabKey = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL adjustLineHeightForEastAsianLanguages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowInputThroughTSF = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowMathZoneInsertion = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowObjectInsertionThroughTSF = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowTableInsertion = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowTSFProofingTips = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowTSFSmartTags = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL alwaysShowScrollBars = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL alwaysShowSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR autoDetectedURLSchemes;
	if(FAILED(hr = autoDetectedURLSchemes.ReadFromStream(pStream))) {
		return hr;
	}
	if(!autoDetectedURLSchemes) {
		autoDetectedURLSchemes = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AutoDetectURLsConstants autoDetectURLs = static_cast<AutoDetectURLsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AutoScrollingConstants autoScrolling = static_cast<AutoScrollingConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL autoSelectWordOnTrackSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL beepOnMaxText = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	CharacterConversionConstants characterConversion = static_cast<CharacterConversionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HAlignmentConstants defaultMathZoneHAlignment = static_cast<HAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_R4, &propertyValue))) {
		return hr;
	}
	FLOAT defaultTabWidth = propertyValue.fltVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DenoteEmptyMathArgumentsConstants denoteEmptyMathArguments = static_cast<DenoteEmptyMathArgumentsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL detectDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL disableIMEOperations = propertyValue.boolVal;
	/*if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL discardCompositionStringOnIMECancel = propertyValue.boolVal;*/
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayHyperlinkTooltips = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayZeroWidthTableCellBorders = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL draftMode = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragScrollTimeBase = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dropWordsOnWordBoundariesOnly = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL emulateSimpleTextBox = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL extendFontBackColorToWholeLine = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS formattingRectangleHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XPOS_PIXELS formattingRectangleLeft = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YPOS_PIXELS formattingRectangleTop = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS formattingRectangleWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL growNAryOperators = propertyValue.boolVal;
	//if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
	//	return hr;
	//}
	//HAlignmentConstants hAlignment = static_cast<HAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I2, &propertyValue))) {
		return hr;
	}
	SHORT hyphenationWordWrapZoneWidth = propertyValue.iVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEConversionModeConstants imeConversionMode = static_cast<IMEConversionModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants imeMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LimitsLocationConstants integralLimitsLocation = static_cast<LimitsLocationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS leftMargin = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL letClientHandleAllIMEOperations = propertyValue.boolVal;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pLinkMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pLinkMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants linkMousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL logicalCaret = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MathLineBreakBehaviorConstants mathLineBreakBehavior = static_cast<MathLineBreakBehaviorConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maxTextLength = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maxUndoQueueSize = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL modified = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL multiLine = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL multiSelect = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LimitsLocationConstants nAryLimitsLocation = static_cast<LimitsLocationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL noInputSequenceCheck = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLEDragImageStyleConstants oleDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL rawSubScriptAndSuperScriptOperators = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL readOnly = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RegisterForOLEDragDropConstants registerForOLEDragDrop = static_cast<RegisterForOLEDragDropConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RichEditVersionConstants richEditVersion = static_cast<RichEditVersionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS rightMargin = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL rightToLeftMathZones = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ScrollBarsConstants scrollBars = static_cast<ScrollBarsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL scrollToTopOnKillFocus = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showSelectionBar = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL smartSpacingOnDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL switchFontOnIMEInput = propertyValue.boolVal;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR text;
	if(FAILED(hr = text.ReadFromStream(pStream))) {
		return hr;
	}
	if(!text) {
		text = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TextFlowConstants textFlow = static_cast<TextFlowConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TextOrientationConstants textOrientation = static_cast<TextOrientationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TSFModeBiasConstants tsfModeBias = static_cast<TSFModeBiasConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useBkAcetateColorForTextSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useBuiltInSpellChecking = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useCustomFormattingRectangle = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useDualFontMode = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSmallerFontForNestedFractions = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useTextServicesFramework = propertyValue.boolVal;
	/*if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useTouchKeyboardAutoCorrection = propertyValue.boolVal;*/
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useWindowsThemes = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_R8, &propertyValue))) {
		return hr;
	}
	DOUBLE zoomRatio = propertyValue.dblVal;

	hr = put_AcceptTabKey(acceptTabKey);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AdjustLineHeightForEastAsianLanguages(adjustLineHeightForEastAsianLanguages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowInputThroughTSF(allowInputThroughTSF);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowMathZoneInsertion(allowMathZoneInsertion);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowObjectInsertionThroughTSF(allowObjectInsertionThroughTSF);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowTableInsertion(allowTableInsertion);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowTSFProofingTips(allowTSFProofingTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowTSFSmartTags(allowTSFSmartTags);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AlwaysShowScrollBars(alwaysShowScrollBars);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AlwaysShowSelection(alwaysShowSelection);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoDetectedURLSchemes(autoDetectedURLSchemes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoDetectURLs(autoDetectURLs);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoScrolling(autoScrolling);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoSelectWordOnTrackSelection(autoSelectWordOnTrackSelection);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BeepOnMaxText(beepOnMaxText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CharacterConversion(characterConversion);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DefaultMathZoneHAlignment(defaultMathZoneHAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DefaultTabWidth(defaultTabWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DenoteEmptyMathArguments(denoteEmptyMathArguments);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DetectDragDrop(detectDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisableIMEOperations(disableIMEOperations);
	if(FAILED(hr)) {
		return hr;
	}
	/*hr = put_DiscardCompositionStringOnIMECancel(discardCompositionStringOnIMECancel);
	if(FAILED(hr)) {
		return hr;
	}*/
	hr = put_DisplayHyperlinkTooltips(displayHyperlinkTooltips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayZeroWidthTableCellBorders(displayZeroWidthTableCellBorders);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DraftMode(draftMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragScrollTimeBase(dragScrollTimeBase);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DropWordsOnWordBoundariesOnly(dropWordsOnWordBoundariesOnly);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EmulateSimpleTextBox(emulateSimpleTextBox);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ExtendFontBackColorToWholeLine(extendFontBackColorToWholeLine);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FormattingRectangleHeight(formattingRectangleHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FormattingRectangleLeft(formattingRectangleLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FormattingRectangleTop(formattingRectangleTop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FormattingRectangleWidth(formattingRectangleWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GrowNAryOperators(growNAryOperators);
	if(FAILED(hr)) {
		return hr;
	}
	//hr = put_HAlignment(hAlignment);
	//if(FAILED(hr)) {
	//	return hr;
	//}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HyphenationWordWrapZoneWidth(hyphenationWordWrapZoneWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IMEConversionMode(imeConversionMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IMEMode(imeMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IntegralLimitsLocation(integralLimitsLocation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LeftMargin(leftMargin);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LetClientHandleAllIMEOperations(letClientHandleAllIMEOperations);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LinkMouseIcon(pLinkMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LinkMousePointer(linkMousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LogicalCaret(logicalCaret);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MathLineBreakBehavior(mathLineBreakBehavior);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxTextLength(maxTextLength);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxUndoQueueSize(maxUndoQueueSize);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Modified(modified);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MultiLine(multiLine);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MultiSelect(multiSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_NAryLimitsLocation(nAryLimitsLocation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_NoInputSequenceCheck(noInputSequenceCheck);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OLEDragImageStyle(oleDragImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RawSubScriptAndSuperScriptOperators(rawSubScriptAndSuperScriptOperators);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ReadOnly(readOnly);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RichEditVersion(richEditVersion);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightMargin(rightMargin);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeftMathZones(rightToLeftMathZones);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollBars(scrollBars);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollToTopOnKillFocus(scrollToTopOnKillFocus);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowSelectionBar(showSelectionBar);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SmartSpacingOnDrop(smartSpacingOnDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SwitchFontOnIMEInput(switchFontOnIMEInput);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TextFlow(textFlow);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TextOrientation(textOrientation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TSFModeBias(tsfModeBias);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseBkAcetateColorForTextSelection(useBkAcetateColorForTextSelection);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseBuiltInSpellChecking(useBuiltInSpellChecking);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseCustomFormattingRectangle(useCustomFormattingRectangle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseDualFontMode(useDualFontMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSmallerFontForNestedFractions(useSmallerFontForNestedFractions);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseTextServicesFramework(useTextServicesFramework);
	if(FAILED(hr)) {
		return hr;
	}
	/*hr = put_UseTouchKeyboardAutoCorrection(useTouchKeyboardAutoCorrection);
	if(FAILED(hr)) {
		return hr;
	}*/
	hr = put_UseWindowsThemes(useWindowsThemes);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ZoomRatio(zoomRatio);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP RichTextBox::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x13131313/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0100;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.acceptTabKey);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.adjustLineHeightForEastAsianLanguages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowInputThroughTSF);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowMathZoneInsertion);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowObjectInsertionThroughTSF);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowTableInsertion);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowTSFProofingTips);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowTSFSmartTags);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.alwaysShowScrollBars);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	VARTYPE vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.autoDetectedURLSchemes.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.lVal = properties.autoDetectURLs;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.autoScrolling;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoSelectWordOnTrackSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.beepOnMaxText);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.characterConversion;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.defaultMathZoneHAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_R4;
	propertyValue.fltVal = properties.defaultTabWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.denoteEmptyMathArguments;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.detectDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.disableIMEOperations);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	/*propertyValue.boolVal = BOOL2VARIANTBOOL(properties.discardCompositionStringOnIMECancel);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}*/
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayHyperlinkTooltips);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayZeroWidthTableCellBorders);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.draftMode);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.dragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dropWordsOnWordBoundariesOnly);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.emulateSimpleTextBox);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.extendFontBackColorToWholeLine);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.formattingRectangle.Height();
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.formattingRectangle.left;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.formattingRectangle.top;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.formattingRectangle.Width();
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.growNAryOperators);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	//propertyValue.lVal = properties.hAlignment;
	//if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
	//	return hr;
	//}
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I2;
	propertyValue.iVal = properties.hyphenationWordWrapZoneWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.IMEConversionMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.IMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.integralLimitsLocation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.leftMargin;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.letClientHandleAllIMEOperations);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.linkMouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.linkMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.linkMousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.logicalCaret);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.mathLineBreakBehavior;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maxTextLength;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maxUndoQueueSize;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.modified);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiLine);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiSelect);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.nAryLimitsLocation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.noInputSequenceCheck);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.oleDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rawSubScriptAndSuperScriptOperators);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.readOnly);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.registerForOLEDragDrop;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.richEditVersion;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.rightMargin;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rightToLeftMathZones);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.scrollBars;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.scrollToTopOnKillFocus);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showSelectionBar);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.smartSpacingOnDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.switchFontOnIMEInput);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.text.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.textFlow;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.textOrientation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.TSFModeBias;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useBkAcetateColorForTextSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useBuiltInSpellChecking);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useCustomFormattingRectangle);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useDualFontMode);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSmallerFontForNestedFractions);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useTextServicesFramework);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	/*propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useTouchKeyboardAutoCorrection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}*/
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useWindowsThemes);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_R8;
	propertyValue.dblVal = properties.zoomRatio;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND RichTextBox::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	LPCTSTR pRichEditClassName = GetWndClassName();
	TCHAR pClassName[128];
	#ifdef UNICODE
		lstrcpyn(pClassName, TEXT("RichTextBoxU_"), 128);
	#else
		lstrcpyn(pClassName, TEXT("RichTextBoxA_"), 128);
	#endif
	StringCchCat(pClassName, 128, pRichEditClassName);

	GetWndClassInfo().m_wc.lpszClassName = pClassName;
	GetWndClassInfo().m_lpszOrigName = pRichEditClassName;

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<RichTextBox>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT RichTextBox::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("RichTextBox ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void RichTextBox::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP RichTextBox::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<RichTextBox>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	if(properties.resetTextToName) {
		properties.resetTextToName = FALSE;

		BSTR buffer = SysAllocString(L"");
		if(SUCCEEDED(GetAmbientDisplayName(buffer))) {
			properties.text.AssignBSTR(buffer);
		} else {
			SysFreeString(buffer);
		}
	}

	return hr;
}

STDMETHODIMP RichTextBox::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL RichTextBox::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		//ATLASSERT((dialogCode & (DLGC_WANTCHARS | DLGC_HASSETSEL | DLGC_WANTARROWS)) == (DLGC_WANTCHARS | DLGC_HASSETSEL | DLGC_WANTARROWS));
		if(pMessage->wParam == VK_TAB) {
			if(dialogCode & DLGC_WANTTAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
		}
	}
	return CComControl<RichTextBox>::PreTranslateAccelerator(pMessage, hReturnValue);
}

HIMAGELIST RichTextBox::CreateLegacyDragImage(ITextRange* pTextRange, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	// retrieve window details
	BOOL layoutRTL = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);

	// create the DCs we'll draw into
	HDC hCompatibleDC = GetDC();
	CDC memoryDC;
	memoryDC.CreateCompatibleDC(hCompatibleDC);
	CDC maskMemoryDC;
	maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

	// calculate the bounding rectangle of the text
	WTL::CRect textBoundingRect;
	// TODO: Make this work with sub-ranges.
	if(SUCCEEDED(pTextRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomStart + TA_TOP + TA_LEFT), &textBoundingRect.left, &textBoundingRect.top)) && SUCCEEDED(pTextRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomEnd + TA_BOTTOM + TA_RIGHT), &textBoundingRect.right, &textBoundingRect.bottom))) {
		if(pBoundingRectangle) {
			*pBoundingRectangle = textBoundingRect;
		}
	}

	// calculate drag image size and upper-left corner
	SIZE dragImageSize = {0};
	if(pUpperLeftPoint) {
		pUpperLeftPoint->x = textBoundingRect.left;
		pUpperLeftPoint->y = textBoundingRect.top;
	}
	dragImageSize.cx = textBoundingRect.Width();
	dragImageSize.cy = textBoundingRect.Height();

	// offset RECTs
	SIZE offset = {0};
	offset.cx = textBoundingRect.left;
	offset.cy = textBoundingRect.top;
	textBoundingRect.OffsetRect(-offset.cx, -offset.cy);

	// setup the DCs we'll draw into
	memoryDC.SetBkColor(GetSysColor(COLOR_WINDOW));
	memoryDC.SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
	memoryDC.SetBkMode(TRANSPARENT);

	// create drag image bitmap
	BOOL doAlphaChannelProcessing = RunTimeHelper::IsCommCtrl6();
	BITMAPINFO bitmapInfo = {0};
	if(doAlphaChannelProcessing) {
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = dragImageSize.cx;
		bitmapInfo.bmiHeader.biHeight = -dragImageSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
	}
	CBitmap dragImage;
	LPRGBQUAD pDragImageBits = NULL;
	if(doAlphaChannelProcessing) {
		dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
	} else {
		dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageSize.cx, dragImageSize.cy);
	}
	HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
	CBitmap dragImageMask;
	dragImageMask.CreateBitmap(dragImageSize.cx, dragImageSize.cy, 1, 1, NULL);
	HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);

	// initialize the bitmap
	// we need a transparent background
	if(doAlphaChannelProcessing && pDragImageBits) {
		// we need a transparent background
		LPRGBQUAD pPixel = pDragImageBits;
		for(int y = 0; y < dragImageSize.cy; ++y) {
			for(int x = 0; x < dragImageSize.cx; ++x, ++pPixel) {
				pPixel->rgbRed = 0xFF;
				pPixel->rgbGreen = 0xFF;
				pPixel->rgbBlue = 0xFF;
				pPixel->rgbReserved = 0x00;
			}
		}
	} else {
		memoryDC.FillRect(&textBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
	}
	maskMemoryDC.FillRect(&textBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

	if(doAlphaChannelProcessing) {
		// correct the alpha channel
		LPRGBQUAD pPixel = pDragImageBits;
		POINT pt;
		for(pt.y = 0; pt.y < dragImageSize.cy; ++pt.y) {
			for(pt.x = 0; pt.x < dragImageSize.cx; ++pt.x, ++pPixel) {
				if(layoutRTL) {
					// we're working on raw data, so we've to handle WS_EX_LAYOUTRTL ourselves
					POINT pt2 = pt;
					pt2.x = dragImageSize.cx - pt.x - 1;
					if(maskMemoryDC.GetPixel(pt2.x, pt2.y) == 0x00000000) {
						// items are not themed
						if(pPixel->rgbReserved == 0x00) {
							pPixel->rgbReserved = 0xFF;
						}
					}
				} else {
					// layout is left to right
					if(maskMemoryDC.GetPixel(pt.x, pt.y) == 0x00000000) {
						if(pPixel->rgbReserved == 0x00) {
							pPixel->rgbReserved = 0xFF;
						}
					}
				}
			}
		}
	}

	memoryDC.SelectBitmap(hPreviousBitmap);
	maskMemoryDC.SelectBitmap(hPreviousBitmapMask);

	// create the imagelist
	HIMAGELIST hDragImageList = ImageList_Create(dragImageSize.cx, dragImageSize.cy, (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
	ImageList_SetBkColor(hDragImageList, CLR_NONE);
	ImageList_Add(hDragImageList, dragImage, dragImageMask);

	ReleaseDC(hCompatibleDC);

	return hDragImageList;
}

BOOL RichTextBox::CreateLegacyOLEDragImage(ITextRange* pTextRange, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pTextRange);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	// use a normal legacy drag image as base
	POINT upperLeftPoint = {0};
	HIMAGELIST hImageList = CreateLegacyDragImage(pTextRange, &upperLeftPoint, NULL);
	if(hImageList) {
		// retrieve the drag image's size
		int bitmapHeight;
		int bitmapWidth;
		ImageList_GetIconSize(hImageList, &bitmapWidth, &bitmapHeight);
		pDragImage->sizeDragImage.cx = bitmapWidth;
		pDragImage->sizeDragImage.cy = bitmapHeight;

		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		pDragImage->hbmpDragImage = NULL;

		if(RunTimeHelper::IsCommCtrl6()) {
			// handle alpha channel
			IImageList* pImgLst = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
				HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
				if(pfnHIMAGELIST_QueryInterface) {
					pfnHIMAGELIST_QueryInterface(hImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
				}
				FreeLibrary(hMod);
			}
			if(!pImgLst) {
				pImgLst = reinterpret_cast<IImageList*>(hImageList);
				pImgLst->AddRef();
			}
			ATLASSUME(pImgLst);

			DWORD flags = 0;
			pImgLst->GetItemFlags(0, &flags);
			if(flags & ILIF_ALPHA) {
				// the drag image makes use of the alpha channel
				IMAGEINFO imageInfo = {0};
				ImageList_GetImageInfo(hImageList, 0, &imageInfo);

				// fetch raw data
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
				bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pSourceBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * sizeof(RGBQUAD)));
				GetDIBits(memoryDC, imageInfo.hbmImage, 0, pDragImage->sizeDragImage.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
				// create target bitmap
				LPRGBQUAD pDragImageBits = NULL;
				pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
				pDragImage->crColorKey = 0xFFFFFFFF;

				// transfer raw data
				CopyMemory(pDragImageBits, pSourceBits, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * 4);

				// clean up
				HeapFree(GetProcessHeap(), 0, pSourceBits);
				::DeleteObject(imageInfo.hbmImage);
				::DeleteObject(imageInfo.hbmMask);
			}
			pImgLst->Release();
		}

		if(!pDragImage->hbmpDragImage) {
			// fallback mode
			memoryDC.SetBkMode(TRANSPARENT);

			// create target bitmap
			HDC hCompatibleDC = ::GetDC(NULL);
			pDragImage->hbmpDragImage = CreateCompatibleBitmap(hCompatibleDC, bitmapWidth, bitmapHeight);
			::ReleaseDC(NULL, hCompatibleDC);
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(pDragImage->hbmpDragImage);

			// draw target bitmap
			pDragImage->crColorKey = RGB(0xF4, 0x00, 0x00);
			CBrush backroundBrush;
			backroundBrush.CreateSolidBrush(pDragImage->crColorKey);
			memoryDC.FillRect(WTL::CRect(0, 0, bitmapWidth, bitmapHeight), backroundBrush);
			ImageList_Draw(hImageList, 0, memoryDC, 0, 0, ILD_NORMAL);

			// clean up
			memoryDC.SelectBitmap(hPreviousBitmap);
		}

		ImageList_Destroy(hImageList);

		if(pDragImage->hbmpDragImage) {
			// retrieve the offset
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			if(GetExStyle() & WS_EX_LAYOUTRTL) {
				pDragImage->ptOffset.x = upperLeftPoint.x + pDragImage->sizeDragImage.cx - mousePosition.x;
			} else {
				pDragImage->ptOffset.x = mousePosition.x - upperLeftPoint.x;
			}
			pDragImage->ptOffset.y = mousePosition.y - upperLeftPoint.y;

			succeeded = TRUE;
		}
	}

	return succeeded;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP RichTextBox::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);

	// For rich edit controls, move is the default effect.
	if(keyState & MK_CONTROL) {
		if(*pEffect & DROPEFFECT_COPY) {
			*pEffect = DROPEFFECT_COPY;
		}
	} else {
		if(*pEffect & DROPEFFECT_MOVE) {
			*pEffect = DROPEFFECT_MOVE;
		}
	}

	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
		if(dragDropStatus.useItemCountLabelHack) {
			dragDropStatus.pDropTargetHelper->DragLeave();
			dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			dragDropStatus.useItemCountLabelHack = FALSE;
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);

	// For rich edit controls, move is the default effect.
	if(keyState & MK_CONTROL) {
		if(*pEffect & DROPEFFECT_COPY) {
			*pEffect = DROPEFFECT_COPY;
		}
	} else {
		if(*pEffect & DROPEFFECT_MOVE) {
			*pEffect = DROPEFFECT_MOVE;
		}
	}

	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	// For rich edit controls, move is the default effect.
	if(keyState & MK_CONTROL) {
		if(*pEffect & DROPEFFECT_COPY) {
			*pEffect = DROPEFFECT_COPY;
		}
	} else {
		if(*pEffect & DROPEFFECT_MOVE) {
			*pEffect = DROPEFFECT_MOVE;
		}
	}

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSource
STDMETHODIMP RichTextBox::GiveFeedback(DWORD effect)
{
	VARIANT_BOOL useDefaultCursors = VARIANT_TRUE;
	//if(flags.usingThemes && RunTimeHelper::IsVista()) {
		ATLASSUME(dragDropStatus.pSourceDataObject);

		BOOL isShowingLayered = FALSE;
		FORMATETC format = {0};
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("IsShowingLayered")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		STGMEDIUM medium = {0};
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				isShowingLayered = *reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		BOOL useDropDescriptionHack = FALSE;
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				useDropDescriptionHack = *reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}

		if(isShowingLayered && properties.oleDragImageStyle != odistClassic) {
			SetCursor(static_cast<HCURSOR>(LoadImage(NULL, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED)));
			useDefaultCursors = VARIANT_FALSE;
		}
		if(useDropDescriptionHack) {
			// this will make drop descriptions work
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragWindow")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
				if(medium.hGlobal) {
					// WM_USER + 1 (with wParam = 0 and lParam = 0) hides the drag image
					#define WM_SETDROPEFFECT				WM_USER + 2     // (wParam = DCID_*, lParam = 0)
					#define DDWM_UPDATEWINDOW				WM_USER + 3     // (wParam = 0, lParam = 0)
					typedef enum DROPEFFECTS
					{
						DCID_NULL = 0,
						DCID_NO = 1,
						DCID_MOVE = 2,
						DCID_COPY = 3,
						DCID_LINK = 4,
						DCID_MAX = 5
					} DROPEFFECTS;

					HWND hWndDragWindow = *reinterpret_cast<HWND*>(GlobalLock(medium.hGlobal));
					GlobalUnlock(medium.hGlobal);

					DROPEFFECTS dropEffect = DCID_NULL;
					switch(effect) {
						case DROPEFFECT_NONE:
							dropEffect = DCID_NO;
							break;
						case DROPEFFECT_COPY:
							dropEffect = DCID_COPY;
							break;
						case DROPEFFECT_MOVE:
							dropEffect = DCID_MOVE;
							break;
						case DROPEFFECT_LINK:
							dropEffect = DCID_LINK;
							break;
					}
					if(::IsWindow(hWndDragWindow)) {
						::PostMessage(hWndDragWindow, WM_SETDROPEFFECT, dropEffect, 0);
					}
				}
				ReleaseStgMedium(&medium);
			}
		}
	//}

	Raise_OLEGiveFeedback(effect, &useDefaultCursors);
	return (useDefaultCursors == VARIANT_FALSE ? S_OK : DRAGDROP_S_USEDEFAULTCURSORS);
}

STDMETHODIMP RichTextBox::QueryContinueDrag(BOOL pressedEscape, DWORD keyState)
{
	HRESULT actionToContinueWith = S_OK;
	if(pressedEscape) {
		actionToContinueWith = DRAGDROP_S_CANCEL;
	} else if(!(keyState & dragDropStatus.draggingMouseButton)) {
		actionToContinueWith = DRAGDROP_S_DROP;
	}
	Raise_OLEQueryContinueDrag(pressedEscape, keyState, &actionToContinueWith);
	return actionToContinueWith;
}
// implementation of IDropSource
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSourceNotify
STDMETHODIMP RichTextBox::DragEnterTarget(HWND hWndTarget)
{
	Raise_OLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	return S_OK;
}

STDMETHODIMP RichTextBox::DragLeaveTarget(void)
{
	Raise_OLEDragLeavePotentialTarget();
	return S_OK;
}
// implementation of IDropSourceNotify
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IRichEditOleCallback
STDMETHODIMP RichTextBox::GetNewStorage(LPSTORAGE FAR* ppStorage)
{
	if(!ppStorage) {
		return E_INVALIDARG;
	}

	if(pDocumentStorage) {
		CAtlString s;
		s.Format(_T("REOLE_%u"), ++nextEmbeddedObjectIndex);
		BSTR pName = s.AllocSysString();
		HRESULT hr = pDocumentStorage->CreateStorage(pName, STGM_TRANSACTED | STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE, 0, 0, ppStorage);
		SysFreeString(pName);
		return hr;
	} else {
		ATLASSERT(FALSE && "IRichEditOleCallback::GetNewStorage - pDocumentStorage is NULL");
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::GetInPlaceContext(LPOLEINPLACEFRAME FAR* ppInPlaceFrame, LPOLEINPLACEUIWINDOW FAR* ppInPlaceUIWindow, LPOLEINPLACEFRAMEINFO pInPlaceFrameInfo)
{
	if(!ppInPlaceFrame || !ppInPlaceUIWindow || !pInPlaceFrameInfo) {
		return E_INVALIDARG;
	}

	// NOTE: This method is called for instance when adding an OLEObject with the ProgID "MediaPlayer.MediaPlayer.1"
	HRESULT hr = E_NOTIMPL;
	if(m_spInPlaceSite) {
		RECT posRect = {0};
		RECT clipRect = {0};
		hr = m_spInPlaceSite->GetWindowContext(ppInPlaceFrame, ppInPlaceUIWindow, &posRect, &clipRect, pInPlaceFrameInfo);
	}
	return hr;
}

STDMETHODIMP RichTextBox::ShowContainerUI(BOOL /*show*/)
{
	// TODO: Maybe raise an event?
	// NOTE: This method is called for instance when adding an OLEObject with the ProgID "MediaPlayer.MediaPlayer.1"
	return S_OK;
}

STDMETHODIMP RichTextBox::QueryInsertObject(LPCLSID pCLSID, LPSTORAGE pStorage, LONG insertionPosition)
{
	if(!pCLSID || !pStorage) {
		return E_INVALIDARG;
	}

	/*CLSID realCLSID = CLSID_NULL;
	HRESULT hr = ReadClassStg(pStorage, &realCLSID);*/

	// TODO: For some reason insertionPosition seems to be always 0 or REO_CP_SELECTION.
	if(flags.silentObjectInsertions == 0) {
		CComPtr<IRichTextRange> pTextRange = NULL;
		if(insertionPosition == REO_CP_SELECTION) {
			get_SelectedTextRange(&pTextRange);
		} else {
			get_TextRange(insertionPosition, insertionPosition, &pTextRange);
		}
		if(pTextRange) {
			LPOLESTR pCLSIDString = NULL;
			StringFromCLSID(*pCLSID, &pCLSIDString);
			BSTR clsid = SysAllocString(pCLSIDString);
			VARIANT_BOOL cancel = VARIANT_FALSE;
			Raise_InsertingOLEObject(pTextRange, &clsid, *reinterpret_cast<LONG*>(&pStorage), &cancel);
			CLSIDFromString(clsid, pCLSID);
			SysFreeString(clsid);
			return (cancel == VARIANT_FALSE ? S_OK : S_FALSE);
		}
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::DeleteObject(LPOLEOBJECT pOLEObject)
{
	if(!pOLEObject) {
		return E_INVALIDARG;
	}

	if(flags.silentObjectDeletions == 0) {
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(SendMessage(EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			LONG objectCount = pRichEditOle->GetObjectCount();
			CComPtr<IRichOLEObject> pDeletedObject = NULL;
			for(LONG objectIndex = 0; objectIndex < objectCount && !pDeletedObject; objectIndex++) {
				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);
				if(SUCCEEDED(pRichEditOle->GetObject(objectIndex, &objectDetails, REO_GETOBJ_POLEOBJ))) {
					if(objectDetails.poleobj == pOLEObject) {
						CComPtr<IRichTextRange> pRichTextRange = NULL;
						CComPtr<ITextRange> pTextRange = NULL;
						if(SUCCEEDED(get_TextRange(objectDetails.cp, objectDetails.cp, &pRichTextRange)) && pRichTextRange && SUCCEEDED(pRichTextRange->QueryInterface(IID_PPV_ARGS(&pTextRange))) && pTextRange) {
							pTextRange->EndOf(tomObject, tomExtend, NULL);
						}
						pDeletedObject = ClassFactory::InitOLEObject(pTextRange, objectDetails.poleobj, this);
					}
					if(objectDetails.poleobj) {
						objectDetails.poleobj->Release();
						objectDetails.poleobj = NULL;
					}
				}
			}

			VARIANT_BOOL cancel = VARIANT_FALSE;
			Raise_RemovingOLEObject(pDeletedObject, &cancel);
			// does not work:
			//return (cancel == VARIANT_FALSE ? S_OK : S_FALSE);
			return S_OK;
		}
	}
	return E_NOTIMPL;
}

STDMETHODIMP RichTextBox::QueryAcceptData(LPDATAOBJECT pDataObject, CLIPFORMAT FAR* pFormat, DWORD operationFlag, BOOL operationIsReal, HGLOBAL hMetaPict)
{
	if(!pDataObject || !pFormat) {
		return E_INVALIDARG;
	}

	HRESULT hr = E_NOTIMPL;
	if(properties.pTextDocument) {
		CComPtr<IOLEDataObject> pOleDataObject = ClassFactory::InitOLEDataObject(pDataObject);
		QueryAcceptDataConstants acceptData = qadLetNativeControlDecide;
		// translate to VB's format constants
		LONG formatID = *pFormat;
		if(static_cast<UINT>(formatID) == RegisterClipboardFormat(CF_RTF)) {
			formatID = 0xffffbf01;
		}
		Raise_QueryAcceptData(pOleDataObject, &formatID, static_cast<ClipboardOperationTypeConstants>(operationFlag), BOOL2VARIANTBOOL(operationIsReal), HandleToLong(hMetaPict), &acceptData);
		// translate from VB's format constants
		if(formatID == 0xffffbf01) {
			formatID = RegisterClipboardFormat(CF_RTF);
		}
		*pFormat = LOWORD(formatID);
		switch(acceptData) {
			case qadAcceptData:
				hr = S_FALSE;
				break;
			case qadRefuseData:
				hr = E_ABORT;
				break;
			case qadLetNativeControlDecide:
				hr = S_OK;
				break;
		}
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP RichTextBox::ContextSensitiveHelp(BOOL /*enterMode*/)
{
	ATLASSERT(FALSE && "TODO: Implement IRichEditOleCallback::ContextSensitiveHelp");
	return E_NOTIMPL;
}

STDMETHODIMP RichTextBox::GetClipboardData(CHARRANGE FAR* pCharRange, DWORD operationFlag, LPDATAOBJECT FAR* ppDataObject)
{
	if(!pCharRange || !ppDataObject) {
		return E_INVALIDARG;
	}

	HRESULT hr = E_NOTIMPL;
	if(properties.pTextDocument) {
		CComPtr<ITextRange> pTextRange = NULL;
		if(SUCCEEDED(properties.pTextDocument->Range(pCharRange->cpMin, pCharRange->cpMax, &pTextRange))) {
			CComPtr<IRichTextRange> pRichTextRange = ClassFactory::InitTextRange(pTextRange, this);
			VARIANT_BOOL useCustomDataObject = VARIANT_FALSE;
			LONG pDataObj = *reinterpret_cast<LONG*>(ppDataObject);
			Raise_GetDataObject(pRichTextRange, static_cast<ClipboardOperationTypeConstants>(operationFlag), &pDataObj, &useCustomDataObject);
			if(useCustomDataObject != VARIANT_FALSE) {
				*ppDataObject = *reinterpret_cast<LPDATAOBJECT*>(&pDataObj);
				hr = S_OK;
			}
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::GetDragDropEffect(BOOL getSourceEffects, DWORD keyState, LPDWORD pEffect)
{
	if(!pEffect) {
		return E_INVALIDARG;
	}

	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
	VARIANT_BOOL skipDefaultProcessing = VARIANT_FALSE;
	Raise_GetDragDropEffect(BOOL2VARIANTBOOL(getSourceEffects), button, shift, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &skipDefaultProcessing);
	if(skipDefaultProcessing != VARIANT_FALSE) {
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
	/*if(properties.registerForOLEDragDrop == rfoddNativeDragDrop) {
		return E_NOTIMPL;
	} else {
		*pEffect = DROPEFFECT_NONE;
		return S_OK;
	}*/
}

STDMETHODIMP RichTextBox::GetContextMenu(WORD selectionType, LPOLEOBJECT pObject, CHARRANGE FAR* pCharRange, HMENU FAR* pHMenu)
{
	if(!pCharRange || !pHMenu) {
		return E_INVALIDARG;
	}
	GETCONTEXTMENUEX* pContextMenuEx = NULL;
	if(RunTimeHelperEx::IsWin8()) {
		pContextMenuEx = reinterpret_cast<GETCONTEXTMENUEX*>(pCharRange);
	}

	MenuTypeConstants menuType = static_cast<MenuTypeConstants>(selectionType & 0xFFFFFF00);
	if(pContextMenuEx) {
		menuType = static_cast<MenuTypeConstants>(pContextMenuEx->dwFlags);
	}
	if(menuType == static_cast<MenuTypeConstants>(0)) {
		menuType = mtMouseMenu;
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	BOOL triggeredByKeyboard = FALSE;
	if(shift & 1/*ShiftConstants.vbShiftMask*/) {
		triggeredByKeyboard = (GetAsyncKeyState(VK_F10) & 0x8000);
	} else {
		triggeredByKeyboard = (GetAsyncKeyState(VK_APPS) & 0x8000);
	}
	if(triggeredByKeyboard) {
		if(!properties.processContextMenuKeys) {
			return S_OK;
		}
	}

	POINT menuPosition = {0};
	if(pContextMenuEx) {
		menuPosition = pContextMenuEx->pt;
		ScreenToClient(&menuPosition);
	} else if(triggeredByKeyboard) {
		if(pCharRange->cpMin == 0 && pCharRange->cpMax == -1) {
			// everything is selected
			WTL::CRect clientRectangle;
			GetClientRect(&clientRectangle);
			WTL::CPoint centerPoint = clientRectangle.CenterPoint();
			menuPosition.x = centerPoint.x;
			menuPosition.y = centerPoint.y;
		} else {
			SendMessage(EM_POSFROMCHAR, reinterpret_cast<WPARAM>(&menuPosition), pCharRange->cpMax);
		}
	} else {
		DWORD position = GetMessagePos();
		menuPosition.x = GET_X_LPARAM(position);
		menuPosition.y = GET_Y_LPARAM(position);
		ScreenToClient(&menuPosition);
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IRichOLEObject> pOLEObject = NULL;
	CComPtr<ITextRange> pObjTextRange = NULL;
	if(selectionType & GCM_RIGHTMOUSEDROP) {
		CComQIPtr<IDataObject> pDataObject = pObject;
		pOLEDataObject = ClassFactory::InitOLEDataObject(pDataObject);
	} else if(pObject) {
		if(SUCCEEDED(properties.pTextDocument->Range(pCharRange->cpMin, pCharRange->cpMax, &pObjTextRange)) && pObjTextRange) {
			VARIANT v;
			VariantInit(&v);
			v.vt = VT_I4;
			v.lVal = WCH_EMBEDDING;
			pObjTextRange->MoveStartUntil(&v, pCharRange->cpMax - pCharRange->cpMin, NULL);
			LONG start = 0;
			pObjTextRange->GetStart(&start);
			pObjTextRange->SetEnd(start);
			pObjTextRange->EndOf(tomObject, tomExtend, NULL);
		}
		pOLEObject = ClassFactory::InitOLEObject(pObjTextRange, pObject, this);
	}

	CComPtr<IRichTextRange> pTextRange = NULL;
	get_TextRange(pCharRange->cpMin, pCharRange->cpMax, &pTextRange);
	//*pHMenu = reinterpret_cast<HMENU>(-1);
	Raise_ContextMenu(menuType, pTextRange, static_cast<SelectionTypeConstants>(selectionType & 0xFF), pOLEObject, button, shift, menuPosition.x, menuPosition.y, BOOL2VARIANTBOOL(selectionType & GCM_RIGHTMOUSEDROP), pOLEDataObject, pHMenu);
	/*if(*pHMenu == reinterpret_cast<HMENU>(-1)) {
		// show default menu
		*pHMenu = NULL;
	} else if(*pHMenu == NULL) {
		// don't show default menu
		return S_FALSE;
	}*/
	return S_OK;
}
// implementation of IRichEditOleCallback
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP RichTextBox::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Accessibility:
			*pName = GetResString(IDPC_ACCESSIBILITY).Detach();
			return S_OK;
			break;
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
		case PROPCAT_IMEAndTSF:
			*pName = GetResString(IDPC_IMEANDTSF).Detach();
			return S_OK;
			break;
		case PROPCAT_MathZones:
			*pName = GetResString(IDPC_MATHZONES).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_RTB_ALWAYSSHOWSELECTION:
		case DISPID_RTB_APPEARANCE:
		case DISPID_RTB_BORDERSTYLE:
		case DISPID_RTB_DISPLAYZEROWIDTHTABLECELLBORDERS:
		case DISPID_RTB_DOCUMENTFONT:
		case DISPID_RTB_DOCUMENTPARAGRAPH:
		case DISPID_RTB_EXTENDFONTBACKCOLORTOWHOLELINE:
		case DISPID_RTB_FIRSTVISIBLELINE:
		case DISPID_RTB_FORMATTINGRECTANGLEHEIGHT:
		case DISPID_RTB_FORMATTINGRECTANGLELEFT:
		case DISPID_RTB_FORMATTINGRECTANGLETOP:
		case DISPID_RTB_FORMATTINGRECTANGLEWIDTH:
		//case DISPID_RTB_HALIGNMENT:
		case DISPID_RTB_LEFTMARGIN:
		case DISPID_RTB_LINKMOUSEICON:
		case DISPID_RTB_LINKMOUSEPOINTER:
		case DISPID_RTB_MOUSEICON:
		case DISPID_RTB_MOUSEPOINTER:
		case DISPID_RTB_RIGHTMARGIN:
		case DISPID_RTB_SHOWSELECTIONBAR:
		case DISPID_RTB_TEXTFLOW:
		case DISPID_RTB_USEBKACETATECOLORFORTEXTSELECTION:
		case DISPID_RTB_USECUSTOMFORMATTINGRECTANGLE:
		case DISPID_RTB_USEWINDOWSTHEMES:
		case DISPID_RTB_ZOOMRATIO:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_RTB_ACCEPTTABKEY:
		case DISPID_RTB_ALLOWTABLEINSERTION:
		case DISPID_RTB_ALWAYSSHOWSCROLLBARS:
		case DISPID_RTB_AUTODETECTEDURLSCHEMES:
		case DISPID_RTB_AUTODETECTURLS:
		case DISPID_RTB_AUTOSCROLLING:
		case DISPID_RTB_AUTOSELECTWORDONTRACKSELECTION:
		case DISPID_RTB_BEEPONMAXTEXT:
		case DISPID_RTB_CHARACTERCONVERSION:
		case DISPID_RTB_DISABLEDEVENTS:
		case DISPID_RTB_DISPLAYHYPERLINKTOOLTIPS:
		case DISPID_RTB_DRAFTMODE:
		case DISPID_RTB_EMULATESIMPLETEXTBOX:
		case DISPID_RTB_HOVERTIME:
		case DISPID_RTB_HYPHENATIONFUNCTION:
		case DISPID_RTB_HYPHENATIONWORDWRAPZONEWIDTH:
		case DISPID_RTB_LOGICALCARET:
		case DISPID_RTB_MAXTEXTLENGTH:
		case DISPID_RTB_MAXUNDOQUEUESIZE:
		case DISPID_RTB_MULTILINE:
		case DISPID_RTB_MULTISELECT:
		case DISPID_RTB_NOINPUTSEQUENCECHECK:
		case DISPID_RTB_PROCESSCONTEXTMENUKEYS:
		case DISPID_RTB_RIGHTTOLEFT:
		case DISPID_RTB_SCROLLBARS:
		case DISPID_RTB_SCROLLTOTOPONKILLFOCUS:
		case DISPID_RTB_USEBUILTINSPELLCHECKING:
		//case DISPID_RTB_USETOUCHKEYBOARDAUTOCORRECTION:
		case DISPID_RTB_WORDBREAKFUNCTION:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_RTB_EFFECTCOLOR:
		case DISPID_RTB_BACKCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_RTB_APPID:
		case DISPID_RTB_APPNAME:
		case DISPID_RTB_APPSHORTNAME:
		case DISPID_RTB_BUILD:
		case DISPID_RTB_CHARSET:
		case DISPID_RTB_EMBEDDEDOBJECTS:
		case DISPID_RTB_HDRAGIMAGELIST:
		case DISPID_RTB_HWND:
		case DISPID_RTB_ISRELEASE:
		case DISPID_RTB_MODIFIED:
		case DISPID_RTB_NEXTREDOACTIONTYPE:
		case DISPID_RTB_NEXTUNDOACTIONTYPE:
		case DISPID_RTB_PROGRAMMER:
		case DISPID_RTB_RICHTEXT:
		//case DISPID_RTB_SELECTEDTEXT:
		case DISPID_RTB_SELECTEDTEXTRANGE:
		case DISPID_RTB_SELECTIONTYPE:
		case DISPID_RTB_TESTER:
		case DISPID_RTB_TEXT:
		//case DISPID_RTB_TEXTLENGTH:
		case DISPID_RTB_TEXTRANGE:
		case DISPID_RTB_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_RTB_DETECTDRAGDROP:
		case DISPID_RTB_DRAGGEDTEXTRANGE:
		case DISPID_RTB_DRAGSCROLLTIMEBASE:
		case DISPID_RTB_DROPWORDSONWORDBOUNDARIESONLY:
		case DISPID_RTB_OLEDRAGIMAGESTYLE:
		case DISPID_RTB_REGISTERFOROLEDRAGDROP:
		case DISPID_RTB_SHOWDRAGIMAGE:
		case DISPID_RTB_SMARTSPACINGONDROP:
		case DISPID_RTB_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_RTB_DEFAULTTABWIDTH:
		case DISPID_RTB_FONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_RTB_ADJUSTLINEHEIGHTFOREASTASIANLANGUAGES:
		case DISPID_RTB_ALLOWINPUTTHROUGHTSF:
		case DISPID_RTB_ALLOWOBJECTINSERTIONTHROUGHTSF:
		case DISPID_RTB_ALLOWTSFPROOFINGTIPS:
		case DISPID_RTB_ALLOWTSFSMARTTAGS:
		case DISPID_RTB_CURRENTIMECOMPOSITIONMODE:
		case DISPID_RTB_DISABLEIMEOPERATIONS:
		//case DISPID_RTB_DISCARDCOMPOSITIONSTRINGONIMECANCEL:
		case DISPID_RTB_IMECONVERSIONMODE:
		case DISPID_RTB_IMEMODE:
		case DISPID_RTB_LETCLIENTHANDLEALLIMEOPERATIONS:
		case DISPID_RTB_SWITCHFONTONIMEINPUT:
		case DISPID_RTB_TEXTORIENTATION:
		case DISPID_RTB_TSFMODEBIAS:
		case DISPID_RTB_USEDUALFONTMODE:
		case DISPID_RTB_USETEXTSERVICESFRAMEWORK:
			*pCategory = PROPCAT_IMEAndTSF;
			return S_OK;
			break;
		case DISPID_RTB_ALLOWMATHZONEINSERTION:
		case DISPID_RTB_DEFAULTMATHZONEHALIGNMENT:
		case DISPID_RTB_DENOTEEMPTYMATHARGUMENTS:
		case DISPID_RTB_GROWNARYOPERATORS:
		case DISPID_RTB_INTEGRALLIMITSLOCATION:
		case DISPID_RTB_MATHLINEBREAKBEHAVIOR:
		case DISPID_RTB_NARYLIMITSLOCATION:
		case DISPID_RTB_RAWSUBSCRIPTANDSUPERSCRIPTOPERATORS:
		case DISPID_RTB_RIGHTTOLEFTMATHZONES:
		case DISPID_RTB_USESMALLERFONTFORNESTEDFRACTIONS:
			*pCategory = PROPCAT_MathZones;
			return S_OK;
			break;
		case DISPID_RTB_ENABLED:
		case DISPID_RTB_READONLY:
		case DISPID_RTB_RICHEDITAPIVERSION:
		case DISPID_RTB_RICHEDITLIBRARYPATH:
		case DISPID_RTB_RICHEDITVERSION:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString RichTextBox::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString RichTextBox::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString RichTextBox::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=RichTextBox&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString RichTextBox::GetSpecialThanks(void)
{
	return TEXT("Wine Headquarters");
}

CAtlString RichTextBox::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString RichTextBox::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IMouseHookHandler
LRESULT RichTextBox::HandleMouseMessage(int /*code*/, WPARAM wParam, LPARAM lParam, BOOL& /*consumeMessage*/)
{
	if(wParam == WM_MOUSEMOVE) {
		LPMOUSEHOOKSTRUCT pDetails = reinterpret_cast<LPMOUSEHOOKSTRUCT>(lParam);
		ATLASSERT_POINTER(pDetails, MOUSEHOOKSTRUCT);

		HWND hWnd = WindowFromPoint(pDetails->pt);
		TCHAR pClassName[201];
		::GetClassName(hWnd, pClassName, 250);
		if(lstrcmpi(pClassName, TEXT("#32768")) == 0) {
			// assume it is the spell checking context menu
			HCURSOR hCursor = static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
			SetCursor(hCursor);
			return TRUE;
		}
	}

	return FALSE;
}
// implementation of IMouseHookHandler
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP RichTextBox::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_RTB_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_RTB_AUTOSCROLLING:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.autoScrolling), description);
			break;
		case DISPID_RTB_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_RTB_CHARACTERCONVERSION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.characterConversion), description);
			break;
		case DISPID_RTB_DEFAULTMATHZONEHALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.defaultMathZoneHAlignment), description);
			break;
		case DISPID_RTB_DENOTEEMPTYMATHARGUMENTS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.denoteEmptyMathArguments), description);
			break;
		//case DISPID_RTB_HALIGNMENT:
		//	hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hAlignment), description);
		//	break;
		case DISPID_RTB_IMECONVERSIONMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEConversionMode), description);
			break;
		case DISPID_RTB_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_RTB_INTEGRALLIMITSLOCATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.integralLimitsLocation), description);
			break;
		case DISPID_RTB_LINKMOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.linkMousePointer), description);
			break;
		case DISPID_RTB_MATHLINEBREAKBEHAVIOR:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mathLineBreakBehavior), description);
			break;
		case DISPID_RTB_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_RTB_OLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.oleDragImageStyle), description);
			break;
		case DISPID_RTB_NARYLIMITSLOCATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.nAryLimitsLocation), description);
			break;
		case DISPID_RTB_REGISTERFOROLEDRAGDROP:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.registerForOLEDragDrop), description);
			break;
		case DISPID_RTB_RICHEDITVERSION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.richEditVersion), description);
			break;
		case DISPID_RTB_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_RTB_SCROLLBARS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.scrollBars), description);
			break;
		case DISPID_RTB_AUTODETECTEDURLSCHEMES:
		case DISPID_RTB_TEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_RTB_TEXTFLOW:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.textFlow), description);
			break;
		case DISPID_RTB_TEXTORIENTATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.textOrientation), description);
			break;
		case DISPID_RTB_TSFMODEBIAS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.TSFModeBias), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<RichTextBox>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP RichTextBox::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_RTB_BORDERSTYLE:
		case DISPID_RTB_INTEGRALLIMITSLOCATION:
		case DISPID_RTB_NARYLIMITSLOCATION:
		case DISPID_RTB_OLEDRAGIMAGESTYLE:
			c = 2;
			break;
		case DISPID_RTB_APPEARANCE:
		case DISPID_RTB_CHARACTERCONVERSION:
		case DISPID_RTB_DEFAULTMATHZONEHALIGNMENT:
		case DISPID_RTB_DENOTEEMPTYMATHARGUMENTS:
		//case DISPID_RTB_HALIGNMENT:
		case DISPID_RTB_IMECONVERSIONMODE:
		case DISPID_RTB_REGISTERFOROLEDRAGDROP:
		case DISPID_RTB_RICHEDITVERSION:
		case DISPID_RTB_TEXTORIENTATION:
			c = 3;
			break;
		case DISPID_RTB_AUTOSCROLLING:
		case DISPID_RTB_RIGHTTOLEFT:
		case DISPID_RTB_SCROLLBARS:
			c = 4;
			break;
		case DISPID_RTB_MATHLINEBREAKBEHAVIOR:
		case DISPID_RTB_TEXTFLOW:
			c = 5;
			break;
		case DISPID_RTB_IMEMODE:
			c = 12;
			break;
		case DISPID_RTB_TSFMODEBIAS:
			c = 13;
			break;
		case DISPID_RTB_LINKMOUSEPOINTER:
		case DISPID_RTB_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<RichTextBox>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if(((property == DISPID_RTB_MOUSEPOINTER) || (property == DISPID_RTB_LINKMOUSEPOINTER)) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		} else if(property == DISPID_RTB_TEXTFLOW && iDescription == pDescriptions->cElems - 1) {
			propertyValue = tfTopToBottomLeftToRight;
		} else if(property == DISPID_RTB_DENOTEEMPTYMATHARGUMENTS) {
			switch(iDescription) {
				case 1:
					propertyValue = demaAlways;
					break;
				case 2:
					propertyValue = demaNever;
					break;
			}
		} else if(property == DISPID_RTB_IMEMODE) {
			// the enum is -1-based
			--propertyValue;
		} else if(property == DISPID_RTB_MATHLINEBREAKBEHAVIOR) {
			switch(iDescription) {
				case 0:
					propertyValue = mlbbBreakBeforeOperator;
					break;
				case 1:
					propertyValue = mlbbBreakAfterOperator;
					break;
				case 2:
					propertyValue = mlbbDuplicateOperatorMinusMinus;
					break;
				case 3:
					propertyValue = mlbbDuplicateOperatorPlusMinus;
					break;
				case 4:
					propertyValue = mlbbDuplicateOperatorMinusPlus;
					break;
			}
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_RTB_APPEARANCE:
		case DISPID_RTB_AUTODETECTEDURLSCHEMES:
		case DISPID_RTB_AUTOSCROLLING:
		case DISPID_RTB_BORDERSTYLE:
		case DISPID_RTB_CHARACTERCONVERSION:
		case DISPID_RTB_DEFAULTMATHZONEHALIGNMENT:
		case DISPID_RTB_DENOTEEMPTYMATHARGUMENTS:
		//case DISPID_RTB_HALIGNMENT:
		case DISPID_RTB_IMECONVERSIONMODE:
		case DISPID_RTB_IMEMODE:
		case DISPID_RTB_INTEGRALLIMITSLOCATION:
		case DISPID_RTB_LINKMOUSEPOINTER:
		case DISPID_RTB_MATHLINEBREAKBEHAVIOR:
		case DISPID_RTB_MOUSEPOINTER:
		case DISPID_RTB_NARYLIMITSLOCATION:
		case DISPID_RTB_OLEDRAGIMAGESTYLE:
		case DISPID_RTB_REGISTERFOROLEDRAGDROP:
		case DISPID_RTB_RICHEDITVERSION:
		case DISPID_RTB_RIGHTTOLEFT:
		case DISPID_RTB_SCROLLBARS:
		case DISPID_RTB_TEXTFLOW:
		case DISPID_RTB_TEXTORIENTATION:
		case DISPID_RTB_TSFMODEBIAS:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<RichTextBox>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_RTB_AUTODETECTEDURLSCHEMES:
		case DISPID_RTB_TEXT:
			/*TODO: *pPropertyPage = CLSID_StringProperties;
			return S_OK;*/
			break;
	}
	return IPerPropertyBrowsingImpl<RichTextBox>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT RichTextBox::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_RTB_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_AUTOSCROLLING:
			switch(cookie) {
				case asNone:
					description = GetResStringWithNumber(IDP_AUTOSCROLLINGNONE, asNone);
					break;
				case asVertical:
					description = GetResStringWithNumber(IDP_AUTOSCROLLINGVERTICAL, asVertical);
					break;
				case asHorizontal:
					description = GetResStringWithNumber(IDP_AUTOSCROLLINGHORIZONTAL, asHorizontal);
					break;
				case asVertical | asHorizontal:
					description = GetResStringWithNumber(IDP_AUTOSCROLLINGBOTH, asVertical | asHorizontal);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_CHARACTERCONVERSION:
			switch(cookie) {
				case ccNone:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONNONE, ccNone);
					break;
				case ccLowerCase:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONLOWERCASE, ccLowerCase);
					break;
				case ccUpperCase:
					description = GetResStringWithNumber(IDP_CHARACTERCONVERSIONUPPERCASE, ccUpperCase);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_DEFAULTMATHZONEHALIGNMENT:
		//case DISPID_RTB_HALIGNMENT:
			switch(cookie) {
				case halLeft:
					description = GetResStringWithNumber(IDP_HALIGNMENTLEFT, halLeft);
					break;
				case halCenter:
					description = GetResStringWithNumber(IDP_HALIGNMENTCENTER, halCenter);
					break;
				case halRight:
					description = GetResStringWithNumber(IDP_HALIGNMENTRIGHT, halRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_DENOTEEMPTYMATHARGUMENTS:
			switch(cookie) {
				case demaMandatoryOnly:
					description = GetResStringWithNumber(IDP_DENOTEEMPTYMATHARGUMENTSMANDATORYONLY, demaMandatoryOnly);
					break;
				case demaAlways:
					description = GetResStringWithNumber(IDP_DENOTEEMPTYMATHARGUMENTSALWAYS, demaAlways);
					break;
				case demaNever:
					description = GetResStringWithNumber(IDP_DENOTEEMPTYMATHARGUMENTSNEVER, demaNever);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_IMECONVERSIONMODE:
			switch(cookie) {
				case imecmDefault:
					description = GetResStringWithNumber(IDP_IMECONVERSIONMODEDEFAULT, imecmDefault);
					break;
				case imecmBiasForNames:
					description = GetResStringWithNumber(IDP_IMECONVERSIONMODEBIASFORNAMES, imecmBiasForNames);
					break;
				case imecmNoConversion:
					description = GetResStringWithNumber(IDP_IMECONVERSIONMODENOCONVERSION, imecmNoConversion);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_IMEMODE:
			switch(cookie) {
				case imeInherit:
					description = GetResStringWithNumber(IDP_IMEMODEINHERIT, imeInherit);
					break;
				case imeNoControl:
					description = GetResStringWithNumber(IDP_IMEMODENOCONTROL, imeNoControl);
					break;
				case imeOn:
					description = GetResStringWithNumber(IDP_IMEMODEON, imeOn);
					break;
				case imeOff:
					description = GetResStringWithNumber(IDP_IMEMODEOFF, imeOff);
					break;
				case imeDisable:
					description = GetResStringWithNumber(IDP_IMEMODEDISABLE, imeDisable);
					break;
				case imeHiragana:
					description = GetResStringWithNumber(IDP_IMEMODEHIRAGANA, imeHiragana);
					break;
				case imeKatakana:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANA, imeKatakana);
					break;
				case imeKatakanaHalf:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANAHALF, imeKatakanaHalf);
					break;
				case imeAlphaFull:
					description = GetResStringWithNumber(IDP_IMEMODEALPHAFULL, imeAlphaFull);
					break;
				case imeAlpha:
					description = GetResStringWithNumber(IDP_IMEMODEALPHA, imeAlpha);
					break;
				case imeHangulFull:
					description = GetResStringWithNumber(IDP_IMEMODEHANGULFULL, imeHangulFull);
					break;
				case imeHangul:
					description = GetResStringWithNumber(IDP_IMEMODEHANGUL, imeHangul);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_INTEGRALLIMITSLOCATION:
		case DISPID_RTB_NARYLIMITSLOCATION:
			switch(cookie) {
				case llOnSide:
					description = GetResStringWithNumber(IDP_LIMITSLOCATIONONSIDE, llOnSide);
					break;
				case llUnderAndOver:
					description = GetResStringWithNumber(IDP_LIMITSLOCATIONUNDERANDOVER, llUnderAndOver);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_LINKMOUSEPOINTER:
		case DISPID_RTB_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_MATHLINEBREAKBEHAVIOR:
			switch(cookie) {
				case mlbbBreakBeforeOperator:
					description = GetResStringWithNumber(IDP_MATHLINEBREAKBEHAVIORBEFORE, mlbbBreakBeforeOperator);
					break;
				case mlbbBreakAfterOperator:
					description = GetResStringWithNumber(IDP_MATHLINEBREAKBEHAVIORAFTER, mlbbBreakAfterOperator);
					break;
				case mlbbDuplicateOperatorMinusMinus:
					description = GetResStringWithNumber(IDP_MATHLINEBREAKBEHAVIORDUPMINUSMINUS, mlbbDuplicateOperatorMinusMinus);
					break;
				case mlbbDuplicateOperatorPlusMinus:
					description = GetResStringWithNumber(IDP_MATHLINEBREAKBEHAVIORDUPPLUSMINUS, mlbbDuplicateOperatorPlusMinus);
					break;
				case mlbbDuplicateOperatorMinusPlus:
					description = GetResStringWithNumber(IDP_MATHLINEBREAKBEHAVIORDUPMINUSPLUS, mlbbDuplicateOperatorMinusPlus);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_OLEDRAGIMAGESTYLE:
			switch(cookie) {
				case odistClassic:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLECLASSIC, odistClassic);
					break;
				case odistAeroIfAvailable:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLEAERO, odistAeroIfAvailable);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_REGISTERFOROLEDRAGDROP:
			switch(cookie) {
				case rfoddNoDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPNONE, rfoddNoDragDrop);
					break;
				case rfoddNativeDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPNATIVE, rfoddNativeDragDrop);
					break;
				case rfoddAdvancedDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPADVANCED, rfoddAdvancedDragDrop);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_RICHEDITVERSION:
			switch(cookie) {
				case revAlwaysUseWindowsRichEdit:
					description = GetResStringWithNumber(IDP_RICHEDITVERSIONALWAYSWINDOWS, revAlwaysUseWindowsRichEdit);
					break;
				case revPreferOfficeRichEdit:
					description = GetResStringWithNumber(IDP_RICHEDITVERSIONOFFICEIFAVAILABLE, revPreferOfficeRichEdit);
					break;
				case revUseHighestVersion:
					description = GetResStringWithNumber(IDP_RICHEDITVERSIONOFFICEIFNEWER, revUseHighestVersion);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_SCROLLBARS:
			switch(cookie) {
				case sbNone:
					description = GetResStringWithNumber(IDP_SCROLLBARSNONE, sbNone);
					break;
				case sbVertical:
					description = GetResStringWithNumber(IDP_SCROLLBARSVERTICAL, sbVertical);
					break;
				case sbHorizontal:
					description = GetResStringWithNumber(IDP_SCROLLBARSHORIZONTAL, sbHorizontal);
					break;
				case sbVertical | sbHorizontal:
					description = GetResStringWithNumber(IDP_SCROLLBARSBOTH, sbVertical | sbHorizontal);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_TEXTFLOW:
			switch(cookie) {
				case tfLeftToRightTopToBottom:
					description = GetResStringWithNumber(IDP_TEXTFLOWLTRTTB, tfLeftToRightTopToBottom);
					break;
				case tfTopToBottomRightToLeft:
					description = GetResStringWithNumber(IDP_TEXTFLOWTTBRTL, tfTopToBottomRightToLeft);
					break;
				case tfRightToLeftBottomToTop:
					description = GetResStringWithNumber(IDP_TEXTFLOWRTLBTT, tfRightToLeftBottomToTop);
					break;
				case tfBottomToTopLeftToRight:
					description = GetResStringWithNumber(IDP_TEXTFLOWBTTLTR, tfBottomToTopLeftToRight);
					break;
				case tfTopToBottomLeftToRight:
					description = GetResStringWithNumber(IDP_TEXTFLOWTTBLTR, tfTopToBottomLeftToRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_TEXTORIENTATION:
			switch(cookie) {
				case toHorizontal:
					description = GetResStringWithNumber(IDP_TEXTORIENTATIONHORIZONTAL, toHorizontal);
					break;
				case toVertical:
					description = GetResStringWithNumber(IDP_TEXTORIENTATIONVERTICAL, toVertical);
					break;
				case toVerticalWithHorizontalFont:
					description = GetResStringWithNumber(IDP_TEXTORIENTATIONVERTICALWITHHORZFONT, toVerticalWithHorizontalFont);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RTB_TSFMODEBIAS:
			switch(cookie) {
				case tsfmbDefault:
					description = GetResStringWithNumber(IDP_TSFMODEBIASDEFAULT, tsfmbDefault);
					break;
				case tsfmbBiasToFilename:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOFILENAME, tsfmbBiasToFilename);
					break;
				case tsfmbBiasToName:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTONAME, tsfmbBiasToName);
					break;
				case tsfmbBiasToReading:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOREADING, tsfmbBiasToReading);
					break;
				case tsfmbBiasToDateOrTime:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTODATEORTIME, tsfmbBiasToDateOrTime);
					break;
				case tsfmbBiasToConversation:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOCONVERSATION, tsfmbBiasToConversation);
					break;
				case tsfmbBiasToNumber:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTONUMBER, tsfmbBiasToNumber);
					break;
				case tsfmbBiasToHiraganaStrings:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOHIRAGANA, tsfmbBiasToHiraganaStrings);
					break;
				case tsfmbBiasToKatakanaStrings:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOKATAKANA, tsfmbBiasToKatakanaStrings);
					break;
				case tsfmbBiasToHangulCharacters:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOHANGUL, tsfmbBiasToHangulCharacters);
					break;
				case tsfmbBiasToHalfWidthKatakanaStrings:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOHALFWIDTHKATAKANA, tsfmbBiasToHalfWidthKatakanaStrings);
					break;
				case tsfmbBiasToFullWidthAlphanumericCharacters:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOFULLWIDTHALPHA, tsfmbBiasToFullWidthAlphanumericCharacters);
					break;
				case tsfmbBiasToHalfWidthAlphanumericCharacters:
					description = GetResStringWithNumber(IDP_TSFMODEBIASTOHALFWIDTHALPHA, tsfmbBiasToHalfWidthAlphanumericCharacters);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP RichTextBox::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 3;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		//TODO: pPropertyPages->pElems[0] = CLSID_StringProperties;
		pPropertyPages->pElems[0] = CLSID_StockColorPage;
		pPropertyPages->pElems[1] = CLSID_StockFontPage;
		pPropertyPages->pElems[2] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP RichTextBox::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<RichTextBox>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP RichTextBox::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP RichTextBox::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT RichTextBox::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT RichTextBox::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_RTB_FONT:
			ApplyFont();
			break;
	}
	return S_OK;
}

HRESULT RichTextBox::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP RichTextBox::get_AcceptTabKey(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.acceptTabKey);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AcceptTabKey(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ACCEPTTABKEY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.acceptTabKey != b) {
		properties.acceptTabKey = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_ACCEPTTABKEY);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AdjustLineHeightForEastAsianLanguages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.adjustLineHeightForEastAsianLanguages = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_NOEALINEHEIGHTADJUST) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.adjustLineHeightForEastAsianLanguages);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AdjustLineHeightForEastAsianLanguages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ADJUSTLINEHEIGHTFOREASTASIANLANGUAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.adjustLineHeightForEastAsianLanguages != b) {
		properties.adjustLineHeightForEastAsianLanguages = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.adjustLineHeightForEastAsianLanguages) {
				SendMessage(EM_SETEDITSTYLE, 0, SES_NOEALINEHEIGHTADJUST);
			} else {
				SendMessage(EM_SETEDITSTYLE, SES_NOEALINEHEIGHTADJUST, SES_NOEALINEHEIGHTADJUST);
			}
		}
		FireOnChanged(DISPID_RTB_ADJUSTLINEHEIGHTFOREASTASIANLANGUAGES);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AllowInputThroughTSF(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowInputThroughTSF = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_CTFNOLOCK) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowInputThroughTSF);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AllowInputThroughTSF(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ALLOWINPUTTHROUGHTSF);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowInputThroughTSF != b) {
		properties.allowInputThroughTSF = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowInputThroughTSF) {
				SendMessage(EM_SETEDITSTYLE, 0, SES_CTFNOLOCK);
			} else {
				SendMessage(EM_SETEDITSTYLE, SES_CTFNOLOCK, SES_CTFNOLOCK);
			}
		}
		FireOnChanged(DISPID_RTB_ALLOWINPUTTHROUGHTSF);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AllowMathZoneInsertion(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowMathZoneInsertion = ((SendMessage(EM_GETEDITSTYLEEX, 0, 0) & SES_EX_NOMATH) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowMathZoneInsertion);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AllowMathZoneInsertion(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ALLOWMATHZONEINSERTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowMathZoneInsertion != b) {
		properties.allowMathZoneInsertion = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowMathZoneInsertion) {
				SendMessage(EM_SETEDITSTYLEEX, 0, SES_EX_NOMATH);
			} else {
				SendMessage(EM_SETEDITSTYLEEX, SES_EX_NOMATH, SES_EX_NOMATH);
			}
		}
		FireOnChanged(DISPID_RTB_ALLOWMATHZONEINSERTION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AllowObjectInsertionThroughTSF(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowObjectInsertionThroughTSF = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_CTFALLOWEMBED) == SES_CTFALLOWEMBED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowObjectInsertionThroughTSF);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AllowObjectInsertionThroughTSF(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ALLOWOBJECTINSERTIONTHROUGHTSF);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowObjectInsertionThroughTSF != b) {
		properties.allowObjectInsertionThroughTSF = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowObjectInsertionThroughTSF) {
				SendMessage(EM_SETEDITSTYLE, SES_CTFALLOWEMBED, SES_CTFALLOWEMBED);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_CTFALLOWEMBED);
			}
		}
		FireOnChanged(DISPID_RTB_ALLOWOBJECTINSERTIONTHROUGHTSF);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AllowTableInsertion(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowTableInsertion = ((SendMessage(EM_GETEDITSTYLEEX, 0, 0) & SES_EX_NOTABLE) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowTableInsertion);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AllowTableInsertion(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ALLOWTABLEINSERTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowTableInsertion != b) {
		properties.allowTableInsertion = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowTableInsertion) {
				SendMessage(EM_SETEDITSTYLEEX, 0, SES_EX_NOTABLE);
			} else {
				SendMessage(EM_SETEDITSTYLEEX, SES_EX_NOTABLE, SES_EX_NOTABLE);
			}
		}
		FireOnChanged(DISPID_RTB_ALLOWTABLEINSERTION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AllowTSFProofingTips(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowTSFProofingTips = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_CTFALLOWPROOFING) == SES_CTFALLOWPROOFING);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowTSFProofingTips);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AllowTSFProofingTips(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ALLOWTSFPROOFINGTIPS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowTSFProofingTips != b) {
		properties.allowTSFProofingTips = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowTSFProofingTips) {
				SendMessage(EM_SETEDITSTYLE, SES_CTFALLOWPROOFING, SES_CTFALLOWPROOFING);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_CTFALLOWPROOFING);
			}
		}
		FireOnChanged(DISPID_RTB_ALLOWTSFPROOFINGTIPS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AllowTSFSmartTags(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowTSFSmartTags = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_CTFALLOWSMARTTAG) == SES_CTFALLOWSMARTTAG);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowTSFSmartTags);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AllowTSFSmartTags(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ALLOWTSFSMARTTAGS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowTSFSmartTags != b) {
		properties.allowTSFSmartTags = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowTSFSmartTags) {
				SendMessage(EM_SETEDITSTYLE, SES_CTFALLOWSMARTTAG, SES_CTFALLOWSMARTTAG);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_CTFALLOWSMARTTAG);
			}
		}
		FireOnChanged(DISPID_RTB_ALLOWTSFSMARTTAGS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AlwaysShowScrollBars(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.alwaysShowScrollBars = ((GetStyle() & ES_DISABLENOSCROLL) == ES_DISABLENOSCROLL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.alwaysShowScrollBars);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AlwaysShowScrollBars(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ALWAYSSHOWSCROLLBARS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.alwaysShowScrollBars != b) {
		properties.alwaysShowScrollBars = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_RTB_ALWAYSSHOWSCROLLBARS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AlwaysShowSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.alwaysShowSelection = ((SendMessage(EM_GETOPTIONS, 0, 0) & ECO_NOHIDESEL) == ECO_NOHIDESEL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AlwaysShowSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ALWAYSSHOWSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.alwaysShowSelection != b) {
		properties.alwaysShowSelection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			// This would work as well, but then we cannot retrieve the current setting anymore:
			/*CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				CComPtr<ITextServices> pTextServices = NULL;
				pRichEditOle->QueryInterface(IID_ITextServices, reinterpret_cast<LPVOID*>(&pTextServices));
				if(pTextServices) {
					pTextServices->OnTxPropertyBitsChange(TXTBIT_HIDESELECTION, (properties.alwaysShowSelection ? 0 : TXTBIT_HIDESELECTION));
				}
			}*/
			if(properties.alwaysShowSelection) {
				SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_NOHIDESEL);
			} else {
				SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_NOHIDESEL);
			}
		}
		FireOnChanged(DISPID_RTB_ALWAYSSHOWSELECTION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_APPEARANCE);
	if(newValue < 0 || newValue > a3DLight) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 19;
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"RichTextBox");
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"RichTextBoxCtl");
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AutoDetectedURLSchemes(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoDetectedURLSchemes.Copy();
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AutoDetectedURLSchemes(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_RTB_AUTODETECTEDURLSCHEMES);
	properties.autoDetectedURLSchemes.AssignBSTR(newValue);
	if(IsWindow()) {
		DWORD autoDetectURLs = properties.autoDetectURLs;
		if(properties.autoDetectedURLSchemes.Length() > 0) {
			if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, reinterpret_cast<LPARAM>(OLE2W(properties.autoDetectedURLSchemes))) != 0) {
				// failure - on Windows 7 and older, try without URL schemes
				if(!RunTimeHelperEx::IsWin8() && SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
					// success
					if(!IsInDesignMode()) {
						properties.autoDetectedURLSchemes = L"";
						properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
					}
				} else {
					// still a failure - ensure that only valid flags are used
					if(RunTimeHelperEx::IsWin8()) {
						autoDetectURLs &= (AURL_ENABLEURL | AURL_ENABLEEMAILADDR | AURL_ENABLETELNO | AURL_ENABLEEAURLS | AURL_ENABLEDRIVELETTERS | AURL_DISABLEMIXEDLGC);
						if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, reinterpret_cast<LPARAM>(OLE2W(properties.autoDetectedURLSchemes))) == 0) {
							// success
							if(!IsInDesignMode()) {
								properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
							}
						}
					} else {
						autoDetectURLs &= AURL_ENABLEURL;
						if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
							// success
							if(!IsInDesignMode()) {
								properties.autoDetectedURLSchemes = L"";
								properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
							}
						}
					}
				}
			}
		} else {
			if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) != 0) {
				// failure - ensure that only valid flags are used
				if(RunTimeHelperEx::IsWin8()) {
					autoDetectURLs = (autoDetectURLs & (AURL_ENABLEURL | AURL_ENABLEEMAILADDR | AURL_ENABLETELNO | AURL_ENABLEEAURLS | AURL_ENABLEDRIVELETTERS | AURL_DISABLEMIXEDLGC));
				} else {
					autoDetectURLs = (autoDetectURLs & AURL_ENABLEURL);
				}
				if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
					// success
					if(!IsInDesignMode()) {
						properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
					}
				}
			}
		}
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_AUTODETECTEDURLSCHEMES);
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AutoDetectURLs(AutoDetectURLsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AutoDetectURLsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoDetectURLs;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AutoDetectURLs(AutoDetectURLsConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_AUTODETECTURLS);
	if(properties.autoDetectURLs != newValue) {
		properties.autoDetectURLs = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD autoDetectURLs = properties.autoDetectURLs;
			if(properties.autoDetectedURLSchemes.Length() > 0) {
				if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, reinterpret_cast<LPARAM>(OLE2W(properties.autoDetectedURLSchemes))) != 0) {
					// failure - on Windows 7 and older, try without URL schemes
					if(!RunTimeHelperEx::IsWin8() && SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
						// success
						if(!IsInDesignMode()) {
							properties.autoDetectedURLSchemes = L"";
							properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
						}
					} else {
						// still a failure - ensure that only valid flags are used
						if(RunTimeHelperEx::IsWin8()) {
							autoDetectURLs &= (AURL_ENABLEURL | AURL_ENABLEEMAILADDR | AURL_ENABLETELNO | AURL_ENABLEEAURLS | AURL_ENABLEDRIVELETTERS | AURL_DISABLEMIXEDLGC);
							if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, reinterpret_cast<LPARAM>(OLE2W(properties.autoDetectedURLSchemes))) == 0) {
								// success
								if(!IsInDesignMode()) {
									properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
								}
							}
						} else {
							autoDetectURLs &= AURL_ENABLEURL;
							if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
								// success
								if(!IsInDesignMode()) {
									properties.autoDetectedURLSchemes = L"";
									properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
								}
							}
						}
					}
				}
			} else {
				if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) != 0) {
					// failure - ensure that only valid flags are used
					if(RunTimeHelperEx::IsWin8()) {
						autoDetectURLs = (autoDetectURLs & (AURL_ENABLEURL | AURL_ENABLEEMAILADDR | AURL_ENABLETELNO | AURL_ENABLEEAURLS | AURL_ENABLEDRIVELETTERS | AURL_DISABLEMIXEDLGC));
					} else {
						autoDetectURLs = (autoDetectURLs & AURL_ENABLEURL);
					}
					if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
						// success
						if(!IsInDesignMode()) {
							properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
						}
					}
				}
			}
		}
		FireOnChanged(DISPID_RTB_AUTODETECTURLS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AutoScrolling(AutoScrollingConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AutoScrollingConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(SendMessage(EM_GETOPTIONS, 0, 0) & (ECO_AUTOVSCROLL | ECO_AUTOHSCROLL)) {
			case 0:
				properties.autoScrolling = asNone;
				break;
			case ECO_AUTOVSCROLL:
				properties.autoScrolling = asVertical;
				break;
			case ECO_AUTOHSCROLL:
				properties.autoScrolling = asHorizontal;
				break;
			case ECO_AUTOVSCROLL | ECO_AUTOHSCROLL:
				properties.autoScrolling = static_cast<AutoScrollingConstants>(asVertical | asHorizontal);
				break;
		}
	}

	*pValue = properties.autoScrolling;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AutoScrolling(AutoScrollingConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_AUTOSCROLLING);
	if(properties.autoScrolling != newValue) {
		properties.autoScrolling = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.autoScrolling & asHorizontal) {
				SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_AUTOHSCROLL);
			} else {
				SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_AUTOHSCROLL);
			}
			if(properties.autoScrolling & asVertical) {
				SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_AUTOVSCROLL);
			} else {
				SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_AUTOVSCROLL);
			}
		}
		FireOnChanged(DISPID_RTB_AUTOSCROLLING);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_AutoSelectWordOnTrackSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.autoSelectWordOnTrackSelection = ((SendMessage(EM_GETOPTIONS, 0, 0) & ECO_AUTOWORDSELECTION) == ECO_AUTOWORDSELECTION);
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoSelectWordOnTrackSelection);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_AutoSelectWordOnTrackSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_AUTOSELECTWORDONTRACKSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoSelectWordOnTrackSelection != b) {
		properties.autoSelectWordOnTrackSelection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			// This would work as well, but then we cannot retrieve the current setting anymore:
			/*CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				CComPtr<ITextServices> pTextServices = NULL;
				pRichEditOle->QueryInterface(IID_ITextServices, reinterpret_cast<LPVOID*>(&pTextServices));
				if(pTextServices) {
					pTextServices->OnTxPropertyBitsChange(TXTBIT_AUTOWORDSEL, (properties.autoSelectWordOnTrackSelection ? TXTBIT_AUTOWORDSEL : 0));
				}
			}*/
			if(properties.autoSelectWordOnTrackSelection) {
				SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_AUTOWORDSELECTION);
			} else {
				SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_AUTOWORDSELECTION);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_AUTOSELECTWORDONTRACKSELECTION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_RTB_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.backColor == (0x80000000 | COLOR_WINDOW)) {
				SendMessage(EM_SETBKGNDCOLOR, TRUE, 0);
			} else {
				SendMessage(EM_SETBKGNDCOLOR, FALSE, OLECOLOR2COLORREF(properties.backColor));
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_BeepOnMaxText(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.beepOnMaxText = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_BEEPONMAXTEXT) == SES_BEEPONMAXTEXT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.beepOnMaxText);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_BeepOnMaxText(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_BEEPONMAXTEXT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.beepOnMaxText != b) {
		properties.beepOnMaxText = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.beepOnMaxText) {
				SendMessage(EM_SETEDITSTYLE, SES_BEEPONMAXTEXT, SES_BEEPONMAXTEXT);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_BEEPONMAXTEXT);
			}
		}
		FireOnChanged(DISPID_RTB_BEEPONMAXTEXT);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP RichTextBox::get_CharacterConversion(CharacterConversionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, CharacterConversionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(SendMessage(EM_GETEDITSTYLE, 0, 0) & (SES_LOWERCASE | SES_UPPERCASE)) {
			case SES_LOWERCASE:
				properties.characterConversion = ccLowerCase;
				break;
			case SES_UPPERCASE:
				properties.characterConversion = ccUpperCase;
				break;
			default:
				properties.characterConversion = ccNone;
				break;
		}
	}

	*pValue = properties.characterConversion;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_CharacterConversion(CharacterConversionConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_CHARACTERCONVERSION);
	if(properties.characterConversion != newValue) {
		properties.characterConversion = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.characterConversion) {
				case ccNone:
					SendMessage(EM_SETEDITSTYLE, 0, SES_LOWERCASE | SES_UPPERCASE);
					break;
				case ccLowerCase:
					SendMessage(EM_SETEDITSTYLE, SES_LOWERCASE, SES_LOWERCASE | SES_UPPERCASE);
					break;
				case ccUpperCase:
					SendMessage(EM_SETEDITSTYLE, SES_UPPERCASE, SES_LOWERCASE | SES_UPPERCASE);
					break;
			}
		}
		FireOnChanged(DISPID_RTB_CHARACTERCONVERSION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP RichTextBox::get_CurrentIMECompositionMode(IMECompositionModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMECompositionModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<IMECompositionModeConstants>(SendMessage(EM_GETIMECOMPMODE, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::get_DefaultMathZoneHAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				switch(options & tomMathDispAlignMask) {
					case tomMathDispAlignLeft:
						properties.defaultMathZoneHAlignment = halLeft;
						break;
					case tomMathDispAlignCenter:
						properties.defaultMathZoneHAlignment = halCenter;
						break;
					case tomMathDispAlignRight:
						properties.defaultMathZoneHAlignment = halRight;
						break;
				}
			}
		}
	}

	*pValue = properties.defaultMathZoneHAlignment;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DefaultMathZoneHAlignment(HAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DEFAULTMATHZONEHALIGNMENT);
	if(properties.defaultMathZoneHAlignment != newValue) {
		properties.defaultMathZoneHAlignment = newValue;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				long options;
				switch(properties.defaultMathZoneHAlignment) {
					case halLeft:
						options = tomMathDispAlignLeft;
						break;
					case halCenter:
						options = tomMathDispAlignCenter;
						break;
					case halRight:
						options = tomMathDispAlignRight;
						break;
					default:
						options = tomMathDispAlignCenter;
						break;
				}
				ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(options, tomMathDispAlignMask)));
			}
		}
		FireOnChanged(DISPID_RTB_DEFAULTMATHZONEHALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DefaultTabWidth(FLOAT* pValue)
{
	ATLASSERT_POINTER(pValue, FLOAT);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && properties.pTextDocument) {
		properties.pTextDocument->GetDefaultTabStop(&properties.defaultTabWidth);
	}
	*pValue = properties.defaultTabWidth;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DefaultTabWidth(FLOAT newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DEFAULTTABWIDTH);
	if(properties.defaultTabWidth != newValue) {
		properties.defaultTabWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow() && properties.pTextDocument) {
			properties.pTextDocument->SetDefaultTabStop(properties.defaultTabWidth);
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_DEFAULTTABWIDTH);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DenoteEmptyMathArguments(DenoteEmptyMathArgumentsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DenoteEmptyMathArgumentsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				properties.denoteEmptyMathArguments = static_cast<DenoteEmptyMathArgumentsConstants>(options & tomMathDocEmptyArgMask);
			}
		}
	}

	*pValue = properties.denoteEmptyMathArguments;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DenoteEmptyMathArguments(DenoteEmptyMathArgumentsConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DENOTEEMPTYMATHARGUMENTS);
	if(properties.denoteEmptyMathArguments != newValue) {
		properties.denoteEmptyMathArguments = newValue;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.denoteEmptyMathArguments, tomMathDocEmptyArgMask)));
			}
		}
		FireOnChanged(DISPID_RTB_DENOTEEMPTYMATHARGUMENTS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DetectDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.detectDragDrop);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DetectDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DETECTDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.detectDragDrop != b) {
		properties.detectDragDrop = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_DETECTDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if(IsWindow()) {
			if((static_cast<long>(properties.disabledEvents) & deMouseEvents) != (static_cast<long>(newValue) & deMouseEvents)) {
				if(static_cast<long>(newValue) & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}

			DWORD events = SendMessage(EM_GETEVENTMASK, 0, 0);
			if((static_cast<long>(properties.disabledEvents) & deScrolling) != (static_cast<long>(newValue) & deScrolling)) {
				if(static_cast<long>(newValue) & deScrolling) {
					// disable scroll events
					events &= ~ENM_SCROLL;
				} else {
					// enable scroll events
					events |= ENM_SCROLL;
				}
			}
			if((static_cast<long>(properties.disabledEvents) & deSelectionChangingEvents) != (static_cast<long>(newValue) & deSelectionChangingEvents)) {
				if(static_cast<long>(newValue) & deSelectionChangingEvents) {
					// disable selection change events
					events &= ~ENM_SELCHANGE;
				} else {
					// enable selection change events
					events |= ENM_SELCHANGE;
				}
			}
			if((static_cast<long>(properties.disabledEvents) & deShouldResizeControlWindow) != (static_cast<long>(newValue) & deShouldResizeControlWindow)) {
				if(static_cast<long>(newValue) & deShouldResizeControlWindow) {
					// disable ShouldResizeControlWindow event
					events &= ~ENM_REQUESTRESIZE;
				} else {
					// enable ShouldResizeControlWindow event
					events |= ENM_REQUESTRESIZE;
				}
			}
			SendMessage(EM_SETEVENTMASK, 0, events);
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DisableIMEOperations(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.disableIMEOperations = ((GetStyle() & ES_NOIME) == ES_NOIME);
	}

	*pValue = BOOL2VARIANTBOOL(properties.disableIMEOperations);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DisableIMEOperations(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DISABLEIMEOPERATIONS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.disableIMEOperations != b) {
		properties.disableIMEOperations = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.disableIMEOperations) {
				ModifyStyle(0, ES_NOIME);
				SendMessage(EM_SETEDITSTYLE, SES_NOIME, SES_NOIME);
			} else {
				ModifyStyle(ES_NOIME, 0);
				SendMessage(EM_SETEDITSTYLE, 0, SES_NOIME);
			}
		}
		FireOnChanged(DISPID_RTB_DISABLEIMEOPERATIONS);
	}
	return S_OK;
}
/*
STDMETHODIMP RichTextBox::get_DiscardCompositionStringOnIMECancel(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.discardCompositionStringOnIMECancel = ((SendMessage(EM_GETLANGOPTIONS, 0, 0) & IMF_IMECANCELCOMPLETE) == IMF_IMECANCELCOMPLETE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.discardCompositionStringOnIMECancel);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DiscardCompositionStringOnIMECancel(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DISCARDCOMPOSITIONSTRINGONIMECANCEL);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.discardCompositionStringOnIMECancel != b) {
		properties.discardCompositionStringOnIMECancel = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD options = SendMessage(EM_GETLANGOPTIONS, 0, 0);
			if(properties.discardCompositionStringOnIMECancel) {
				options |= IMF_IMECANCELCOMPLETE;
			} else {
				options &= ~IMF_IMECANCELCOMPLETE;
			}
			SendMessage(EM_SETLANGOPTIONS, 0, options);
		}
		FireOnChanged(DISPID_RTB_DISCARDCOMPOSITIONSTRINGONIMECANCEL);
	}
	return S_OK;
}*/

STDMETHODIMP RichTextBox::get_DisplayHyperlinkTooltips(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.displayHyperlinkTooltips = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_HYPERLINKTOOLTIPS) == SES_HYPERLINKTOOLTIPS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayHyperlinkTooltips);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DisplayHyperlinkTooltips(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DISPLAYHYPERLINKTOOLTIPS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayHyperlinkTooltips != b) {
		properties.displayHyperlinkTooltips = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.displayHyperlinkTooltips) {
				SendMessage(EM_SETEDITSTYLE, SES_HYPERLINKTOOLTIPS, SES_HYPERLINKTOOLTIPS);
			} else {
				//SendMessage(EM_SETEDITSTYLE, 0, SES_HYPERLINKTOOLTIPS);
				RecreateControlWindow();
			}
		}
		FireOnChanged(DISPID_RTB_DISPLAYHYPERLINKTOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DisplayZeroWidthTableCellBorders(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.displayZeroWidthTableCellBorders = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_HIDEGRIDLINES) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayZeroWidthTableCellBorders);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DisplayZeroWidthTableCellBorders(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DISPLAYZEROWIDTHTABLECELLBORDERS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayZeroWidthTableCellBorders != b) {
		properties.displayZeroWidthTableCellBorders = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			/*if(properties.displayZeroWidthTableCellBorders) {
				SendMessage(EM_SETEDITSTYLE, 0, SES_HIDEGRIDLINES);
			} else {
				SendMessage(EM_SETEDITSTYLE, SES_HIDEGRIDLINES, SES_HIDEGRIDLINES);
			}*/
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_RTB_DISPLAYZEROWIDTHTABLECELLBORDERS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DocumentFont(IRichTextFont** ppTextFont)
{
	ATLASSERT_POINTER(ppTextFont, IRichTextFont*);
	if(!ppTextFont) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			CComPtr<ITextFont2> pFont = NULL;
			hr = pTextDocument2->GetDocumentFont(&pFont);
			if(pFont) {
				ClassFactory::InitTextFont(pFont, IID_IRichTextFont, reinterpret_cast<LPUNKNOWN*>(ppTextFont));
			}
		} else {
			hr = E_NOTIMPL;
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::put_DocumentFont(IRichTextFont* pNewTextFont)
{
	PUTPROPPROLOG(DISPID_RTB_DOCUMENTFONT);

	// invalid value - raise VB runtime error 380
	HRESULT hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	if(pNewTextFont) {
		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				CComQIPtr<ITextFont2> pFont = pNewTextFont;
				if(pFont) {
					hr = pTextDocument2->SetDocumentFont(pFont);
					FireOnChanged(DISPID_RTB_DOCUMENTFONT);
				}
			} else {
				hr = E_NOTIMPL;
			}
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::get_DocumentParagraph(IRichTextParagraph** ppTextParagraph)
{
	ATLASSERT_POINTER(ppTextParagraph, IRichTextParagraph*);
	if(!ppTextParagraph) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			CComPtr<ITextPara2> pParagraph = NULL;
			hr = pTextDocument2->GetDocumentPara(&pParagraph);
			if(pParagraph) {
				ClassFactory::InitTextParagraph(pParagraph, IID_IRichTextParagraph, reinterpret_cast<LPUNKNOWN*>(ppTextParagraph));
			}
		} else {
			hr = E_NOTIMPL;
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::put_DocumentParagraph(IRichTextParagraph* pNewTextParagraph)
{
	PUTPROPPROLOG(DISPID_RTB_DOCUMENTPARAGRAPH);

	// invalid value - raise VB runtime error 380
	HRESULT hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	if(pNewTextParagraph) {
		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				CComQIPtr<ITextPara2> pParagraph = pNewTextParagraph;
				if(pParagraph) {
					hr = pTextDocument2->SetDocumentPara(pParagraph);
					FireOnChanged(DISPID_RTB_DOCUMENTPARAGRAPH);
				}
			} else {
				hr = E_NOTIMPL;
			}
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::get_DraftMode(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.draftMode = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_DRAFTMODE) == SES_DRAFTMODE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.draftMode);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DraftMode(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DRAFTMODE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.draftMode != b) {
		properties.draftMode = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.draftMode) {
				SendMessage(EM_SETEDITSTYLE, SES_DRAFTMODE, SES_DRAFTMODE);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_DRAFTMODE);
			}
		}

		FireOnChanged(DISPID_RTB_DRAFTMODE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DraggedTextRange(IRichTextRange** ppTextRange)
{
	ATLASSERT_POINTER(ppTextRange, IRichTextRange*);
	if(!ppTextRange) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(dragDropStatus.pDraggedTextRange) {
		ClassFactory::InitTextRange(dragDropStatus.pDraggedTextRange, this, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppTextRange));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP RichTextBox::get_DragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DRAGSCROLLTIMEBASE);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragScrollTimeBase != newValue) {
		properties.dragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_DRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_DropWordsOnWordBoundariesOnly(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.dropWordsOnWordBoundariesOnly = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_WORDDRAGDROP) == SES_WORDDRAGDROP);
	}

	*pValue = BOOL2VARIANTBOOL(properties.dropWordsOnWordBoundariesOnly);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_DropWordsOnWordBoundariesOnly(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_DROPWORDSONWORDBOUNDARIESONLY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dropWordsOnWordBoundariesOnly != b) {
		properties.dropWordsOnWordBoundariesOnly = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.dropWordsOnWordBoundariesOnly) {
				SendMessage(EM_SETEDITSTYLE, SES_WORDDRAGDROP, SES_WORDDRAGDROP);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_WORDDRAGDROP);
			}
		}
		FireOnChanged(DISPID_RTB_DROPWORDSONWORDBOUNDARIESONLY);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_EffectColor(LONG index, OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			hr = pTextDocument2->GetEffectColor(index, reinterpret_cast<LONG*>(pValue));
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::put_EffectColor(LONG index, OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_RTB_EFFECTCOLOR);
	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			hr = pTextDocument2->SetEffectColor(index, newValue);
		}
	}
	if(SUCCEEDED(hr)) {
		FireViewChange();
	}
	return hr;
}

STDMETHODIMP RichTextBox::get_EmbeddedObjects(IRichOLEObjects** ppOLEObjects)
{
	ATLASSERT_POINTER(ppOLEObjects, IRichOLEObjects*);
	if(!ppOLEObjects) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		ClassFactory::InitOLEObjects(this, IID_IRichOLEObjects, reinterpret_cast<LPUNKNOWN*>(ppOLEObjects));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP RichTextBox::get_EmulateSimpleTextBox(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.emulateSimpleTextBox = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_EMULATESYSEDIT) == SES_EMULATESYSEDIT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.emulateSimpleTextBox);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_EmulateSimpleTextBox(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_EMULATESIMPLETEXTBOX);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.emulateSimpleTextBox != b) {
		properties.emulateSimpleTextBox = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			// Most of the times this wont work:
			/*if(properties.emulateSimpleTextBox) {
				SendMessage(EM_SETEDITSTYLE, SES_EMULATESYSEDIT, SES_EMULATESYSEDIT);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_EMULATESYSEDIT);
			}*/
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_RTB_EMULATESIMPLETEXTBOX);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_RTB_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_ExtendFontBackColorToWholeLine(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.extendFontBackColorToWholeLine = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_EXTENDBACKCOLOR) == SES_EXTENDBACKCOLOR);
	}

	*pValue = BOOL2VARIANTBOOL(properties.extendFontBackColorToWholeLine);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_ExtendFontBackColorToWholeLine(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_EXTENDFONTBACKCOLORTOWHOLELINE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.extendFontBackColorToWholeLine != b) {
		properties.extendFontBackColorToWholeLine = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.extendFontBackColorToWholeLine) {
				SendMessage(EM_SETEDITSTYLE, SES_EXTENDBACKCOLOR, SES_EXTENDBACKCOLOR);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_EXTENDBACKCOLOR);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_EXTENDFONTBACKCOLORTOWHOLELINE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_FirstVisibleLine(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(GetStyle() & ES_MULTILINE) {
			*pValue = static_cast<LONG>(SendMessage(EM_GETFIRSTVISIBLELINE, 0, 0));
		}
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_RTB_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	ApplyFont();

	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_FONT);
	return S_OK;
}

STDMETHODIMP RichTextBox::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_RTB_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	ApplyFont();

	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_FONT);
	return S_OK;
}

STDMETHODIMP RichTextBox::get_FormattingRectangleHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	*pValue = properties.formattingRectangle.Height();
	return S_OK;
}

STDMETHODIMP RichTextBox::put_FormattingRectangleHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_RTB_FORMATTINGRECTANGLEHEIGHT);
	if(properties.formattingRectangle.Height() != newValue) {
		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
		}
		properties.formattingRectangle.bottom = properties.formattingRectangle.top + newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, FALSE, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, FALSE, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_FORMATTINGRECTANGLEHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_FormattingRectangleLeft(OLE_XPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	*pValue = properties.formattingRectangle.left;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_FormattingRectangleLeft(OLE_XPOS_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_RTB_FORMATTINGRECTANGLELEFT);
	if(properties.formattingRectangle.left != newValue) {
		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
		}
		properties.formattingRectangle.MoveToX(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, FALSE, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, FALSE, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_FORMATTINGRECTANGLELEFT);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_FormattingRectangleTop(OLE_YPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	*pValue = properties.formattingRectangle.top;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_FormattingRectangleTop(OLE_YPOS_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_RTB_FORMATTINGRECTANGLETOP);
	if(properties.formattingRectangle.top != newValue) {
		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
		}
		properties.formattingRectangle.MoveToY(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, FALSE, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, FALSE, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_FORMATTINGRECTANGLETOP);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_FormattingRectangleWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	*pValue = properties.formattingRectangle.Width();
	return S_OK;
}

STDMETHODIMP RichTextBox::put_FormattingRectangleWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_RTB_FORMATTINGRECTANGLEWIDTH);
	if(properties.formattingRectangle.Width() != newValue) {
		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_GETRECT, 0, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
		}
		properties.formattingRectangle.right = properties.formattingRectangle.left + newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, FALSE, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, FALSE, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_FORMATTINGRECTANGLEWIDTH);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_GrowNAryOperators(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				properties.growNAryOperators = ((options & tomMathDispNaryGrow) == tomMathDispNaryGrow);
			}
		}
	}

	*pValue = BOOL2VARIANTBOOL(properties.growNAryOperators);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_GrowNAryOperators(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_GROWNARYOPERATORS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.growNAryOperators != b) {
		properties.growNAryOperators = b;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.growNAryOperators ? tomMathDispNaryGrow : 0, tomMathDispNaryGrow)));
			}
		}
		FireOnChanged(DISPID_RTB_GROWNARYOPERATORS);
	}
	return S_OK;
}
//
//STDMETHODIMP RichTextBox::get_HAlignment(HAlignmentConstants* pValue)
//{
//	ATLASSERT_POINTER(pValue, HAlignmentConstants);
//	if(!pValue) {
//		return E_POINTER;
//	}
//
//	if(!IsInDesignMode() && IsWindow()) {
//		switch(GetStyle() & (ES_LEFT | ES_CENTER | ES_RIGHT)) {
//			case ES_CENTER:
//				properties.hAlignment = halCenter;
//				break;
//			case ES_RIGHT:
//				properties.hAlignment = halRight;
//				break;
//			case ES_LEFT:
//				properties.hAlignment = halLeft;
//				break;
//		}
//	}
//
//	*pValue = properties.hAlignment;
//	return S_OK;
//}
//
//STDMETHODIMP RichTextBox::put_HAlignment(HAlignmentConstants newValue)
//{
//	PUTPROPPROLOG(DISPID_RTB_HALIGNMENT);
//	if(properties.hAlignment != newValue) {
//		properties.hAlignment = newValue;
//		SetDirty(TRUE);
//
//		if(IsWindow()) {
//			RecreateControlWindow();
//		}
//		FireOnChanged(DISPID_RTB_HALIGNMENT);
//	}
//	return S_OK;
//}

STDMETHODIMP RichTextBox::get_hDragImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(dragDropStatus.IsDragging()) {
		*pValue = HandleToLong(dragDropStatus.hDragImageList);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_RTB_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_HyphenationFunction(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		HYPHENATEINFO hyphenateInfo = {0};
		hyphenateInfo.cbSize = sizeof(HYPHENATEINFO);
		SendMessage(EM_GETHYPHENATEINFO, reinterpret_cast<WPARAM>(&hyphenateInfo), 0);
		properties.hyphenationFunction = hyphenateInfo.pfnHyphenate;
	}

	*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.hyphenationFunction));
	return S_OK;
}

STDMETHODIMP RichTextBox::put_HyphenationFunction(LONG newValue)
{
	PUTPROPPROLOG(DISPID_RTB_HYPHENATIONFUNCTION);
	PFNHYPHENATEPROC pfnHyphenateProc = reinterpret_cast<PFNHYPHENATEPROC>(static_cast<LONG_PTR>(newValue));
	if(properties.hyphenationFunction != pfnHyphenateProc) {
		properties.hyphenationFunction = pfnHyphenateProc;
		SetDirty(TRUE);

		if(IsWindow()) {
			HYPHENATEINFO hyphenateInfo = {0};
			hyphenateInfo.cbSize = sizeof(HYPHENATEINFO);
			SendMessage(EM_GETHYPHENATEINFO, reinterpret_cast<WPARAM>(&hyphenateInfo), 0);
			hyphenateInfo.pfnHyphenate = reinterpret_cast<PFNHYPHENATEPROC>(properties.hyphenationFunction);
			SendMessage(EM_SETHYPHENATEINFO, reinterpret_cast<WPARAM>(&hyphenateInfo), 0);
		}
		FireOnChanged(DISPID_RTB_HYPHENATIONFUNCTION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_HyphenationWordWrapZoneWidth(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		HYPHENATEINFO hyphenateInfo = {0};
		hyphenateInfo.cbSize = sizeof(HYPHENATEINFO);
		SendMessage(EM_GETHYPHENATEINFO, reinterpret_cast<WPARAM>(&hyphenateInfo), 0);
		properties.hyphenationWordWrapZoneWidth = hyphenateInfo.dxHyphenateZone;
	}

	*pValue = properties.hyphenationWordWrapZoneWidth;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_HyphenationWordWrapZoneWidth(SHORT newValue)
{
	PUTPROPPROLOG(DISPID_RTB_HYPHENATIONWORDWRAPZONEWIDTH);
	if(properties.hyphenationWordWrapZoneWidth != newValue) {
		properties.hyphenationWordWrapZoneWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			HYPHENATEINFO hyphenateInfo = {0};
			hyphenateInfo.cbSize = sizeof(HYPHENATEINFO);
			SendMessage(EM_GETHYPHENATEINFO, reinterpret_cast<WPARAM>(&hyphenateInfo), 0);
			hyphenateInfo.dxHyphenateZone = properties.hyphenationWordWrapZoneWidth;
			SendMessage(EM_SETHYPHENATEINFO, reinterpret_cast<WPARAM>(&hyphenateInfo), 0);
		}
		FireOnChanged(DISPID_RTB_HYPHENATIONWORDWRAPZONEWIDTH);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP RichTextBox::get_IMEConversionMode(IMEConversionModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEConversionModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && SendMessage(EM_ISIME, 0, 0)) {
		properties.IMEConversionMode = static_cast<IMEConversionModeConstants>(SendMessage(EM_GETIMEMODEBIAS, 0, 0));
	}

	*pValue = properties.IMEConversionMode;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_IMEConversionMode(IMEConversionModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_IMECONVERSIONMODE);
	if(properties.IMEConversionMode != newValue) {
		properties.IMEConversionMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode() && IsWindow() && SendMessage(EM_ISIME, 0, 0)) {
			SendMessage(EM_SETIMEMODEBIAS, properties.IMEConversionMode, properties.IMEConversionMode);
		}
		FireOnChanged(DISPID_RTB_IMECONVERSIONMODE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_IMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.IMEMode;
	} else {
		if((GetFocus() == *this) && (GetEffectiveIMEMode() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(*this);
			if((ime != imeInherit) && (properties.IMEMode != imeInherit)) {
				properties.IMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode();
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == *this) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(*this, ime);
				}
			}
		}
		FireOnChanged(DISPID_RTB_IMEMODE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_IntegralLimitsLocation(LimitsLocationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, LimitsLocationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				if(options & tomMathDispIntUnderOver) {
					properties.integralLimitsLocation = llUnderAndOver;
				} else {
					properties.integralLimitsLocation = llOnSide;
				}
			}
		}
	}

	*pValue = properties.integralLimitsLocation;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_IntegralLimitsLocation(LimitsLocationConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_INTEGRALLIMITSLOCATION);
	if(properties.integralLimitsLocation != newValue) {
		properties.integralLimitsLocation = newValue;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				switch(properties.integralLimitsLocation) {
					case llOnSide:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(0, tomMathDispIntUnderOver)));
						break;
					case llUnderAndOver:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathDispIntUnderOver, tomMathDispIntUnderOver)));
						break;
				}
			}
		}
		FireOnChanged(DISPID_RTB_INTEGRALLIMITSLOCATION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP RichTextBox::get_LeftMargin(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	// Not supported by RichEdit! And because of EC_USEFONTINFO we cannot really intercept EM_GETMARGINS.
	/*if(!IsInDesignMode() && IsWindow()) {
		properties.leftMargin = LOWORD(SendMessage(EM_GETMARGINS, 0, 0));
	}*/

	*pValue = properties.leftMargin;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_LeftMargin(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_RTB_LEFTMARGIN);
	if(properties.leftMargin != newValue) {
		properties.leftMargin = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			// Does not work for RichEdit:
			//SendMessage(EM_SETMARGINS, EC_LEFTMARGIN, MAKELPARAM((properties.leftMargin == -1 ? EC_USEFONTINFO : properties.leftMargin), 0));
			// EC_USEFONTINFO does not seem to work at all!
			/*if(properties.leftMargin == -1 || properties.rightMargin == -1) {
				SendMessage(EM_SETMARGINS, EC_USEFONTINFO, 0);
			} else {*/
				SendMessage(EM_SETMARGINS, EC_LEFTMARGIN, MAKELPARAM((properties.leftMargin == -1 ? 1 : properties.leftMargin), 0));
			//}
		}
		FireOnChanged(DISPID_RTB_LEFTMARGIN);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_LetClientHandleAllIMEOperations(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.letClientHandleAllIMEOperations = ((GetStyle() & ES_SELFIME) == ES_SELFIME);
	}

	*pValue = BOOL2VARIANTBOOL(properties.letClientHandleAllIMEOperations);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_LetClientHandleAllIMEOperations(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_LETCLIENTHANDLEALLIMEOPERATIONS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.letClientHandleAllIMEOperations != b) {
		properties.letClientHandleAllIMEOperations = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.letClientHandleAllIMEOperations) {
				ModifyStyle(0, ES_SELFIME);
			} else {
				ModifyStyle(ES_SELFIME, 0);
			}
		}
		FireOnChanged(DISPID_RTB_LETCLIENTHANDLEALLIMEOPERATIONS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_LinkMouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.linkMouseIcon.pPictureDisp) {
		properties.linkMouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::put_LinkMouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_RTB_LINKMOUSEICON);
	if(properties.linkMouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.linkMouseIcon.StopWatching();
		if(properties.linkMouseIcon.pPictureDisp) {
			properties.linkMouseIcon.pPictureDisp->Release();
			properties.linkMouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.linkMouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.linkMouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_LINKMOUSEICON);
	return S_OK;
}

STDMETHODIMP RichTextBox::putref_LinkMouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_RTB_LINKMOUSEICON);
	if(properties.linkMouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.linkMouseIcon.StopWatching();
		if(properties.linkMouseIcon.pPictureDisp) {
			properties.linkMouseIcon.pPictureDisp->Release();
			properties.linkMouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.linkMouseIcon.pPictureDisp));
		}
		properties.linkMouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_LINKMOUSEICON);
	return S_OK;
}

STDMETHODIMP RichTextBox::get_LinkMousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.linkMousePointer;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_LinkMousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_LINKMOUSEPOINTER);
	if(properties.linkMousePointer != newValue) {
		properties.linkMousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_LINKMOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_LogicalCaret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.logicalCaret = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_LOGICALCARET) == SES_LOGICALCARET);
	}

	*pValue = BOOL2VARIANTBOOL(properties.logicalCaret);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_LogicalCaret(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_LOGICALCARET);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.logicalCaret != b) {
		properties.logicalCaret = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.logicalCaret) {
				SendMessage(EM_SETEDITSTYLE, SES_LOGICALCARET, SES_LOGICALCARET);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_LOGICALCARET);
			}
		}
		FireOnChanged(DISPID_RTB_LOGICALCARET);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_MathLineBreakBehavior(MathLineBreakBehaviorConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MathLineBreakBehaviorConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				switch(options & tomMathBrkBinMask) {
					case tomMathBrkBinBefore:
						properties.mathLineBreakBehavior = mlbbBreakBeforeOperator;
						break;
					case tomMathBrkBinAfter:
						properties.mathLineBreakBehavior = mlbbBreakAfterOperator;
						break;
					case tomMathBrkBinDup:
						switch(options & tomMathBrkBinSubMask) {
							case tomMathBrkBinSubMM:
								properties.mathLineBreakBehavior = mlbbDuplicateOperatorMinusMinus;
								break;
							case tomMathBrkBinSubPM:
								properties.mathLineBreakBehavior = mlbbDuplicateOperatorPlusMinus;
								break;
							case tomMathBrkBinSubMP:
								properties.mathLineBreakBehavior = mlbbDuplicateOperatorMinusPlus;
								break;
						}
						break;
				}
			}
		}
	}

	*pValue = properties.mathLineBreakBehavior;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_MathLineBreakBehavior(MathLineBreakBehaviorConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_MATHLINEBREAKBEHAVIOR);
	if(properties.mathLineBreakBehavior != newValue) {
		properties.mathLineBreakBehavior = newValue;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				switch(properties.mathLineBreakBehavior) {
					case mlbbBreakBeforeOperator:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathBrkBinBefore, tomMathBrkBinMask)));
						break;
					case mlbbBreakAfterOperator:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathBrkBinAfter, tomMathBrkBinMask)));
						break;
					case mlbbDuplicateOperatorMinusMinus:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathBrkBinDup | tomMathBrkBinSubMM, tomMathBrkBinMask | tomMathBrkBinSubMask)));
						break;
					case mlbbDuplicateOperatorPlusMinus:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathBrkBinDup | tomMathBrkBinSubPM, tomMathBrkBinMask | tomMathBrkBinSubMask)));
						break;
					case mlbbDuplicateOperatorMinusPlus:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathBrkBinDup | tomMathBrkBinSubMP, tomMathBrkBinMask | tomMathBrkBinSubMask)));
						break;
				}
			}
		}
		FireOnChanged(DISPID_RTB_MATHLINEBREAKBEHAVIOR);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_MaxTextLength(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.maxTextLength = static_cast<LONG>(SendMessage(EM_GETLIMITTEXT, 0, 0));
	}

	*pValue = properties.maxTextLength;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_MaxTextLength(LONG newValue)
{
	PUTPROPPROLOG(DISPID_RTB_MAXTEXTLENGTH);
	if(properties.maxTextLength != newValue) {
		properties.maxTextLength = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_EXLIMITTEXT, 0, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength));
		}
		FireOnChanged(DISPID_RTB_MAXTEXTLENGTH);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_MaxUndoQueueSize(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			pTextDocument2->GetProperty(tomUndoLimit, &properties.maxUndoQueueSize);
		}
	}
	*pValue = properties.maxUndoQueueSize;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_MaxUndoQueueSize(LONG newValue)
{
	PUTPROPPROLOG(DISPID_RTB_MAXUNDOQUEUESIZE);
	if(properties.maxUndoQueueSize != newValue) {
		properties.maxUndoQueueSize = newValue;
		SetDirty(TRUE);

		BOOL useFallback = TRUE;
		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				useFallback = FAILED(pTextDocument2->SetProperty(tomUndoLimit, newValue));
			}
		}
		if(useFallback && IsWindow()) {
			SendMessage(EM_SETUNDOLIMIT, properties.maxUndoQueueSize, 0);
		}
		FireOnChanged(DISPID_RTB_MAXUNDOQUEUESIZE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Modified(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.modified = SendMessage(EM_GETMODIFY, 0, 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.modified);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_Modified(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_MODIFIED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.modified != b) {
		properties.modified = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SETMODIFY, properties.modified, 0);
		}
		FireOnChanged(DISPID_RTB_MODIFIED);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_RTB_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP RichTextBox::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_RTB_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP RichTextBox::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_MultiLine(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiLine = ((GetStyle() & ES_MULTILINE) == ES_MULTILINE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.multiLine);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_MultiLine(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_MULTILINE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiLine != b) {
		properties.multiLine = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			// This would work as well, but then we cannot retrieve the current setting anymore:
			/*CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				CComPtr<ITextServices> pTextServices = NULL;
				pRichEditOle->QueryInterface(IID_ITextServices, reinterpret_cast<LPVOID*>(&pTextServices));
				if(pTextServices) {
					pTextServices->OnTxPropertyBitsChange(TXTBIT_MULTILINE, (properties.multiLine ? TXTBIT_MULTILINE : 0));
				}
			}*/
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_RTB_MULTILINE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_MultiSelect(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiSelect = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_MULTISELECT) == SES_MULTISELECT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.multiSelect);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_MultiSelect(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_MULTISELECT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiSelect != b) {
		properties.multiSelect = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.multiSelect) {
				SendMessage(EM_SETEDITSTYLE, SES_MULTISELECT, SES_MULTISELECT);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_MULTISELECT);
			}
		}
		FireOnChanged(DISPID_RTB_MULTISELECT);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_NAryLimitsLocation(LimitsLocationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, LimitsLocationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				if(options & tomMathDispNarySubSup) {
					properties.nAryLimitsLocation = llOnSide;
				} else {
					properties.nAryLimitsLocation = llUnderAndOver;
				}
			}
		}
	}

	*pValue = properties.nAryLimitsLocation;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_NAryLimitsLocation(LimitsLocationConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_NARYLIMITSLOCATION);
	if(properties.nAryLimitsLocation != newValue) {
		properties.nAryLimitsLocation = newValue;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				switch(properties.nAryLimitsLocation) {
					case llOnSide:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathDispNarySubSup, tomMathDispNarySubSup)));
						break;
					case llUnderAndOver:
						ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(0, tomMathDispNarySubSup)));
						break;
				}
			}
		}
		FireOnChanged(DISPID_RTB_NARYLIMITSLOCATION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_NextRedoActionType(UndoActionTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UndoActionTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = uatUnknown;
	if(IsWindow()) {
		*pValue = static_cast<UndoActionTypeConstants>(SendMessage(EM_GETREDONAME, 0, 0));
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_NextUndoActionType(UndoActionTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UndoActionTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = uatUnknown;
	if(IsWindow()) {
		*pValue = static_cast<UndoActionTypeConstants>(SendMessage(EM_GETUNDONAME, 0, 0));
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_NoInputSequenceCheck(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.noInputSequenceCheck = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_NOINPUTSEQUENCECHK) == SES_NOINPUTSEQUENCECHK);
	}

	*pValue = BOOL2VARIANTBOOL(properties.noInputSequenceCheck);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_NoInputSequenceCheck(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_NOINPUTSEQUENCECHECK);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.noInputSequenceCheck != b) {
		properties.noInputSequenceCheck = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.noInputSequenceCheck) {
				SendMessage(EM_SETEDITSTYLE, SES_NOINPUTSEQUENCECHK, SES_NOINPUTSEQUENCECHK);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_NOINPUTSEQUENCECHK);
			}
		}
		FireOnChanged(DISPID_RTB_NOINPUTSEQUENCECHECK);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.oleDragImageStyle;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_OLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_OLEDRAGIMAGESTYLE);
	if(properties.oleDragImageStyle != newValue) {
		properties.oleDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_OLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP RichTextBox::get_RawSubScriptAndSuperScriptOperators(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				properties.rawSubScriptAndSuperScriptOperators = ((options & tomMathDocSbSpOpUnchanged) == tomMathDocSbSpOpUnchanged);
			}
		}
	}

	*pValue = BOOL2VARIANTBOOL(properties.rawSubScriptAndSuperScriptOperators);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_RawSubScriptAndSuperScriptOperators(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_RAWSUBSCRIPTANDSUPERSCRIPTOPERATORS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.rawSubScriptAndSuperScriptOperators != b) {
		properties.rawSubScriptAndSuperScriptOperators = b;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.rawSubScriptAndSuperScriptOperators ? tomMathDocSbSpOpUnchanged : 0, tomMathDocSbSpOpUnchanged)));
			}
		}
		FireOnChanged(DISPID_RTB_RAWSUBSCRIPTANDSUPERSCRIPTOPERATORS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_ReadOnly(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.readOnly = ((GetStyle() & ES_READONLY) == ES_READONLY);
	}

	*pValue = BOOL2VARIANTBOOL(properties.readOnly);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_ReadOnly(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_READONLY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.readOnly != b) {
		properties.readOnly = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			// This would work as well, but then we cannot retrieve the current setting anymore:
			/*CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				CComPtr<ITextServices> pTextServices = NULL;
				pRichEditOle->QueryInterface(IID_ITextServices, reinterpret_cast<LPVOID*>(&pTextServices));
				if(pTextServices) {
					pTextServices->OnTxPropertyBitsChange(TXTBIT_READONLY, (properties.readOnly ? TXTBIT_READONLY : 0));
				}
			}*/
			SendMessage(EM_SETREADONLY, properties.readOnly, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_READONLY);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RegisterForOLEDragDropConstants);
	if(!pValue) {
		return E_POINTER;
	}

	/*if(!IsInDesignMode() && IsWindow()) {
		if((GetStyle() & ES_NOOLEDRAGDROP) == 0) {
			properties.registerForOLEDragDrop = rfoddNativeDragDrop;
		}
	}*/

	*pValue = properties.registerForOLEDragDrop;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_REGISTERFOROLEDRAGDROP);
	if(properties.registerForOLEDragDrop != newValue) {
		properties.registerForOLEDragDrop = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				CComPtr<ITextServices> pTextServices = NULL;
				pRichEditOle->QueryInterface(IID_ITextServices, reinterpret_cast<LPVOID*>(&pTextServices));
				if(pTextServices) {
					switch(properties.registerForOLEDragDrop) {
						case rfoddNoDragDrop:
							pTextServices->OnTxPropertyBitsChange(TXTBIT_DISABLEDRAG, TXTBIT_DISABLEDRAG);
							//ModifyStyle(0, ES_NOOLEDRAGDROP);
							RevokeDragDrop(*this);					// even with ES_NOOLEDRAGDROP the native control calls RegisterDragDrop
							break;
						case rfoddNativeDragDrop:
							pTextServices->OnTxPropertyBitsChange(TXTBIT_DISABLEDRAG, 0);
							//ModifyStyle(ES_NOOLEDRAGDROP, 0);
							break;
						case rfoddAdvancedDragDrop: {
							pTextServices->OnTxPropertyBitsChange(TXTBIT_DISABLEDRAG, TXTBIT_DISABLEDRAG);
							//ModifyStyle(0, ES_NOOLEDRAGDROP);
							RevokeDragDrop(*this);					// even with ES_NOOLEDRAGDROP the native control calls RegisterDragDrop
							ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
							break;
						}
					}
				}
			}
			//RecreateControlWindow();
		}
		FireOnChanged(DISPID_RTB_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_RichEditAPIVersion(RichEditAPIVersionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RichEditAPIVersionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.richEditAPIVersion;
	return S_OK;
}

STDMETHODIMP RichTextBox::get_RichEditLibraryPath(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.richEditLibraryPath;
	return S_OK;
}

STDMETHODIMP RichTextBox::get_RichEditVersion(RichEditVersionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RichEditVersionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.richEditVersion;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_RichEditVersion(RichEditVersionConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_RICHEDITVERSION);
	if(properties.richEditVersion != newValue) {
		if(IsWindow() && !IsInDesignMode()) {
			// set not supported at runtime - raise VB runtime error 382
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382);
		}
		properties.richEditVersion = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_RTB_RICHEDITVERSION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_RichText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(!IsInDesignMode() && IsWindow()) {
		CComPtr<IRichTextRange> pTextRange = NULL;
		hr = get_TextRange(0, -1, &pTextRange);
		if(SUCCEEDED(hr)) {
			hr = pTextRange->get_RichText(pValue);
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::put_RichText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_RTB_RICHTEXT);

	HRESULT hr = E_FAIL;
	if(!IsInDesignMode() && IsWindow()) {
		CComPtr<IRichTextRange> pTextRange = NULL;
		hr = get_TextRange(0, -1, &pTextRange);
		if(SUCCEEDED(hr)) {
			hr = pTextRange->put_RichText(newValue);
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::get_RightMargin(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	// Not supported by RichEdit! And because of EC_USEFONTINFO we cannot really intercept EM_GETMARGINS.
	/*if(!IsInDesignMode() && IsWindow()) {
		properties.rightMargin = HIWORD(SendMessage(EM_GETMARGINS, 0, 0));
	}*/

	*pValue = properties.rightMargin;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_RightMargin(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_RTB_RIGHTMARGIN);
	if(properties.rightMargin != newValue) {
		properties.rightMargin = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			// Does not work for RichEdit:
			//SendMessage(EM_SETMARGINS, EC_RIGHTMARGIN, MAKELPARAM(0, (properties.rightMargin == -1 ? EC_USEFONTINFO : properties.rightMargin)));
			// EC_USEFONTINFO does not seem to work at all!
			/*if(properties.leftMargin == -1 || properties.rightMargin == -1) {
				SendMessage(EM_SETMARGINS, EC_USEFONTINFO, 0);
			} else {*/
				SendMessage(EM_SETMARGINS, EC_RIGHTMARGIN, MAKELPARAM(0, (properties.rightMargin == -1 ? 1 : properties.rightMargin)));
			//}
		}
		FireOnChanged(DISPID_RTB_RIGHTMARGIN);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		DWORD style = GetExStyle();
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			BOOL requiresRecreation = FALSE;
			if(properties.rightToLeft & rtlText) {
				if(!(GetExStyle() & WS_EX_RTLREADING)) {
					requiresRecreation = TRUE;
				}
			} else {
				if(GetExStyle() & WS_EX_RTLREADING) {
					requiresRecreation = TRUE;
				}
			}
			if(properties.rightToLeft & rtlLayout) {
				if(!(GetExStyle() & WS_EX_LAYOUTRTL) && !RunTimeHelper::IsVista()) {
					requiresRecreation = TRUE;
				}
			} else {
				if(GetExStyle() & WS_EX_LAYOUTRTL && !RunTimeHelper::IsVista()) {
					requiresRecreation = TRUE;
				}
			}

			if(requiresRecreation) {
				RecreateControlWindow();
			} else {
				if(properties.rightToLeft & rtlLayout) {
					ModifyStyleEx(0, WS_EX_LAYOUTRTL);
				} else {
					ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
				}
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_RTB_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_RightToLeftMathZones(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				properties.rightToLeftMathZones = ((options & tomMathEnableRtl) == tomMathEnableRtl);
			}
		}
	}

	*pValue = properties.rightToLeftMathZones;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_RightToLeftMathZones(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_RIGHTTOLEFTMATHZONES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.rightToLeftMathZones != b) {
		properties.rightToLeftMathZones = b;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.rightToLeftMathZones ? tomMathEnableRtl : 0, tomMathEnableRtl)));
			}
		}
		FireOnChanged(DISPID_RTB_RIGHTTOLEFTMATHZONES);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_ScrollBars(ScrollBarsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ScrollBarsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (WS_VSCROLL | WS_HSCROLL)) {
			case 0:
				properties.scrollBars = sbNone;
				break;
			case WS_VSCROLL:
				properties.scrollBars = sbVertical;
				break;
			case WS_HSCROLL:
				properties.scrollBars = sbHorizontal;
				break;
			case WS_VSCROLL | WS_HSCROLL:
				properties.scrollBars = static_cast<ScrollBarsConstants>(sbVertical | sbHorizontal);
				break;
		}
	}

	*pValue = properties.scrollBars;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_ScrollBars(ScrollBarsConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_SCROLLBARS);
	if(properties.scrollBars != newValue) {
		properties.scrollBars = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SHOWSCROLLBAR, SB_VERT, (properties.scrollBars & sbVertical) == sbVertical);
			SendMessage(EM_SHOWSCROLLBAR, SB_HORZ, (properties.scrollBars & sbHorizontal) == sbHorizontal);
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_SCROLLBARS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_ScrollToTopOnKillFocus(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.scrollToTopOnKillFocus = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_SCROLLONKILLFOCUS) == SES_SCROLLONKILLFOCUS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.scrollToTopOnKillFocus);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_ScrollToTopOnKillFocus(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_SCROLLTOTOPONKILLFOCUS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.scrollToTopOnKillFocus != b) {
		properties.scrollToTopOnKillFocus = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.scrollToTopOnKillFocus) {
				SendMessage(EM_SETEDITSTYLE, SES_SCROLLONKILLFOCUS, SES_SCROLLONKILLFOCUS);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_SCROLLONKILLFOCUS);
			}
		}
		FireOnChanged(DISPID_RTB_SCROLLTOTOPONKILLFOCUS);
	}
	return S_OK;
}
//
//STDMETHODIMP RichTextBox::get_SelectedText(BSTR* pValue)
//{
//	ATLASSERT_POINTER(pValue, BSTR);
//	if(!pValue) {
//		return E_POINTER;
//	}
//
//	if(IsWindow()) {
//		CHARRANGE range;
//		SendMessage(EM_EXGETSEL, 0, reinterpret_cast<LPARAM>(&range));
//		int bufferSize = 0;
//		if(range.cpMin == 0 && range.cpMax == -1) {
//			bufferSize = GetWindowTextLength();
//		} else if(range.cpMax != range.cpMin) {
//			bufferSize = abs(range.cpMax - range.cpMin);
//		}
//		LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (bufferSize + 1) * sizeof(TCHAR)));
//		if(pBuffer) {
//			ZeroMemory(pBuffer, (bufferSize + 1) * sizeof(TCHAR));
//			SendMessage(EM_GETSELTEXT, 0, reinterpret_cast<LPARAM>(pBuffer));
//			*pValue = _bstr_t(pBuffer).Detach();
//			HeapFree(GetProcessHeap(), 0, pBuffer);
//			return S_OK;
//		}
//	}
//	return E_FAIL;
//}

STDMETHODIMP RichTextBox::get_SelectedTextRange(IRichTextRange** ppTextRange)
{
	ATLASSERT_POINTER(ppTextRange, IRichTextRange*);
	if(!ppTextRange) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		if(properties.pTextDocument) {
			CComPtr<ITextSelection> pTextSelection = NULL;
			hr = properties.pTextDocument->GetSelection(&pTextSelection);
			// TODO: Better provide an ITextSelection object (ITextDocument::GetSelection)
			CComQIPtr<ITextRange> pTextRange = pTextSelection;
			ClassFactory::InitTextRange(pTextRange, this, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppTextRange));
		}
	}
	return hr;
}

STDMETHODIMP RichTextBox::get_SelectionType(SelectionTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ScrollBarsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = setEmpty;
	if(IsWindow()) {
		*pValue = static_cast<SelectionTypeConstants>(SendMessage(EM_SELECTIONTYPE, 0, 0));
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	*pValue = BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP RichTextBox::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_SHOWDRAGIMAGE);
	if(!dragDropStatus.hDragImageList && !dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage(FALSE);
	} else {
		dragDropStatus.ShowDragImage(FALSE);
	}

	FireOnChanged(DISPID_RTB_SHOWDRAGIMAGE);
	SendOnDataChange(NULL);
	return S_OK;
}

STDMETHODIMP RichTextBox::get_ShowSelectionBar(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showSelectionBar = ((SendMessage(EM_GETOPTIONS, 0, 0) & ECO_SELECTIONBAR) == ECO_SELECTIONBAR);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showSelectionBar);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_ShowSelectionBar(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_SHOWSELECTIONBAR);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showSelectionBar != b) {
		properties.showSelectionBar = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showSelectionBar) {
				SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_SELECTIONBAR);
			} else {
				SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_SELECTIONBAR);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_SHOWSELECTIONBAR);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_SmartSpacingOnDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.smartSpacingOnDrop = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_SMARTDRAGDROP) == SES_SMARTDRAGDROP);
	}

	*pValue = BOOL2VARIANTBOOL(properties.smartSpacingOnDrop);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_SmartSpacingOnDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_SMARTSPACINGONDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.smartSpacingOnDrop != b) {
		properties.smartSpacingOnDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.smartSpacingOnDrop) {
				SendMessage(EM_SETEDITSTYLE, SES_SMARTDRAGDROP, SES_SMARTDRAGDROP);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_SMARTDRAGDROP);
			}
		}
		FireOnChanged(DISPID_RTB_SMARTSPACINGONDROP);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_SwitchFontOnIMEInput(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.switchFontOnIMEInput = ((SendMessage(EM_GETLANGOPTIONS, 0, 0) & IMF_AUTOFONT) == IMF_AUTOFONT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.switchFontOnIMEInput);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_SwitchFontOnIMEInput(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_SWITCHFONTONIMEINPUT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.switchFontOnIMEInput != b) {
		properties.switchFontOnIMEInput = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD options = SendMessage(EM_GETLANGOPTIONS, 0, 0);
			if(properties.switchFontOnIMEInput) {
				options |= IMF_AUTOFONT;
			} else {
				options &= ~IMF_AUTOFONT;
			}
			SendMessage(EM_SETLANGOPTIONS, 0, options);
		}
		FireOnChanged(DISPID_RTB_SWITCHFONTONIMEINPUT);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		GetWindowText(&properties.text);
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP RichTextBox::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_RTB_TEXT);
	properties.text.AssignBSTR(newValue);
	if(IsWindow()) {
		SetWindowText(COLE2CT(properties.text));
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_TEXT);
	SendOnDataChange();
	return S_OK;
}

STDMETHODIMP RichTextBox::get_TextFlow(TextFlowConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TextFlowConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.textFlow = static_cast<TextFlowConstants>(SendMessage(EM_GETPAGEROTATE, 0, 0));
	}

	*pValue = properties.textFlow;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_TextFlow(TextFlowConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_TEXTFLOW);
	if(properties.textFlow != newValue) {
		properties.textFlow = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(EM_SETPAGEROTATE, properties.textFlow, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_TEXTFLOW);
	}
	return S_OK;
}
//
//STDMETHODIMP RichTextBox::get_TextLength(LONG* pValue)
//{
//	ATLASSERT_POINTER(pValue, LONG);
//	if(!pValue) {
//		return E_POINTER;
//	}
//
//	if(IsWindow()) {
//		*pValue = GetWindowTextLength();
//	}
//	return S_OK;
//}

STDMETHODIMP RichTextBox::get_TextOrientation(TextOrientationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TextOrientationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if((SendMessage(EM_GETOPTIONS, 0, 0) & ECO_VERTICAL) == ECO_VERTICAL) {
			if((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_USEATFONT) == SES_USEATFONT) {
				properties.textOrientation = toVertical;
			} else {
				properties.textOrientation = toVerticalWithHorizontalFont;
			}
		} else {
			properties.textOrientation = toHorizontal;
		}
	}

	*pValue = properties.textOrientation;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_TextOrientation(TextOrientationConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_TEXTORIENTATION);
	if(properties.textOrientation != newValue) {
		properties.textOrientation = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.textOrientation) {
				case toHorizontal:
					SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_VERTICAL);
					SendMessage(EM_SETEDITSTYLE, 0, SES_USEATFONT);
					break;
				case toVertical:
					SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_VERTICAL);
					SendMessage(EM_SETEDITSTYLE, SES_USEATFONT, SES_USEATFONT);
					break;
				case toVerticalWithHorizontalFont:
					SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_VERTICAL);
					SendMessage(EM_SETEDITSTYLE, 0, SES_USEATFONT);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_TEXTORIENTATION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_TextRange(LONG rangeStart/* = 0*/, LONG rangeEnd/* = -1*/, IRichTextRange** ppTextRange/* = NULL*/)
{
	ATLASSERT_POINTER(ppTextRange, IRichTextRange*);
	if(!ppTextRange) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		CComPtr<ITextRange> pTextRange = NULL;
		if(rangeStart == 0 && rangeEnd == -1) {
			/*hr = properties.pTextDocument->Range(rangeStart, rangeStart, &pTextRange);
			if(SUCCEEDED(hr) && pTextRange) {
				// actually something like tomDocument would be better
				hr = pTextRange->MoveEnd(tomStory, 1, NULL);
			}*/
			rangeEnd = GetWindowTextLength();
			hr = properties.pTextDocument->Range(rangeStart, rangeEnd, &pTextRange);
		} else {
			hr = properties.pTextDocument->Range(rangeStart, rangeEnd, &pTextRange);
		}
		ClassFactory::InitTextRange(pTextRange, this, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppTextRange));
	}
	return hr;
}

STDMETHODIMP RichTextBox::get_TSFModeBias(TSFModeBiasConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TSFModeBiasConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.TSFModeBias = static_cast<TSFModeBiasConstants>(SendMessage(EM_GETCTFMODEBIAS, 0, 0));
	}

	*pValue = properties.TSFModeBias;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_TSFModeBias(TSFModeBiasConstants newValue)
{
	PUTPROPPROLOG(DISPID_RTB_TSFMODEBIAS);
	if(properties.TSFModeBias != newValue) {
		properties.TSFModeBias = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode() && IsWindow()) {
			SendMessage(EM_SETCTFMODEBIAS, properties.TSFModeBias, 0);
		}
		FireOnChanged(DISPID_RTB_TSFMODEBIAS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_UseBkAcetateColorForTextSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useBkAcetateColorForTextSelection = ((SendMessage(EM_GETEDITSTYLEEX, 0, 0) & SES_EX_NOACETATESELECTION) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useBkAcetateColorForTextSelection);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_UseBkAcetateColorForTextSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_USEBKACETATECOLORFORTEXTSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useBkAcetateColorForTextSelection != b) {
		properties.useBkAcetateColorForTextSelection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useBkAcetateColorForTextSelection) {
				SendMessage(EM_SETEDITSTYLEEX, 0, SES_EX_NOACETATESELECTION);
			} else {
				SendMessage(EM_SETEDITSTYLEEX, SES_EX_NOACETATESELECTION, SES_EX_NOACETATESELECTION);
			}
		}
		FireOnChanged(DISPID_RTB_USEBKACETATECOLORFORTEXTSELECTION);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_UseBuiltInSpellChecking(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useBuiltInSpellChecking = ((SendMessage(EM_GETLANGOPTIONS, 0, 0) & IMF_SPELLCHECKING) == IMF_SPELLCHECKING);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useBuiltInSpellChecking);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_UseBuiltInSpellChecking(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_USEBUILTINSPELLCHECKING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useBuiltInSpellChecking != b) {
		properties.useBuiltInSpellChecking = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD options = SendMessage(EM_GETLANGOPTIONS, 0, 0);
			if(properties.useBuiltInSpellChecking) {
				options |= IMF_SPELLCHECKING;
			} else {
				options &= ~IMF_SPELLCHECKING;
			}
			SendMessage(EM_SETLANGOPTIONS, 0, options);
		}
		FireOnChanged(DISPID_RTB_USEBUILTINSPELLCHECKING);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_UseCustomFormattingRectangle(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useCustomFormattingRectangle);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_UseCustomFormattingRectangle(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_USECUSTOMFORMATTINGRECTANGLE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useCustomFormattingRectangle != b) {
		properties.useCustomFormattingRectangle = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useCustomFormattingRectangle) {
				SendMessage(EM_SETRECT, FALSE, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
			} else {
				SendMessage(EM_SETRECT, FALSE, NULL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_USECUSTOMFORMATTINGRECTANGLE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_UseDualFontMode(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useDualFontMode = ((SendMessage(EM_GETLANGOPTIONS, 0, 0) & IMF_DUALFONT) == IMF_DUALFONT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useDualFontMode);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_UseDualFontMode(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_USEDUALFONTMODE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useDualFontMode != b) {
		properties.useDualFontMode = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD options = SendMessage(EM_GETLANGOPTIONS, 0, 0);
			if(properties.useDualFontMode) {
				options |= IMF_DUALFONT;
			} else {
				options &= ~IMF_DUALFONT;
			}
			SendMessage(EM_SETLANGOPTIONS, 0, options);
		}
		FireOnChanged(DISPID_RTB_USEDUALFONTMODE);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_UseSmallerFontForNestedFractions(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			if(SUCCEEDED(pTextDocument2->GetMathProperties(&options))) {
				properties.useSmallerFontForNestedFractions = ((options & tomMathDispFracTeX) == tomMathDispFracTeX);
			}
		}
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSmallerFontForNestedFractions);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_UseSmallerFontForNestedFractions(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_USESMALLERFONTFORNESTEDFRACTIONS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSmallerFontForNestedFractions != b) {
		properties.useSmallerFontForNestedFractions = b;
		SetDirty(TRUE);

		if(properties.pTextDocument) {
			CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
			if(pTextDocument2) {
				ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.useSmallerFontForNestedFractions ? tomMathDispFracTeX : 0, tomMathDispFracTeX)));
			}
		}
		FireOnChanged(DISPID_RTB_USESMALLERFONTFORNESTEDFRACTIONS);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_UseTextServicesFramework(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useTextServicesFramework = ((SendMessage(EM_GETEDITSTYLE, 0, 0) & SES_USECTF) == SES_USECTF);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useTextServicesFramework);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_UseTextServicesFramework(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_USETEXTSERVICESFRAMEWORK);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useTextServicesFramework != b) {
		properties.useTextServicesFramework = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useTextServicesFramework) {
				SendMessage(EM_SETEDITSTYLE, SES_USECTF, SES_USECTF);
			} else {
				SendMessage(EM_SETEDITSTYLE, 0, SES_USECTF);
			}
		}
		FireOnChanged(DISPID_RTB_USETEXTSERVICESFRAMEWORK);
	}
	return S_OK;
}
/*
STDMETHODIMP RichTextBox::get_UseTouchKeyboardAutoCorrection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useTouchKeyboardAutoCorrection = ((SendMessage(EM_GETLANGOPTIONS, 0, 0) & IMF_TKBAUTOCORRECTION) == IMF_TKBAUTOCORRECTION);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useTouchKeyboardAutoCorrection);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_UseTouchKeyboardAutoCorrection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_USETOUCHKEYBOARDAUTOCORRECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useTouchKeyboardAutoCorrection != b) {
		properties.useTouchKeyboardAutoCorrection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD options = SendMessage(EM_GETLANGOPTIONS, 0, 0);
			if(properties.useTouchKeyboardAutoCorrection) {
				options |= IMF_TKBAUTOCORRECTION;
			} else {
				options &= ~IMF_TKBAUTOCORRECTION;
			}
			SendMessage(EM_SETLANGOPTIONS, 0, options);
		}
		FireOnChanged(DISPID_RTB_USETOUCHKEYBOARDAUTOCORRECTION);
	}
	return S_OK;
}*/

STDMETHODIMP RichTextBox::get_UseWindowsThemes(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useWindowsThemes = ((SendMessage(EM_GETEDITSTYLEEX, 0, 0) & SES_EX_NOTHEMING) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useWindowsThemes);
	return S_OK;
}

STDMETHODIMP RichTextBox::put_UseWindowsThemes(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RTB_USEWINDOWSTHEMES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useWindowsThemes != b) {
		properties.useWindowsThemes = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
			/*if(properties.useWindowsThemes) {
				SendMessage(EM_SETEDITSTYLEEX, 0, SES_EX_NOTHEMING);
			} else {
				SendMessage(EM_SETEDITSTYLEEX, SES_EX_NOTHEMING, SES_EX_NOTHEMING);
			}
			SetWindowPos(NULL, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_DRAWFRAME | SWP_FRAMECHANGED);*/
		}
		FireOnChanged(DISPID_RTB_USEWINDOWSTHEMES);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP RichTextBox::get_WordBreakFunction(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<LONG>(SendMessage(EM_GETWORDBREAKPROC, 0, 0));
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::put_WordBreakFunction(LONG newValue)
{
	PUTPROPPROLOG(DISPID_RTB_WORDBREAKFUNCTION);
	if(IsWindow()) {
		SendMessage(EM_SETWORDBREAKPROC, 0, static_cast<LPARAM>(newValue));
	}

	FireOnChanged(DISPID_RTB_WORDBREAKFUNCTION);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP RichTextBox::get_ZoomRatio(DOUBLE* pValue)
{
	ATLASSERT_POINTER(pValue, DOUBLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		int numerator = 0;
		int denominator = 0;
		if(SendMessage(EM_GETZOOM, reinterpret_cast<WPARAM>(&numerator), reinterpret_cast<LPARAM>(&denominator))) {
			if(numerator == 0 && denominator == 0) {
				properties.zoomRatio = 0.0;
			} else if(denominator != 0) {
				properties.zoomRatio = static_cast<DOUBLE>(numerator) / static_cast<DOUBLE>(denominator);
			}
		}
	}

	*pValue = properties.zoomRatio;
	return S_OK;
}

STDMETHODIMP RichTextBox::put_ZoomRatio(DOUBLE newValue)
{
	PUTPROPPROLOG(DISPID_RTB_ZOOMRATIO);
	if(newValue >= -0.0000001 && newValue <= 0.0000001) {
		newValue = 0.0;
	}
	if(newValue != 0.0 && (newValue <= (1.0 / 64.0) || newValue >= 64.0)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.zoomRatio != newValue) {
		properties.zoomRatio = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			int numerator = 0;
			int denominator = 0;
			if(properties.zoomRatio != 0.0) {
				numerator = static_cast<int>(properties.zoomRatio);
				denominator = 1;
				DOUBLE ratio = static_cast<DOUBLE>(numerator) / static_cast<DOUBLE>(denominator);
				while((ratio != properties.zoomRatio) && numerator < 1000000000 && denominator < 1000000000) {
					denominator *= 10;
					numerator = static_cast<int>(properties.zoomRatio * static_cast<DOUBLE>(denominator));
					ratio = static_cast<DOUBLE>(numerator) / static_cast<DOUBLE>(denominator);
				}
			}
			SendMessage(EM_SETZOOM, numerator, denominator);
			FireViewChange();
		}
		FireOnChanged(DISPID_RTB_ZOOMRATIO);
	}
	return S_OK;
}

STDMETHODIMP RichTextBox::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP RichTextBox::BeginNewUndoAction(void)
{
	if(IsWindow()) {
		SendMessage(EM_STOPGROUPTYPING, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::CanCopy(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	BOOL useFallback = TRUE;
	/*if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long value = 0;
			// Always returns tomTrue?!
			useFallback = FAILED(pTextDocument2->GetProperty(tomCanCopy, &value));
			if(!useFallback) {
				*pValue = BOOL2VARIANTBOOL(value);
				return S_OK;
			}
		}
	}*/
	if(useFallback && IsWindow()) {
		if(properties.pTextDocument) {
			CComPtr<ITextSelection> pTextSelection = NULL;
			HRESULT hr = properties.pTextDocument->GetSelection(&pTextSelection);
			if(SUCCEEDED(hr)) {
				LONG start = 0;
				LONG end = 0;
				hr = pTextSelection->GetStart(&start);
				if(SUCCEEDED(hr)) {
					hr = pTextSelection->GetEnd(&end);
					if(SUCCEEDED(hr)) {
						*pValue = BOOL2VARIANTBOOL(end - start > 0);
					}
				}
			}
			return hr;
		}
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::CanRedo(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	BOOL useFallback = TRUE;
	if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long value = 0;
			useFallback = FAILED(pTextDocument2->GetProperty(tomCanRedo, &value));
			if(!useFallback) {
				*pValue = BOOL2VARIANTBOOL(value);
				return S_OK;
			}
		}
	}
	if(useFallback && IsWindow()) {
		*pValue = BOOL2VARIANTBOOL(SendMessage(EM_CANREDO, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::CanUndo(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	BOOL useFallback = TRUE;
	if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long value = 0;
			useFallback = FAILED(pTextDocument2->GetProperty(tomCanUndo, &value));
			if(!useFallback) {
				*pValue = BOOL2VARIANTBOOL(value);
				return S_OK;
			}
		}
	}
	if(useFallback && IsWindow()) {
		*pValue = BOOL2VARIANTBOOL(SendMessage(EM_CANUNDO, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}
//
//STDMETHODIMP RichTextBox::CharIndexToPosition(LONG characterIndex, OLE_XPOS_PIXELS* pX/* = NULL*/, OLE_YPOS_PIXELS* pY/* = NULL*/)
//{
//	if(IsWindow()) {
//		POINTL pt = {0};
//		LRESULT lr = SendMessage(EM_POSFROMCHAR, reinterpret_cast<WPARAM>(&pt), characterIndex);
//		if(lr != -1) {
//			if(pX) {
//				*pX = pt.x;
//			}
//			if(pY) {
//				*pY = pt.y;
//			}
//			return S_OK;
//		}
//	}
//	return E_FAIL;
//}

STDMETHODIMP RichTextBox::CheckForOptimalControlWindowSize(void)
{
	if(IsWindow()) {
		SendMessage(EM_REQUESTRESIZE, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::CloseTSFKeyboard(void)
{
	if(IsWindow()) {
		SendMessage(EM_SETCTFOPENSTATUS, FALSE, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::CreateNewDocument(void)
{
	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		/* ITextDocument::New() saves the current document automatically, but as plain text.
		   So always prevent auto-saving. */
		properties.pTextDocument->SetSaved(tomTrue);
		hr = properties.pTextDocument->New();
	}

	return hr;
}

STDMETHODIMP RichTextBox::EmptyUndoBuffer(void)
{
	/*if(IsWindow()) {
		SendMessage(EM_EMPTYUNDOBUFFER, 0, 0);
		return S_OK;
	}*/
	if(properties.pTextDocument) {
		HRESULT hr = properties.pTextDocument->Undo(tomFalse, NULL);
		if(SUCCEEDED(hr)) {
			hr = properties.pTextDocument->Undo(tomTrue, NULL);
		}
		return hr;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP RichTextBox::Freeze(LONG* pNewCount)
{
	ATLASSERT_POINTER(pNewCount, LONG);
	if(!pNewCount) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		hr = properties.pTextDocument->Freeze(pNewCount);
	}

	return hr;
}

STDMETHODIMP RichTextBox::GetCurrentIMECompositionText(IMECompositionTextType textType/* = imecttFinalComposedString*/, BSTR* pCompositionText/* = NULL*/)
{
	ATLASSERT_POINTER(pCompositionText, BSTR);
	if(!pCompositionText) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		IMECOMPTEXT params = {0};
		params.cb = 400;
		params.flags = textType;
		WCHAR pBuffer[200];
		ZeroMemory(pBuffer, 200);
		if(SendMessage(EM_GETIMECOMPTEXT, reinterpret_cast<WPARAM>(&params), reinterpret_cast<LPARAM>(pBuffer))) {
			*pCompositionText = _bstr_t(pBuffer).Detach();
		} else {
			*pCompositionText = NULL;
		}
		hr = S_OK;
	}

	return hr;
}

//STDMETHODIMP RichTextBox::GetFirstCharOfLine(LONG lineIndex, LONG* pValue)
//{
//	ATLASSERT_POINTER(pValue, LONG);
//	if(!pValue) {
//		return E_POINTER;
//	}
//
//	if(IsWindow()) {
//		*pValue = static_cast<LONG>(SendMessage(EM_LINEINDEX, lineIndex, 0));
//		// alternative implementation (lineIndex being one-based!):
//		// get_TextRange(0, -1) -> MoveToUnitIndex(uLine, lineIndex + 1) -> get_RangeStart
//		return S_OK;
//	}
//	return E_FAIL;
//}
//
//STDMETHODIMP RichTextBox::GetLine(LONG lineIndex, BSTR* pValue)
//{
//	ATLASSERT_POINTER(pValue, BSTR);
//	if(!pValue) {
//		return E_POINTER;
//	}
//
//	HRESULT hr = E_FAIL;
//	// NOTE: EM_LINELENGTH doesn't work as expected, e. g. it may return garbage if ES_CENTER is set.
//	if(properties.pTextDocument) {
//		*pValue = NULL;
//		LONG lineStart = static_cast<LONG>(SendMessage(EM_LINEINDEX, lineIndex, 0));
//		if(lineStart >= 0) {
//			CComPtr<ITextRange> pTextRange = NULL;
//			hr = properties.pTextDocument->Range(lineStart, lineStart, &pTextRange);
//			if(SUCCEEDED(hr) && pTextRange) {
//				LONG delta = -1;
//				hr = pTextRange->EndOf(tomLine, tomExtend, &delta);
//				if(SUCCEEDED(hr)) {
//					// lineEnd always includes the CR, so move back by one character
//					hr = pTextRange->MoveEnd(tomCharacter, -1, &delta);
//					if(SUCCEEDED(hr)) {
//						hr = pTextRange->GetText(pValue);
//					}
//				}
//			}
//		} else {
//			// illegal line index, let's stay compatible to EM_LINELENGTH and not fail
//			hr = S_OK;
//		}
//		if(!*pValue) {
//			*pValue = SysAllocStringLen(NULL, 0);
//		}
//	}
//	return hr;
//}

STDMETHODIMP RichTextBox::GetLineCount(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = static_cast<LONG>(SendMessage(EM_GETLINECOUNT, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}
//
//STDMETHODIMP RichTextBox::GetLineFromChar(LONG characterIndex, LONG* pValue)
//{
//	ATLASSERT_POINTER(pValue, LONG);
//	if(!pValue) {
//		return E_POINTER;
//	}
//
//	if(IsWindow()) {
//		*pValue = static_cast<LONG>(SendMessage(EM_EXLINEFROMCHAR, 0, characterIndex));
//		// alternative implementation (lineIndex being one-based!):
//		// get_TextRange(characterIndex, characterIndex) -> get_UnitIndex(uLine)
//		return S_OK;
//	}
//	return E_FAIL;
//}
//
//STDMETHODIMP RichTextBox::GetLineLength(LONG lineIndex, LONG* pValue)
//{
//	ATLASSERT_POINTER(pValue, LONG);
//	if(!pValue) {
//		return E_POINTER;
//	}
//
//	HRESULT hr = E_FAIL;
//	// NOTE: EM_LINELENGTH doesn't work as expected, e. g. it may return garbage if ES_CENTER is set.
//	if(properties.pTextDocument) {
//		*pValue = 0;
//		LONG lineStart = static_cast<LONG>(SendMessage(EM_LINEINDEX, lineIndex, 0));
//		if(lineStart >= 0) {
//			CComPtr<ITextRange> pTextRange = NULL;
//			hr = properties.pTextDocument->Range(lineStart, lineStart, &pTextRange);
//			if(SUCCEEDED(hr) && pTextRange) {
//				LONG delta = -1;
//				hr = pTextRange->EndOf(tomLine, tomExtend, &delta);
//				if(SUCCEEDED(hr)) {
//					LONG lineEnd = 0;
//					hr = pTextRange->GetEnd(&lineEnd);
//					if(SUCCEEDED(hr)) {
//						// lineEnd always includes the CR
//						*pValue = lineEnd - 1 - lineStart;
//					}
//				}
//			}
//		} else {
//			// illegal line index, let's stay compatible to EM_LINELENGTH and not fail
//			hr = S_OK;
//		}
//	}
//	return hr;
//}
//
//STDMETHODIMP RichTextBox::GetSelection(LONG* pSelectionStart/* = NULL*/, LONG* pSelectionEnd/* = NULL*/)
//{
//	if(IsWindow()) {
//		CHARRANGE range;
//		SendMessage(EM_EXGETSEL, 0, reinterpret_cast<LPARAM>(&range));
//		if(pSelectionStart) {
//			*pSelectionStart = range.cpMin;
//		}
//		if(pSelectionEnd) {
//			*pSelectionEnd = range.cpMax;
//		}
//		return S_OK;
//	}
//	return E_FAIL;
//}

STDMETHODIMP RichTextBox::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IRichTextRange** ppHitRange)
{
	ATLASSERT_POINTER(ppHitRange, IRichTextRange*);
	if(!ppHitRange) {
		return E_POINTER;
	}

	if(IsWindow() && properties.pTextDocument) {
		*pHitTestDetails = HitTest(x, y);
		if(*pHitTestDetails & (htAbove | htToRight | htBelow | htToLeft)) {
			*ppHitRange = NULL;
		} else {
			CComPtr<ITextRange> pRange;
			POINT screenMousePosition = {x, y};
			ClientToScreen(&screenMousePosition);
			properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);

			ClassFactory::InitTextRange(pRange, this, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppHitRange));
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::LoadFromFile(BSTR file, FileCreationDispositionConstants fileCreationDisposition/* = fcdOpenExisting*/, FileTypeConstants fileType/* = ftyAutoDetect*/, FileAccessOptionConstants fileAccessOptions/* = faoDefault*/, LONG codePage/* = 0*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		VARIANT varFile;
		VariantInit(&varFile);
		varFile.vt = VT_BSTR;
		varFile.bstrVal = file;

		long flags = fileCreationDisposition | fileType | fileAccessOptions;
		hr = properties.pTextDocument->Open(&varFile, flags, codePage);
		*pSucceeded = BOOL2VARIANTBOOL(SUCCEEDED(hr));
		hr = S_OK;

		VariantClear(&varFile);
	}

	return hr;
}

STDMETHODIMP RichTextBox::OLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, IRichTextRange* pDraggedTextRange/* = NULL*/, LONG itemCountToDisplay/* = 0*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
{
	if(supportedEffects == odeNone) {
		// don't waste time
		return S_OK;
	}
	if(hWndToAskForDragImage == -1) {
		ATLASSERT(pDraggedTextRange);
		if(!pDraggedTextRange) {
			return E_INVALIDARG;
		}
	}
	ATLASSERT_POINTER(pDataObject, LONG);
	LONG dummy = NULL;
	if(!pDataObject) {
		pDataObject = &dummy;
	}
	ATLASSERT_POINTER(pPerformedEffects, OLEDropEffectConstants);
	OLEDropEffectConstants performedEffect = odeNone;
	if(!pPerformedEffects) {
		pPerformedEffects = &performedEffect;
	}

	HWND hWnd = NULL;
	if(hWndToAskForDragImage == -1) {
		hWnd = m_hWnd;
	} else {
		hWnd = static_cast<HWND>(LongToHandle(hWndToAskForDragImage));
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IDataObject> pDataObjectToUse = NULL;
	if(*pDataObject) {
		pDataObjectToUse = *reinterpret_cast<LPDATAOBJECT*>(pDataObject);

		CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->Attach(pDataObjectToUse);
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->Release();
	} else {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->SetOwner(this);
		if(itemCountToDisplay >= 0) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				pOLEDataObjectObj->properties.numberOfItemsToDisplay = itemCountToDisplay;
			}
		}
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&pDataObjectToUse));
		pOLEDataObjectObj->Release();
	}
	ATLASSUME(pDataObjectToUse);
	pDataObjectToUse->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
	CComQIPtr<IDropSource, &IID_IDropSource> pDragSource(this);

	if(pDraggedTextRange) {
		CComQIPtr<ITextRange> pDraggedRange = pDraggedTextRange;
		dragDropStatus.pDraggedTextRange = pDraggedRange;
	}
	POINT mousePosition = {0};
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);

	if(properties.supportOLEDragImages) {
		IDragSourceHelper* pDragSourceHelper = NULL;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDragSourceHelper));
		if(pDragSourceHelper) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				IDragSourceHelper2* pDragSourceHelper2 = NULL;
				pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&pDragSourceHelper2));
				if(pDragSourceHelper2) {
					pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
					// this was the only place we actually use IDragSourceHelper2
					pDragSourceHelper->Release();
					pDragSourceHelper = static_cast<IDragSourceHelper*>(pDragSourceHelper2);
				}
			}

			HRESULT hr = pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
			if(FAILED(hr)) {
				/* This happens if full window dragging is deactivated. Actually, InitializeFromWindow() contains a
				   fallback mechanism for this case. This mechanism retrieves the passed window's class name and
				   builds the drag image using TVM_CREATEDRAGIMAGE if it's SysTreeView32, LVM_CREATEDRAGIMAGE if
				   it's SysListView32 and so on. Our class name is RichTextBox[U|A]_RICHEDIT[5|6]0W, so we're doomed.
				   So how can we have drag images anyway? Well, we use a very ugly hack: We temporarily activate
				   full window dragging. */
				BOOL fullWindowDragging;
				SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &fullWindowDragging, 0);
				if(!fullWindowDragging) {
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, TRUE, NULL, 0);
					pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FALSE, NULL, 0);
				}
			}

			if(pDragSourceHelper) {
				pDragSourceHelper->Release();
			}
		}
	}

	if(IsLeftMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_LBUTTON;
	} else if(IsRightMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_RBUTTON;
	}
	if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
		dragDropStatus.useItemCountLabelHack = TRUE;
	}

	if(pOLEDataObject) {
		Raise_OLEStartDrag(pOLEDataObject);
	}
	HRESULT hr = DoDragDrop(pDataObjectToUse, pDragSource, supportedEffects, reinterpret_cast<LPDWORD>(pPerformedEffects));
	KillTimer(timers.ID_DRAGSCROLL);
	if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
		Raise_OLECompleteDrag(pOLEDataObject, *pPerformedEffects);
	}

	dragDropStatus.Reset();
	return S_OK;
}

STDMETHODIMP RichTextBox::OpenTSFKeyboard(void)
{
	if(IsWindow()) {
		SendMessage(EM_SETCTFOPENSTATUS, TRUE, 0);
		return S_OK;
	}
	return E_FAIL;
}
//
//STDMETHODIMP RichTextBox::PositionToCharIndex(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG* pCharacterIndex/* = NULL*/, LONG* pLineIndex/* = NULL*/)
//{
//	if(IsWindow()) {
//		if(pCharacterIndex) {
//			*pCharacterIndex = -1;
//		}
//		if(pLineIndex) {
//			*pLineIndex = -1;
//		}
//		POINTL pt = {x, y};
//		LRESULT lr = SendMessage(EM_CHARFROMPOS, 0, reinterpret_cast<LPARAM>(&pt));
//		if((LOWORD(lr) != 65535) || (HIWORD(lr) != 65535)) {
//			if(pCharacterIndex) {
//				*pCharacterIndex = lr;
//			}
//			if(pLineIndex) {
//				*pLineIndex = static_cast<LONG>(SendMessage(EM_EXLINEFROMCHAR, 0, lr));
//			}
//		}
//		return S_OK;
//	}
//	return E_FAIL;
//}

STDMETHODIMP RichTextBox::Redo(LONG numberOfActions/* = 1*/, LONG* pCount/* = NULL*/)
{
	ATLASSERT_POINTER(pCount, LONG);
	if(!pCount) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		hr = properties.pTextDocument->Redo(numberOfActions, pCount);
	}
	/*if(IsWindow()) {
		*pSucceeded = BOOL2VARIANTBOOL(SendMessage(EM_REDO, 0, 0));
		hr = S_OK;
	}*/
	return hr;
}

STDMETHODIMP RichTextBox::Refresh(void)
{
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		Invalidate();
		UpdateWindow();
		dragDropStatus.ShowDragImage(TRUE);
	}
	return S_OK;
}
//
//STDMETHODIMP RichTextBox::ReplaceSelectedText(BSTR replacementText, VARIANT_BOOL undoable/* = VARIANT_FALSE*/)
//{
//	if(IsWindow()) {
//		SendMessage(EM_REPLACESEL, VARIANTBOOL2BOOL(undoable), reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(COLE2CT(replacementText))));
//		return S_OK;
//	}
//	return E_FAIL;
//}

STDMETHODIMP RichTextBox::SaveToFile(BSTR file/* = NULL*/, FileCreationDispositionConstants fileCreationDisposition/* = static_cast<FileCreationDispositionConstants>(0)*/, FileTypeConstants fileType/* = ftyAutoDetect*/, FileAccessOptionConstants fileAccessOptions/* = faoDefault*/, LONG codePage/* = 0*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		VARIANT varFile;
		VariantInit(&varFile);
		varFile.vt = VT_BSTR;

		BOOL hasFileName = FALSE;
		if(file) {
			hasFileName = (lstrlenW(OLE2W(file)) > 0);
		}
		if(hasFileName) {
			varFile.bstrVal = file;
		} else {
			properties.pTextDocument->GetName(&varFile.bstrVal);
		}

		/* NOTE: According to the documentation, calling ITextDocument::Save with all parameters set to 0,
		         saves the document with the options it has been opened with. However, this seems to be true
						 only for the file name. If you really call the Save method with (NULL, 0, 0), you'll get
						 an ANSI text file even if you loaded a Unicode rtf file. */
		/*long flags = fileCreationDisposition | fileType | fileAccessOptions;
		if(!hasFileName && flags == 0 && codePage == 0) {
			if(fileType == ftyAutoDetect) {
				LPWSTR pExtension = PathFindExtensionW(OLE2W(varFile.bstrVal));
				if(lstrcmpiW(pExtension, L".rtf") == 0) {
					fileType = ftyRTF;
				} else if(lstrcmpiW(pExtension, L".txt") == 0) {
					fileType = ftyText;
				} else {
					fileType = ftyRTF;
				}
			}

			flags = fileCreationDisposition | fileType | fileAccessOptions;
			hr = properties.pTextDocument->Save(NULL, flags, 0);
		} else {*/
			if(fileType == ftyAutoDetect) {
				LPWSTR pExtension = PathFindExtensionW(OLE2W(varFile.bstrVal));
				if(lstrcmpiW(pExtension, L".rtf") == 0) {
					fileType = ftyRTF;
				} else if(lstrcmpiW(pExtension, L".txt") == 0) {
					fileType = ftyText;
				} else {
					fileType = ftyRTF;
				}
			}

			long flags = fileCreationDisposition | fileType | fileAccessOptions;
			hr = properties.pTextDocument->Save(&varFile, flags, codePage);
		//}
		VariantClear(&varFile);
		*pSucceeded = BOOL2VARIANTBOOL(SUCCEEDED(hr));
		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP RichTextBox::Scroll(ScrollAxisConstants axis, ScrollDirectionConstants directionAndIntensity, LONG linesToScrollVertically/* = 0*/, LONG charactersToScrollHorizontally/* = 0*/)
{
	if(IsWindow()) {
		if(directionAndIntensity == sdCustom) {
			SendMessage(EM_LINESCROLL, ((axis & saHorizontal) == saHorizontal) ? charactersToScrollHorizontally : 0, ((axis & saVertical) == saVertical) ? linesToScrollVertically : 0);
		} else {
			if(axis & saVertical) {
				SendMessage(EM_SCROLL, directionAndIntensity, 0);
			}
			if(axis & saHorizontal) {
				SendMessage(WM_HSCROLL, directionAndIntensity, 0);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

//STDMETHODIMP RichTextBox::ScrollCaretIntoView(void)
//{
//	if(IsWindow()) {
//		SendMessage(EM_SCROLLCARET, 0, 0);
//		return S_OK;
//	}
//	return E_FAIL;
//}

//STDMETHODIMP RichTextBox::SetSelection(LONG selectionStart, LONG selectionEnd)
//{
//	if(IsWindow()) {
//		CHARRANGE range = {selectionStart, selectionEnd};
//		SendMessage(EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&range));
//		return S_OK;
//	}
//	return E_FAIL;
//}

STDMETHODIMP RichTextBox::SetTargetDevice(OLE_HANDLE hDC, LONG lineWidth)
{
	if(IsWindow()) {
		SendMessage(EM_SETTARGETDEVICE, hDC, lineWidth);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::StartIMEReconversion(void)
{
	if(IsWindow()) {
		SendMessage(EM_RECONVERSION, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP RichTextBox::Undo(LONG numberOfActions/* = 1*/, LONG* pCount/* = NULL*/)
{
	ATLASSERT_POINTER(pCount, LONG);
	if(!pCount) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		hr = properties.pTextDocument->Undo(numberOfActions, pCount);
	}
	/*if(IsWindow()) {
		*pSucceeded = BOOL2VARIANTBOOL(SendMessage(EM_UNDO, 0, 0));
		hr = S_OK;
	}*/
	return hr;
}

STDMETHODIMP RichTextBox::Unfreeze(LONG* pNewCount)
{
	ATLASSERT_POINTER(pNewCount, LONG);
	if(!pNewCount) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextDocument) {
		hr = properties.pTextDocument->Unfreeze(pNewCount);
	}

	return hr;
}


LRESULT RichTextBox::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT RichTextBox::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	previousSelectedRange.cpMin = -1;
	previousSelectedRange.cpMax = -1;
	nextEmbeddedObjectIndex = 0;
	ATLASSERT(SUCCEEDED(StgCreateStorageEx(NULL, STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DELETEONRELEASE, STGFMT_STORAGE, 0, NULL, NULL, IID_IStorage, reinterpret_cast<LPVOID*>(&pDocumentStorage))));

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(SendMessage(EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			pRichEditOle->QueryInterface(IID_PPV_ARGS(&properties.pTextDocument));
			/*CComPtr<ITextServices> pTextServices = NULL;
			pRichEditOle->QueryInterface(IID_ITextServices, reinterpret_cast<LPVOID*>(&pTextServices));
			if(pTextServices) {
			}*/
		}
		Raise_RecreatedControlWindow(*this);
	}
	return lr;
}

LRESULT RichTextBox::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(properties.pTextDocument) {
		// this will prevent MS Office Rich Edit from destroying opened files
		properties.pTextDocument->SetSaved(tomTrue);
		properties.pTextDocument->Release();
		properties.pTextDocument = NULL;
	}
	Raise_DestroyedControlWindow(*this);

	if(pDocumentStorage) {
		pDocumentStorage->Release();
		pDocumentStorage = NULL;
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT RichTextBox::OnEraseBkGnd(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(properties.useCustomFormattingRectangle) {
		SendMessage(EM_SETRECTNP, FALSE, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	}

	return lr;
}

LRESULT RichTextBox::OnGetDlgCode(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(properties.acceptTabKey) {
		lr |= DLGC_WANTTAB;
	} else {
		lr &= ~DLGC_WANTTAB;
	}
	return lr;
}

LRESULT RichTextBox::OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if((ime != imeNoControl) && (GetFocus() == *this)) {
		// we've the focus, so configure the IME
		SetCurrentIMEContextMode(*this, ime);
	}
	return lr;
}

LRESULT RichTextBox::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT RichTextBox::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT RichTextBox::OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT RichTextBox::OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = 0;
	if(properties.detectDragDrop && properties.pTextDocument) {
		CComPtr<ITextSelection> pTextSelection = NULL;
		properties.pTextDocument->GetSelection(&pTextSelection);
		long selectionStart = 0;
		pTextSelection->GetStart(&selectionStart);
		long selectionEnd = 0;
		pTextSelection->GetEnd(&selectionEnd);
		if(abs(selectionEnd - selectionStart) > 0) {
			// we've a selection - check whether the mouse cursor is over selected text
			if(IsOverSelectedText(lParam)) {
				dragDropStatus.candidate.pTextRange = pTextSelection;
				dragDropStatus.candidate.button = MK_LBUTTON;
				dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
				dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);

				dragDropStatus.candidate.bufferedLButtonDown.wParam = wParam;
				dragDropStatus.candidate.bufferedLButtonDown.lParam = lParam;

				SetCapture();
			} else {
				lr = DefWindowProc(message, wParam, lParam);
			}
		} else {
			// there's no selected text
			lr = DefWindowProc(message, wParam, lParam);
		}
	} else {
		lr = DefWindowProc(message, wParam, lParam);
	}

	return lr;
}

LRESULT RichTextBox::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.detectDragDrop) {
		if(dragDropStatus.candidate.button == MK_LBUTTON) {
			if((dragDropStatus.candidate.bufferedLButtonDown.wParam != 0) && (dragDropStatus.candidate.bufferedLButtonDown.lParam != 0)) {
				DefWindowProc(WM_LBUTTONDOWN, dragDropStatus.candidate.bufferedLButtonDown.wParam, dragDropStatus.candidate.bufferedLButtonDown.lParam);
			}
			dragDropStatus.candidate.Reset();
		}
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT RichTextBox::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT RichTextBox::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT RichTextBox::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT RichTextBox::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};

	CComPtr<ITextRange> pRange = NULL;
	if(properties.pTextDocument) {
		POINT screenMousePosition = mousePosition;
		ClientToScreen(&screenMousePosition);
		properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
	}
	CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);

	Raise_MouseHover(pTextRange, button, shift, mousePosition.x, mousePosition.y, HitTest(mousePosition.x, mousePosition.y));

	return 0;
}

LRESULT RichTextBox::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	CComPtr<ITextRange> pRange = NULL;
	if(properties.pTextDocument) {
		POINT screenMousePosition = mouseStatus.lastPosition;
		ClientToScreen(&screenMousePosition);
		properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
	}
	CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);

	Raise_MouseLeave(pTextRange, button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y, HitTest(mouseStatus.lastPosition.x, mouseStatus.lastPosition.y));

	return 0;
}

LRESULT RichTextBox::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	if(properties.detectDragDrop) {
		if(dragDropStatus.candidate.pTextRange) {
			// calculate the rectangle, that the mouse cursor must leave to start dragging
			int clickRectWidth = GetSystemMetrics(SM_CXDRAG);
			if(clickRectWidth < 4) {
				clickRectWidth = 4;
			}
			int clickRectHeight = GetSystemMetrics(SM_CYDRAG);
			if(clickRectHeight < 4) {
				clickRectHeight = 4;
			}
			WTL::CRect rc(dragDropStatus.candidate.position.x - clickRectWidth, dragDropStatus.candidate.position.y - clickRectHeight, dragDropStatus.candidate.position.x + clickRectWidth, dragDropStatus.candidate.position.y + clickRectHeight);

			if(!rc.PtInRect(WTL::CPoint(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)))) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(dragDropStatus.candidate.pTextRange, this);
				HitTestConstants hitTestDetails = HitTest(dragDropStatus.candidate.position.x, dragDropStatus.candidate.position.y);
				switch(dragDropStatus.candidate.button) {
					case MK_LBUTTON:
						Raise_BeginDrag(pTextRange, button, shift, dragDropStatus.candidate.position.x, dragDropStatus.candidate.position.y, hitTestDetails);
						break;
					case MK_RBUTTON:
						Raise_BeginRDrag(pTextRange, button, shift, dragDropStatus.candidate.position.x, dragDropStatus.candidate.position.y, hitTestDetails);
						break;
				}
				dragDropStatus.candidate.pTextRange = NULL;
			}
		}
	}

	return 0;
}

LRESULT RichTextBox::OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT RichTextBox::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT RichTextBox::OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT RichTextBox::OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = 0;
	if(properties.detectDragDrop && properties.pTextDocument) {
		CComPtr<ITextSelection> pTextSelection = NULL;
		properties.pTextDocument->GetSelection(&pTextSelection);
		long selectionStart = 0;
		pTextSelection->GetStart(&selectionStart);
		long selectionEnd = 0;
		pTextSelection->GetEnd(&selectionEnd);
		if(abs(selectionEnd - selectionStart) > 0) {
			// we've a selection - check whether the mouse cursor is over selected text
			if(IsOverSelectedText(lParam)) {
				dragDropStatus.candidate.pTextRange = pTextSelection;
				dragDropStatus.candidate.button = MK_RBUTTON;
				dragDropStatus.candidate.position.x = GET_X_LPARAM(lParam);
				dragDropStatus.candidate.position.y = GET_Y_LPARAM(lParam);

				dragDropStatus.candidate.bufferedRButtonDown.wParam = wParam;
				dragDropStatus.candidate.bufferedRButtonDown.lParam = lParam;

				SetCapture();
			} else {
				lr = DefWindowProc(message, wParam, lParam);
			}
		} else {
			// there's no selected text
			lr = DefWindowProc(message, wParam, lParam);
		}
	} else {
		lr = DefWindowProc(message, wParam, lParam);
	}

	return lr;
}

LRESULT RichTextBox::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.detectDragDrop) {
		if(dragDropStatus.candidate.button == MK_RBUTTON) {
			if((dragDropStatus.candidate.bufferedRButtonDown.wParam != 0) && (dragDropStatus.candidate.bufferedRButtonDown.lParam != 0)) {
				DefWindowProc(WM_RBUTTONDOWN, dragDropStatus.candidate.bufferedRButtonDown.wParam, dragDropStatus.candidate.bufferedRButtonDown.lParam);
			}
			dragDropStatus.candidate.Reset();
		}
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT RichTextBox::OnScroll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!(properties.disabledEvents & deScrolling)) {
		if((LOWORD(wParam) == SB_THUMBPOSITION) || (LOWORD(wParam) == SB_THUMBTRACK)) {
			switch(message) {
				case WM_HSCROLL:
					Raise_Scrolling(saHorizontal);
					break;
				case WM_VSCROLL:
					Raise_Scrolling(saVertical);
					break;
			}
		}
	}

	return lr;
}

LRESULT RichTextBox::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	WTL::CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	ScreenToClient(&mousePosition);
	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			BOOL isOverLink = (HitTest(mousePosition.x, mousePosition.y) & htLink);
			if(!isOverLink) {
				SetCursor(hCursor);
				return TRUE;
			}
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT RichTextBox::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_RTB_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_RTB_FONT);
	}

	return lr;
}

LRESULT RichTextBox::OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_RTB_TEXT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_TEXT);
	SendOnDataChange();
	return lr;
}

LRESULT RichTextBox::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT RichTextBox::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_DRAGSCROLL:
			AutoScroll();
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT RichTextBox::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	WTL::CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull()/* && pDetails->x != 0 && pDetails->y != 0*/)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		Invalidate();
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT RichTextBox::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT RichTextBox::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT RichTextBox::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT RichTextBox::OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL succeeded = FALSE;
	BOOL useVistaDragImage = FALSE;
	if(dragDropStatus.pDraggedTextRange) {
		if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
			succeeded = TRUE; //CreateVistaOLEDragImage(dragDropStatus.draggedTextFirstChar, dragDropStatus.draggedTextLastChar, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
			useVistaDragImage = succeeded;
		}
		if(!succeeded) {
			// use a legacy drag image as fallback
			succeeded = CreateLegacyOLEDragImage(dragDropStatus.pDraggedTextRange, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
		}

		if(succeeded && RunTimeHelper::IsVista()) {
			FORMATETC format = {0};
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			STGMEDIUM medium = {0};
			medium.tymed = TYMED_HGLOBAL;
			medium.hGlobal = GlobalAlloc(GPTR, sizeof(BOOL));
			if(medium.hGlobal) {
				LPBOOL pUseVistaDragImage = reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				*pUseVistaDragImage = useVistaDragImage;
				GlobalUnlock(medium.hGlobal);

				dragDropStatus.pSourceDataObject->SetData(&format, &medium, TRUE);
			}
		}
	}

	wasHandled = succeeded;
	// TODO: Why do we have to return FALSE to have the set offset not ignored if a Vista drag image is used?
	return succeeded && !useVistaDragImage;
}

LRESULT RichTextBox::OnSetBkgndColor(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(wParam) {
		properties.backColor = 0x80000000 | COLOR_WINDOW;
	} else if(static_cast<COLORREF>(lParam) != OLECOLOR2COLORREF(properties.backColor)) {
		properties.backColor = static_cast<COLORREF>(lParam);
	}
	return lr;
}

LRESULT RichTextBox::OnSetUndoLimit(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(!IsInDesignMode()) {
		properties.maxUndoQueueSize = lr;
	}
	return lr;
}


LRESULT RichTextBox::OnDragDropDoneNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	Raise_DragDropDone();
	return FALSE;
}

LRESULT RichTextBox::OnLinkNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	wasHandled = FALSE;

	ENLINK* pDetails = reinterpret_cast<ENLINK*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, ENLINK);
	if(!pDetails) {
		return FALSE; //(RunTimeHelperEx::IsWin8() ? EN_LINK_DO_DEFAULT : FALSE);
	}

	switch(pDetails->msg) {
		case WM_SETCURSOR:
		{
			HCURSOR hCursor = NULL;
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			if(properties.linkMousePointer == mpCustom) {
				if(properties.linkMouseIcon.pPictureDisp) {
					CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.linkMouseIcon.pPictureDisp);
					if(pPicture) {
						OLE_HANDLE h = NULL;
						pPicture->get_Handle(&h);
						hCursor = static_cast<HCURSOR>(LongToHandle(h));
					}
				}
			} else {
				hCursor = MousePointerConst2hCursor(properties.linkMousePointer);
			}

			if(hCursor) {
				SetCursor(hCursor);
				wasHandled = TRUE;
				return TRUE;
			}
			return FALSE; //(RunTimeHelperEx::IsWin8() ? EN_LINK_DO_DEFAULT : FALSE);
			break;
		}

		case WM_LBUTTONDBLCLK:
		{
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(pDetails->wParam, button, shift);
			button = 1/*MouseButtonConstants.vbLeftButton*/;
			if(!(properties.disabledEvents & deClickEvents)) {
				Raise_DblClick(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			}
			Raise_MouseUp(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			break;
		}

		case WM_LBUTTONDOWN:
		{
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(pDetails->wParam, button, shift);
			button = 1/*MouseButtonConstants.vbLeftButton*/;
			Raise_MouseDown(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			break;
		}

		case WM_LBUTTONUP:
		{
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(pDetails->wParam, button, shift);
			button = 1/*MouseButtonConstants.vbLeftButton*/;
			Raise_MouseUp(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			break;
		}

		case WM_MOUSEMOVE:
			if(!(properties.disabledEvents & deMouseEvents)) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(pDetails->wParam, button, shift);
				Raise_MouseMove(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			} else if(!mouseStatus.enteredControl) {
				mouseStatus.EnterControl();
			}
			break;

		case WM_RBUTTONDBLCLK:
		{
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(pDetails->wParam, button, shift);
			button = 2/*MouseButtonConstants.vbRightButton*/;
			if(!(properties.disabledEvents & deClickEvents)) {
				Raise_RDblClick(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			}
			Raise_MouseUp(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			break;
		}

		case WM_RBUTTONDOWN:
		{
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(pDetails->wParam, button, shift);
			button = 2/*MouseButtonConstants.vbRightButton*/;
			Raise_MouseDown(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			break;
		}

		case WM_RBUTTONUP:
		{
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(pDetails->wParam, button, shift);
			button = 2/*MouseButtonConstants.vbRightButton*/;
			Raise_MouseUp(button, shift, GET_X_LPARAM(pDetails->lParam), GET_Y_LPARAM(pDetails->lParam), &pDetails->chrg, htLink);
			break;
		}
	}
	return FALSE; //(RunTimeHelperEx::IsWin8() ? EN_LINK_DO_DEFAULT : FALSE);
}

LRESULT RichTextBox::OnRequestResizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	REQRESIZE* pDetails = reinterpret_cast<REQRESIZE*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, REQRESIZE);
	if(!pDetails) {
		return FALSE;
	}

	if(!(properties.disabledEvents & deShouldResizeControlWindow)) {
		Raise_ShouldResizeControlWindow(reinterpret_cast<RECTANGLE*>(&pDetails->rc));
	}
	return FALSE;
}

LRESULT RichTextBox::OnSelChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SELCHANGE* pDetails = reinterpret_cast<SELCHANGE*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, SELCHANGE);
	if(!pDetails) {
		return FALSE;
	}
	if(pDetails->chrg.cpMin == previousSelectedRange.cpMin && pDetails->chrg.cpMax == previousSelectedRange.cpMax) {
		// no real change - this happens for instance when calling ITextFont::SetBold
		return FALSE;
	}
	previousSelectedRange = pDetails->chrg;

	if(!(properties.disabledEvents & deSelectionChangingEvents) && flags.silentSelectionChanges == 0) {
		CComPtr<IRichTextRange> pSelectionTextRange = NULL;
		get_TextRange(pDetails->chrg.cpMin, pDetails->chrg.cpMax, &pSelectionTextRange);
		Raise_SelectionChanged(pSelectionTextRange, static_cast<SelectionTypeConstants>(pDetails->seltyp));
	}
	return FALSE;
}


LRESULT RichTextBox::OnReflectedChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	if(!(properties.disabledEvents & deTextChangedEvents)) {
		Raise_TextChanged();
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_RTB_TEXT);
	SendOnDataChange();

	wasHandled = FALSE;
	return 0;
}

LRESULT RichTextBox::OnReflectedMaxText(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	Raise_TruncatedText();
	wasHandled = FALSE;
	return 0;
}

LRESULT RichTextBox::OnReflectedScroll(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deScrolling)) {
		switch(notifyCode) {
			case EN_HSCROLL:
				Raise_Scrolling(saHorizontal);
				break;
			case EN_VSCROLL:
				Raise_Scrolling(saVertical);
				break;
		}
	}

	return 0;
}

LRESULT RichTextBox::OnReflectedSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(*this);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(*this, ime);
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}


void RichTextBox::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			  still valid, so simply update our font. */
		if(properties.font.pFontDisp) {
			CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
			if(pFont) {
				HFONT hFont = NULL;
				pFont->get_hFont(&hFont);
				properties.font.currentFont.Attach(hFont);
				properties.font.owningFontResource = FALSE;

				SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
		} else {
			SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
		}
		Invalidate();
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT RichTextBox::Raise_BeginDrag(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeginDrag(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_BeginRDrag(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeginRDrag(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_Click(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_ContextMenu(MenuTypeConstants menuType, IRichTextRange* pTextRange, SelectionTypeConstants selectionType, IRichOLEObject* pOLEObject, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL isRightMouseDrop, IOLEDataObject* pDraggedData, HMENU* pHMenuToDisplay)
{
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		OLE_HANDLE h = HandleToLong(*pHMenuToDisplay);
		hr = Fire_ContextMenu(menuType, pTextRange, selectionType, pOLEObject, button, shift, x, y, HitTest(x, y), isRightMouseDrop, pDraggedData, &h);
		*pHMenuToDisplay = static_cast<HMENU>(LongToHandle(h));
	//}
	return hr;
}

inline HRESULT RichTextBox::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange/* = NULL*/, HitTestConstants hitTestDetails/* = static_cast<HitTestConstants>(0)*/)
{
	CComPtr<ITextRange> pRange = NULL;
	if(properties.pTextDocument) {
		if(pCharRange) {
			properties.pTextDocument->Range(pCharRange->cpMin, pCharRange->cpMax, &pRange);
			hitTestDetails = static_cast<HitTestConstants>(hitTestDetails | HitTest(x, y));
		} else {
			POINT screenMousePosition = {x, y};
			ClientToScreen(&screenMousePosition);
			properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
			hitTestDetails = HitTest(x, y);
			if(hitTestDetails & htLink) {
				// we'll receive an EN_LINK notification with better details
				return S_OK;
			}
		}
	}

	//if(m_nFreezeEvents == 0) {
		CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);
		return Fire_DblClick(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_DestroyedControlWindow(HWND hWnd)
{
	if(properties.registerForOLEDragDrop == rfoddAdvancedDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	if(RunTimeHelperEx::IsWin8()) {
		// remove mouse hook
		MouseHookProvider::RemoveHook(this);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_DragDropDone(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DragDropDone();
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_GetDataObject(IRichTextRange* pTextRange, ClipboardOperationTypeConstants operationType, LONG* ppDataObject, VARIANT_BOOL* pUseCustomDataObject)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetDataObject(pTextRange, operationType, ppDataObject, pUseCustomDataObject);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_GetDragDropEffect(VARIANT_BOOL getSourceEffects, SHORT button, SHORT shift, OLEDropEffectConstants* pEffect, VARIANT_BOOL* pSkipDefaultProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetDragDropEffect(getSourceEffects, button, shift, pEffect, pSkipDefaultProcessing);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_InsertingOLEObject(IRichTextRange* pTextRange, BSTR* pClassID, LONG pStorage, VARIANT_BOOL* pCancelInsertion)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingOLEObject(pTextRange, pClassID, pStorage, pCancelInsertion);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_MClick(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	CComPtr<ITextRange> pRange = NULL;
	if(properties.pTextDocument) {
		POINT screenMousePosition = {x, y};
		ClientToScreen(&screenMousePosition);
		properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
	}

	//if(m_nFreezeEvents == 0) {
		CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);
		return Fire_MDblClick(pTextRange, button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange/* = NULL*/, HitTestConstants hitTestDetails/* = static_cast<HitTestConstants>(0)*/)
{
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		CComPtr<ITextRange> pRange = NULL;
		if(properties.pTextDocument) {
			if(pCharRange) {
				properties.pTextDocument->Range(pCharRange->cpMin, pCharRange->cpMax, &pRange);
				hitTestDetails = static_cast<HitTestConstants>(hitTestDetails | HitTest(x, y));
			} else {
				POINT screenMousePosition = {x, y};
				ClientToScreen(&screenMousePosition);
				properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
				hitTestDetails = HitTest(x, y);
				if(hitTestDetails & htLink) {
					// we'll receive an EN_LINK notification with better details
					return S_OK;
				}
			}
		}
		CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);

		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(pTextRange, button, shift, x, y, hitTestDetails);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(pTextRange, button, shift, x, y, hitTestDetails);
		}
		mouseStatus.StoreClickCandidate(button);
		SetCapture();

		return Fire_MouseDown(pTextRange, button, shift, x, y, hitTestDetails);
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
		mouseStatus.StoreClickCandidate(button);
		SetCapture();
		return S_OK;
	}
}

inline HRESULT RichTextBox::Raise_MouseEnter(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseEnter(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_MouseHover(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(pTextRange, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT RichTextBox::Raise_MouseLeave(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	mouseStatus.LeaveControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(pTextRange, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT RichTextBox::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange/* = NULL*/, HitTestConstants hitTestDetails/* = static_cast<HitTestConstants>(0)*/)
{
	CComPtr<ITextRange> pRange = NULL;
	if(properties.pTextDocument) {
		if(pCharRange) {
			properties.pTextDocument->Range(pCharRange->cpMin, pCharRange->cpMax, &pRange);
			hitTestDetails = static_cast<HitTestConstants>(hitTestDetails | HitTest(x, y));
		} else {
			POINT screenMousePosition = {x, y};
			ClientToScreen(&screenMousePosition);
			properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
			hitTestDetails = HitTest(x, y);
			if(hitTestDetails & htLink) {
				// we'll receive an EN_LINK notification with better details
				return S_OK;
			}
		}
	}
	CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);

	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(pTextRange, button, shift, x, y, hitTestDetails);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange/* = NULL*/, HitTestConstants hitTestDetails/* = static_cast<HitTestConstants>(0)*/)
{
	CComPtr<ITextRange> pRange = NULL;
	BOOL dontRaise = FALSE;
	if(properties.pTextDocument) {
		if(pCharRange) {
			properties.pTextDocument->Range(pCharRange->cpMin, pCharRange->cpMax, &pRange);
			hitTestDetails = static_cast<HitTestConstants>(hitTestDetails | HitTest(x, y));
		} else {
			POINT screenMousePosition = {x, y};
			ClientToScreen(&screenMousePosition);
			properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
			hitTestDetails = HitTest(x, y);
			if(hitTestDetails & htLink) {
				// we'll receive an EN_LINK notification with better details
				dontRaise = TRUE;
			}
		}
	}
	CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);

	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents) && !dontRaise) {
		hr = Fire_MouseUp(pTextRange, button, shift, x, y, hitTestDetails);
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(GetCapture() == *this/* && !dontRaise*/) {
			ReleaseCapture();
		}

		BOOL removeClickCandidate = TRUE;
		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						if(dontRaise) {
							removeClickCandidate = FALSE;
						} else {
							Raise_Click(pTextRange, button, shift, x, y, hitTestDetails);
						}
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						if(dontRaise) {
							removeClickCandidate = FALSE;
						} else {
							Raise_RClick(pTextRange, button, shift, x, y, hitTestDetails);
						}
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(pTextRange, button, shift, x, y, hitTestDetails);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(pTextRange, button, shift, x, y, hitTestDetails);
					}
					break;
			}
		}

		if(removeClickCandidate) {
			mouseStatus.RemoveClickCandidate(button);
		}
	}

	return hr;
}

inline HRESULT RichTextBox::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	CComPtr<ITextRange> pRange = NULL;
	if(properties.pTextDocument) {
		POINT screenMousePosition = {x, y};
		ClientToScreen(&screenMousePosition);
		properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
	}
	CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);
	HitTestConstants hitTestDetails = HitTest(x, y);

	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(pTextRange, button, shift, x, y, hitTestDetails);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(pTextRange, button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	CComPtr<IRichTextRange> pDropTarget = NULL;
	ATLVERIFY(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, &pDropTarget)));

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT RichTextBox::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	CComPtr<IRichTextRange> pDropTarget = NULL;
	ATLVERIFY(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, &pDropTarget)));

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		WTL::CPoint mousePos(mousePosition.x, mousePosition.y);
		WTL::CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT RichTextBox::Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_OLEDragLeave(void)
{
	KillTimer(timers.ID_DRAGSCROLL);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	CComPtr<IRichTextRange> pDropTarget = NULL;
	ATLVERIFY(SUCCEEDED(HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails, &pDropTarget)));

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, hitTestDetails);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT RichTextBox::Raise_OLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0);
	CComPtr<IRichTextRange> pDropTarget = NULL;
	ATLVERIFY(SUCCEEDED(HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, &pDropTarget)));

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		WTL::CPoint mousePos(mousePosition.x, mousePosition.y);
		WTL::CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = noScrollZone.PtInRect(mousePos);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, hitTestDetails, &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() / 4 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT RichTextBox::Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
{
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
		return Fire_OLEQueryContinueDrag(BOOL2VARIANTBOOL(pressedEscape), button, shift, reinterpret_cast<OLEActionToContinueWithConstants*>(pActionToContinueWith));
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT RichTextBox::Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT RichTextBox::Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLESetData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_OLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_QueryAcceptData(IOLEDataObject* pData, LONG* pFormatID, ClipboardOperationTypeConstants operationType, VARIANT_BOOL performOperation, OLE_HANDLE hMetafilePicture, QueryAcceptDataConstants* pAcceptData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_QueryAcceptData(pData, pFormatID, operationType, performOperation, hMetafilePicture, pAcceptData);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_RClick(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange/* = NULL*/, HitTestConstants hitTestDetails/* = static_cast<HitTestConstants>(0)*/)
{
	CComPtr<ITextRange> pRange = NULL;
	if(properties.pTextDocument) {
		if(pCharRange) {
			properties.pTextDocument->Range(pCharRange->cpMin, pCharRange->cpMax, &pRange);
			hitTestDetails = static_cast<HitTestConstants>(hitTestDetails | HitTest(x, y));
		} else {
			POINT screenMousePosition = {x, y};
			ClientToScreen(&screenMousePosition);
			properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
			hitTestDetails = HitTest(x, y);
			if(hitTestDetails & htLink) {
				// we'll receive an EN_LINK notification with better details
				return S_OK;
			}
		}
	}

	//if(m_nFreezeEvents == 0) {
		CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);
		return Fire_RDblClick(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_RecreatedControlWindow(HWND hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop == rfoddAdvancedDragDrop) {
		RevokeDragDrop(*this);					// even with ES_NOOLEDRAGDROP the native control calls RegisterDragDrop
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(RunTimeHelperEx::IsWin8()) {
		// setup mouse hook
		/* This is a really ugly hack. For some reason (does not seem to be our fault) the spell checking
		 * context menu of Windows 8 and newer does not change the mouse cursor to the default arrow. The
		 * only way to replace the I-beam cursor with the default arrow when this context menu is active,
		 * seems to be a mouse hook that watches for WM_MOUSEMOVE that was meant for a menu window.
		 * The following alternatives have been tried:
		 * - Detect the menu in OnSetCursor or Raise_MouseUp.
		 * - Don't call SetCapture. -> It does not seem to be the root of the problem.
		 * - SetCursor(OCR_NORMAL) on WM_CANCELMODE -> WM_CANCELMODE is not sent.
		 * - CallWndProc-Hook, calling SetCursor(OCR_NORMAL) on WM_SETCURSOR -> WM_SETCURSOR is not sent or
		 *   the hook was not working.
		 * - CallWndProc-Hook, calling SetCursor(OCR_NORMAL) on WM_MOUSEMOVE -> did not work or the hook was
		 *   not working.
		 * - Mouse-Hook, calling SetCursor(OCR_NORMAL) on WM_SETCURSOR -> WM_SETCURSOR is not sent.
		 */
		MouseHookProvider::InstallHook(this);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_RemovingOLEObject(IRichOLEObject* pOLEObject, VARIANT_BOOL* pCancelDeletion)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingOLEObject(pOLEObject, pCancelDeletion);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_Scrolling(ScrollAxisConstants axis)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Scrolling(axis);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_SelectionChanged(IRichTextRange* pSelectedTextRange, SelectionTypeConstants selectionType)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectionChanged(pSelectedTextRange, selectionType);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_ShouldResizeControlWindow(RECTANGLE* pSuggestedControlSize)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ShouldResizeControlWindow(pSuggestedControlSize);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_TextChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TextChanged();
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_TruncatedText(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TruncatedText();
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_XClick(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(pTextRange, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT RichTextBox::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	CComPtr<ITextRange> pRange = NULL;
	if(properties.pTextDocument) {
		POINT screenMousePosition = {x, y};
		ClientToScreen(&screenMousePosition);
		properties.pTextDocument->RangeFromPoint(screenMousePosition.x, screenMousePosition.y, &pRange);
	}

	//if(m_nFreezeEvents == 0) {
		CComPtr<IRichTextRange> pTextRange = ClassFactory::InitTextRange(pRange, this);
		return Fire_XDblClick(pTextRange, button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}


void RichTextBox::EnterSilentSelectionChangesSection(void)
{
	++flags.silentSelectionChanges;
}

void RichTextBox::LeaveSilentSelectionChangesSection(void)
{
	--flags.silentSelectionChanges;
}

void RichTextBox::EnterSilentObjectDeletionsSection(void)
{
	++flags.silentObjectDeletions;
}

void RichTextBox::LeaveSilentObjectDeletionsSection(void)
{
	--flags.silentObjectDeletions;
}

void RichTextBox::EnterSilentObjectInsertionsSection(void)
{
	++flags.silentObjectInsertions;
}

void RichTextBox::LeaveSilentObjectInsertionsSection(void)
{
	--flags.silentObjectInsertions;
}


void RichTextBox::RecreateControlWindow(void)
{
	if(m_bInPlaceActive && !flags.dontRecreate) {
		BSTR richText = NULL;
		get_RichText(&richText);

		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));

		if(richText) {
			put_RichText(richText);
			SysFreeString(richText);
		}
	}
}

IMEModeConstants RichTextBox::GetCurrentIMEContextMode(HWND hWnd)
{
	IMEModeConstants imeContextMode = imeNoControl;

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		imeContextMode = imeInherit;
	} else {
		HIMC hIMC = ImmGetContext(hWnd);
		if(hIMC) {
			if(ImmGetOpenStatus(hIMC)) {
				DWORD conversionMode = 0;
				DWORD sentenceMode = 0;
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				if(conversionMode & IME_CMODE_NATIVE) {
					if(conversionMode & IME_CMODE_KATAKANA) {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeKatakanaHalf];
						} else {
							imeContextMode = countryTable[imeAlphaFull];
						}
					} else {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeHiragana];
						} else {
							imeContextMode = countryTable[imeKatakana];
						}
					}
				} else {
					if(conversionMode & IME_CMODE_FULLSHAPE) {
						imeContextMode = countryTable[imeAlpha];
					} else {
						imeContextMode = countryTable[imeHangulFull];
					}
				}
			} else {
				imeContextMode = countryTable[imeDisable];
			}
			ImmReleaseContext(hWnd, hIMC);
		} else {
			imeContextMode = countryTable[imeOn];
		}
	}
	return imeContextMode;
}

void RichTextBox::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
{
	if(!::IsWindow(hWnd)) {
		return;
	}
	if((IMEMode == imeInherit) || (IMEMode == imeNoControl)) {
		if(::SendMessage(hWnd, EM_ISIME, 0, 0)) {
			::SendMessage(hWnd, EM_SETIMEMODEBIAS, properties.IMEConversionMode, properties.IMEConversionMode);
		}
		::SendMessage(hWnd, EM_SETCTFMODEBIAS, properties.TSFModeBias, 0);
		return;
	}

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		if(::SendMessage(hWnd, EM_ISIME, 0, 0)) {
			::SendMessage(hWnd, EM_SETIMEMODEBIAS, properties.IMEConversionMode, properties.IMEConversionMode);
		}
		::SendMessage(hWnd, EM_SETCTFMODEBIAS, properties.TSFModeBias, 0);
		return;
	}

	// update IME mode
	HIMC hIMC = ImmGetContext(hWnd);
	if(IMEMode == imeDisable) {
		// disable IME
		if(hIMC) {
			// close the IME
			if(ImmGetOpenStatus(hIMC)) {
				ImmSetOpenStatus(hIMC, FALSE);
			}
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
		// remove the control's association to the IME context
		HIMC h = ImmAssociateContext(hWnd, NULL);
		if(h) {
			IMEFlags.hDefaultIMC = h;
		}
		return;
	} else {
		// enable IME
		if(!hIMC) {
			if(!IMEFlags.hDefaultIMC) {
				// create an IME context
				hIMC = ImmCreateContext();
				if(hIMC) {
					// associate the control with the IME context
					ImmAssociateContext(hWnd, hIMC);
				}
			} else {
				// associate the control with the default IME context
				ImmAssociateContext(hWnd, IMEFlags.hDefaultIMC);
			}
		} else {
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
	}

	hIMC = ImmGetContext(hWnd);
	if(hIMC) {
		DWORD conversionMode = 0;
		DWORD sentenceMode = 0;
		switch(IMEMode) {
			case imeOn:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				break;
			case imeOff:
				// close IME
				ImmSetOpenStatus(hIMC, FALSE);
				break;
			default:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				// switch conversion
				switch(IMEMode) {
					case imeHiragana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_KATAKANA;
						break;
					case imeKatakana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeKatakanaHalf:
						conversionMode |= (IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
					case imeAlphaFull:
						conversionMode |= IME_CMODE_FULLSHAPE;
						conversionMode &= ~(IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeAlpha:
						conversionMode |= IME_CMODE_ALPHANUMERIC;
						conversionMode &= ~(IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeHangulFull:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeHangul:
						conversionMode |= IME_CMODE_NATIVE;
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
				}
				ImmSetConversionStatus(hIMC, conversionMode, sentenceMode);
				break;
		}
		// each ImmGetContext() needs a ImmReleaseContext()
		ImmReleaseContext(hWnd, hIMC);
		hIMC = NULL;

		if(::SendMessage(hWnd, EM_ISIME, 0, 0)) {
			::SendMessage(hWnd, EM_SETIMEMODEBIAS, properties.IMEConversionMode, properties.IMEConversionMode);
		}
		::SendMessage(hWnd, EM_SETCTFMODEBIAS, properties.TSFModeBias, 0);
	}
}

IMEModeConstants RichTextBox::GetEffectiveIMEMode(void)
{
	IMEModeConstants IMEMode = properties.IMEMode;
	if((IMEMode == imeInherit) && IsWindow()) {
		CWindow wnd = GetParent();
		while((IMEMode == imeInherit) && wnd.IsWindow()) {
			// retrieve the parent's IME mode
			IMEMode = GetCurrentIMEContextMode(wnd);
			wnd = wnd.GetParent();
		}
	}

	if(IMEMode == imeInherit) {
		// use imeNoControl as fallback
		IMEMode = imeNoControl;
	}
	return IMEMode;
}

LPCTSTR RichTextBox::GetWndClassName(void)
{
	properties.richEditAPIVersion = reavUnknown;
	properties.richEditLibraryPath = L"";
	RichEditAPIVersionConstants availableVersionByWindows = (RunTimeHelperEx::IsWin8() ? reav75 : reav41);
	if(properties.richEditVersion != revAlwaysUseWindowsRichEdit) {
		// try MS Office rich edit

		TCHAR pCommonFilesPath[1024];
		ZeroMemory(pCommonFilesPath, 1024 * sizeof(TCHAR));
		if(!SHGetSpecialFolderPath(NULL, pCommonFilesPath, CSIDL_PROGRAM_FILES_COMMON, FALSE)) {
			SHGetSpecialFolderPath(NULL, pCommonFilesPath, CSIDL_PROGRAM_FILES, FALSE);
			if(lstrlen(pCommonFilesPath) > 0) {
				PathAddBackslash(pCommonFilesPath);
				StringCchCat(pCommonFilesPath, 1024, TEXT("Common Files"));
			}
		}
		PathAddBackslash(pCommonFilesPath);

		RichEditAPIVersionConstants availableVersionByOffice = reavUnknown;
		if(lstrlen(pCommonFilesPath) > 0) {
			TCHAR pRichEd20Path[1024];
			ZeroMemory(pRichEd20Path, 1024 * sizeof(TCHAR));

			if(availableVersionByOffice == reavUnknown) {
				// check for Office 2016
				SHGetSpecialFolderPath(NULL, pRichEd20Path, CSIDL_PROGRAM_FILES, FALSE);
				if(lstrlen(pRichEd20Path) > 0) {
					PathAddBackslash(pRichEd20Path);
					StringCchCat(pRichEd20Path, 1024, TEXT("Microsoft Office\\root\\VFS\\ProgramFilesCommonX86\\Microsoft Shared\\OFFICE16\\RICHED20.DLL"));
					if(PathFileExists(pRichEd20Path)) {
						availableVersionByOffice = reav80;
					}
				}
			}
			if(availableVersionByOffice == reavUnknown) {
				// check for Office 2013
				lstrcpyn(pRichEd20Path, pCommonFilesPath, 1024);
				StringCchCat(pRichEd20Path, 1024, TEXT("Microsoft Shared\\OFFICE15\\RICHED20.DLL"));
				if(PathFileExists(pRichEd20Path)) {
					availableVersionByOffice = reav80;
				}
			}
			if(availableVersionByOffice == reavUnknown) {
				// check for Office 2010
				lstrcpyn(pRichEd20Path, pCommonFilesPath, 1024);
				StringCchCat(pRichEd20Path, 1024, TEXT("Microsoft Shared\\OFFICE14\\RICHED20.DLL"));
				if(PathFileExists(pRichEd20Path)) {
					availableVersionByOffice = reav70;
				}
			}
			if(availableVersionByOffice == reavUnknown) {
				// check for Office 2007
				lstrcpyn(pRichEd20Path, pCommonFilesPath, 1024);
				StringCchCat(pRichEd20Path, 1024, TEXT("Microsoft Shared\\OFFICE12\\RICHED20.DLL"));
				if(PathFileExists(pRichEd20Path)) {
					availableVersionByOffice = reav60;
				}
			}
			if(availableVersionByOffice == reavUnknown) {
				// check for Office 2003
				lstrcpyn(pRichEd20Path, pCommonFilesPath, 1024);
				StringCchCat(pRichEd20Path, 1024, TEXT("Microsoft Shared\\OFFICE11\\RICHED20.DLL"));
				if(PathFileExists(pRichEd20Path)) {
					availableVersionByOffice = reav50;
				}
			}

			BOOL useOfficeVersion = FALSE;
			switch(properties.richEditVersion) {
				case revPreferOfficeRichEdit:
					useOfficeVersion = TRUE;
					break;
				case revUseHighestVersion:
					useOfficeVersion = (availableVersionByOffice >= availableVersionByWindows);
					break;
			}

			if(useOfficeVersion) {
				HMODULE hRichEdit = GetModuleHandle(pRichEd20Path);
				if(!hRichEdit && PathFileExists(pRichEd20Path)) {
					hRichEdit = LoadLibrary(pRichEd20Path);
				}
				if(hRichEdit) {
					TCHAR pBuffer[1024];
					GetModuleFileName(hRichEdit, pBuffer, 1024);
					properties.richEditAPIVersion = availableVersionByOffice;
					properties.richEditLibraryPath = pBuffer;
					return WNDCLASS_RICHEDIT60;
				}
			}
		}
	}

	HMODULE hMsftEdit = GetModuleHandle(TEXT("msftedit.dll"));
	if(!hMsftEdit) {
		hMsftEdit = LoadLibrary(TEXT("msftedit.dll"));
	}
	if(hMsftEdit) {
		TCHAR pBuffer[1024];
		GetModuleFileName(hMsftEdit, pBuffer, 1024);
		properties.richEditAPIVersion = availableVersionByWindows;
		properties.richEditLibraryPath = pBuffer;
	} else {
		MessageBox(TEXT("Could not load 'msftedit.dll'. This program won't work."), TEXT("Error"), MB_ICONERROR);
	}
	return WNDCLASS_RICHEDIT50;
}

DWORD RichTextBox::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD RichTextBox::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(properties.alwaysShowScrollBars) {
		style |= ES_DISABLENOSCROLL;
	}
	if(properties.alwaysShowSelection) {
		/* Yes, also set the style, because it seems like EM_SETOPTIONS is not meant to be called while
		   handling WM_CREATE. */
		style |= ES_NOHIDESEL;
	}
	if(properties.autoScrolling & asVertical) {
		/* Yes, also set the style, because it seems like EM_SETOPTIONS is not meant to be called while
		   handling WM_CREATE. */
		style |= ES_AUTOVSCROLL;
	}
	if(properties.autoScrolling & asHorizontal) {
		/* Yes, also set the style, because it seems like EM_SETOPTIONS is not meant to be called while
		   handling WM_CREATE. */
		style |= ES_AUTOHSCROLL;
	}
	if(properties.disableIMEOperations) {
		style |= ES_NOIME;
	}
	//switch(properties.hAlignment) {
	//	case halLeft:
	//		style |= ES_LEFT;
	//		break;
	//	case halCenter:
	//		style |= ES_CENTER;
	//		break;
	//	case halRight:
	//		style |= ES_RIGHT;
	//		break;
	//}
	if(properties.letClientHandleAllIMEOperations) {
		style |= ES_SELFIME;
	}
	if(properties.multiLine) {
		style |= ES_MULTILINE;
	}
	if(properties.registerForOLEDragDrop == rfoddNoDragDrop) {
		style |= ES_NOOLEDRAGDROP;
	}
	if(properties.showSelectionBar) {
		/* Yes, also set the style, because it seems like EM_SETOPTIONS is not meant to be called while
		   handling WM_CREATE. */
		style |= ES_SELECTIONBAR;
	}
	return style;
}

void RichTextBox::SendConfigurationMessages(void)
{
	DWORD editStyle = 0;
	if(!properties.adjustLineHeightForEastAsianLanguages) {
		editStyle |= SES_NOEALINEHEIGHTADJUST;
	}
	if(!properties.allowInputThroughTSF) {
		editStyle |= SES_CTFNOLOCK;
	}
	if(properties.allowObjectInsertionThroughTSF) {
		editStyle |= SES_CTFALLOWEMBED;
	}
	if(properties.allowTSFProofingTips) {
		editStyle |= SES_CTFALLOWPROOFING;
	}
	if(properties.allowTSFSmartTags) {
		editStyle |= SES_CTFALLOWSMARTTAG;
	}
	if(properties.beepOnMaxText) {
		editStyle |= SES_BEEPONMAXTEXT;
	}
	switch(properties.characterConversion) {
		case ccLowerCase:
			editStyle |= SES_LOWERCASE;
			break;
		case ccUpperCase:
			editStyle |= SES_UPPERCASE;
			break;
	}
	if(properties.disableIMEOperations) {
		editStyle |= SES_NOIME;
	}
	if(properties.displayHyperlinkTooltips) {
		editStyle |= SES_HYPERLINKTOOLTIPS;
	}
	if(!properties.displayZeroWidthTableCellBorders) {
		editStyle |= SES_HIDEGRIDLINES;
	}
	if(properties.draftMode) {
		editStyle |= SES_DRAFTMODE;
	}
	if(properties.dropWordsOnWordBoundariesOnly) {
		editStyle |= SES_WORDDRAGDROP;
	}
	if(properties.emulateSimpleTextBox) {
		editStyle |= SES_EMULATESYSEDIT;
	}
	if(properties.extendFontBackColorToWholeLine) {
		editStyle |= SES_EXTENDBACKCOLOR;
	}
	if(properties.logicalCaret) {
		editStyle |= SES_LOGICALCARET;
	}
	if(properties.multiSelect) {
		editStyle |= SES_MULTISELECT;
	}
	if(properties.noInputSequenceCheck) {
		editStyle |= SES_NOINPUTSEQUENCECHK;
	}
	if(properties.scrollToTopOnKillFocus) {
		editStyle |= SES_SCROLLONKILLFOCUS;
	}
	if(properties.smartSpacingOnDrop) {
		editStyle |= SES_SMARTDRAGDROP;
	}
	if(properties.useTextServicesFramework) {
		// SES_USECTF is being told to enable dictation support. (https://blogs.msdn.com/b/tsfaware/archive/2007/06/12/easy-dictation-support-for-windowless-richedit-controls.aspx)
		// But on Windows 7 it does not seem to be necessary.
		editStyle |= SES_USECTF;
	}
	SendMessage(EM_SETEDITSTYLE, editStyle, SES_NOEALINEHEIGHTADJUST | SES_CTFNOLOCK | SES_CTFALLOWEMBED | SES_CTFALLOWPROOFING | SES_CTFALLOWSMARTTAG | SES_BEEPONMAXTEXT | SES_LOWERCASE | SES_UPPERCASE | SES_NOIME | SES_HYPERLINKTOOLTIPS | SES_HIDEGRIDLINES | SES_DRAFTMODE | SES_WORDDRAGDROP | SES_EMULATESYSEDIT | SES_EXTENDBACKCOLOR | SES_LOGICALCARET | SES_MULTISELECT | SES_NOINPUTSEQUENCECHK | SES_SCROLLONKILLFOCUS | SES_SMARTDRAGDROP | SES_USECTF);
	editStyle = SES_EX_MULTITOUCH;
	if(!properties.allowMathZoneInsertion) {
		editStyle |= SES_EX_NOMATH;
	}
	if(!properties.allowTableInsertion) {
		editStyle |= SES_EX_NOTABLE;
	}
	if(!properties.useBkAcetateColorForTextSelection) {
		editStyle |= SES_EX_NOACETATESELECTION;
	}
	if(!properties.useWindowsThemes) {
		editStyle |= SES_EX_NOTHEMING;
	}
	SendMessage(EM_SETEDITSTYLEEX, editStyle, SES_EX_MULTITOUCH | SES_EX_NOMATH | SES_EX_NOTABLE | SES_EX_NOACETATESELECTION | SES_EX_NOTHEMING);
	SendMessage(EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY | TO_ADVANCEDLAYOUT, TO_ADVANCEDTYPOGRAPHY | TO_SIMPLELINEBREAK | TO_ADVANCEDLAYOUT);
	DWORD options = SendMessage(EM_GETLANGOPTIONS, 0, 0);
	/*if(properties.discardCompositionStringOnIMECancel) {
		options |= IMF_IMECANCELCOMPLETE;
	} else {
		options &= ~IMF_IMECANCELCOMPLETE;
	}*/
	if(properties.switchFontOnIMEInput) {
		options |= IMF_AUTOFONT;
	} else {
		options &= ~IMF_AUTOFONT;
	}
	if(properties.useBuiltInSpellChecking) {
		options |= IMF_SPELLCHECKING;
	} else {
		options &= ~IMF_SPELLCHECKING;
	}
	if(properties.useDualFontMode) {
		options |= IMF_DUALFONT;
	} else {
		options &= ~IMF_DUALFONT;
	}
	/*if(properties.useTouchKeyboardAutoCorrection) {
		options |= IMF_TKBAUTOCORRECTION;
	} else {
		options &= ~IMF_TKBAUTOCORRECTION;
	}*/
	SendMessage(EM_SETLANGOPTIONS, 0, options);
	
	LPRICHEDITOLECALLBACK pOleCallback = NULL;
	QueryInterface(IID_IRichEditOleCallback, reinterpret_cast<LPVOID*>(&pOleCallback));
	ATLASSUME(pOleCallback);
	if(pOleCallback) {
		#ifdef DEBUG
			ATLASSERT(SendMessage(EM_SETOLECALLBACK, 0, reinterpret_cast<LPARAM>(pOleCallback)));
		#else
			SendMessage(EM_SETOLECALLBACK, 0, reinterpret_cast<LPARAM>(pOleCallback));
		#endif
		pOleCallback->Release();
		pOleCallback = NULL;
	}

	if(properties.backColor == (0x80000000 | COLOR_WINDOW)) {
		SendMessage(EM_SETBKGNDCOLOR, TRUE, 0);
	} else {
		SendMessage(EM_SETBKGNDCOLOR, FALSE, OLECOLOR2COLORREF(properties.backColor));
	}
	SendMessage(EM_SETREADONLY, properties.readOnly, 0);

	DWORD events = ENM_CHANGE | ENM_DRAGDROPDONE | ENM_LINK;					// NOTE: We always want EN_CHANGE notifications because of data-binding.
	if(!(properties.disabledEvents & deScrolling)) {
		events |= ENM_SCROLL;
	}
	if(!(properties.disabledEvents & deSelectionChangingEvents)) {
		events |= ENM_SELCHANGE;
	}
	if(!(properties.disabledEvents & deShouldResizeControlWindow)) {
		events |= ENM_REQUESTRESIZE;
	}
	SendMessage(EM_SETEVENTMASK, 0, events);

	SendMessage(EM_SETPAGEROTATE, properties.textFlow, 0);
	if(properties.useCustomFormattingRectangle) {
		SendMessage(EM_SETRECTNP, FALSE, reinterpret_cast<LPARAM>(&properties.formattingRectangle));
	} else {
		SendMessage(EM_SETRECTNP, FALSE, NULL);
	}

	SetWindowText(COLE2CT(properties.text));

	SendMessage(EM_SHOWSCROLLBAR, SB_VERT, (properties.scrollBars & sbVertical) == sbVertical);
	SendMessage(EM_SHOWSCROLLBAR, SB_HORZ, (properties.scrollBars & sbHorizontal) == sbHorizontal);

	if(properties.alwaysShowSelection) {
		SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_NOHIDESEL);
	} else {
		SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_NOHIDESEL);
	}
	if(properties.autoScrolling & asHorizontal) {
		SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_AUTOHSCROLL);
	} else {
		SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_AUTOHSCROLL);
	}
	if(properties.autoScrolling & asVertical) {
		SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_AUTOVSCROLL);
	} else {
		SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_AUTOVSCROLL);
	}
	if(properties.autoSelectWordOnTrackSelection) {
		SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_AUTOWORDSELECTION);
	} else {
		SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_AUTOWORDSELECTION);
	}
	if(properties.showSelectionBar) {
		SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_SELECTIONBAR);
	} else {
		SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_SELECTIONBAR);
	}

	SendMessage(EM_EXLIMITTEXT, 0, (properties.maxTextLength == -1 ? 0 : properties.maxTextLength));

	ApplyFont();

	switch(properties.textOrientation) {
		case toHorizontal:
			SendMessage(EM_SETOPTIONS, ECOOP_AND, 0xFFFFFFFF & ~ECO_VERTICAL);
			SendMessage(EM_SETEDITSTYLE, 0, SES_USEATFONT);
			break;
		case toVertical:
			SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_VERTICAL);
			SendMessage(EM_SETEDITSTYLE, SES_USEATFONT, SES_USEATFONT);
			break;
		case toVerticalWithHorizontalFont:
			SendMessage(EM_SETOPTIONS, ECOOP_OR, ECO_VERTICAL);
			SendMessage(EM_SETEDITSTYLE, 0, SES_USEATFONT);
			break;
	}

	int numerator = 0;
	int denominator = 0;
	if(properties.zoomRatio != 0.0) {
		numerator = static_cast<int>(properties.zoomRatio);
		denominator = 1;
		DOUBLE ratio = static_cast<DOUBLE>(numerator) / static_cast<DOUBLE>(denominator);
		while((ratio != properties.zoomRatio) && numerator < 1000000000 && denominator < 1000000000) {
			denominator *= 10;
			numerator = static_cast<int>(properties.zoomRatio * static_cast<DOUBLE>(denominator));
			ratio = static_cast<DOUBLE>(numerator) / static_cast<DOUBLE>(denominator);
		}
	}
	SendMessage(EM_SETZOOM, numerator, denominator);

	// Does not work for RichEdit:
	//SendMessage(EM_SETMARGINS, (EC_LEFTMARGIN | EC_RIGHTMARGIN), MAKELPARAM((properties.leftMargin == -1 ? EC_USEFONTINFO : properties.leftMargin), (properties.rightMargin == -1 ? EC_USEFONTINFO : properties.rightMargin)));
	if(properties.leftMargin != -1 && properties.rightMargin != -1) {
		SendMessage(EM_SETMARGINS, (EC_LEFTMARGIN | EC_RIGHTMARGIN), MAKELPARAM(properties.leftMargin, properties.rightMargin));
	}

	BOOL useFallback = TRUE;
	if(properties.pTextDocument) {
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			useFallback = FAILED(pTextDocument2->SetProperty(tomUndoLimit, properties.maxUndoQueueSize));
		}
	}
	if(useFallback) {
		SendMessage(EM_SETUNDOLIMIT, properties.maxUndoQueueSize, 0);
	}
	SendMessage(EM_SETMODIFY, properties.modified, 0);

	DWORD autoDetectURLs = properties.autoDetectURLs;
	if(properties.autoDetectedURLSchemes.Length() > 0) {
		if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, reinterpret_cast<LPARAM>(OLE2W(properties.autoDetectedURLSchemes))) != 0) {
			// failure - on Windows 7 and older, try without URL schemes
			if(!RunTimeHelperEx::IsWin8() && SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
				// success
				if(!IsInDesignMode()) {
					properties.autoDetectedURLSchemes = L"";
					properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
				}
			} else {
				// still a failure - ensure that only valid flags are used
				if(RunTimeHelperEx::IsWin8()) {
					autoDetectURLs &= (AURL_ENABLEURL | AURL_ENABLEEMAILADDR | AURL_ENABLETELNO | AURL_ENABLEEAURLS | AURL_ENABLEDRIVELETTERS | AURL_DISABLEMIXEDLGC);
					if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, reinterpret_cast<LPARAM>(OLE2W(properties.autoDetectedURLSchemes))) == 0) {
						// success
						if(!IsInDesignMode()) {
							properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
						}
					}
				} else {
					autoDetectURLs &= AURL_ENABLEURL;
					if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
						// success
						if(!IsInDesignMode()) {
							properties.autoDetectedURLSchemes = L"";
							properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
						}
					}
				}
			}
		}
	} else {
		if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) != 0) {
			// failure - ensure that only valid flags are used
			if(RunTimeHelperEx::IsWin8()) {
				autoDetectURLs = (autoDetectURLs & (AURL_ENABLEURL | AURL_ENABLEEMAILADDR | AURL_ENABLETELNO | AURL_ENABLEEAURLS | AURL_ENABLEDRIVELETTERS | AURL_DISABLEMIXEDLGC));
			} else {
				autoDetectURLs = (autoDetectURLs & AURL_ENABLEURL);
			}
			if(SendMessage(EM_AUTOURLDETECT, autoDetectURLs, NULL) == 0) {
				// success
				if(!IsInDesignMode()) {
					properties.autoDetectURLs = static_cast<AutoDetectURLsConstants>(autoDetectURLs);
				}
			}
		}
	}

	HYPHENATEINFO hyphenateInfo = {0};
	hyphenateInfo.cbSize = sizeof(HYPHENATEINFO);
	hyphenateInfo.dxHyphenateZone = properties.hyphenationWordWrapZoneWidth;
	hyphenateInfo.pfnHyphenate = reinterpret_cast<PFNHYPHENATEPROC>(properties.hyphenationFunction);
	SendMessage(EM_SETHYPHENATEINFO, reinterpret_cast<WPARAM>(&hyphenateInfo), 0);

	if(properties.pTextDocument) {
		properties.pTextDocument->SetDefaultTabStop(properties.defaultTabWidth);
		CComQIPtr<ITextDocument2> pTextDocument2 = properties.pTextDocument;
		if(pTextDocument2) {
			long options = 0;
			switch(properties.mathLineBreakBehavior) {
				case mlbbBreakBeforeOperator:
					options = tomMathBrkBinBefore;
					break;
				case mlbbBreakAfterOperator:
					options = tomMathBrkBinAfter;
					break;
				case mlbbDuplicateOperatorMinusMinus:
					options = tomMathBrkBinDup | tomMathBrkBinSubMM;
					break;
				case mlbbDuplicateOperatorPlusMinus:
					options = tomMathBrkBinDup | tomMathBrkBinSubPM;
					break;
				case mlbbDuplicateOperatorMinusPlus:
					options = tomMathBrkBinDup | tomMathBrkBinSubMP;
					break;
			}
			ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(options, tomMathBrkBinMask | tomMathBrkBinSubMask)));
			options = 0;
			switch(properties.defaultMathZoneHAlignment) {
				case halLeft:
					options = tomMathDispAlignLeft;
					break;
				case halCenter:
					options = tomMathDispAlignCenter;
					break;
				case halRight:
					options = tomMathDispAlignRight;
					break;
				default:
					options = tomMathDispAlignCenter;
					break;
			}
			ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(options, tomMathDispAlignMask)));
			ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.denoteEmptyMathArguments, tomMathDocEmptyArgMask)));
			switch(properties.integralLimitsLocation) {
				case llOnSide:
					ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(0, tomMathDispIntUnderOver)));
					break;
				case llUnderAndOver:
					ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathDispIntUnderOver, tomMathDispIntUnderOver)));
					break;
			}
			switch(properties.nAryLimitsLocation) {
				case llOnSide:
					ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(tomMathDispNarySubSup, tomMathDispNarySubSup)));
					break;
				case llUnderAndOver:
					ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(0, tomMathDispNarySubSup)));
					break;
			}
			ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.growNAryOperators ? tomMathDispNaryGrow : 0, tomMathDispNaryGrow)));
			ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.rawSubScriptAndSuperScriptOperators ? tomMathDocSbSpOpUnchanged : 0, tomMathDocSbSpOpUnchanged)));
			ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.rightToLeftMathZones ? tomMathEnableRtl : 0, tomMathEnableRtl)));
			ATLASSERT(SUCCEEDED(pTextDocument2->SetMathProperties(properties.useSmallerFontForNestedFractions ? tomMathDispFracTeX : 0, tomMathDispFracTeX)));
		}
	}
}

HCURSOR RichTextBox::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

HitTestConstants RichTextBox::HitTest(LONG x, LONG y)
{
	ATLASSERT(IsWindow());

	POINT pt = {x, y};
	HitTestConstants ret = static_cast<HitTestConstants>(0);

	// Are we outside the window?
	WTL::CRect windowRectangle;
	GetWindowRect(&windowRectangle);
	ScreenToClient(&windowRectangle);
	if(windowRectangle.PtInRect(pt)) {
		WTL::CRect lineRectangle;
		BOOL isOverText = FALSE;

		// Retrieve the bounding rectangle of the line that the point belongs to.
		POINT screenCoordinates = {x, y};
		ClientToScreen(&screenCoordinates);
		CComPtr<ITextRange> pTextRange = NULL;
		if(properties.pTextDocument && SUCCEEDED(properties.pTextDocument->RangeFromPoint(screenCoordinates.x, screenCoordinates.y, &pTextRange)) && pTextRange) {
			CComPtr<ITextRange> pLineRange = NULL;
			for(int i = 0; i < 2 && !isOverText && !pLineRange; i++) {
				if(SUCCEEDED(pTextRange->GetDuplicate(&pLineRange)) && pLineRange) {
					if(SUCCEEDED(pLineRange->StartOf(tomLine, tomMove, NULL)) && SUCCEEDED(pLineRange->EndOf(tomLine, tomExtend, NULL))) {
						if(SUCCEEDED(pLineRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomStart + TA_TOP + TA_LEFT), &lineRectangle.left, &lineRectangle.top)) && SUCCEEDED(pLineRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomEnd + TA_BOTTOM + TA_RIGHT), &lineRectangle.right, &lineRectangle.bottom))) {
							if(lineRectangle.PtInRect(pt)) {
								isOverText = TRUE;
							} else if(i < 2 && pt.y < lineRectangle.top) {
								// maybe the position is at the very end of the line, so that the rectangle is for the next line already
								LONG lineIndex1 = 0;
								LONG lineIndex2 = 0;
								if(SUCCEEDED(pTextRange->GetIndex(tomLine, &lineIndex1)) && SUCCEEDED(pTextRange->MoveStart(tomCharacter, -1, NULL)) && SUCCEEDED(pTextRange->GetIndex(tomLine, &lineIndex2)) && lineIndex2 < lineIndex1) {
									// Bingo!
									pLineRange = NULL;
								}
							}
						}
					}
				}
			}
		}

		BOOL isOverLink = FALSE;
		BOOL isOverSelection = FALSE;

		if(isOverText) {
			// the point is over text, so check for selection and for links
			CComPtr<ITextSelection> pTextSelection = NULL;
			if(pTextRange && SUCCEEDED(properties.pTextDocument->GetSelection(&pTextSelection))) {
				CHARRANGE selection;
				pTextSelection->GetStart(&selection.cpMin);
				pTextSelection->GetEnd(&selection.cpMax);

				// set isOverSelection
				CComPtr<ITextRange> pSelRange = NULL;
				if(SUCCEEDED(pTextSelection->GetDuplicate(&pSelRange)) && pSelRange && selection.cpMin != selection.cpMax && SUCCEEDED(pSelRange->MoveEnd(tomCharacter, -1, NULL))) {
					WTL::CRect selectionBoundingRectangle;
					if(SUCCEEDED(pSelRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomStart + TA_TOP + TA_LEFT), &selectionBoundingRectangle.left, &selectionBoundingRectangle.top)) && SUCCEEDED(pSelRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomStart + TA_BOTTOM + TA_RIGHT), &selectionBoundingRectangle.right, &selectionBoundingRectangle.bottom)) && selectionBoundingRectangle.PtInRect(pt)) {
						// the position is over the first selected character
						isOverSelection = TRUE;
					} else if(SUCCEEDED(pSelRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomEnd + TA_TOP + TA_LEFT), &selectionBoundingRectangle.left, &selectionBoundingRectangle.top)) && SUCCEEDED(pSelRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomEnd + TA_BOTTOM + TA_RIGHT), &selectionBoundingRectangle.right, &selectionBoundingRectangle.bottom)) && selectionBoundingRectangle.PtInRect(pt)) {
						// the position is over the last selected character
						isOverSelection = TRUE;
					} else if(selection.cpMin != selection.cpMax) {
						// the position might be within the selection
						CHARRANGE chr;
						pTextRange->GetStart(&chr.cpMin);
						pTextRange->GetEnd(&chr.cpMax);

						isOverSelection = (min(chr.cpMin, chr.cpMax) > min(selection.cpMin, selection.cpMax) && max(chr.cpMin, chr.cpMax) < max(selection.cpMin, selection.cpMax));
					}
				}

				// now set isOverLink - at first try to find out where the potential link starts and ends
				LONG hitLineIndex = 0;
				CComPtr<ITextRange> pLinkRange = NULL;
				if(SUCCEEDED(pTextRange->GetIndex(tomLine, &hitLineIndex)) && SUCCEEDED(pTextRange->GetDuplicate(&pLinkRange)) && pLinkRange && SUCCEEDED(pLinkRange->StartOf(tomCharFormat, tomExtend, NULL)) && SUCCEEDED(pLinkRange->EndOf(tomCharFormat, tomExtend, NULL))) {
					BOOL retrieveBoundingRectangle = FALSE;
					BOOL checkForLinkEffect = FALSE;
					WTL::CRect linkBoundingRectangle;
					for(int i = 0; i < 2 && !retrieveBoundingRectangle; i++) {
						LONG startLineIndex = 0;
						LONG endLineIndex = 0;
						CComPtr<ITextRange> pTextRangeCopy = NULL;
						if(SUCCEEDED(pTextRange->GetDuplicate(&pTextRangeCopy)) && pTextRangeCopy && i > 0) {
							pLinkRange = NULL;
							if(SUCCEEDED(pTextRangeCopy->MoveStart(tomCharacter, -i, NULL)) && SUCCEEDED(pTextRangeCopy->MoveEnd(tomCharacter, -i, NULL)) && SUCCEEDED(pTextRangeCopy->GetDuplicate(&pLinkRange)) && pLinkRange && SUCCEEDED(pLinkRange->StartOf(tomCharFormat, tomExtend, NULL)) && SUCCEEDED(pLinkRange->EndOf(tomCharFormat, tomExtend, NULL))) {
								//
							} else {
								break;
							}
						}
						CComPtr<ITextRange> p = NULL;
						if(SUCCEEDED(pLinkRange->GetIndex(tomLine, &startLineIndex)) && SUCCEEDED(pLinkRange->GetDuplicate(&p)) && p && SUCCEEDED(p->Collapse(tomEnd)) && SUCCEEDED(p->GetIndex(tomLine, &endLineIndex))) {
							if(startLineIndex < endLineIndex) {
								// multiple lines, but we are interested in the hit line only
								if(startLineIndex == hitLineIndex) {
									// we hit the first line, so move the end position to the line's end
									endLineIndex = 0;
									p = NULL;
									if(SUCCEEDED(pLinkRange->Collapse(tomStart)) && SUCCEEDED(pLinkRange->EndOf(tomLine, tomExtend, NULL)) && SUCCEEDED(pLinkRange->GetDuplicate(&p)) && p && SUCCEEDED(p->Collapse(tomEnd)) && SUCCEEDED(p->MoveStart(tomCharacter, -1, NULL)) && SUCCEEDED(p->GetIndex(tomLine, &endLineIndex))) {
										//
									}
								} else if(hitLineIndex == endLineIndex) {
									// we hit the last line, so move the start position to the line's start
									startLineIndex = 0;
									if(SUCCEEDED(pLinkRange->Collapse(tomEnd)) && SUCCEEDED(pLinkRange->StartOf(tomLine, tomExtend, NULL)) && SUCCEEDED(pLinkRange->GetIndex(tomLine, &startLineIndex))) {
										//
									}
								} else if(startLineIndex < hitLineIndex && hitLineIndex < endLineIndex) {
									// we hit a line inbetween, so move the start and end position to the line's boundaries
									pLinkRange = NULL;
									startLineIndex = 0;
									endLineIndex = 0;
									if(SUCCEEDED(pTextRangeCopy->GetDuplicate(&pLinkRange)) && pLinkRange) {
										p = NULL;
										if(SUCCEEDED(pLinkRange->StartOf(tomLine, tomExtend, NULL)) && SUCCEEDED(pLinkRange->GetIndex(tomLine, &startLineIndex)) && SUCCEEDED(pLinkRange->GetDuplicate(&p)) && p && SUCCEEDED(p->Collapse(tomEnd)) && SUCCEEDED(p->MoveStart(tomCharacter, -1, NULL)) && SUCCEEDED(p->GetIndex(tomLine, &endLineIndex))) {
											//
										}
									}
								}
							}

							if(startLineIndex == hitLineIndex && hitLineIndex == endLineIndex) {
								retrieveBoundingRectangle = TRUE;
							}

							if(retrieveBoundingRectangle) {
								if(SUCCEEDED(pLinkRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomStart + TA_TOP + TA_LEFT), &linkBoundingRectangle.left, &linkBoundingRectangle.top)) && SUCCEEDED(pLinkRange->GetPoint(tomClientCoord | tomAllowOffClient | (tomEnd + TA_BOTTOM + TA_RIGHT), &linkBoundingRectangle.right, &linkBoundingRectangle.bottom)) && linkBoundingRectangle.PtInRect(pt)) {
									// the position is over the potential link text range
									checkForLinkEffect = TRUE;
								} else if(pt.x < linkBoundingRectangle.left) {
									// this seems strange - maybe we are right at the end of the link text range
									// next loop iteration will go backward by one character and try again
									retrieveBoundingRectangle = FALSE;
								}
							}
						}
					}

					if(!checkForLinkEffect) {
						LONG s, e;
						pLinkRange->GetStart(&s);
						pLinkRange->GetEnd(&e);

						LONG hitLineIndex = 0;
						LONG startLineIndex = 0;
						LONG endLineIndex = 0;
						CComPtr<ITextRange> p = NULL;
						if(SUCCEEDED(pTextRange->GetIndex(tomLine, &hitLineIndex)) && SUCCEEDED(pLinkRange->GetIndex(tomLine, &startLineIndex)) && SUCCEEDED(pLinkRange->GetDuplicate(&p)) && p && SUCCEEDED(p->Collapse(tomEnd)) && SUCCEEDED(p->GetIndex(tomLine, &endLineIndex))) {
							if(startLineIndex < hitLineIndex && hitLineIndex < endLineIndex) {
								// the position is within the potential link text range
								checkForLinkEffect = TRUE;
							} else if(startLineIndex < endLineIndex) {
								if(startLineIndex == hitLineIndex) {
									// reduce the bounding rectangle to that of the first line of the potential link
									p = NULL;
									if(SUCCEEDED(pLinkRange->GetDuplicate(&p)) && p && SUCCEEDED(p->Collapse(tomStart)) && SUCCEEDED(p->EndOf(tomLine, tomExtend, NULL))) {
										if(SUCCEEDED(p->GetPoint(tomClientCoord | tomAllowOffClient | (tomEnd + TA_BOTTOM + TA_RIGHT), &linkBoundingRectangle.right, &linkBoundingRectangle.bottom)) && linkBoundingRectangle.PtInRect(pt)) {
											checkForLinkEffect = TRUE;
										}
									}
								} else if(hitLineIndex == endLineIndex) {
									// reduce the bounding rectangle to that of the last line of the potential link
									p = NULL;
									if(SUCCEEDED(pLinkRange->GetDuplicate(&p)) && p && SUCCEEDED(p->Collapse(tomEnd)) && SUCCEEDED(p->StartOf(tomLine, tomExtend, NULL))) {
										if(SUCCEEDED(p->GetPoint(tomClientCoord | tomAllowOffClient | (tomStart + TA_TOP + TA_LEFT), &linkBoundingRectangle.left, &linkBoundingRectangle.top)) && linkBoundingRectangle.PtInRect(pt)) {
											checkForLinkEffect = TRUE;
										}
									}
								}
							}
						}
					}

					if(checkForLinkEffect) {
						// at first check the beginning of the text range
						CComPtr<ITextFont> pTextFont = NULL;
						CComPtr<ITextFont2> pTextFont2 = NULL;
						if(SUCCEEDED(pLinkRange->GetFont(&pTextFont)) && pTextFont && SUCCEEDED(pTextFont->QueryInterface(IID_PPV_ARGS(&pTextFont2))) && pTextFont2) {
							long effects = 0;
							long mask = tomLink;
							if(SUCCEEDED(pTextFont2->GetEffects(&effects, &mask))) {
								isOverLink = (effects & tomLink);
							}
						} else {
							// save selection
							long flags = 0;
							pTextSelection->GetFlags(&flags);
							if(flags & tomSelStartActive) {
								LONG tmp = selection.cpMin;
								selection.cpMin = selection.cpMax;
								selection.cpMax = tmp;
							}

							CHARRANGE chr;
							pLinkRange->GetStart(&chr.cpMin);
							pLinkRange->GetEnd(&chr.cpMax);

							BOOL needsChangeSelection = (min(chr.cpMin, chr.cpMax) != min(selection.cpMin, selection.cpMax) || max(chr.cpMin, chr.cpMax) != max(selection.cpMin, selection.cpMax));
							BOOL didChangeSelection = FALSE;
							long freeze = 0;
							if(needsChangeSelection) {
								EnterSilentSelectionChangesSection();
								/* Unfortunately this also freezes the caret while moving the mouse cursor.
									* The reasons are:
									*   1) The WM_MOUSEMOVE message. This problem exists only if mouse events are enabled.
									*   2) The WM_SETCURSOR message. This problem exists only if LinkMousePointer is not mpDefault.
									*/
								if(SUCCEEDED(properties.pTextDocument->Freeze(&freeze))) {
									didChangeSelection = SUCCEEDED(pLinkRange->Select());
								}
							}

							if(!needsChangeSelection || didChangeSelection) {
								CHARFORMAT2 format;
								ZeroMemory(&format, sizeof(format));
								format.cbSize = sizeof(format);
								format.dwMask = CFM_LINK;
								SendMessage(EM_GETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&format));
								isOverLink = (format.dwEffects & CFE_LINK);
							}

							if(needsChangeSelection) {
								if(didChangeSelection) {
									// restore selection - pTextSelection->Select does not work
									SendMessage(EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&selection));
								}
								if(freeze > 0) {
									properties.pTextDocument->Unfreeze(&freeze);
								}
								LeaveSilentSelectionChangesSection();
							}
						}
					}
				}
			}
			if(isOverSelection) {
				ret = static_cast<HitTestConstants>(ret | htSelectedText);
			}
			if(isOverLink) {
				ret = static_cast<HitTestConstants>(ret | htLink);
			} else {
				ret = static_cast<HitTestConstants>(ret | htText);
			}
		} else {
			ret = static_cast<HitTestConstants>(ret | htBackground);
		}
	} else {
		if(x < windowRectangle.left) {
			ret = static_cast<HitTestConstants>(ret | htToLeft);
		} else if(x > windowRectangle.right) {
			ret = static_cast<HitTestConstants>(ret | htToRight);
		}
		if(y < windowRectangle.top) {
			ret = static_cast<HitTestConstants>(ret | htAbove);
		} else if(y > windowRectangle.bottom) {
			ret = static_cast<HitTestConstants>(ret | htBelow);
		}
	}
	return ret;
}

BOOL RichTextBox::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void RichTextBox::AutoScroll(void)
{
	LONG realScrollTimeBase = properties.dragScrollTimeBase;
	if(realScrollTimeBase == -1) {
		realScrollTimeBase = GetDoubleClickTime() / 4;
	}

	if((dragDropStatus.autoScrolling.currentHScrollVelocity != 0) && ((GetStyle() & WS_HSCROLL) == WS_HSCROLL)) {
		if(dragDropStatus.autoScrolling.currentHScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll to the left?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Left) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll left
					dragDropStatus.autoScrolling.lastScroll_Left = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINELEFT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll to the right?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Right) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll right
					dragDropStatus.autoScrolling.lastScroll_Right = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINERIGHT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}

	if((dragDropStatus.autoScrolling.currentVScrollVelocity != 0) && ((GetStyle() & WS_VSCROLL) == WS_VSCROLL)) {
		if(dragDropStatus.autoScrolling.currentVScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll upwardly?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Up) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll up
					dragDropStatus.autoScrolling.lastScroll_Up = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEUP, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll downwards?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Down) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll down
					dragDropStatus.autoScrolling.lastScroll_Down = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEDOWN, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}
}

BOOL RichTextBox::IsOverSelectedText(LPARAM pt)
{
	BOOL ret = FALSE;

	if(IsWindow()) {
		HitTestConstants hitTestDetails = HitTest(GET_X_LPARAM(pt), GET_Y_LPARAM(pt));
		ret = ((hitTestDetails & htSelectedText) == htSelectedText);
	}
	return ret;
}

BOOL RichTextBox::IsLeftMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	}
}

BOOL RichTextBox::IsRightMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	}
}


HRESULT RichTextBox::CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing)
{
	if(!hIcon || !pBits || !pWICImagingFactory) {
		return E_FAIL;
	}

	ICONINFO iconInfo;
	GetIconInfo(hIcon, &iconInfo);
	ATLASSERT(iconInfo.hbmColor);
	BITMAP bitmapInfo = {0};
	if(iconInfo.hbmColor) {
		GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bitmapInfo);
	} else if(iconInfo.hbmMask) {
		GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bitmapInfo);
	}
	bitmapInfo.bmHeight = abs(bitmapInfo.bmHeight);
	BOOL needsFrame = ((bitmapInfo.bmWidth < size.cx) || (bitmapInfo.bmHeight < size.cy));
	if(iconInfo.hbmColor) {
		::DeleteObject(iconInfo.hbmColor);
	}
	if(iconInfo.hbmMask) {
		::DeleteObject(iconInfo.hbmMask);
	}

	HRESULT hr = E_FAIL;

	CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
	if(!needsFrame) {
		hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapScaler);
	}
	if(needsFrame || SUCCEEDED(hr)) {
		CComPtr<IWICBitmap> pWICBitmapSource = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmapSource);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapSource);
		if(SUCCEEDED(hr)) {
			if(!needsFrame) {
				hr = pWICBitmapScaler->Initialize(pWICBitmapSource, size.cx, size.cy, WICBitmapInterpolationModeFant);
			}
			if(SUCCEEDED(hr)) {
				WICRect rc = {0};
				if(needsFrame) {
					rc.Height = bitmapInfo.bmHeight;
					rc.Width = bitmapInfo.bmWidth;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					LPRGBQUAD pIconBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, rc.Width * rc.Height * sizeof(RGBQUAD)));
					hr = pWICBitmapSource->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pIconBits));
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						// center the icon
						int xIconStart = (size.cx - bitmapInfo.bmWidth) / 2;
						int yIconStart = (size.cy - bitmapInfo.bmHeight) / 2;
						LPRGBQUAD pIconPixel = pIconBits;
						LPRGBQUAD pPixel = pBits;
						pPixel += yIconStart * size.cx;
						for(int y = yIconStart; y < yIconStart + bitmapInfo.bmHeight; ++y, pPixel += size.cx, pIconPixel += bitmapInfo.bmWidth) {
							CopyMemory(pPixel + xIconStart, pIconPixel, bitmapInfo.bmWidth * sizeof(RGBQUAD));
						}
						HeapFree(GetProcessHeap(), 0, pIconBits);

						rc.Height = size.cy;
						rc.Width = size.cx;

						// TODO: now draw a frame around it
					}
				} else {
					rc.Height = size.cy;
					rc.Width = size.cx;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					hr = pWICBitmapScaler->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pBits));
					ATLASSERT(SUCCEEDED(hr));

					if(SUCCEEDED(hr) && doAlphaChannelPostProcessing) {
						for(int i = 0; i < rc.Width * rc.Height; ++i, ++pBits) {
							if(pBits->rgbReserved == 0x00) {
								ZeroMemory(pBits, sizeof(RGBQUAD));
							}
						}
					}
				}
			} else {
				ATLASSERT(FALSE && "Bitmap scaler failed");
			}
		}
	}
	return hr;
}