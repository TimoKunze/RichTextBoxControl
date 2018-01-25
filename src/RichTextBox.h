//////////////////////////////////////////////////////////////////////
/// \class RichTextBox
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c RichEditx0W</em>
///
/// This class superclasses \c RichEditx0W and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo \c IMEFlags is the name of a struct as well as a variable.
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTextBox
/// \else
///   \sa RichTextBoxCtlLibA::IRichTextBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_IRichTextBoxEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "IMouseHookHandler.h"
#include "MouseHookProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "TextRange.h"
#include "TargetOLEDataObject.h"
#include "SourceOLEDataObject.h"


class ATL_NO_VTABLE RichTextBox : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IRichTextBox, &IID_IRichTextBox, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IRichTextBox, &IID_IRichTextBox, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<RichTextBox>,
    public IOleControlImpl<RichTextBox>,
    public IOleObjectImpl<RichTextBox>,
    public IOleInPlaceActiveObjectImpl<RichTextBox>,
    public IViewObjectExImpl<RichTextBox>,
    public IOleInPlaceObjectWindowlessImpl<RichTextBox>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<RichTextBox>,
    public Proxy_IRichTextBoxEvents<RichTextBox>,
    public IPersistStorageImpl<RichTextBox>,
    public IPersistPropertyBagImpl<RichTextBox>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<RichTextBox>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_RichTextBox, &__uuidof(_IRichTextBoxEvents), &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_RichTextBox, &__uuidof(_IRichTextBoxEvents), &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<RichTextBox>,
    public CComCoClass<RichTextBox, &CLSID_RichTextBox>,
    public CComControl<RichTextBox>,
    public IPerPropertyBrowsingImpl<RichTextBox>,
    public IDropTarget,
    public IDropSource,
    public IDropSourceNotify,
    public ICategorizeProperties,
    public ICreditsProvider,
    public IRichEditOleCallback,
    public IMouseHookHandler
{
	friend class OLEObject;
	friend class SourceOLEDataObject;
	friend class TextRange;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	///
	/// \sa ~RichTextBox
	RichTextBox();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used for cleaning up.
	///
	/// \sa RichTextBox
	~RichTextBox();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_IMEMODE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_RICHTEXTBOX)
		
		/*#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("RichTextBoxU"), NULL)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("RichTextBoxA"), NULL)
		#endif*/

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(RichTextBox)
			COM_INTERFACE_ENTRY(IRichTextBox)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IQuickActivate)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IDropTarget)
			COM_INTERFACE_ENTRY(IDropSource)
			COM_INTERFACE_ENTRY(IDropSourceNotify)
			COM_INTERFACE_ENTRY_IID(IID_IRichEditOleCallback, IRichEditOleCallback)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextDocument), 0, QueryITextDocumentInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextDocument2), 0, QueryITextDocumentInterface)
		END_COM_MAP()

		BEGIN_PROP_MAP(RichTextBox)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AcceptTabKey", DISPID_RTB_ACCEPTTABKEY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AdjustLineHeightForEastAsianLanguages", DISPID_RTB_ADJUSTLINEHEIGHTFOREASTASIANLANGUAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowInputThroughTSF", DISPID_RTB_ALLOWINPUTTHROUGHTSF, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowMathZoneInsertion", DISPID_RTB_ALLOWMATHZONEINSERTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowObjectInsertionThroughTSF", DISPID_RTB_ALLOWOBJECTINSERTIONTHROUGHTSF, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowTableInsertion", DISPID_RTB_ALLOWTABLEINSERTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowTSFProofingTips", DISPID_RTB_ALLOWTSFPROOFINGTIPS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AllowTSFSmartTags", DISPID_RTB_ALLOWTSFSMARTTAGS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AlwaysShowScrollBars", DISPID_RTB_ALWAYSSHOWSCROLLBARS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AlwaysShowSelection", DISPID_RTB_ALWAYSSHOWSELECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_RTB_APPEARANCE, CLSID_NULL, VT_I4)
			//PROP_ENTRY_TYPE("AutoDetectedURLSchemes", DISPID_RTB_AUTODETECTEDURLSCHEMES, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("AutoDetectURLs", DISPID_RTB_AUTODETECTURLS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutoScrolling", DISPID_RTB_AUTOSCROLLING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutoSelectWordOnTrackSelection", DISPID_RTB_AUTOSELECTWORDONTRACKSELECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("BackColor", DISPID_RTB_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("BeepOnMaxText", DISPID_RTB_BEEPONMAXTEXT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_RTB_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CharacterConversion", DISPID_RTB_CHARACTERCONVERSION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DefaultMathZoneHAlignment", DISPID_RTB_DEFAULTMATHZONEHALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DefaultTabWidth", DISPID_RTB_DEFAULTTABWIDTH, CLSID_NULL, VT_R4)
			PROP_ENTRY_TYPE("DenoteEmptyMathArguments", DISPID_RTB_DENOTEEMPTYMATHARGUMENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DetectDragDrop", DISPID_RTB_DETECTDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_RTB_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisableIMEOperations", DISPID_RTB_DISABLEIMEOPERATIONS, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("DiscardCompositionStringOnIMECancel", DISPID_RTB_DISCARDCOMPOSITIONSTRINGONIMECANCEL, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisplayHyperlinkTooltips", DISPID_RTB_DISPLAYHYPERLINKTOOLTIPS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisplayZeroWidthTableCellBorders", DISPID_RTB_DISPLAYZEROWIDTHTABLECELLBORDERS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DraftMode", DISPID_RTB_DRAFTMODE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragScrollTimeBase", DISPID_RTB_DRAGSCROLLTIMEBASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DropWordsOnWordBoundariesOnly", DISPID_RTB_DROPWORDSONWORDBOUNDARIESONLY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("EmulateSimpleTextBox", DISPID_RTB_EMULATESIMPLETEXTBOX, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Enabled", DISPID_RTB_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ExtendFontBackColorToWholeLine", DISPID_RTB_EXTENDFONTBACKCOLORTOWHOLELINE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_RTB_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("FormattingRectangleHeight", DISPID_RTB_FORMATTINGRECTANGLEHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("FormattingRectangleLeft", DISPID_RTB_FORMATTINGRECTANGLELEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("FormattingRectangleTop", DISPID_RTB_FORMATTINGRECTANGLETOP, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("FormattingRectangleWidth", DISPID_RTB_FORMATTINGRECTANGLEWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("GrowNAryOperators", DISPID_RTB_GROWNARYOPERATORS, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("HAlignment", DISPID_RTB_HALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HoverTime", DISPID_RTB_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HyphenationWordWrapZoneWidth", DISPID_RTB_HYPHENATIONWORDWRAPZONEWIDTH, CLSID_NULL, VT_I2)
			PROP_ENTRY_TYPE("IMEConversionMode", DISPID_RTB_IMECONVERSIONMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IMEMode", DISPID_RTB_IMEMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IntegralLimitsLocation", DISPID_RTB_INTEGRALLIMITSLOCATION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("LeftMargin", DISPID_RTB_LEFTMARGIN, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("LetClientHandleAllIMEOperations", DISPID_RTB_LETCLIENTHANDLEALLIMEOPERATIONS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("LinkMouseIcon", DISPID_RTB_LINKMOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("LinkMousePointer", DISPID_RTB_LINKMOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("LogicalCaret", DISPID_RTB_LOGICALCARET, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MathLineBreakBehavior", DISPID_RTB_MATHLINEBREAKBEHAVIOR, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MaxTextLength", DISPID_RTB_MAXTEXTLENGTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MaxUndoQueueSize", DISPID_RTB_MAXUNDOQUEUESIZE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Modified", DISPID_RTB_MODIFIED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_RTB_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_RTB_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MultiLine", DISPID_RTB_MULTILINE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MultiSelect", DISPID_RTB_MULTISELECT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("NAryLimitsLocation", DISPID_RTB_NARYLIMITSLOCATION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("NoInputSequenceCheck", DISPID_RTB_NOINPUTSEQUENCECHECK, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("OLEDragImageStyle", DISPID_RTB_OLEDRAGIMAGESTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_RTB_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RawSubScriptAndSuperScriptOperators", DISPID_RTB_RAWSUBSCRIPTANDSUPERSCRIPTOPERATORS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ReadOnly", DISPID_RTB_READONLY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_RTB_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RichEditVersion", DISPID_RTB_RICHEDITVERSION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RightMargin", DISPID_RTB_RIGHTMARGIN, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_RTB_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RightToLeftMathZones", DISPID_RTB_RIGHTTOLEFTMATHZONES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ScrollBars", DISPID_RTB_SCROLLBARS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ScrollToTopOnKillFocus", DISPID_RTB_SCROLLTOTOPONKILLFOCUS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ShowSelectionBar", DISPID_RTB_SHOWSELECTIONBAR, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("SmartSpacingOnDrop", DISPID_RTB_SMARTSPACINGONDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_RTB_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("SwitchFontOnIMEInput", DISPID_RTB_SWITCHFONTONIMEINPUT, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("Text", DISPID_RTB_TEXT, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("TextFlow", DISPID_RTB_TEXTFLOW, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("TextOrientation", DISPID_RTB_TEXTORIENTATION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("TSFModeBias", DISPID_RTB_TSFMODEBIAS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("UseBkAcetateColorForTextSelection", DISPID_RTB_USEBKACETATECOLORFORTEXTSELECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseBuiltInSpellChecking", DISPID_RTB_USEBUILTINSPELLCHECKING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseCustomFormattingRectangle", DISPID_RTB_USECUSTOMFORMATTINGRECTANGLE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseDualFontMode", DISPID_RTB_USEDUALFONTMODE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseSmallerFontForNestedFractions", DISPID_RTB_USESMALLERFONTFORNESTEDFRACTIONS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseTextServicesFramework", DISPID_RTB_USETEXTSERVICESFRAMEWORK, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("UseTouchKeyboardAutoCorrection", DISPID_RTB_USETOUCHKEYBOARDAUTOCORRECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseWindowsThemes", DISPID_RTB_USEWINDOWSTHEMES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ZoomRatio", DISPID_RTB_ZOOMRATIO, CLSID_NULL, VT_R8)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(RichTextBox)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTextBoxEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(RichTextBox)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_ERASEBKGND, OnEraseBkGnd)
			MESSAGE_HANDLER(WM_GETDLGCODE, OnGetDlgCode)
			MESSAGE_HANDLER(WM_HSCROLL, OnScroll)
			MESSAGE_HANDLER(WM_INPUTLANGCHANGE, OnInputLangChange)
			MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			MESSAGE_HANDLER(WM_PRINTCLIENT, OnPaint)
			MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETTEXT, OnSetText)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_VSCROLL, OnScroll)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(GetDragImageMessage(), OnGetDragImage)

			MESSAGE_HANDLER(EM_SETBKGNDCOLOR, OnSetBkgndColor)
			MESSAGE_HANDLER(EM_SETUNDOLIMIT, OnSetUndoLimit)

			REFLECTED_NOTIFY_CODE_HANDLER(EN_DRAGDROPDONE, OnDragDropDoneNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(EN_LINK, OnLinkNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(EN_REQUESTRESIZE, OnRequestResizeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(EN_SELCHANGE, OnSelChangeNotification)

			REFLECTED_COMMAND_CODE_HANDLER(EN_CHANGE, OnReflectedChange)
			REFLECTED_COMMAND_CODE_HANDLER(EN_HSCROLL, OnReflectedScroll)
			REFLECTED_COMMAND_CODE_HANDLER(EN_MAXTEXT, OnReflectedMaxText)
			REFLECTED_COMMAND_CODE_HANDLER(EN_SETFOCUS, OnReflectedSetFocus)
			REFLECTED_COMMAND_CODE_HANDLER(EN_VSCROLL, OnReflectedScroll)

			CHAIN_MSG_MAP(CComControl<RichTextBox>)
		END_MSG_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Will be called if this class is \c QueryInterface()'d for \c ITextDocument or \c ITextDocument2</em>
	///
	/// This method will be called if the class's \c QueryInterface() method is called with
	/// \c IID_ITextDocument or \c IID_ITextDocument2. We forward the call to the wrapped \c ITextDocument or
	/// \c ITextDocument2 implementation.
	///
	/// \param[in] pThis The instance of this class, that the interface is queried from.
	/// \param[in] queriedInterface Should be \c IID_ITextDocument or \c IID_ITextDocument2.
	/// \param[out] ppImplementation Receives the wrapped object's \c ITextDocument or \c ITextDocument2
	///             implementation.
	/// \param[in] cookie A \c DWORD value specified in the COM interface map.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774052.aspx">ITextDocument</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/hh768436.aspx">ITextDocument2</a>
	static HRESULT CALLBACK QueryITextDocumentInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Load to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Load and read directly from the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] pErrorLog The caller's \c IErrorLog implementation which will receive any errors
	///            that occur during property loading.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768206.aspx">IPersistPropertyBag::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog);
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Save to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Save and write directly into the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	/// \param[in] saveAllProperties Flag indicating whether the control should save all its properties
	///            (\c TRUE) or only those that have changed from the default value (\c FALSE).
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768207.aspx">IPersistPropertyBag::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTextBox
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AcceptTabKey property</em>
	///
	/// Retrieves whether pressing the [TAB] key inserts a tabulator into the control. If set to
	/// \c VARIANT_TRUE, a tabulator is inserted; otherwise the keyboard focus is transfered to the next
	/// control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AcceptTabKey, get_TabStops, get_TabWidth, get_Text, Raise_KeyDown
	virtual HRESULT STDMETHODCALLTYPE get_AcceptTabKey(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AcceptTabKey property</em>
	///
	/// Sets whether pressing the [TAB] key inserts a tabulator into the control. If set to
	/// \c VARIANT_TRUE, a tabulator is inserted; otherwise the keyboard focus is transfered to the next
	/// control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AcceptTabKey, put_TabStops, put_TabWidth, put_Text, Raise_KeyDown
	virtual HRESULT STDMETHODCALLTYPE put_AcceptTabKey(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AdjustLineHeightForEastAsianLanguages property</em>
	///
	/// Retrieves whether the line height is increased by 15%, if the line contains text in an east-asian
	/// language. If set to \c VARIANT_TRUE, the line height is increased; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa put_AdjustLineHeightForEastAsianLanguages, TextParagraph::get_LineSpacingRule,
	///     TextParagraph::get_SpaceBefore, TextParagraph::get_SpaceAfter
	virtual HRESULT STDMETHODCALLTYPE get_AdjustLineHeightForEastAsianLanguages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AdjustLineHeightForEastAsianLanguages property</em>
	///
	/// Sets whether the line height is increased by 15%, if the line contains text in an east-asian
	/// language. If set to \c VARIANT_TRUE, the line height is increased; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa get_AdjustLineHeightForEastAsianLanguages, TextParagraph::put_LineSpacingRule,
	///     TextParagraph::put_SpaceBefore, TextParagraph::put_SpaceAfter
	virtual HRESULT STDMETHODCALLTYPE put_AdjustLineHeightForEastAsianLanguages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowInputThroughTSF property</em>
	///
	/// Retrieves whether it is possible to insert text using Windows Text Services Framework (TSF).
	/// If set to \c VARIANT_TRUE, text can be inserted through TSF; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Usually input is disabled only temporarily to make sure the control's content is not
	///          locked by TSF.\n
	///          Requires Rich Edit 7.5 or newer.
	///
	/// \sa put_AllowInputThroughTSF, get_AllowObjectInsertionThroughTSF, get_UseTextServicesFramework,
	///     get_IMEMode
	virtual HRESULT STDMETHODCALLTYPE get_AllowInputThroughTSF(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowInputThroughTSF property</em>
	///
	/// Sets whether it is possible to insert text using Windows Text Services Framework (TSF).
	/// If set to \c VARIANT_TRUE, text can be inserted through TSF; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Usually input is disabled only temporarily to make sure the control's content is not
	///          locked by TSF.\n
	///          Requires Rich Edit 7.5 or newer.
	///
	/// \sa get_AllowInputThroughTSF, put_AllowObjectInsertionThroughTSF, put_UseTextServicesFramework,
	///     put_IMEMode
	virtual HRESULT STDMETHODCALLTYPE put_AllowInputThroughTSF(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowMathZoneInsertion property</em>
	///
	/// Retrieves whether it is possible to insert math zones, i.e. paragraphs that contain formulas.
	/// If set to \c VARIANT_TRUE, math zones can be inserted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property can be changed only before inserting any math zone.\n
	///          Requires Rich Edit 8.0 or newer.
	///
	/// \sa put_AllowMathZoneInsertion, TextRange::BuildUpMath, TextFont::get_IsMathZone
	virtual HRESULT STDMETHODCALLTYPE get_AllowMathZoneInsertion(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowMathZoneInsertion property</em>
	///
	/// Sets whether it is possible to insert math zones, i.e. paragraphs that contain formulas.
	/// If set to \c VARIANT_TRUE, math zones can be inserted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property can be changed only before inserting any math zone.\n
	///          Requires Rich Edit 8.0 or newer.
	///
	/// \sa get_AllowMathZoneInsertion, TextRange::BuildUpMath, TextFont::put_IsMathZone
	virtual HRESULT STDMETHODCALLTYPE put_AllowMathZoneInsertion(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowObjectInsertionThroughTSF property</em>
	///
	/// Retrieves whether it is possible to insert embedded objects using Windows Text Services Framework
	/// (TSF). If set to \c VARIANT_TRUE, embedded objects can be inserted through TSF; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowObjectInsertionThroughTSF, get_AllowInputThroughTSF, get_UseTextServicesFramework,
	///     get_IMEMode
	virtual HRESULT STDMETHODCALLTYPE get_AllowObjectInsertionThroughTSF(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowObjectInsertionThroughTSF property</em>
	///
	/// Sets whether it is possible to insert embedded objects using Windows Text Services Framework
	/// (TSF). If set to \c VARIANT_TRUE, embedded objects can be inserted through TSF; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowObjectInsertionThroughTSF, put_AllowInputThroughTSF, put_UseTextServicesFramework,
	///     put_IMEMode
	virtual HRESULT STDMETHODCALLTYPE put_AllowObjectInsertionThroughTSF(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowTableInsertion property</em>
	///
	/// Retrieves whether it is possible to insert tables. If set to \c VARIANT_TRUE, tables can be
	/// inserted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property can be changed only before inserting any table.\n
	///          Requires Rich Edit 7.5 or newer.
	///
	/// \sa put_AllowTableInsertion, TextRange::ReplaceWithTable
	virtual HRESULT STDMETHODCALLTYPE get_AllowTableInsertion(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowTableInsertion property</em>
	///
	/// Sets whether it is possible to insert tables. If set to \c VARIANT_TRUE, tables can be
	/// inserted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property can be changed only before inserting any table.\n
	///          Requires Rich Edit 7.5 or newer.
	///
	/// \sa get_AllowTableInsertion, TextRange::ReplaceWithTable
	virtual HRESULT STDMETHODCALLTYPE put_AllowTableInsertion(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowTSFProofingTips property</em>
	///
	/// Retrieves whether Windows Text Services Framework (TSF) may display proofing tips. If set to
	/// \c VARIANT_TRUE, proofing tips are allowed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowTSFProofingTips, get_AllowTSFSmartTags, get_UseTextServicesFramework
	virtual HRESULT STDMETHODCALLTYPE get_AllowTSFProofingTips(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowTSFProofingTips property</em>
	///
	/// Sets whether Windows Text Services Framework (TSF) may display proofing tips. If set to
	/// \c VARIANT_TRUE, proofing tips are allowed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowTSFProofingTips, put_AllowTSFSmartTags, put_UseTextServicesFramework
	virtual HRESULT STDMETHODCALLTYPE put_AllowTSFProofingTips(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AllowTSFSmartTags property</em>
	///
	/// Retrieves whether Windows Text Services Framework (TSF) may display SmartTag tips. If set to
	/// \c VARIANT_TRUE, SmartTag tips are allowed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowTSFSmartTags, get_AllowTSFProofingTips, get_UseTextServicesFramework
	virtual HRESULT STDMETHODCALLTYPE get_AllowTSFSmartTags(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowTSFSmartTags property</em>
	///
	/// Sets whether Windows Text Services Framework (TSF) may display SmartTag tips. If set to
	/// \c VARIANT_TRUE, SmartTag tips are allowed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowTSFSmartTags, put_AllowTSFProofingTips, put_UseTextServicesFramework
	virtual HRESULT STDMETHODCALLTYPE put_AllowTSFSmartTags(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AlwaysShowScrollBars property</em>
	///
	/// Retrieves whether the scroll bars are disabled instead of hidden if not needed. If set to
	/// \c VARIANT_TRUE, the scroll bars are disabled; otherwise they are hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_AlwaysShowScrollBars, get_ScrollBars
	virtual HRESULT STDMETHODCALLTYPE get_AlwaysShowScrollBars(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AlwaysShowScrollBars property</em>
	///
	/// Sets whether the scroll bars are disabled instead of hidden if not needed. If set to
	/// \c VARIANT_TRUE, the scroll bars are disabled; otherwise they are hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_AlwaysShowScrollBars, put_ScrollBars
	virtual HRESULT STDMETHODCALLTYPE put_AlwaysShowScrollBars(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AlwaysShowSelection property</em>
	///
	/// Retrieves whether the selected text will be highlighted even if the control doesn't have the focus.
	/// If set to \c VARIANT_TRUE, selected text is drawn as selected if the control does not have the focus;
	/// otherwise it's drawn as normal text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AlwaysShowSelection, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_AlwaysShowSelection(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AlwaysShowSelection property</em>
	///
	/// Sets whether the selected text will be highlighted even if the control doesn't have the focus.
	/// If set to \c VARIANT_TRUE, selected text is drawn as selected if the control does not have the focus;
	/// otherwise it's drawn as normal text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AlwaysShowSelection, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_AlwaysShowSelection(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Appearance property</em>
	///
	/// Retrieves the kind of border that is drawn around the control. Any of the values defined by
	/// the \c AppearanceConstants enumeration except \c aDefault is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Appearance, get_BorderStyle, get_UseWindowsThemes, RichTextBoxCtlLibU::AppearanceConstants
	/// \else
	///   \sa put_Appearance, get_BorderStyle, get_UseWindowsThemes, RichTextBoxCtlLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Appearance(AppearanceConstants* pValue);
	/// \brief <em>Sets the \c Appearance property</em>
	///
	/// Sets the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration except \c aDefault is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Appearance, put_BorderStyle, put_UseWindowsThemes, RichTextBoxCtlLibU::AppearanceConstants
	/// \else
	///   \sa get_Appearance, put_BorderStyle, put_UseWindowsThemes, RichTextBoxCtlLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Appearance(AppearanceConstants newValue);
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c AutoDetectedURLSchemes property</em>
	///
	/// Retrieves the URL schemes recognized if the \c AutoDetectURLs property contains the \c aduURLs flag.
	/// If set to an empty string, the system's default scheme name list is used.\n
	/// The schemes have to include the colon (":") and are simply concatenated, e.g.
	/// "news:http:ftp:telnet:".
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AutoDetectedURLSchemes, get_AutoDetectURLs, TextRange::get_URL
	virtual HRESULT STDMETHODCALLTYPE get_AutoDetectedURLSchemes(BSTR* pValue);
	/// \brief <em>Sets the \c AutoDetectedURLSchemes property</em>
	///
	/// Sets the URL schemes recognized if the \c AutoDetectURLs property contains the \c aduURLs flag.
	/// If set to an empty string, the system's default scheme name list is used.\n
	/// The schemes have to include the colon (":") and are simply concatenated, e.g.
	/// "news:http:ftp:telnet:".
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AutoDetectedURLSchemes, put_AutoDetectURLs, TextRange::put_URL
	virtual HRESULT STDMETHODCALLTYPE put_AutoDetectedURLSchemes(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c AutoDetectURLs property</em>
	///
	/// Retrieves whether and how the user input is parsed to detect URLs, email addresses, phone
	/// numbers etc. If such content is found in the user input, it is formatted accordingly. Any
	/// combination of the values defined by the \c AutoDetectURLsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_AutoDetectURLs, get_AutoDetectedURLSchemes, get_DisplayHyperlinkTooltips,
	///       TextRange::get_URL, RichTextBoxCtlLibU::AutoDetectURLsConstants
	/// \else
	///   \sa put_AutoDetectURLs, get_AutoDetectedURLSchemes, get_DisplayHyperlinkTooltips,
	///       TextRange::get_URL, RichTextBoxCtlLibA::AutoDetectURLsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AutoDetectURLs(AutoDetectURLsConstants* pValue);
	/// \brief <em>Sets the \c AutoDetectURLs property</em>
	///
	/// Sets whether and how the user input is parsed to detect URLs, email addresses, phone
	/// numbers etc. If such content is found in the user input, it is formatted accordingly. Any
	/// combination of the values defined by the \c AutoDetectURLsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_AutoDetectURLs, put_AutoDetectedURLSchemes, put_DisplayHyperlinkTooltips,
	///       TextRange::put_URL, RichTextBoxCtlLibU::AutoDetectURLsConstants
	/// \else
	///   \sa get_AutoDetectURLs, put_AutoDetectedURLSchemes, put_DisplayHyperlinkTooltips,
	///       TextRange::put_URL, RichTextBoxCtlLibA::AutoDetectURLsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_AutoDetectURLs(AutoDetectURLsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c AutoScrolling property</em>
	///
	/// Retrieves the directions into which the control scrolls automatically, if the caret reaches the
	/// borders of the control's client area. Any combination of the values defined by the
	/// \c AutoScrollingConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_AutoScrolling, get_ScrollBars, get_MultiLine, Raise_TruncatedText,
	///       RichTextBoxCtlLibU::AutoScrollingConstants
	/// \else
	///   \sa put_AutoScrolling, get_ScrollBars, get_MultiLine, Raise_TruncatedText,
	///       RichTextBoxCtlLibA::AutoScrollingConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AutoScrolling(AutoScrollingConstants* pValue);
	/// \brief <em>Sets the \c AutoScrolling property</em>
	///
	/// Sets the directions into which the control scrolls automatically, if the caret reaches the
	/// borders of the control's client area. Any combination of the values defined by the
	/// \c AutoScrollingConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_AutoScrolling, put_ScrollBars, put_MultiLine, Raise_TruncatedText,
	///       RichTextBoxCtlLibU::AutoScrollingConstants
	/// \else
	///   \sa get_AutoScrolling, put_ScrollBars, put_MultiLine, Raise_TruncatedText,
	///       RichTextBoxCtlLibA::AutoScrollingConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_AutoScrolling(AutoScrollingConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c AutoSelectWordOnTrackSelection property</em>
	///
	/// Retrieves whether whole words are selected, if the user selects text using the mouse. If set to
	/// \c VARIANT_TRUE, track-selection selects whole words; otherwise the selection ends at the current
	/// mouse position which can be within a word.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa put_AutoSelectWordOnTrackSelection, get_DropWordsOnWordBoundariesOnly, Raise_DblClick
	virtual HRESULT STDMETHODCALLTYPE get_AutoSelectWordOnTrackSelection(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoSelectWordOnTrackSelection property</em>
	///
	/// Sets whether whole words are selected, if the user selects text using the mouse. If set to
	/// \c VARIANT_TRUE, track-selection selects whole words; otherwise the selection ends at the current
	/// mouse position which can be within a word.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa get_AutoSelectWordOnTrackSelection, put_DropWordsOnWordBoundariesOnly, Raise_DblClick
	virtual HRESULT STDMETHODCALLTYPE put_AutoSelectWordOnTrackSelection(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BackColor, TextFont::get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BackColor, TextFont::put_BackColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BeepOnMaxText property</em>
	///
	/// Retrieves whether the control plays the system beep sound if the user tries to enter more characters
	/// than allowed by the \c MaxTextLength property. If set to \c VARIANT_TRUE, the system beep is played;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BeepOnMaxText, get_MaxTextLength
	virtual HRESULT STDMETHODCALLTYPE get_BeepOnMaxText(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c BeepOnMaxText property</em>
	///
	/// Sets whether the control plays the system beep sound if the user tries to enter more characters
	/// than allowed by the \c MaxTextLength property. If set to \c VARIANT_TRUE, the system beep is played;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BeepOnMaxText, put_MaxTextLength
	virtual HRESULT STDMETHODCALLTYPE put_BeepOnMaxText(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of inner border that is drawn around the control. Any of the values defined
	/// by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BorderStyle, get_Appearance, get_UseWindowsThemes, RichTextBoxCtlLibU::BorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, get_Appearance, get_UseWindowsThemes, RichTextBoxCtlLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(BorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of inner border that is drawn around the control. Any of the values defined by
	/// the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BorderStyle, put_Appearance, put_UseWindowsThemes, RichTextBoxCtlLibU::BorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, put_Appearance, put_UseWindowsThemes, RichTextBoxCtlLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(BorderStyleConstants newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c CharacterConversion property</em>
	///
	/// Retrieves the kind of conversion that is applied to characters that are typed into the control. Any
	/// of the values defined by the \c CharacterConversionConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_CharacterConversion, get_Text, RichTextBoxCtlLibU::CharacterConversionConstants
	/// \else
	///   \sa put_CharacterConversion, get_Text, RichTextBoxCtlLibA::CharacterConversionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CharacterConversion(CharacterConversionConstants* pValue);
	/// \brief <em>Sets the \c CharacterConversion property</em>
	///
	/// Sets the kind of conversion that is applied to characters that are typed into the control. Any
	/// of the values defined by the \c CharacterConversionConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_CharacterConversion, put_Text, RichTextBoxCtlLibU::CharacterConversionConstants
	/// \else
	///   \sa get_CharacterConversion, put_Text, RichTextBoxCtlLibA::CharacterConversionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_CharacterConversion(CharacterConversionConstants newValue);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c CurrentIMECompositionMode property</em>
	///
	/// Retrieves the control's current IME composition mode. IME (Input Method Editor) is a Windows feature
	/// making it easy to enter Asian characters. Any of the values defined by the
	/// \c IMECompositionModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa GetCurrentIMECompositionText, get_IMEMode, get_UseTextServicesFramework,
	///       RichTextBoxCtlLibU::IMECompositionModeConstants
	/// \else
	///   \sa GetCurrentIMECompositionText, get_IMEMode, get_UseTextServicesFramework,
	///       RichTextBoxCtlLibA::IMECompositionModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CurrentIMECompositionMode(IMECompositionModeConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c DefaultMathZoneHAlignment property</em>
	///
	/// Retrieves the default horizontal alignment of math zones. Some of the values defined by the
	/// \c HAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_DefaultMathZoneHAlignment, TextParagraph::get_HAlignment, TextFont::get_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa put_DefaultMathZoneHAlignment, TextParagraph::get_HAlignment, TextFont::get_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DefaultMathZoneHAlignment(HAlignmentConstants* pValue);
	/// \brief <em>Sets the \c DefaultMathZoneHAlignment property</em>
	///
	/// Sets the default horizontal alignment of math zones. Some of the values defined by the
	/// \c HAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_DefaultMathZoneHAlignment, TextParagraph::put_HAlignment, TextFont::put_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa get_DefaultMathZoneHAlignment, TextParagraph::put_HAlignment, TextFont::put_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DefaultMathZoneHAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DefaultTabWidth property</em>
	///
	/// Retrieves the width (in points) of a tabulator character. It is used if no tab stop is specified
	/// for the paragraph.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The system's default tab width is 36.0 points.
	///
	/// \sa put_DefaultTabWidth, TextParagraph::get_Tabs
	virtual HRESULT STDMETHODCALLTYPE get_DefaultTabWidth(FLOAT* pValue);
	/// \brief <em>Sets the \c DefaultTabWidth property</em>
	///
	/// Sets the width (in points) of a tabulator character. It is used if no tab stop is specified
	/// for the paragraph.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DefaultTabWidth, TextParagraph::get_Tabs
	virtual HRESULT STDMETHODCALLTYPE put_DefaultTabWidth(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c DenoteEmptyMathArguments property</em>
	///
	/// Retrieves which kinds of empty arguments in a math zone are denoted by a dotted square. Any of the
	/// values defined by the \c DenoteEmptyMathArgumentsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_DenoteEmptyMathArguments, TextFont::get_IsMathZone, TextRange::BuildUpMath,
	///       RichTextBoxCtlLibU::DenoteEmptyMathArgumentsConstants
	/// \else
	///   \sa put_DenoteEmptyMathArguments, TextFont::get_IsMathZone, TextRange::BuildUpMath,
	///       RichTextBoxCtlLibA::DenoteEmptyMathArgumentsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DenoteEmptyMathArguments(DenoteEmptyMathArgumentsConstants* pValue);
	/// \brief <em>Sets the \c DenoteEmptyMathArguments property</em>
	///
	/// Sets which kinds of empty arguments in a math zone are denoted by a dotted square. Any of the
	/// values defined by the \c DenoteEmptyMathArgumentsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_DenoteEmptyMathArguments, TextFont::put_IsMathZone, TextRange::BuildUpMath,
	///       RichTextBoxCtlLibU::DenoteEmptyMathArgumentsConstants
	/// \else
	///   \sa get_DenoteEmptyMathArguments, TextFont::put_IsMathZone, TextRange::BuildUpMath,
	///       RichTextBoxCtlLibA::DenoteEmptyMathArgumentsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DenoteEmptyMathArguments(DenoteEmptyMathArgumentsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DetectDragDrop property</em>
	///
	/// Retrieves whether <b>advanced</b> drag'n'drop mode can be entered. If set to \c VARIANT_TRUE,
	/// <b>advanced</b> drag'n'drop mode can be entered by pressing the left or right mouse button over
	/// selected text and then moving the mouse with the button still pressed. If set to \c VARIANT_FALSE,
	/// <b>advanced</b> drag'n'drop mode is not available.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This affects advanced drag'n'drop mode only. Built-in drag'n'drop is not affected by this
	///          property.
	///
	/// \sa put_DetectDragDrop, get_RegisterForOLEDragDrop, get_DragScrollTimeBase, Raise_BeginDrag,
	///     Raise_BeginRDrag
	virtual HRESULT STDMETHODCALLTYPE get_DetectDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DetectDragDrop property</em>
	///
	/// Sets whether <b>advanced</b> drag'n'drop mode can be entered. If set to \c VARIANT_TRUE,
	/// <b>advanced</b> drag'n'drop mode can be entered by pressing the left or right mouse button over
	/// selected text and then moving the mouse with the button still pressed. If set to \c VARIANT_FALSE,
	/// <b>advanced</b> drag'n'drop mode is not available.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This affects advanced drag'n'drop mode only. Built-in drag'n'drop is not affected by this
	///          property.
	///
	/// \sa get_DetectDragDrop, put_RegisterForOLEDragDrop, put_DragScrollTimeBase, Raise_BeginDrag,
	///     Raise_BeginRDrag
	virtual HRESULT STDMETHODCALLTYPE put_DetectDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisabledEvents, RichTextBoxCtlLibU::DisabledEventsConstants
	/// \else
	///   \sa put_DisabledEvents, RichTextBoxCtlLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisabledEvents, RichTextBoxCtlLibU::DisabledEventsConstants
	/// \else
	///   \sa get_DisabledEvents, RichTextBoxCtlLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisableIMEOperations property</em>
	///
	/// Retrieves whether IME operation is enabled. IME (Input Method Editor) is a Windows feature making
	/// it easy to enter Asian characters. On east-asian systems it can be explicitly disabled by setting
	/// this property to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DisableIMEOperations, get_LetClientHandleAllIMEOperations, get_IMEMode, get_IMEConversionMode
	virtual HRESULT STDMETHODCALLTYPE get_DisableIMEOperations(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisableIMEOperations property</em>
	///
	/// Sets whether IME operation is enabled. IME (Input Method Editor) is a Windows feature making
	/// it easy to enter Asian characters. On east-asian systems it can be explicitly disabled by setting
	/// this property to \c VARIANT_FALSE.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DisableIMEOperations, put_LetClientHandleAllIMEOperations, put_IMEMode, put_IMEConversionMode
	virtual HRESULT STDMETHODCALLTYPE put_DisableIMEOperations(VARIANT_BOOL newValue);
	// \brief <em>Retrieves the current setting of the \c DiscardCompositionStringOnIMECancel property</em>
	//
	// Retrieves whether the current composition string of an IME (Input Method Editor) is discarded, if the
	// user cancels it. If set to \c VARIANT_TRUE, the composition string is discarded; otherwise it is used
	// as result string.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \sa put_DiscardCompositionStringOnIMECancel, get_IMEMode
	//virtual HRESULT STDMETHODCALLTYPE get_DiscardCompositionStringOnIMECancel(VARIANT_BOOL* pValue);
	// \brief <em>Sets the \c DiscardCompositionStringOnIMECancel property</em>
	//
	// Sets whether the current composition string of an IME (Input Method Editor) is discarded, if the
	// user cancels it. If set to \c VARIANT_TRUE, the composition string is discarded; otherwise it is used
	// as result string.
	//
	// \param[in] newValue The setting to apply.
	//
	// \return An \c HRESULT error code.
	//
	// \sa get_DiscardCompositionStringOnIMECancel, put_IMEMode
	//virtual HRESULT STDMETHODCALLTYPE put_DiscardCompositionStringOnIMECancel(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayHyperlinkTooltips property</em>
	///
	/// Retrieves whether a tool tip with the link target is displayed if the mouse cursor is located
	/// over a link. If set to \c VARIANT_TRUE, tool tips are displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property from \c VARIANT_TRUE to \c VARIANT_FALSE destroys and recreates the
	///            control window.
	///
	/// \remarks Requires Rich Edit 5.0 or newer.
	///
	/// \sa put_DisplayHyperlinkTooltips, get_AutoDetectURLs, get_AutoDetectedURLSchemes
	virtual HRESULT STDMETHODCALLTYPE get_DisplayHyperlinkTooltips(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayHyperlinkTooltips property</em>
	///
	/// Sets whether a tool tip with the link target is displayed if the mouse cursor is located
	/// over a link. If set to \c VARIANT_TRUE, tool tips are displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property from \c VARIANT_TRUE to \c VARIANT_FALSE destroys and recreates the
	///            control window.
	///
	/// \remarks Requires Rich Edit 5.0 or newer.
	///
	/// \sa get_DisplayHyperlinkTooltips, put_AutoDetectURLs, put_AutoDetectedURLSchemes
	virtual HRESULT STDMETHODCALLTYPE put_DisplayHyperlinkTooltips(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayZeroWidthTableCellBorders property</em>
	///
	/// Retrieves whether table cell borders are displayed as thin lines if the border width is set to 0.
	/// If set to \c VARIANT_TRUE, such borders are displayed as thin lines; otherwise they are not displayed
	/// at all.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_DisplayZeroWidthTableCellBorders, TextRange::ReplaceWithTable
	virtual HRESULT STDMETHODCALLTYPE get_DisplayZeroWidthTableCellBorders(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayZeroWidthTableCellBorders property</em>
	///
	/// Sets whether table cell borders are displayed as thin lines if the border width is set to 0.
	/// If set to \c VARIANT_TRUE, such borders are displayed as thin lines; otherwise they are not displayed
	/// at all.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_DisplayZeroWidthTableCellBorders, TextRange::ReplaceWithTable
	virtual HRESULT STDMETHODCALLTYPE put_DisplayZeroWidthTableCellBorders(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DocumentFont property</em>
	///
	/// Retrieves the document's default font settings.
	///
	/// \param[out] ppTextFont The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa put_DocumentFont, TextFont, get_DocumentParagraph, TextRange::get_Font
	virtual HRESULT STDMETHODCALLTYPE get_DocumentFont(IRichTextFont** ppTextFont);
	/// \brief <em>Sets the \c DocumentFont property</em>
	///
	/// Sets the document's default font settings.
	///
	/// \param[in] pNewTextFont The new styling object's \c IRichTextFont implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa get_DocumentFont, TextFont, put_DocumentParagraph, TextRange::put_Font
	virtual HRESULT STDMETHODCALLTYPE put_DocumentFont(IRichTextFont* pNewTextFont);
	/// \brief <em>Retrieves the current setting of the \c DocumentParagraph property</em>
	///
	/// Retrieves the document's default paragraph settings.
	///
	/// \param[out] ppTextParagraph The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DocumentParagraph, TextParagraph, get_DocumentFont, TextRange::get_Paragraph
	virtual HRESULT STDMETHODCALLTYPE get_DocumentParagraph(IRichTextParagraph** ppTextParagraph);
	/// \brief <em>Sets the \c DocumentParagraph property</em>
	///
	/// Sets the document's default paragraph settings.
	///
	/// \param[in] pNewTextParagraph The new styling object's \c IRichTextParagraph implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DocumentParagraph, TextParagraph, put_DocumentFont, TextRange::put_Paragraph
	virtual HRESULT STDMETHODCALLTYPE put_DocumentParagraph(IRichTextParagraph* pNewTextParagraph);
	/// \brief <em>Retrieves the current setting of the \c DraftMode property</em>
	///
	/// Retrieves whether a single font is used to draw all content. The font is determined by the system
	/// setting for the font used in message boxes. For example, accessible users may read text easier if it
	/// is uniform, rather than a mix of fonts and styles.\n
	/// If set to \c VARIANT_TRUE, a uniform font is used; otherwise the used fonts are specified by the
	/// document.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DraftMode, TextFont::get_Name
	virtual HRESULT STDMETHODCALLTYPE get_DraftMode(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DraftMode property</em>
	///
	/// Sets whether a single font is used to draw all content. The font is determined by the system
	/// setting for the font used in message boxes. For example, accessible users may read text easier if it
	/// is uniform, rather than a mix of fonts and styles.\n
	/// If set to \c VARIANT_TRUE, a uniform font is used; otherwise the used fonts are specified by the
	/// document.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DraftMode, TextFont::put_Name
	virtual HRESULT STDMETHODCALLTYPE put_DraftMode(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DraggedTextRange property</em>
	///
	/// Retrieves a \c TextRange object for the currently dragged text.
	///
	/// \param[out] ppTextRange The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TextRange, OLEDrag, get_TextRange
	virtual HRESULT STDMETHODCALLTYPE get_DraggedTextRange(IRichTextRange** ppTextRange);
	/// \brief <em>Retrieves the current setting of the \c DragScrollTimeBase property</em>
	///
	/// Retrieves the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time, divided by 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_DragScrollTimeBase, get_DetectDragDrop, get_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragScrollTimeBase(LONG* pValue);
	/// \brief <em>Sets the \c DragScrollTimeBase property</em>
	///
	/// Sets the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time divided by 4 is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_DragScrollTimeBase, put_DetectDragDrop, put_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragScrollTimeBase(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DropWordsOnWordBoundariesOnly property</em>
	///
	/// Retrieves whether the insertion mark is placed on word boundaries only, if the dragged text range is
	/// a whole word. If set to \c VARIANT_TRUE, words can be dropped on word boundaries only; otherwise they
	/// can be dropped within other words.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa put_DropWordsOnWordBoundariesOnly, get_SmartSpacingOnDrop, get_RegisterForOLEDragDrop,
	///     get_AutoSelectWordOnTrackSelection
	virtual HRESULT STDMETHODCALLTYPE get_DropWordsOnWordBoundariesOnly(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DropWordsOnWordBoundariesOnly property</em>
	///
	/// Sets whether the insertion mark is placed on word boundaries only, if the dragged text range is
	/// a whole word. If set to \c VARIANT_TRUE, words can be dropped on word boundaries only; otherwise they
	/// can be dropped within other words.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa get_DropWordsOnWordBoundariesOnly, put_SmartSpacingOnDrop, put_RegisterForOLEDragDrop,
	///     put_AutoSelectWordOnTrackSelection
	virtual HRESULT STDMETHODCALLTYPE put_DropWordsOnWordBoundariesOnly(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c EffectColor property</em>
	///
	/// Retrieves the colors that the control uses to draw some text effects. For instance these colors can
	/// be used for text underlines.\n
	/// Currently 16 colors can be defined.
	///
	/// \param[in] index The one-based index of the color to retrieve.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa put_EffectColor, TextFont::get_UndelineColorIndex, TextFont::get_UnderlineType,
	///     TextFont::get_UnderlinePosition, TextFont::get_ForeColor
	virtual HRESULT STDMETHODCALLTYPE get_EffectColor(LONG index, OLE_COLOR* pValue);
	/// \brief <em>Sets the \c EffectColor property</em>
	///
	/// Sets the colors that the control uses to draw some text effects. For instance these colors can
	/// be used for text underlines.\n
	/// Currently 16 colors can be defined.
	///
	/// \param[in] index The one-based index of the color to set.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa get_EffectColor, TextFont::put_UndelineColorIndex, TextFont::put_UnderlineType,
	///     TextFont::put_UnderlinePosition, TextFont::put_ForeColor
	virtual HRESULT STDMETHODCALLTYPE put_EffectColor(LONG index, OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c EmbeddedObjects property</em>
	///
	/// Retrieves a collection of all OLE objects embedded in the document.
	///
	/// \param[out] ppOLEObjects The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa OLEObjects, TextRange::get_EmbeddedObject
	virtual HRESULT STDMETHODCALLTYPE get_EmbeddedObjects(IRichOLEObjects** ppOLEObjects);
	/// \brief <em>Retrieves the current setting of the \c EmulateSimpleTextBox property</em>
	///
	/// Retrieves whether the control deactivates a couple of features like most keyboard shortcuts and
	/// pasting rich text in order to give the client application much more control over the available rich
	/// edit features. If set to \c VARIANT_TRUE, the control disables some rich edit features; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_EmulateSimpleTextBox
	virtual HRESULT STDMETHODCALLTYPE get_EmulateSimpleTextBox(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c EmulateSimpleTextBox property</em>
	///
	/// Sets whether the control deactivates a couple of features like most keyboard shortcuts and
	/// pasting rich text in order to give the client application much more control over the available rich
	/// edit features. If set to \c VARIANT_TRUE, the control disables some rich edit features; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_EmulateSimpleTextBox
	virtual HRESULT STDMETHODCALLTYPE put_EmulateSimpleTextBox(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled, get_ReadOnly
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the control for user input. If set to \c VARIANT_TRUE, the control reacts
	/// to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled, put_ReadOnly
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ExtendFontBackColorToWholeLine property</em>
	///
	/// Retrieves whether the background color of the first character in a line of text is used as the
	/// background color of the whole line. If set to \c VARIANT_TRUE, the first character's background color
	/// is used as background color of the whole line. If subsequent text ranges have different background
	/// colors, those are applied to the corresponding text range only. If set to \c VARIANT_FALSE, the
	/// line's background color remains transparent.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ExtendFontBackColorToWholeLine, TextFont::get_BackColor, get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_ExtendFontBackColorToWholeLine(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ExtendFontBackColorToWholeLine property</em>
	///
	/// Sets whether the background color of the first character in a line of text is used as the
	/// background color of the whole line. If set to \c VARIANT_TRUE, the first character's background color
	/// is used as background color of the whole line. If subsequent text ranges have different background
	/// colors, those are applied to the corresponding text range only. If set to \c VARIANT_FALSE, the
	/// line's background color remains transparent.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ExtendFontBackColorToWholeLine, TextFont::put_BackColor, put_BackColor
	virtual HRESULT STDMETHODCALLTYPE put_ExtendFontBackColorToWholeLine(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FirstVisibleLine property</em>
	///
	/// Retrieves the zero-based index of the uppermost visible line in a multiline control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MultiLine, GetLineCount
	virtual HRESULT STDMETHODCALLTYPE get_FirstVisibleLine(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the control's content.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, putref_Font, get_Text, get_ZoomRatio,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.
	///
	/// \sa get_Font, putref_Font, put_Text, put_ZoomRatio,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Font, put_Font, put_Text, put_ZoomRatio,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleHeight property</em>
	///
	/// Retrieves the height (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_FormattingRectangleHeight, get_FormattingRectangleLeft, get_FormattingRectangleTop,
	///     get_FormattingRectangleWidth, get_UseCustomFormattingRectangle, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c FormattingRectangleHeight property</em>
	///
	/// Sets the height (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_FormattingRectangleHeight, put_FormattingRectangleLeft, put_FormattingRectangleTop,
	///     put_FormattingRectangleWidth, put_UseCustomFormattingRectangle, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_FormattingRectangleHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleLeft property</em>
	///
	/// Retrieves the distance (in pixels) between the left borders of the control's formatting rectangle and
	/// its client area.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_FormattingRectangleLeft, get_FormattingRectangleHeight, get_FormattingRectangleTop,
	///     get_FormattingRectangleWidth, get_UseCustomFormattingRectangle, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleLeft(OLE_XPOS_PIXELS* pValue);
	/// \brief <em>Sets the \c FormattingRectangleLeft property</em>
	///
	/// Retrieves the distance (in pixels) between the left borders of the control's formatting rectangle and
	/// its client area.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_FormattingRectangleLeft, put_FormattingRectangleHeight, put_FormattingRectangleTop,
	///     put_FormattingRectangleWidth, put_UseCustomFormattingRectangle, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_FormattingRectangleLeft(OLE_XPOS_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleTop property</em>
	///
	/// Retrieves the distance (in pixels) between the upper borders of the control's formatting rectangle
	/// and its client area.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_FormattingRectangleTop, get_FormattingRectangleHeight, get_FormattingRectangleLeft,
	///     get_FormattingRectangleWidth, get_UseCustomFormattingRectangle, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleTop(OLE_YPOS_PIXELS* pValue);
	/// \brief <em>Sets the \c FormattingRectangleTop property</em>
	///
	/// Sets the distance (in pixels) between the upper borders of the control's formatting rectangle
	/// and its client area.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_FormattingRectangleTop, put_FormattingRectangleHeight, put_FormattingRectangleLeft,
	///     put_FormattingRectangleWidth, put_UseCustomFormattingRectangle, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_FormattingRectangleTop(OLE_YPOS_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FormattingRectangleWidth property</em>
	///
	/// Retrieves the width (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_FormattingRectangleWidth, get_FormattingRectangleHeight, get_FormattingRectangleLeft,
	///     get_FormattingRectangleTop, get_UseCustomFormattingRectangle, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_FormattingRectangleWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c FormattingRectangleWidth property</em>
	///
	/// Sets the width (in pixels) of the control's formatting rectangle.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for formatting
	/// the text displayed in the window rectangle. When the control is first displayed, the two rectangles
	/// are identical on the screen. An application can make the formatting rectangle larger than the window
	/// rectangle (thereby limiting the visibility of the control's text) or smaller than the window
	/// rectangle (thereby creating extra white space around the text).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_FormattingRectangleWidth, put_FormattingRectangleHeight, put_FormattingRectangleLeft,
	///     put_FormattingRectangleTop, put_UseCustomFormattingRectangle, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_FormattingRectangleWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c GrowNAryOperators property</em>
	///
	/// Retrieves whether n-ary operators (e.g. ∑) in math-zones are displayed in a font size that matches
	/// the height of the operand. If set to \c VARIANT_TRUE, those operators grow with the operand's height;
	/// otherwise they remain in the initial size.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa put_GrowNAryOperators, TextFont::get_IsMathZone, TextRange::BuildUpMath
	virtual HRESULT STDMETHODCALLTYPE get_GrowNAryOperators(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c GrowNAryOperators property</em>
	///
	/// Sets whether n-ary operators (e.g. ∑) in math-zones are displayed in a font size that matches
	/// the height of the operand. If set to \c VARIANT_TRUE, those operators grow with the operand's height;
	/// otherwise they remain in the initial size.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa get_GrowNAryOperators, TextFont::put_IsMathZone, TextRange::BuildUpMath
	virtual HRESULT STDMETHODCALLTYPE put_GrowNAryOperators(VARIANT_BOOL newValue);
	// \brief <em>Retrieves the current setting of the \c HAlignment property</em>
	//
	// Retrieves the horizontal alignment of the control's content. Some of the values defined by the
	// \c HAlignmentConstants enumeration is valid.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \attention Changing this property destroys and recreates the control window.
	//
	// \if UNICODE
	//   \sa put_HAlignment, get_Text, TextParagraph::get_HAlignment, get_DefaultMathZoneHAlignment,
	//       RichTextBoxCtlLibU::HAlignmentConstants
	// \else
	//   \sa put_HAlignment, get_Text, TextParagraph::get_HAlignment, get_DefaultMathZoneHAlignment,
	//       RichTextBoxCtlLibA::HAlignmentConstants
	// \endif
	//virtual HRESULT STDMETHODCALLTYPE get_HAlignment(HAlignmentConstants* pValue);
	// \brief <em>Sets the \c HAlignment property</em>
	//
	// Sets the horizontal alignment of the control's content. Some of the values defined by the
	// \c HAlignmentConstants enumeration is valid.
	//
	// \param[in] newValue The setting to apply.
	//
	// \return An \c HRESULT error code.
	//
	// \attention Changing this property destroys and recreates the control window.
	//
	// \if UNICODE
	//   \sa get_HAlignment, put_Text, TextParagraph::put_HAlignment, put_DefaultMathZoneHAlignment,
	//       RichTextBoxCtlLibU::HAlignmentConstants
	// \else
	//   \sa get_HAlignment, put_Text, TextParagraph::put_HAlignment, put_DefaultMathZoneHAlignment,
	//       RichTextBoxCtlLibA::HAlignmentConstants
	// \endif
	//virtual HRESULT STDMETHODCALLTYPE put_HAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c hDragImageList property</em>
	///
	/// Retrieves the handle to the imagelist containing the drag image that is used during a
	/// drag'n'drop operation to visualize the dragged data.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, OLEDrag, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_hDragImageList(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE get_HoverTime(LONG* pValue);
	/// \brief <em>Sets the \c HoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE put_HoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c HyphenationFunction property</em>
	///
	/// Retrieves the address of the \c HyphenateProc implementation that will be called to hyphenate words.
	/// The address has to point to a function that has the same signature as the \c HyphenateProc prototype
	/// (see <a href="https://msdn.microsoft.com/en-us/library/bb774370.aspx">MSDN</a>).\n
	/// If set to 0, hyphenation is deactivated.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Hyphenation is known to work only with the native rich edit control of Windows and not
	///          with the MS Office version.
	///
	/// \sa put_HyphenationFunction, TextParagraph::get_Hyphenation, get_HyphenationWordWrapZoneWidth,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774370.aspx">HyphenateProc</a>
	virtual HRESULT STDMETHODCALLTYPE get_HyphenationFunction(LONG* pValue);
	/// \brief <em>Sets the \c HyphenationFunction property</em>
	///
	/// Sets the address of the \c HyphenateProc implementation that will be called to hyphenate words.
	/// The address has to point to a function that has the same signature as the \c HyphenateProc prototype
	/// (see <a href="https://msdn.microsoft.com/en-us/library/bb774370.aspx">MSDN</a>).\n
	/// If set to 0, hyphenation is deactivated.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Hyphenation is known to work only with the native rich edit control of Windows and not
	///          with the MS Office version.
	///
	/// \sa get_HyphenationFunction, TextParagraph::put_Hyphenation, put_HyphenationWordWrapZoneWidth,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774370.aspx">HyphenateProc</a>
	virtual HRESULT STDMETHODCALLTYPE put_HyphenationFunction(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c HyphenationWordWrapZoneWidth property</em>
	///
	/// Retrieves the width (in twips) of the zone, in which wrapping words is prefered over hyphenation.
	/// The zone is measured from the right margin to the left side. If there is a word delimiter (e.g. a
	/// space char) within this zone, the word that follows the delimiter will be wrapped to the next line.
	/// If there is no word delimiter within the zone, the word that exceeds the line length will be
	/// hyphenated.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Hyphenation is known to work only with the native rich edit control of Windows and not
	///          with the MS Office version.
	///
	/// \sa put_HyphenationWordWrapZoneWidth, TextParagraph::get_Hyphenation, get_HyphenationFunction
	virtual HRESULT STDMETHODCALLTYPE get_HyphenationWordWrapZoneWidth(SHORT* pValue);
	/// \brief <em>Sets the \c HyphenationWordWrapZoneWidth property</em>
	///
	/// Sets the width (in twips) of the zone, in which wrapping words is prefered over hyphenation.
	/// The zone is measured from the right margin to the left side. If there is a word delimiter (e.g. a
	/// space char) within this zone, the word that follows the delimiter will be wrapped to the next line.
	/// If there is no word delimiter within the zone, the word that exceeds the line length will be
	/// hyphenated.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Hyphenation is known to work only with the native rich edit control of Windows and not
	///          with the MS Office version.
	///
	/// \sa get_HyphenationWordWrapZoneWidth, TextParagraph::put_Hyphenation, put_HyphenationFunction
	virtual HRESULT STDMETHODCALLTYPE put_HyphenationWordWrapZoneWidth(SHORT newValue);
	/// \brief <em>Retrieves the current setting of the \c IMEConversionMode property</em>
	///
	/// Retrieves the control's IME conversion mode, controlling which words are proposed by IME. IME (Input
	/// Method Editor) is a Windows feature making it easy to enter Asian characters. Any of the values
	/// defined by the \c IMEConversionModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is useful only for some input languages, for instance Japanese.
	///
	/// \if UNICODE
	///   \sa put_IMEConversionMode, get_DisableIMEOperations, get_TSFModeBias, get_IMEMode,
	///       RichTextBoxCtlLibU::IMEConversionModeConstants
	/// \else
	///   \sa put_IMEConversionMode, get_DisableIMEOperations, get_TSFModeBias, get_IMEMode,
	///       RichTextBoxCtlLibA::IMEConversionModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IMEConversionMode(IMEConversionModeConstants* pValue);
	/// \brief <em>Sets the \c IMEConversionMode property</em>
	///
	/// Sets the control's IME conversion mode, controlling which words are proposed by IME. IME (Input
	/// Method Editor) is a Windows feature making it easy to enter Asian characters. Any of the values
	/// defined by the \c IMEConversionModeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is useful only for some input languages, for instance Japanese.
	///
	/// \if UNICODE
	///   \sa get_IMEConversionMode, put_DisableIMEOperations, put_TSFModeBias, put_IMEMode,
	///       RichTextBoxCtlLibU::IMEConversionModeConstants
	/// \else
	///   \sa get_IMEConversionMode, put_DisableIMEOperations, put_TSFModeBias, put_IMEMode,
	///       RichTextBoxCtlLibA::IMEConversionModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IMEConversionMode(IMEConversionModeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c IMEMode property</em>
	///
	/// Retrieves the control's IME mode. IME (Input Method Editor) is a Windows feature making it easy to
	/// enter Asian characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IMEMode, get_IMEConversionMode, get_CurrentIMECompositionMode, get_DisableIMEOperations,
	///       get_UseTextServicesFramework, RichTextBoxCtlLibU::IMEModeConstants
	/// \else
	///   \sa put_IMEMode, get_IMEConversionMode, get_CurrentIMECompositionMode, get_DisableIMEOperations,
	///       get_UseTextServicesFramework, RichTextBoxCtlLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IMEMode(IMEModeConstants* pValue);
	/// \brief <em>Sets the \c IMEMode property</em>
	///
	/// Sets the control's IME mode. IME (Input Method Editor) is a Windows feature making it easy to
	/// enter Asian characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IMEMode, put_IMEConversionMode, put_CurrentIMECompositionMode, put_DisableIMEOperations,
	///       put_UseTextServicesFramework, RichTextBoxCtlLibU::IMEModeConstants
	/// \else
	///   \sa get_IMEMode, put_IMEConversionMode, put_CurrentIMECompositionMode, put_DisableIMEOperations,
	///       put_UseTextServicesFramework, RichTextBoxCtlLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IMEMode(IMEModeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c IntegralLimitsLocation property</em>
	///
	/// Retrieves where the limits of an integral function are displayed in math zones. Any of the
	/// values defined by the \c LimitsLocationConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_IntegralLimitsLocation, get_NAryLimitsLocation, TextFont::get_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibU::LimitsLocationConstants
	/// \else
	///   \sa put_IntegralLimitsLocation, get_NAryLimitsLocation, TextFont::get_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibA::LimitsLocationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IntegralLimitsLocation(LimitsLocationConstants* pValue);
	/// \brief <em>Sets the \c IntegralLimitsLocation property</em>
	///
	/// Sets where the limits of an integral function are displayed in math zones. Any of the
	/// values defined by the \c LimitsLocationConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_IntegralLimitsLocation, put_NAryLimitsLocation, TextFont::put_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibU::LimitsLocationConstants
	/// \else
	///   \sa get_IntegralLimitsLocation, put_NAryLimitsLocation, TextFont::put_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibA::LimitsLocationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IntegralLimitsLocation(LimitsLocationConstants newValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c LeftMargin property</em>
	///
	/// Retrieves the width (in pixels) of the control's left margin. If set to -1, a value, that depends on
	/// the control's font, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_LeftMargin, get_RightMargin, get_Font, get_ShowSelectionBar, TextParagraph::get_LeftIndent
	virtual HRESULT STDMETHODCALLTYPE get_LeftMargin(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c LeftMargin property</em>
	///
	/// Sets the width (in pixels) of the control's left margin. If set to -1, a value, that depends on
	/// the control's font, is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LeftMargin, put_RightMargin, put_Font, put_ShowSelectionBar, TextParagraph::put_LeftIndent
	virtual HRESULT STDMETHODCALLTYPE put_LeftMargin(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c LetClientHandleAllIMEOperations property</em>
	///
	/// Retrieves whether the control lets the client application handle all IME operations. IME (Input
	/// Method Editor) is a Windows feature making it easy to enter Asian characters. On east-asian systems
	/// the control can be directed to allow the client application to handle all IME operations, by setting
	/// this property to \c VARIANT_TRUE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_LetClientHandleAllIMEOperations, get_DisableIMEOperations, get_IMEMode, get_IMEConversionMode
	virtual HRESULT STDMETHODCALLTYPE get_LetClientHandleAllIMEOperations(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c LetClientHandleAllIMEOperations property</em>
	///
	/// Sets whether the control lets the client application handle all IME operations. IME (Input
	/// Method Editor) is a Windows feature making it easy to enter Asian characters. On east-asian systems
	/// the control can be directed to allow the client application to handle all IME operations, by setting
	/// this property to \c VARIANT_TRUE.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LetClientHandleAllIMEOperations, put_DisableIMEOperations, put_IMEMode, put_IMEConversionMode
	virtual HRESULT STDMETHODCALLTYPE put_LetClientHandleAllIMEOperations(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c LinkMouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor that is used when the mouse cursor is located over a
	/// link, if the \c LinkMousePointer property is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_LinkMouseIcon, putref_LinkMouseIcon, get_LinkMousePointer, get_MouseIcon,
	///       RichTextBoxCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_LinkMouseIcon, putref_LinkMouseIcon, get_LinkMousePointer, get_MouseIcon,
	///       RichTextBoxCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_LinkMouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c LinkMouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor that is used when the mouse cursor is located over a
	/// link, if the \c LinkMousePointer property is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_LinkMouseIcon, putref_LinkMouseIcon, put_LinkMousePointer, put_MouseIcon,
	///       RichTextBoxCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_LinkMouseIcon, putref_LinkMouseIcon, put_LinkMousePointer, put_MouseIcon,
	///       RichTextBoxCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_LinkMouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c LinkMouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor that is used when the mouse cursor is located over a
	/// link, if the \c LinkMousePointer property is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_LinkMouseIcon, put_LinkMouseIcon, put_LinkMousePointer, putref_MouseIcon,
	///       RichTextBoxCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_LinkMouseIcon, put_LinkMouseIcon, put_LinkMousePointer, putref_MouseIcon,
	///       RichTextBoxCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_LinkMouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c LinkMousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is located over a link.
	/// Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_LinkMousePointer, get_LinkMouseIcon, get_MousePointer,
	///       RichTextBoxCtlLibU::MousePointerConstants
	/// \else
	///   \sa put_LinkMousePointer, get_LinkMouseIcon, get_MousePointer,
	///       RichTextBoxCtlLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_LinkMousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c LinkMousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is located over a link.
	/// Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_LinkMousePointer, put_LinkMouseIcon, put_MousePointer,
	///       RichTextBoxCtlLibU::MousePointerConstants
	/// \else
	///   \sa get_LinkMousePointer, put_LinkMouseIcon, put_MousePointer,
	///       RichTextBoxCtlLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_LinkMousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c LogicalCaret property</em>
	///
	/// Retrieves whether the caret is a real, visible caret or just a logical one that behaves like a
	/// normal caret, but never is displayed. If set to \c VARIANT_TRUE, the caret is a logical caret;
	/// otherwise a real one.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa put_LogicalCaret, get_ReadOnly
	virtual HRESULT STDMETHODCALLTYPE get_LogicalCaret(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c LogicalCaret property</em>
	///
	/// Sets whether the caret is a real, visible caret or just a logical one that behaves like a
	/// normal caret, but never is displayed. If set to \c VARIANT_TRUE, the caret is a logical caret;
	/// otherwise a real one.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \sa get_LogicalCaret, put_ReadOnly
	virtual HRESULT STDMETHODCALLTYPE put_LogicalCaret(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MathLineBreakBehavior property</em>
	///
	/// Retrieves where automatic line breaks may be inserted within math zones. Any of the values
	/// defined by the \c MathLineBreakBehaviorConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_MathLineBreakBehavior, TextFont::get_IsMathZone, TextRange::BuildUpMath,
	///       RichTextBoxCtlLibU::MathLineBreakBehaviorConstants
	/// \else
	///   \sa put_MathLineBreakBehavior, TextFont::get_IsMathZone, TextRange::BuildUpMath,
	///       RichTextBoxCtlLibA::MathLineBreakBehaviorConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MathLineBreakBehavior(MathLineBreakBehaviorConstants* pValue);
	/// \brief <em>Sets the \c MathLineBreakBehavior property</em>
	///
	/// Sets where automatic line breaks may be inserted within math zones. Any of the values
	/// defined by the \c MathLineBreakBehaviorConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_MathLineBreakBehavior, TextFont::put_IsMathZone, TextRange::BuildUpMath,
	///       RichTextBoxCtlLibU::MathLineBreakBehaviorConstants
	/// \else
	///   \sa get_MathLineBreakBehavior, TextFont::put_IsMathZone, TextRange::BuildUpMath,
	///       RichTextBoxCtlLibA::MathLineBreakBehaviorConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MathLineBreakBehavior(MathLineBreakBehaviorConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c MaxTextLength property</em>
	///
	/// Retrieves the maximum number of characters, that the user can type into the control. If set to -1,
	/// the system's default setting is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Text, that is set through the \c Text property may exceed this limit.
	///
	/// \sa put_MaxTextLength, TextRange::get_RangeLength, get_Text, get_BeepOnMaxText, Raise_TruncatedText
	virtual HRESULT STDMETHODCALLTYPE get_MaxTextLength(LONG* pValue);
	/// \brief <em>Sets the \c MaxTextLength property</em>
	///
	/// Sets the maximum number of characters, that the user can type into the control. If set to -1,
	/// the system's default setting is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Text, that is set through the \c Text property may exceed this limit.
	///
	/// \sa get_MaxTextLength, TextRange::get_RangeLength, put_Text, put_BeepOnMaxText, Raise_TruncatedText
	virtual HRESULT STDMETHODCALLTYPE put_MaxTextLength(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MaxUndoQueueSize property</em>
	///
	/// Retrieves the maximum number of actions that can be stored in the control's undo queue.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The actual maximum depends on the available memory and can be lower.
	///
	/// \sa put_MaxUndoQueueSize, BeginNewUndoAction, EmptyUndoBuffer, CanUndo, Undo
	virtual HRESULT STDMETHODCALLTYPE get_MaxUndoQueueSize(LONG* pValue);
	/// \brief <em>Sets the \c MaxUndoQueueSize property</em>
	///
	/// Sets the maximum number of actions that can be stored in the control's undo queue.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The actual maximum depends on the available memory and can be lower.
	///
	/// \sa get_MaxUndoQueueSize, BeginNewUndoAction, EmptyUndoBuffer, CanUndo, Undo
	virtual HRESULT STDMETHODCALLTYPE put_MaxUndoQueueSize(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Modified property</em>
	///
	/// Retrieves a flag indicating whether the control's content has changed. A value of \c VARIANT_TRUE
	/// stands for changed content, a value of \c VARIANT_FALSE for unchanged content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Modified, get_Text, Raise_ContentChanged
	virtual HRESULT STDMETHODCALLTYPE get_Modified(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Modified property</em>
	///
	/// Sets a flag indicating whether the control's content has changed. A value of \c VARIANT_TRUE
	/// stands for changed content, a value of \c VARIANT_FALSE for unchanged content.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Modified, put_Text, Raise_ContentChanged
	virtual HRESULT STDMETHODCALLTYPE put_Modified(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, get_LinkMouseIcon,
	///       RichTextBoxCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, get_LinkMouseIcon,
	///       RichTextBoxCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, put_LinkMouseIcon,
	///       RichTextBoxCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, put_LinkMouseIcon,
	///       RichTextBoxCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, putref_LinkMouseIcon,
	///       RichTextBoxCtlLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, putref_LinkMouseIcon,
	///       RichTextBoxCtlLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c MousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MousePointer, get_MouseIcon, get_LinkMousePointer,
	///       RichTextBoxCtlLibU::MousePointerConstants
	/// \else
	///   \sa put_MousePointer, get_MouseIcon, get_LinkMousePointer,
	///       RichTextBoxCtlLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c MousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MousePointer, put_MouseIcon, put_LinkMousePointer,
	///       RichTextBoxCtlLibU::MousePointerConstants
	/// \else
	///   \sa get_MousePointer, put_MouseIcon, put_LinkMousePointer,
	///       RichTextBoxCtlLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiLine property</em>
	///
	/// Retrieves whether the control processes carriage returns and displays the content on multiple
	/// lines. If set to \c VARIANT_TRUE, the content is displayed on multiple lines; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_MultiLine, get_Text, get_ScrollBars, GetLineCount, get_FirstVisibleLine,
	///     TextParagraph::get_HAlignment
	virtual HRESULT STDMETHODCALLTYPE get_MultiLine(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MultiLine property</em>
	///
	/// Sets whether the control processes carriage returns and displays the content on multiple
	/// lines. If set to \c VARIANT_TRUE, the content is displayed on multiple lines; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_MultiLine, put_Text, put_ScrollBars, TextParagraph::put_HAlignment
	virtual HRESULT STDMETHODCALLTYPE put_MultiLine(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiSelect property</em>
	///
	/// Retrieves whether the control allows simultaneous selection of multiple text ranges by pressing the
	/// [CTRL] key while selecting text. If set to \c VARIANT_TRUE, multiple text ranges can be selected;
	/// otherwise only one text range can be selected.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 5.0 or newer.
	///
	/// \sa put_MultiSelect, TextRange::get_SubRanges, get_SelectedTextRange
	virtual HRESULT STDMETHODCALLTYPE get_MultiSelect(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MultiSelect property</em>
	///
	/// Sets whether the control allows simultaneous selection of multiple text ranges by pressing the
	/// [CTRL] key while selecting text. If set to \c VARIANT_TRUE, multiple text ranges can be selected;
	/// otherwise only one text range can be selected.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 5.0 or newer.
	///
	/// \sa get_MultiSelect, TextRange::get_SubRanges, get_SelectedTextRange
	virtual HRESULT STDMETHODCALLTYPE put_MultiSelect(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c NAryLimitsLocation property</em>
	///
	/// Retrieves where the limits of a n-ary function are displayed in math zones. Any of the values
	/// defined by the \c LimitsLocationConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_NAryLimitsLocation, get_IntegralLimitsLocation, TextFont::get_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibU::LimitsLocationConstants
	/// \else
	///   \sa put_NAryLimitsLocation, get_IntegralLimitsLocation, TextFont::get_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibA::LimitsLocationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_NAryLimitsLocation(LimitsLocationConstants* pValue);
	/// \brief <em>Sets the \c NAryLimitsLocation property</em>
	///
	/// Sets where the limits of a n-ary function are displayed in math zones. Any of the values
	/// defined by the \c LimitsLocationConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_NAryLimitsLocation, put_IntegralLimitsLocation, TextFont::put_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibU::LimitsLocationConstants
	/// \else
	///   \sa get_NAryLimitsLocation, put_IntegralLimitsLocation, TextFont::put_IsMathZone,
	///       TextRange::BuildUpMath, RichTextBoxCtlLibA::LimitsLocationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_NAryLimitsLocation(LimitsLocationConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c NextRedoActionType property</em>
	///
	/// Retrieves the kind of the action that will be redone when calling the \c Redo method. Any of the
	/// values defined by the \c UndoActionTypeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa CanRedo, Redo, get_NextUndoActionType, RichTextBoxCtlLibU::UndoActionTypeConstants
	/// \else
	///   \sa CanRedo, Redo, get_NextUndoActionType, RichTextBoxCtlLibA::UndoActionTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_NextRedoActionType(UndoActionTypeConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c NextUndoActionType property</em>
	///
	/// Retrieves the kind of the action that will be undone when calling the \c Undo method. Any of the
	/// values defined by the \c UndoActionTypeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa CanUndo, Undo, get_NextRedoActionType, RichTextBoxCtlLibU::UndoActionTypeConstants
	/// \else
	///   \sa CanUndo, Undo, get_NextRedoActionType, RichTextBoxCtlLibA::UndoActionTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_NextUndoActionType(UndoActionTypeConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c NoInputSequenceCheck property</em>
	///
	/// Retrieves whether the sequence of typed text is verified in languages that require it, for instance
	/// Thai and Vietnamese. If set to \c VARIANT_TRUE, the sequence of typed text is verified before
	/// accepting the input; otherwise the sequence is not verified and Unicode characters, that are meant
	/// to be combined to a complex character, are displayed separately.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_NoInputSequenceCheck, get_UseTextServicesFramework, get_IMEMode
	virtual HRESULT STDMETHODCALLTYPE get_NoInputSequenceCheck(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c NoInputSequenceCheck property</em>
	///
	/// Sets whether the sequence of typed text is verified in languages that require it, for instance
	/// Thai and Vietnamese. If set to \c VARIANT_TRUE, the sequence of typed text is verified before
	/// accepting the input; otherwise the sequence is not verified and Unicode characters, that are meant
	/// to be combined to a complex character, are displayed separately.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_NoInputSequenceCheck, put_UseTextServicesFramework, put_IMEMode
	virtual HRESULT STDMETHODCALLTYPE put_NoInputSequenceCheck(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c OLEDragImageStyle property</em>
	///
	/// Retrieves the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       RichTextBoxCtlLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       RichTextBoxCtlLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue);
	/// \brief <em>Sets the \c OLEDragImageStyle property</em>
	///
	/// Sets the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       RichTextBoxCtlLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       RichTextBoxCtlLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OLEDragImageStyle(OLEDragImageStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE get_ProcessContextMenuKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessContextMenuKeys property</em>
	///
	/// Sets whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE put_ProcessContextMenuKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c RawSubScriptAndSuperScriptOperators property</em>
	///
	/// Retrieves whether subscript (_) and superscript (^) operators are displayed as themselves in math
	/// zones that are in linear format. If set to \c VARIANT_TRUE, those operators are displayed as
	/// themselves; otherwise they are replaced by an arrow pointing down and an arrow pointing up.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa put_RawSubScriptAndSuperScriptOperators, TextFont::get_IsMathZone, TextRange::BuildDownMath
	virtual HRESULT STDMETHODCALLTYPE get_RawSubScriptAndSuperScriptOperators(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RawSubScriptAndSuperScriptOperators property</em>
	///
	/// Sets whether subscript (_) and superscript (^) operators are displayed as themselves in math
	/// zones that are in linear format. If set to \c VARIANT_TRUE, those operators are displayed as
	/// themselves; otherwise they are replaced by an arrow pointing down and an arrow pointing up.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa get_RawSubScriptAndSuperScriptOperators, TextFont::put_IsMathZone, TextRange::BuildDownMath
	virtual HRESULT STDMETHODCALLTYPE put_RawSubScriptAndSuperScriptOperators(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ReadOnly property</em>
	///
	/// Retrieves whether the control accepts user input, that would change the control's content. If set to
	/// \c VARIANT_FALSE, such user input is accepted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ReadOnly, get_Enabled, get_LogicalCaret, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_ReadOnly(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ReadOnly property</em>
	///
	/// Sets whether the control accepts user input, that would change the control's content. If set to
	/// \c VARIANT_FALSE, such user input is accepted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ReadOnly, put_Enabled, put_LogicalCaret, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_ReadOnly(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. Any of the values
	/// defined by the \c RegisterForOLEDragDropConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RegisterForOLEDragDrop, get_DetectDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter,
	///       RichTextBoxCtlLibU::RegisterForOLEDragDropConstants
	/// \else
	///   \sa put_RegisterForOLEDragDrop, get_DetectDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter,
	///       RichTextBoxCtlLibA::RegisterForOLEDragDropConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants* pValue);
	/// \brief <em>Sets the \c RegisterForOLEDragDrop property</em>
	///
	/// Sets whether the control is registered as a target for OLE drag'n'drop. Any of the values
	/// defined by the \c RegisterForOLEDragDropConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_RegisterForOLEDragDrop, put_DetectDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter,
	///       RichTextBoxCtlLibU::RegisterForOLEDragDropConstants
	/// \else
	///   \sa get_RegisterForOLEDragDrop, put_DetectDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter,
	///       RichTextBoxCtlLibA::RegisterForOLEDragDropConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants newValue);
	/// \brief <em>Retrieves the supported version of the Rich Edit API set</em>
	///
	/// The control can use the rich edit implementation of Windows and different versions of Office. The
	/// supported API set depends on the Windows version, if the native rich edit is used, and on the
	/// Office version, if the MS Office rich edit control is used. This property retrieves the version of
	/// the supported API set. Any of the values defined by the \c RichEditAPIVersionConstants enumeration
	/// is valid.
	///
	/// \param[out] pValue The control's API version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_RichEditVersion, RichTextBoxCtlLibU::RichEditAPIVersionConstants
	/// \else
	///   \sa get_RichEditVersion, RichTextBoxCtlLibA::RichEditAPIVersionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RichEditAPIVersion(RichEditAPIVersionConstants* pValue);
	/// \brief <em>Retrieves the fully qualified path to the currently used rich edit library</em>
	///
	/// The control can use the rich edit implementation of Windows and different versions of Office. This
	/// property retrieves the fully qualified path to the library that implements the currently used
	/// native rich edit control.
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_RichEditVersion, get_RichEditAPIVersion
	virtual HRESULT STDMETHODCALLTYPE get_RichEditLibraryPath(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c RichEditVersion property</em>
	///
	/// Retrieves whether to use the rich edit control of Microsoft Office or the rich edit control of
	/// Microsoft Windows. Any of the values defined by the \c RichEditVersionConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The rich edit control of Microsoft Office is implemented in the file richedit20.dll, which
	///          is installed to <code>%COMMONPROGRAMFILES(x86)%\Microsoft Shared\OFFICE<em>n</em></code>.
	///          The rich edit control of Microsoft Windows is implemented in the file msftedit.dll in the
	///          system directory. Neither of both files can be redistributed.\n
	///          The rich edit control of Microsoft Office usually has more features than the rich edit
	///          control of Microsoft Windows.\n
	///          Currently the following rich edit versions can be used:
	///          - Rich Edit 4.1, native rich edit control of Windows XP, 2003, 2003 R2, Vista, 2008, 7,
	///            2008 R2.
	///          - Rich Edit 5.0, rich edit control of Office 2003.
	///          - Rich Edit 6.0, rich edit control of Office 2007.
	///          - Rich Edit 7.0, rich edit control of Office 2010.
	///          - Rich Edit 7.5, native rich edit control of Windows 8, 2012, 8.1, 2012 R2, 10.
	///          - Rich Edit 8.0, rich edit control of Office 2013, 2016.
	///
	/// \attention This property cannot be changed at runtime.
	///
	/// \if UNICODE
	///   \sa put_RichEditVersion, get_RichEditAPIVersion, get_RichEditLibraryPath,
	///       RichTextBoxCtlLibU::RichEditVersionConstants
	/// \else
	///   \sa put_RichEditVersion, get_RichEditAPIVersion, get_RichEditLibraryPath,
	///       RichTextBoxCtlLibA::RichEditVersionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RichEditVersion(RichEditVersionConstants* pValue);
	/// \brief <em>Sets the \c RichEditVersion property</em>
	///
	/// Sets whether to use the rich edit control of Microsoft Office or the rich edit control of
	/// Microsoft Windows. Any of the values defined by the \c RichEditVersionConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The rich edit control of Microsoft Office is implemented in the file richedit20.dll, which
	///          is installed to <code>%COMMONPROGRAMFILES(x86)%\Microsoft Shared\OFFICE<em>n</em></code>.
	///          The rich edit control of Microsoft Windows is implemented in the file msftedit.dll in the
	///          system directory. Neither of both files can be redistributed.\n
	///          The rich edit control of Microsoft Office usually has more features than the rich edit
	///          control of Microsoft Windows.\n
	///          Currently the following rich edit versions can be used:
	///          - Rich Edit 4.1, native rich edit control of Windows XP, 2003, 2003 R2, Vista, 2008, 7,
	///            2008 R2.
	///          - Rich Edit 5.0, rich edit control of Office 2003.
	///          - Rich Edit 6.0, rich edit control of Office 2007.
	///          - Rich Edit 7.0, rich edit control of Office 2010.
	///          - Rich Edit 7.5, native rich edit control of Windows 8, 2012, 8.1, 2012 R2, 10.
	///          - Rich Edit 8.0, rich edit control of Office 2013, 2016.
	///
	/// \attention This property cannot be changed at runtime.
	///
	/// \if UNICODE
	///   \sa get_RichEditVersion, get_RichEditAPIVersion, get_RichEditLibraryPath,
	///       RichTextBoxCtlLibU::RichEditVersionConstants
	/// \else
	///   \sa get_RichEditVersion, get_RichEditAPIVersion, get_RichEditLibraryPath,
	///       RichTextBoxCtlLibA::RichEditVersionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RichEditVersion(RichEditVersionConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c RichText property</em>
	///
	/// Retrieves the control's content, including RTF control words for formatting.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RichText, get_Text, TextFont, TextParagraph, TextRange::get_RichText
	virtual HRESULT STDMETHODCALLTYPE get_RichText(BSTR* pValue);
	/// \brief <em>Sets the \c RichText property</em>
	///
	/// Sets the control's content, including RTF control words for formatting.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RichText, put_Text, TextFont, TextParagraph, TextRange::put_RichText
	virtual HRESULT STDMETHODCALLTYPE put_RichText(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c RightMargin property</em>
	///
	/// Retrieves the width (in pixels) of the control's right margin. If set to -1, a value, that depends on
	/// the control's font, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RightMargin, get_LeftMargin, get_Font, TextParagraph::get_RightIndent
	virtual HRESULT STDMETHODCALLTYPE get_RightMargin(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c RightMargin property</em>
	///
	/// Sets the width (in pixels) of the control's right margin. If set to -1, a value, that depends on
	/// the control's font, is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RightMargin, put_LeftMargin, put_Font, TextParagraph::put_RightIndent
	virtual HRESULT STDMETHODCALLTYPE put_RightMargin(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether bidirectional features are enabled or disabled. Any combination of the values
	/// defined by the \c RightToLeftConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention On Windows XP, changing this property destroys and recreates the control window.\n
	///            Setting or clearing the \c rtlText flag destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa put_RightToLeft, get_IMEMode, TextParagraph::get_RightToLeft, get_RightToLeftMathZones,
	///       Raise_WritingDirectionChanged, RichTextBoxCtlLibU::RightToLeftConstants
	/// \else
	///   \sa put_RightToLeft, get_IMEMode, TextParagraph::get_RightToLeft, get_RightToLeftMathZones,
	///       Raise_WritingDirectionChanged, RichTextBoxCtlLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(RightToLeftConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Enables or disables bidirectional features. Any combination of the values defined by the
	/// \c RightToLeftConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention On Windows XP, changing this property destroys and recreates the control window.\n
	///            Setting or clearing the \c rtlText flag destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, put_IMEMode, TextParagraph::put_RightToLeft, put_RightToLeftMathZones,
	///       Raise_WritingDirectionChanged, RichTextBoxCtlLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, put_IMEMode, TextParagraph::put_RightToLeft, put_RightToLeftMathZones,
	///       Raise_WritingDirectionChanged, RichTextBoxCtlLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeftMathZones property</em>
	///
	/// Retrieves whether bidirectional features of math zones within right-to-left paragraphs are enabled or
	/// disabled. If set to \c VARIANT_TRUE, math-zones in right-to-left paragraphs will have right-to-left
	/// layout as well; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa put_RightToLeftMathZones, TextParagraph::get_RightToLeft, get_RightToLeft,
	///     TextFont::get_IsMathZone, TextRange::BuildUpMath
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeftMathZones(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RightToLeftMathZones property</em>
	///
	/// Enables or disables bidirectional features of math zones within right-to-left paragraphs. If set to
	/// \c VARIANT_TRUE, math-zones in right-to-left paragraphs will have right-to-left layout as well;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa get_RightToLeftMathZones, TextParagraph::put_RightToLeft, put_RightToLeft,
	///     TextFont::put_IsMathZone, TextRange::BuildUpMath
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeftMathZones(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ScrollBars property</em>
	///
	/// Retrieves the scrollbars to show. Any combination of the values defined by the \c ScrollBarsConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ScrollBars, get_AlwaysShowScrollBars, get_AutoScrolling, get_MultiLine, Scroll,
	///       Raise_Scrolling, RichTextBoxCtlLibU::ScrollBarsConstants
	/// \else
	///   \sa put_ScrollBars, get_AlwaysShowScrollBars, get_AutoScrolling, get_MultiLine, Scroll,
	///       Raise_Scrolling, RichTextBoxCtlLibA::ScrollBarsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ScrollBars(ScrollBarsConstants* pValue);
	/// \brief <em>Sets the \c ScrollBars property</em>
	///
	/// Sets the scrollbars to show. Any combination of the values defined by the \c ScrollBarsConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ScrollBars, put_AlwaysShowScrollBars, put_AutoScrolling, put_MultiLine, Scroll,
	///       Raise_Scrolling, RichTextBoxCtlLibU::ScrollBarsConstants
	/// \else
	///   \sa get_ScrollBars, put_AlwaysShowScrollBars, put_AutoScrolling, put_MultiLine, Scroll,
	///       Raise_Scrolling, RichTextBoxCtlLibA::ScrollBarsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ScrollBars(ScrollBarsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ScrollToTopOnKillFocus property</em>
	///
	/// Retrieves whether the control scrolls automatically to character position 0 when it loses the
	/// keyboard focus. If set to \c VARIANT_TRUE, the control scrolls to top when losing focus; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ScrollToTopOnKillFocus, get_ScrollBars, TextRange::ScrollIntoView, Scroll
	virtual HRESULT STDMETHODCALLTYPE get_ScrollToTopOnKillFocus(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ScrollToTopOnKillFocus property</em>
	///
	/// Sets whether the control scrolls automatically to character position 0 when it loses the
	/// keyboard focus. If set to \c VARIANT_TRUE, the control scrolls to top when losing focus; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ScrollToTopOnKillFocus, put_ScrollBars, TextRange::ScrollIntoView, Scroll
	virtual HRESULT STDMETHODCALLTYPE put_ScrollToTopOnKillFocus(VARIANT_BOOL newValue);
	// \brief <em>Retrieves the current setting of the \c SelectedText property</em>
	//
	// Retrieves the currently selected text.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks If \c MultiSelect is enabled, this property refers only to the active sub-range.\n
	//          This property is read-only.
	//
	// \sa get_SelectedTextRange, get_MultiSelect, ReplaceSelectedText, get_AutoSelectWordOnTrackSelection,
	//     get_Text
	//virtual HRESULT STDMETHODCALLTYPE get_SelectedText(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c SelectedText property</em>
	///
	/// Retrieves a \c TextRange object for the currently selected text.
	///
	/// \param[out] ppTextRange The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TextRange, get_MultiSelect, get_TextRange
	virtual HRESULT STDMETHODCALLTYPE get_SelectedTextRange(IRichTextRange** ppTextRange);
	/// \brief <em>Retrieves the current setting of the \c SelectionType property</em>
	///
	/// Retrieves what kind of data the current selection contains. Any combination of the values defined
	/// by the \c SelectionTypeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_SelectedTextRange, get_MultiSelect, RichTextBoxCtlLibU::SelectionTypeConstants
	/// \else
	///   \sa get_SelectedTextRange, get_MultiSelect, RichTextBoxCtlLibA::SelectionTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SelectionType(SelectionTypeConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c ShowDragImage property</em>
	///
	/// Retrieves whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible;
	/// otherwise hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowDragImage, get_hDragImageList, get_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_ShowDragImage(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowDragImage property</em>
	///
	/// Sets whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible; otherwise
	/// hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, get_hDragImageList, put_SupportOLEDragImages, Raise_DragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_ShowDragImage(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowSelectionBar property</em>
	///
	/// Retrieves whether the control adds space to the left margin where the cursor changes to a right-up
	/// arrow, allowing the user to select full lines of text. If set to \c VARIANT_TRUE, the selection bar
	/// is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowSelectionBar, get_LeftMargin
	virtual HRESULT STDMETHODCALLTYPE get_ShowSelectionBar(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowSelectionBar property</em>
	///
	/// Sets whether the control adds space to the left margin where the cursor changes to a right-up
	/// arrow, allowing the user to select full lines of text. If set to \c VARIANT_TRUE, the selection bar
	/// is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowSelectionBar, put_LeftMargin
	virtual HRESULT STDMETHODCALLTYPE put_ShowSelectionBar(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SmartSpacingOnDrop property</em>
	///
	/// Retrieves whether a space character is inserted automatically when text is being dropped right before
	/// or right after a word. If set to \c VARIANT_TRUE, the space character is inserted automatically;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_SmartSpacingOnDrop, get_DropWordsOnWordBoundariesOnly, get_RegisterForOLEDragDrop
	virtual HRESULT STDMETHODCALLTYPE get_SmartSpacingOnDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SmartSpacingOnDrop property</em>
	///
	/// Sets whether a space character is inserted automatically when text is being dropped right before
	/// or right after a word. If set to \c VARIANT_TRUE, the space character is inserted automatically;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_SmartSpacingOnDrop, put_DropWordsOnWordBoundariesOnly, put_RegisterForOLEDragDrop
	virtual HRESULT STDMETHODCALLTYPE put_SmartSpacingOnDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SupportOLEDragImages property</em>
	///
	/// Retrieves whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, get_ShowDragImage, get_OLEDragImageStyle,
	///     FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE get_SupportOLEDragImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SupportOLEDragImages property</em>
	///
	/// Sets whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, put_ShowDragImage, put_OLEDragImageStyle,
	///     FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SwitchFontOnIMEInput property</em>
	///
	/// Retrieves whether the control changes to another font when the user explicitly changes to a different
	/// keyboard layout and starts typing text. If set to \c VARIANT_TRUE, the font is switched
	/// automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_SwitchFontOnIMEInput, get_UseDualFontMode, get_IMEMode, TextRange::get_Font
	virtual HRESULT STDMETHODCALLTYPE get_SwitchFontOnIMEInput(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SwitchFontOnIMEInput property</em>
	///
	/// Sets whether the control changes to another font when the user explicitly changes to a different
	/// keyboard layout and starts typing text. If set to \c VARIANT_TRUE, the font is switched
	/// automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_SwitchFontOnIMEInput, put_UseDualFontMode, put_IMEMode, TextRange::get_Font
	virtual HRESULT STDMETHODCALLTYPE put_SwitchFontOnIMEInput(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the control's content. It is the plain text, without RTF formatting.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa put_Text, get_RichText, get_TextRange, TextRange::get_RangeLength, get_MaxTextLength,
	///     get_TextFlow, LoadFromFile, SaveToFile, TextParagraph::get_HAlignment, get_MultiLine, get_Font,
	///     Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the control's content. It is the plain text, without RTF formatting.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa get_Text, put_RichText, get_TextRange, TextRange::get_RangeLength, put_MaxTextLength,
	///     put_TextFlow, LoadFromFile, SaveToFile, TextParagraph::put_HAlignment, put_MultiLine, put_Font,
	///     Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c TextFlow property</em>
	///
	/// Retrieves a value specifying how the text flows. Any of the values defined by the
	/// \c TextFlowConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_TextFlow, get_TextOrientation, get_Text, RichTextBoxCtlLibU::TextFlowConstants
	/// \else
	///   \sa put_TextFlow, get_TextOrientation, get_Text, RichTextBoxCtlLibA::TextFlowConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TextFlow(TextFlowConstants* pValue);
	/// \brief <em>Sets the \c TextFlow property</em>
	///
	/// Sets a value specifying how the text flows. Any of the values defined by the
	/// \c TextFlowConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_TextFlow, put_TextOrientation, put_Text, RichTextBoxCtlLibU::TextFlowConstants
	/// \else
	///   \sa get_TextFlow, put_TextOrientation, put_Text, RichTextBoxCtlLibA::TextFlowConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TextFlow(TextFlowConstants newValue);
	// \brief <em>Retrieves the current setting of the \c TextLength property</em>
	//
	// Retrieves the length of the text specified by the \c Text property.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks This property is read-only.
	//
	// \sa get_MaxTextLength, get_Text
	//virtual HRESULT STDMETHODCALLTYPE get_TextLength(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c TextOrientation property</em>
	///
	/// Retrieves a value that specifies whether the control's text is organized in lines or columns. By
	/// default text is written in horizontal lines, from left to right and top to bottom. Some languages,
	/// e.g. traditional Chinese andtraditional Japanese, are written in vertical columns, from top to
	/// bottom and right to left.\n
	/// Any of the values defined by the \c TextOrientationConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_TextOrientation, get_IMEMode, get_TextFlow, get_Text,
	///       RichTextBoxCtlLibU::TextOrientationConstants
	/// \else
	///   \sa put_TextOrientation, get_IMEMode, get_TextFlow, get_Text,
	///       RichTextBoxCtlLibA::TextOrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TextOrientation(TextOrientationConstants* pValue);
	/// \brief <em>Sets the \c TextOrientation property</em>
	///
	/// Sets a value that specifies whether the control's text is organized in lines or columns. By
	/// default text is written in horizontal lines, from left to right and top to bottom. Some languages,
	/// e.g. traditional Chinese andtraditional Japanese, are written in vertical columns, from top to
	/// bottom and right to left.\n
	/// Any of the values defined by the \c TextOrientationConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_TextOrientation, put_IMEMode, put_TextFlow, put_Text,
	///       RichTextBoxCtlLibU::TextOrientationConstants
	/// \else
	///   \sa get_TextOrientation, put_IMEMode, put_TextFlow, put_Text,
	///       RichTextBoxCtlLibA::TextOrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TextOrientation(TextOrientationConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c TextRange property</em>
	///
	/// Retrieves a \c TextRange object for the specified range of text.
	///
	/// \param[in] rangeStart The zero-based index of the character at which the range starts.
	/// \param[in] rangeEnd The zero-based index of the first character after the end of the range.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.\n
	///          If \c rangeStart is set to 0 and \c rangeEnd is set to -1, the range includes the whole
	///          content of the control.
	///
	/// \sa TextRange, get_Text, get_SelectedTextRange
	virtual HRESULT STDMETHODCALLTYPE get_TextRange(LONG rangeStart = 0, LONG rangeEnd = -1, IRichTextRange** ppTextRange = NULL);
	/// \brief <em>Retrieves the current setting of the \c TSFModeBias property</em>
	///
	/// Retrieves the control's Text Services Framework (TSF) mode bias, controlling which strings are
	/// proposed by IME. IME (Input Method Editor) is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c TSFModeBiasConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_TSFModeBias, get_IMEMode, get_IMEConversionMode, RichTextBoxCtlLibU::TSFModeBiasConstants
	/// \else
	///   \sa put_TSFModeBias, get_IMEMode, get_IMEConversionMode, RichTextBoxCtlLibA::TSFModeBiasConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TSFModeBias(TSFModeBiasConstants* pValue);
	/// \brief <em>Sets the \c TSFModeBias property</em>
	///
	/// Sets the control's Text Services Framework (TSF) mode bias, controlling which strings are
	/// proposed by IME. IME (Input Method Editor) is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c TSFModeBiasConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_TSFModeBias, put_IMEMode, put_IMEConversionMode, RichTextBoxCtlLibU::TSFModeBiasConstants
	/// \else
	///   \sa get_TSFModeBias, put_IMEMode, put_IMEConversionMode, RichTextBoxCtlLibA::TSFModeBiasConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TSFModeBias(TSFModeBiasConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c UseBkAcetateColorForTextSelection property</em>
	///
	/// Retrieves whether the background acetate color is used for drawing the selection background. If set
	/// to \c VARIANT_TRUE, the background acetate color is used; otherwise classic Windows colors are used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.
	///
	/// \sa put_UseBkAcetateColorForTextSelection, get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_UseBkAcetateColorForTextSelection(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseBkAcetateColorForTextSelection property</em>
	///
	/// Sets whether the background acetate color is used for drawing the selection background. If set
	/// to \c VARIANT_TRUE, the background acetate color is used; otherwise classic Windows colors are used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.
	///
	/// \sa get_UseBkAcetateColorForTextSelection, put_BackColor
	virtual HRESULT STDMETHODCALLTYPE put_UseBkAcetateColorForTextSelection(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseBuiltInSpellChecking property</em>
	///
	/// Retrieves whether the built-in spell-checking of Windows is activated. If set to \c VARIANT_TRUE,
	/// built-in spell-checking is activated; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 and Windows 8 or newer.
	///
	/// \sa put_UseBuiltInSpellChecking, get_UseTouchKeyboardAutoCorrection, TextFont::get_Locale
	virtual HRESULT STDMETHODCALLTYPE get_UseBuiltInSpellChecking(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseBuiltInSpellChecking property</em>
	///
	/// Sets whether the built-in spell-checking of Windows is activated. If set to \c VARIANT_TRUE,
	/// built-in spell-checking is activated; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 and Windows 8 or newer.
	///
	/// \sa get_UseBuiltInSpellChecking, put_UseTouchKeyboardAutoCorrection, TextFont::put_Locale
	virtual HRESULT STDMETHODCALLTYPE put_UseBuiltInSpellChecking(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseCustomFormattingRectangle property</em>
	///
	/// Retrieves whether the control uses the formatting rectangle defined by the \c FormattingRectangle*
	/// properties.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for
	/// formatting the text displayed in the window rectangle. When an edit control is first displayed, the
	/// two rectangles are identical on the screen. An application can make the formatting rectangle larger
	/// than the window rectangle (thereby limiting the visibility of the control's text) or smaller than
	/// the window rectangle (thereby creating extra white space around the text).\n
	/// If this property is set to \c VARIANT_FALSE, the formatting rectangle is set to its default values.
	/// Otherwise it's defined by the \c FormattingRectangle* properties.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa put_UseCustomFormattingRectangle, get_FormattingRectangleHeight, get_FormattingRectangleLeft,
	///     get_FormattingRectangleTop, get_FormattingRectangleWidth, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE get_UseCustomFormattingRectangle(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseCustomFormattingRectangle property</em>
	///
	/// Sets whether the control uses the formatting rectangle defined by the \c FormattingRectangle*
	/// properties.\n
	/// The visibility of the control's text is governed by the dimensions of its window rectangle and its
	/// formatting rectangle. The formatting rectangle is a construct maintained by the system for
	/// formatting the text displayed in the window rectangle. When an edit control is first displayed, the
	/// two rectangles are identical on the screen. An application can make the formatting rectangle larger
	/// than the window rectangle (thereby limiting the visibility of the control's text) or smaller than
	/// the window rectangle (thereby creating extra white space around the text).\n
	/// If this property is set to \c VARIANT_FALSE, the formatting rectangle is set to its default values.
	/// Otherwise it's defined by the \c FormattingRectangle* properties.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \sa get_UseCustomFormattingRectangle, put_FormattingRectangleHeight, put_FormattingRectangleLeft,
	///     put_FormattingRectangleTop, put_FormattingRectangleWidth, put_MultiLine
	virtual HRESULT STDMETHODCALLTYPE put_UseCustomFormattingRectangle(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseDualFontMode property</em>
	///
	/// Retrieves whether the control switches to an English font for ASCII text and to an Asian font for
	/// Asian text. If set to \c VARIANT_TRUE, the font is switched automatically depending on the kind of
	/// text. This is the so-called dual font mode. If set to \c VARIANT_FALSE, the same font (usually an
	/// Asian font) is used for all text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseDualFontMode, get_SwitchFontOnIMEInput, get_IMEMode, TextRange::get_Font
	virtual HRESULT STDMETHODCALLTYPE get_UseDualFontMode(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseDualFontMode property</em>
	///
	/// Sets whether the control switches to an English font for ASCII text and to an Asian font for
	/// Asian text. If set to \c VARIANT_TRUE, the font is switched automatically depending on the kind of
	/// text. This is the so-called dual font mode. If set to \c VARIANT_FALSE, the same font (usually an
	/// Asian font) is used for all text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseDualFontMode, put_SwitchFontOnIMEInput, put_IMEMode, TextRange::get_Font
	virtual HRESULT STDMETHODCALLTYPE put_UseDualFontMode(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSmallerFontForNestedFractions property</em>
	///
	/// Retrieves whether nested fractions within math zones are displayed in a smaller font or in normal
	/// font size. If set to \c VARIANT_TRUE, a smaller font is used; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa put_UseSmallerFontForNestedFractions, TextFont::get_Size, TextFont::get_IsMathZone,
	///     TextRange::BuildUpMath
	virtual HRESULT STDMETHODCALLTYPE get_UseSmallerFontForNestedFractions(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSmallerFontForNestedFractions property</em>
	///
	/// Sets whether nested fractions within math zones are displayed in a smaller font or in normal
	/// font size. If set to \c VARIANT_TRUE, a smaller font is used; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \sa get_UseSmallerFontForNestedFractions, TextFont::put_Size, TextFont::put_IsMathZone,
	///     TextRange::BuildUpMath
	virtual HRESULT STDMETHODCALLTYPE put_UseSmallerFontForNestedFractions(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseTextServicesFramework property</em>
	///
	/// Retrieves whether the control uses
	/// <a href="https://msdn.microsoft.com/en-us/library/ms629032.aspx">Windows Text Services Framework</a>
	/// (TSF). If set to \c VARIANT_TRUE, TSF is used; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseTextServicesFramework, get_IMEMode, get_AllowInputThroughTSF,
	///     get_AllowObjectInsertionThroughTSF, get_AllowTSFProofingTips, get_AllowTSFSmartTags,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms629032.aspx">Text Services Framework</a>
	virtual HRESULT STDMETHODCALLTYPE get_UseTextServicesFramework(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseTextServicesFramework property</em>
	///
	/// Sets whether the control uses
	/// <a href="https://msdn.microsoft.com/en-us/library/ms629032.aspx">Windows Text Services Framework</a>
	/// (TSF). If set to \c VARIANT_TRUE, TSF is used; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseTextServicesFramework, put_IMEMode, put_AllowInputThroughTSF,
	///     put_AllowObjectInsertionThroughTSF, put_AllowTSFProofingTips, put_AllowTSFSmartTags,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms629032.aspx">Text Services Framework</a>
	virtual HRESULT STDMETHODCALLTYPE put_UseTextServicesFramework(VARIANT_BOOL newValue);
	// \brief <em>Retrieves the current setting of the \c UseTouchKeyboardAutoCorrection property</em>
	//
	// Retrieves whether the built-in auto-correction of Windows is used when text is entered via a touch
	// screen keyboard. If set to \c VARIANT_TRUE, built-in auto-correction is activated; otherwise not.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks Requires Rich Edit 7.5 and Windows 8 or newer.
	//
	// \sa put_UseTouchKeyboardAutoCorrection, get_UseBuiltInSpellChecking, TextFont::get_Locale
	//virtual HRESULT STDMETHODCALLTYPE get_UseTouchKeyboardAutoCorrection(VARIANT_BOOL* pValue);
	// \brief <em>Sets the \c UseTouchKeyboardAutoCorrection property</em>
	//
	// Sets whether the built-in auto-correction of Windows is used when text is entered via a touch
	// screen keyboard. If set to \c VARIANT_TRUE, built-in auto-correction is activated; otherwise not.
	//
	// \param[in] newValue The setting to apply.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks Requires Rich Edit 7.5 and Windows 8 or newer.
	//
	// \sa get_UseTouchKeyboardAutoCorrection, put_UseBuiltInSpellChecking, TextFont::put_Locale
	//virtual HRESULT STDMETHODCALLTYPE put_UseTouchKeyboardAutoCorrection(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseWindowsThemes property</em>
	///
	/// Retrieves whether the control draws itself using the Windows Theming engine. For instance this
	/// affects the appearance of the control's borders.\n
	/// If set to \c VARIANT_TRUE, the theming engine is used; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.
	///
	/// \sa put_UseWindowsThemes, get_Appearance, get_BorderStyle
	virtual HRESULT STDMETHODCALLTYPE get_UseWindowsThemes(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseWindowsThemes property</em>
	///
	/// Sets whether the control draws itself using the Windows Theming engine. For instance this
	/// affects the appearance of the control's borders.\n
	/// If set to \c VARIANT_TRUE, the theming engine is used; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.
	///
	/// \sa get_UseWindowsThemes, put_Appearance, put_BorderStyle
	virtual HRESULT STDMETHODCALLTYPE put_UseWindowsThemes(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c WordBreakFunction property</em>
	///
	/// Retrieves the function that is responsible to tell the control where a word starts and where it ends.
	/// This property takes the address of a function having the following signature:\n
	/// \code
	///   int CALLBACK FindWorkBreak(LPTSTR pText, int startPosition, int textLength, int flags);
	/// \endcode
	/// The \c pText argument is a pointer to the control's text.\n
	/// The \c startPosition argument specifies the (zero-based) position within the text, at which the
	/// function should begin checking for a word break.\n
	/// The \c textLength argument specifies the length of the text pointed to by \c pText in characters.\n
	/// The \c flags argument specifies the action to be taken by the function. This can be one of the
	/// following values:
	/// - \c WB_CLASSIFY Retrieve the character class and word break flags of the character at the
	///   specified position.
	/// - \c WB_ISDELIMITER Check whether the character at the specified position is a delimiter.
	/// - \c WB_LEFT Find the beginning of a word to the left of the specified position.
	/// - \c WB_LEFTBREAK Find the end-of-word delimiter to the left of the specified position.
	/// - \c WB_MOVEWORDLEFT Find the beginning of a word to the left of the specified position. This value
	///   is used during [CTRL]+[LEFT] key processing.
	/// - \c WB_MOVEWORDRIGHT Find the beginning of a word to the right of the specified position. This
	///   value is used during [CTRL]+[RIGHT] key processing.
	/// - \c WB_RIGHT Find the beginning of a word to the right of the specified position. This is useful
	///   in right-aligned rich edit controls.
	/// - \c WB_RIGHTBREAK Find the end-of-word delimiter to the right of the specified position. This is
	///   useful in right-aligned rich edit controls.
	///
	/// If the \c flags parameter specifies \c WB_CLASSIFY, the return value is the character class and
	/// word break flags of the character at the specified position.\n
	/// If the \c flags parameter specifies \c WB_ISDELIMITER and the character at the specified position
	/// is a delimiter, the function must return \c TRUE.\n
	/// If the \c flags parameter specifies \c WB_ISDELIMITER and the character at the specified position
	/// is not a delimiter, the function must return \c FALSE.\n
	/// If the \c flags parameter specifies \c WB_LEFT or \c WB_RIGHT, the function must return the
	/// (zero-based) index to the beginning of a word in the specified text.\n\n
	/// If this property is set to 0, the system's internal function is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_WordBreakFunction, get_Text, TextParagraph::get_HAlignment,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672125.aspx">EditWordBreakProc</a>
	virtual HRESULT STDMETHODCALLTYPE get_WordBreakFunction(LONG* pValue);
	/// \brief <em>Sets the \c WordBreakFunction property</em>
	///
	/// Sets the function that is responsible to tell the control where a word starts and where it ends.
	/// This property takes the address of a function having the following signature:\n
	/// \code
	///   int CALLBACK FindWorkBreak(LPTSTR pText, int startPosition, int textLength, int flags);
	/// \endcode
	/// The \c pText argument is a pointer to the control's text.\n
	/// The \c startPosition argument specifies the (zero-based) position within the text, at which the
	/// function should begin checking for a word break.\n
	/// The \c textLength argument specifies the length of the text pointed to by \c pText in characters.\n
	/// The \c flags argument specifies the action to be taken by the function. This can be one of the
	/// following values:
	/// - \c WB_CLASSIFY Retrieve the character class and word break flags of the character at the
	///   specified position.
	/// - \c WB_ISDELIMITER Check whether the character at the specified position is a delimiter.
	/// - \c WB_LEFT Find the beginning of a word to the left of the specified position.
	/// - \c WB_LEFTBREAK Find the end-of-word delimiter to the left of the specified position.
	/// - \c WB_MOVEWORDLEFT Find the beginning of a word to the left of the specified position. This value
	///   is used during [CTRL]+[LEFT] key processing.
	/// - \c WB_MOVEWORDRIGHT Find the beginning of a word to the right of the specified position. This
	///   value is used during [CTRL]+[RIGHT] key processing.
	/// - \c WB_RIGHT Find the beginning of a word to the right of the specified position. This is useful
	///   in right-aligned rich edit controls.
	/// - \c WB_RIGHTBREAK Find the end-of-word delimiter to the right of the specified position. This is
	///   useful in right-aligned rich edit controls.
	///
	/// If the \c flags parameter specifies \c WB_CLASSIFY, the return value is the character class and
	/// word break flags of the character at the specified position.\n
	/// If the \c flags parameter specifies \c WB_ISDELIMITER and the character at the specified position
	/// is a delimiter, the function must return \c TRUE.\n
	/// If the \c flags parameter specifies \c WB_ISDELIMITER and the character at the specified position
	/// is not a delimiter, the function must return \c FALSE.\n
	/// If the \c flags parameter specifies \c WB_LEFT or \c WB_RIGHT, the function must return the
	/// (zero-based) index to the beginning of a word in the specified text.\n\n
	/// If this property is set to 0, the system's internal function is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_WordBreakFunction, put_Text, TextParagraph::put_HAlignment,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672125.aspx">EditWordBreakProc</a>
	virtual HRESULT STDMETHODCALLTYPE put_WordBreakFunction(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ZoomRatio property</em>
	///
	/// Retrieves the zoom ratio. Any value between 1/64 and 64 is valid. If set to 0, no zooming is applied.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ZoomRatio, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_ZoomRatio(DOUBLE* pValue);
	/// \brief <em>Sets the \c ZoomRatio property</em>
	///
	/// Sets the zoom ratio. Any value between 1/64 and 64 is valid. If set to 0, no zooming is applied.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ZoomRatio, put_Font
	virtual HRESULT STDMETHODCALLTYPE put_ZoomRatio(DOUBLE newValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Closes the current undo action and starts a new one</em>
	///
	/// Stops the control from collecting additional typing actions into the current undo action. The
	/// control stores the next typing action, if any, into a new action in the undo queue.\n
	/// The control groups consecutive typing actions, including characters deleted by using the [BACKSPACE]
	/// key, into a single undo action until one of the following events occurs:
	/// - \c BeginNewUndoAction is called.
	/// - The control loses focus.
	/// - The user moves the current selection, either by using the arrow keys or by clicking the mouse.
	/// - The user presses the [DEL] key.
	/// - The user performs any other action, such as a paste operation that does not involve typing.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CanUndo, Undo, EmptyUndoBuffer, get_MaxUndoQueueSize
	virtual HRESULT STDMETHODCALLTYPE BeginNewUndoAction(void);
	/// \brief <em>Determines whether data can be copied to the clipboard</em>
	///
	/// \param[out] pValue \c VARIANT_TRUE if data can be copied (i.e. there is a selection); otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Copy, Paste
	virtual HRESULT STDMETHODCALLTYPE CanCopy(VARIANT_BOOL* pValue);
	/// \brief <em>Determines whether there are any actions in the control's redo queue</em>
	///
	/// \param[out] pValue \c VARIANT_TRUE if there are actions in the redo queue; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Redo, CanUndo, get_NextRedoActionType, EmptyUndoBuffer
	virtual HRESULT STDMETHODCALLTYPE CanRedo(VARIANT_BOOL* pValue);
	/// \brief <em>Determines whether there are any actions in the control's undo queue</em>
	///
	/// \param[out] pValue \c VARIANT_TRUE if there are actions in the undo queue; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Undo, CanRedo, get_NextRedoActionType, BeginNewUndoAction, EmptyUndoBuffer
	virtual HRESULT STDMETHODCALLTYPE CanUndo(VARIANT_BOOL* pValue);
	// \brief <em>Retrieves the specified character's position in client coordinates</em>
	//
	// \param[in] characterIndex The zero-based index of the character within the control, for which to
	//            retrieve the position. If the character is a line delimiter, the returned coordinates
	//            indicate a point just beyond the last visible character in the line. If the specified
	//            index is greater than the index of the last character in the control, the function fails.
	// \param[out] pX The x-coordinate (in pixels) of the character relative to the control's upper-left
	//             corner.
	// \param[out] pY The y-coordinate (in pixels) of the character relative to the control's upper-left
	//             corner.
	//
	// \return An \c HRESULT error code.
	//
	// \sa PositionToCharIndex
	//virtual HRESULT STDMETHODCALLTYPE CharIndexToPosition(LONG characterIndex, OLE_XPOS_PIXELS* pX = NULL, OLE_YPOS_PIXELS* pY = NULL);
	/// \brief <em>Raises the ShouldResizeControlWindow event</em>
	///
	/// Checks whether the control's content is smaller or larger than the control's window size and raises
	/// the \c ShouldResizeControlWindow event.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_ShouldResizeControlWindow
	virtual HRESULT STDMETHODCALLTYPE CheckForOptimalControlWindowSize(void);
	/// \brief <em>Closes the Text Services Framework (TSF) keyboard</em>
	///
	/// Closes the Text Services Framework (TSF) keyboard.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OpenTSFKeyboard
	virtual HRESULT STDMETHODCALLTYPE CloseTSFKeyboard(void);
	/// \brief <em>Starts a new document</em>
	///
	/// Starts a new document, resetting any input.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadFromFile, SaveToFile, get_Text, put_Text
	virtual HRESULT STDMETHODCALLTYPE CreateNewDocument(void);
	/// \brief <em>Clears the control's undo and redo queues</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CanUndo, Undo, BeginNewUndoAction, get_MaxUndoQueueSize
	virtual HRESULT STDMETHODCALLTYPE EmptyUndoBuffer(void);
	/// \brief <em>Finishes a pending drop operation</em>
	///
	/// During a drag'n'drop operation the drag image is displayed until the \c OLEDragDrop event has been
	/// handled. This order is intended by Microsoft Windows. However, if a message box is displayed from
	/// within the \c OLEDragDrop event, or the drop operation cannot be performed asynchronously and takes
	/// a long time, it may be desirable to remove the drag image earlier.\n
	/// This method will break the intended order and finish the drag'n'drop operation (including removal
	/// of the drag image) immediately.
	///
	/// \remarks This method will fail if not called from the \c OLEDragDrop event handler or if no drag
	///          images are used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_OLEDragDrop, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE FinishOLEDragDrop(void);
	/// \brief <em>Deactivates screen updating</em>
	///
	/// Disables automatic redrawing of the control. Disabling redraw while doing large changes on the
	/// control may increase performance.\n
	/// On each call the internal freeze counter is <strong>incremented</strong>. As long as this counter
	/// is nonzero, the control is not redrawn.
	///
	/// \param[out] pNewCount Will be set to the updated freeze count.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Unfreeze
	virtual HRESULT STDMETHODCALLTYPE Freeze(LONG* pNewCount);
	/// \brief <em>Retrieves the current IME composition text</em>
	///
	/// Retrieves the text currently being composed by the Input Method Editor (IME).
	///
	/// \param[in] textType The type of text to retrieve. Any of the values defined by the
	///            \c IMECompositionTextType enumeration is valid.
	/// \param[out] pCompositionText Will be set to the current composition string.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IMEMode, get_CurrentIMECompositionMode, RichTextBoxCtlLibU::IMECompositionTextType
	/// \else
	///   \sa get_IMEMode, get_CurrentIMECompositionMode, RichTextBoxCtlLibA::IMECompositionTextType
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetCurrentIMECompositionText(IMECompositionTextType textType = imecttFinalComposedString, BSTR* pCompositionText = NULL);
	// \brief <em>Retrieves the zero-based index of the first character of the specified line</em>
	//
	// \param[in] lineIndex The zero-based index of the line to retrieve the first character for. If set
	//            to -1, the index of the line containing the caret is used.
	// \param[out] pValue The zero-based index of the first character of the line. -1 if the specified line
	//             index is greater than the total number of lines.
	//
	// \return An \c HRESULT error code.
	//
	// \sa GetLineFromChar, get_MultiLine
	//virtual HRESULT STDMETHODCALLTYPE GetFirstCharOfLine(LONG lineIndex, LONG* pValue);
	// \brief <em>Retrieves the text of the specified line</em>
	//
	// \param[in] lineIndex The zero-based index of the line to retrieve the text for.
	// \param[out] pValue The line's text.
	//
	// \return An \c HRESULT error code.
	//
	// \sa GetLineLength, get_MultiLine, get_Text
	//virtual HRESULT STDMETHODCALLTYPE GetLine(LONG lineIndex, BSTR* pValue);
	/// \brief <em>Retrieves the number of lines in the control</em>
	///
	/// \param[out] pValue The number of lines.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstVisibleLine, get_MultiLine
	virtual HRESULT STDMETHODCALLTYPE GetLineCount(LONG* pValue);
	// \brief <em>Retrieves the zero-based index of the line that contains the specified character</em>
	//
	// \param[in] characterIndex The zero-based index of the character within the control. If set to
	//            -1, the index of the character at which the selection begins, is used. If there's no
	//            selection, the index of the character next to the caret is used.
	// \param[out] pValue The zero-based index of the line containing the character.
	//
	// \return An \c HRESULT error code.
	//
	// \sa get_FirstVisibleChar, GetFirstCharOfLine, get_MultiLine, get_SelectionStart
	//virtual HRESULT STDMETHODCALLTYPE GetLineFromChar(LONG characterIndex, LONG* pValue);
	// \brief <em>Retrieves the number of characters in the specified line</em>
	//
	// \param[in] lineIndex The zero-based index of the line to retrieve the length for. If set to -1, the
	//            number of unselected characters on lines containing selected characters is retrieved.
	//            E. g. if the selection extended from the fourth character of one line through the eighth
	//            character from the end of the next line, the return value would be 10 (three characters on
	//            the first line and seven on the next).
	// \param[out] pValue The number of characters in the line.
	//
	// \return An \c HRESULT error code.
	//
	// \sa GetLine, MultiLine
	//virtual HRESULT STDMETHODCALLTYPE GetLineLength(LONG lineIndex, LONG* pValue);
	// \brief <em>Retrieves the current selection's start and end</em>
	//
	// Retrieves the zero-based character indices of the current selection's start and end.
	//
	// \param[out] pSelectionStart The zero-based index of the character at which the selection starts.
	// \param[out] pSelectionEnd The zero-based index of the first unselected character after the end of the
	//             selection.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks If \c MultiSelect is activated, the text range returned by this method refers to the
	//          active sub-range.
	//
	// \sa SetSelection, ReplaceSelectedText, get_SelectedText, get_SelectedTextRange, get_MultiSelect
	//virtual HRESULT STDMETHODCALLTYPE GetSelection(LONG* pSelectionStart = NULL, LONG* pSelectionEnd = NULL);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in,out] pHitTestDetails Receives a value specifying the exact part of the control the
	///                specified point lies in. Any of the values defined by the \c HitTestConstants
	///                enumeration is valid.
	/// \param[out] ppHitRange Receives the "hit" text range's \c IRichTextRange implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa IRichTextRange, RichTextBoxCtlLibU::HitTestConstants, HitTest
	/// \else
	///   \sa IRichTextRange, RichTextBoxCtlLibA::HitTestConstants, HitTest
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IRichTextRange** ppHitRange);
	/// \brief <em>Loads rich text from a file</em>
	///
	/// Loads rich text from a file. It can be a plain text file, or a rtf file.
	///
	/// \param[in] file The file to load.
	/// \param[in] fileCreationDisposition Specifies how to open the file. Any of the values defined by the
	///            \c FileCreationDispositionConstants enumeration is valid.
	/// \param[in] fileType The type of the file to load. Any of the values defined by the
	///            \c FileTypeConstants enumeration is valid.
	/// \param[in] fileAccessOptions Further options to apply. Any combination of the values defined by the
	///            \c FileAccessOptionConstants enumeration is valid.
	/// \param[in] codePage The code page to use for the file. If set to 0, the ANSI code page (\c CP_ACP)
	///            will be used unless the file begins with a Unicode BOM (0xFEFF), in which case the file
	///            is considered to be Unicode.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE if the file could be loaded and to
	///             \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SaveToFile, CreateNewDocument, get_Text, RichTextBoxCtlLibU::FileCreationDispositionConstants,
	///       RichTextBoxCtlLibU::FileTypeConstants, RichTextBoxCtlLibU::FileAccessOptionConstants
	/// \else
	///   \sa SaveToFile, CreateNewDocument, get_Text, RichTextBoxCtlLibA::FileCreationDispositionConstants,
	///       RichTextBoxCtlLibA::FileTypeConstants, RichTextBoxCtlLibA::FileAccessOptionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE LoadFromFile(BSTR file, FileCreationDispositionConstants fileCreationDisposition = fcdOpenExisting, FileTypeConstants fileType = ftyAutoDetect, FileAccessOptionConstants fileAccessOptions = faoDefault, LONG codePage = 0, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Enters OLE drag'n'drop mode</em>
	///
	/// \param[in] pDataObject A pointer to the \c IDataObject implementation to use during OLE
	///            drag'n'drop. If not specified, the control's own implementation is used.
	/// \param[in] supportedEffects A bit field defining all drop effects the client wants to support.
	///            Any combination of the values defined by the \c OLEDropEffectConstants enumeration
	///            (except \c odeScroll) is valid.
	/// \param[in] hWndToAskForDragImage The handle of the window, that is awaiting the
	///            \c DI_GETDRAGIMAGE message to specify the drag image to use. If -1, the method
	///            creates the drag image itself. If \c SupportOLEDragImages is set to \c VARIANT_FALSE,
	///            no drag image is used.
	/// \param[in] pDraggedTextRange The range of text to drag. This parameter is used to generate the
	///            drag image, if \c hWndToAskForDragImage is set to -1.
	/// \param[in] itemCountToDisplay The number to display in the item count label of Aero drag images.
	///            If set to 0 or 1, no item count label is displayed. If set to any value larger than 1,
	///            this value is displayed in the item count label.
	/// \param[out] pPerformedEffects The performed drop effect. Any of the values defined by the
	///             \c OLEDropEffectConstants enumeration (except \c odeScroll) is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TextRange, Raise_BeginDrag, Raise_BeginRDrag, Raise_OLEStartDrag, get_SupportOLEDragImages,
	///       get_OLEDragImageStyle, RichTextBoxCtlLibU::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \else
	///   \sa TextRange, Raise_BeginDrag, Raise_BeginRDrag, Raise_OLEStartDrag, get_SupportOLEDragImages,
	///       get_OLEDragImageStyle, RichTextBoxCtlLibA::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE OLEDrag(LONG* pDataObject = NULL, OLEDropEffectConstants supportedEffects = odeCopyOrMove, OLE_HANDLE hWndToAskForDragImage = -1, IRichTextRange* pDraggedTextRange = NULL, LONG itemCountToDisplay = 0, OLEDropEffectConstants* pPerformedEffects = NULL);
	/// \brief <em>Opens the Text Services Framework (TSF) keyboard</em>
	///
	/// Opens the Text Services Framework (TSF) keyboard.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CloseTSFKeyboard
	virtual HRESULT STDMETHODCALLTYPE OpenTSFKeyboard(void);
	// \brief <em>Retrieves the character closest to the specified position</em>
	//
	// Retrieves the zero-based index of the character nearest the specified position.
	//
	// \param[in] x The x-coordinate (in pixels) of the position to retrieve the nearest character for. It
	//            is relative to the control's upper-left corner.
	// \param[in] y The y-coordinate (in pixels) of the position to retrieve the nearest character for. It
	//            is relative to the control's upper-left corner.
	// \param[out] pCharacterIndex The zero-based index of the character within the control, that is
	//             nearest to the specified position. If the specified point is beyond the last character
	//             in the control, this value indicates the last character in the control. The index
	//             indicates the line delimiter if the specified point is beyond the last visible
	//             character in a line.
	// \param[out] pLineIndex The zero-based index of the line, that contains the character specified by
	//             the \c pCharacterIndex parameter.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks If a point outside the bounds of the control is passed, the function fails.
	//
	// \sa CharIndexToPosition
	//virtual HRESULT STDMETHODCALLTYPE PositionToCharIndex(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, LONG* pCharacterIndex = NULL, LONG* pLineIndex = NULL);
	/// \brief <em>Redoes the next action(s) in the control's redo queue</em>
	///
	/// \param[in] numberOfActions The count of actions to redo.
	/// \param[out] pCount Will be set to the count of actions that actually have been redone.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CanRedo, Undo, get_NextRedoActionType, EmptyUndoBuffer
	virtual HRESULT STDMETHODCALLTYPE Redo(LONG numberOfActions = 1, LONG* pCount = NULL);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	// \brief <em>Replaces the currently selected text</em>
	//
	// Replaces the control's currently selected text.
	//
	// \param[in] replacementText The text that replaces the currently selected text.
	// \param[in] undoable If \c VARIANT_TRUE, this action is inserted into the control's undo queue;
	//            otherwise not.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks If \c MultiSelect is activated, only the text of the active sub-range is replaced.
	//
	// \sa get_SelectedTextRange, TextRange::put_Text
	//virtual HRESULT STDMETHODCALLTYPE ReplaceSelectedText(BSTR replacementText, VARIANT_BOOL undoable = VARIANT_FALSE);
	/// \brief <em>Saves rich text to a file</em>
	///
	/// Saves rich text to a file. It can be a plain text file, or a rtf file.
	///
	/// \param[in] file The file to save to. If not specified, the document's current filename is used. If
	///            this is not specified as well, the method fails.
	/// \param[in] fileCreationDisposition Specifies how to save the file. Any of the values defined by the
	///            \c FileCreationDispositionConstants enumeration is valid.
	/// \param[in] fileType The format to save the file in. Any of the values defined by the
	///            \c FileTypeConstants enumeration is valid.
	/// \param[in] fileAccessOptions Further options to apply. Any combination of the values defined by the
	///            \c FileAccessOptionConstants enumeration is valid.
	/// \param[in] codePage The code page to use for the file. Common values are 0 (\c CP_ACP, ANSI code
	///            page) and 65001 (UTF-8 code page).
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE if the file could be saved and to
	///             \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa LoadFromFile, CreateNewDocument, get_Text,
	///       RichTextBoxCtlLibU::FileCreationDispositionConstants, RichTextBoxCtlLibU::FileTypeConstants,
	///       RichTextBoxCtlLibU::FileAccessOptionConstants
	/// \else
	///   \sa LoadFromFile, CreateNewDocument, get_Text,
	///       RichTextBoxCtlLibA::FileCreationDispositionConstants, RichTextBoxCtlLibA::FileTypeConstants,
	///       RichTextBoxCtlLibA::FileAccessOptionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SaveToFile(BSTR file = NULL, FileCreationDispositionConstants fileCreationDisposition = static_cast<FileCreationDispositionConstants>(0), FileTypeConstants fileType = ftyAutoDetect, FileAccessOptionConstants fileAccessOptions = faoDefault, LONG codePage = 0, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Scrolls the control</em>
	///
	/// \param[in] axis The axis which is to be scrolled. Any combination of the values defined by the
	///            \c ScrollAxisConstants enumeration is valid.
	/// \param[in] directionAndIntensity The intensity and direction of the action. Any of the values defined
	///            by the \c ScrollDirectionConstants enumeration is valid.
	/// \param[in] linesToScrollVertically The number of lines to scroll vertically. This parameter is
	///            ignored, if \c directionAndIntensity is not set to \c sdCustom.
	/// \param[in] charactersToScrollHorizontally The number of characters to scroll horizontally. This
	///            parameter is ignored, if \c directionAndIntensity is not set to \c sdCustom.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method has no effect if the \c MultiLine property is set to \c VARIANT_FALSE.
	///
	/// \if UNICODE
	///   \sa TextRange::ScrollIntoView, get_ScrollToTopOnKillFocus, get_ScrollBars, get_AutoScrolling,
	///       get_MultiLine, Raise_Scrolling, RichTextBoxCtlLibU::ScrollAxisConstants,
	///       RichTextBoxCtlLibU::ScrollDirectionConstants
	/// \else
	///   \sa TextRange::ScrollIntoView, get_ScrollToTopOnKillFocus, get_ScrollBars, get_AutoScrolling,
	///       get_MultiLine, Raise_Scrolling, RichTextBoxCtlLibA::ScrollAxisConstants,
	///       RichTextBoxCtlLibA::ScrollDirectionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Scroll(ScrollAxisConstants axis, ScrollDirectionConstants directionAndIntensity, LONG linesToScrollVertically = 0, LONG charactersToScrollHorizontally = 0);
	// \brief <em>Scrolls the control so that the caret is visible</em>
	//
	// Ensures that the control's caret is visible by scrolling the control if necessary.
	//
	// \return An \c HRESULT error code.
	//
	// \sa Scroll, get_ScrollToTopOnKillFocus, get_ScrollBars, get_AutoScrolling, get_MultiLine,
	//     Raise_Scrolling
	//virtual HRESULT STDMETHODCALLTYPE ScrollCaretIntoView(void);
	// \brief <em>Sets the selection's start and end</em>
	//
	// Sets the zero-based character indices of the selection's start and end.
	//
	// \param[in] selectionStart The zero-based index of the character at which the selection starts. If set
	//            to -1, the current selection is cleared.
	// \param[in] selectionEnd The zero-based index of the first unselected character after the end of the
	//            selection.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks To select all text in the control, set \c selectionStart to 0 and \c selectionEnd to -1.\n
	//          If \c MultiSelect is enabled, this method replaces only the active sub-range.
	//
	// \sa GetSelection, ReplaceSelectedText, get_SelectedText, get_SelectedTextRange, get_MultiSelect
	//virtual HRESULT STDMETHODCALLTYPE SetSelection(LONG selectionStart, LONG selectionEnd);
	/// \brief <em>Sets the target device</em>
	///
	/// Sets the device context and the line width (in twips) that the control will optimize output for.
	/// The device context usually is a printer. The line width can be set to the width of a paper format
	/// like DIN A4. It has to be in twips though and the client application should ensure that the target
	/// device supports this format.
	///
	/// \param[in] hDC The handle of the device context to optimize output for.
	/// \param[in] lineWidth The line width (in twips) to use.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Setting both parameters to 0 will make words wrap on the control boundaries.\n
	///          Setting \c hDC to NULL and \c lineWidth to 1 will disable word-wrapping.
	///
	/// \sa put_AutoScrolling, put_ScrollBars
	virtual HRESULT STDMETHODCALLTYPE SetTargetDevice(OLE_HANDLE hDC, LONG lineWidth);
	/// \brief <em>Displays the IME reconversion dialog</em>
	///
	/// Invokes the Input Method Editor (IME) reconversion dialog box.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_IMEMode
	virtual HRESULT STDMETHODCALLTYPE StartIMEReconversion(void);
	/// \brief <em>Undoes the last action(s) in the control's undo queue</em>
	///
	/// \param[in] numberOfActions The count of actions to undo. This parameter can also be set to some
	///            special values:
	///            - 0 - Stops collecting undo actions and empties the undo buffer.
	///            - -1 - Restarts undo again.
	///            - -9999995 - Suspends collecting undo actions.
	///            - -9999994 - Resumes collecting undo actions.
	/// \param[out] pCount Will be set to the count of actions that actually have been undone.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CanUndo, Redo, get_NextUndoActionType, BeginNewUndoAction, EmptyUndoBuffer, get_MaxUndoQueueSize
	virtual HRESULT STDMETHODCALLTYPE Undo(LONG numberOfActions = 1, LONG* pCount = NULL);
	/// \brief <em>Activates screen updating</em>
	///
	/// Enables automatic redrawing of the control. Disabling redraw while doing large changes on the
	/// control may increase performance.\n
	/// On each call the internal freeze counter is <strong>decremented</strong>. As long as this counter
	/// is nonzero, the control is not redrawn.
	///
	/// \param[out] pNewCount Will be set to the updated freeze count.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Freeze
	virtual HRESULT STDMETHODCALLTYPE Unfreeze(LONG* pNewCount);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Property object changes
	///
	//@{
	/// \brief <em>Will be called after a property object was changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnChanged
	HRESULT OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/);
	/// \brief <em>Will be called before a property object is changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that is about to be changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnRequestEdit
	HRESULT OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Called to create the control window</em>
	///
	/// Called to create the control window. This method overrides \c CWindowImpl::Create() and is
	/// used to customize the window styles and window class name.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] szWindowName The control's window name.
	/// \param[in] dwStyle The control's window style. Will be ignored.
	/// \param[in] dwExStyle The control's extended window style. Will be ignored.
	/// \param[in] MenuOrID The control's ID.
	/// \param[in] lpCreateParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnCreate, GetStyleBits, GetExStyleBits
	HWND Create(HWND hWndParent, ATL::_U_RECT rect = NULL, LPCTSTR szWindowName = NULL, DWORD dwStyle = 0, DWORD dwExStyle = 0, ATL::_U_MENUorID MenuOrID = 0U, LPVOID lpCreateParam = NULL);
	/// \brief <em>Called to draw the control</em>
	///
	/// Called to draw the control. This method overrides \c CComControlBase::OnDraw() and is used to prevent
	/// the "ATL 9.0" drawing in user mode and replace it in design mode.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/056hw3hs.aspx">CComControlBase::OnDraw</a>
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Called after receiving the last message (typically \c WM_NCDESTROY)</em>
	///
	/// \param[in] hWnd The window being destroyed.
	///
	/// \sa OnCreate, OnDestroy
	void OnFinalMessage(HWND /*hWnd*/);
	/// \brief <em>Informs an embedded object of its display location within its container</em>
	///
	/// \param[in] pClientSite The \c IOleClientSite implementation of the container application's
	///            client side.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms684013.aspx">IOleObject::SetClientSite</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite);
	/// \brief <em>Notifies the control when the container's document window is activated or deactivated</em>
	///
	/// ATL's implementation of \c OnDocWindowActivate calls \c IOleInPlaceObject_UIDeactivate if the control
	/// is deactivated. This causes a bug in MDI apps. If the control sits on a \c MDI child window and has
	/// the focus and the title bar of another top-level window (not the MDI parent window) of the same
	/// process is clicked, the focus is moved from the ATL based ActiveX control to the next control on the
	/// MDI child before it is moved to the other top-level window that was clicked. If the focus is set back
	/// to the MDI child, the ATL based control no longer has the focus.
	///
	/// \param[in] fActivate If \c TRUE, the document window is activated; otherwise deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/0kz79wfc.aspx">IOleInPlaceActiveObjectImpl::OnDocWindowActivate</a>
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL /*fActivate*/);

	/// \brief <em>A keyboard or mouse message needs to be translated</em>
	///
	/// The control's container calls this method if it receives a keyboard or mouse message. It gives
	/// us the chance to customize keystroke translation (i. e. to react to them in a non-default way).
	/// This method overrides \c CComControlBase::PreTranslateAccelerator.
	///
	/// \param[in] pMessage A \c MSG structure containing details about the received window message.
	/// \param[out] hReturnValue A reference parameter of type \c HRESULT which will be set to \c S_OK,
	///             if the message was translated, and to \c S_FALSE otherwise.
	///
	/// \return \c FALSE if the object's container should translate the message; otherwise \c TRUE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hxa56938.aspx">CComControlBase::PreTranslateAccelerator</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646373.aspx">TranslateAccelerator</a>
	BOOL PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue);

	//////////////////////////////////////////////////////////////////////
	/// \name Drag image creation
	///
	//@{
	/// \brief <em>Creates a legacy drag image for the specified text</em>
	///
	/// Creates a drag image for the specified text in the style of Windows versions prior to Vista. The
	/// drag image is added to an imagelist which is returned.
	///
	/// \param[in] pTextRange The range of text for which to create the drag image.
	/// \param[out] pUpperLeftPoint Receives the coordinates (in pixels) of the drag image's upper-left
	///             corner relative to the control's upper-left corner.
	/// \param[out] pBoundingRectangle Receives the drag image's bounding rectangle (in pixels) relative to
	///             the control's upper-left corner.
	///
	/// \return An imagelist containing the drag image.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa CreateLegacyOLEDragImage
	HIMAGELIST CreateLegacyDragImage(__in ITextRange* pTextRange, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle);
	/// \brief <em>Creates a legacy OLE drag image for the specified text</em>
	///
	/// Creates an OLE drag image for the specified text in the style of Windows versions prior to Vista.
	///
	/// \param[in] pTextRange The range of text for which to create the drag image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateVistaOLEDragImage, CreateLegacyDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateLegacyOLEDragImage(__in ITextRange* pTextRange, __in LPSHDRAGIMAGE pDragImage);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropTarget
	///
	//@{
	/// \brief <em>Indicates whether a drop can be accepted, and, if so, the effect of the drop</em>
	///
	/// This method is called by the \c DoDragDrop function to determine the target's preferred drop
	/// effect the first time the user moves the mouse into the control during OLE drag'n'Drop. The
	/// target communicates the \c DoDragDrop function the drop effect it wants to be used on drop.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the dragged
	///            data.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragOver, DragLeave, Drop, Raise_OLEDragEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Notifies the target that it no longer is the target of the current OLE drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse out of the
	/// control during OLE drag'n'Drop or if the user canceled the operation. The target must release
	/// any references it holds to the data object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, Drop, Raise_OLEDragLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeave(void);
	/// \brief <em>Communicates the current drop effect to the \c DoDragDrop function</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse over the
	/// control during OLE drag'n'Drop. The target communicates the \c DoDragDrop function the drop
	/// effect it wants to be used on drop.
	///
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, the current drop effect (defined by the \c DROPEFFECT
	///                enumeration). On return, this paramter must be set to the drop effect that the
	///                target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragLeave, Drop, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
	virtual HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Incorporates the source data into the target and completes the drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user completes the drag'n'drop
	/// operation. The target must incorporate the dragged data into itself and pass the used drop
	/// effect back to the \c DoDragDrop function. The target must release any references it holds to
	/// the data object.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the data
	///            to transfer.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target finally executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, DragLeave, Raise_OLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSource
	///
	//@{
	/// \brief <em>Notifies the source of the current drop effect during OLE drag'n'drop</em>
	///
	/// This method is called frequently by the \c DoDragDrop function to notify the source of the
	/// last drop effect that the target has chosen. The source should set an appropriate mouse cursor.
	///
	/// \param[in] effect The drop effect chosen by the target. Any of the values defined by the
	///            \c DROPEFFECT enumeration is valid.
	///
	/// \return \c S_OK if the method has set a custom mouse cursor; \c DRAGDROP_S_USEDEFAULTCURSORS to
	///         use default mouse cursors; or an error code otherwise.
	///
	/// \sa QueryContinueDrag, Raise_OLEGiveFeedback,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693723.aspx">IDropSource::GiveFeedback</a>
	virtual HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD effect);
	/// \brief <em>Determines whether a drag'n'drop operation should be continued, canceled or completed</em>
	///
	/// This method is called by the \c DoDragDrop function to determine whether a drag'n'drop
	/// operation should be continued, canceled or completed.
	///
	/// \param[in] pressedEscape Indicates whether the user has pressed the \c ESC key since the
	///            previous call of this method. If \c TRUE, the key has been pressed; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	///
	/// \return \c S_OK if the drag'n'drop operation should continue; \c DRAGDROP_S_DROP if it should
	///         be completed; \c DRAGDROP_S_CANCEL if it should be canceled; or an error code otherwise.
	///
	/// \sa GiveFeedback, Raise_OLEQueryContinueDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690076.aspx">IDropSource::QueryContinueDrag</a>
	virtual HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL pressedEscape, DWORD keyState);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSourceNotify
	///
	//@{
	/// \brief <em>Notifies the source that the user drags the mouse cursor into a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor into a potential drop target window.
	///
	/// \param[in] hWndTarget The potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragLeaveTarget, Raise_OLEDragEnterPotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragEnterTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnterTarget(HWND hWndTarget);
	/// \brief <em>Notifies the source that the user drags the mouse cursor out of a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor out of a potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnterTarget, Raise_OLEDragLeavePotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragLeaveTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeaveTarget(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichEditOleCallback
	///
	//@{
	/// \brief <em>Provides storage for a new object pasted from the clipboard or read in from an Rich Text Format (RTF) stream</em>
	///
	/// Provides storage for a new object pasted from the clipboard or read in from an Rich Text Format (RTF)
	/// stream. This method must be implemented to allow cut, copy, paste, drag, and drop operations of
	/// Component Object Model (COM) objects.
	///
	/// \param[out] ppStorage Receives the address of the \c IStorage interface created for the new object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774323.aspx">IRichEditOleCallback::GetNewStorage</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380015.aspx">IStorage</a>
	virtual HRESULT STDMETHODCALLTYPE GetNewStorage(LPSTORAGE FAR* ppStorage);
	/// \brief <em>Provides the application and document-level interfaces and information required to support in-place activation</em>
	///
	/// Provides the application and document-level interfaces and information required to support in-place
	/// activation.
	///
	/// \param[out] ppInPlaceFrame Receives the address of the \c IOleInPlaceFrame interface that represents
	///             the frame window of a rich edit control client.
	/// \param[out] ppInPlaceUIWindow The address of the \c IOleInPlaceUIWindow interface that represents the
	///             document window of the rich edit control client. An interface need not be returned if the
	///             frame and document windows are the same.
	/// \param[in] pInPlaceFrameInfo A \c OLEINPLACEFRAMEINFO with information about the accelerators
	///            supported by a container during an in-place session.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774321.aspx">IRichEditOleCallback::GetInPlaceContext</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692770.aspx">IOleInPlaceFrame</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680716.aspx">IOleInPlaceUIWindow</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693737.aspx">OLEINPLACEFRAMEINFO</a>
	virtual HRESULT STDMETHODCALLTYPE GetInPlaceContext(LPOLEINPLACEFRAME FAR* ppInPlaceFrame, LPOLEINPLACEUIWINDOW FAR* ppInPlaceUIWindow, LPOLEINPLACEFRAMEINFO pInPlaceFrameInfo);
	/// \brief <em>Indicates whether or not the application is to display its container UI</em>
	///
	/// Indicates whether or not the application is to display its container UI. The rich edit control looks
	/// ahead for double-clicks and defers the call if appropriate.
	///
	/// \param[in] show Specifies whether to show or hide the container UI. If set to \c TRUE, the container
	///            UI is to be displayed; otherwise it is to be hidden.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774329.aspx">IRichEditOleCallback::ShowContainerUI</a>
	virtual HRESULT STDMETHODCALLTYPE ShowContainerUI(BOOL /*show*/);
	/// \brief <em>Queries the application as to whether an object should be inserted</em>
	///
	/// Queries the application as to whether an object should be inserted. The member is called when
	/// pasting and when reading Rich Text Format (RTF).
	///
	/// \param[in] pCLSID Class identifier of the object to be inserted.
	/// \param[in] pStorage Storage containing the object.
	/// \param[in] insertionPosition Character position, at which the object will be inserted.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_InsertingOLEObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774327.aspx">IRichEditOleCallback::QueryInsertObject</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380015.aspx">IStorage</a>
	virtual HRESULT STDMETHODCALLTYPE QueryInsertObject(LPCLSID pCLSID, LPSTORAGE pStorage, LONG insertionPosition);
	/// \brief <em>Sends notification that an object is about to be deleted from a rich edit control</em>
	///
	/// Sends notification that an object is about to be deleted from a rich edit control. The object is not
	/// necessarily being released when this member is called.
	///
	/// \param[in] pOLEObject The object that is being deleted.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_RemovingOLEObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774313.aspx">IRichEditOleCallback::DeleteObject</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd542709.aspx">IOleObject</a>
	virtual HRESULT STDMETHODCALLTYPE DeleteObject(LPOLEOBJECT pOLEObject);
	/// \brief <em>During a paste operation or a drag event, determines if the data that is pasted or dragged should be accepted</em>
	///
	/// During a paste operation or a drag event, determines if the data that is pasted or dragged should be
	/// accepted.
	///
	/// \param[in] pDataObject The data object being pasted or dragged.
	/// \param[in,out] pFormat The clipboard format that will be used for the paste or drop operation. If the
	///                value pointed to by \c pFormat is zero, the best available format will be used. If the
	///                callback changes the value pointed to by \c pFormat, the rich edit control only uses
	///                that format and the operation will fail if the format is not available.
	/// \param[in] operationFlag If \c RECO_DROP, this is a drop operation; if it is \c RECO_PASTE, it is a
	///            paste from the clipboard.
	/// \param[in] operationFlag If \c RECO_DROP, this is a drop operation; if it is \c RECO_PASTE, it is a
	///            paste from the clipboard.
	/// \param[in] operationIsReal If \c TRUE, the specified operation is really performed; otherwise it is
	///            just a query ("what if" scenario).
	/// \param[in] hMetaPict Handle to a metafile containing the icon view of an object if \c DVASPECT_ICON
	///            is being imposed on an object by a paste special operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_QueryAcceptData,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774325.aspx">IRichEditOleCallback::QueryAcceptData</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	virtual HRESULT STDMETHODCALLTYPE QueryAcceptData(LPDATAOBJECT pDataObject, CLIPFORMAT FAR* pFormat, DWORD operationFlag, BOOL operationIsReal, HGLOBAL hMetaPict);
	/// \brief <em>Indicates if the application should transition into or out of context-sensitive help mode</em>
	///
	/// Indicates if the application should transition into or out of context-sensitive help mode.
	///
	/// \param[in] enterMode If \c TRUE, the application should enter context-sensitive help mode. If
	///            \c FALSE, the application should leave context-sensitive help mode.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774310.aspx">IRichEditOleCallback::ContextSensitiveHelp</a>
	virtual HRESULT STDMETHODCALLTYPE ContextSensitiveHelp(BOOL /*enterMode*/);
	/// \brief <em>Allows the client to supply its own clipboard object</em>
	///
	/// Allows the client to supply its own clipboard object.
	///
	/// \param[in] pCharRange The range (in form of character positions) of the text that the clipboard
	///            object refers to.
	/// \param[in] operationFlag The operation being performed:
	///            - \c RECO_COPY - Data is copied to the clipboard.
	///            - \c RECO_CUT - Data is cut to the clipboard.
	///            - \c RECO_DRAG - Drag operation (drag-and-drop).
	///            - \c RECO_DROP - Drop operation (drag-and-drop).
	///            - \c RECO_PASTE - Data is pasted from the clipboard.
	/// \param[out] ppDataObject Receives the address of the \c IDataObject implementation representing the
	///             range specified in the \c pCharRange parameter.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_GetDataObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774315.aspx">IRichEditOleCallback::GetClipboardData</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787885.aspx">CHARRANGE</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	virtual HRESULT STDMETHODCALLTYPE GetClipboardData(CHARRANGE FAR* pCharRange, DWORD operationFlag, LPDATAOBJECT FAR* ppDataObject);
	/// \brief <em>Allows the client to specify the effects of a drop operation</em>
	///
	/// Allows the client to specify the effects of a drop operation.
	///
	/// \param[in] getSourceEffects \c TRUE if we're within a call to \c IDropTarget::DragEnter or
	///            \c IDropTarget::DragOver. \c FALSE if we're within a call to \c IDropTarget::Drop.
	/// \param[in] keyState Key state as defined by OLE.
	/// \param[out] pEffect Receives the effect used by a rich edit control. When \c getSourceEffects is
	///             \c TRUE, on return, its content is set to the effect allowable by the rich edit control.
	///             When \c getSourceEffects is \c FALSE, on return, the variable is set to the effect to
	///             use.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_GetDragDropEffect,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774319.aspx">IRichEditOleCallback::GetDragDropEffect</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE GetDragDropEffect(BOOL getSourceEffects, DWORD keyState, LPDWORD pEffect);
	/// \brief <em>Queries the application for a context menu to use on a right-click event</em>
	///
	/// Queries the application for a context menu to use on a right-click event.
	///
	/// \param[in] selectionType Selection type. The value, which specifies the contents of the new
	///            selection, can be one or more of the following values:
	///            - \c SEL_EMPTY - The selection is empty.
	///            - \c SEL_TEXT - The selection consists of text.
	///            - \c SEL_OBJECT - The selection contains at least one COM object.
	///            - \c SEL_MULTICHAR - The selection consists of more than one character of text.
	///            - \c SEL_MULTIOBJECT - The selection contains more than one COM object.
	///            - \c GCM_RIGHTMOUSEDROP - Indicates that a context menu for a right-mouse drag drop should
	///              be generated. The \c pObject parameter is a pointer to the \c IDataObject interface for
	///              the object being dropped.
	/// \param[in] pObject Pointer to an interface. If the \c selectionType parameter includes the
	///            \c SEL_OBJECT flag, \c pObject is a pointer to the \c IOleObject interface for the first
	///            selected COM object. If \c selectionType includes the \c GCM_RIGHTMOUSEDROP flag,
	///            \c pObject is a pointer to an \c IDataObject interface. Otherwise, \c pObject is \c NULL.
	/// \param[in] pCharRange The range (in form of character positions) of the current selection.
	/// \param[out] pHMenu Receives the context menu to display. The rich edit control will destroy it after
	///             use.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks When the user selects an item from the context window, a \c WM_COMMAND message is sent to
	///          the parent window of the rich edit control.
	///
	/// \sa Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774310.aspx">IRichEditOleCallback::GetContextMenu</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd542709.aspx">IOleObject</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787885.aspx">CHARRANGE</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_COMMAND</a>
	virtual HRESULT STDMETHODCALLTYPE GetContextMenu(WORD selectionType, LPOLEOBJECT pObject, CHARRANGE FAR* pCharRange, HMENU FAR* pHMenu);
	//@}
	//////////////////////////////////////////////////////////////////////
	
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	/// \param[in] languageID The locale identifier identifying the language in which name should be
	///            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IMouseHookHandler
	///
	//@{
	/// \brief <em>Processes a hooked mouse message</em>
	///
	/// This method is called by the callback function that we defined when we installed a mouse hook to trap
	/// \c WM_MOUSEMOVE messages for the spell checking context menu window.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam The identifier of the mouse message.
	/// \param[in] lParam Points to a \c MOUSEHOOKSTRUCT structure that contains more information.
	/// \param[out] consumeMessage If set to \c TRUE, the message will not be passed to any further handlers
	///             and \c CallNextHookEx will not be called.
	///
	/// \return The value to return if consuming the message.
	///
	/// \sa OnCreate, MouseHookProvider::MouseHookProc, IMouseHookHandler,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644968.aspx">MOUSEHOOKSTRUCT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644988.aspx">MouseProc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644974.aspx">CallNextHookEx</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT HandleMouseMessage(int /*code*/, WPARAM wParam, LPARAM lParam, BOOL& /*consumeMessage*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties, StringProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnKeyDown, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to configure the control window and to raise the \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_ERASEBKGND handler</em>
	///
	/// Will be called if the control's window background must be erased.
	/// We use this handler to set the custom formatting rectangle again, because it is overridden by
	/// Windows.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms648055.aspx">WM_ERASEBKGND</a>
	LRESULT OnEraseBkGnd(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_GETDLGCODE handler</em>
	///
	/// Will be called if the system needs to know which input messages the control wants to handle.
	/// We use this handler for the \c AcceptTabKey property.
	///
	/// \sa OnKeyDown, OnKeyUp, get_AcceptTabKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645425.aspx">WM_GETDLGCODE</a>
	LRESULT OnGetDlgCode(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_INPUTLANGCHANGE handler</em>
	///
	/// Will be called after an application's input language has been changed.
	/// We use this handler to update the IME mode of the control.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632629.aspx">WM_INPUTLANGCHANGE</a>
	LRESULT OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the left mouse
	/// button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnMButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnMButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the control's client area for a previously
	/// specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, OnMouseWheel, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnMouseMove, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534913.aspx">WM_PRINTCLIENT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnXButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_*SCROLL handler</em>
	///
	/// Will be called if the control shall be scrolled.
	/// We use this handler to raise the \c Scrolling event for the case that the scrollbar thumb is dragged,
	/// because in this case no \c EN_HSCROLL and \c EN_VSCROLL notification respectively is sent.
	///
	/// \sa OnReflectedScroll, Raise_Scrolling,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651283.aspx">WM_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651284.aspx">WM_VSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672113.aspx">EN_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672118.aspx">EN_VSCROLL</a>
	LRESULT OnScroll(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTEXT handler</em>
	///
	/// Will be called if the control's text shall be changed.
	/// We use this handler for data-binding.
	///
	/// \sa get_Text,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632644.aspx">WM_SETTEXT</a>
	LRESULT OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_THEMECHANGED handler</em>
	///
	/// Will be called on themable systems if the theme was changed.
	/// We use this handler to update our appearance.
	///
	/// \sa Flags::usingThemes,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632650.aspx">WM_THEMECHANGED</a>
	LRESULT OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called when a timer expires that's associated with the control.
	/// We use this handler for auto-scrolling during drag'n'drop.
	///
	/// \sa AutoScroll,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control was moved.
	/// We use this handler to resize the control on COM level and to raise the \c ResizedControlWindow
	/// event.
	///
	/// \sa Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c DI_GETDRAGIMAGE handler</em>
	///
	/// Will be called during OLE drag'n'drop if the control is queried for a drag image.
	///
	/// \sa OLEDrag, CreateLegacyOLEDragImage, CreateVistaOLEDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	LRESULT OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c EM_SETBKGNDCOLOR handler</em>
	///
	/// Will be called if the control's background color shall be set.
	/// We use this handler to synchronize the \c BackColor property.
	///
	/// \sa get_BackColor, put_BackColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774228.aspx">EM_SETBKGNDCOLOR</a>
	LRESULT OnSetBkgndColor(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c EM_SETUNDOLIMIT handler</em>
	///
	/// Will be called if the maximum size of the control's undo queue shall be set.
	/// We use this handler to synchronize the \c MaxUndoQueueSize property.
	///
	/// \sa get_MaxUndoQueueSize, put_MaxUndoQueueSize,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774290.aspx">EM_SETUNDOLIMIT</a>
	LRESULT OnSetUndoLimit(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c EN_DRAGDROPDONE handler</em>
	///
	/// Will be called if the control's parent window is notified, that a drag'n'drop operation has
	/// completed.
	/// We use this handler to raise the \c DragDropDone event.
	///
	/// \sa Raise_DragDropDone,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787964.aspx">EN_DRAGDROPDONE</a>
	LRESULT OnDragDropDoneNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c EN_LINK handler</em>
	///
	/// Will be called if the control's parent window is notified, that some kind of mouse interaction with a
	/// link occured.
	/// We use this handler for the \c LinkMousePointer property and to raise link specific mouse events.
	///
	/// \sa get_AutoDetectURLs, get_LinkMousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787987.aspx">EN_LINK</a>
	LRESULT OnLinkNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	/// \brief <em>\c EN_REQUESTRESIZE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's content is smaller or
	/// larger than the control's window size. The client application might want to resize the control
	/// window.
	/// We use this handler to raise the \c ShouldResizeControlWindow event.
	///
	/// \sa Raise_ShouldResizeControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787983.aspx">EN_REQUESTRESIZE</a>
	LRESULT OnRequestResizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c EN_SELCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's current selection has
	/// changed.
	/// We use this handler to raise the \c SelectionChanged event.
	///
	/// \sa get_SelectedTextRange, Raise_SelectionChanged, OnReflectedChange,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787987.aspx">EN_SELCHANGE</a>
	LRESULT OnSelChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c EN_CHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's content has changed.
	/// We use this handler to raise the \c TextChanged event.
	///
	/// \sa OnSetText, get_Text, Raise_TextChanged, OnSelChangeNotification,
	///     <a href="https://msdn.microsoft.com/en-us/library/hh768384.aspx">EN_CHANGE (rich edit)</a>
	LRESULT OnReflectedChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_MAXTEXT handler</em>
	///
	/// Will be called if the control's parent window is notified, that the text, that was entered into the
	/// control, got truncated.
	/// We use this handler to raise the \c TruncatedText event.
	///
	/// \sa Raise_TruncatedText,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672115.aspx">EN_MAXTEXT</a>
	LRESULT OnReflectedMaxText(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_*SCROLL handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control is being scrolled.
	/// We use this handler to raise the \c Scrolling event.
	///
	/// \sa OnScroll, Raise_Scrolling,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672113.aspx">EN_HSCROLL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672118.aspx">EN_VSCROLL</a>
	LRESULT OnReflectedScroll(WORD notifyCode, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c EN_SETFOCUS handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has gained the keyboard
	/// focus.
	/// We use this handler to initialize IME.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672116.aspx">EN_SETFOCUS</a>
	LRESULT OnReflectedSetFocus(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c BeginDrag event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the text that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Some of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_BeginDrag, RichTextBoxCtlLibU::_IRichTextBoxEvents::BeginDrag,
	///       OLEDrag, get_AllowDragDrop, Raise_BeginRDrag, RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_BeginDrag, RichTextBoxCtlLibA::_IRichTextBoxEvents::BeginDrag,
	///       OLEDrag, get_AllowDragDrop, Raise_BeginRDrag, RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_BeginDrag(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c BeginRDrag event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the text that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbRightButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Some of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_BeginRDrag, RichTextBoxCtlLibU::_IRichTextBoxEvents::BeginRDrag,
	///       OLEDrag, get_AllowDragDrop, Raise_BeginDrag, RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_BeginRDrag, RichTextBoxCtlLibA::_IRichTextBoxEvents::BeginRDrag,
	///       OLEDrag, get_AllowDragDrop, Raise_BeginDrag, RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_BeginRDrag(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c Click event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the clicked text.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_Click, RichTextBoxCtlLibU::_IRichTextBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick, RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_Click, RichTextBoxCtlLibA::_IRichTextBoxEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick, RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_Click(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
	/// \param[in] menuType Specifies which kind of menu is being requested. Some combinations of the values
	///            defined by the \c MenuTypeConstants enumeration are valid.
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the selected text.
	/// \param[in] selectionType Specifies what kind of data the current selection contains. Any combination
	///            of the values defined by the \c SelectionTypeConstants enumeration is valid.
	/// \param[in] pOLEObject A \c IRichOLEObject object wrapping the (first) embedded OLE object within
	///            the current selection. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] isRightMouseDrop If \c VARIANT_TRUE, the context menu for a drag'n'drop operation with the
	///            right mouse button should be generated. If \c VARIANT_FALSE, the context menu request has
	///            been triggered by a normal right click or by pressing the context menu key (or
	///            [SHIFT]+[F10]) on the keyboard.
	/// \param[in] pDraggedData A \c IOLEDataObject object that holds the dragged data. May be \c NULL.
	/// \param[in,out] pHMenuToDisplay Receives the handle of the context menu to display. The menu will
	///                be destroyed by the control after use.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_ContextMenu,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::ContextMenu, Raise_RClick, TextRange, OLEObject,
	///       TargetOLEDataObject, RichTextBoxCtlLibU::MenuTypeConstants,
	///       RichTextBoxCtlLibU::SelectionTypeConstants, RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_ContextMenu,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::ContextMenu, Raise_RClick, TextRange, OLEObject,
	///       TargetOLEDataObject, RichTextBoxCtlLibA::MenuTypeConstants,
	///       RichTextBoxCtlLibA::SelectionTypeConstants, RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ContextMenu(MenuTypeConstants menuType, IRichTextRange* pTextRange, SelectionTypeConstants selectionType, IRichOLEObject* pOLEObject, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL isRightMouseDrop, IOLEDataObject* pDraggedData, HMENU* pHMenuToDisplay);
	/// \brief <em>Raises the \c DblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] pCharRange A \c CHARRANGE structure specifying the text that was double-clicked. If not
	///            specified, the method performs hit-testing itself.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid. If \c pCharRange is not specified, the
	///            method performs hit-testing itself.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_DblClick, RichTextBoxCtlLibU::_IRichTextBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick,
	///       RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_DblClick, RichTextBoxCtlLibA::_IRichTextBoxEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick,
	///       RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange = NULL, HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0));
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_DestroyedControlWindow,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_DestroyedControlWindow,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(HWND hWnd);
	/// \brief <em>Raises the \c DragDropDone event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event is raised only if the \c RegisterForOLEDragDrop property is set to
	///          \c rfoddNativeDragDrop.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_DragDropDone,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::DragDropDone, get_RegisterForOLEDragDrop,
	///       Raise_OLECompleteDrag
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_DragDropDone,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::DragDropDone, get_RegisterForOLEDragDrop,
	///       Raise_OLECompleteDrag
	/// \endif
	inline HRESULT Raise_DragDropDone(void);
	/// \brief <em>Raises the \c GetDataObject event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the text that the operation is performed
	///            on.
	/// \param[in] operationType Specifies the kind of operation that will take place. Some of the values
	///            defined by the \c ClipboardOperationTypeConstants enumeration is valid.
	/// \param[in,out] ppDataObject The instance of \c IDataObject that is used to transfer data during the
	///                operation. The client application may set this value to a custom object that
	///                implements \c IDataObject.
	/// \param[in,out] pUseCustomDataObject Must be set to \c VARIANT_TRUE by the client application, if the
	///                data object has been replaced with a custom implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_GetDataObject,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::GetDataObject, TextRange,
	///       RichTextBoxCtlLibU::ClipboardOperationTypeConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_GetDataObject,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::GetDataObject, TextRange,
	///       RichTextBoxCtlLibA::ClipboardOperationTypeConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	/// \endif
	inline HRESULT Raise_GetDataObject(IRichTextRange* pTextRange, ClipboardOperationTypeConstants operationType, LONG* ppDataObject, VARIANT_BOOL* pUseCustomDataObject);
	/// \brief <em>Raises the \c GetDragDropEffect event</em>
	///
	/// \param[in] getSourceEffects If \c VARIANT_TRUE, the event is raised to determine the effects allowed
	///            by the drag source. Otherwise it is raised to determine the effects allowed by the drop
	///            target.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] pSkipDefaultProcessing Must be set to \c VARIANT_TRUE to prevent the native rich edit
	///                control chose an effect on its own.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_GetDragDropEffect,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::GetDragDropEffect, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDrop, RichTextBoxCtlLibU::OLEDropEffectConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_GetDragDropEffect,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::GetDragDropEffect, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDrop, RichTextBoxCtlLibA::OLEDropEffectConstants
	/// \endif
	inline HRESULT Raise_GetDragDropEffect(VARIANT_BOOL getSourceEffects, SHORT button, SHORT shift, OLEDropEffectConstants* pEffect, VARIANT_BOOL* pSkipDefaultProcessing);
	/// \brief <em>Raises the \c InsertingOLEObject event</em>
	///
	/// \param[in] pTextRange The insertion position at which the new object is about to be inserted.
	/// \param[in,out] pClassID The class identifier (\c CLSID) of the object that is about to be inserted.
	/// \param[in] pStorage The \c IStorage object holding the object data.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_InsertingOLEObject,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::InsertingOLEObject, Raise_RemovingOLEObject,
	///       TextRange, OLEObject::get_ClassID,
	///       <a href="https://msdn.microsoft.com/en-us/library/aa380015.aspx">IStorage</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_InsertingOLEObject,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::InsertingOLEObject, Raise_RemovingOLEObject,
	///       TextRange, OLEObject::get_ClassID,
	///       <a href="https://msdn.microsoft.com/en-us/library/aa380015.aspx">IStorage</a>
	/// \endif
	inline HRESULT Raise_InsertingOLEObject(IRichTextRange* pTextRange, BSTR* pClassID, LONG pStorage, VARIANT_BOOL* pCancelInsertion);
	/// \brief <em>Raises the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_KeyDown, RichTextBoxCtlLibU::_IRichTextBoxEvents::KeyDown,
	///       Raise_KeyUp, Raise_KeyPress
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_KeyDown, RichTextBoxCtlLibA::_IRichTextBoxEvents::KeyDown,
	///       Raise_KeyUp, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_KeyPress, RichTextBoxCtlLibU::_IRichTextBoxEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_KeyPress, RichTextBoxCtlLibA::_IRichTextBoxEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp
	/// \endif
	inline HRESULT Raise_KeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_KeyUp, RichTextBoxCtlLibU::_IRichTextBoxEvents::KeyUp,
	///       Raise_KeyDown, Raise_KeyPress
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_KeyUp, RichTextBoxCtlLibA::_IRichTextBoxEvents::KeyUp,
	///       Raise_KeyDown, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c MClick event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the clicked text.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MClick, RichTextBoxCtlLibU::_IRichTextBoxEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick, RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MClick, RichTextBoxCtlLibA::_IRichTextBoxEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick, RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MClick(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MDblClick, RichTextBoxCtlLibU::_IRichTextBoxEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick,
	///       RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MDblClick, RichTextBoxCtlLibA::_IRichTextBoxEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick,
	///       RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] pCharRange A \c CHARRANGE structure specifying the text that the mouse cursor is located
	///            over If not specified, the method performs hit-testing itself.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid. If
	///            \c pCharRange is not specified, the method performs hit-testing itself.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseDown, RichTextBoxCtlLibU::_IRichTextBoxEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseDown, RichTextBoxCtlLibA::_IRichTextBoxEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange = NULL, HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0));
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the text that the mouse cursor is located
	///            over.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseEnter, RichTextBoxCtlLibU::_IRichTextBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove, RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseEnter, RichTextBoxCtlLibA::_IRichTextBoxEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove, RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseEnter(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the text that the mouse cursor is located
	///            over.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseHover, RichTextBoxCtlLibU::_IRichTextBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime,
	///       RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseHover, RichTextBoxCtlLibA::_IRichTextBoxEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime,
	///       RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseHover(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the text that the mouse cursor is located
	///            over.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseLeave, RichTextBoxCtlLibU::_IRichTextBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove, RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseLeave, RichTextBoxCtlLibA::_IRichTextBoxEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove, RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseLeave(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] pCharRange A \c CHARRANGE structure specifying the text that the mouse cursor is located
	///            over If not specified, the method performs hit-testing itself.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid. If
	///            \c pCharRange is not specified, the method performs hit-testing itself.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseMove, RichTextBoxCtlLibU::_IRichTextBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseMove, RichTextBoxCtlLibA::_IRichTextBoxEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange = NULL, HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0));
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] pCharRange A \c CHARRANGE structure specifying the text that the mouse cursor is located
	///            over If not specified, the method performs hit-testing itself.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid. If
	///            \c pCharRange is not specified, the method performs hit-testing itself.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseUp, RichTextBoxCtlLibU::_IRichTextBoxEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseUp, RichTextBoxCtlLibA::_IRichTextBoxEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick,
	///       RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange = NULL, HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0));
	/// \brief <em>Raises the \c MouseWheel event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseWheel, RichTextBoxCtlLibU::_IRichTextBoxEvents::MouseWheel,
	///       Raise_MouseMove, RichTextBoxCtlLibU::ExtendedMouseButtonConstants,
	///       RichTextBoxCtlLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_MouseWheel, RichTextBoxCtlLibA::_IRichTextBoxEvents::MouseWheel,
	///       Raise_MouseMove, RichTextBoxCtlLibA::ExtendedMouseButtonConstants,
	///       RichTextBoxCtlLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event is raised only if the drag'n'drop operation has been initiated by calling
	///          \c OLEDrag.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLECompleteDrag,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLECompleteDrag, Raise_OLEStartDrag,
	///       RichTextBoxCtlLibU::IOLEDataObject::GetData, OLEDrag, Raise_DragDropDone
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLECompleteDrag,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLECompleteDrag, Raise_OLEStartDrag,
	///       RichTextBoxCtlLibA::IOLEDataObject::GetData, OLEDrag, Raise_DragDropDone
	/// \endif
	inline HRESULT Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect);
	/// \brief <em>Raises the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client finally executed.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragDrop,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEDragDrop, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_GetDragDropEffect, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragDrop,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEDragDrop, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_GetDragDropEffect, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data. If \c NULL, the cached data object is used. We use this when
	///            we call this method from other places than \c DragEnter.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragEnter,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEDragEnter, Raise_OLEDragMouseMove,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_GetDragDropEffect, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragEnter,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEDragEnter,Raise_OLEDragMouseMove,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_GetDragDropEffect, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The potential drop target window's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragEnterPotentialTarget,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEDragEnterPotentialTarget,
	///       Raise_OLEDragLeavePotentialTarget, OLEDrag
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragEnterPotentialTarget,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEDragEnterPotentialTarget,
	///       Raise_OLEDragLeavePotentialTarget, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragLeave,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEDragLeave, Raise_OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragLeave,
	///        RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEDragLeave,Raise_OLEDragEnter,
	///        Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(void);
	/// \brief <em>Raises the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragLeavePotentialTarget,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEDragLeavePotentialTarget,
	///       Raise_OLEDragEnterPotentialTarget, OLEDrag
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragLeavePotentialTarget,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEDragLeavePotentialTarget,
	///       Raise_OLEDragEnterPotentialTarget, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragLeavePotentialTarget(void);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragMouseMove,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEDragMouseMove, Raise_OLEDragEnter,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_GetDragDropEffect, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEDragMouseMove,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEDragMouseMove, Raise_OLEDragEnter,
	///       Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_GetDragDropEffect, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c DROPEFFECT enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEGiveFeedback,
	///        RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEGiveFeedback,Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEGiveFeedback,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEGiveFeedback, Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors);
	/// \brief <em>Raises the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c TRUE, the user has pressed the \c ESC key since the last time
	///            this event was raised; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue (\c S_OK), cancel
	///                (\c DRAGDROP_S_CANCEL) or complete (\c DRAGDROP_S_DROP) the drag'n'drop operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEQueryContinueDrag,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEQueryContinueDrag,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \endif
	inline HRESULT Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith);
	/// \brief <em>Raises the \c OLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEReceivedNewData,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEReceivedNewData,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLESetData, RichTextBoxCtlLibU::_IRichTextBoxEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLESetData, RichTextBoxCtlLibA::_IRichTextBoxEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event is raised only if the drag'n'drop operation has been initiated by calling
	///          \c OLEDrag.\n
	///          This event won't be fired if a custom \c IDataObject implementation was passed to
	///          the \c OLEDrag method.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEStartDrag,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::OLEStartDrag, Raise_OLESetData, Raise_OLECompleteDrag,
	///       SourceOLEDataObject::SetData, OLEDrag
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_OLEStartDrag,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::OLEStartDrag, Raise_OLESetData, Raise_OLECompleteDrag,
	///       SourceOLEDataObject::SetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEStartDrag(IOLEDataObject* pData);
	/// \brief <em>Raises the \c QueryAcceptData event</em>
	///
	/// \param[in] pData The object that holds the copied or dragged data.
	/// \param[in,out] pFormatID An integer value specifying the format in which the copied or dragged
	///                data will be pasted or dropped. The client application may change the format to
	///                any other format provided by the data object.\n
	///                Valid values are those defined by VB's \c ClipBoardConstants enumeration, but
	///                also any other format that has been registered using the \c RegisterClipboardFormat
	///                API function.
	/// \param[in] operationType Specifies the kind of operation that will take place. Some of the values
	///            defined by the \c ClipboardOperationTypeConstants enumeration are valid.
	/// \param[in] performOperation If \c VARIANT_TRUE, the operation is actually being performed; otherwise
	///            the event is fired only to check whether the data would be accepted or refused.
	/// \param[in] hMetafilePicture A handle to a metafile image that should be used to display the
	///            inserted object as icon.
	/// \param[in,out] acceptData Specifies whether to accept or refuse the provided data or whether to
	///                let the native rich edit control decide. Any of the values defined by the
	///                \c QueryAcceptDataConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_QueryAcceptData,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::QueryAcceptData, Raise_GetDataObject,
	///       IOLEDataObject, RichTextBoxCtlLibU::ClipboardOperationTypeConstants,
	///       RichTextBoxCtlLibU::QueryAcceptDataConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_QueryAcceptData,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::QueryAcceptData, Raise_GetDataObject,
	///       IOLEDataObject, RichTextBoxCtlLibA::ClipboardOperationTypeConstants,
	///       RichTextBoxCtlLibA::QueryAcceptDataConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>
	/// \endif
	inline HRESULT Raise_QueryAcceptData(IOLEDataObject* pData, LONG* pFormatID, ClipboardOperationTypeConstants operationType, VARIANT_BOOL performOperation, OLE_HANDLE hMetafilePicture, QueryAcceptDataConstants* pAcceptData);
	/// \brief <em>Raises the \c RClick event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the clicked text.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_RClick, RichTextBoxCtlLibU::_IRichTextBoxEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick,
	///       RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_RClick, RichTextBoxCtlLibA::_IRichTextBoxEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick,
	///       RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_RClick(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] pCharRange A \c CHARRANGE structure specifying the text that was double-clicked. If not
	///            specified, the method performs hit-testing itself.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid. If \c pCharRange is not specified, the
	///            method performs hit-testing itself.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_RDblClick, RichTextBoxCtlLibU::_IRichTextBoxEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick,
	///       RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_RDblClick, RichTextBoxCtlLibA::_IRichTextBoxEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick,
	///       RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, CHARRANGE* pCharRange = NULL, HitTestConstants hitTestDetails = static_cast<HitTestConstants>(0));
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_RecreatedControlWindow,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_RecreatedControlWindow,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(HWND hWnd);
	/// \brief <em>Raises the \c RemovingOLEObject event</em>
	///
	/// \param[in] pOLEObject The object being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_RemovingOLEObject,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::RemovingOLEObject, Raise_InsertingOLEObject
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_RemovingOLEObject,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::RemovingOLEObject, Raise_InsertingOLEObject
	/// \endif
	inline HRESULT Raise_RemovingOLEObject(IRichOLEObject* pOLEObject, VARIANT_BOOL* pCancelDeletion);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_ResizedControlWindow,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_ResizedControlWindow,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c Scrolling event</em>
	///
	/// \param[in] axis The axis which is scrolled. Any of the values defined by the \c ScrollAxisConstants
	///            enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_Scrolling, RichTextBoxCtlLibU::_IRichTextBoxEvents::Scrolling,
	///       RichTextBoxCtlLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_Scrolling, RichTextBoxCtlLibA::_IRichTextBoxEvents::Scrolling,
	///       RichTextBoxCtlLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_Scrolling(ScrollAxisConstants axis);
	/// \brief <em>Raises the \c SelectionChanged event</em>
	///
	/// \param[in] pSelectedTextRange A \c IRichTextRange object wrapping the (new) selected text.
	/// \param[in] selectionType Specifies what kind of data the current selection contains. Any
	///            combination of the values defined by the \c SelectionTypeConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_SelectionChanged,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::SelectionChanged, IRichTextRange,
	///       get_SelectedTextRange, RichTextBoxCtlLibU::SelectionTypeConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_SelectionChanged,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::SelectionChanged, IRichTextRange,
	///       get_SelectedTextRange, RichTextBoxCtlLibA::SelectionTypeConstants
	/// \endif
	inline HRESULT Raise_SelectionChanged(IRichTextRange* pSelectedTextRange, SelectionTypeConstants selectionType);
	/// \brief <em>Raises the \c ShouldResizeControlWindow event</em>
	///
	/// \param[in] pSuggestedControlSize The rectangle that the control should be resized to.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_ShouldResizeControlWindow,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::ShouldResizeControlWindow,
	///       CheckForOptimalControlWindowSize, TBarCtlsLibU::RECTANGLE
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_ShouldResizeControlWindow,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::ShouldResizeControlWindow,
	///       CheckForOptimalControlWindowSize, TBarCtlsLibA::RECTANGLE
	/// \endif
	inline HRESULT Raise_ShouldResizeControlWindow(RECTANGLE* pSuggestedControlSize);
	/// \brief <em>Raises the \c TextChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default event.\n
	///          This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_TextChanged,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::TextChanged
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_TextChanged,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::TextChanged
	/// \endif
	inline HRESULT Raise_TextChanged(void);
	/// \brief <em>Raises the \c TruncatedText event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_TruncatedText,
	///       RichTextBoxCtlLibU::_IRichTextBoxEvents::TruncatedText
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_TruncatedText,
	///       RichTextBoxCtlLibA::_IRichTextBoxEvents::TruncatedText
	/// \endif
	inline HRESULT Raise_TruncatedText(void);
	/// \brief <em>Raises the \c XClick event</em>
	///
	/// \param[in] pTextRange A \c IRichTextRange object wrapping the clicked text.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_XClick, RichTextBoxCtlLibU::_IRichTextBoxEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick, RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_XClick, RichTextBoxCtlLibA::_IRichTextBoxEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick, RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_XClick(IRichTextRange* pTextRange, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IRichTextBoxEvents::Fire_XDblClick, RichTextBoxCtlLibU::_IRichTextBoxEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick
	/// \else
	///   \sa Proxy_IRichTextBoxEvents::Fire_XDblClick, RichTextBoxCtlLibA::_IRichTextBoxEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleControl
	///
	//@{
	/// \brief <em>Implementation of \c IOleControl::GetControlInfo</em>
	///
	/// Will be called if the container requests details about the control's keyboard mnemonics and keyboard
	/// behavior.
	///
	/// \param[in, out] pControlInfo The requested details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered actions (verbs) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	//////////////////////////////////////////////////////////////////////
	/// \name Selection changes
	///
	//@{
	/// \brief <em>Increments the \c silentSelectionChanges flag</em>
	///
	/// \sa LeaveSilentSelectionChangesSection, Flags::silentSelectionChanges, HitTest, Raise_SelectionChanged
	void EnterSilentSelectionChangesSection(void);
	/// \brief <em>Decrements the \c silentSelectionChanges flag</em>
	///
	/// \sa EnterSilentSelectionChangesSection, Flags::silentSelectionChanges, HitTest, Raise_SelectionChanged
	void LeaveSilentSelectionChangesSection(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Object insertions
	///
	//@{
	/// \brief <em>Increments the \c silentObjectInsertions flag</em>
	///
	/// \sa LeaveSilentObjectInsertionsSection, Flags::silentObjectInsertions, Raise_InsertingOLEObject
	void EnterSilentObjectInsertionsSection(void);
	/// \brief <em>Decrements the \c silentObjectInsertions flag</em>
	///
	/// \sa EnterSilentObjectInsertionsSection, Flags::silentObjectInsertions, Raise_InsertingOLEObject
	void LeaveSilentObjectInsertionsSection(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Object deletions
	///
	//@{
	/// \brief <em>Increments the \c silentObjectDeletions flag</em>
	///
	/// \sa LeaveSilentObjectDeletionsSection, Flags::silentObjectDeletions, Raise_RemovingOLEObject
	void EnterSilentObjectDeletionsSection(void);
	/// \brief <em>Decrements the \c silentObjectDeletions flag</em>
	///
	/// \sa EnterSilentObjectDeletionsSection, Flags::silentObjectDeletions, Raise_RemovingOLEObject
	void LeaveSilentObjectDeletionsSection(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name MFC clones
	///
	//@{
	/// \brief <em>A rewrite of MFC's \c COleControl::RecreateControlWindow</em>
	///
	/// Destroys and re-creates the control window.
	///
	/// \remarks This rewrite probably isn't complete.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
	void RecreateControlWindow(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name IME support
	///
	//@{
	/// \brief <em>Retrieves a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is requested.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa SetCurrentIMEContextMode, RichTextBoxCtlLibU::IMEModeConstants, get_IMEMode
	/// \else
	///   \sa SetCurrentIMEContextMode, RichTextBoxCtlLibA::IMEModeConstants, get_IMEMode
	/// \endif
	IMEModeConstants GetCurrentIMEContextMode(HWND hWnd);
	/// \brief <em>Sets a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is set.
	/// \param[in] IMEMode A constant out of the \c IMEModeConstants enumeration specifying the IME
	///            context mode to apply.
	///
	/// \if UNICODE
	///   \sa GetCurrentIMEContextMode, RichTextBoxCtlLibU::IMEModeConstants, put_IMEMode
	/// \else
	///   \sa GetCurrentIMEContextMode, RichTextBoxCtlLibA::IMEModeConstants, put_IMEMode
	/// \endif
	void SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode);
	/// \brief <em>Retrieves the control's effective IME context mode</em>
	///
	/// Retrieves the IME context mode that is set for the control after resolving recursive modes like
	/// \c imeInherit.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::IMEModeConstants, get_IMEMode
	/// \else
	///   \sa RichTextBoxCtlLibA::IMEModeConstants, get_IMEMode
	/// \endif
	IMEModeConstants GetEffectiveIMEMode(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Control window configuration
	///
	//@{
	/// \brief <em>Retrieves the rich edit window class name to use and loads the required dll</em>
	///
	/// We can use 5 different versions of rich edit:
	/// - 4.1 - available on Windows XP and newer
	/// - 5.0 - requires Office 2003 or newer
	/// - 6.0 - requires Office 2007 or newer
	/// - 7.0 - requires Office 2010 or newer
	/// - 7.5 - available on Windows 8 or newer
	/// - 8.0 - requires Office 2013 or newer
	/// The versions that are part of Windows, are located in msftedit.dll and use a different window class
	/// name than the versions that are part of Office, which are located in riched20.dll.\n
	/// \c GetWndClassName retrieves the class name to use, based on the \c RichEditVersion property. It also
	/// loads the required dll - either msftedit.dll from the system directory or a riched20.dll from the
	/// Office installation directory.
	///
	/// \return The window class name to use.
	///
	/// \sa get_RichEditVersion, Create
	LPCTSTR GetWndClassName(void);
	/// \brief <em>Calculates the extended window style bits</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetStyleBits, SendConfigurationMessages, Create
	DWORD GetExStyleBits(void);
	/// \brief <em>Calculates the window style bits</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* constants specifying the required window styles.
	///
	/// \sa GetExStyleBits, SendConfigurationMessages, Create
	DWORD GetStyleBits(void);
	/// \brief <em>Configures the control by sending messages</em>
	///
	/// Sends \c WM_* and \c CB_* messages to the control window to make it match the current property
	/// settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa GetExStyleBits, GetStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates a \c MousePointerConstants type mouse cursor into a \c HCURSOR type mouse cursor</em>
	///
	/// \param[in] mousePointer The \c MousePointerConstants type mouse cursor to translate.
	///
	/// \return The translated \c HCURSOR type mouse cursor.
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \else
	///   \sa RichTextBoxCtlLibA::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \endif
	HCURSOR MousePointerConst2hCursor(MousePointerConstants mousePointer);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves details about the part of the control that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	///
	/// \return A combination of \c HitTestConstants flags classifying the point ('x'; 'y').
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::HitTestConstants
	/// \else
	///   \sa RichTextBoxCtlLibA::HitTestConstants
	/// \endif
	HitTestConstants HitTest(LONG x, LONG y);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);
	/// \brief <em>Auto-scrolls the control</em>
	///
	/// \sa OnTimer, Raise_DragMouseMove, DragDropStatus::AutoScrolling
	void AutoScroll(void);
	/// \brief <em>Retrieves whether the specified point is located over selected text</em>
	///
	/// \param[in] pt The point to check.
	///
	/// \return \c TRUE if the specified point is located over selected text; otherwise \c FALSE.
	///
	/// \sa get_SelectedTextRange
	BOOL IsOverSelectedText(LPARAM pt);

	/// \brief <em>Retrieves whether the logical left mouse button is held down</em>
	///
	/// \return \c TRUE if the logical left mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsRightMouseButtonDown
	BOOL IsLeftMouseButtonDown(void);
	/// \brief <em>Retrieves whether the logical right mouse button is held down</em>
	///
	/// \return \c TRUE if the logical right mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsLeftMouseButtonDown
	BOOL IsRightMouseButtonDown(void);


	/// \brief <em>Holds constants and flags used with IME support</em>
	struct IMEFlags
	{
	protected:
		/// \brief <em>A table of IME modes to use for Chinese input language</em>
		///
		/// \sa GetIMECountryTable, japaneseIMETable, koreanIMETable
		static IMEModeConstants chineseIMETable[10];
		/// \brief <em>A table of IME modes to use for Japanese input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants japaneseIMETable[10];
		/// \brief <em>A table of IME modes to use for Korean input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants koreanIMETable[10];

	public:
		/// \brief <em>The handle of the default IME context</em>
		HIMC hDefaultIMC;

		IMEFlags()
		{
			hDefaultIMC = NULL;
		}

		/// \brief <em>Retrieves a table of IME modes to use for the current keyboard layout</em>
		///
		/// Retrieves a table of IME modes which can be used to map \c IME_CMODE_* constants to
		/// \c IMEModeConstants constants. The table depends on the current keyboard layout.
		///
		/// \param[in,out] table The IME mode table for the currently active keyboard layout.
		///
		/// \if UNICODE
		///   \sa RichTextBoxCtlLibU::IMEModeConstants, GetCurrentIMEContextMode
		/// \else
		///   \sa RichTextBoxCtlLibA::IMEModeConstants, GetCurrentIMEContextMode
		/// \endif
		static void GetIMECountryTable(IMEModeConstants table[10])
		{
			WORD languageID = LOWORD(GetKeyboardLayout(0));
			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				switch(languageID) {
					case MAKELANGID(LANG_JAPANESE, SUBLANG_DEFAULT):
						CopyMemory(table, japaneseIMETable, sizeof(japaneseIMETable));
						return;
						break;
					case MAKELANGID(LANG_KOREAN, SUBLANG_DEFAULT):
						CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
						return;
						break;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
				if(languageID == MAKELANGID(LANG_KOREAN, SUBLANG_SYS_DEFAULT)) {
					CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
					return;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if((languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SINGAPORE)) && (languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_MACAU))) {
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
		}
	} IMEFlags;

	/// \brief <em>Holds a control instance's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>Holds a font property's settings</em>
		typedef struct FontProperty
		{
		protected:
			/// \brief <em>Holds the control's default font</em>
			///
			/// \sa GetDefaultFont
			static FONTDESC defaultFont;

		public:
			/// \brief <em>Holds whether we're listening for events fired by the font object</em>
			///
			/// If greater than 0, we're advised to the \c IFontDisp object identified by \c pFontDisp. I. e.
			/// we will be notified if a property of the font object changes. If 0, we won't receive any events
			/// fired by the \c IFontDisp object.
			///
			/// \sa pFontDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>Flag telling \c OnSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnSetFont
			UINT dontGetFontObject : 1;
			/// \brief <em>The control's current font</em>
			///
			/// \sa ApplyFont, owningFontResource
			CFont currentFont;
			/// \brief <em>If \c TRUE, \c currentFont may destroy the font resource; otherwise not</em>
			///
			/// \sa currentFont
			UINT owningFontResource : 1;
			/// \brief <em>A pointer to the font object's implementation of \c IFontDisp</em>
			IFontDisp* pFontDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<RichTextBox> >* pPropertyNotifySink;

			FontProperty()
			{
				watching = 0;
				dontGetFontObject = FALSE;
				owningFontResource = TRUE;
				pFontDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~FontProperty()
			{
				Release();
			}

			FontProperty& operator =(const FontProperty& source)
			{
				Release();

				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				pFontDisp = source.pFontDisp;
				if(pFontDisp) {
					pFontDisp->AddRef();
				}
				owningFontResource = source.owningFontResource;
				if(!owningFontResource) {
					currentFont.Attach(source.currentFont.m_hFont);
				}
				dontGetFontObject = source.dontGetFontObject;

				if(source.watching > 0) {
					StartWatching();
				}

				return *this;
			}

			/// \brief <em>Retrieves a default font that may be used</em>
			///
			/// \return A \c FONTDESC structure containing the default font.
			///
			/// \sa defaultFont
			static FONTDESC GetDefaultFont(void)
			{
				return defaultFont;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(RichTextBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<RichTextBox> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(owningFontResource) {
					if(!currentFont.IsNull()) {
						currentFont.DeleteObject();
					}
				} else {
					currentFont.Detach();
				}
				if(pFontDisp) {
					pFontDisp->Release();
					pFontDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pFontDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} FontProperty;

		/// \brief <em>Holds a picture property's settings</em>
		typedef struct PictureProperty
		{
			/// \brief <em>Holds whether we're listening for events fired by the picture object</em>
			///
			/// If greater than 0, we're advised to the \c IPictureDisp object identified by \c pPictureDisp.
			/// I. e. we will be notified if a property of the picture object changes. If 0, we won't receive any
			/// events fired by the \c IPictureDisp object.
			///
			/// \sa pPictureDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>A pointer to the picture object's implementation of \c IPictureDisp</em>
			IPictureDisp* pPictureDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<RichTextBox> >* pPropertyNotifySink;

			PictureProperty()
			{
				watching = 0;
				pPictureDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~PictureProperty()
			{
				Release();
			}

			PictureProperty& operator =(const PictureProperty& source)
			{
				Release();

				pPictureDisp = source.pPictureDisp;
				if(pPictureDisp) {
					pPictureDisp->AddRef();
				}
				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				if(source.watching > 0) {
					StartWatching();
				}
				return *this;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(RichTextBox* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<RichTextBox> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(pPictureDisp) {
					pPictureDisp->Release();
					pPictureDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pPictureDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} PictureProperty;

		/// \brief <em>The autoscroll-zone's default width</em>
		///
		/// The default width (in pixels) of the border around the control's client area, that's
		/// sensitive for auto-scrolling during a drag'n'drop operation. If the mouse cursor's position
		/// lies within this area during a drag'n'drop operation, the control will be auto-scrolled.
		///
		/// \sa dragScrollTimeBase, Raise_DragMouseMove
		static const int DRAGSCROLLZONEWIDTH = 16;

		/// \brief <em>Holds the \c AcceptTabKey property's setting</em>
		///
		/// \sa get_AcceptTabKey, put_AcceptTabKey
		UINT acceptTabKey : 1;
		/// \brief <em>Holds the \c AdjustLineHeightForEastAsianLanguages property's setting</em>
		///
		/// \sa get_AdjustLineHeightForEastAsianLanguages, put_AdjustLineHeightForEastAsianLanguages
		UINT adjustLineHeightForEastAsianLanguages : 1;
		/// \brief <em>Holds the \c AllowInputThroughTSF property's setting</em>
		///
		/// \sa get_AllowInputThroughTSF, put_AllowInputThroughTSF
		UINT allowInputThroughTSF : 1;
		/// \brief <em>Holds the \c AllowMathZoneInsertion property's setting</em>
		///
		/// \sa get_AllowMathZoneInsertion, put_AllowMathZoneInsertion
		UINT allowMathZoneInsertion : 1;
		/// \brief <em>Holds the \c AllowObjectInsertionThroughTSF property's setting</em>
		///
		/// \sa get_AllowObjectInsertionThroughTSF, put_AllowObjectInsertionThroughTSF
		UINT allowObjectInsertionThroughTSF : 1;
		/// \brief <em>Holds the \c AllowTableInsertion property's setting</em>
		///
		/// \sa get_AllowTableInsertion, put_AllowTableInsertion
		UINT allowTableInsertion : 1;
		/// \brief <em>Holds the \c AllowTSFProofingTips property's setting</em>
		///
		/// \sa get_AllowTSFProofingTips, put_AllowTSFProofingTips
		UINT allowTSFProofingTips : 1;
		/// \brief <em>Holds the \c AllowTSFSmartTags property's setting</em>
		///
		/// \sa get_AllowTSFSmartTags, put_AllowTSFSmartTags
		UINT allowTSFSmartTags : 1;
		/// \brief <em>Holds the \c AlwaysShowScrollBars property's setting</em>
		///
		/// \sa get_AlwaysShowScrollBars, put_AlwaysShowScrollBars
		UINT alwaysShowScrollBars : 1;
		/// \brief <em>Holds the \c AlwaysShowSelection property's setting</em>
		///
		/// \sa get_AlwaysShowSelection, put_AlwaysShowSelection
		UINT alwaysShowSelection : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c AutoDetectedURLSchemes property's setting</em>
		///
		/// \sa get_AutoDetectedURLSchemes, put_AutoDetectedURLSchemes
		CComBSTR autoDetectedURLSchemes;
		/// \brief <em>Holds the \c AutoDetectURLs property's setting</em>
		///
		/// \sa get_AutoDetectURLs, put_AutoDetectURLs
		AutoDetectURLsConstants autoDetectURLs;
		/// \brief <em>Holds the \c AutoScrolling property's setting</em>
		///
		/// \sa get_AutoScrolling, put_AutoScrolling
		AutoScrollingConstants autoScrolling;
		/// \brief <em>Holds the \c AutoSelectWordOnTrackSelection property's setting</em>
		///
		/// \sa get_AutoSelectWordOnTrackSelection, put_AutoSelectWordOnTrackSelection
		UINT autoSelectWordOnTrackSelection : 1;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
		/// \brief <em>Holds the \c BeepOnMaxText property's setting</em>
		///
		/// \sa get_BeepOnMaxText, put_BeepOnMaxText
		UINT beepOnMaxText : 1;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c CharacterConversion property's setting</em>
		///
		/// \sa get_CharacterConversion, put_CharacterConversion
		CharacterConversionConstants characterConversion;
		/// \brief <em>Holds the \c DefaultMathZoneHAlignment property's setting</em>
		///
		/// \sa get_DefaultMathZoneHAlignment, put_DefaultMathZoneHAlignment
		HAlignmentConstants defaultMathZoneHAlignment;
		/// \brief <em>Holds the \c DefaultTabWidth property's setting</em>
		///
		/// \sa get_DefaultTabWidth, put_DefaultTabWidth
		FLOAT defaultTabWidth;
		/// \brief <em>Holds the \c DenoteEmptyMathArguments property's setting</em>
		///
		/// \sa get_DenoteEmptyMathArguments, put_DenoteEmptyMathArguments
		DenoteEmptyMathArgumentsConstants denoteEmptyMathArguments;
		/// \brief <em>Holds the \c DetectDragDrop property's setting</em>
		///
		/// \sa get_DetectDragDrop, put_DetectDragDrop
		UINT detectDragDrop : 1;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DisableIMEOperations property's setting</em>
		///
		/// \sa get_DisableIMEOperations, put_DisableIMEOperations
		UINT disableIMEOperations : 1;
		// \brief <em>Holds the \c DiscardCompositionStringOnIMECancel property's setting</em>
		//
		// \sa get_DiscardCompositionStringOnIMECancel, put_DiscardCompositionStringOnIMECancel
		//UINT discardCompositionStringOnIMECancel : 1;
		/// \brief <em>Holds the \c DisplayHyperlinkTooltips property's setting</em>
		///
		/// \sa get_DisplayHyperlinkTooltips, put_DisplayHyperlinkTooltips
		UINT displayHyperlinkTooltips : 1;
		/// \brief <em>Holds the \c DisplayZeroWidthTableCellBorders property's setting</em>
		///
		/// \sa get_DisplayZeroWidthTableCellBorders, put_DisplayZeroWidthTableCellBorders
		UINT displayZeroWidthTableCellBorders : 1;
		/// \brief <em>Holds the \c DraftMode property's setting</em>
		///
		/// \sa get_DraftMode, put_DraftMode
		UINT draftMode : 1;
		/// \brief <em>Holds the \c DragScrollTimeBase property's setting</em>
		///
		/// \sa get_DragScrollTimeBase, put_DragScrollTimeBase
		long dragScrollTimeBase;
		/// \brief <em>Holds the \c DropWordsOnWordBoundariesOnly property's setting</em>
		///
		/// \sa get_DropWordsOnWordBoundariesOnly, put_DropWordsOnWordBoundariesOnly
		UINT dropWordsOnWordBoundariesOnly : 1;
		/// \brief <em>Holds the \c EmulateSimpleTextBox property's setting</em>
		///
		/// \sa get_EmulateSimpleTextBox, put_EmulateSimpleTextBox
		UINT emulateSimpleTextBox : 1;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c ExtendFontBackColorToWholeLine property's setting</em>
		///
		/// \sa get_ExtendFontBackColorToWholeLine, put_ExtendFontBackColorToWholeLine
		UINT extendFontBackColorToWholeLine : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the \c FormattingRectangle* properties' settings</em>
		///
		/// \sa get_FormattingRectangleHeight, put_FormattingRectangleHeight,
		///     get_FormattingRectangleLeft, put_FormattingRectangleLeft,
		///     get_FormattingRectangleTop, put_FormattingRectangleTop,
		///     get_FormattingRectangleWidth, put_FormattingRectangleWidth
		CRect formattingRectangle;
		/// \brief <em>Holds the \c GrowNAryOperators property's setting</em>
		///
		/// \sa get_GrowNAryOperators, put_GrowNAryOperators
		UINT growNAryOperators : 1;
		// \brief <em>Holds the \c HAlignment property's setting</em>
		//
		// \sa get_HAlignment, put_HAlignment
		//HAlignmentConstants hAlignment;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c HyphenationFunction property's setting</em>
		///
		/// \sa get_HyphenationFunction, put_HyphenationFunction
		LPVOID hyphenationFunction;
		/// \brief <em>Holds the \c HyphenationWordWrapZoneWidth property's setting</em>
		///
		/// \sa get_HyphenationWordWrapZoneWidth, put_HyphenationWordWrapZoneWidth
		short hyphenationWordWrapZoneWidth;
		/// \brief <em>Holds the \c IMEConversionMode property's setting</em>
		///
		/// \sa get_IMEConversionMode, put_IMEConversionMode
		IMEConversionModeConstants IMEConversionMode;
		/// \brief <em>Holds the \c IMEMode property's setting</em>
		///
		/// \sa get_IMEMode, put_IMEMode
		IMEModeConstants IMEMode;
		/// \brief <em>Holds the \c IntegralLimitsLocation property's setting</em>
		///
		/// \sa get_IntegralLimitsLocation, put_IntegralLimitsLocation
		LimitsLocationConstants integralLimitsLocation;
		/// \brief <em>Holds the \c LeftMargin property's setting</em>
		///
		/// \sa get_LeftMargin, put_LeftMargin
		OLE_XSIZE_PIXELS leftMargin;
		/// \brief <em>Holds the \c LetClientHandleAllIMEOperations property's setting</em>
		///
		/// \sa get_LetClientHandleAllIMEOperations, put_LetClientHandleAllIMEOperations
		UINT letClientHandleAllIMEOperations : 1;
		/// \brief <em>Holds the \c LinkMouseIcon property's settings</em>
		///
		/// \sa get_LinkMouseIcon, put_LinkMouseIcon, putref_LinkMouseIcon
		PictureProperty linkMouseIcon;
		/// \brief <em>Holds the \c LinkMousePointer property's setting</em>
		///
		/// \sa get_LinkMousePointer, put_LinkMousePointer
		MousePointerConstants linkMousePointer;
		/// \brief <em>Holds the \c LogicalCaret property's setting</em>
		///
		/// \sa get_LogicalCaret, put_LogicalCaret
		UINT logicalCaret : 1;
		/// \brief <em>Holds the \c MathLineBreakBehavior property's setting</em>
		///
		/// \sa get_MathLineBreakBehavior, put_MathLineBreakBehavior
		MathLineBreakBehaviorConstants mathLineBreakBehavior;
		/// \brief <em>Holds the \c MaxTextLength property's setting</em>
		///
		/// \sa get_MaxTextLength, put_MaxTextLength
		long maxTextLength;
		/// \brief <em>Holds the \c MaxUndoQueueSize property's setting</em>
		///
		/// \sa get_MaxUndoQueueSize, put_MaxUndoQueueSize
		long maxUndoQueueSize;
		/// \brief <em>Holds the \c Modified property's setting</em>
		///
		/// \sa get_Modified, put_Modified
		UINT modified : 1;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c MultiLine property's setting</em>
		///
		/// \sa get_MultiLine, put_MultiLine
		UINT multiLine : 1;
		/// \brief <em>Holds the \c MultiSelect property's setting</em>
		///
		/// \sa get_MultiSelect, put_MultiSelect
		UINT multiSelect : 1;
		/// \brief <em>Holds the \c NAryLimitsLocation property's setting</em>
		///
		/// \sa get_NAryLimitsLocation, put_NAryLimitsLocation
		LimitsLocationConstants nAryLimitsLocation;
		/// \brief <em>Holds the \c NoInputSequenceCheck property's setting</em>
		///
		/// \sa get_NoInputSequenceCheck, put_NoInputSequenceCheck
		UINT noInputSequenceCheck : 1;
		/// \brief <em>Holds the \c OLEDragImageStyle property's setting</em>
		///
		/// \sa get_OLEDragImageStyle, put_OLEDragImageStyle
		OLEDragImageStyleConstants oleDragImageStyle;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c RawSubScriptAndSuperScriptOperators property's setting</em>
		///
		/// \sa get_RawSubScriptAndSuperScriptOperators, put_RawSubScriptAndSuperScriptOperators
		UINT rawSubScriptAndSuperScriptOperators : 1;
		/// \brief <em>Holds the \c ReadOnly property's setting</em>
		///
		/// \sa get_ReadOnly, put_ReadOnly
		UINT readOnly : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		RegisterForOLEDragDropConstants registerForOLEDragDrop;
		/// \brief <em>Holds the \c RichEditAPIVersion property's setting</em>
		///
		/// \sa get_RichEditAPIVersion
		RichEditAPIVersionConstants richEditAPIVersion;
		/// \brief <em>Holds the \c RichEditLibraryPath property's setting</em>
		///
		/// \sa get_RichEditLibraryPath
		CComBSTR richEditLibraryPath;
		/// \brief <em>Holds the \c RichEditVersion property's setting</em>
		///
		/// \sa get_RichEditVersion, put_RichEditVersion
		RichEditVersionConstants richEditVersion;
		/// \brief <em>Holds the \c RightMargin property's setting</em>
		///
		/// \sa get_RightMargin, put_RightMargin
		OLE_XSIZE_PIXELS rightMargin;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c RightToLeftMathZones property's setting</em>
		///
		/// \sa get_RightToLeftMathZones, put_RightToLeftMathZones
		UINT rightToLeftMathZones : 1;
		/// \brief <em>Holds the \c ScrollBars property's setting</em>
		///
		/// \sa get_ScrollBars, put_ScrollBars
		ScrollBarsConstants scrollBars;
		/// \brief <em>Holds the \c ScrollToTopOnKillFocus property's setting</em>
		///
		/// \sa get_ScrollToTopOnKillFocus, put_ScrollToTopOnKillFocus
		UINT scrollToTopOnKillFocus : 1;
		/// \brief <em>Holds the \c ShowSelectionBar property's setting</em>
		///
		/// \sa get_ShowSelectionBar, put_ShowSelectionBar
		UINT showSelectionBar : 1;
		/// \brief <em>Holds the \c SmartSpacingOnDrop property's setting</em>
		///
		/// \sa get_SmartSpacingOnDrop, put_SmartSpacingOnDrop
		UINT smartSpacingOnDrop : 1;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c SwitchFontOnIMEInput property's setting</em>
		///
		/// \sa get_SwitchFontOnIMEInput, put_SwitchFontOnIMEInput
		UINT switchFontOnIMEInput : 1;
		/// \brief <em>If \c TRUE, \c Create will set \c text to the control's name</em>
		///
		/// \sa Create, text
		UINT resetTextToName : 1;
		/// \brief <em>Holds the \c Text property's setting</em>
		///
		/// \sa get_Text, put_Text
		CComBSTR text;
		/// \brief <em>Holds the \c TextFlow property's setting</em>
		///
		/// \sa get_TextFlow, put_TextFlow
		TextFlowConstants textFlow;
		/// \brief <em>Holds the \c TextOrientation property's setting</em>
		///
		/// \sa get_TextOrientation, put_TextOrientation
		TextOrientationConstants textOrientation;
		/// \brief <em>Holds the \c TSFModeBias property's setting</em>
		///
		/// \sa get_TSFModeBias, put_TSFModeBias
		TSFModeBiasConstants TSFModeBias;
		/// \brief <em>Holds the \c UseBkAcetateColorForTextSelection property's setting</em>
		///
		/// \sa get_UseBkAcetateColorForTextSelection, put_UseBkAcetateColorForTextSelection
		UINT useBkAcetateColorForTextSelection : 1;
		/// \brief <em>Holds the \c UseBuiltInSpellChecking property's setting</em>
		///
		/// \sa get_UseBuiltInSpellChecking, put_UseBuiltInSpellChecking
		UINT useBuiltInSpellChecking : 1;
		/// \brief <em>Holds the \c UseCustomFormattingRectangle property's setting</em>
		///
		/// \sa get_UseCustomFormattingRectangle, put_UseCustomFormattingRectangle
		UINT useCustomFormattingRectangle : 1;
		/// \brief <em>Holds the \c UseDualFontMode property's setting</em>
		///
		/// \sa get_UseDualFontMode, put_UseDualFontMode
		UINT useDualFontMode : 1;
		/// \brief <em>Holds the \c UseSmallerFontForNestedFractions property's setting</em>
		///
		/// \sa get_UseSmallerFontForNestedFractions, put_UseSmallerFontForNestedFractions
		UINT useSmallerFontForNestedFractions : 1;
		/// \brief <em>Holds the \c UseTextServicesFramework property's setting</em>
		///
		/// \sa get_UseTextServicesFramework, put_UseTextServicesFramework
		UINT useTextServicesFramework : 1;
		// \brief <em>Holds the \c UseTouchKeyboardAutoCorrection property's setting</em>
		//
		// \sa get_UseTouchKeyboardAutoCorrection, put_UseTouchKeyboardAutoCorrection
		//UINT useTouchKeyboardAutoCorrection : 1;
		/// \brief <em>Holds the \c UseWindowsThemes property's setting</em>
		///
		/// \sa get_UseWindowsThemes, put_UseWindowsThemes
		UINT useWindowsThemes : 1;
		/// \brief <em>Holds the \c ZoomRatio property's setting</em>
		///
		/// \sa get_ZoomRatio, put_ZoomRatio
		DOUBLE zoomRatio;

		/// \brief <em>Holds the control's \c ITextDocument instance</em>
		///
		/// \sa OnCreate,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb774052.aspx">ITextDocument</a>
		ITextDocument* pTextDocument;

		Properties()
		{
			pTextDocument = NULL;
			ResetToDefaults();
		}

		~Properties()
		{
			Release();
		}

		/// \brief <em>Prepares the object for destruction</em>
		void Release(void)
		{
			font.Release();
			linkMouseIcon.Release();
			mouseIcon.Release();
			if(pTextDocument) {
				pTextDocument->Release();
				pTextDocument = NULL;
			}
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			acceptTabKey = TRUE;
			adjustLineHeightForEastAsianLanguages = TRUE;
			allowInputThroughTSF = TRUE;
			allowMathZoneInsertion = FALSE;
			allowObjectInsertionThroughTSF = FALSE;
			allowTableInsertion = TRUE;
			allowTSFProofingTips = FALSE;
			allowTSFSmartTags = FALSE;
			alwaysShowScrollBars = FALSE;
			alwaysShowSelection = FALSE;
			appearance = a3D;
			autoDetectedURLSchemes = L"";
			autoDetectURLs = static_cast<AutoDetectURLsConstants>(aduURLs | aduEMailAddresses);
			autoScrolling = static_cast<AutoScrollingConstants>(asHorizontal | asVertical);
			autoSelectWordOnTrackSelection = FALSE;
			backColor = 0x80000000 | COLOR_WINDOW;
			beepOnMaxText = FALSE;
			borderStyle = bsNone;
			characterConversion = ccNone;
			defaultMathZoneHAlignment = halCenter;
			defaultTabWidth = 36.0f;
			denoteEmptyMathArguments = demaMandatoryOnly;
			detectDragDrop = FALSE;
			disabledEvents = static_cast<DisabledEventsConstants>(deClickEvents | deMouseEvents | deScrolling | deShouldResizeControlWindow | deSelectionChangingEvents);
			disableIMEOperations = FALSE;
			//discardCompositionStringOnIMECancel = FALSE;
			displayHyperlinkTooltips = TRUE;
			displayZeroWidthTableCellBorders = TRUE;
			draftMode = FALSE;
			dragScrollTimeBase = -1;
			dropWordsOnWordBoundariesOnly = TRUE;
			emulateSimpleTextBox = FALSE;
			enabled = TRUE;
			extendFontBackColorToWholeLine = FALSE;
			ZeroMemory(&formattingRectangle, sizeof(formattingRectangle));
			growNAryOperators = TRUE;
			//hAlignment = halLeft;
			hoverTime = -1;

			hyphenationFunction = NULL;
			hyphenationWordWrapZoneWidth = 0;
			HDC hDC = ::GetDC(NULL);
			if(hDC) {
				// 12 Pixel * x Twips per Pixel
				hyphenationWordWrapZoneWidth = static_cast<short>(12 * 1440 / GetDeviceCaps(hDC, LOGPIXELSX));
				::ReleaseDC(NULL, hDC);
			}

			IMEConversionMode = imecmDefault;
			IMEMode = imeInherit;
			integralLimitsLocation = llOnSide;
			leftMargin = -1;
			letClientHandleAllIMEOperations = FALSE;
			linkMousePointer = mpDefault;
			logicalCaret = FALSE;
			mathLineBreakBehavior = mlbbBreakBeforeOperator;
			maxTextLength = -1;
			maxUndoQueueSize = 100;
			modified = FALSE;
			mousePointer = mpDefault;
			multiLine = TRUE;
			multiSelect = FALSE;
			nAryLimitsLocation = llUnderAndOver;
			noInputSequenceCheck = FALSE;
			oleDragImageStyle = odistClassic;
			processContextMenuKeys = TRUE;
			rawSubScriptAndSuperScriptOperators = TRUE;
			readOnly = FALSE;
			registerForOLEDragDrop = rfoddNativeDragDrop;
			richEditAPIVersion = reavUnknown;
			richEditLibraryPath = L"";
			richEditVersion = revAlwaysUseWindowsRichEdit;
			rightMargin = -1;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			rightToLeftMathZones = TRUE;
			scrollBars = static_cast<ScrollBarsConstants>(sbVertical | sbHorizontal);
			scrollToTopOnKillFocus = FALSE;
			showSelectionBar = FALSE;
			smartSpacingOnDrop = TRUE;
			supportOLEDragImages = TRUE;
			switchFontOnIMEInput = TRUE;
			text = L"RichTextBox";
			resetTextToName = TRUE;
			textFlow = tfLeftToRightTopToBottom;
			textOrientation = toHorizontal;
			TSFModeBias = tsfmbDefault;
			useBkAcetateColorForTextSelection = TRUE;
			useBuiltInSpellChecking = FALSE;
			useCustomFormattingRectangle = FALSE;
			useDualFontMode = TRUE;
			useSmallerFontForNestedFractions = FALSE;
			useTextServicesFramework = TRUE;
			//useTouchKeyboardAutoCorrection = FALSE;
			useWindowsThemes = TRUE;
			zoomRatio = 0.0;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, \c RecreateControlWindow won't recreate the control window</em>
		///
		/// \sa RecreateControlWindow, Load
		UINT dontRecreate : 1;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;
		/// \brief <em>If not 0, we won't raise \c SelectionChanged events</em>
		///
		/// \sa EnterSilentSelectionChangesSection, LeaveSilentSelectionChangesSection, HitTest,
		///     Raise_SelectionChanged
		UINT silentSelectionChanges;
		/// \brief <em>If not 0, we won't raise \c RemovingOLEObject events</em>
		///
		/// \sa EnterSilentObjectDeletionsSection, LeaveSilentObjectDeletionsSection, Raise_RemovingOLEObject
		UINT silentObjectDeletions;
		/// \brief <em>If not 0, we won't raise \c InsertingOLEObject events</em>
		///
		/// \sa EnterSilentObjectInsertionsSection, LeaveSilentObjectInsertionsSection, Raise_InsertingOLEObject
		UINT silentObjectInsertions;

		Flags()
		{
			dontRecreate = FALSE;
			usingThemes = FALSE;
			silentSelectionChanges = 0;
			silentObjectDeletions = 0;
			silentObjectInsertions = 0;
		}
	} flags;

	/// \brief <em>Holds the previous text range that the \c EN_SELCHANGE notification has been sent for</em>
	///
	/// \sa OnSelChangeNotification
	CHARRANGE previousSelectedRange;
	/// \brief <em>A storage used to provide memory for loading embedded objects</em>
	///
	/// \sa OnCreate, OnDestroy, GetNewStorage
	LPSTORAGE pDocumentStorage;
	/// \brief <em>Holds the index of the next embedded object</em>
	///
	/// \sa OnCreate, GetNewStorage
	int nextEmbeddedObjectIndex;

	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
	protected:
		/// \brief <em>Holds all mouse buttons that may cause a click event in the immediate future</em>
		///
		/// A bit field of \c SHORT values representing those mouse buttons that are currently pressed and
		/// may cause a click event in the immediate future.
		///
		/// \sa StoreClickCandidate, IsClickCandidate, RemoveClickCandidate, Raise_Click, Raise_MClick,
		///     Raise_RClick, Raise_XClick
		SHORT clickCandidates;

	public:
		/// \brief <em>If \c TRUE, the \c MouseEnter event already was raised</em>
		///
		/// \sa Raise_MouseEnter
		UINT enteredControl : 1;
		/// \brief <em>If \c TRUE, the \c MouseHover event already was raised</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa Raise_MouseHover
		UINT hoveredControl : 1;
		/// \brief <em>Holds the mouse cursor's last position</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		POINT lastPosition;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseHover event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl, LeaveControl
		void HoverControl(void)
		{
			enteredControl = TRUE;
			hoveredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseLeave event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl
		void LeaveControl(void)
		{
			enteredControl = FALSE;
			hoveredControl = FALSE;
		}

		/// \brief <em>Stores a mouse button as click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa clickCandidates, IsClickCandidate, RemoveClickCandidate
		void StoreClickCandidate(SHORT button)
		{
			// avoid combined click events
			if(clickCandidates == 0) {
				clickCandidates |= button;
			}
		}

		/// \brief <em>Retrieves whether a mouse button is a click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a click candidate; otherwise \c FALSE.
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa clickCandidates, StoreClickCandidate, RemoveClickCandidate
		BOOL IsClickCandidate(SHORT button)
		{
			return (clickCandidates & button);
		}

		/// \brief <em>Removes a mouse button from the list of click candidates</em>
		///
		/// \param[in] button The mouse button to remove.
		///
		/// \sa clickCandidates, RemoveAllClickCandidates, StoreClickCandidate, IsClickCandidate
		void RemoveClickCandidate(SHORT button)
		{
			clickCandidates &= ~button;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa clickCandidates, RemoveClickCandidate, StoreClickCandidate, IsClickCandidate
		void RemoveAllClickCandidates(void)
		{
			clickCandidates = 0;
		}
	} MouseStatus;

	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus;

	//////////////////////////////////////////////////////////////////////
	/// \name Drag'n'Drop
	///
	//@{
	/// \brief <em>The \c CLSID_WICImagingFactory object used to create WIC objects that are required during drag image creation</em>
	///
	/// \sa OnGetDragImage, CreateThumbnail
	CComPtr<IWICImagingFactory> pWICImagingFactory;
	/// \brief <em>Creates a thumbnail of the specified icon in the specified size</em>
	///
	/// \param[in] hIcon The icon to create the thumbnail for.
	/// \param[in] size The thumbnail's size in pixels.
	/// \param[in,out] pBits The thumbnail's DIB bits.
	/// \param[in] doAlphaChannelPostProcessing WIC has problems to handle the alpha channel of the icon
	///            specified by \c hIcon. If this parameter is set to \c TRUE, some post-processing is done
	///            to correct the pixel failures. Otherwise the failures are not corrected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnGetDragImage, pWICImagingFactory
	HRESULT CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing);

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>The dragged text</em>
		CComPtr<ITextRange> pDraggedTextRange;
		/// \brief <em>The handle of the imagelist containing the drag image</em>
		///
		/// \sa get_hDragImageList
		HIMAGELIST hDragImageList;
		/// \brief <em>Enables or disables auto-destruction of \c hDragImageList</em>
		///
		/// Controls whether the imagelist defined by \c hDragImageList is destroyed automatically. If set to
		/// \c TRUE, it is destroyed in \c EndDrag; otherwise not.
		///
		/// \sa hDragImageList, EndDrag
		UINT autoDestroyImgLst : 1;
		/// \brief <em>Indicates whether the drag image is visible or hidden</em>
		///
		/// If this value is 0, the drag image is visible; otherwise not.
		///
		/// \sa get_hDragImageList, get_ShowDragImage, put_ShowDragImage, ShowDragImage, HideDragImage,
		///     IsDragImageVisible
		int dragImageIsHidden;

		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>The currently dragged data for the case that the we're the drag source</em>
		CComPtr<IDataObject> pSourceDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the mouse button (as \c MK_* constant) that the drag'n'drop operation is performed with</em>
		DWORD draggingMouseButton;
		/// \brief <em>If \c TRUE, we'll hide and re-show the drag image in \c IDropTarget::DragEnter so that the item count label is displayed</em>
		///
		/// \sa DragEnter, OLEDrag
		UINT useItemCountLabelHack : 1;
		/// \brief <em>Holds the \c IDataObject to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		IDataObject* drop_pDataObject;
		/// \brief <em>Holds the mouse position to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		POINT drop_mousePosition;
		/// \brief <em>Holds the drop effect to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		DWORD drop_effect;
		//@}
		//////////////////////////////////////////////////////////////////////

		/// \brief <em>Holds data and flags related to auto-scrolling</em>
		///
		/// \sa AutoScroll
		struct AutoScrolling
		{
			/// \brief <em>Holds the current speed multiplier used for horizontal auto-scrolling</em>
			LONG currentHScrollVelocity;
			/// \brief <em>Holds the current speed multiplier used for vertical auto-scrolling</em>
			LONG currentVScrollVelocity;
			/// \brief <em>Holds the current interval of the auto-scroll timer</em>
			LONG currentTimerInterval;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled downwards</em>
			DWORD lastScroll_Down;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the left</em>
			DWORD lastScroll_Left;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the right</em>
			DWORD lastScroll_Right;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled upwardly</em>
			DWORD lastScroll_Up;

			AutoScrolling()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				currentHScrollVelocity = 0;
				currentVScrollVelocity = 0;
				currentTimerInterval = 0;
				lastScroll_Down = 0;
				lastScroll_Left = 0;
				lastScroll_Right = 0;
				lastScroll_Up = 0;
			}
		} autoScrolling;

		/// \brief <em>Holds data required for drag-detection</em>
		struct Candidate
		{
			/// \brief <em>The range of text, that might be dragged</em>
			CComPtr<ITextRange> pTextRange;
			/// \brief <em>The position in pixels at which dragging might start</em>
			POINT position;
			/// \brief <em>The mouse button (as \c MK_* constant) that the drag'n'drop operation might be performed with</em>
			int button;

			/// \brief <em>Holds the parameters of a buffered \c WM_*BUTTONDOWN message</em>
			typedef struct BufferedMessage
			{
				/// \brief <em>Holds the \c wParam parameter of the buffered message</em>
				WPARAM wParam;
				/// \brief <em>Holds the \c lParam parameter of the buffered message</em>
				LPARAM lParam;

				BufferedMessage()
				{
					Reset();
				}

				/// \brief <em>Resets all member variables to their defaults</em>
				void Reset(void)
				{
					wParam = 0;
					lParam = 0;
				}
			} BufferedMessage;
			/// \brief <em>Holds the parameters of a buffered \c WM_LBUTTONDOWN message</em>
			BufferedMessage bufferedLButtonDown;
			/// \brief <em>Holds the parameters of a buffered \c WM_RBUTTONDOWN message</em>
			BufferedMessage bufferedRButtonDown;

			Candidate()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				button = 0;
				pTextRange = NULL;
				bufferedLButtonDown.Reset();
				bufferedRButtonDown.Reset();
			}
		} candidate;

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pSourceDataObject = NULL;
			pDropTargetHelper = NULL;
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;
			drop_pDataObject = NULL;

			pDraggedTextRange = NULL;
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
			}
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			this->pDraggedTextRange = NULL;
			if(this->hDragImageList && autoDestroyImgLst) {
				ImageList_Destroy(this->hDragImageList);
			}
			this->hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;

			candidate.Reset();

			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			if(this->pSourceDataObject) {
				this->pSourceDataObject = NULL;
			}
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;
			drop_pDataObject = NULL;
		}

		/// \brief <em>Decrements the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, HideDragImage, IsDragImageVisible
		void ShowDragImage(BOOL commonDragDropOnly)
		{
			if(this->hDragImageList) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					ImageList_DragShowNolock(TRUE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					pDropTargetHelper->Show(TRUE);
				}
			}
		}

		/// \brief <em>Increments the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, ShowDragImage, IsDragImageVisible
		void HideDragImage(BOOL commonDragDropOnly)
		{
			if(this->hDragImageList) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					ImageList_DragShowNolock(FALSE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					pDropTargetHelper->Show(FALSE);
				}
			}
		}

		/// \brief <em>Retrieves whether we're currently displaying a drag image</em>
		///
		/// \return \c TRUE, if we're displaying a drag image; otherwise \c FALSE.
		///
		/// \sa dragImageIsHidden, ShowDragImage, HideDragImage
		BOOL IsDragImageVisible(void)
		{
			return (dragImageIsHidden == 0);
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation started</em>
		///
		/// \param[in] pRTB The \c RichTextBox control whose \c CreateLegacyDragImage method shall be called.
		/// \param[in] hWndRTB The edit control window, that the method will work on to calculate the position
		///            of the drag image's hotspot.
		/// \param[in] pDraggedTxtRange The range of text to drag.
		/// \param[in] hDragImgLst The imagelist containing the drag image that shall be used to
		///            visualize the drag'n'drop operation. If -1, the method will create the drag image
		///            itself; if \c NULL, no drag image will be displayed.
		/// \param[in,out] pXHotSpot The x-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImgLst parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImgLst parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		/// \param[in,out] pYHotSpot The y-coordinate (in pixels) of the drag image's hotspot relative to the
		///                drag image's upper-left corner. If the \c hDragImgLst parameter is set to
		///                \c NULL, this parameter is ignored. If the \c hDragImgLst parameter is set to
		///                -1, this parameter is set to the hotspot calculated by the method.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa EndDrag, TextBox::CreateLegacyDragImage
		HRESULT BeginDrag(RichTextBox* pRTB, HWND hWndRTB, ITextRange* pDraggedTxtRange, HIMAGELIST hDragImgLst, int* pXHotSpot, int* pYHotSpot)
		{
			ATLASSUME(pRTB);

			UINT b = FALSE;
			if(hDragImgLst == static_cast<HIMAGELIST>(LongToHandle(-1))) {
				POINT upperLeftPoint = {0};
				hDragImgLst = pRTB->CreateLegacyDragImage(pDraggedTxtRange, &upperLeftPoint, NULL);
				if(!hDragImgLst) {
					return E_FAIL;
				}
				b = TRUE;

				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				::ScreenToClient(hWndRTB, &mousePosition);
				if(CWindow(hWndRTB).GetExStyle() & WS_EX_LAYOUTRTL) {
					SIZE dragImageSize = {0};
					ImageList_GetIconSize(hDragImgLst, reinterpret_cast<PINT>(&dragImageSize.cx), reinterpret_cast<PINT>(&dragImageSize.cy));
					*pXHotSpot = upperLeftPoint.x + dragImageSize.cx - mousePosition.x;
				} else {
					*pXHotSpot = mousePosition.x - upperLeftPoint.x;
				}
				*pYHotSpot = mousePosition.y - upperLeftPoint.y;
			}

			if(this->hDragImageList && this->autoDestroyImgLst) {
				ImageList_Destroy(this->hDragImageList);
			}

			this->autoDestroyImgLst = b;
			this->hDragImageList = hDragImgLst;
			this->pDraggedTextRange = pDraggedTxtRange;

			dragImageIsHidden = 1;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done after a drag'n'drop operation ended</em>
		///
		/// \sa BeginDrag
		void EndDrag(void)
		{
			this->pDraggedTextRange = NULL;
			if(autoDestroyImgLst && this->hDragImageList) {
				ImageList_Destroy(this->hDragImageList);
			}
			this->hDragImageList = NULL;
			dragImageIsHidden = 1;
			autoScrolling.Reset();
		}

		/// \brief <em>Retrieves whether we're in drag'n'drop mode</em>
		///
		/// \return \c TRUE if we're in drag'n'drop mode; otherwise \c FALSE.
		///
		/// \sa BeginDrag, EndDrag
		BOOL IsDragging(void)
		{
			return (this->pDraggedTextRange != NULL);
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			autoScrolling.Reset();
		}
	} dragDropStatus;
	///@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to auto-scroll the control window during drag'n'drop</em>
		static const UINT_PTR ID_DRAGSCROLL = 13;
	} timers;

	/// \brief <em>Holds the token that is used to shutdown GDI+</em>
	ULONG_PTR gdiplusToken;

private:
};     // RichTextBox

OBJECT_ENTRY_AUTO(__uuidof(RichTextBox), RichTextBox)