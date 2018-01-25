//////////////////////////////////////////////////////////////////////
/// \class TextFont
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a text range's styling, e.g. font and colors</em>
///
/// This class is a wrapper around the styling of a text range.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTextFont, RichTextBox, TextRange
/// \else
///   \sa RichTextBoxCtlLibA::IRichTextFont, RichTextBox, TextRange
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITextFontEvents_CP.h"
#include "helpers.h"


class ATL_NO_VTABLE TextFont : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TextFont, &CLSID_TextFont>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TextFont>,
    public Proxy_IRichTextFontEvents<TextFont>, 
    #ifdef UNICODE
    	public IDispatchImpl<IRichTextFont, &IID_IRichTextFont, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTextFont, &IID_IRichTextFont, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TEXTFONT)

		BEGIN_COM_MAP(TextFont)
			COM_INTERFACE_ENTRY(IRichTextFont)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextFont), 0, QueryITextFontInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextFont2), 0, QueryITextFontInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TextFont)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTextFontEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
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

	/// \brief <em>Will be called if this class is \c QueryInterface()'d for \c ITextFont or \c ITextFont2</em>
	///
	/// This method will be called if the class's \c QueryInterface() method is called with
	/// \c IID_ITextFont or \c IID_ITextFont2. We forward the call to the wrapped \c ITextFont
	/// implementation.
	///
	/// \param[in] pThis The instance of this class, that the interface is queried from.
	/// \param[in] queriedInterface Should be \c IID_ITextFont or \c IID_ITextFont2.
	/// \param[out] ppImplementation Receives the wrapped object's \c ITextFont or \c ITextFont2
	///             implementation.
	/// \param[in] cookie A \c DWORD value specified in the COM interface map.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774054.aspx">ITextFont</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/hh768487.aspx">ITextFont2</a>
	static HRESULT CALLBACK QueryITextFontInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTextFont
	///
	//@{
	// \brief <em>Retrieves the current setting of the \c AlignEquationsAtOperator property</em>
	//
	// To be documented
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \if UNICODE
	//   \sa put_AlignEquationsAtOperator, get_IsMathZone, RichTextBox::get_DefaultMathZoneHAlignment
	// \else
	//   \sa put_AlignEquationsAtOperator, get_IsMathZone, RichTextBox::get_DefaultMathZoneHAlignment
	// \endif
	//virtual HRESULT STDMETHODCALLTYPE get_AlignEquationsAtOperator(SHORT* pValue);
	// \brief <em>Sets the \c AlignEquationsAtOperator property</em>
	//
	// To be documented
	//
	// \param[in] newValue The setting to apply.
	//
	// \return An \c HRESULT error code.
	//
	// \if UNICODE
	//   \sa get_AlignEquationsAtOperator, put_IsMathZone, RichTextBox::put_DefaultMathZoneHAlignment
	// \else
	//   \sa get_AlignEquationsAtOperator, put_IsMathZone, RichTextBox::put_DefaultMathZoneHAlignment
	// \endif
	//virtual HRESULT STDMETHODCALLTYPE put_AlignEquationsAtOperator(SHORT newValue);
	/// \brief <em>Retrieves the current setting of the \c AllCaps property</em>
	///
	/// Retrieves whether the characters are all uppercase. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are all
	/// uppercase; if set to \c bpvFalse, they are not all uppercase; and if set to \c bpvUndefined, some
	/// of the characters are uppercase and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_AllCaps, get_SmallCaps, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_AllCaps, get_SmallCaps, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AllCaps(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c AllCaps property</em>
	///
	/// Sets whether the characters are all uppercase. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are all
	/// uppercase; if set to \c bpvFalse, they are not all uppercase; and if set to \c bpvUndefined, some
	/// of the characters are uppercase and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_AllCaps, put_SmallCaps, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_AllCaps, put_SmallCaps, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_AllCaps(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c AnimationType property</em>
	///
	/// Retrieves whether and how the text is animated. Any of the values defined by the
	/// \c AnimationTypeConstants enumeration is valid. If the text range uses multiple animation
	/// types, the type will be reported as \c atUndefined.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_AnimationType, RichTextBoxCtlLibU::AnimationTypeConstants
	/// \else
	///   \sa put_AnimationType, RichTextBoxCtlLibA::AnimationTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AnimationType(AnimationTypeConstants* pValue);
	/// \brief <em>Sets the \c AnimationType property</em>
	///
	/// Sets whether and how the text is animated. Any of the values defined by the
	/// \c AnimationTypeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_AnimationType, RichTextBoxCtlLibU::AnimationTypeConstants
	/// \else
	///   \sa get_AnimationType, RichTextBoxCtlLibA::AnimationTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_AnimationType(AnimationTypeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c AutoLigatures property</em>
	///
	/// Retrieves whether character combinations like \c f+i are detected and replaced by ligatures like
	/// \c ﬁ automatically. Any of the values defined by the \c BooleanPropertyValueConstants enumeration
	/// is valid. If set to \c bpvTrue, ligatures are detected and used automatically; if set to \c bpvFalse,
	/// they are not detected and used automatically; and if set to \c bpvUndefined, some parts of the text
	/// range use automatic ligature detection and others don't.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_AutoLigatures, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_AutoLigatures, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_AutoLigatures(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c AutoLigatures property</em>
	///
	/// Sets whether character combinations like \c f+i are detected and replaced by ligatures like
	/// \c ﬁ automatically. Any of the values defined by the \c BooleanPropertyValueConstants enumeration
	/// is valid. If set to \c bpvTrue, ligatures are detected and used automatically; if set to \c bpvFalse,
	/// they are not detected and used automatically; and if set to \c bpvUndefined, some parts of the text
	/// range use automatic ligature detection and others don't.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_AutoLigatures, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_AutoLigatures, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_AutoLigatures(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the text range's background color. If set to -1, the default background color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BackColor, get_ForeColor, RichTextBox::get_ExtendFontBackColorToWholeLine,
	///     RichTextBox::get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the text range's background color. If set to -1, the default background color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BackColor, put_ForeColor, RichTextBox::put_ExtendFontBackColorToWholeLine,
	///     RichTextBox::put_BackColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c Bold property</em>
	///
	/// Retrieves whether the characters are bold. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// bold; if set to \c bpvFalse, they are not bold; and if set to \c bpvUndefined, some of the
	/// characters are bold and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Bold, get_Weight, get_Italic, get_Name, get_Size,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Bold, get_Weight, get_Italic, get_Name, get_Size,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Bold(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Bold property</em>
	///
	/// Sets whether the characters are bold. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// bold; if set to \c bpvFalse, they are not bold; and if set to \c bpvUndefined, some of the
	/// characters are bold and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Bold, put_Weight, put_Italic, put_Name, put_Size,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Bold, put_Weight, put_Italic, put_Name, put_Size,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Bold(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c CanChange property</em>
	///
	/// Retrieves whether the styling of the text range can be changed. If set to \c bpvTrue, the styling can
	/// be changed; if set to \c bpvFalse, it cannot be changed; and if set to \c bpvUndefined, parts of the
	/// text range's styling can be changed and others can't.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CanChange(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayAsOrdinaryTextWithMathZone property</em>
	///
	/// Retrieves whether the text, which is located inside a math zone, is displayed as normal text. Any of
	/// the values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to
	/// \c bpvTrue, the text is displayed as ordinary text; if set to \c bpvFalse, they it is displayed as
	/// part of the mathematics; and if set to \c bpvUndefined, parts of the text range are displayed as
	/// ordinary text and others as part of the mathematics.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_DisplayAsOrdinaryTextWithMathZone, get_IsMathZone,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_DisplayAsOrdinaryTextWithMathZone, get_IsMathZone,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisplayAsOrdinaryTextWithMathZone(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c DisplayAsOrdinaryTextWithMathZone property</em>
	///
	/// Sets whether the text, which is located inside a math zone, is displayed as normal text. Any of
	/// the values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to
	/// \c bpvTrue, the text is displayed as ordinary text; if set to \c bpvFalse, they it is displayed as
	/// part of the mathematics; and if set to \c bpvUndefined, parts of the text range are displayed as
	/// ordinary text and others as part of the mathematics.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_DisplayAsOrdinaryTextWithMathZone, put_IsMathZone,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_DisplayAsOrdinaryTextWithMathZone, put_IsMathZone,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisplayAsOrdinaryTextWithMathZone(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Emboss property</em>
	///
	/// Retrieves whether the characters are embossed. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// embossed; if set to \c bpvFalse, they are not embossed; and if set to \c bpvUndefined, some of the
	/// characters are embossed and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Emboss, get_Engrave, get_Outline, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Emboss, get_Engrave, get_Outline, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Emboss(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Emboss property</em>
	///
	/// Sets whether the characters are embossed. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// embossed; if set to \c bpvFalse, they are not embossed; and if set to \c bpvUndefined, some of the
	/// characters are embossed and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Emboss, put_Engrave, put_Outline, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Emboss, put_Engrave, put_Outline, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Emboss(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Engrave property</em>
	///
	/// Retrieves whether the characters are displayed as imprinted characters. Any of the values
	/// defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the
	/// characters are displayed as imprinted characters; if set to \c bpvFalse, they are not displayed as
	/// imprinted characters; and if set to \c bpvUndefined, some of the characters are displayed as
	/// imprinted characters and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Engrave, get_Emboss, get_Outline, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Engrave, get_Emboss, get_Outline, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Engrave(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Engrave property</em>
	///
	/// Sets whether the characters are displayed as imprinted characters. Any of the values
	/// defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the
	/// characters are displayed as imprinted characters; if set to \c bpvFalse, they are not displayed as
	/// imprinted characters; and if set to \c bpvUndefined, some of the characters are displayed as
	/// imprinted characters and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Engrave, put_Emboss, put_Outline, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Engrave, put_Emboss, put_Outline, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Engrave(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the text range's text color. If set to -1, the default text color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ForeColor, get_BackColor, get_Hidden
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the text range's text color. If set to -1, the default text color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ForeColor, put_BackColor, put_Hidden
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c Hidden property</em>
	///
	/// Retrieves whether the characters are hidden. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// hidden; if set to \c bpvFalse, they are not hidden; and if set to \c bpvUndefined, some of the
	/// characters are hidden and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Hidden, get_ForeColor, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Hidden, get_ForeColor, ichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Hidden(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Hidden property</em>
	///
	/// Sets whether the characters are hidden. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// hidden; if set to \c bpvFalse, they are not hidden; and if set to \c bpvUndefined, some of the
	/// characters are hidden and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Hidden, put_ForeColor, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Hidden, put_ForeColor, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Hidden(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c HorizontalScaling property</em>
	///
	/// Retrieves the percentage by which the width of each character in the range is scaled.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.
	///
	/// \sa put_HorizontalScaling, get_Size, get_Spacing
	virtual HRESULT STDMETHODCALLTYPE get_HorizontalScaling(LONG* pValue);
	/// \brief <em>Sets the \c HorizontalScaling property</em>
	///
	/// Sets the percentage by which the width of each character in the range is scaled.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.
	///
	/// \sa get_HorizontalScaling, put_Size, put_Spacing
	virtual HRESULT STDMETHODCALLTYPE put_HorizontalScaling(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c IsAutomaticLink property</em>
	///
	/// Retrieves whether the text range is a hyperlink, that has been detected automatically. Any of the
	/// values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue,
	/// the text range is an automatically detected hyperlink; if set to \c bpvFalse, it is not a hyperlink
	/// or has not been detected automatically; and if set to \c bpvUndefined, some parts of the text range
	/// are an automatically detected hyperlink and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_IsLink, TextRange::get_URL, RichTextBox::get_AutoDetectURLs,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_IsLink, TextRange::get_URL, RichTextBox::get_AutoDetectURLs,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IsAutomaticLink(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c IsInlineMathZone property</em>
	///
	/// Retrieves whether the text range is a zone for entering formulas, and is embedded in a normal line of
	/// text. Any of the values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set
	/// to \c bpvTrue, the text range is an inline math zone; if set to \c bpvFalse, it is not an inline math
	/// zone (might be a math zone that has its own paragraph); and if set to \c bpvUndefined, some parts of
	/// the text range are an inline math zone and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_IsMathZone, TextRange::BuildUpMath, RichTextBox::get_AllowMathZoneInsertion,
	///       get_DisplayAsOrdinaryTextWithMathZone, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_IsMathZone, TextRange::BuildUpMath, RichTextBox::get_AllowMathZoneInsertion,
	///       get_DisplayAsOrdinaryTextWithMathZone, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IsInlineMathZone(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c IsLink property</em>
	///
	/// Retrieves whether the text range is a hyperlink, either detected automatically or created manually.
	/// Any of the values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to
	/// \c bpvTrue, the text range is a hyperlink; if set to \c bpvFalse, it is not a hyperlink; and if set
	/// to \c bpvUndefined, some parts of the text range are a hyperlink and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_IsAutomaticLink, TextRange::get_URL, RichTextBox::get_AutoDetectURLs,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_IsAutomaticLink, TextRange::get_URL, RichTextBox::get_AutoDetectURLs,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IsLink(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c IsMathZone property</em>
	///
	/// Retrieves whether the text range is a zone for entering formulas. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the text range is a
	/// math zone; if set to \c bpvFalse, it is not a math zone; and if set to \c bpvUndefined, some parts of
	/// the text range are a math zone and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_IsMathZone, get_IsInlineMathZone, TextRange::BuildUpMath,
	///       RichTextBox::get_AllowMathZoneInsertion, get_DisplayAsOrdinaryTextWithMathZone,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_IsMathZone, get_IsInlineMathZone, TextRange::BuildUpMath,
	///       RichTextBox::get_AllowMathZoneInsertion, get_DisplayAsOrdinaryTextWithMathZone,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IsMathZone(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c IsMathZone property</em>
	///
	/// Sets whether the text range is a zone for entering formulas. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the text range is a
	/// math zone; if set to \c bpvFalse, it is not a math zone; and if set to \c bpvUndefined, some parts of
	/// the text range are a math zone and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_IsMathZone, TextRange::BuildUpMath, RichTextBox::put_AllowMathZoneInsertion,
	///       put_DisplayAsOrdinaryTextWithMathZone, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_IsMathZone, TextRange::BuildUpMath, RichTextBox::put_AllowMathZoneInsertion,
	///       put_DisplayAsOrdinaryTextWithMathZone, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IsMathZone(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Italic property</em>
	///
	/// Retrieves whether the characters are in italics. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are in
	/// italics; if set to \c bpvFalse, they are not in italics; and if set to \c bpvUndefined, some of the
	/// characters are in italics and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Italic, get_Bold, get_Name, get_Size, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Italic, get_Bold, get_Name, get_Size, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Italic(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Italic property</em>
	///
	/// Sets whether the characters are in italics. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are in
	/// italics; if set to \c bpvFalse, they are not in italics; and if set to \c bpvUndefined, some of the
	/// characters are in italics and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Italic, put_Bold, put_Name, put_Size, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Italic, put_Bold, put_Name, put_Size, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Italic(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c KerningSize property</em>
	///
	/// Retrieves the minimum font size at which kerning is applied. If the text range uses
	/// multiple kerning sizes, the size will be reported as -9999999.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_KerningSize, get_Size, get_Spacing
	virtual HRESULT STDMETHODCALLTYPE get_KerningSize(FLOAT* pValue);
	/// \brief <em>Sets the \c KerningSize property</em>
	///
	/// Sets the minimum font size at which kerning is applied.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_KerningSize, put_Size, put_Spacing
	virtual HRESULT STDMETHODCALLTYPE put_KerningSize(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c Locale property</em>
	///
	/// Retrieves the identifier of the locale that the text is in.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Locale, RichTextBox::get_UseBuiltInSpellChecking
	virtual HRESULT STDMETHODCALLTYPE get_Locale(LONG* pValue);
	/// \brief <em>Sets the \c Locale property</em>
	///
	/// Sets the identifier of the locale that the text is in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Locale, RichTextBox::put_UseBuiltInSpellChecking
	virtual HRESULT STDMETHODCALLTYPE put_Locale(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Name property</em>
	///
	/// Retrieves the name of the font the text is drawn in. If the text range uses multiple fonts,
	/// the name will be reported as an empty string.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Name, get_Size, RichTextBox::get_Font
	virtual HRESULT STDMETHODCALLTYPE get_Name(BSTR* pValue);
	/// \brief <em>Sets the \c Name property</em>
	///
	/// Sets the name of the font the text is drawn in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Name, put_Size, RichTextBox::putref_Font
	virtual HRESULT STDMETHODCALLTYPE put_Name(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c Outline property</em>
	///
	/// Retrieves whether the characters are displayed as outlined characters. Any of the values
	/// defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the
	/// characters are displayed as outlined characters; if set to \c bpvFalse, they are not displayed as
	/// outlined characters; and if set to \c bpvUndefined, some of the characters are displayed as
	/// outlined characters and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Outline, get_Emboss, get_Engrave, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Outline, get_Emboss, get_Engrave, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Outline(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Outline property</em>
	///
	/// Sets whether the characters are displayed as outlined characters. Any of the values
	/// defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the
	/// characters are displayed as outlined characters; if set to \c bpvFalse, they are not displayed as
	/// outlined characters; and if set to \c bpvUndefined, some of the characters are displayed as
	/// outlined characters and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Outline, put_Emboss, put_Engrave, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Outline, put_Emboss, put_Engrave, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Outline(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Protected property</em>
	///
	/// Retrieves whether the text range is protected against modifications. Any of the values
	/// defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the
	/// text range is protected against modifications; if set to \c bpvFalse, it is not protected; and if
	/// set to \c bpvUndefined, parts of the text range are protected and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Protected, get_CanChange, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Protected, get_CanChange, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Protected(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Protected property</em>
	///
	/// Sets whether the text range is protected against modifications. Any of the values
	/// defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the
	/// text range is protected against modifications; if set to \c bpvFalse, it is not protected; and if
	/// set to \c bpvUndefined, parts of the text range are protected and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Protected, get_CanChange, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	///      
	/// \else
	///   \sa get_Protected, get_CanChange, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Protected(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Revised property</em>
	///
	/// Retrieves whether the text range has been revised. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the text range has
	/// been revised; if set to \c bpvFalse, it has not been revised; and if set to \c bpvUndefined, parts
	/// of the text range have been revised and others not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa put_Revised, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Revised, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Revised(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Revised property</em>
	///
	/// Sets whether the text range has been revised. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the text range has
	/// been revised; if set to \c bpvFalse, it has not been revised; and if set to \c bpvUndefined, parts
	/// of the text range have been revised and others not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa get_Revised, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	///      
	/// \else
	///   \sa get_Revised, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Revised(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Shadow property</em>
	///
	/// Retrieves whether the characters are displayed as shadowed characters. Any of the values
	/// defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the
	/// characters are displayed as shadowed characters; if set to \c bpvFalse, they are displayed
	/// normally; and if set to \c bpvUndefined, some of the characters are displayed as shadowed
	/// characters and others are displayed normally.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Shadow, get_Bold, get_Italic, get_Name, get_Size,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Shadow, get_Bold, get_Italic, get_Name, get_Size,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Shadow(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Shadow property</em>
	///
	/// Sets whether the characters are displayed as shadowed characters. Any of the values
	/// defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the
	/// characters are displayed as shadowed characters; if set to \c bpvFalse, they are displayed
	/// normally; and if set to \c bpvUndefined, some of the characters are displayed as shadowed
	/// characters and others are displayed normally.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Shadow, put_Bold, put_Italic, put_Name, put_Size,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Shadow, put_Bold, put_Italic, put_Name, put_Size,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Shadow(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Size property</em>
	///
	/// Retrieves the size of the font the text is drawn in. If the text range uses multiple font
	/// sizes, the size will be reported as -9999999.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Size, get_Name, get_Spacing, get_KerningSize, get_HorizontalScaling, RichTextBox::get_Font
	virtual HRESULT STDMETHODCALLTYPE get_Size(FLOAT* pValue);
	/// \brief <em>Sets the \c Size property</em>
	///
	/// Sets the size of the font the text is drawn in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Size, put_Name, put_Spacing, put_KerningSize, put_HorizontalScaling, RichTextBox::putref_Font
	virtual HRESULT STDMETHODCALLTYPE put_Size(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c SmallCaps property</em>
	///
	/// Retrieves whether the characters are in small capital letters. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are in
	/// small capital letters; if set to \c bpvFalse, they are not in small capital letters; and if set to
	/// \c bpvUndefined, some of the characters are in small capital letters and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SmallCaps, get_AllCaps, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_SmallCaps, get_AllCaps, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SmallCaps(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c SmallCaps property</em>
	///
	/// Sets whether the characters are in small capital letters. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are in
	/// small capital letters; if set to \c bpvFalse, they are not in small capital letters; and if set to
	/// \c bpvUndefined, some of the characters are in small capital letters and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SmallCaps, put_AllCaps, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_SmallCaps, put_AllCaps, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SmallCaps(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Spacing property</em>
	///
	/// Retrieves the amount of horizontal spacing between characters, in floating-point points.
	/// If the text range uses multiple spacings, the spacing will be reported as -9999999.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Spacing, get_HorizontalScaling, get_Size, get_Name
	virtual HRESULT STDMETHODCALLTYPE get_Spacing(FLOAT* pValue);
	/// \brief <em>Sets the \c Spacing property</em>
	///
	/// Retrieves the amount of horizontal spacing between characters, in floating-point points.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Spacing, put_HorizontalScaling, put_Size, put_Name
	virtual HRESULT STDMETHODCALLTYPE put_Spacing(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c StrikeThrough property</em>
	///
	/// Retrieves whether the characters are displayed with a horizontal line through the center. Any of the
	/// values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue,
	/// the characters are stroke out; if set to \c bpvFalse, they are not stroke out; and if set to
	/// \c bpvUndefined, some of the characters are stroke out and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_StrikeThrough, get_UnderlineType, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_StrikeThrough, get_UnderlineType, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StrikeThrough(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c StrikeThrough property</em>
	///
	/// Sets whether the characters are displayed with a horizontal line through the center. Any of the
	/// values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue,
	/// the characters are stroke out; if set to \c bpvFalse, they are not stroke out; and if set to
	/// \c bpvUndefined, some of the characters are stroke out and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_StrikeThrough, put_UnderlineType, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_StrikeThrough, put_UnderlineType, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_StrikeThrough(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c StyleID property</em>
	///
	/// Retrieves the unique ID of the style assigned to the text. If the text range uses multiple styles,
	/// the style will be reported as -9999999.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_StyleID, RichTextBoxCtlLibU::BuiltInStyleConstants
	/// \else
	///   \sa put_StyleID, RichTextBoxCtlLibA::BuiltInStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StyleID(LONG* pValue);
	/// \brief <em>Sets the \c StyleID property</em>
	///
	/// Retrieves the unique ID of the style assigned to the text.\n
	/// This property can also be set to one of the built-in styles defined by the \c BuiltInStyleConstants
	/// enumeration.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_StyleID, RichTextBoxCtlLibU::BuiltInStyleConstants
	/// \else
	///   \sa get_StyleID, RichTextBoxCtlLibA::BuiltInStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_StyleID(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Subscript property</em>
	///
	/// Retrieves whether the characters are displayed as subscript. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// displayed as subscript; if set to \c bpvFalse, they are not displayed as subscript; and if set to
	/// \c bpvUndefined, some of the characters are displayed as subscript and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Subscript, get_Superscript, get_VerticalOffset,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Subscript, get_Superscript, get_VerticalOffset,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Subscript(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Subscript property</em>
	///
	/// Sets whether the characters are displayed as subscript. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// displayed as subscript; if set to \c bpvFalse, they are not displayed as subscript; and if set to
	/// \c bpvUndefined, some of the characters are displayed as subscript and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Subscript, put_Superscript, put_VerticalOffset,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Subscript, put_Superscript, put_VerticalOffset,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Subscript(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Superscript property</em>
	///
	/// Retrieves whether the characters are displayed as superscript. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// displayed as superscript; if set to \c bpvFalse, they are not displayed as superscript; and if set to
	/// \c bpvUndefined, some of the characters are displayed as superscript and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Superscript, get_Subscript, get_VerticalOffset,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Superscript, get_Subscript, get_VerticalOffset,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Superscript(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Superscript property</em>
	///
	/// Sets whether the characters are displayed as superscript. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the characters are
	/// displayed as superscript; if set to \c bpvFalse, they are not displayed as superscript; and if set to
	/// \c bpvUndefined, some of the characters are displayed as superscript and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Superscript, put_Subscript, put_VerticalOffset,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Superscript, put_Subscript, put_VerticalOffset,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Superscript(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c TeXStyle property</em>
	///
	/// Retrieves the TeX style of the font, if the font refers to a math zone. Any of the values defined
	/// by the \c TeXFontStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_TeXStyle, get_IsMathZone, RichTextBoxCtlLibU::TeXFontStyleConstants
	/// \else
	///   \sa put_TeXStyle, get_IsMathZone, RichTextBoxCtlLibA::TeXFontStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TeXStyle(TeXFontStyleConstants* pValue);
	/// \brief <em>Sets the \c TeXStyle property</em>
	///
	/// Sets the TeX style of the font, if the font refers to a math zone. Any of the values defined
	/// by the \c TeXFontStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_TeXStyle, put_IsMathZone, RichTextBoxCtlLibU::TeXFontStyleConstants
	/// \else
	///   \sa get_TeXStyle, put_IsMathZone, RichTextBoxCtlLibA::TeXFontStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TeXStyle(TeXFontStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c UnderlineColorIndex property</em>
	///
	/// Retrieves the zero-based index of the color in which the text's underline is drawn. If set to zero,
	/// the underline is drawn in the same color as the text; otherwise the value is the index of the effect
	/// color to use. Effect colors are specified by the \c IRichTextBox::EffectColor property.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UnderlineColorIndex, RichTextBox::get_EffectColor, get_ForeColor, get_UnderlineType,
	///     get_UnderlinePosition
	virtual HRESULT STDMETHODCALLTYPE get_UnderlineColorIndex(LONG* pValue);
	/// \brief <em>Sets the \c UnderlineColorIndex property</em>
	///
	/// Sets the zero-based index of the color in which the text's underline is drawn. If set to zero,
	/// the underline is drawn in the same color as the text; otherwise the value is the index of the effect
	/// color to use. Effect colors are specified by the \c IRichTextBox::EffectColor property.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UnderlineColorIndex, RichTextBox::put_EffectColor, put_ForeColor, put_UnderlineType,
	///     put_UnderlinePosition
	virtual HRESULT STDMETHODCALLTYPE put_UnderlineColorIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c UnderlinePosition property</em>
	///
	/// Retrieves the position that the text's underlining is drawn at. Any of the values defined by the
	/// \c UnderlinePositionConstants enumeration is valid. If the text range uses multiple underline
	/// positions, the position will be reported as \c upUndefined.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa put_UnderlinePosition, get_UnderlineType, get_UnderlineColorIndex,
	///       RichTextBoxCtlLibU::UnderlinePositionConstants
	/// \else
	///   \sa put_UnderlinePosition, get_UnderlineType, get_UnderlineColorIndex,
	///       RichTextBoxCtlLibA::UnderlinePositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_UnderlinePosition(UnderlinePositionConstants* pValue);
	/// \brief <em>Sets the \c UnderlinePosition property</em>
	///
	/// Sets the position that the text's underlining is drawn at. Any of the values defined by the
	/// \c UnderlinePositionConstants enumeration is valid. If the text range uses multiple underline
	/// positions, the position will be reported as \c upUndefined.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa get_UnderlinePosition, put_UnderlineType, put_UnderlineColorIndex,
	///       RichTextBoxCtlLibU::UnderlinePositionConstants
	/// \else
	///   \sa get_UnderlinePosition, put_UnderlineType, put_UnderlineColorIndex,
	///       RichTextBoxCtlLibA::UnderlinePositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_UnderlinePosition(UnderlinePositionConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c UnderlineType property</em>
	///
	/// Retrieves whether the characters are underlined. Any of the values defined by the
	/// \c UnderlineTypeConstants enumeration is valid. If the text range uses multiple underline
	/// types, the type will be reported as \c utUndefined.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_UnderlineType, get_UnderlineColorIndex, get_Weight, get_Italic, get_Name, get_Size,
	///       get_UnderlinePosition, RichTextBoxCtlLibU::UnderlineTypeConstants
	/// \else
	///   \sa put_UnderlineType, get_UnderlineColorIndex, get_Weight, get_Italic, get_Name, get_Size,
	///       get_UnderlinePosition, RichTextBoxCtlLibA::UnderlineTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_UnderlineType(UnderlineTypeConstants* pValue);
	/// \brief <em>Sets the \c UnderlineType property</em>
	///
	/// Sets whether the characters are underlined. Any of the values defined by the
	/// \c UnderlineTypeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_UnderlineType, put_UnderlineType, put_Weight, put_Italic, put_Name, put_Size,
	///       put_UnderlinePosition, RichTextBoxCtlLibU::UnderlineTypeConstants
	/// \else
	///   \sa get_UnderlineType, put_UnderlineType, put_Weight, put_Italic, put_Name, put_Size,
	///       put_UnderlinePosition, RichTextBoxCtlLibA::UnderlineTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_UnderlineType(UnderlineTypeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c VerticalOffset property</em>
	///
	/// Retrieves the amount that characters are offset vertically relative to the baseline, in
	/// floating-point points. If the text range uses multiple offsets, the offset will be reported as
	/// -9999999.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_VerticalOffset, get_Subscript, put_Superscript
	virtual HRESULT STDMETHODCALLTYPE get_VerticalOffset(FLOAT* pValue);
	/// \brief <em>Sets the \c VerticalOffset property</em>
	///
	/// Sets the amount that characters are offset vertically relative to the baseline, in
	/// floating-point points.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_VerticalOffset, put_Subscript, put_Superscript
	virtual HRESULT STDMETHODCALLTYPE put_VerticalOffset(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c Weight property</em>
	///
	/// Retrieves the font weight the characters are drawn with. Any of the values defined by the
	/// \c FontWeightConstants enumeration is valid. If the text range uses multiple font weights,
	/// the weight will be reported as \c fwUndefined.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Weight, get_Bold, get_Italic, get_Name, get_Size, RichTextBoxCtlLibU::FontWeightConstants
	/// \else
	///   \sa put_Weight, get_Bold, get_Italic, get_Name, get_Size, RichTextBoxCtlLibA::FontWeightConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Weight(FontWeightConstants* pValue);
	/// \brief <em>Sets the \c Weight property</em>
	///
	/// Sets the font weight the characters are drawn with. Any of the values defined by the
	/// \c FontWeightConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Weight, put_Bold, put_Italic, put_Name, put_Size, RichTextBoxCtlLibU::FontWeightConstants
	/// \else
	///   \sa get_Weight, put_Bold, put_Italic, put_Name, put_Size, RichTextBoxCtlLibA::FontWeightConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Weight(FontWeightConstants newValue);

	/// \brief <em>Creates a new \c IRichTextFont instance with identical properties</em>
	///
	/// \param[out] ppClone Receives the new \c IRichTextFont instance.
	///
	/// \sa CopySettings, Equals
	virtual HRESULT STDMETHODCALLTYPE Clone(IRichTextFont** ppClone);
	/// \brief <em>Copies all settings from an other \c IRichTextFont instance</em>
	///
	/// \param[in] pSourceObject The \c IRichTextFont instance to copy from.
	///
	/// \sa Clone, Equals
	virtual HRESULT STDMETHODCALLTYPE CopySettings(IRichTextFont* pSourceObject);
	/// \brief <em>Retrieves whether two \c IRichTextFont instances have the same properties</em>
	///
	/// \param[in] pCompareAgainst The \c IRichTextFont instance to compare with.
	/// \param[out] pValue \c VARIANT_TRUE, if the objects have the same properties; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \sa Clone, CopySettings
	virtual HRESULT STDMETHODCALLTYPE Equals(IRichTextFont* pCompareAgainst, VARIANT_BOOL* pValue);
	/// \brief <em>Resets the text range's character formatting</em>
	///
	/// Resets the text range's character formatting, applying the specified rules.
	///
	/// \param[in] rules The rules to apply. Any of the values defined by the \c ResetRulesConstants
	///            enumeration is valid.
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::ResetRulesConstants
	/// \else
	///   \sa RichTextBoxCtlLibA::ResetRulesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Reset(ResetRulesConstants rules);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given text range styling</em>
	///
	/// Attaches this object to a given text range styling, so that the text range styling's properties can
	/// be retrieved and set using this object's methods.
	///
	/// \param[in] pTextFont The \c ITextFont object to attach to.
	///
	/// \sa Detach,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774054.aspx">ITextFont</a>
	void Attach(ITextFont* pTextFont);
	/// \brief <em>Detaches this object from a text range styling</em>
	///
	/// Detaches this object from the text range styling it currently wraps, so that it doesn't wrap any text
	/// range styling anymore.
	///
	/// \sa Attach
	void Detach(void);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ITextFont object corresponding to the text range styling wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774054.aspx">ITextFont</a>
		ITextFont* pTextFont;

		Properties()
		{
			pTextFont = NULL;
		}

		~Properties();
	} properties;
};     // TextFont

OBJECT_ENTRY_AUTO(__uuidof(TextFont), TextFont)