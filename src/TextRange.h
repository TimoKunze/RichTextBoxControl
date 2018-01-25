//////////////////////////////////////////////////////////////////////
/// \class TextRange
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a range of text</em>
///
/// This class is a wrapper around a range of text.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTextRange, RichTextBox, TextFont, TextParagraph
/// \else
///   \sa RichTextBoxCtlLibA::IRichTextRange, RichTextBox, TextFont, TextParagraph
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITextRangeEvents_CP.h"
#include "helpers.h"
#include "OLEObject.h"
#include "Table.h"
#include "TextFont.h"
#include "TextParagraph.h"
#include "TextSubRanges.h"
#include "RichTextBox.h"


class ATL_NO_VTABLE TextRange : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TextRange, &CLSID_TextRange>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TextRange>,
    public Proxy_IRichTextRangeEvents<TextRange>, 
    #ifdef UNICODE
    	public IDispatchImpl<IRichTextRange, &IID_IRichTextRange, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTextRange, &IID_IRichTextRange, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class RichTextBox;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TEXTRANGE)

		BEGIN_COM_MAP(TextRange)
			COM_INTERFACE_ENTRY(IRichTextRange)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange2), 0, QueryITextRangeInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TextRange)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTextRangeEvents))
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

	/// \brief <em>Will be called if this class is \c QueryInterface()'d for \c ITextRange or \c ITextRange2</em>
	///
	/// This method will be called if the class's \c QueryInterface() method is called with
	/// \c IID_ITextRange or \c IID_ITextRange2. We forward the call to the wrapped \c ITextRange or
	/// \c ITextRange2 implementation.
	///
	/// \param[in] pThis The instance of this class, that the interface is queried from.
	/// \param[in] queriedInterface Should be \c IID_ITextRange or \c IID_ITextRange2.
	/// \param[out] ppImplementation Receives the wrapped object's \c ITextRange or \c ITextRange2
	///             implementation.
	/// \param[in] cookie A \c DWORD value specified in the COM interface map.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/hh768626.aspx">ITextRange2</a>
	static HRESULT CALLBACK QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTextRange
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c CanEdit property</em>
	///
	/// Retrieves whether the text in the text range can be edited. If set to \c bpvTrue, the text can
	/// be edited; if set to \c bpvFalse, it cannot be edited; and if set to \c bpvUndefined, parts of the
	/// text range can be edited and others can't.
	/// The text cannot be edited, if the document is read-only or any part of the text is protected.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa CanPaste, get_Text, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa CanPaste, get_Text, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CanEdit(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c EmbeddedObject property</em>
	///
	/// Retrieves the embedded OLE object that is positioned at the text range's start point. If the text
	/// range spans more than the embedded object, this property will be \c NULL.
	///
	/// \param[out] ppOLEObject The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa OLEObject, RichTextBox::get_EmbeddedObjects
	virtual HRESULT STDMETHODCALLTYPE get_EmbeddedObject(IRichOLEObject** ppOLEObject);
	/// \brief <em>Retrieves the current setting of the \c FirstChar property</em>
	///
	/// Retrieves the first character of the text range.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FirstChar, get_Text, get_RangeStart
	virtual HRESULT STDMETHODCALLTYPE get_FirstChar(LONG* pValue);
	/// \brief <em>Sets the \c FirstChar property</em>
	///
	/// Sets the first character of the text range.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstChar, put_Text, put_RangeStart
	virtual HRESULT STDMETHODCALLTYPE put_FirstChar(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves an object that can be used to retrieve and set the text range's styling, e.g. font and
	/// colors.
	///
	/// \param[out] ppTextFont The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, TextFont, get_Paragraph, RichTextBox::get_DocumentFont
	virtual HRESULT STDMETHODCALLTYPE get_Font(IRichTextFont** ppTextFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the text range's styling, e.g. font and colors.
	///
	/// \param[in] pNewTextFont The new styling object's \c IRichTextFont implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Font, TextFont, put_Paragraph, RichTextBox::put_DocumentFont
	virtual HRESULT STDMETHODCALLTYPE put_Font(IRichTextFont* pNewTextFont);
	/// \brief <em>Retrieves the current setting of the \c Paragraph property</em>
	///
	/// Retrieves an object that can be used to retrieve and set the text range's styling that applies
	/// to the whole paragraph, e.g. alignment and line spacing.
	///
	/// \param[out] ppTextParagraph The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Paragraph, TextParagraph, get_Font, RichTextBox::get_DocumentParagraph
	virtual HRESULT STDMETHODCALLTYPE get_Paragraph(IRichTextParagraph** ppTextParagraph);
	/// \brief <em>Sets the \c Paragraph property</em>
	///
	/// Sets the text range's styling that applies to the whole paragraph, e.g. alignment and line spacing.
	///
	/// \param[in] pNewTextParagraph The new styling object's \c IRichTextParagraph implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Paragraph, TextParagraph, put_Font, RichTextBox::put_DocumentParagraph
	virtual HRESULT STDMETHODCALLTYPE put_Paragraph(IRichTextParagraph* pNewTextParagraph);
	/// \brief <em>Retrieves the current setting of the \c RangeEnd property</em>
	///
	/// Retrieves the zero-based index of the first character after the text range.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RangeEnd, get_RangeStart, get_RangeLength, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_RangeEnd(LONG* pValue);
	/// \brief <em>Sets the current setting of the \c RangeEnd property</em>
	///
	/// Sets the zero-based index of the first character after the text range.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RangeEnd, put_RangeStart, SetStartAndEnd, get_RangeLength, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_RangeEnd(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c RangeLength property</em>
	///
	/// Retrieves the length (in characters) of the text range.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_RangeStart, get_RangeEnd, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_RangeLength(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c RangeStart property</em>
	///
	/// Retrieves the zero-based index of the first character in the text range.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RangeStart, get_RangeEnd, get_RangeLength, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_RangeStart(LONG* pValue);
	/// \brief <em>Sets the current setting of the \c RangeStart property</em>
	///
	/// Sets the zero-based index of the first character in the text range.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RangeStart, put_RangeEnd, SetStartAndEnd, get_RangeLength, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_RangeStart(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c RichText property</em>
	///
	/// Retrieves the text range's content, including RTF control words for formatting.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RichText, get_Text, get_Font, get_Paragraph, CopyRichTextFromTextRange
	virtual HRESULT STDMETHODCALLTYPE get_RichText(BSTR* pValue);
	/// \brief <em>Sets the \c RichText property</em>
	///
	/// Sets the text range's content, including RTF control words for formatting.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RichText, put_Text, put_Font, put_Paragraph, CopyRichTextFromTextRange
	virtual HRESULT STDMETHODCALLTYPE put_RichText(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c StoryLength property</em>
	///
	/// Retrieves the count of characters in the story that the text range belongs to.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_StoryType
	virtual HRESULT STDMETHODCALLTYPE get_StoryLength(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c StoryType property</em>
	///
	/// Retrieves the type of the story that the text range belongs to. Any of the values defined by the
	/// \c StoryTypeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_StoryLength, RichTextBoxCtlLibU::StoryTypeConstants
	/// \else
	///   \sa get_StoryLength, RichTextBoxCtlLibA::StoryTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StoryType(StoryTypeConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c SubRanges property</em>
	///
	/// Retrieves a collection object wrapping the additional text ranges that are assigned to this text
	/// range. For instance sub-ranges are used with multi-selection.
	///
	/// \param[out] ppSubRanges Receives the collection object's \c IRichTextSubRanges implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.\n
	///          This property is read-only.
	///
	/// \sa RichTextBox::get_MultiSelect, TextSubRanges
	virtual HRESULT STDMETHODCALLTYPE get_SubRanges(IRichTextSubRanges** ppSubRanges);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the text range's content. It is the plain text, without RTF formatting.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, get_RichText, get_URL, FindText, Delete, get_Font, RichTextBox::get_Text
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the text range's content. It is the plain text, without RTF formatting.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, put_RichText, put_URL, FindText, Delete, put_Font, RichTextBox::put_Text
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c UnitIndex property</em>
	///
	/// Retrieves the one-based index of the specified unit (within the story) at the text range's start
	/// position. For instance, if the range starts within line 10, its index of the \c uLine unit is 10.
	///
	/// \param[in] unit The unit for which to retrieve the index. Any of the values defined by the
	///            \c UnitConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa MoveToUnitIndex, MoveInsertionPointByUnit, RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa MoveToUnitIndex, MoveInsertionPointByUnit, RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_UnitIndex(UnitConstants unit, LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c URL property</em>
	///
	/// Retrieves the URL that is associated with the text range. It is used to turn arbitrary text into a
	/// hyperlink.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.\n
	///          Transforming the link into normal text by setting the URL to an empty string requires Rich
	///          Edit 7.0 or newer.
	///
	/// \sa put_URL, get_RichText, get_Text, RichTextBox::get_AutoDetectURLs
	virtual HRESULT STDMETHODCALLTYPE get_URL(BSTR* pValue);
	/// \brief <em>Sets the \c URL property</em>
	///
	/// Sets the URL that is associated with the text range. It is used to turn arbitrary text into a
	/// hyperlink.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 6.0 or newer.\n
	///          Transforming the link into normal text by setting the URL to an empty string requires Rich
	///          Edit 7.0 or newer.
	///
	/// \sa get_URL, put_RichText, put_Text, RichTextBox::put_AutoDetectURLs
	virtual HRESULT STDMETHODCALLTYPE put_URL(BSTR newValue);

	/// \brief <em>Changes the math in the text range from built-up form to linear format</em>
	///
	/// Modifies the representation of math in the text range.\n
	/// Math can be displayed in linear format, e.g. \c a=(b^2+c^2)/d, or it can be displayed in built-up
	/// form: \f$a=\frac{b^2+c^2}{d}\f$. This method changes from built-up form to linear format.
	///
	/// \param[in] flags Influences the build-down operation. Some combinations of the values defined by the
	///            \c BuildUpMathConstants enumeration are valid.
	/// \param[out] pDidAnyChanges \c VARIANT_TRUE, if the range has been changed; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa BuildUpMath, RichTextBox::get_AllowMathZoneInsertion, RichTextFont::get_IsMathZone,
	///       RichTextBoxCtlLibU::BuildUpMathConstants
	/// \else
	///   \sa BuildUpMath, RichTextBox::get_AllowMathZoneInsertion, RichTextFont::get_IsMathZone,
	///       RichTextBoxCtlLibA::BuildUpMathConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE BuildDownMath(BuildUpMathConstants flags, VARIANT_BOOL* pDidAnyChanges);
	/// \brief <em>Changes the math in the text range from linear format to built-up form</em>
	///
	/// Modifies the representation of math in the text range.\n
	/// Math can be displayed in linear format, e.g. \c a=(b^2+c^2)/d, or it can be displayed in built-up
	/// form: \f$a=\frac{b^2+c^2}{d}\f$. This method changes from linear format to built-up form, or it
	/// modifies the built-up form.
	///
	/// \param[in] flags Influences the build-up operation. Some combinations of the values defined by the
	///            \c BuildUpMathConstants enumeration are valid.
	/// \param[out] pDidAnyChanges \c VARIANT_TRUE, if the range has been changed; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 8.0 or newer.
	///
	/// \if UNICODE
	///   \sa BuildDownMath, RichTextBox::get_AllowMathZoneInsertion, RichTextFont::get_IsMathZone,
	///       RichTextBoxCtlLibU::BuildUpMathConstants
	/// \else
	///   \sa BuildDownMath, RichTextBox::get_AllowMathZoneInsertion, RichTextFont::get_IsMathZone,
	///       RichTextBoxCtlLibA::BuildUpMathConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE BuildUpMath(BuildUpMathConstants flags, VARIANT_BOOL* pDidAnyChanges);
	/// \brief <em>Determines whether the content of the clipboard can be pasted to the text range</em>
	///
	/// Determines whether the content of the clipboard or a specific data object can be pasted to the text
	/// range, replacing its content.
	///
	/// \param[in] pOLEDataObject The \c OLEDataObject object from which to paste. If this parameter is not
	///            specified or does not refer to an object, the text will be pasted from the clipboard.
	/// \param[in] formatID An integer value specifying the format in which to paste the data. Valid values
	///            are those defined by VB's \c ClipBoardConstants enumeration, but also any other format
	///            that has been registered using the \c RegisterClipboardFormat API function.\n
	///            If set to \c 0, the format will be chosen automatically.
	/// \param[out] pCanPaste \c VARIANT_TRUE if the clipboard's content can be pasted; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Paste, Copy, Cut, get_Text, TargetOLEDataObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>
	virtual HRESULT STDMETHODCALLTYPE CanPaste(IOLEDataObject* pOLEDataObject = NULL, LONG formatID = 0, VARIANT_BOOL* pCanPaste = NULL);
	/// \brief <em>Changes the case of the letters in the text range</em>
	///
	/// \param[in] newCase The case to apply. Any of the values defined by the
	///            \c CaseConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Text, RichTextBoxCtlLibU::CaseConstants
	/// \else
	///   \sa get_Text, RichTextBoxCtlLibA::CaseConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE ChangeCase(CaseConstants newCase);
	/// \brief <em>Creates a new \c IRichTextRange instance with identical properties</em>
	///
	/// \param[out] ppClone Receives the new \c IRichTextRange instance.
	///
	/// \sa CopyRichTextFromTextRange, Equals
	virtual HRESULT STDMETHODCALLTYPE Clone(IRichTextRange** ppClone);
	/// \brief <em>Collapses the text range into a degenerate point at either end of the text range</em>
	///
	/// Collapses the text range into a degenerate point (insertion point) at either the beginning or end
	/// of the text range.
	///
	/// \param[in] collapseToStart If \c VARIANT_TRUE, the text range is collapsed to its beginning,
	///            otherwise to its end.
	/// \param[out] pSucceeded \c VARIANT_TRUE on success; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExpandToUnit, MoveStartToStartOfUnit, MoveEndToEndOfUnit
	virtual HRESULT STDMETHODCALLTYPE Collapse(VARIANT_BOOL collapseToStart = VARIANT_TRUE, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Retrieves whether the text range contains another text range</em>
	///
	/// Retrieves whether the text range contains the specified text range.
	///
	/// \param[in] compareAgainst The \c IRichTextRange instance to compare with.
	/// \param[out] pValue \c VARIANT_TRUE, if the specified text range is contained by this text range;
	///             otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Equals
	virtual HRESULT STDMETHODCALLTYPE ContainsRange(IRichTextRange* pCompareAgainst, VARIANT_BOOL* pValue);
	/// \brief <em>Copies the content of the text range to the clipboard</em>
	///
	/// Copies the text range's content to the clipboard or to a specific data object.
	///
	/// \param[out] pOLEDataObject Receives an instance of \c IOLEDataObject, wrapping the data object to
	///             which the text range's content will have been copied. If this parameter is not specified
	///             or does not refer to an object, the content will have been copied to the clipboard.
	/// \param[out] pSucceeded \c VARIANT_TRUE on success; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Cut, Paste, get_Text, TargetOLEDataObject
	virtual HRESULT STDMETHODCALLTYPE Copy(VARIANT* pOLEDataObject = NULL, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Replaces the text range's content with the formatted (rich) text of another text range</em>
	///
	/// Replaces the text range's content with the content of another text range, including RTF control
	/// words for formatting.
	///
	/// \param[in] pSourceObject The \c TextRange object from which to copy the content.
	/// \param[out] pSucceeded \c VARIANT_TRUE on success; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RichText, Copy
	virtual HRESULT STDMETHODCALLTYPE CopyRichTextFromTextRange(IRichTextRange* pSourceObject, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Cuts the content of the text range to the clipboard</em>
	///
	/// Copies the text range's content to the clipboard or to a specific data object and deletes it from
	/// the control.
	///
	/// \param[out] pOLEDataObject Receives an instance of \c IOLEDataObject, wrapping the data object to
	///             which the text range's content will have been copied. If this parameter is not specified
	///             or does not refer to an object, the content will have been copied to the clipboard.
	/// \param[out] pSucceeded \c VARIANT_TRUE on success; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Copy, Paste, get_Text, TargetOLEDataObject
	virtual HRESULT STDMETHODCALLTYPE Cut(VARIANT* pOLEDataObject = NULL, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Deletes text from the text range</em>
	///
	/// Mimics the [DEL] or [BACKSPACE] key deleting the specified number of units, e.g. the next 2 words,
	/// or the previous 3 characters.
	///
	/// \param[in] unit The unit by which to select the text to delete. Any of the values defined by the
	///            \c UnitConstants enumeration is valid.
	/// \param[in] Count The number of units to delete. If \c Count is greater than zero, the method acts
	///            as if the [DEL] key was pressed \c Count times. If \c Count is less than zero, the
	///            method acts as if the [BACKSPACE] key was pressed \c Count times. If \c Count is zero,
	///            the whole text in the range is deleted.
	/// \param[out] pDelta Receives the actual number of units that has been deleted.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Text, get_UnitIndex, RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa put_Text, get_UnitIndex, RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Delete(UnitConstants unit = uCharacter, LONG count = 0, LONG* pDelta = NULL);
	/// \brief <em>Retrieves whether two \c IRichTextRange instances have the same properties</em>
	///
	/// \param[in] pCompareAgainst The \c IRichTextRange instance to compare with.
	/// \param[out] pValue \c VARIANT_TRUE, if the objects have the same properties; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \sa EqualsStory, Clone, ContainsRange
	virtual HRESULT STDMETHODCALLTYPE Equals(IRichTextRange* pCompareAgainst, VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves whether two \c IRichTextRange instances belong to the same story</em>
	///
	/// \param[in] pCompareAgainst The \c IRichTextRange instance to compare with.
	/// \param[out] pValue \c VARIANT_TRUE, if the objects belong to the same story; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \sa Equals
	virtual HRESULT STDMETHODCALLTYPE EqualsStory(IRichTextRange* pCompareAgainst, VARIANT_BOOL* pValue);
	/// \brief <em>Expand the text range's boundaries to include the specified overlapping unit</em>
	///
	/// Moves the start and/or end position of the text range to completely include the specified unit.
	///
	/// \param[in] unit The unit which to include in the text range. Any of the values defined by the
	///            \c UnitConstants enumeration is valid.
	/// \param[out] pDelta Receives the count of characters that have been added to the text range.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Collapse, MoveStartToStartOfUnit, MoveEndToEndOfUnit, RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa Collapse, MoveStartToStartOfUnit, MoveEndToEndOfUnit, RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE ExpandToUnit(UnitConstants unit, LONG* pDelta);
	/// \brief <em>Searches for the specified text and moves the text range to the match if found</em>
	///
	/// Moves the text range or one of its boundaries to the first occurance of the specified string. The
	/// method is able to search for text, words and regular expressions. The search can be limited to a
	/// specific range of text.
	///
	/// \param[in] searchFor The string to search for. Depending on the \c searchMode parameter, this string
	///            will be treated as plain text, as a word, or as a regular expression.
	/// \param[in] searchDirectionAndMaxCharsToPass Specifies whether to search forward, backward or within
	///            the text range only. Any of the values defined by the \c SearchDirectionConstants
	///            enumeration and any integer value is valid.\n
	///            If \c searchDirectionAndMaxCharsToPass is greater than zero, the search moves forward,
	///            toward the end of the story, and is aborted after this number of characters has been
	///            passed, starting at the text range's start position. If
	///            \c searchDirectionAndMaxCharsToPass is less than zero, the search moves backward, toward
	///            the beginning of the story, and is aborted after this number of characters has been
	///            passed, starting at the text range's start position.\n
	///            If \c searchDirectionAndMaxCharsToPass is \c sdWithinRange and the text range is
	///            degenerated (an insertion point), the search starts after the range and is not limited
	///            in its length. If \c searchDirectionAndMaxCharsToPass is \c sdWithinRange and the text
	///            range has a length greater than zero, the search is limited to the text range.
	/// \param[in] searchMode Specifies whether to perform a case-sensitive or case-insensitive search,
	///            whether to search for whole words, and whether to treat the search string as regular
	///            expression. Any combination of the values defined by the \c SearchModeConstants
	///            enumeration is valid.
	/// \param[in] moveRangeBoundaries Specifies which ends of the text range shall be moved to the match,
	///            in case a match is found. Any of the values defined by the
	///            \c MoveRangeBoundariesConstants enumeration is valid.
	/// \param[out] pLengthMatched Receives the length of the string that did match.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Text, RichTextBoxCtlLibU::SearchDirectionConstants,
	///       RichTextBoxCtlLibU::SearchModeConstants, RichTextBoxCtlLibU::MoveRangeBoundariesConstants
	/// \else
	///   \sa get_Text, RichTextBoxCtlLibA::SearchDirectionConstants,
	///       RichTextBoxCtlLibA::SearchModeConstants, RichTextBoxCtlLibA::MoveRangeBoundariesConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE FindText(BSTR searchFor, SearchDirectionConstants searchDirectionAndMaxCharsToPass = sdForward, SearchModeConstants searchMode = smDefault, MoveRangeBoundariesConstants moveRangeBoundaries = mrbMoveBoth, LONG* pLengthMatched = NULL);
	/// \brief <em>Retrieves the coordinates of the text range's end position</em>
	///
	/// Retrieves the coordinates of the text range's end position.
	///
	/// \param[in] horizontalPosition Specifies the horizontal position for which to retrieve the
	///            coordinates. Any of the values defined by the \c HorizontalPositionConstants enumeration
	///            is valid.
	/// \param[in] verticalPosition Specifies the vertical position for which to retrieve the coordinates.
	///            Any of the values defined by the \c VerticalPositionConstants enumeration is valid.
	/// \param[in] flags Controls the method's behavior. Any combination of the values defined by the
	///            \c RangePositionConstants enumeration is valid.
	/// \param[out] pX Receives the x-coordinate (in twips) of the text range's end position relative to
	///             either the control's or the screen's upper-left corner, as specified by the \c flags
	///             parameter.
	/// \param[out] pY Receives the y-coordinate (in twips) of the text range's end position relative to
	///             either the control's or the screen's upper-left corner, as specified by the \c flags
	///             parameter.
	/// \param[out] pSucceeded \c VARIANT_TRUE on success; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetStartPosition, MoveEndToPosition, RichTextBoxCtlLibU::HorizontalPositionConstants,
	///       RichTextBoxCtlLibU::VerticalPositionConstants, RichTextBoxCtlLibU::RangePositionConstants
	/// \else
	///   \sa GetStartPosition, MoveEndToPosition, RichTextBoxCtlLibA::HorizontalPositionConstants,
	///       RichTextBoxCtlLibA::VerticalPositionConstants, RichTextBoxCtlLibA::RangePositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetEndPosition(HorizontalPositionConstants horizontalPosition, VerticalPositionConstants verticalPosition, RangePositionConstants flags, OLE_XPOS_PIXELS* pX, OLE_YPOS_PIXELS* pY, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Retrieves the coordinates of the text range's start position</em>
	///
	/// Retrieves the coordinates of the text range's start position.
	///
	/// \param[in] horizontalPosition Specifies the horizontal position for which to retrieve the
	///            coordinates. Any of the values defined by the \c HorizontalPositionConstants enumeration
	///            is valid.
	/// \param[in] verticalPosition Specifies the vertical position for which to retrieve the coordinates.
	///            Any of the values defined by the \c VerticalPositionConstants enumeration is valid.
	/// \param[in] flags Controls the method's behavior. Any combination of the values defined by the
	///            \c RangePositionConstants enumeration is valid.
	/// \param[out] pX Receives the x-coordinate (in twips) of the text range's end position relative to
	///             either the control's or the screen's upper-left corner, as specified by the \c flags
	///             parameter.
	/// \param[out] pY Receives the y-coordinate (in twips) of the text range's end position relative to
	///             either the control's or the screen's upper-left corner, as specified by the \c flags
	///             parameter.
	/// \param[out] pSucceeded \c VARIANT_TRUE on success; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetEndPosition, MoveStartToPosition, RichTextBoxCtlLibU::HorizontalPositionConstants,
	///       RichTextBoxCtlLibU::VerticalPositionConstants, RichTextBoxCtlLibU::RangePositionConstants
	/// \else
	///   \sa GetEndPosition, MoveStartToPosition, RichTextBoxCtlLibA::HorizontalPositionConstants,
	///       RichTextBoxCtlLibA::VerticalPositionConstants, RichTextBoxCtlLibA::RangePositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetStartPosition(HorizontalPositionConstants horizontalPosition, VerticalPositionConstants verticalPosition, RangePositionConstants flags, OLE_XPOS_PIXELS* pX, OLE_YPOS_PIXELS* pY, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Retrieves whether the text range is located within a table</em>
	///
	/// Retrieves whether the text range is contained entirely by a table and returns the \c Table,
	/// \c TableRow, and \c TableCell object if applicable.
	///
	/// \param[out] pTable Receives the \c IRichTable object that the text range lies within.
	/// \param[out] pTableRow Receives the \c IRichTableRow object that the text range lies within.
	/// \param[out] pTableCell Receives the \c IRichTableCell object that the text range lies within.
	/// \param[out] pValue \c VARIANT_TRUE, if the text range lies entirely within a table; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \sa ReplaceWithTable, Table, TableRow, TableCell
	virtual HRESULT STDMETHODCALLTYPE IsWithinTable(VARIANT* pTable = NULL, VARIANT* pTableRow = NULL, VARIANT* pTableCell = NULL, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Moves the text range's end position by the specified number of the specified units</em>
	///
	/// Moves the text range's end position forward or backward by the specified number of units, e.g. by 2
	/// words forward. If the new end position precedes the old start position, the range is collapsed to
	/// an insertion point at the new end position.
	///
	/// \param[in] unit The unit by which to move. Any of the values defined by the \c UnitConstants
	///            enumeration is valid.
	/// \param[in] count The number of units by which to move. If \c count is greater than zero, motion
	///            is forward, toward the end of the story; if it is less than zero, motion is backward.
	/// \param[out] pDelta Receives the actual number of units that the range's end position moved past.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa MoveStartByUnit, MoveInsertionPointByUnit, MoveEndUntil, MoveEndWhile,
	///       RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa MoveStartByUnit, MoveInsertionPointByUnit, MoveEndUntil, MoveEndWhile,
	///       RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE MoveEndByUnit(UnitConstants unit, LONG count, LONG* pDelta);
	/// \brief <em>Moves the text range's end to the specified screen coordinates</em>
	///
	/// Moves the end position of the text range to the specified point or the nearest point with
	/// selectable text.
	///
	/// \param[in] x The x-coordinate (in twips) to which to move the text range's end position, relative
	///            to the screen's upper-left corner.
	/// \param[in] y The y-coordinate (in twips) to which to move the text range's end position, relative
	///            to the screen's upper-left corner.
	/// \param[in] doNotMoveStart If \c VARIANT_TRUE, only the text range's end position is moved to the
	///            specified point, while the range's start position is not moved. If \c VARIANT_FALSE, the
	///            text range is collapsed to the specified position, creating an insertion point.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MoveStartToPosition, GetEndPosition
	virtual HRESULT STDMETHODCALLTYPE MoveEndToPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL doNotMoveStart);
	/// \brief <em>Moves the text range's end to the end of the specified overlapping unit</em>
	///
	/// Moves the end position of the text range to the end of the last overlapping specified unit in the
	/// text range.
	///
	/// \param[in] unit The unit to which to move. Any of the values defined by the \c UnitConstants
	///            enumeration is valid.
	/// \param[in] doNotMoveStart If \c VARIANT_TRUE, only the text range's end position is moved to the end
	///            of the specified unit, while the range's start position is not moved. If \c VARIANT_FALSE,
	///            the text range is collapsed to the end of the specified unit, creating an insertion point.
	/// \param[out] pDelta Receives the actual number of units that the range's end position has been moved
	///             past.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa MoveStartToStartOfUnit, ExpandToUnit, MoveToUnitIndex, MoveInsertionPointByUnit,
	///       RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa MoveStartToStartOfUnit, ExpandToUnit, MoveToUnitIndex, MoveInsertionPointByUnit,
	///       RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE MoveEndToEndOfUnit(UnitConstants unit, VARIANT_BOOL doNotMoveStart, LONG* pDelta);
	/// \brief <em>Moves the text range's end position to the position of the first character found that is in the specified character set</em>
	///
	/// Moves the text range's end position forward or backward, to the position of the first character
	/// found that is in the set of characters specified by \c pCharacterSet, provided that the character is
	/// found within \c maximumCharactersToPass characters of the range's end.\n
	/// If the new end position precedes the old start position, the range is collapsed to an insertion point
	/// at the new end position. If no character from the set specified by \c pCharacterSet is found within
	/// \c maximumCharactersToPass positions of the range's end, the range is left unchanged.
	///
	/// \param[in] pCharacterSet The character set to use in the match. This could be an explicit string of
	///            characters or a character-set index. For instance passing "0123456789" has the same
	///            effect as passing 4 (\c C1_DIGIT): The method finds the first digit.\n
	///            For more information see <a href="https://msdn.microsoft.com/en-us/library/bb787724.aspx#tom_character_match_sets">Character Match Sets</a>.
	/// \param[in] maximumCharactersToPass The maximum number of characters to move past. If
	///            \c maximumCharactersToPass is greater than zero, the search moves forward, toward the
	///            end of the story; if it is less than zero, the search moves backward.
	/// \param[out] pDelta Receives the actual number of characters that the range's end position moved past,
	///             plus 1 for a match.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MoveStartUntil, MoveInsertionPointUntil, MoveEndWhile, MoveEndByUnit
	virtual HRESULT STDMETHODCALLTYPE MoveEndUntil(VARIANT* pCharacterSet, LONG maximumCharactersToPass = tomForward, LONG* pDelta = NULL);
	/// \brief <em>Moves the text range's end position to the position of the first character found that is not in the specified character set</em>
	///
	/// Moves the text range's end position forward or backward, to the position of the first character
	/// found that is not in the set of characters specified by \c pCharacterSet, or by
	/// \c maximumCharactersToPass characters, whichever is less.\n
	/// If the new end position precedes the old start position, the range is collapsed to an insertion
	/// point at the new end position.
	///
	/// \param[in] pCharacterSet The character set to use in the match. This could be an explicit string of
	///            characters or a character-set index. For instance passing "0123456789" has the same
	///            effect as passing 4 (\c C1_DIGIT): The method finds the first digit.\n
	///            For more information see <a href="https://msdn.microsoft.com/en-us/library/bb787724.aspx#tom_character_match_sets">Character Match Sets</a>.
	/// \param[in] maximumCharactersToPass The maximum number of characters to move past. If
	///            \c maximumCharactersToPass is greater than zero, the search moves forward, toward the
	///            end of the story; if it is less than zero, the search moves backward.
	/// \param[out] pDelta Receives the actual number of characters that the range's end position moved past.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MoveStartWhile, MoveInsertionPointWhile, MoveEndUntil, MoveEndByUnit
	virtual HRESULT STDMETHODCALLTYPE MoveEndWhile(VARIANT* pCharacterSet, LONG maximumCharactersToPass = tomForward, LONG* pDelta = NULL);
	/// \brief <em>Moves the insertion point by the specified number of the specified units</em>
	///
	/// Moves the insertion point forward or backward by the specified number of units, e.g. by 2 words
	/// forward. If the text range is nondegenerate (i.e. it is not an insertion point), it is collapsed to
	/// an insertion point at either end before it is moved.
	///
	/// \param[in] unit The unit by which to move. Any of the values defined by the \c UnitConstants
	///            enumeration is valid.
	/// \param[in] count The number of units by which to move. If \c count is greater than zero, motion
	///            is forward, toward the end of the story; if it is less than zero, motion is backward.
	/// \param[out] pDelta Receives the actual number of units that the insertion point moved past.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The insertion point is never moved beyond the story boundaries.
	///
	/// \if UNICODE
	///   \sa MoveStartByUnit, MoveEndByUnit, MoveToUnitIndex, MoveStartToStartOfUnit, MoveEndToEndOfUnit,
	///       RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa MoveStartByUnit, MoveEndByUnit, MoveToUnitIndex, MoveStartToStartOfUnit, MoveEndToEndOfUnit,
	///       RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE MoveInsertionPointByUnit(UnitConstants unit, LONG count, LONG* pDelta);
	/// \brief <em>Moves the insertion point to the position of the first character found that is in the specified character set</em>
	///
	/// Moves the insertion point forward or backward, to the position of the first character found that is
	/// in the set of characters specified by \c pCharacterSet, provided that the character is found within
	/// \c maximumCharactersToPass characters of the range's boundaries.\n
	/// If no character from the set specified by \c pCharacterSet is found within \c maximumCharactersToPass
	/// positions of the range's boundaries, the range is left unchanged.
	///
	/// \param[in] pCharacterSet The character set to use in the match. This could be an explicit string of
	///            characters or a character-set index. For instance passing "0123456789" has the same
	///            effect as passing 4 (\c C1_DIGIT): The method finds the first digit.\n
	///            For more information see <a href="https://msdn.microsoft.com/en-us/library/bb787724.aspx#tom_character_match_sets">Character Match Sets</a>.
	/// \param[in] maximumCharactersToPass The maximum number of characters to move past. If
	///            \c maximumCharactersToPass is greater than zero, the search moves forward starting at the
	///            text range's end position, toward the end of the story; if it is less than zero, the
	///            search moves backward starting at the text range's start position.
	/// \param[out] pDelta Receives the actual number of characters that the insertion point moved past, plus
	///             1 for a match.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MoveStartUntil, MoveEndUntil, MoveInsertionPointWhile, MoveInsertionPointByUnit
	virtual HRESULT STDMETHODCALLTYPE MoveInsertionPointUntil(VARIANT* pCharacterSet, LONG maximumCharactersToPass = tomForward, LONG* pDelta = NULL);
	/// \brief <em>Moves the insertion point to the position of the first character found that is not in the specified character set</em>
	///
	/// Moves the insertion point forward or backward, to the position of the first character found that is
	/// not in the set of characters specified by \c pCharacterSet, or by \c maximumCharactersToPass
	/// characters, whichever is less.
	///
	/// \param[in] pCharacterSet The character set to use in the match. This could be an explicit string of
	///            characters or a character-set index. For instance passing "0123456789" has the same
	///            effect as passing 4 (\c C1_DIGIT): The method finds the first digit.\n
	///            For more information see <a href="https://msdn.microsoft.com/en-us/library/bb787724.aspx#tom_character_match_sets">Character Match Sets</a>.
	/// \param[in] maximumCharactersToPass The maximum number of characters to move past. If
	///            \c maximumCharactersToPass is greater than zero, the search moves forward starting at the
	///            text range's end position, toward the end of the story; if it is less than zero, the
	///            search moves backward starting at the text range's start position.
	/// \param[out] pDelta Receives the actual number of characters that the insertion point moved past.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MoveStartWhile, MoveEndWhile, MoveInsertionPointUntil, MoveInsertionPointByUnit
	virtual HRESULT STDMETHODCALLTYPE MoveInsertionPointWhile(VARIANT* pCharacterSet, LONG maximumCharactersToPass = tomForward, LONG* pDelta = NULL);
	/// \brief <em>Moves the text range's start position by the specified number of the specified units</em>
	///
	/// Moves the text range's start position forward or backward by the specified number of units, e.g. by 2
	/// words backward. If the new start position follows the old end position, the range is collapsed to
	/// an insertion point at the new start position.
	///
	/// \param[in] unit The unit by which to move. Any of the values defined by the \c UnitConstants
	///            enumeration is valid.
	/// \param[in] count The number of units by which to move. If \c count is greater than zero, motion
	///            is forward, toward the end of the story; if it is less than zero, motion is backward.
	/// \param[out] pDelta Receives the actual number of units that the range's start position moved past.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa MoveEndByUnit, MoveInsertionPointByUnit, MoveStartUntil, MoveStartWhile,
	///       RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa MoveEndByUnit, MoveInsertionPointByUnit, MoveStartUntil, MoveStartWhile,
	///       RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE MoveStartByUnit(UnitConstants unit, LONG count, LONG* pDelta);
	/// \brief <em>Moves the text range's start to the specified screen coordinates</em>
	///
	/// Moves the start position of the text range to the specified point or the nearest point with
	/// selectable text.
	///
	/// \param[in] x The x-coordinate (in twips) to which to move the text range's start position, relative
	///            to the screen's upper-left corner.
	/// \param[in] y The y-coordinate (in twips) to which to move the text range's start position, relative
	///            to the screen's upper-left corner.
	/// \param[in] doNotMoveEnd If \c VARIANT_TRUE, only the text range's start position is moved to the
	///            specified point, while the range's end position is not moved. If \c VARIANT_FALSE, the
	///            text range is collapsed to the specified position, creating an insertion point.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MoveEndToPosition, GetStartPosition
	virtual HRESULT STDMETHODCALLTYPE MoveStartToPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL doNotMoveEnd);
	/// \brief <em>Moves the text range's start to the start of the specified overlapping unit</em>
	///
	/// Moves the start position of the text range to the start of the first overlapping specified unit in
	/// the text range.
	///
	/// \param[in] unit The unit to which to move. Any of the values defined by the \c UnitConstants
	///            enumeration is valid.
	/// \param[in] doNotMoveEnd If \c VARIANT_TRUE, only the text range's start position is moved to the
	///            start of the specified unit, while the range's end position is not moved. If
	///            \c VARIANT_FALSE, the text range is collapsed to the start of the specified unit, creating
	///            an insertion point.
	/// \param[out] pDelta Receives the actual number of units that the range's start position has been moved
	///             past.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa MoveEndToEndOfUnit, ExpandToUnit, MoveToUnitIndex, MoveInsertionPointByUnit,
	///       RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa MoveEndToEndOfUnit, ExpandToUnit, MoveToUnitIndex, MoveInsertionPointByUnit,
	///       RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE MoveStartToStartOfUnit(UnitConstants unit, VARIANT_BOOL doNotMoveEnd, LONG* pDelta);
	/// \brief <em>Moves the text range's start position to the position of the first character found that is in the specified character set</em>
	///
	/// Moves the text range's start position forward or backward, to the position of the first character
	/// found that is in the set of characters specified by \c pCharacterSet, provided that the character is
	/// found within \c maximumCharactersToPass characters of the range's start.\n
	/// If the new start position follows the old end position, the range is collapsed to an insertion point
	/// at the new start position. If no character from the set specified by \c pCharacterSet is found within
	/// \c maximumCharactersToPass positions of the range's start, the range is left unchanged.
	///
	/// \param[in] pCharacterSet The character set to use in the match. This could be an explicit string of
	///            characters or a character-set index. For instance passing "0123456789" has the same
	///            effect as passing 4 (\c C1_DIGIT): The method finds the first digit.\n
	///            For more information see <a href="https://msdn.microsoft.com/en-us/library/bb787724.aspx#tom_character_match_sets">Character Match Sets</a>.
	/// \param[in] maximumCharactersToPass The maximum number of characters to move past. If
	///            \c maximumCharactersToPass is greater than zero, the search moves forward, toward the
	///            end of the story; if it is less than zero, the search moves backward.
	/// \param[out] pDelta Receives the actual number of characters that the range's start position moved
	///             past, plus 1 for a match.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MoveEndUntil, MoveInsertionPointUntil, MoveStartWhile, MoveStartByUnit
	virtual HRESULT STDMETHODCALLTYPE MoveStartUntil(VARIANT* pCharacterSet, LONG maximumCharactersToPass = tomForward, LONG* pDelta = NULL);
	/// \brief <em>Moves the text range's start position to the position of the first character found that is not in the specified character set</em>
	///
	/// Moves the text range's start position forward or backward, to the position of the first character
	/// found that is not in the set of characters specified by \c pCharacterSet, or by
	/// \c maximumCharactersToPass characters, whichever is less.\n
	/// If the new start position follows the old end position, the range is collapsed to an insertion
	/// point at the new start position.
	///
	/// \param[in] pCharacterSet The character set to use in the match. This could be an explicit string of
	///            characters or a character-set index. For instance passing "0123456789" has the same
	///            effect as passing 4 (\c C1_DIGIT): The method finds the first digit.\n
	///            For more information see <a href="https://msdn.microsoft.com/en-us/library/bb787724.aspx#tom_character_match_sets">Character Match Sets</a>.
	/// \param[in] maximumCharactersToPass The maximum number of characters to move past. If
	///            \c maximumCharactersToPass is greater than zero, the search moves forward, toward the
	///            end of the story; if it is less than zero, the search moves backward.
	/// \param[out] pDelta Receives the actual number of characters that the range's start position moved
	///             past.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MoveEndWhile, MoveInsertionPointWhile, MoveStartUntil, MoveStartByUnit
	virtual HRESULT STDMETHODCALLTYPE MoveStartWhile(VARIANT* pCharacterSet, LONG maximumCharactersToPass = tomForward, LONG* pDelta = NULL);
	/// \brief <em>Changes the range to the x-th unit of the current story</em>
	///
	/// Relocates the text range to the specified unit that has the specified index number, e.g. to the
	/// third word from the start of the story, or to the second word from the end of the story.
	///
	/// \param[in] unit The unit used to index the range. Any of the values defined by the \c UnitConstants
	///            enumeration is valid.
	/// \param[in] index The one-based index of the unit. If \c index is greater than zero, the numbering
	///            of units begins at the start of the story and proceeds forward; if it is less than zero,
	///            the numbering begins at the end of the story and proceeds backward.
	/// \param[in] setRangeToEntireUnit If \c VARIANT_TRUE, the range is set to the entire unit; otherwise it
	///            is collapsed to an insertion point at the start position of the unit.
	/// \param[out] pSucceeded \c VARIANT_TRUE if the text range has been relocated successfully; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The text range is never moved beyond the story boundaries.
	///
	/// \if UNICODE
	///   \sa MoveInsertionPointByUnit, MoveStartToStartOfUnit, MoveEndToEndOfUnit, get_UnitIndex,
	///       RichTextBoxCtlLibU::UnitConstants
	/// \else
	///   \sa MoveInsertionPointByUnit, MoveStartToStartOfUnit, MoveEndToEndOfUnit, get_UnitIndex,
	///       RichTextBoxCtlLibA::UnitConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE MoveToUnitIndex(UnitConstants unit, LONG index, VARIANT_BOOL setRangeToEntireUnit = VARIANT_FALSE, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Replaces the content of the text range with the content of the clipboard</em>
	///
	/// Pastes text from the clipboard or a specific data object, replacing the text range's current content.
	///
	/// \param[in] pOLEDataObject The \c OLEDataObject object from which to paste. If this parameter is not
	///            specified or does not refer to an object, the text will be pasted from the clipboard.
	/// \param[in] formatID An integer value specifying the format in which to paste the data. Valid values
	///            are those defined by VB's \c ClipBoardConstants enumeration, but also any other format
	///            that has been registered using the \c RegisterClipboardFormat API function.\n
	///            If set to \c 0, the format will be chosen automatically.
	/// \param[out] pSucceeded \c VARIANT_TRUE on success; otherwise \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CanPaste, Copy, Cut, get_Text, TargetOLEDataObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>
	virtual HRESULT STDMETHODCALLTYPE Paste(IOLEDataObject* pOLEDataObject = NULL, LONG formatID = 0, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Inserts a table</em>
	///
	/// Replaces the text range with a table.
	///
	/// \param[in] columnCount The number of columns that the table will have. A maximum of 63 columns is
	///            supported.
	/// \param[in] rowCount The number of rows that the table will have.
	/// \param[in] allowIndividualCellStyles If set to \c VARIANT_TRUE, each table cell can be styled
	///            individually; otherwise all cells in a row will have the same style. This refers to cell
	///            properties like margins, vertical alignments etc. Individual font styling (e.g. colors,
	///            bold, italics) always is possible, regardless of this setting.
	/// \param[in] borderSize The width (in twips) of the cell borders. If set to -1, the width of 1 pixel
	///            will be used.
	/// \param[in] horizontalRowAlignment The horizontal alignment of the table rows. The following of the
	///            values defined by the \c HAlignmentConstants enumeration are valid: \c halLeft,
	///            \c halCenter, \c halRight.
	/// \param[in] verticalCellAlignment The vertical alignment of the cells' content within the table cell.
	///            Any of the values defined by the \c VAlignmentConstants enumeration is valid.
	/// \param[in] rowIndent The size (in twips) by which the table rows are indented horizontally. The
	///            indent is measured from the control's left border.
	/// \param[in] columnWidth The width (in twips) of each column. If set to -1, the columns are sized
	///            evenly over the available space.
	/// \param[in] horizontalCellMargin The size (in twips) of the left and right margin in each cell. If set
	///            to -1, the default margin is used.
	/// \param[in] rowHeight The height (in twips) of each row. If set to -1, the default height is used.
	/// \param[out] ppAddedTable Receives the added table's \c IRichTable implementation.
	///
	/// \if UNICODE
	///   \sa RichTextBox::get_AllowTableInsertion, RichTextBox::get_DisplayZeroWidthTableCellBorders, Table,
	///       RichTextBoxCtlLibU::HAlignmentConstants, RichTextBoxCtlLibU::VAlignmentConstants
	/// \else
	///   \sa RichTextBox::get_AllowTableInsertion, RichTextBox::get_DisplayZeroWidthTableCellBorders, Table,
	///       RichTextBoxCtlLibA::HAlignmentConstants, RichTextBoxCtlLibA::VAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE ReplaceWithTable(LONG columnCount, LONG rowCount, VARIANT_BOOL allowIndividualCellStyles = VARIANT_TRUE, SHORT borderSize = -1, HAlignmentConstants horizontalRowAlignment = halLeft, VAlignmentConstants verticalCellAlignment = valCenter, LONG rowIndent = -1, LONG columnWidth = -1, LONG horizontalCellMargin = -1, LONG rowHeight = -1, IRichTable** ppAddedTable = NULL);
	/// \brief <em>Scrolls the text range into view</em>
	///
	/// \param[in] flags A value controlling which part of the range is scrolled into view. Some combinations
	///            of the values defined by the \c ScrollIntoViewConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::ScrollIntoViewConstants
	/// \else
	///   \sa RichTextBoxCtlLibA::ScrollIntoViewConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE ScrollIntoView(ScrollIntoViewConstants flags = sivScrollRangeStartToTop);
	/// \brief <em>Selects the current range of text</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RichTextBox::get_SelectedTextRange
	virtual HRESULT STDMETHODCALLTYPE Select(void);
	/// \brief <em>Sets the text range's start and end</em>
	///
	/// Sets the text range's start and end. The range's start position will be set to the minimum of
	/// \c anchorEnd and \c activeEnd; the range's end position will be set to the maximum.
	///
	/// \param[in] anchorEnd The character position of the anchor end of the text range.
	/// \param[in] activeEnd The character position of the active end of the text range.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RangeStart, RangeEnd
	virtual HRESULT STDMETHODCALLTYPE SetStartAndEnd(LONG anchorEnd, LONG activeEnd);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given text range</em>
	///
	/// Attaches this object to a given text range, so that the text range's properties can be retrieved and
	/// set using this object's methods.
	///
	/// \param[in] pTextRange The \c ITextRange object to attach to.
	///
	/// \sa Detach,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	void Attach(ITextRange* pTextRange);
	/// \brief <em>Detaches this object from a text range</em>
	///
	/// Detaches this object from the text range it currently wraps, so that it doesn't wrap any text range
	/// anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this button</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerRTB
	void SetOwner(__in_opt RichTextBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c RichTextBox object that owns this text range</em>
		///
		/// \sa SetOwner
		RichTextBox* pOwnerRTB;
		/// \brief <em>The \c ITextRange object corresponding to the text range wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pTextRange;

		Properties()
		{
			pOwnerRTB = NULL;
			pTextRange = NULL;
		}

		~Properties();

		/// \brief <em>Retrieves the owning rich text box' window handle</em>
		///
		/// \return The window handle of the rich text box that contains this text range.
		///
		/// \sa pOwnerRTB
		HWND GetRTBHWnd(void);
	} properties;
};     // TextRange

OBJECT_ENTRY_AUTO(__uuidof(TextRange), TextRange)