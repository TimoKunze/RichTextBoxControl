//////////////////////////////////////////////////////////////////////
/// \class TextParagraph
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a text range paragraph's styling, e.g. alignment and line spacing</em>
///
/// This class is a wrapper around the styling of a text range's paragraph.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTextParagraph, RichTextBox, TextRange
/// \else
///   \sa RichTextBoxCtlLibA::IRichTextParagraph, RichTextBox, TextRange
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITextParagraphEvents_CP.h"
#include "helpers.h"
#include "TextParagraphTabs.h"


class ATL_NO_VTABLE TextParagraph : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TextParagraph, &CLSID_TextParagraph>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TextParagraph>,
    public Proxy_IRichTextParagraphEvents<TextParagraph>, 
    #ifdef UNICODE
    	public IDispatchImpl<IRichTextParagraph, &IID_IRichTextParagraph, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTextParagraph, &IID_IRichTextParagraph, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TEXTPARAGRAPH)

		BEGIN_COM_MAP(TextParagraph)
			COM_INTERFACE_ENTRY(IRichTextParagraph)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextPara), 0, QueryITextParaInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextPara2), 0, QueryITextParaInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TextParagraph)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTextParagraphEvents))
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

	/// \brief <em>Will be called if this class is \c QueryInterface()'d for \c ITextPara or \c ITextPara2</em>
	///
	/// This method will be called if the class's \c QueryInterface() method is called with
	/// \c IID_ITextPara or \c IID_ITextPara2. We forward the call to the wrapped \c ITextPara or
	/// \c ITextPara2 implementation.
	///
	/// \param[in] pThis The instance of this class, that the interface is queried from.
	/// \param[in] queriedInterface Should be \c IID_ITextPara or \c IID_ITextPara2.
	/// \param[out] ppImplementation Receives the wrapped object's \c ITextPara or \c ITextPara2
	///             implementation.
	/// \param[in] cookie A \c DWORD value specified in the COM interface map.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774056.aspx">ITextPara</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/hh768585.aspx">ITextPara2</a>
	static HRESULT CALLBACK QueryITextParaInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTextParagraph
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c CanChange property</em>
	///
	/// Retrieves whether the paragraph's formatting can be changed. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the formatting can be
	/// changed; if set to \c bpvFalse, it cannot be changed; and if set to \c bpvUndefined, parts of the
	/// paragraph can be changed and others can't.
	/// The formatting cannot be changed, if the document is read-only or any part of the paragraph is
	/// protected.
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
	/// \brief <em>Retrieves the current setting of the \c FirstLineIndent property</em>
	///
	/// Retrieves the indent (in points) of the first line in the paragraph, relative to the control's left
	/// margin.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FirstLineIndent, get_LeftIndent, get_RightIndent
	virtual HRESULT STDMETHODCALLTYPE get_FirstLineIndent(FLOAT* pValue);
	/// \brief <em>Sets the \c FirstLineIndent property</em>
	///
	/// Sets the indent (in points) of the first line in the paragraph, relative to the control's left
	/// margin.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstLineIndent, put_LeftIndent, put_RightIndent
	virtual HRESULT STDMETHODCALLTYPE put_FirstLineIndent(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c HAlignment property</em>
	///
	/// Retrieves the horizontal alignment of the paragraph's content. Any of the values defined by the
	/// \c HAlignmentConstants enumeration is valid. If the text range uses multiple alignments, the
	/// alignment will be reported as \c halUndefined.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_HAlignment, TextRange::get_Text, get_ListNumberingHAlignment,
	///       TextParagraphTab::get_Alignment, RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa put_HAlignment, TextRange::get_Text, get_ListNumberingHAlignment,
	///       TextParagraphTab::get_Alignment, RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HAlignment(HAlignmentConstants* pValue);
	/// \brief <em>Sets the \c HAlignment property</em>
	///
	/// Sets the horizontal alignment of the paragraph's content. Any of the values defined by the
	/// \c HAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_HAlignment, TextRange::put_Text, RichTextBox::put_HAlignment, put_ListNumberingHAlignment,
	///       TextParagraphTab::put_Alignment, RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa get_HAlignment, TextRange::put_Text, RichTextBox::put_HAlignment, put_ListNumberingHAlignment,
	///       TextParagraphTab::put_Alignment, RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Hyphenation property</em>
	///
	/// Retrieves whether automatic hyphenation is applied to the paragraphs in the text range. Any of the
	/// values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue,
	/// automatic hyphenation is enabled; if set to \c bpvFalse, it is disabled; and if set to
	/// \c bpvUndefined, some paragraphs in the text range use automatic hyphenation and others don't.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Hyphenation, RichTextBox::get_HyphenationFunction,
	///       RichTextBox::get_HyphenationWordWrapZoneWidth,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Hyphenation, RichTextBox::get_HyphenationFunction,
	///       RichTextBox::get_HyphenationWordWrapZoneWidth,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Hyphenation(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Hyphenation property</em>
	///
	/// Sets whether automatic hyphenation is applied to the paragraphs in the text range. Any of the
	/// values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue,
	/// automatic hyphenation is enabled; if set to \c bpvFalse, it is disabled; and if set to
	/// \c bpvUndefined, some paragraphs in the text range use automatic hyphenation and others don't.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Hyphenation, RichTextBox::put_HyphenationFunction,
	///       RichTextBox::put_HyphenationWordWrapZoneWidth,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Hyphenation, RichTextBox::put_HyphenationFunction,
	///       RichTextBox::put_HyphenationWordWrapZoneWidth,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Hyphenation(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c InsertPageBreakBeforeParagraph property</em>
	///
	/// Retrieves whether the paragraph is forced to start on a new page by inserting a page break
	/// before the paragraph. Any of the values defined by the \c BooleanPropertyValueConstants enumeration
	/// is valid. If set to \c bpvTrue, the paragraph is forced to start on a new page; if set to
	/// \c bpvFalse, the paragraph does not need to start on a new page; and if set to \c bpvUndefined,
	/// some paragraphs in the text range are forced to start on a new page and others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_InsertPageBreakBeforeParagraph, get_KeepTogether, get_KeepWithNext,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_InsertPageBreakBeforeParagraph, get_KeepTogether, get_KeepWithNext,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_InsertPageBreakBeforeParagraph(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c InsertPageBreakBeforeParagraph property</em>
	///
	/// Sets whether the paragraph is forced to start on a new page by inserting a page break
	/// before the paragraph. Any of the values defined by the \c BooleanPropertyValueConstants enumeration
	/// is valid. If set to \c bpvTrue, the paragraph is forced to start on a new page; if set to
	/// \c bpvFalse, the paragraph does not need to start on a new page; and if set to \c bpvUndefined,
	/// some paragraphs in the text range are forced to start on a new page and others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_InsertPageBreakBeforeParagraph, put_KeepTogether, put_KeepWithNext,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_InsertPageBreakBeforeParagraph, put_KeepTogether, put_KeepWithNext,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_InsertPageBreakBeforeParagraph(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c KeepTogether property</em>
	///
	/// Retrieves whether page breaks are allowed within the paragraph. Any of the values defined
	/// by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvFalse, page breaks
	/// are allowed; if set to \c bpvTrue, they are forbidden; and if set to \c bpvUndefined, some
	/// paragraphs in the text range allow page breaks and others don't.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_KeepTogether, get_KeepWithNext, get_InsertPageBreakBeforeParagraph,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_KeepTogether, get_KeepWithNext, get_InsertPageBreakBeforeParagraph,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_KeepTogether(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c KeepTogether property</em>
	///
	/// Sets whether page breaks are allowed within the paragraph. Any of the values defined
	/// by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvFalse, page breaks
	/// are allowed; if set to \c bpvTrue, they are forbidden; and if set to \c bpvUndefined, some
	/// paragraphs in the text range allow page breaks and others don't.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_KeepTogether, put_KeepWithNext, put_InsertPageBreakBeforeParagraph,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_KeepTogether, put_KeepWithNext, put_InsertPageBreakBeforeParagraph,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_KeepTogether(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c KeepWithNext property</em>
	///
	/// Retrieves whether page breaks are allowed between paragraphs in the text range. Any of the
	/// values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvFalse,
	/// page breaks are allowed; if set to \c bpvTrue, they are forbidden; and if set to \c bpvUndefined,
	/// settings in the text range are mixed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_KeepWithNext, get_KeepTogether, get_InsertPageBreakBeforeParagraph,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_KeepWithNext, get_KeepTogether, get_InsertPageBreakBeforeParagraph,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_KeepWithNext(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c KeepWithNext property</em>
	///
	/// Sets whether page breaks are allowed between paragraphs in the text range. Any of the
	/// values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvFalse,
	/// page breaks are allowed; if set to \c bpvTrue, they are forbidden; and if set to \c bpvUndefined,
	/// settings in the text range are mixed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_KeepWithNext, put_KeepTogether, put_InsertPageBreakBeforeParagraph,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_KeepWithNext, put_KeepTogether, put_InsertPageBreakBeforeParagraph,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_KeepWithNext(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c LeftIndent property</em>
	///
	/// Retrieves the indent (in points) of each line in the paragraph, relative to the control's left
	/// margin.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_LeftIndent, get_FirstLineIndent, get_RightIndent, RichTextBox::get_LeftMargin
	virtual HRESULT STDMETHODCALLTYPE get_LeftIndent(FLOAT* pValue);
	/// \brief <em>Sets the \c LeftIndent property</em>
	///
	/// Sets the indent (in points) of each line in the paragraph, relative to the control's left
	/// margin.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LeftIndent, put_FirstLineIndent, put_RightIndent, RichTextBox::put_LeftMargin
	virtual HRESULT STDMETHODCALLTYPE put_LeftIndent(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c LineSpacing property</em>
	///
	/// Retrieves the value by which lines are spaced within the paragraph. The meaning of this
	/// value depends on the value of the \c LineSpacingRule property.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_LineSpacing, get_LineSpacingRule, get_SpaceBefore, get_SpaceAfter
	virtual HRESULT STDMETHODCALLTYPE get_LineSpacing(FLOAT* pValue);
	/// \brief <em>Sets the \c LineSpacing property</em>
	///
	/// Sets the value by which lines are spaced within the paragraph. The meaning of this
	/// value depends on the value of the \c LineSpacingRule property.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_LineSpacing, put_LineSpacingRule, put_SpaceBefore, put_SpaceAfter
	virtual HRESULT STDMETHODCALLTYPE put_LineSpacing(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c LineSpacingRule property</em>
	///
	/// Retrieves the rule by which lines are spaced within the paragraph. Any of the values defined by the
	/// \c LineSpacingRuleConstants enumeration is valid. If the text range uses multiple line spacing rules,
	/// the rule will be reported as \c lsrUndefined.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_LineSpacingRule, get_LineSpacing, get_SpaceBefore, get_SpaceAfter,
	///       RichTextBox::get_AdjustLineHeightForEastAsianLanguages,
	///       RichTextBoxCtlLibU::LineSpacingRuleConstants
	/// \else
	///   \sa put_LineSpacingRule, get_LineSpacing, get_SpaceBefore, get_SpaceAfter,
	///       RichTextBox::get_AdjustLineHeightForEastAsianLanguages,
	///       RichTextBoxCtlLibA::LineSpacingRuleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_LineSpacingRule(LineSpacingRuleConstants* pValue);
	/// \brief <em>Sets the \c LineSpacingRule property</em>
	///
	/// Sets the rule by which lines are spaced within the paragraph. Any of the values defined by the
	/// \c LineSpacingRuleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_LineSpacingRule, put_LineSpacing, put_SpaceBefore, put_SpaceAfter,
	///       RichTextBox::put_AdjustLineHeightForEastAsianLanguages,
	///       RichTextBoxCtlLibU::LineSpacingRuleConstants
	/// \else
	///   \sa get_LineSpacingRule, put_LineSpacing, put_SpaceBefore, put_SpaceAfter,
	///       RichTextBox::put_AdjustLineHeightForEastAsianLanguages,
	///       RichTextBoxCtlLibA::LineSpacingRuleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_LineSpacingRule(LineSpacingRuleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ListLevelIndex property</em>
	///
	/// Retrieves a zero-based index specifying the list level of the paragraph. If set to 0, the paragraph
	/// is not part of a list. If set to 1, the paragraph is part of the outermost list. If set to 2, the
	/// paragraph is part of the second-level (nested) list, which is nested under a level 1 list item. If
	/// set to 3, the paragraph is part of the third-level (nested) list, which is nested under a level 2
	/// list item and so forth.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListLevelIndex, get_ListNumberingStyle, get_ListNumberingStart, get_ListNumberingHAlignment,
	///     get_ListNumberingTabStop
	virtual HRESULT STDMETHODCALLTYPE get_ListLevelIndex(LONG* pValue);
	/// \brief <em>Sets the \c ListLevelIndex property</em>
	///
	/// Sets a zero-based index specifying the list level of the paragraph. If set to 0, the paragraph
	/// is not part of a list. If set to 1, the paragraph is part of the outermost list. If set to 2, the
	/// paragraph is part of the second-level (nested) list, which is nested under a level 1 list item. If
	/// set to 3, the paragraph is part of the third-level (nested) list, which is nested under a level 2
	/// list item and so forth.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListLevelIndex, put_ListNumberingStyle, put_ListNumberingStart, put_ListNumberingHAlignment,
	///     put_ListNumberingTabStop
	virtual HRESULT STDMETHODCALLTYPE put_ListLevelIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ListNumberingHAlignment property</em>
	///
	/// Retrieves the horizontal alignment of the paragraph's list item numbers. Some of the values
	/// defined by the \c HAlignmentConstants enumeration is valid. If the text range uses multiple
	/// alignments, the alignment will be reported as \c halUndefined.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Some alignment options require a left indent.
	///
	/// \if UNICODE
	///   \sa put_ListNumberingHAlignment, get_LeftIndent, get_ListNumberingTabStop, get_ListLevelIndex,
	///       get_ListNumberingStyle, get_HAlignment, RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa put_ListNumberingHAlignment, get_LeftIndent, get_ListNumberingTabStop, get_ListLevelIndex,
	///       get_ListNumberingStyle, get_HAlignment, RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ListNumberingHAlignment(HAlignmentConstants* pValue);
	/// \brief <em>Sets the \c ListNumberingHAlignment property</em>
	///
	/// Sets the horizontal alignment of the paragraph's list item numbers. Some of the values
	/// defined by the \c HAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Some alignment options require a left indent.
	///
	/// \if UNICODE
	///   \sa get_ListNumberingHAlignment, put_LeftIndent, put_ListNumberingTabStop, put_ListLevelIndex,
	///       put_ListNumberingStyle, put_HAlignment, RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa get_ListNumberingHAlignment, put_LeftIndent, put_ListNumberingTabStop, put_ListLevelIndex,
	///       put_ListNumberingStyle, put_HAlignment, RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ListNumberingHAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ListNumberingStart property</em>
	///
	/// Retrieves the number of the first list item in the current list. If the \c ListNumberingStyle
	/// property is set to \c lnsCustomSequence, this property specifies the unicode code point that is used
	/// as the first numbering sign.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListNumberingStart, get_ListNumberingStyle, get_ListLevelIndex, get_ListNumberingHAlignment
	virtual HRESULT STDMETHODCALLTYPE get_ListNumberingStart(LONG* pValue);
	/// \brief <em>Sets the \c ListNumberingStart property</em>
	///
	/// Sets the number of the first list item in the current list. If the \c ListNumberingStyle
	/// property is set to \c lnsCustomSequence, this property specifies the unicode code point that is used
	/// as the first numbering sign.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListNumberingStart, put_ListNumberingStyle, put_ListLevelIndex, put_ListNumberingHAlignment
	virtual HRESULT STDMETHODCALLTYPE put_ListNumberingStart(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ListNumberingStyle property</em>
	///
	/// Retrieves the kind of numbering of list items in the paragraph. Any of the values defined by the
	/// \c ListNumberingStyleConstants enumeration is valid. If the text range uses multiple list numbering
	/// types, the type will be reported as \c lnsUndefined.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ListNumberingStyle, get_ListLevelIndex, get_ListNumberingStart,
	///       get_ListNumberingHAlignment, RichTextBoxCtlLibU::ListNumberingStyleConstants
	/// \else
	///   \sa put_ListNumberingStyle, get_ListLevelIndex, get_ListNumberingStart,
	///       get_ListNumberingHAlignment, RichTextBoxCtlLibA::ListNumberingStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ListNumberingStyle(ListNumberingStyleConstants* pValue);
	/// \brief <em>Sets the \c ListNumberingStyle property</em>
	///
	/// Sets the kind of numbering of list items in the paragraph. Any of the values defined by the
	/// \c ListNumberingStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ListNumberingStyle, put_ListLevelIndex, put_ListNumberingStart,
	///       put_ListNumberingHAlignment, RichTextBoxCtlLibU::ListNumberingStyleConstants
	/// \else
	///   \sa get_ListNumberingStyle, put_ListLevelIndex, put_ListNumberingStart,
	///       put_ListNumberingHAlignment, RichTextBoxCtlLibA::ListNumberingStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ListNumberingStyle(ListNumberingStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ListNumberingTabStop property</em>
	///
	/// Retrieves the distance (in pixels) between the right first list item's number and the first list
	/// item's text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ListNumberingTabStop, get_ListNumberingHAlignment, get_ListLevelIndex, get_ListNumberingStyle
	virtual HRESULT STDMETHODCALLTYPE get_ListNumberingTabStop(FLOAT* pValue);
	/// \brief <em>Sets the \c ListNumberingTabStop property</em>
	///
	/// Sets the distance (in pixels) between the right first list item's number and the first list
	/// item's text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ListNumberingTabStop, put_ListNumberingHAlignment, put_ListLevelIndex, put_ListNumberingStyle
	virtual HRESULT STDMETHODCALLTYPE put_ListNumberingTabStop(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c Numbered property</em>
	///
	/// Retrieves whether the paragraphs in the text range are numbered and the number is displayed in the
	/// first line of each paragraph. Any of the values defined by the \c BooleanPropertyValueConstants
	/// enumeration is valid. If set to \c bpvTrue, the paragraphs are numbered; if set to \c bpvFalse, they
	/// are not numbered; and if set to \c bpvUndefined, some paragraphs in the text range are numbered and
	/// others are not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Numbered, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_Numbered, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Numbered(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c Numbered property</em>
	///
	/// Sets whether the paragraphs in the text range are numbered and the number is displayed in the
	/// first line of each paragraph. Any of the values defined by the \c BooleanPropertyValueConstants
	/// enumeration is valid. If set to \c bpvTrue, the paragraphs are numbered; if set to \c bpvFalse, they
	/// are not numbered; and if set to \c bpvUndefined, some paragraphs in the text range are numbered and
	/// others are not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Numbered, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_Numbered, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Numbered(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c PreventWidowsAndOrphans property</em>
	///
	/// Retrieves whether the control prevents widows and orphans within the paragraph. A widow is
	/// created when the last line of a paragraph is printed by itself at the top of a page. An orphan is
	/// when the first line of a paragraph is printed by itself at the bottom of a page.\n
	/// Any of the values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to
	/// \c bpvTrue, widows and orphans are prevented; if set to \c bpvFalse, they are allowed; and if set
	/// to \c bpvUndefined, some paragraphs in the text range prevent widows and orphans and others don't.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_PreventWidowsAndOrphans, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_PreventWidowsAndOrphans, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_PreventWidowsAndOrphans(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c PreventWidowsAndOrphans property</em>
	///
	/// Sets whether the control prevents widows and orphans within the paragraph. A widow is
	/// created when the last line of a paragraph is printed by itself at the top of a page. An orphan is
	/// when the first line of a paragraph is printed by itself at the bottom of a page.\n
	/// Any of the values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to
	/// \c bpvTrue, widows and orphans are prevented; if set to \c bpvFalse, they are allowed; and if set
	/// to \c bpvUndefined, some paragraphs in the text range prevent widows and orphans and others don't.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_PreventWidowsAndOrphans, RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_PreventWidowsAndOrphans, RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_PreventWidowsAndOrphans(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c RightIndent property</em>
	///
	/// Retrieves the indent (in points) of each line in the paragraph, relative to the control's right
	/// margin.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RightIndent, get_LeftIndent, get_FirstLineIndent, RichTextBox::get_RightMargin
	virtual HRESULT STDMETHODCALLTYPE get_RightIndent(FLOAT* pValue);
	/// \brief <em>Sets the \c RightIndent property</em>
	///
	/// Sets the indent (in points) of each line in the paragraph, relative to the control's right
	/// margin.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RightIndent, put_LeftIndent, put_FirstLineIndent, RichTextBox::put_RightMargin
	virtual HRESULT STDMETHODCALLTYPE put_RightIndent(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether the paragraph uses right-to-left layout.\n
	/// Any of the values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to
	/// \c bpvTrue, the paragraph's content uses right-to-left layout; if set to \c bpvFalse, it uses
	/// left-to-right layout; and if set to \c bpvUndefined, some parts in the text range use right-to-left
	/// layout and others use left-to-right layout.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa put_RightToLeft, RichTextBox::get_RightToLeft, RichTextBox::get_RightToLeftMathZones,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa put_RightToLeft, RichTextBox::get_RightToLeft, RichTextBox::get_RightToLeftMathZones,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Sets whether the paragraph uses right-to-left layout.\n
	/// Any of the values defined by the \c BooleanPropertyValueConstants enumeration is valid. If set to
	/// \c bpvTrue, the paragraph's content uses right-to-left layout; if set to \c bpvFalse, it uses
	/// left-to-right layout; and if set to \c bpvUndefined, some parts in the text range use right-to-left
	/// layout and others use left-to-right layout.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, RichTextBox::put_RightToLeft, RichTextBox::put_RightToLeftMathZones,
	///       RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa get_RightToLeft, RichTextBox::put_RightToLeft, RichTextBox::put_RightToLeftMathZones,
	///       RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(BooleanPropertyValueConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c SpaceAfter property</em>
	///
	/// Retrieves the amount of vertical space below the paragraph. If the text range uses multiple
	/// spacings, the vertical space will be reported as -9999999.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_SpaceAfter, get_SpaceBefore, get_LineSpacing,
	///     RichTextBox::get_AdjustLineHeightForEastAsianLanguages
	virtual HRESULT STDMETHODCALLTYPE get_SpaceAfter(FLOAT* pValue);
	/// \brief <em>Sets the \c SpaceAfter property</em>
	///
	/// Sets the amount of vertical space below the paragraph.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_SpaceAfter, put_SpaceBefore, put_LineSpacing,
	///     RichTextBox::put_AdjustLineHeightForEastAsianLanguages
	virtual HRESULT STDMETHODCALLTYPE put_SpaceAfter(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c SpaceBefore property</em>
	///
	/// Retrieves the amount of vertical space above the paragraph. If the text range uses multiple
	/// spacings, the vertical space will be reported as -9999999.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_SpaceBefore, get_SpaceAfter, get_LineSpacing,
	///     RichTextBox::get_AdjustLineHeightForEastAsianLanguages
	virtual HRESULT STDMETHODCALLTYPE get_SpaceBefore(FLOAT* pValue);
	/// \brief <em>Sets the \c SpaceBefore property</em>
	///
	/// Sets the amount of vertical space above the paragraph.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_SpaceBefore, put_SpaceAfter, put_LineSpacing,
	///     RichTextBox::put_AdjustLineHeightForEastAsianLanguages
	virtual HRESULT STDMETHODCALLTYPE put_SpaceBefore(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c StyleID property</em>
	///
	/// Retrieves the unique ID of the style assigned to the paragraph. If the text range uses multiple
	/// paragraph styles, the style will be reported as -9999999.\n
	/// This property can also be set to one of the built-in styles defined by the \c BuiltInStyleConstants
	/// enumeration.
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
	/// Sets the unique ID of the style assigned to the paragraph.\n
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
	/// \brief <em>Retrieves the current setting of the \c Tabs property</em>
	///
	/// Retrieves a collection object wrapping the paragraph's tabs.
	///
	/// \param[out] ppTabs Receives the collection object's \c IRichTextParagraphTabs implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TextParagraphTabs
	virtual HRESULT STDMETHODCALLTYPE get_Tabs(IRichTextParagraphTabs** ppTabs);

	/// \brief <em>Creates a new \c IRichTextParagraph instance with identical properties</em>
	///
	/// \param[out] ppClone Receives the new \c IRichTextParagraph instance.
	///
	/// \sa CopySettings, Equals
	virtual HRESULT STDMETHODCALLTYPE Clone(IRichTextParagraph** ppClone);
	/// \brief <em>Copies all settings from an other \c IRichTextParagraph instance</em>
	///
	/// \param[in] pSourceObject The \c IRichTextParagraph instance to copy from.
	///
	/// \sa Clone, Equals
	virtual HRESULT STDMETHODCALLTYPE CopySettings(IRichTextParagraph* pSourceObject);
	/// \brief <em>Retrieves whether two \c IRichTextParagraph instances have the same properties</em>
	///
	/// \param[in] pCompareAgainst The \c IRichTextParagraph instance to compare with.
	/// \param[out] pValue \c VARIANT_TRUE, if the objects have the same properties; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \sa Clone, CopySettings
	virtual HRESULT STDMETHODCALLTYPE Equals(IRichTextParagraph* pCompareAgainst, VARIANT_BOOL* pValue);
	/// \brief <em>Resets the text range's paragraph formatting</em>
	///
	/// Resets the text range's paragraph formatting, applying the specified rules.
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

	/// \brief <em>Attaches this object to a given text range paragraph</em>
	///
	/// Attaches this object to a given text range paragraph, so that the text range paragraph's properties
	/// can be retrieved and set using this object's methods.
	///
	/// \param[in] pTextParagraph The \c ITextPara object to attach to.
	///
	/// \sa Detach,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774056.aspx">ITextPara</a>
	void Attach(ITextPara* pTextParagraph);
	/// \brief <em>Detaches this object from a text range paragraph</em>
	///
	/// Detaches this object from the text range paragraph it currently wraps, so that it doesn't wrap any
	/// text range paragraph anymore.
	///
	/// \sa Attach
	void Detach(void);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ITextPara object corresponding to the text range paragraph wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774056.aspx">ITextPara</a>
		ITextPara* pTextParagraph;

		Properties()
		{
			pTextParagraph = NULL;
		}

		~Properties();
	} properties;
};     // TextParagraph

OBJECT_ENTRY_AUTO(__uuidof(TextParagraph), TextParagraph)