//////////////////////////////////////////////////////////////////////
/// \class TableRow
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a table row's styling, e.g. margins</em>
///
/// This class is a wrapper around the styling of a table row.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTableRow, RichTextBox, TextRange
/// \else
///   \sa RichTextBoxCtlLibA::IRichTableRow, RichTextBox, TextRange
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITableRowEvents_CP.h"
#include "helpers.h"
#include "TableCells.h"


class ATL_NO_VTABLE TableRow : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TableRow, &CLSID_TableRow>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TableRow>,
    public Proxy_IRichTableRowEvents<TableRow>, 
    #ifdef UNICODE
    	public IDispatchImpl<IRichTableRow, &IID_IRichTableRow, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTableRow, &IID_IRichTableRow, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TABLEROW)

		BEGIN_COM_MAP(TableRow)
			COM_INTERFACE_ENTRY(IRichTableRow)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange2), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRow), 0, QueryITextRowInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TableRow)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTableRowEvents))
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
	/// \brief <em>Will be called if this class is \c QueryInterface()'d for \c ITextRow</em>
	///
	/// This method will be called if the class's \c QueryInterface() method is called with
	/// \c IID_ITextRow. We forward the call to the wrapped \c ITextRow implementation.
	///
	/// \param[in] pThis The instance of this class, that the interface is queried from.
	/// \param[in] queriedInterface Should be \c IID_ITextRow.
	/// \param[out] ppImplementation Receives the wrapped object's \c ITextRow implementation.
	/// \param[in] cookie A \c DWORD value specified in the COM interface map.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hh768670.aspx">ITextRow</a>
	static HRESULT CALLBACK QueryITextRowInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTableRow
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AllowPageBreakWithinRow property</em>
	///
	/// Retrieves whether a page break is allowed within the table row. If set to \c VARIANT_TRUE, page
	/// breaks are allowed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa put_AllowPageBreakWithinRow, get_KeepTogetherWithNextRow
	virtual HRESULT STDMETHODCALLTYPE get_AllowPageBreakWithinRow(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowPageBreakWithinRow property</em>
	///
	/// Sets whether a page break is allowed within the table row. If set to \c VARIANT_TRUE, page
	/// breaks are allowed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa get_AllowPageBreakWithinRow, put_KeepTogetherWithNextRow
	virtual HRESULT STDMETHODCALLTYPE put_AllowPageBreakWithinRow(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CanChange property</em>
	///
	/// Retrieves whether the table row can be changed. Any of the values defined by the
	/// \c BooleanPropertyValueConstants enumeration is valid. If set to \c bpvTrue, the row can be
	/// changed; if set to \c bpvFalse, it cannot be changed; and if set to \c bpvUndefined, parts of the
	/// row can be changed and others can't.
	/// The row cannot be changed, if the document is read-only or any part of the row is protected.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::BooleanPropertyValueConstants
	/// \else
	///   \sa RichTextBoxCtlLibA::BooleanPropertyValueConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CanChange(BooleanPropertyValueConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c Cells property</em>
	///
	/// Retrieves a collection object wrapping the table row's cells.
	///
	/// \param[out] ppCells Receives the collection object's \c IRichTableCells implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TableCells
	virtual HRESULT STDMETHODCALLTYPE get_Cells(IRichTableCells** ppCells);
	/// \brief <em>Retrieves the current setting of the \c HAlignment property</em>
	///
	/// Retrieves the horizontal alignment of the table row. The following of the values defined by
	/// the \c HAlignmentConstants enumeration are valid: \c halLeft, \c halCenter, \c halRight.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa put_HAlignment, get_Indent, TableCell::get_Width, RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa put_HAlignment, get_Indent, TableCell::get_Width, RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HAlignment(HAlignmentConstants* pValue);
	/// \brief <em>Sets the \c HAlignment property</em>
	///
	/// Sets the horizontal alignment of the table row. The following of the values defined by
	/// the \c HAlignmentConstants enumeration are valid: \c halLeft, \c halCenter, \c halRight.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa get_HAlignment, put_Indent, TableCell::put_Width, RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa get_HAlignment, put_Indent, TableCell::put_Width, RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Height property</em>
	///
	/// Retrieves the table row's height (in twips).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa put_Height, TableCell::get_Width, TableCell::get_VAlignment
	virtual HRESULT STDMETHODCALLTYPE get_Height(LONG* pValue);
	/// \brief <em>Sets the \c Height property</em>
	///
	/// Sets the table row's height (in twips).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa get_Height, TableCell::put_Width, TableCell::put_VAlignment
	virtual HRESULT STDMETHODCALLTYPE put_Height(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c HorizontalCellMargin property</em>
	///
	/// Retrieves the size (in twips) of the left and right margin in each table cell.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa put_HorizontalCellMargin, TextParagraph::get_LeftIndent
	virtual HRESULT STDMETHODCALLTYPE get_HorizontalCellMargin(LONG* pValue);
	/// \brief <em>Sets the \c HorizontalCellMargin property</em>
	///
	/// Sets the size (in twips) of the left and right margin in each table cell.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa get_HorizontalCellMargin, TextParagraph::put_LeftIndent
	virtual HRESULT STDMETHODCALLTYPE put_HorizontalCellMargin(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Indent property</em>
	///
	/// Retrieves the size (in twips) by which the table row is indented horizontally. The
	/// indent is measured from the control's left border.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa put_Indent, get_HAlignment
	virtual HRESULT STDMETHODCALLTYPE get_Indent(LONG* pValue);
	/// \brief <em>Sets the \c Indent property</em>
	///
	/// Sets the size (in twips) by which the table row is indented horizontally. The
	/// indent is measured from the control's left border.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa get_Indent, put_HAlignment
	virtual HRESULT STDMETHODCALLTYPE put_Indent(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index identifying this table row.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::RowIdentifierTypeConstants
	/// \else
	///   \sa RichTextBoxCtlLibA::RowIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c KeepTogetherWithNextRow property</em>
	///
	/// Retrieves whether the table row is kept on the same page as the following row. If set to
	/// \c VARIANT_TRUE, the rows are kept together; otherwise a page break is allowed in between.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa put_KeepTogetherWithNextRow, get_AllowPageBreakWithinRow
	virtual HRESULT STDMETHODCALLTYPE get_KeepTogetherWithNextRow(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c KeepTogetherWithNextRow property</em>
	///
	/// Sets whether the table row is kept on the same page as the following row. If set to
	/// \c VARIANT_TRUE, the rows are kept together; otherwise a page break is allowed in between.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa get_KeepTogetherWithNextRow, put_AllowPageBreakWithinRow
	virtual HRESULT STDMETHODCALLTYPE put_KeepTogetherWithNextRow(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MasterCell property</em>
	///
	/// Retrieves the cell that all subsequent cells in the same row inherit their style from. If set to
	/// \c NULL, no cell in the row inherits its style from another cell.
	///
	/// \param[out] ppMasterCell Receives the master cell's \c IRichTableCell implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_MasterCell, TableCell::get_DefinesStyleForSubsequentCells
	virtual HRESULT STDMETHODCALLTYPE get_MasterCell(IRichTableCell** ppMasterCell);
	/// \brief <em>Sets the \c MasterCell property</em>
	///
	/// Sets the cell that all subsequent cells in the same row inherit their style from. If set to
	/// \c NULL, no cell in the row inherits its style from another cell.
	///
	/// \param[in] pNewMasterCell The new master cell's \c IRichTableCell implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MasterCell, TableCell::get_DefinesStyleForSubsequentCells
	virtual HRESULT STDMETHODCALLTYPE putref_MasterCell(IRichTableCell* pNewMasterCell);
	/// \brief <em>Retrieves the current setting of the \c ParentTable property</em>
	///
	/// Retrieves a \c Table object that wraps the table that the table row refers to.
	///
	/// \param[out] ppTable The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Table
	virtual HRESULT STDMETHODCALLTYPE get_ParentTable(IRichTable** ppTable);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether the cells in the table row have right-to-left precedence. If set to
	/// \c VARIANT_TRUE, the row's cells are prepared for right-to-left text, otherwise for left-to-right
	/// text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa put_RightToLeft, get_TextRange, TextParagraph::get_RightToLeft
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Sets whether the cells in the table row have right-to-left precedence. If set to
	/// \c VARIANT_TRUE, the row's cells are prepared for right-to-left text, otherwise for left-to-right
	/// text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa get_RightToLeft, get_TextRange, TextParagraph::put_RightToLeft
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c TextRange property</em>
	///
	/// Retrieves a \c TextRange object for the table row.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TextRange
	virtual HRESULT STDMETHODCALLTYPE get_TextRange(IRichTextRange** ppTextRange);

	/// \brief <em>Retrieves whether two \c IRichTableRow instances have the same properties</em>
	///
	/// \param[in] pCompareAgainst The \c IRichTableRow instance to compare with.
	/// \param[out] pValue \c VARIANT_TRUE, if the objects have the same properties; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	virtual HRESULT STDMETHODCALLTYPE Equals(IRichTableRow* pCompareAgainst, VARIANT_BOOL* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given text range</em>
	///
	/// Attaches this object to a given text range that spans the whole table row, so that the table row's
	/// properties can be retrieved and set using this object's methods.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table to attach to.
	/// \param[in] pRowTextRange The \c ITextRange object of the table row to attach to.
	///
	/// \sa Detach,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	void Attach(ITextRange* pTableTextRange, ITextRange* pRowTextRange);
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
		/// \brief <em>The \c RichTextBox object that owns this table row</em>
		///
		/// \sa SetOwner
		RichTextBox* pOwnerRTB;
		/// \brief <em>The \c ITextRange object corresponding to the table wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pTableTextRange;
		/// \brief <em>The \c ITextRange object corresponding to the table row wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pRowTextRange;

		Properties()
		{
			pOwnerRTB = NULL;
			pTableTextRange = NULL;
			pRowTextRange = NULL;
		}

		~Properties();

		/// \brief <em>Retrieves the owning rich text box' window handle</em>
		///
		/// \return The window handle of the rich text box that contains this table row.
		///
		/// \sa pOwnerRTB
		HWND GetRTBHWnd(void);
	} properties;
};     // TableRow

OBJECT_ENTRY_AUTO(__uuidof(TableRow), TableRow)