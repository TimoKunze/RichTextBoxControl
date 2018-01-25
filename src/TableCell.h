//////////////////////////////////////////////////////////////////////
/// \class TableCell
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a table cell's styling, e.g. margins</em>
///
/// This class is a wrapper around the styling of a table cell.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTableCell, RichTextBox, TextRange
/// \else
///   \sa RichTextBoxCtlLibA::IRichTableCell, RichTextBox, TextRange
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITableCellEvents_CP.h"
#include "helpers.h"


class ATL_NO_VTABLE TableCell : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TableCell, &CLSID_TableCell>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TableCell>,
    public Proxy_IRichTableCellEvents<TableCell>, 
    #ifdef UNICODE
    	public IDispatchImpl<IRichTableCell, &IID_IRichTableCell, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTableCell, &IID_IRichTableCell, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TABLECELL)

		BEGIN_COM_MAP(TableCell)
			COM_INTERFACE_ENTRY(IRichTableCell)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange2), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRow), 0, QueryITextRowInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TableCell)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTableCellEvents))
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
	/// \name Implementation of IRichTableCell
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c BackColor1 property</em>
	///
	/// Retrieves the table cell's first background color. The actual background color is calculated
	/// from \c BackColor1 and \c BackColor2, with \c BackColorShading controlling the mixture.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.\n
	///          If this property is set for a cell, it also applies to all cells in the same row that
	///          follow this cell.
	///
	/// \sa put_BackColor1, get_BackColor2, get_BackColorShading, get_BorderColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor1(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor1 property</em>
	///
	/// Sets the table cell's first background color. The actual background color is calculated
	/// from \c BackColor1 and \c BackColor2, with \c BackColorShading controlling the mixture.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.\n
	///          If this property is set for a cell, it also applies to all cells in the same row that
	///          follow this cell.
	///
	/// \sa get_BackColor1, put_BackColor2, put_BackColorShading, put_BorderColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor1(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor2 property</em>
	///
	/// Retrieves the table cell's second background color. The actual background color is calculated
	/// from \c BackColor1 and \c BackColor2, with \c BackColorShading controlling the mixture.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.\n
	///          If this property is set for a cell, it also applies to all cells in the same row that
	///          follow this cell.
	///
	/// \sa put_BackColor2, get_BackColor1, get_BackColorShading, get_BorderColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor2(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor2 property</em>
	///
	/// Sets the table cell's second background color. The actual background color is calculated
	/// from \c BackColor1 and \c BackColor2, with \c BackColorShading controlling the mixture.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.\n
	///          If this property is set for a cell, it also applies to all cells in the same row that
	///          follow this cell.
	///
	/// \sa get_BackColor2, put_BackColor1, put_BackColorShading, put_BorderColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor2(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColorShading property</em>
	///
	/// Retrieves the proportion by which both background colors are mixed to create the table cell's
	/// actual background color. \c BackColorShading can be a value between 0 and 10000. If set to 0, the
	/// actual background color will be the same as \c BackColor1. If set to 10000, the actual background
	/// color will be the same as \c BackColor2. Any value inbetween leads to a mixture of both colors.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.\n
	///          If this property is set for a cell, it also applies to all cells in the same row that
	///          follow this cell.
	///
	/// \sa put_BackColorShading, get_BackColor1, get_BackColor2, get_BorderColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColorShading(LONG* pValue);
	/// \brief <em>Sets the \c BackColorShading property</em>
	///
	/// Sets the proportion by which both background colors are mixed to create the table cell's
	/// actual background color. \c BackColorShading can be a value between 0 and 10000. If set to 0, the
	/// actual background color will be the same as \c BackColor1. If set to 10000, the actual background
	/// color will be the same as \c BackColor2. Any value inbetween leads to a mixture of both colors.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.\n
	///          If this property is set for a cell, it also applies to all cells in the same row that
	///          follow this cell.
	///
	/// \sa get_BackColorShading, put_BackColor1, put_BackColor2, put_BorderColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColorShading(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderColor property</em>
	///
	/// Retrieves the color of a cell border.
	///
	/// \param[in] border The border for which to retrieve or set the border color. Any of the values
	///            defined by the \c CellBorderConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa put_BorderColor, get_BorderSize, get_BackColor1, get_BackColor2, get_BackColorShading,
	///       RichTextBoxCtlLibU::CellBorderConstants
	/// \else
	///   \sa put_BorderColor, get_BorderSize, get_BackColor1, get_BackColor2, get_BackColorShading,
	///       RichTextBoxCtlLibA::CellBorderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderColor(CellBorderConstants border, OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BorderColor property</em>
	///
	/// Sets the color of a cell border.
	///
	/// \param[in] border The border for which to retrieve or set the border color. Any of the values
	///            defined by the \c CellBorderConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa get_BorderColor, put_BorderSize, put_BackColor1, put_BackColor2, put_BackColorShading,
	///       RichTextBoxCtlLibU::CellBorderConstants
	/// \else
	///   \sa get_BorderColor, put_BorderSize, put_BackColor1, put_BackColor2, put_BackColorShading,
	///       RichTextBoxCtlLibA::CellBorderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderColor(CellBorderConstants border, OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderSize property</em>
	///
	/// Retrieves the width (in twips) of a cell border. If set to -1, the width of 1 pixel will be used.
	///
	/// \param[in] border The border for which to retrieve or set the border size. Any of the values
	///            defined by the \c CellBorderConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa put_BorderSize, get_BorderColor, get_CellMergeFlags, RichTextBoxCtlLibU::CellBorderConstants
	/// \else
	///   \sa put_BorderSize, get_BorderColor, get_CellMergeFlags, RichTextBoxCtlLibA::CellBorderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderSize(CellBorderConstants border, SHORT* pValue);
	/// \brief <em>Sets the \c BorderSize property</em>
	///
	/// Sets the width (in twips) of a cell border. If set to -1, the width of 1 pixel will be used.
	///
	/// \param[in] border The border for which to retrieve or set the border size. Any of the values
	///            defined by the \c CellBorderConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa get_BorderSize, put_BorderColor, put_CellMergeFlags, RichTextBoxCtlLibU::CellBorderConstants
	/// \else
	///   \sa get_BorderSize, put_BorderColor, put_CellMergeFlags, RichTextBoxCtlLibA::CellBorderConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderSize(CellBorderConstants border, SHORT newValue);
	/// \brief <em>Retrieves the current setting of the \c CellMergeFlags property</em>
	///
	/// Retrieves whether the table cell is merged with one or more surrounding cell. Any
	/// combination of the values defined by the \c CellMergeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa put_CellMergeFlags, get_BorderSize, RichTextBoxCtlLibU::CellMergeConstants
	/// \else
	///   \sa put_CellMergeFlags, get_BorderSize, RichTextBoxCtlLibA::CellMergeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CellMergeFlags(CellMergeConstants* pValue);
	/// \brief <em>Sets the \c CellMergeFlags property</em>
	///
	/// Sets whether the table cell is merged with one or more surrounding cell. Any
	/// combination of the values defined by the \c CellMergeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa get_CellMergeFlags, put_BorderSize, RichTextBoxCtlLibU::CellMergeConstants
	/// \else
	///   \sa get_CellMergeFlags, put_BorderSize, RichTextBoxCtlLibA::CellMergeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_CellMergeFlags(CellMergeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DefinesStyleForSubsequentCells property</em>
	///
	/// Retrieves whether subsequent table cells inherit their style from this cell. If set to
	/// \c VARIANT_TRUE, this cell inherits its style to subsequent cells; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.\n
	///          This property is read-only.
	///
	/// \sa TableRow::get_MasterCell
	virtual HRESULT STDMETHODCALLTYPE get_DefinesStyleForSubsequentCells(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index identifying this table cell.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa RichTextBoxCtlLibU::CellIdentifierTypeConstants
	/// \else
	///   \sa RichTextBoxCtlLibA::CellIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ParentRow property</em>
	///
	/// Retrieves a \c TableRow object that wraps the table row that the table cell refers to.
	///
	/// \param[out] ppRow The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TableRow
	virtual HRESULT STDMETHODCALLTYPE get_ParentRow(IRichTableRow** ppRow);
	/// \brief <em>Retrieves the current setting of the \c ParentTable property</em>
	///
	/// Retrieves a \c Table object that wraps the table that the table cell refers to.
	///
	/// \param[out] ppTable The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Table
	virtual HRESULT STDMETHODCALLTYPE get_ParentTable(IRichTable** ppTable);
	/// \brief <em>Retrieves the current setting of the \c TextFlow property</em>
	///
	/// Retrieves how the text flows within the table cell. Some of the values defined by the
	/// \c TextFlowConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_TextFlow, RichTextBox::get_TextFlow, RichTextBoxCtlLibU::TextFlowConstants
	/// \else
	///   \sa put_TextFlow, RichTextBox::get_TextFlow, RichTextBoxCtlLibA::TextFlowConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TextFlow(TextFlowConstants* pValue);
	/// \brief <em>Sets the \c TextFlow property</em>
	///
	/// Sets how the text flows within the table cell. Some of the values defined by the
	/// \c TextFlowConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_TextFlow, RichTextBox::put_TextFlow, RichTextBoxCtlLibU::TextFlowConstants
	/// \else
	///   \sa get_TextFlow, RichTextBox::put_TextFlow, RichTextBoxCtlLibA::TextFlowConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TextFlow(TextFlowConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c TextRange property</em>
	///
	/// Retrieves a \c TextRange object for the table cell.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TextRange
	virtual HRESULT STDMETHODCALLTYPE get_TextRange(IRichTextRange** ppTextRange);
	/// \brief <em>Retrieves the current setting of the \c VAlignment property</em>
	///
	/// Retrieves the vertical alignment of the table cell's content within the cell.
	/// Any of the values defined by the \c VAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa put_VAlignment, TableRow::get_Height, RichTextBoxCtlLibU::VAlignmentConstants
	/// \else
	///   \sa put_VAlignment, TableRow::get_Height, RichTextBoxCtlLibA::VAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_VAlignment(VAlignmentConstants* pValue);
	/// \brief <em>Sets the \c VAlignment property</em>
	///
	/// Sets the vertical alignment of the table cell's content within the cell.
	/// Any of the values defined by the \c VAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa get_VAlignment, TableRow::put_Height, RichTextBoxCtlLibU::VAlignmentConstants
	/// \else
	///   \sa get_VAlignment, TableRow::put_Height, RichTextBoxCtlLibA::VAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_VAlignment(VAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Width property</em>
	///
	/// Retrieves the table cell's width (in twips).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa put_Width, TableRow::get_Height
	virtual HRESULT STDMETHODCALLTYPE get_Width(LONG* pValue);
	/// \brief <em>Sets the \c Width property</em>
	///
	/// Sets the table cell's width (in twips).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing this property requires Rich Edit 7.5 or newer.
	///
	/// \sa get_Width, TableRow::put_Height
	virtual HRESULT STDMETHODCALLTYPE put_Width(LONG newValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given text range</em>
	///
	/// Attaches this object to a given text range that spans the whole table cell, so that the table cell's
	/// properties can be retrieved and set using this object's methods.
	///
	/// \param[in] pTableTextRange The \c ITextRange object of the table to attach to.
	/// \param[in] pRowTextRange The \c ITextRange object of the table row to attach to.
	/// \param[in] pCellTextRange The \c ITextRange object of the table cell to attach to.
	///
	/// \sa Detach,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	void Attach(ITextRange* pTableTextRange, ITextRange* pRowTextRange, ITextRange* pCellTextRange);
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
		/// \brief <em>The \c RichTextBox object that owns this table cell</em>
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
		/// \brief <em>The \c ITextRange object corresponding to the table cell wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pCellTextRange;

		Properties()
		{
			pOwnerRTB = NULL;
			pTableTextRange = NULL;
			pRowTextRange = NULL;
			pCellTextRange = NULL;
		}

		~Properties();

		/// \brief <em>Retrieves the owning rich text box' window handle</em>
		///
		/// \return The window handle of the rich text box that contains this table cell.
		///
		/// \sa pOwnerRTB
		HWND GetRTBHWnd(void);
	} properties;
};     // TableCell

OBJECT_ENTRY_AUTO(__uuidof(TableCell), TableCell)