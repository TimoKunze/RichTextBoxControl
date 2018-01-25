//////////////////////////////////////////////////////////////////////
/// \class TableRows
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c TableRowsRow objects</em>
///
/// This class provides easy access to collections of \c TableRowsRow objects.
/// \c TableRowsRows objects are used to group table rows that have certain properties in common.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTableRows, RichTextBox, TextRange
/// \else
///   \sa RichTextBoxCtlLibA::IRichTableRows, RichTextBox, TextRange
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITableRowsEvents_CP.h"
#include "helpers.h"
#include "TableRow.h"


class ATL_NO_VTABLE TableRows : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TableRows, &CLSID_TableRows>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TableRows>,
    public Proxy_IRichTableRowsEvents<TableRows>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IRichTableRows, &IID_IRichTableRows, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTableRows, &IID_IRichTableRows, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TABLEROWS)

		BEGIN_COM_MAP(TableRows)
			COM_INTERFACE_ENTRY(IRichTableRows)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange2), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRow), 0, QueryITextRowInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TableRows)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTableRowsEvents))
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the table rows</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c TableRow objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x table rows</em>
	///
	/// Retrieves the next \c numberOfMaxRows table rows from the iterator.
	///
	/// \param[in] numberOfMaxRows The maximum number of table rows the array identified by \c pRows can
	///            contain.
	/// \param[in,out] pRows An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a table row's \c IRichTableRow implementation.
	/// \param[out] pNumberOfRowsReturned The number of table rows that actually were copied to the array
	///             identified by \c pRows.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, TableRow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxRows, VARIANT* pRows, ULONG* pNumberOfRowsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// table row in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x table rows</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfRowsToSkip items.
	///
	/// \param[in] numberOfRowsToSkip The number of table rows to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfRowsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTableRows
	///
	//@{
	/// \brief <em>Retrieves a \c TableRow object from the collection</em>
	///
	/// Retrieves a \c TableRow object from the collection that wraps the table row identified
	/// by \c rowIdentifier.
	///
	/// \param[in] rowIdentifier A value that identifies the table row to be retrieved.
	/// \param[in] rowIdentifierType A value specifying the meaning of \c rowIdentifier. Any of the
	///            values defined by the \c RowIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppRow Receives the row's \c IRichTableRow implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa TableRow, Add, Remove, Contains, RichTextBoxCtlLibU::RowIdentifierTypeConstants
	/// \else
	///   \sa TableRow, Add, Remove, Contains, RichTextBoxCtlLibA::RowIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG rowIdentifier, RowIdentifierTypeConstants rowIdentifierType = ritIndex, IRichTableRow** ppRow = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c TableRow objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator A pointer to the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);
	/// \brief <em>Retrieves the current setting of the \c ParentTable property</em>
	///
	/// Retrieves a \c Table object that wraps the table that the collection refers to.
	///
	/// \param[out] ppTable The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Table
	virtual HRESULT STDMETHODCALLTYPE get_ParentTable(IRichTable** ppTable);

	/// \brief <em>Adds a row to the table</em>
	///
	/// Adds a row with the specified properties at the specified position in the table and returns
	/// a \c TableRow object wrapping the inserted row.
	///
	/// \param[in] insertAt The zero-based index at which to insert the new row. If set to -1, the row is
	///            inserted after the last row in the table.
	/// \param[in] horizontalAlignment The horizontal alignment of the table row. The following of the
	///            values defined by the \c HAlignmentConstants enumeration are valid: \c halLeft,
	///            \c halCenter, \c halRight.
	/// \param[in] indent The size (in twips) by which the table row is indented horizontally. The indent
	///            is measured from the control's left border.
	/// \param[in] horizontalCellMargin The size (in twips) of the left and right margin in each cell. If
	///            set to -1, the default margin is used.
	/// \param[in] height The row's height (in twips). If set to -1, the default height is used.
	/// \param[out] ppAddedRow Receives the added row's \c IRichTableRow implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, TableRow::get_Height, TableRow::get_Index, TableRow::get_HAlignment,
	///       TableRow::get_Indent, TableRow::get_HorizontalCellMargin,
	///       RichTextBoxCtlLibU::HAlignmentConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, TableRow::get_Height, TableRow::get_Index, TableRow::get_HAlignment,
	///       TableRow::get_Indent, TableRow::get_HorizontalCellMargin,
	///       RichTextBoxCtlLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(LONG insertAt = -1, HAlignmentConstants horizontalAlignment = halLeft, LONG indent = -1, LONG horizontalCellMargin = -1, LONG height = -1, IRichTableRow** ppAddedRow = NULL);
	/// \brief <em>Retrieves whether the specified table row is part of the table row collection</em>
	///
	/// \param[in] rowIdentifier A value that identifies the table row to be checked.
	/// \param[in] rowIdentifierType A value specifying the meaning of \c rowIdentifier. Any of the
	///            values defined by the \c RowIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the table row is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Remove, RichTextBoxCtlLibU::RowIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, RichTextBoxCtlLibA::RowIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(LONG rowIdentifier, RowIdentifierTypeConstants rowIdentifierType = ritIndex, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the table rows in the collection</em>
	///
	/// Retrieves the number of \c TableRow objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified row in the collection from the table</em>
	///
	/// \param[in] rowIdentifier A value that identifies the row to be removed.
	/// \param[in] rowIdentifierType A value specifying the meaning of \c rowIdentifier. Any of the
	///            values defined by the \c RowIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibU::RowIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibA::RowIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG rowIdentifier, RowIdentifierTypeConstants rowIdentifierType = ritIndex);
	/// \brief <em>Removes all rows in the collection from the table</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given text range</em>
	///
	/// Attaches this object to a given text range that spans the whole table, so that the table's properties
	/// can be retrieved and set using this object's methods.
	///
	/// \param[in] pTableTextRange The \c ITextRange object to attach to.
	///
	/// \sa Detach,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
	void Attach(ITextRange* pTableTextRange);
	/// \brief <em>Detaches this object from a text range</em>
	///
	/// Detaches this object from the text range it currently wraps, so that it doesn't wrap any text range
	/// anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerRTB
	void SetOwner(__in_opt RichTextBox* pOwner);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Filter appliance
	///
	//@{
	/// \brief <em>Retrieves the table's first row that might be in the collection</em>
	///
	/// \return The table row being the collection's first candidate. \c NULL if no row was found.
	///
	/// \sa GetNextRowToProcess, Count, RemoveAll, Next
	ITextRange* GetFirstRowToProcess(void);
	/// \brief <em>Retrieves the table's next row that might be in the collection</em>
	///
	/// \param[in] pPreviousRow The row at which to start the search for a new collection candidate.
	///
	/// \return The next table row being a candidate for the collection. \c NULL if no row was found.
	///
	/// \sa GetFirstRowToProcess, Count, RemoveAll, Next
	ITextRange* GetNextRowToProcess(ITextRange* pPreviousRow);
	/// \brief <em>Retrieves whether a table row is part of the collection (applying the filters)</em>
	///
	/// \param[in] pRowTextRange The text range defining the table row to check.
	///
	/// \return \c TRUE, if the table row is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(ITextRange* pRowTextRange);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c RichTextBox object that owns these table rows</em>
		///
		/// \sa SetOwner
		RichTextBox* pOwnerRTB;
		/// \brief <em>The \c ITextRange object corresponding to the table wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pTableTextRange;
		/// \brief <em>The \c ITextRange object corresponding to the last enumerated table row</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pLastEnumeratedRowTextRange;

		Properties()
		{
			pOwnerRTB = NULL;
			pTableTextRange = NULL;
			pLastEnumeratedRowTextRange = NULL;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
		/// \brief <em>Retrieves the owning rich text box' window handle</em>
		///
		/// \return The window handle of the rich text box that contains these table rows.
		///
		/// \sa pOwnerRTB
		HWND GetRTBHWnd(void);
	} properties;
};     // TableRows

OBJECT_ENTRY_AUTO(__uuidof(TableRows), TableRows)