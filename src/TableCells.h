//////////////////////////////////////////////////////////////////////
/// \class TableCells
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c TableCellsCell objects</em>
///
/// This class provides easy access to collections of \c TableCell objects.
/// \c TableCell objects are used to group table cells that have certain properties in common.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTableCells, RichTextBox, TextRange
/// \else
///   \sa RichTextBoxCtlLibA::IRichTableCells, RichTextBox, TextRange
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITableCellsEvents_CP.h"
#include "helpers.h"
#include "TableCell.h"


class ATL_NO_VTABLE TableCells : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TableCells, &CLSID_TableCells>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TableCells>,
    public Proxy_IRichTableCellsEvents<TableCells>, 
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IRichTableCells, &IID_IRichTableCells, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTableCells, &IID_IRichTableCells, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TABLEROWS)

		BEGIN_COM_MAP(TableCells)
			COM_INTERFACE_ENTRY(IRichTableCells)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange2), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRow), 0, QueryITextRowInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TableCells)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTableCellsEvents))
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
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the table cells</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c TableCell objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x table cells</em>
	///
	/// Retrieves the next \c numberOfMaxCells table cells from the iterator.
	///
	/// \param[in] numberOfMaxCells The maximum number of table cells the array identified by \c pCells can
	///            contain.
	/// \param[in,out] pCells An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a table cell's \c IRichTableCell implementation.
	/// \param[out] pNumberOfCellsReturned The number of table cells that actually were copied to the array
	///             identified by \c pCells.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, TableCell,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxCells, VARIANT* pCells, ULONG* pNumberOfCellsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// table cell in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x table cells</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfCellsToSkip items.
	///
	/// \param[in] numberOfCellsToSkip The number of table cells to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfCellsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTableCells
	///
	//@{
	/// \brief <em>Retrieves a \c TableCell object from the collection</em>
	///
	/// Retrieves a \c TableCell object from the collection that wraps the table cell identified
	/// by \c cellIdentifier.
	///
	/// \param[in] cellIdentifier A value that identifies the table cell to be retrieved.
	/// \param[in] cellIdentifierType A value specifying the meaning of \c cellIdentifier. Any of the
	///            values defined by the \c CellIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppCell Receives the cell's \c IRichTableCell implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa TableCell, Add, Remove, Contains, RichTextBoxCtlLibU::CellIdentifierTypeConstants
	/// \else
	///   \sa TableCell, Add, Remove, Contains, RichTextBoxCtlLibA::CellIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG cellIdentifier, CellIdentifierTypeConstants cellIdentifierType = citIndex, IRichTableCell** ppCell = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c TableCell objects
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
	/// \brief <em>Retrieves the current setting of the \c ParentRow property</em>
	///
	/// Retrieves a \c TableRow object that wraps the table row that the collection refers to.
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

	/// \brief <em>Adds a cell to the table row</em>
	///
	/// Adds a cell with the specified properties at the specified position in the table row and returns
	/// a \c TableCell object wrapping the inserted cell.
	///
	/// \param[in] width The width (in twips) of the new cell.
	/// \param[in] insertAt The zero-based index at which to insert the new cell. If set to -1, the cell is
	///            inserted after the last cell in the row.
	/// \param[in] borderSize The width (in twips) of the cell borders. If set to -1, the width of 1 pixel
	///            will be used.
	/// \param[in] verticalAlignment The vertical alignment of the cell's content within the table
	///            cell. Any of the values defined by the \c VAlignmentConstants enumeration is valid.
	/// \param[out] ppAddedCell Receives the added cell's \c IRichTableCell implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, TableCell::get_Width, TableCell::get_Index, TableCell::get_BorderSize,
	///       TableCell::get_BorderSize, TableCell::get_VAlignment, RichTextBoxCtlLibU::VAlignmentConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, TableCell::get_Width, TableCell::get_Index, TableCell::get_BorderSize,
	///       TableCell::get_BorderSize, TableCell::get_VAlignment, RichTextBoxCtlLibA::VAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(LONG width, LONG insertAt = -1, SHORT borderSize = -1, VAlignmentConstants verticalAlignment = valCenter, IRichTableCell** ppAddedCell = NULL);
	/// \brief <em>Retrieves whether the specified table cell is part of the table cell collection</em>
	///
	/// \param[in] cellIdentifier A value that identifies the table cell to be checked.
	/// \param[in] cellIdentifierType A value specifying the meaning of \c cellIdentifier. Any of the
	///            values defined by the \c CellIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the table cell is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Remove, RichTextBoxCtlLibU::CellIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, RichTextBoxCtlLibA::CellIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(LONG cellIdentifier, CellIdentifierTypeConstants cellIdentifierType = citIndex, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the table cells in the collection</em>
	///
	/// Retrieves the number of \c TableCell objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified cell in the collection from the table row</em>
	///
	/// \param[in] cellIdentifier A value that identifies the cell to be removed.
	/// \param[in] cellIdentifierType A value specifying the meaning of \c cellIdentifier. Any of the
	///            values defined by the \c CellIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires Rich Edit 7.5 or newer.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibU::CellIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibA::CellIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG cellIdentifier, CellIdentifierTypeConstants cellIdentifierType = citIndex);
	/// \brief <em>Removes all cells in the collection from the table row</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given text range</em>
	///
	/// Attaches this object to a given text range that spans the whole table row, so that the table row's
	/// cells can be retrieved and modified using this object's methods.
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
	//////////////////////////////////////////////////////////////////////
	/// \name Filter appliance
	///
	//@{
	/// \brief <em>Retrieves the table's first cell that might be in the collection</em>
	///
	/// \return The table cell being the collection's first candidate. -1 if no cell was found.
	///
	/// \sa GetNextCellToProcess, Count, RemoveAll, Next
	ITextRange* GetFirstCellToProcess(void);
	/// \brief <em>Retrieves the table's next cell that might be in the collection</em>
	///
	/// \param[in] pPreviousCell The cell at which to start the search for a new collection candidate.
	///
	/// \return The next table cell being a candidate for the collection. -1 if no cell was found.
	///
	/// \sa GetFirstCellToProcess, Count, RemoveAll, Next
	ITextRange* GetNextCellToProcess(ITextRange* pPreviousCell);
	/// \brief <em>Retrieves whether a table cell is part of the collection (applying the filters)</em>
	///
	/// \param[in] pCellTextRange The text range defining the table cell to check.
	///
	/// \return \c TRUE, if the table cell is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(ITextRange* pCellTextRange);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c RichTextBox object that owns these table cells</em>
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
		/// \brief <em>The \c ITextRange object corresponding to the last enumerated table cell</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pLastEnumeratedCellTextRange;

		Properties()
		{
			pOwnerRTB = NULL;
			pTableTextRange = NULL;
			pRowTextRange = NULL;
			pLastEnumeratedCellTextRange = NULL;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
		/// \brief <em>Retrieves the owning rich text box' window handle</em>
		///
		/// \return The window handle of the rich text box that contains these table cells.
		///
		/// \sa pOwnerRTB
		HWND GetRTBHWnd(void);
	} properties;
};     // TableCells

OBJECT_ENTRY_AUTO(__uuidof(TableCells), TableCells)