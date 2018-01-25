//////////////////////////////////////////////////////////////////////
/// \class Table
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a table's styling, e.g. alignment</em>
///
/// This class is a wrapper around the styling of a table.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTable, RichTextBox, TextRange
/// \else
///   \sa RichTextBoxCtlLibA::IRichTable, RichTextBox, TextRange
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITableEvents_CP.h"
#include "helpers.h"
#include "TableRows.h"


class ATL_NO_VTABLE Table : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<Table, &CLSID_Table>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<Table>,
    public Proxy_IRichTableEvents<Table>, 
    #ifdef UNICODE
    	public IDispatchImpl<IRichTable, &IID_IRichTable, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTable, &IID_IRichTable, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TABLE)

		BEGIN_COM_MAP(Table)
			COM_INTERFACE_ENTRY(IRichTable)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange2), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRow), 0, QueryITextRowInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(Table)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTableEvents))
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
	/// \name Implementation of IRichTable
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c NestingLevel property</em>
	///
	/// Retrieves an integer value describing the nesting of tables. If the table is the outermost table,
	/// its nesting level is 1. If it is embedded in another table, that is the outermost table, its
	/// nesting level is 2 and so on.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TextRange::ReplaceWithTable
	virtual HRESULT STDMETHODCALLTYPE get_NestingLevel(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Rows property</em>
	///
	/// Retrieves a collection object wrapping the table's rows.
	///
	/// \param[out] ppRows Receives the collection object's \c IRichTableRows implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TableRows
	virtual HRESULT STDMETHODCALLTYPE get_Rows(IRichTableRows** ppRows);
	/// \brief <em>Retrieves the current setting of the \c TextRange property</em>
	///
	/// Retrieves a \c TextRange object for the table.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TextRange
	virtual HRESULT STDMETHODCALLTYPE get_TextRange(IRichTextRange** ppTextRange);
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
	/// \brief <em>Sets the owner of this table</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerRTB
	void SetOwner(__in_opt RichTextBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c RichTextBox object that owns this table</em>
		///
		/// \sa SetOwner
		RichTextBox* pOwnerRTB;
		/// \brief <em>The \c ITextRange object corresponding to the table wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pTableTextRange;

		Properties()
		{
			pOwnerRTB = NULL;
			pTableTextRange = NULL;
		}

		~Properties();

		/// \brief <em>Retrieves the owning rich text box' window handle</em>
		///
		/// \return The window handle of the rich text box that contains this table.
		///
		/// \sa pOwnerRTB
		HWND GetRTBHWnd(void);
	} properties;
};     // Table

OBJECT_ENTRY_AUTO(__uuidof(Table), Table)