//////////////////////////////////////////////////////////////////////
/// \class TextSubRanges
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c TextRange objects</em>
///
/// This class provides easy access to collections of \c TextRange objects that are sub-ranges of another
/// \c TextRange object.
/// \c TextSubRanges objects are used to group text sub-ranges that have certain properties in common.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTextSubRanges, RichTextBox, TextRange
/// \else
///   \sa RichTextBoxCtlLibA::IRichTextSubRanges, RichTextBox, TextRange
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITextSubRangesEvents_CP.h"
#include "helpers.h"
#include "TextRange.h"
#include "RichTextBox.h"


class ATL_NO_VTABLE TextSubRanges : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TextSubRanges, &CLSID_TextSubRanges>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TextSubRanges>,
    public Proxy_IRichTextSubRangesEvents<TextSubRanges>, 
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IRichTextSubRanges, &IID_IRichTextSubRanges, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTextSubRanges, &IID_IRichTextSubRanges, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class RichTextBox;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TEXTSUBRANGES)

		BEGIN_COM_MAP(TextSubRanges)
			COM_INTERFACE_ENTRY(IRichTextSubRanges)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange2), 0, QueryITextRangeInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TextSubRanges)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTextSubRangesEvents))
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the sub-ranges</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c TextRange objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x sub-ranges</em>
	///
	/// Retrieves the next \c numberOfMaxRanges sub-ranges from the iterator.
	///
	/// \param[in] numberOfMaxRanges The maximum number of sub-ranges the array identified by \c pTextRanges
	///            can contain.
	/// \param[in,out] pTextRanges An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a sub-range's \c IRichTextRange implementation.
	/// \param[out] pNumberOfRangesReturned The number of sub-ranges that actually were copied to the array
	///             identified by \c pTextRanges.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, TextRange,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxRanges, VARIANT* pTextRanges, ULONG* pNumberOfRangesReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// sub-range in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x sub-ranges</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfRangesToSkip items.
	///
	/// \param[in] numberOfRangesToSkip The number of sub-ranges to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfRangesToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTextSubRanges
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c ActiveSubRange property</em>
	///
	/// Retrieves the sub-range that is affected by operations such as [SHIFT]+Arrow keys if the parent
	/// text range is the selection.
	///
	/// \param[out] ppActiveSubRange Receives the active sub-range's \c IRichTextRange implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_ActiveSubRange, RichTextBox::get_MultiSelect
	// \sa RichTextBox::get_MultiSelect
	virtual HRESULT STDMETHODCALLTYPE get_ActiveSubRange(IRichTextRange** ppActiveSubRange);
	// \brief <em>Sets the \c ActiveSubRange property</em>
	//
	// Sets the sub-range that is affected by operations such as [SHIFT]+Arrow keys if the parent
	// text range is the selection.
	//
	// \param[in] pNewActiveSubRange The new active sub-range's \c IRichTextRange implementation.
	//
	// \return An \c HRESULT error code.
	//
	// \sa get_ActiveSubRange, RichTextBox::put_MultiSelect
	//virtual HRESULT STDMETHODCALLTYPE putref_ActiveSubRange(IRichTextRange* pNewActiveSubRange);
	/// \brief <em>Retrieves a \c TextRange object from the collection</em>
	///
	/// Retrieves a \c TextRange object from the collection that wraps the sub-range identified by
	/// \c subRangeIdentifier.
	///
	/// \param[in] subRangeIdentifier A value that identifies the sub-range to be retrieved.
	/// \param[in] subRangeIdentifierType A value specifying the meaning of \c subRangeIdentifier. Any of
	///            the values defined by the \c SubRangeIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppSubRange Receives the sub-range's \c IRichTextRange implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa TextRange, Add, Remove, Contains, RichTextBoxCtlLibU::SubRangeIdentifierTypeConstants
	/// \else
	///   \sa TextRange, Add, Remove, Contains, RichTextBoxCtlLibA::SubRangeIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(VARIANT subRangeIdentifier, SubRangeIdentifierTypeConstants subRangeIdentifierType = sritIndex, IRichTextRange** ppSubRange = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c TextRange objects
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

	/// \brief <em>Adds a sub-range to the text range</em>
	///
	/// Adds a sub-range with the specified properties to the text range and returns a \c TextRange
	/// object wrapping the inserted sub-range.
	///
	/// \param[in] activeSubRangeEnd The character position of the active end of the sub-range. The active
	///            end is the end with the cursor.
	/// \param[in] anchorSubRangeEnd The character position of the anchor end of the sub-range. The anchor
	///            end is the end at which selecting the text range starts.
	/// \param[in] activateSubRange Specifies whether the new sub-range will become the active sub-range
	///            within its parent text range. If set to \c VARIANT_TRUE, it will become the active
	///            sub-range; otherwise not.
	/// \param[out] ppAddedSubRange Receives the added sub-range's \c IRichTextRange implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ActiveSubRange, Count, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Add(LONG activeSubRangeEnd, LONG anchorSubRangeEnd, VARIANT_BOOL activateSubRange = VARIANT_TRUE, IRichTextRange** ppAddedSubRange = NULL);
	/// \brief <em>Retrieves whether the specified sub-range is part of the sub-range collection</em>
	///
	/// \param[in] subRangeIdentifier A value that identifies the sub-range to be checked.
	/// \param[in] subRangeIdentifierType A value specifying the meaning of \c subRangeIdentifier. Any of
	///            the values defined by the \c SubRangeIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the sub-range is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Remove, RichTextBoxCtlLibU::SubRangeIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, RichTextBoxCtlLibA::SubRangeIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(VARIANT subRangeIdentifier, SubRangeIdentifierTypeConstants subRangeIdentifierType = sritIndex, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the sub-ranges in the collection</em>
	///
	/// Retrieves the number of \c TextRange objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified sub-range in the collection from the text range</em>
	///
	/// \param[in] subRangeIdentifier A value that identifies the sub-range to be removed.
	/// \param[in] subRangeIdentifierType A value specifying the meaning of \c subRangeIdentifier. Any of
	///            the values defined by the \c SubRangeIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibU::SubRangeIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibA::SubRangeIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(VARIANT subRangeIdentifier, SubRangeIdentifierTypeConstants subRangeIdentifierType = sritIndex);
	/// \brief <em>Removes all sub-ranges in the collection from the text range</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
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
		/// \brief <em>Holds the last enumerated sub-range</em>
		int lastEnumeratedSubRange;

		Properties()
		{
			pOwnerRTB = NULL;
			pTextRange = NULL;
			lastEnumeratedSubRange = 1;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
	} properties;
};     // TextSubRanges

OBJECT_ENTRY_AUTO(__uuidof(TextSubRanges), TextSubRanges)