//////////////////////////////////////////////////////////////////////
/// \class TextParagraphTabs
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c TextParagraphTab objects</em>
///
/// This class provides easy access to collections of \c TextParagraphTab objects.
/// \c TextParagraphTabs objects are used to group tabs that have certain properties in common.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTextParagraphTab, TextParagraphTab, RichTextBox
/// \else
///   \sa RichTextBoxCtlLibA::IRichTextParagraphTab, TextParagraphTab, RichTextBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITextParagraphTabsEvents_CP.h"
#include "helpers.h"
#include "TextParagraph.h"
#include "TextParagraphTab.h"


class ATL_NO_VTABLE TextParagraphTabs : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TextParagraphTabs, &CLSID_TextParagraphTabs>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TextParagraphTabs>,
    public Proxy_IRichTextParagraphTabsEvents<TextParagraphTabs>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IRichTextParagraphTabs, &IID_IRichTextParagraphTabs, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTextParagraphTabs, &IID_IRichTextParagraphTabs, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class TextParagraphTab;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TEXTPARAGRAPHTABS)

		BEGIN_COM_MAP(TextParagraphTabs)
			COM_INTERFACE_ENTRY(IRichTextParagraphTabs)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TextParagraphTabs)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTextParagraphTabsEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportsErrorInfo
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

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the tabs</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c TextParagraphTab objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x tabs</em>
	///
	/// Retrieves the next \c numberOfMaxTabs tabs from the iterator.
	///
	/// \param[in] numberOfMaxTabs The maximum number of tabs the array identified by \c pTabs can
	///            contain.
	/// \param[in,out] pTabs An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a tab's \c IRichTextParagraphTab implementation.
	/// \param[out] pNumberOfTabsReturned The number of tabs that actually were copied to the array
	///             identified by \c pTabs.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, TextParagraphTab,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxTabs, VARIANT* pTabs, ULONG* pNumberOfTabsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// tab in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x tabs</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfTabsToSkip items.
	///
	/// \param[in] numberOfTabsToSkip The number of tabs to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfTabsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTextParagraphTabs
	///
	//@{
	/// \brief <em>Retrieves a \c TextParagraphTab object from the collection</em>
	///
	/// Retrieves a \c TextParagraphTab object from the collection that wraps the tab identified by
	/// \c tabIdentifier.
	///
	/// \param[in] tabIdentifier A value that identifies the tab to be retrieved.
	/// \param[in] tabIdentifierType A value specifying the meaning of \c tabIdentifier. Any of the
	///            values defined by the \c TabIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppTab Receives the tab's \c IRichTextParagraphTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa TextParagraphTab, Add, Remove, Contains, RichTextBoxCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa TextParagraphTab, Add, Remove, Contains, RichTextBoxCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(VARIANT tabIdentifier, TabIdentifierTypeConstants tabIdentifierType = titPosition, IRichTextParagraphTab** ppTab = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c TextParagraphTab objects
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

	/// \brief <em>Adds a tab to the paragraph</em>
	///
	/// Adds a tab with the specified properties at the specified position in the paragraph and returns
	/// a \c TextParagraphTab object wrapping the inserted tab.
	///
	/// \param[in] tabPosition The new tab's position in pixels.
	/// \param[in] tabAlignment The text alignment at the new tab. Any of the values defined by the
	///            \c TabAlignmentConstants enumeration is valid.
	/// \param[in] leaderCharacter The character that is used to fill the space taken by the tab. Any of
	///            the values defined by the \c TabLeaderCharacterConstants enumeration is valid.
	/// \param[out] ppAddedTab Receives the added tab's \c IRichTextParagraphTab implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, TextParagraphTab::get_Position, TextParagraphTab::get_Alignment,
	///       TextParagraphTab::get_LeaderCharacter, RichTextBoxCtlLibU::TabAlignmentConstants,
	///       RichTextBoxCtlLibU::TabLeaderCharacterConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, TextParagraphTab::get_Position, TextParagraphTab::get_Alignment,
	///       TextParagraphTab::get_LeaderCharacter, RichTextBoxCtlLibA::TabAlignmentConstants,
	///       RichTextBoxCtlLibA::TabLeaderCharacterConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(FLOAT tabPosition, TabAlignmentConstants tabAlignment = taLeft, TabLeaderCharacterConstants leaderCharacter = tlcSpaces, IRichTextParagraphTab** ppAddedTab = NULL);
	/// \brief <em>Retrieves whether the specified tab is part of the tab collection</em>
	///
	/// \param[in] tabIdentifier A value that identifies the tab to be checked.
	/// \param[in] tabIdentifierType A value specifying the meaning of \c tabIdentifier. Any of the
	///            values defined by the \c TabIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the tab is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Remove, RichTextBoxCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, RichTextBoxCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(VARIANT tabIdentifier, TabIdentifierTypeConstants tabIdentifierType = titPosition, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the tabs in the collection</em>
	///
	/// Retrieves the number of \c TextParagraphTab objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified tab in the collection from the paragraph</em>
	///
	/// \param[in] tabIdentifier A value that identifies the tab to be removed.
	/// \param[in] tabIdentifierType A value specifying the meaning of \c tabIdentifier. Any of the
	///            values defined by the \c TabIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(VARIANT tabIdentifier, TabIdentifierTypeConstants tabIdentifierType = titPosition);
	/// \brief <em>Removes all tabs in the collection from the paragraph</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given text range paragraph</em>
	///
	/// Attaches this object to a given text range paragraph, so that the text range paragraph's properties
	/// can be retrieved and set using this object's methods.
	///
	/// \param[in] pTextParagraphTab The \c ITextPara object to attach to.
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
	//////////////////////////////////////////////////////////////////////
	/// \name Filter appliance
	///
	//@{
	/// \brief <em>Retrieves the paragraph's first tab that might be in the collection</em>
	///
	/// \param[in] numberOfTabs The number of tabs in the paragraph.
	///
	/// \return The tab being the collection's first candidate. -1 if no tab was found.
	///
	/// \sa GetNextTabToProcess, Count, RemoveAll, Next
	int GetFirstTabToProcess(int numberOfTabs);
	/// \brief <em>Retrieves the paragraph's next tab that might be in the collection</em>
	///
	/// \param[in] previousTab The tab at which to start the search for a new collection candidate.
	/// \param[in] numberOfTabs The number of tabs in the paragraph.
	///
	/// \return The next tab being a candidate for the collection. -1 if no tab was found.
	///
	/// \sa GetFirstTabToProcess, Count, RemoveAll, Next
	int GetNextTabToProcess(int previousTab, int numberOfTabs);
	/// \brief <em>Retrieves whether a tab is part of the collection (applying the filters)</em>
	///
	/// \param[in] tabIndex The tab to check.
	///
	/// \return \c TRUE, if the tab is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(int tabIndex);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ITextPara object corresponding to the text range paragraph wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774056.aspx">ITextPara</a>
		ITextPara* pTextParagraph;
		/// \brief <em>Holds the last enumerated tab</em>
		int lastEnumeratedTab;

		Properties()
		{
			pTextParagraph = NULL;
			lastEnumeratedTab = -1;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
	} properties;
};     // TextParagraphTabs

OBJECT_ENTRY_AUTO(__uuidof(TextParagraphTabs), TextParagraphTabs)