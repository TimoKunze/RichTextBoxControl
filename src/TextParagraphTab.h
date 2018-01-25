//////////////////////////////////////////////////////////////////////
/// \class TextParagraphTab
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps one of a text range paragraph's tabs</em>
///
/// This class is a wrapper around a specific tab of a text range's paragraph.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichTextParagraphTab, TextParagraphTabs, RichTextBox
/// \else
///   \sa RichTextBoxCtlLibA::IRichTextParagraphTab, TextParagraphTabs, RichTextBox
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_ITextParagraphTabEvents_CP.h"
#include "helpers.h"


class ATL_NO_VTABLE TextParagraphTab : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<TextParagraphTab, &CLSID_TextParagraphTab>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<TextParagraphTab>,
    public Proxy_IRichTextParagraphTabEvents<TextParagraphTab>, 
    #ifdef UNICODE
    	public IDispatchImpl<IRichTextParagraphTab, &IID_IRichTextParagraphTab, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichTextParagraphTab, &IID_IRichTextParagraphTab, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TEXTPARAGRAPHTAB)

		BEGIN_COM_MAP(TextParagraphTab)
			COM_INTERFACE_ENTRY(IRichTextParagraphTab)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(TextParagraphTab)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichTextParagraphTabEvents))
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

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichTextParagraphTab
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Alignment property</em>
	///
	/// Retrieves the horizontal alignment of text at the tab. Any of the values defined by the
	/// \c TabAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Alignment, get_Position, get_LeaderCharacter, TextParagraph::get_HAlignment,
	///       RichTextBoxCtlLibU::TabAlignmentConstants
	/// \else
	///   \sa put_Alignment, get_Position, get_LeaderCharacter, TextParagraph::get_HAlignment,
	///       RichTextBoxCtlLibA::TabAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Alignment(TabAlignmentConstants* pValue);
	/// \brief <em>Sets the \c Alignment property</em>
	///
	/// Sets the horizontal alignment of text at the tab. Any of the values defined by the
	/// \c TabAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Alignment, put_Position, put_LeaderCharacter, TextParagraph::put_HAlignment,
	///       RichTextBoxCtlLibU::TabAlignmentConstants
	/// \else
	///   \sa get_Alignment, put_Position, put_LeaderCharacter, TextParagraph::put_HAlignment,
	///       RichTextBoxCtlLibA::TabAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Alignment(TabAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index identifying this tab.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Changing a tab's position, or removing or adding new tabs may change a tab's index.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Position, RichTextBoxCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa get_Position, RichTextBoxCtlLibA::TabIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c LeaderCharacter property</em>
	///
	/// Retrieves the character that is used to fill the space taken by the tab. Any of the values
	/// defined by the \c TabLeaderCharacterConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_LeaderCharacter, get_Position, get_Alignment,
	///       RichTextBoxCtlLibU::TabLeaderCharacterConstants
	/// \else
	///   \sa put_LeaderCharacter, get_Position, get_Alignment,
	///       RichTextBoxCtlLibA::TabLeaderCharacterConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_LeaderCharacter(TabLeaderCharacterConstants* pValue);
	/// \brief <em>Sets the \c LeaderCharacter property</em>
	///
	/// Sets the character that is used to fill the space taken by the tab. Any of the values
	/// defined by the \c TabLeaderCharacterConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_LeaderCharacter, put_Position, put_Alignment,
	///       RichTextBoxCtlLibU::TabLeaderCharacterConstants
	/// \else
	///   \sa get_LeaderCharacter, put_Position, put_Alignment,
	///       RichTextBoxCtlLibA::TabLeaderCharacterConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_LeaderCharacter(TabLeaderCharacterConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c NextTab property</em>
	///
	/// Retrieves the tab that proceeds this tab.
	///
	/// \param[out] ppTab The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_PreviousTab, get_Index
	virtual HRESULT STDMETHODCALLTYPE get_NextTab(IRichTextParagraphTab** ppTab);
	/// \brief <em>Retrieves the current setting of the \c Position property</em>
	///
	/// Retrieves the tab's position (in points).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Position, get_Index, get_LeaderCharacter, get_Alignment,
	///       RichTextBoxCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa put_Position, get_Index, get_LeaderCharacter, get_Alignment,
	///       RichTextBoxCtlLibA::TabLeaderCharacterConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Position(FLOAT* pValue);
	/// \brief <em>Sets the \c Position property</em>
	///
	/// Sets the tab's position (in points).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Position, get_Index, put_LeaderCharacter, put_Alignment,
	///       RichTextBoxCtlLibU::TabIdentifierTypeConstants
	/// \else
	///   \sa get_Position, get_Index, put_LeaderCharacter, put_Alignment,
	///       RichTextBoxCtlLibA::TabLeaderCharacterConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Position(FLOAT newValue);
	/// \brief <em>Retrieves the current setting of the \c PreviousTab property</em>
	///
	/// Retrieves the tab that precedes this tab.
	///
	/// \param[out] ppTab The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_NextTab, get_Index
	virtual HRESULT STDMETHODCALLTYPE get_PreviousTab(IRichTextParagraphTab** ppTab);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given text range paragraph tab</em>
	///
	/// Attaches this object to a given text range paragraph tab, so that the text range paragraph tab's
	/// properties can be retrieved and set using this object's methods.
	///
	/// \param[in] pTextParagraph The \c ITextPara object to attach to.
	/// \param[in] tabPosition The position of the tab to attach to.
	///
	/// \sa Detach,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774056.aspx">ITextPara</a>
	void Attach(ITextPara* pTextParagraph, float tabPosition);
	/// \brief <em>Detaches this object from a text range paragraph tab</em>
	///
	/// Detaches this object from the text range paragraph tab it currently wraps, so that it doesn't wrap
	/// any text range paragraph tab anymore.
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
		/// \brief <em>The position of the tab wrapped by this object</em>
		float tabPosition;

		Properties()
		{
			pTextParagraph = NULL;
			tabPosition = 0.0f;
		}

		~Properties();
	} properties;
};     // TextParagraphTab

OBJECT_ENTRY_AUTO(__uuidof(TextParagraphTab), TextParagraphTab)