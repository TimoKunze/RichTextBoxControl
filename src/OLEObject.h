//////////////////////////////////////////////////////////////////////
/// \class OLEObject
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an embedded OLE object</em>
///
/// This class is a wrapper around an embedded OLE object.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichOLEObject, RichTextBox, OLEObjects
/// \else
///   \sa RichTextBoxCtlLibA::IRichOLEObject, RichTextBox, OLEObjects
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_IOLEObjectEvents_CP.h"
#include "helpers.h"
#include "TextRange.h"
#include "RichTextBox.h"


class ATL_NO_VTABLE OLEObject : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<OLEObject, &CLSID_OLEObject>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<OLEObject>,
    public Proxy_IRichOLEObjectEvents<OLEObject>,
    #ifdef UNICODE
    	public IDispatchImpl<IRichOLEObject, &IID_IRichOLEObject, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichOLEObject, &IID_IRichOLEObject, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_OLEOBJECT)

		BEGIN_COM_MAP(OLEObject)
			COM_INTERFACE_ENTRY(IRichOLEObject)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(ITextRange2), 0, QueryITextRangeInterface)
			COM_INTERFACE_ENTRY_FUNC(__uuidof(IOleObject), 0, QueryIOleObjectInterface)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(OLEObject)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichOLEObjectEvents))
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
	/// \brief <em>Will be called if this class is \c QueryInterface()'d for \c IOleObject</em>
	///
	/// This method will be called if the class's \c QueryInterface() method is called with
	/// \c IID_IOleObject. We forward the call to the wrapped \c IOleObject implementation.
	///
	/// \param[in] pThis The instance of this class, that the interface is queried from.
	/// \param[in] queriedInterface Should be \c IID_IOleObject.
	/// \param[out] ppImplementation Receives the wrapped object's \c IOleObject implementation.
	/// \param[in] cookie A \c DWORD value specified in the COM interface map.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/dd542709.aspx">IOleObject</a>
	static HRESULT CALLBACK QueryIOleObjectInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichOLEObject
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c ClassID property</em>
	///
	/// Retrieves an OLE object's class identifier, the \c CLSID corresponding to the string identifying the
	/// object to an end user.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ProgID, get_TypeName
	virtual HRESULT STDMETHODCALLTYPE get_ClassID(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayAspect property</em>
	///
	/// Retrieves the embedded OLE object's display aspect, i.e. whether it is displayed as an icon or as an
	/// embedded object. Any of the values defined by the \c DisplayAspectConstants enumeration is valid.
	/// Other display aspects as specified on MSDN may work as well.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisplayAspect, RichTextBoxCtlLibU::DisplayAspectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">Display Aspects</a>
	/// \else
	///   \sa put_DisplayAspect, RichTextBoxCtlLibA::DisplayAspectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">Display Aspects</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisplayAspect(DisplayAspectConstants* pValue);
	/// \brief <em>Sets the \c DisplayAspect property</em>
	///
	/// Sets the embedded OLE object's display aspect, i.e. whether it is displayed as an icon or as an
	/// embedded object. Any of the values defined by the \c DisplayAspectConstants enumeration is valid.
	/// Other display aspects as specified on MSDN may work as well.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisplayAspect, RichTextBoxCtlLibU::DisplayAspectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">Display Aspects</a>
	/// \else
	///   \sa get_DisplayAspect, RichTextBoxCtlLibA::DisplayAspectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">Display Aspects</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisplayAspect(DisplayAspectConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Flags property</em>
	///
	/// Retrieves flags that control how the OLE object is embedded, e.g. how it is aligned. Any combination
	/// of the values defined by the \c OLEObjectFlagsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Flags, RichTextBoxCtlLibU::OLEObjectFlagsConstants
	/// \else
	///   \sa put_Flags, RichTextBoxCtlLibA::OLEObjectFlagsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Flags(OLEObjectFlagsConstants* pValue);
	/// \brief <em>Sets the \c Flags property</em>
	///
	/// Sets flags that control how the OLE object is embedded, e.g. how it is aligned. Any combination
	/// of the values defined by the \c OLEObjectFlagsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Flags, RichTextBoxCtlLibU::OLEObjectFlagsConstants
	/// \else
	///   \sa get_Flags, RichTextBoxCtlLibA::OLEObjectFlagsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Flags(OLEObjectFlagsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index identifying this embedded OLE object.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c LinkSource property</em>
	///
	/// Retrieves the fully qualified path to the source file, if the OLE object is a link.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property will fail if the \c Flags property does not include the \c oofIsLink flag.
	///
	/// \sa put_LinkSource, get_Flags, get_TypeName
	virtual HRESULT STDMETHODCALLTYPE get_LinkSource(BSTR* pValue);
	/// \brief <em>Sets the current setting of the \c LinkSource property</em>
	///
	/// Sets the fully qualified path to the source file, if the OLE object is a link.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property will fail if the \c Flags property does not include the \c oofIsLink flag.
	///
	/// \sa get_LinkSource, put_Flags, get_TypeName
	virtual HRESULT STDMETHODCALLTYPE put_LinkSource(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c ProgID property</em>
	///
	/// Retrieves an OLE object's programmatic identifier, the \c ProgID corresponding to the string
	/// identifying the object to an end user.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IRichOLEObject interface.\n
	///          This property is read-only.
	///
	/// \sa get_ClassID, get_TypeName
	virtual HRESULT STDMETHODCALLTYPE get_ProgID(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c TextRange property</em>
	///
	/// Retrieves a \c TextRange object for the embedded OLE object.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa TextRange
	virtual HRESULT STDMETHODCALLTYPE get_TextRange(IRichTextRange** ppTextRange);
	/// \brief <em>Retrieves the current setting of the \c TypeName property</em>
	///
	/// Retrieves the user-type name of an object for display in user-interface elements such as menus,
	/// list boxes, and dialog boxes.
	///
	/// \param[in] format The kind of type name to retrieve. Any of the values defined by the
	///            \c OLEObjectTypeNameFormatConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ClassID, get_ProgID, RichTextBoxCtlLibU::OLEObjectTypeNameFormatConstants
	/// \else
	///   \sa get_ClassID, get_ProgID, RichTextBoxCtlLibA::OLEObjectTypeNameFormatConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TypeName(OLEObjectTypeNameFormatConstants format, BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c UserData property</em>
	///
	/// Retrieves the \c LONG value associated with the OLE object. Use this property to associate any data
	/// with the object.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UserData
	virtual HRESULT STDMETHODCALLTYPE get_UserData(LONG* pValue);
	/// \brief <em>Sets the \c UserData property</em>
	///
	/// Sets the \c LONG value associated with the OLE object. Use this property to associate any data
	/// with the object.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UserData
	virtual HRESULT STDMETHODCALLTYPE put_UserData(LONG newValue);

	/// \brief <em>Executes the specified verb</em>
	///
	/// Executes the specified verb on the embedded OLE object.
	///
	/// \param[in] verbID The unique identifier of the verb to execute.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetVerbs, RichTextBoxCtlLibU::OLEVERBDETAILS::ID
	/// \else
	///   \sa GetVerbs, RichTextBoxCtlLibA::OLEVERBDETAILS::ID
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE ExecuteVerb(LONG verbID);
	/// \brief <em>Retrieves the OLE object's bounding rectangle</em>
	///
	/// Retrieves the OLE object's bounding rectangle in pixels relative to the control's upper-left corner.
	///
	/// \param[out] pXLeft Receives the x-coordinate (in pixels) of the bounding rectangle's left border.
	/// \param[out] pYTop Receives the y-coordinate (in pixels) of the bounding rectangle's top border.
	/// \param[out] pXRight Receives the x-coordinate (in pixels) of the bounding rectangle's right border.
	/// \param[out] pYBottom Receives the y-coordinate (in pixels) of the bounding rectangle's bottom border.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetSize, SetSize
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Retrieves the OLE object's extent</em>
	///
	/// Retrieves the OLE object's extent (in himetrics).
	///
	/// \param[out] pWidth Receives the OLE object's current width in himetrics.
	/// \param[out] pHeight Receives the OLE object's current height in himetrics.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetSize, GetRectangle
	virtual HRESULT STDMETHODCALLTYPE GetSize(OLE_XSIZE_HIMETRIC* pWidth = NULL, OLE_YSIZE_HIMETRIC* pHeight = NULL);
	/// \brief <em>Retrieves the OLE object's verbs</em>
	///
	/// Retrieves the verbs provided by the object. The client application may execute these verbs to
	/// perform operations on the object.
	///
	/// \param[out] pVerbs Receives an array of \c OLEVERBDETAILS structures, each string defining a verb.
	/// \param[out] pCount Recevies the number of verbs.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExecuteVerb, RichTextBoxCtlLibU::OLEVERBDETAILS
	/// \else
	///   \sa ExecuteVerb, RichTextBoxCtlLibA::OLEVERBDETAILS
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetVerbs(VARIANT* pVerbs, LONG* pCount);
	/// \brief <em>Sets the OLE object's extent</em>
	///
	/// Sets the OLE object's extent (in himetrics).
	///
	/// \param[in] width The OLE object's desired width in himetrics.
	/// \param[in] height The OLE object's desired height in himetrics.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetSize, GetRectangle
	virtual HRESULT STDMETHODCALLTYPE SetSize(OLE_XSIZE_HIMETRIC width, OLE_YSIZE_HIMETRIC height);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given embedded OLE object</em>
	///
	/// Attaches this object to a given embedded OLE object, so that the embedded object's properties can be
	/// retrieved and set using this object's methods.
	///
	/// \param[in] pTextRange The \c ITextRange object of the OLE object to attach to.
	/// \param[in] pOleObject The \c IOleObject object to attach to.
	///
	/// \sa Detach,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd542709.aspx">IOleObject</a>
	void Attach(ITextRange* pTextRange, IOleObject* pOleObject);
	/// \brief <em>Detaches this object from an embedded OLE object</em>
	///
	/// Detaches this object from the embedded OLE object it currently wraps, so that it doesn't wrap any
	/// embedded object anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this object</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerRTB
	void SetOwner(__in_opt RichTextBox* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c RichTextBox object that owns this object</em>
		///
		/// \sa SetOwner
		RichTextBox* pOwnerRTB;
		/// \brief <em>The \c ITextRange object corresponding to the OLE object wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774058.aspx">ITextRange</a>
		ITextRange* pTextRange;
		/// \brief <em>The \c IOleObject object corresponding to the embedded OLE object wrapped by this object</em>
		///
		/// \sa <a href="https://msdn.microsoft.com/en-us/library/dd542709.aspx">IOleObject</a>
		IOleObject* pOleObject;

		Properties()
		{
			pOwnerRTB = NULL;
			pTextRange = NULL;
			pOleObject = NULL;
		}

		~Properties();

		/// \brief <em>Retrieves the owning rich text box' window handle</em>
		///
		/// \return The window handle of the rich text box that contains this object.
		///
		/// \sa pOwnerRTB
		HWND GetRTBHWnd(void);
	} properties;
};     // OLEObject

OBJECT_ENTRY_AUTO(__uuidof(OLEObject), OLEObject)