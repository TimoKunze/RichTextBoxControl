//////////////////////////////////////////////////////////////////////
/// \class OLEObjects
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a collection of embedded OLE objects</em>
///
/// This class is a wrapper around a collection of embedded OLE objects.
///
/// \if UNICODE
///   \sa RichTextBoxCtlLibU::IRichOLEObjects, RichTextBox, OLEObject
/// \else
///   \sa RichTextBoxCtlLibA::IRichOLEObjects, RichTextBox, OLEObject
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "RichTextBoxCtlU.h"
#else
	#include "RichTextBoxCtlA.h"
#endif
#include "_IOLEObjectsEvents_CP.h"
#include "helpers.h"
#include "RichTextBox.h"


class ATL_NO_VTABLE OLEObjects : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<OLEObjects, &CLSID_OLEObjects>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<OLEObjects>,
    public Proxy_IRichOLEObjectsEvents<OLEObjects>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IRichOLEObjects, &IID_IRichOLEObjects, &LIBID_RichTextBoxCtlLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IRichOLEObjects, &IID_IRichOLEObjects, &LIBID_RichTextBoxCtlLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_OLEOBJECTS)

		BEGIN_COM_MAP(OLEObjects)
			COM_INTERFACE_ENTRY(IRichOLEObjects)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(OLEObjects)
			CONNECTION_POINT_ENTRY(__uuidof(_IRichOLEObjectsEvents))
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the objects</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c OLEObject objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x objects</em>
	///
	/// Retrieves the next \c numberOfMaxObjects OLE objects from the iterator.
	///
	/// \param[in] numberOfMaxObjects The maximum number of objects the array identified by \c pObjects can
	///            contain.
	/// \param[in,out] pObjects An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to an OLE object's \c IRichOLEObject implementation.
	/// \param[out] pNumberOfObjectsReturned The number of objects that actually were copied to the array
	///             identified by \c pObjects.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, OLEObject,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxObjects, VARIANT* pObjects, ULONG* pNumberOfObjectsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// object in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x objects</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfObjectsToSkip items.
	///
	/// \param[in] numberOfObjectsToSkip The number of objects to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfObjectsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IRichOLEObjects
	///
	//@{
	/// \brief <em>Retrieves a \c OLEObject object from the collection</em>
	///
	/// Retrieves a \c OLEObject object from the collection that wraps the embedded OLE object identified by
	/// \c objectIdentifier.
	///
	/// \param[in] objectIdentifier A value that identifies the embedded OLE object to be retrieved.
	/// \param[in] objectIdentifierType A value specifying the meaning of \c objectIdentifier. Any of the
	///            values defined by the \c ObjectIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppObject Receives the tab's \c IRichOLEObject implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa OLEObject, Add, Remove, Contains, RichTextBoxCtlLibU::ObjectIdentifierTypeConstants
	/// \else
	///   \sa OLEObject, Add, Remove, Contains, RichTextBoxCtlLibA::ObjectIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(VARIANT objectIdentifier, ObjectIdentifierTypeConstants objectIdentifierType = oitIndex, IRichOLEObject** ppObject = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c OLEObject objects managed
	/// by this collection object. This iterator is used by Visual Basic's \c For...Each construct.
	///
	/// \param[out] ppEnumerator A pointer to the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds an embedded OLE object to the control</em>
	///
	/// Adds an embedded OLE object with the specified properties at the specified position into the
	/// document and returns a \c OLEObject object wrapping the inserted object.
	///
	/// \param[in] objectReference Identifies the object to insert. May be:
	///            - A programmatic identifier (e.g. "WordPad.Document.1"). The control will create an
	///              instance of this class.
	///            - A \c CLSID (e.g. "{73FDDC80-AEA9-101A-98A7-00AA00374959}"). The control will create an
	///              instance of this class.
	///            - An object that implements the \c IOleObject interface. The control will use this
	///              object directly.
	/// \param[in] fileToCreateFrom Optionally an absolute path to a file to create the object from. This
	///            parameter will be ignored, if \c objectReference contains an object.
	/// \param[in] insertAsLink If the object is created from a file, set this parameter to \c VARIANT_TRUE
	///            to link the inserted object with the file. Otherwise a copy of the file will be inserted.
	///            If the object is linked to the file, it will reflect changes to the file.
	/// \param[in] characterPosition The zero-based character position at which to insert the new object. If
	///            set to -1, the object will replace the current selection.
	/// \param[in] displayAspect The object's display aspect, i.e. whether it is displayed as an icon or as
	///            an embedded object. Any of the values defined by the \c DisplayAspectConstants
	///            enumeration is valid. Other display aspects as specified on MSDN may work as well.
	/// \param[in] flags A bit field that controls how the OLE object is embedded, e.g. how it is aligned.
	///            Any combination of the values defined by the \c OLEObjectFlagsConstants enumeration is
	///            valid.
	/// \param[in] width The object's width (in himetric). If set to 0, the object defines its size itself.
	/// \param[in] height The object's height (in himetric). If set to 0, the object defines its size itself.
	/// \param[in] userData A \c LONG value associated with the OLE object. Use this property to associate
	///            any data with the object.
	/// \param[out] ppAddedObject Receives the added object's \c IRichOLEObject implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa AddImage, Count, Remove, RemoveAll, OLEObject::get_ClassID, OLEObject::get_ProgID,
	///       OLEObject::get_LinkSource, OLEObject::get_TextRange, OLEObject::get_DisplayAspect,
	///       RichTextBoxCtlLibU::DisplayAspectConstants, OLEObject::get_Flags,
	///       RichTextBoxCtlLibU::OLEObjectFlagsConstants, OLEObject::GetSize, OLEObject::get_UserData,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">Display Aspects</a>
	/// \else
	///   \sa AddImage, Count, Remove, RemoveAll, OLEObject::get_ClassID, OLEObject::get_ProgID,
	///       OLEObject::get_LinkSource, OLEObject::get_TextRange, OLEObject::get_DisplayAspect,
	///       RichTextBoxCtlLibA::DisplayAspectConstants, OLEObject::get_Flags,
	///       RichTextBoxCtlLibA::OLEObjectFlagsConstants, OLEObject::GetSize, OLEObject::get_UserData,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">Display Aspects</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT objectReference = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), BSTR fileToCreateFrom = NULL, VARIANT_BOOL insertAsLink = VARIANT_FALSE, LONG characterPosition = REO_CP_SELECTION, DisplayAspectConstants displayAspect = daContent, OLEObjectFlagsConstants flags = static_cast<OLEObjectFlagsConstants>(oofDynamicSize | oofResizable), OLE_XSIZE_HIMETRIC width = 0, OLE_YSIZE_HIMETRIC height = 0, LONG userData = 0, IRichOLEObject** ppAddedObject = NULL);
	/// \brief <em>Adds an embedded image to the control</em>
	///
	/// Adds an embedded image with the specified properties at the specified position into the document
	/// and returns a \c OLEObject object wrapping the inserted image.
	///
	/// \param[in] imageData Identifies the image to insert. May be:
	///            - A fully qualified path to a file that can be loaded by GDI+ (e.g. JPG, PNG, BMP, GIF,
	///              EMF, WMF). The image will be embedded as bitmap.
	///            - A handle to a metafile, enhanced metafile, DIB, or bitmap.
	///            - An \c IPictureDisp object (VB6 \c StdPicture).
	///            - An object that implements the \c IDataObject interface.
	/// \param[in] characterPosition The zero-based character position at which to insert the new image.
	///            If set to -1, the image will replace the current selection.
	/// \param[in] flags A bit field that controls how the image is embedded, e.g. how it is aligned.
	///            Any combination of the values defined by the \c OLEObjectFlagsConstants enumeration is
	///            valid.
	/// \param[in] width The image's width (in himetric). If set to 0, the image defines its size itself.
	/// \param[in] height The image's height (in himetric). If set to 0, the image defines its size itself.
	/// \param[in] userData A \c LONG value associated with the image. Use this property to associate any
	///            data with the image.
	/// \param[out] ppAddedObject Receives the added image OLE object's \c IRichOLEObject implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, Remove, RemoveAll, OLEObject::get_TextRange, OLEObject::get_Flags,
	///       RichTextBoxCtlLibU::OLEObjectFlagsConstants, OLEObject::GetSize, OLEObject::get_UserData,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	/// \else
	///   \sa Add, Count, Remove, RemoveAll, OLEObject::get_TextRange, OLEObject::get_Flags,
	///       RichTextBoxCtlLibA::OLEObjectFlagsConstants, OLEObject::GetSize, OLEObject::get_UserData,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE AddImage(VARIANT imageData, LONG characterPosition = REO_CP_SELECTION, OLEObjectFlagsConstants flags = static_cast<OLEObjectFlagsConstants>(oofDynamicSize | oofResizable), OLE_XSIZE_HIMETRIC width = 0, OLE_YSIZE_HIMETRIC height = 0, LONG userData = 0, IRichOLEObject** ppAddedObject = NULL);
	/// \brief <em>Retrieves whether the specified object is part of the OLE object collection</em>
	///
	/// \param[in] objectIdentifier A value that identifies the embedded OLE object to be checked.
	/// \param[in] objectIdentifierType A value specifying the meaning of \c objectIdentifier. Any of the
	///            values defined by the \c ObjectIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the OLE object is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Remove, RichTextBoxCtlLibU::ObjectIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, RichTextBoxCtlLibA::ObjectIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(VARIANT objectIdentifier, ObjectIdentifierTypeConstants objectIdentifierType = oitIndex, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the OLE objects in the collection</em>
	///
	/// Retrieves the number of \c OLEObject objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified embedded OLE object in the collection from the document</em>
	///
	/// \param[in] objectIdentifier A value that identifies the embedded OLE object to be removed.
	/// \param[in] objectIdentifierType A value specifying the meaning of \c objectIdentifier. Any of the
	///            values defined by the \c ObjectIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibU::ObjectIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, RichTextBoxCtlLibA::ObjectIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(VARIANT objectIdentifier, ObjectIdentifierTypeConstants objectIdentifierType = oitIndex);
	/// \brief <em>Removes all embedded OLE objects in the collection from the document</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

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
	/// \brief <em>Retrieves the control's first embedded OLE object that might be in the collection</em>
	///
	/// \param[in] numberOfObjects The number of OLE objects in the control.
	///
	/// \return The zero-based index of the OLE object being the collection's first candidate. -1 if no
	///         object was found.
	///
	/// \sa GetNextObjectToProcess, Count, RemoveAll, Next
	LONG GetFirstObjectToProcess(LONG numberOfObjects);
	/// \brief <em>Retrieves the control's next embedded OLE object that might be in the collection</em>
	///
	/// \param[in] previousObjectIndex The OLE object at which to start the search for a new collection
	///            candidate.
	/// \param[in] numberOfObjects The number of OLE objects in the control.
	///
	/// \return The zero-based index of the next OLE object being a candidate for the collection. -1 if no
	///         object was found.
	///
	/// \sa GetFirstObjectToProcess, Count, RemoveAll, Next
	LONG GetNextObjectToProcess(LONG previousObjectIndex, LONG numberOfObjects);
	/// \brief <em>Retrieves whether am embedded OLE object is part of the collection (applying the filters)</em>
	///
	/// \param[in] objectIndex The zero-based index of the OLE object to check.
	///
	/// \return \c TRUE, if the OLE object is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(LONG objectIndex);
	//@}
	//////////////////////////////////////////////////////////////////////

	#ifdef USE_STL
		/// \brief <em>Removes the specified objects</em>
		///
		/// \param[in] objectsToRemove A vector containing all objects to remove.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveObjects(std::list<LONG>& objectsToRemove);
	#else
		/// \brief <em>Removes the specified objects</em>
		///
		/// \param[in] objectsToRemove A vector containing all objects to remove.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveObjects(CAtlList<LONG>& objectsToRemove);
	#endif

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c RichTextBox object that owns these objects</em>
		///
		/// \sa SetOwner
		RichTextBox* pOwnerRTB;
		/// \brief <em>The zero-based index of the last enumerated object</em>
		LONG lastEnumeratedObjectIndex;

		Properties()
		{
			pOwnerRTB = NULL;
			lastEnumeratedObjectIndex = -1;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);
		/// \brief <em>Retrieves the owning rich text box' window handle</em>
		///
		/// \return The window handle of the rich text box that contains these objects.
		///
		/// \sa pOwnerRTB
		HWND GetRTBHWnd(void);
	} properties;
};     // OLEObjects

OBJECT_ENTRY_AUTO(__uuidof(OLEObjects), OLEObjects)