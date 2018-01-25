// OLEObjects.cpp: A wrapper for a collection of embedded OLE objects.

#include "stdafx.h"
#include "OLEObjects.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP OLEObjects::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichOLEObjects, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP OLEObjects::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<OLEObjects>* pOLEObjectsObj = NULL;
	CComObject<OLEObjects>::CreateInstance(&pOLEObjectsObj);
	pOLEObjectsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pOLEObjectsObj->properties);

	pOLEObjectsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pOLEObjectsObj->Release();
	return S_OK;
}

STDMETHODIMP OLEObjects::Next(ULONG numberOfMaxObjects, VARIANT* pObjects, ULONG* pNumberOfObjectsReturned)
{
	ATLASSERT_POINTER(pObjects, VARIANT);
	if(!pObjects) {
		return E_POINTER;
	}

	// check each object in the control
	LONG numberOfObjects = 0;
	HWND hWndRTB = properties.GetRTBHWnd();
	CComPtr<IRichEditOle> pRichEditOle = NULL;
	if(IsWindow(hWndRTB)) {
		if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			numberOfObjects = pRichEditOle->GetObjectCount();
		} else {
			return E_FAIL;
		}
	} else {
		return E_FAIL;
	}
	ULONG i = 0;
	for(i = 0; i < numberOfMaxObjects; ++i) {
		VariantInit(&pObjects[i]);

		do {
			if(properties.lastEnumeratedObjectIndex >= 0) {
				if(properties.lastEnumeratedObjectIndex < numberOfObjects) {
					properties.lastEnumeratedObjectIndex = GetNextObjectToProcess(properties.lastEnumeratedObjectIndex, numberOfObjects);
				}
			} else {
				properties.lastEnumeratedObjectIndex = GetFirstObjectToProcess(numberOfObjects);
			}
			if(properties.lastEnumeratedObjectIndex >= numberOfObjects) {
				properties.lastEnumeratedObjectIndex = -1;
			}
		} while((properties.lastEnumeratedObjectIndex != -1) && (!IsPartOfCollection(properties.lastEnumeratedObjectIndex)));

		if(properties.lastEnumeratedObjectIndex != -1) {
			REOBJECT objectDetails = {0};
			objectDetails.cbStruct = sizeof(REOBJECT);
			if(SUCCEEDED(pRichEditOle->GetObject(properties.lastEnumeratedObjectIndex, &objectDetails, REO_GETOBJ_POLEOBJ))) {
				CComPtr<IRichTextRange> pRichTextRange = NULL;
				CComPtr<ITextRange> pTextRange = NULL;
				if(SUCCEEDED(properties.pOwnerRTB->get_TextRange(objectDetails.cp, objectDetails.cp, &pRichTextRange)) && pRichTextRange && SUCCEEDED(pRichTextRange->QueryInterface(IID_PPV_ARGS(&pTextRange))) && pTextRange) {
					pTextRange->EndOf(tomObject, tomExtend, NULL);
				}
				ClassFactory::InitOLEObject(pTextRange, objectDetails.poleobj, properties.pOwnerRTB, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pObjects[i].pdispVal));
				pObjects[i].vt = VT_DISPATCH;
				if(objectDetails.poleobj) {
					objectDetails.poleobj->Release();
					objectDetails.poleobj = NULL;
				}
			} else {
				i--;
				break;
			}
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfObjectsReturned) {
		*pNumberOfObjectsReturned = i;
	}

	return (i == numberOfMaxObjects ? S_OK : S_FALSE);
}

STDMETHODIMP OLEObjects::Reset(void)
{
	properties.lastEnumeratedObjectIndex = -1;
	return S_OK;
}

STDMETHODIMP OLEObjects::Skip(ULONG numberOfObjectsToSkip)
{
	VARIANT dummy;
	ULONG numObjectsReturned = 0;
	// we could skip all rows at once, but it's easier to skip them one by one
	for(ULONG i = 1; i <= numberOfObjectsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numObjectsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numObjectsReturned == 0) {
			// there're no more objects to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////

LONG OLEObjects::GetFirstObjectToProcess(LONG numberOfObjects)
{
	if(numberOfObjects == 0) {
		return -1;
	}
	return 0;
}

LONG OLEObjects::GetNextObjectToProcess(LONG previousObjectIndex, LONG numberOfObjects)
{
	if(previousObjectIndex < numberOfObjects - 1) {
		return previousObjectIndex + 1;
	} else {
		return -1;
	}
}

BOOL OLEObjects::IsPartOfCollection(LONG objectIndex)
{
	HWND hWndRTB = properties.GetRTBHWnd();
	if(IsWindow(hWndRTB)) {
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			return (objectIndex < pRichEditOle->GetObjectCount());
		}
	}
	return FALSE;
}

#ifdef USE_STL
	HRESULT OLEObjects::RemoveObjects(std::list<LONG>& objectsToRemove)
#else
	HRESULT OLEObjects::RemoveObjects(CAtlList<LONG>& objectsToRemove)
#endif
{
	VARIANT v;
	VariantInit(&v);
	v.vt = VT_I4;

	// sort in reverse order
	#ifdef USE_STL
		objectsToRemove.sort(std::greater<LONG>());
		for(std::list<LONG>::const_iterator iter = objectsToRemove.begin(); iter != objectsToRemove.end(); ++iter) {
			v.intVal = *iter;
			Remove(v, oitIndex);
		}
	#else
		// perform a crude bubble sort
		for(size_t j = 0; j < objectsToRemove.GetCount(); ++j) {
			for(size_t i = 0; i < objectsToRemove.GetCount() - 1; ++i) {
				if(objectsToRemove.GetAt(objectsToRemove.FindIndex(i)) < objectsToRemove.GetAt(objectsToRemove.FindIndex(i + 1))) {
					objectsToRemove.SwapElements(objectsToRemove.FindIndex(i), objectsToRemove.FindIndex(i + 1));
				}
			}
		}

		for(size_t i = 0; i < objectsToRemove.GetCount(); ++i) {
			v.intVal = objectsToRemove.GetAt(objectsToRemove.FindIndex(i));
			Remove(v, oitIndex);
		}
	#endif
	VariantClear(&v);

	return S_OK;
}


OLEObjects::Properties::~Properties()
{
	if(pOwnerRTB) {
		pOwnerRTB->Release();
		pOwnerRTB = NULL;
	}
}

void OLEObjects::Properties::CopyTo(OLEObjects::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerRTB = this->pOwnerRTB;
		if(pTarget->pOwnerRTB) {
			pTarget->pOwnerRTB->AddRef();
		}
		pTarget->lastEnumeratedObjectIndex = this->lastEnumeratedObjectIndex;
	}
}

HWND OLEObjects::Properties::GetRTBHWnd(void)
{
	ATLASSUME(pOwnerRTB);

	OLE_HANDLE handle = NULL;
	pOwnerRTB->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

void OLEObjects::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP OLEObjects::get_Item(VARIANT objectIdentifier, ObjectIdentifierTypeConstants objectIdentifierType/* = oitIndex*/, IRichOLEObject** ppObject/* = NULL*/)
{
	ATLASSERT_POINTER(ppObject, IRichOLEObject*);
	if(!ppObject) {
		return E_POINTER;
	}

	// retrieve the object's index
	LONG objectIndex = -1;
	HRESULT hr = E_FAIL;
	VARIANT v;
	VariantInit(&v);
	switch(objectIdentifierType) {
		case oitIndex:
			if(SUCCEEDED(VariantChangeType(&v, &objectIdentifier, 0, VT_I4))) {
				objectIndex = v.intVal;
				hr = S_OK;
			}
			break;
	}
	VariantClear(&v);

	if(SUCCEEDED(hr) && hr != S_FALSE && IsPartOfCollection(objectIndex)) {
		HWND hWndRTB = properties.GetRTBHWnd();
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(IsWindow(hWndRTB)) {
			if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);
				if(SUCCEEDED(pRichEditOle->GetObject(objectIndex, &objectDetails, REO_GETOBJ_POLEOBJ))) {
					CComPtr<IRichTextRange> pRichTextRange = NULL;
					CComPtr<ITextRange> pTextRange = NULL;
					if(SUCCEEDED(properties.pOwnerRTB->get_TextRange(objectDetails.cp, objectDetails.cp, &pRichTextRange)) && pRichTextRange && SUCCEEDED(pRichTextRange->QueryInterface(IID_PPV_ARGS(&pTextRange))) && pTextRange) {
						pTextRange->EndOf(tomObject, tomExtend, NULL);
					}
					ClassFactory::InitOLEObject(pTextRange, objectDetails.poleobj, properties.pOwnerRTB, IID_IRichOLEObject, reinterpret_cast<LPUNKNOWN*>(ppObject));
					if(objectDetails.poleobj) {
						objectDetails.poleobj->Release();
						objectDetails.poleobj = NULL;
					}
					if(*ppObject) {
						return S_OK;
					}
				}
			}
		}
	}

	// object not found
	if(objectIdentifierType == oitIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP OLEObjects::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP OLEObjects::Add(VARIANT objectReference/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, BSTR fileToCreateFrom/* = NULL*/, VARIANT_BOOL insertAsLink/* = VARIANT_FALSE*/, LONG characterPosition/* = REO_CP_SELECTION*/, DisplayAspectConstants displayAspect/* = daContent*/, OLEObjectFlagsConstants flags/* = static_cast<OLEObjectFlagsConstants>(oofDynamicSize | oofResizable)*/, OLE_XSIZE_HIMETRIC width/* = 0*/, OLE_YSIZE_HIMETRIC height/* = 0*/, LONG userData/* = 0*/, IRichOLEObject** ppAddedObject/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedObject, IRichOLEObject*);
	if(!ppAddedObject) {
		return E_POINTER;
	}

	*ppAddedObject = NULL;
	HRESULT hr = E_FAIL;
	HWND hWndRTB = properties.GetRTBHWnd();
	if(IsWindow(hWndRTB)) {
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			REOBJECT objectDetails = {0};
			objectDetails.cbStruct = sizeof(REOBJECT);

			CLSID clsid = CLSID_NULL;
			CComPtr<ILockBytes> pLockBytes = NULL;
			if(SUCCEEDED(CreateILockBytesOnHGlobal(NULL, TRUE, &pLockBytes))) {
				if(SUCCEEDED(StgCreateDocfileOnILockBytes(pLockBytes, STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_READWRITE, 0, &objectDetails.pstg))) {
					if(SUCCEEDED(pRichEditOle->GetClientSite(&objectDetails.polesite))) {
						if(fileToCreateFrom && SysStringLen(fileToCreateFrom) > 0) {
							if(insertAsLink != VARIANT_FALSE) {
								hr = OleCreateLinkToFile(fileToCreateFrom, IID_IOleObject, OLERENDER_DRAW, NULL, objectDetails.polesite, objectDetails.pstg, reinterpret_cast<LPVOID*>(&objectDetails.poleobj));
							} else {
								if(objectReference.vt == VT_BSTR) {
									if(FAILED(CLSIDFromString(objectReference.bstrVal, &clsid))) {
										CLSIDFromProgID(objectReference.bstrVal, &clsid);
									}
								}
								hr = OleCreateFromFile(clsid, fileToCreateFrom, IID_IOleObject, OLERENDER_DRAW, NULL, objectDetails.polesite, objectDetails.pstg, reinterpret_cast<LPVOID*>(&objectDetails.poleobj));
							}
							if(SUCCEEDED(hr) && objectDetails.poleobj && clsid == CLSID_NULL) {
								objectDetails.poleobj->GetUserClassID(&clsid);
							}
						} else if(objectReference.vt == VT_BSTR) {
							if(SUCCEEDED(CLSIDFromString(objectReference.bstrVal, &clsid)) || SUCCEEDED(CLSIDFromProgID(objectReference.bstrVal, &clsid))) {
								hr = OleCreate(clsid, IID_IOleObject, OLERENDER_DRAW, NULL, objectDetails.polesite, objectDetails.pstg, reinterpret_cast<LPVOID*>(&objectDetails.poleobj));
							}
						} else if(objectReference.vt == VT_DISPATCH && objectReference.pdispVal) {
							if(SUCCEEDED(objectReference.pdispVal->QueryInterface(IID_PPV_ARGS(&objectDetails.poleobj))) && objectDetails.poleobj) {
								hr = objectDetails.poleobj->GetUserClassID(&clsid);
							}
						} else if(objectReference.vt == VT_UNKNOWN && objectReference.punkVal) {
							if(SUCCEEDED(objectReference.punkVal->QueryInterface(IID_PPV_ARGS(&objectDetails.poleobj))) && objectDetails.poleobj) {
								hr = objectDetails.poleobj->GetUserClassID(&clsid);
							}
						}
					}
				}
			}
			if(!objectDetails.pstg || !objectDetails.polesite) {
				hr = E_FAIL;
			} else if(!objectDetails.poleobj || clsid == CLSID_NULL) {
				hr = E_INVALIDARG;
			} else {
				hr = E_FAIL;
				OleSetContainedObject(objectDetails.poleobj, TRUE);
				objectDetails.clsid = clsid;
				objectDetails.cp = characterPosition;
				objectDetails.dvaspect = displayAspect;
				objectDetails.dwFlags = flags;
				objectDetails.dwUser = userData;
				objectDetails.sizel.cx = width;
				objectDetails.sizel.cy = height;

				if(objectDetails.dvaspect == DVASPECT_ICON) {
					FORMATETC fmt = {0};
					fmt.dwAspect = objectDetails.dvaspect;
					fmt.lindex = -1;
					CComQIPtr<IOleCache> pCache = objectDetails.poleobj;
					if(pCache && SUCCEEDED(pCache->Cache(&fmt, objectDetails.dvaspect == DVASPECT_ICON ? ADVF_NODATA : ADVF_DATAONSTOP, NULL))) {
						fmt.cfFormat = CF_METAFILEPICT;
						fmt.dwAspect = objectDetails.dvaspect;
						fmt.lindex = -1;
						fmt.tymed = TYMED_MFPICT;

						STGMEDIUM stg = {0};
						clsid = CLSID_NULL;
						objectDetails.poleobj->GetUserClassID(&clsid);
						if(fileToCreateFrom && SysStringLen(fileToCreateFrom) > 0) {
							stg.hMetaFilePict = OleGetIconOfClass(clsid, fileToCreateFrom, FALSE);
						} else {
							stg.hMetaFilePict = OleGetIconOfClass(clsid, NULL, TRUE);
						}
						pCache->SetData(&fmt, &stg, TRUE);
					}
				}

				hr = pRichEditOle->InsertObject(&objectDetails);
				if(SUCCEEDED(hr)) {
					CComPtr<IRichTextRange> pRichTextRange = NULL;
					CComPtr<ITextRange> pTextRange = NULL;
					if(SUCCEEDED(properties.pOwnerRTB->get_TextRange(objectDetails.cp, objectDetails.cp, &pRichTextRange)) && pRichTextRange && SUCCEEDED(pRichTextRange->QueryInterface(IID_PPV_ARGS(&pTextRange))) && pTextRange) {
						pTextRange->EndOf(tomObject, tomExtend, NULL);
					}
					ClassFactory::InitOLEObject(pTextRange, objectDetails.poleobj, properties.pOwnerRTB, IID_IRichOLEObject, reinterpret_cast<LPUNKNOWN*>(ppAddedObject));
					// should be done by the client application
					/*if(*ppAddedObject) {
						(*ppAddedObject)->ExecuteVerb(OLEIVERB_PRIMARY);
					}*/
				}
			}
			
			if(objectDetails.poleobj) {
				objectDetails.poleobj->Release();
				objectDetails.poleobj = NULL;
			}
			if(objectDetails.polesite) {
				objectDetails.polesite->Release();
				objectDetails.polesite = NULL;
			}
			if(objectDetails.pstg) {
				objectDetails.pstg->Release();
				objectDetails.pstg = NULL;
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObjects::AddImage(VARIANT imageData, LONG characterPosition/* = REO_CP_SELECTION*/, OLEObjectFlagsConstants flags/* = static_cast<OLEObjectFlagsConstants>(oofDynamicSize | oofResizable)*/, OLE_XSIZE_HIMETRIC width/* = 0*/, OLE_YSIZE_HIMETRIC height/* = 0*/, LONG userData/* = 0*/, IRichOLEObject** ppAddedObject/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedObject, IRichOLEObject*);
	if(!ppAddedObject) {
		return E_POINTER;
	}

	*ppAddedObject = NULL;
	HRESULT hr = E_FAIL;
	BOOL useFallback = TRUE;

	CComPtr<IRichTextRange> pRichTextRange = NULL;
	if(characterPosition == REO_CP_SELECTION) {
		properties.pOwnerRTB->get_SelectedTextRange(&pRichTextRange);
	} else {
		properties.pOwnerRTB->get_TextRange(characterPosition, characterPosition, &pRichTextRange);
	}
	// pTextRange2->InsertImage always fails with E_INVALIDARG - maybe the stream must contain a plain bitmap without file header?
	// EMF files do not seem to be supported at all.
	/*if(pRichTextRange) {
		CComQIPtr<ITextRange2> pTextRange2 = pRichTextRange;
		if(pTextRange2) {
			CComPtr<IStream> pStream = NULL;
			if(imageData.vt == VT_BSTR) {
				SHCreateStreamOnFileEx(OLE2W(imageData.bstrVal), STGM_READ | STGM_SHARE_DENY_WRITE, GENERIC_READ, FALSE, NULL, &pStream);
			} else if(imageData.vt == VT_DISPATCH) {
				imageData.pdispVal->QueryInterface(IID_PPV_ARGS(&pStream));
			} else if(imageData.vt == VT_UNKNOWN) {
				imageData.punkVal->QueryInterface(IID_PPV_ARGS(&pStream));
			}
			if(pStream) {
				hr = pTextRange2->InsertImage(width, height, 0, TA_TOP, NULL, pStream);
				if(SUCCEEDED(hr)) {
					useFallback = FALSE;
					pRichTextRange->get_EmbeddedObject(ppAddedObject);
				}
			}
		}
	}*/

	// NOTE: See revisions around 160 for attempts to make this work without GDI+
	CComPtr<IDataObject> pDataObject = NULL;
	if(useFallback && imageData.vt == VT_BSTR) {
		HBITMAP hBmp = NULL;
		Gdiplus::Bitmap* pBitmap = Gdiplus::Bitmap::FromFile(OLE2W(imageData.bstrVal), TRUE);
		if(pBitmap) {
			pBitmap->GetHBITMAP(Gdiplus::Color::Transparent, &hBmp);
			delete pBitmap;
		}

		//CLIPFORMAT usePasteHackFormat = 0;
		if(hBmp) {
			CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
			CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
			pOLEDataObjectObj->AddRef();
			pOLEDataObjectObj->QueryInterface(IID_PPV_ARGS(&pDataObject));
			pOLEDataObjectObj->Release();

			FORMATETC format = {CF_BITMAP, NULL, DVASPECT_CONTENT, -1, TYMED_GDI};
			STGMEDIUM storageMedium = {0};
			storageMedium.tymed = format.tymed;
			storageMedium.hBitmap = hBmp;
			hr = pDataObject->SetData(&format, &storageMedium, FALSE);
			//usePasteHackFormat = format.cfFormat;
		}
		/*if(SUCCEEDED(hr) && pDataObject && usePasteHackFormat != 0) {
			CComQIPtr<ITextRange> pTextRange = pRichTextRange;
			if(pTextRange) {
				VARIANT v;
				VariantInit(&v);
				pDataObject->QueryInterface(IID_PPV_ARGS(&v.punkVal));
				v.vt = VT_UNKNOWN;
				hr = pTextRange->Paste(&v, usePasteHackFormat);
				VariantClear(&v);
				if(SUCCEEDED(hr)) {
					useFallback = FALSE;
					pRichTextRange->MoveStartByUnit(uCharacter, -1, NULL);
					// NOTE: For background images this may fail!
					HRESULT hr2 = pRichTextRange->get_EmbeddedObject(ppAddedObject);
					if(hr2 == S_FALSE && (objectDetails.dwFlags & REO_USEASBACKGROUND)) {
						ClassFactory::InitOLEObject(NULL, objectDetails.poleobj, properties.pOwnerRTB, IID_IRichOLEObject, reinterpret_cast<LPUNKNOWN*>(ppAddedObject));
					} else {
						hr = hr2;
					}
					pRichTextRange->MoveStartByUnit(uCharacter, 1, NULL);
					(*ppAddedObject)->put_DisplayAspect(daContent);
					(*ppAddedObject)->put_Flags(flags);
					(*ppAddedObject)->put_UserData(userData);
					if(width != 0 && height != 0) {
						(*ppAddedObject)->SetSize(width, height);
					} else if(width != 0) {
						(*ppAddedObject)->GetSize(NULL, &height);
						(*ppAddedObject)->SetSize(width, height);
					} else if(height != 0) {
						(*ppAddedObject)->GetSize(&width, NULL);
						(*ppAddedObject)->SetSize(width, height);
					}
				}
			}
		}*/
	}

	if(useFallback) {
		hr = E_FAIL;
		HWND hWndRTB = properties.GetRTBHWnd();
		if(IsWindow(hWndRTB)) {
			CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);

				CComPtr<IPicture> pPicture = NULL;
				CLSID clsid = CLSID_NULL;
				CComPtr<ILockBytes> pLockBytes = NULL;
				if(SUCCEEDED(CreateILockBytesOnHGlobal(NULL, TRUE, &pLockBytes))) {
					if(SUCCEEDED(StgCreateDocfileOnILockBytes(pLockBytes, STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_READWRITE, 0, &objectDetails.pstg))) {
						if(SUCCEEDED(pRichEditOle->GetClientSite(&objectDetails.polesite))) {
							if(imageData.vt == VT_DISPATCH && imageData.pdispVal) {
								if(FAILED(imageData.pdispVal->QueryInterface(IID_PPV_ARGS(&pDataObject)))) {
									imageData.pdispVal->QueryInterface(IID_PPV_ARGS(&pPicture));
								}
							} else if(imageData.vt == VT_UNKNOWN && imageData.punkVal) {
								if(FAILED(imageData.punkVal->QueryInterface(IID_PPV_ARGS(&pDataObject)))) {
									imageData.punkVal->QueryInterface(IID_PPV_ARGS(&pPicture));
								}
							}

							if(pDataObject) {
								hr = S_OK;
							} else if(pPicture) {
								OLE_HANDLE handle = NULL;
								hr = E_FAIL;
								if(SUCCEEDED(pPicture->get_Handle(&handle))) {
									CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
									CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
									pOLEDataObjectObj->AddRef();
									pOLEDataObjectObj->QueryInterface(IID_PPV_ARGS(&pDataObject));
									pOLEDataObjectObj->Release();

									FORMATETC format = {0, NULL, DVASPECT_CONTENT, -1, 0};
									STGMEDIUM storageMedium = {0};
									SHORT type = PICTYPE_BITMAP;
									if(SUCCEEDED(pPicture->get_Type(&type))) {
										switch(type) {
											case PICTYPE_BITMAP:
												format.cfFormat = CF_BITMAP;
												format.tymed = TYMED_GDI;
												storageMedium.tymed = format.tymed;
												storageMedium.hBitmap = static_cast<HBITMAP>(LongToHandle(handle));
												hr = pDataObject->SetData(&format, &storageMedium, FALSE);
												break;
											case PICTYPE_ENHMETAFILE:
												format.cfFormat = CF_ENHMETAFILE;
												format.tymed = TYMED_ENHMF;
												storageMedium.tymed = format.tymed;
												storageMedium.hEnhMetaFile = static_cast<HENHMETAFILE>(LongToHandle(handle));
												hr = pDataObject->SetData(&format, &storageMedium, FALSE);
												break;
											case PICTYPE_METAFILE:
												format.cfFormat = CF_METAFILEPICT;
												format.tymed = TYMED_MFPICT;
												storageMedium.tymed = format.tymed;
												storageMedium.hMetaFilePict = static_cast<HMETAFILE>(LongToHandle(handle));
												hr = pDataObject->SetData(&format, &storageMedium, FALSE);
												break;
										}
									}
								}
							}

							if(!pDataObject) {
								hr = E_FAIL;
								CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
								CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
								pOLEDataObjectObj->AddRef();
								pOLEDataObjectObj->QueryInterface(IID_PPV_ARGS(&pDataObject));
								pOLEDataObjectObj->Release();

								VARIANT v;
								VariantInit(&v);
								if(SUCCEEDED(VariantChangeType(&v, &imageData, 0, VT_I4))) {
									FORMATETC format = {0, NULL, DVASPECT_CONTENT, -1, 0};
									STGMEDIUM storageMedium = {0};

									HGDIOBJ hImage = LongToHandle(v.lVal);
									switch(GetObjectType(hImage)) {
										case OBJ_BITMAP:
											format.cfFormat = CF_BITMAP;
											format.tymed = TYMED_GDI;
											storageMedium.tymed = format.tymed;
											storageMedium.hBitmap = static_cast<HBITMAP>(hImage);
											hr = pDataObject->SetData(&format, &storageMedium, FALSE);
											break;
										case OBJ_ENHMETAFILE:
											format.cfFormat = CF_ENHMETAFILE;
											format.tymed = TYMED_ENHMF;
											storageMedium.tymed = format.tymed;
											storageMedium.hEnhMetaFile = static_cast<HENHMETAFILE>(hImage);
											hr = pDataObject->SetData(&format, &storageMedium, FALSE);
											break;
										case OBJ_METAFILE:
											format.cfFormat = CF_METAFILEPICT;
											format.tymed = TYMED_MFPICT;
											storageMedium.tymed = format.tymed;
											storageMedium.hEnhMetaFile = static_cast<HENHMETAFILE>(hImage);
											hr = pDataObject->SetData(&format, &storageMedium, FALSE);
											break;
									}
								}
								VariantClear(&v);
							}

							if(SUCCEEDED(hr) && pDataObject) {
								hr = OleCreateStaticFromData(pDataObject, IID_IOleObject, OLERENDER_DRAW, NULL, objectDetails.polesite, objectDetails.pstg, reinterpret_cast<LPVOID*>(&objectDetails.poleobj));
								if(SUCCEEDED(hr) && objectDetails.poleobj && clsid == CLSID_NULL) {
									objectDetails.poleobj->GetUserClassID(&clsid);
								}
							}
						}
					}
				}
				if(*ppAddedObject) {
					// nothing to do
				} else if(!objectDetails.pstg || !objectDetails.polesite) {
					hr = E_FAIL;
				} else if(!objectDetails.poleobj || clsid == CLSID_NULL) {
					hr = E_INVALIDARG;
				} else {
					hr = E_FAIL;
					OleSetContainedObject(objectDetails.poleobj, TRUE);
					objectDetails.clsid = clsid;
					objectDetails.cp = characterPosition;
					objectDetails.dvaspect = DVASPECT_CONTENT;
					objectDetails.dwFlags = flags;
					objectDetails.dwUser = userData;
					objectDetails.sizel.cx = width;
					objectDetails.sizel.cy = height;

					if(objectDetails.dvaspect == DVASPECT_ICON) {
						FORMATETC fmt = {0};
						fmt.dwAspect = objectDetails.dvaspect;
						fmt.lindex = -1;
						CComQIPtr<IOleCache> pCache = objectDetails.poleobj;
						if(pCache && SUCCEEDED(pCache->Cache(&fmt, objectDetails.dvaspect == DVASPECT_ICON ? ADVF_NODATA : ADVF_DATAONSTOP, NULL))) {
							fmt.cfFormat = CF_METAFILEPICT;
							fmt.dwAspect = objectDetails.dvaspect;
							fmt.lindex = -1;
							fmt.tymed = TYMED_MFPICT;

							STGMEDIUM stg = {0};
							stg.hMetaFilePict = OleGetIconOfClass(clsid, NULL, TRUE);
							pCache->SetData(&fmt, &stg, TRUE);
						}
					}

					hr = pRichEditOle->InsertObject(&objectDetails);
					if(SUCCEEDED(hr)) {
						pRichTextRange->MoveStartByUnit(uCharacter, -1, NULL);
						// NOTE: For background images this may fail!
						HRESULT hr2 = pRichTextRange->get_EmbeddedObject(ppAddedObject);
						if(hr2 == S_FALSE && (objectDetails.dwFlags & REO_USEASBACKGROUND)) {
							ClassFactory::InitOLEObject(NULL, objectDetails.poleobj, properties.pOwnerRTB, IID_IRichOLEObject, reinterpret_cast<LPUNKNOWN*>(ppAddedObject));
						} else {
							hr = hr2;
						}
						pRichTextRange->MoveStartByUnit(uCharacter, 1, NULL);
						// should be done by the client application
						/*if(*ppAddedObject) {
							(*ppAddedObject)->ExecuteVerb(OLEIVERB_PRIMARY);
						}*/
					}
				}
			
				if(objectDetails.poleobj) {
					objectDetails.poleobj->Release();
					objectDetails.poleobj = NULL;
				}
				if(objectDetails.polesite) {
					objectDetails.polesite->Release();
					objectDetails.polesite = NULL;
				}
				if(objectDetails.pstg) {
					objectDetails.pstg->Release();
					objectDetails.pstg = NULL;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObjects::Contains(VARIANT objectIdentifier, ObjectIdentifierTypeConstants objectIdentifierType/* = oitIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VARIANT_FALSE;

	// retrieve the object's index
	LONG objectIndex = -1;
	HRESULT hr = E_FAIL;
	VARIANT v;
	VariantInit(&v);
	switch(objectIdentifierType) {
		case oitIndex:
			if(SUCCEEDED(VariantChangeType(&v, &objectIdentifier, 0, VT_I4))) {
				objectIndex = v.intVal;
				hr = S_OK;
			}
			break;
	}
	VariantClear(&v);

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(objectIndex));
	return S_OK;
}

STDMETHODIMP OLEObjects::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRTB = properties.GetRTBHWnd();
	if(IsWindow(hWndRTB)) {
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			*pValue = pRichEditOle->GetObjectCount();
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP OLEObjects::Remove(VARIANT objectIdentifier, ObjectIdentifierTypeConstants objectIdentifierType/* = oitIndex*/)
{
	// retrieve the object's index
	LONG objectIndex = -1;
	HRESULT hr = E_FAIL;
	VARIANT v;
	VariantInit(&v);
	switch(objectIdentifierType) {
		case oitIndex:
			if(SUCCEEDED(VariantChangeType(&v, &objectIdentifier, 0, VT_I4))) {
				objectIndex = v.intVal;
				hr = S_OK;
			}
			break;
	}
	VariantClear(&v);

	if(SUCCEEDED(hr) && hr != S_FALSE && IsPartOfCollection(objectIndex)) {
		HWND hWndRTB = properties.GetRTBHWnd();
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(IsWindow(hWndRTB)) {
			if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);
				if(SUCCEEDED(pRichEditOle->GetObject(objectIndex, &objectDetails, REO_GETOBJ_POLEOBJ))) {
					CComPtr<IRichTextRange> pRichTextRange = NULL;
					CComPtr<ITextRange> pTextRange = NULL;
					if(SUCCEEDED(properties.pOwnerRTB->get_TextRange(objectDetails.cp, objectDetails.cp, &pRichTextRange)) && pRichTextRange && SUCCEEDED(pRichTextRange->QueryInterface(IID_PPV_ARGS(&pTextRange))) && pTextRange) {
						pTextRange->EndOf(tomObject, tomExtend, NULL);
					}
					hr = pTextRange->Delete(tomObject, 1, NULL);
					if(objectDetails.poleobj) {
						objectDetails.poleobj->Release();
						objectDetails.poleobj = NULL;
					}
					return hr;
				}
			}
		}
		return E_FAIL;
	} else {
		// object not found
		if(objectIdentifierType == oitIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}
}

STDMETHODIMP OLEObjects::RemoveAll(void)
{
	LONG numberOfObjects = 0;
	HRESULT hr = Count(&numberOfObjects);
	if(FAILED(hr)) {
		return hr;
	}
	#ifdef USE_STL
		std::list<LONG> objectsToRemove;
	#else
		CAtlList<LONG> objectsToRemove;
	#endif
	for(LONG objectIndex = numberOfObjects - 1; objectIndex >= 0; objectIndex--) {
		#ifdef USE_STL
			objectsToRemove.push_back(objectIndex);
		#else
			objectsToRemove.AddTail(objectIndex);
		#endif
	}
	/*LONG objectIndex = GetFirstObjectToProcess(numberOfObjects);
	while(objectIndex != -1) {
		if(IsPartOfCollection(objectIndex)) {
			#ifdef USE_STL
				objectsToRemove.push_back(objectIndex);
			#else
				objectsToRemove.AddTail(objectIndex);
			#endif
		}
		objectIndex = GetNextObjectToProcess(objectIndex, numberOfObjects);
	}*/
	return RemoveObjects(objectsToRemove);
}