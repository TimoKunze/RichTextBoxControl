// OLEObject.cpp: A wrapper for an embedded OLE object.

#include "stdafx.h"
#include "OLEObject.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP OLEObject::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IRichOLEObject, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


HRESULT CALLBACK OLEObject::QueryITextRangeInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(ITextRange), queriedInterface) || InlineIsEqualGUID(__uuidof(ITextRange2), queriedInterface)) {
		OLEObject* pOLEObject = reinterpret_cast<OLEObject*>(pThis);
		if(pOLEObject->properties.pTextRange) {
			return pOLEObject->properties.pTextRange->QueryInterface(queriedInterface, ppImplementation);
		}
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}

HRESULT CALLBACK OLEObject::QueryIOleObjectInterface(LPVOID pThis, REFIID queriedInterface, LPVOID* ppImplementation, DWORD_PTR /*cookie*/)
{
	ATLASSERT_POINTER(ppImplementation, LPVOID);
	if(!ppImplementation) {
		return E_POINTER;
	}

	if(InlineIsEqualGUID(__uuidof(IOleObject), queriedInterface)) {
		OLEObject* pOLEObject = reinterpret_cast<OLEObject*>(pThis);
		return pOLEObject->properties.pOleObject->QueryInterface(queriedInterface, ppImplementation);
	}

	*ppImplementation = NULL;
	return E_NOINTERFACE;
}


OLEObject::Properties::~Properties()
{
	if(pOwnerRTB) {
		pOwnerRTB->Release();
		pOwnerRTB = NULL;
	}
	if(pTextRange) {
		pTextRange->Release();
		pTextRange = NULL;
	}
	if(pOleObject) {
		pOleObject->Release();
		pOleObject = NULL;
	}
}

HWND OLEObject::Properties::GetRTBHWnd(void)
{
	ATLASSUME(pOwnerRTB);

	OLE_HANDLE handle = NULL;
	pOwnerRTB->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void OLEObject::Attach(ITextRange* pTextRange, IOleObject* pOleObject)
{
	if(properties.pTextRange) {
		properties.pTextRange->Release();
		properties.pTextRange = NULL;
	}
	if(properties.pOleObject) {
		properties.pOleObject->Release();
		properties.pOleObject = NULL;
	}
	if(pTextRange) {
		//pTextRange->QueryInterface(IID_PPV_ARGS(&properties.pTextRange));
		properties.pTextRange = pTextRange;
		properties.pTextRange->AddRef();
	}
	if(pOleObject) {
		//pOleObject->QueryInterface(IID_PPV_ARGS(&properties.pOleObject));
		properties.pOleObject = pOleObject;
		properties.pOleObject->AddRef();
	}
}

void OLEObject::Detach(void)
{
	if(properties.pTextRange) {
		properties.pTextRange->Release();
		properties.pTextRange = NULL;
	}
	if(properties.pOleObject) {
		properties.pOleObject->Release();
		properties.pOleObject = NULL;
	}
}

void OLEObject::SetOwner(RichTextBox* pOwner)
{
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->Release();
	}
	properties.pOwnerRTB = pOwner;
	if(properties.pOwnerRTB) {
		properties.pOwnerRTB->AddRef();
	}
}


STDMETHODIMP OLEObject::get_ClassID(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		CLSID clsid = CLSID_NULL;
		hr = properties.pOleObject->GetUserClassID(&clsid);
		if(SUCCEEDED(hr)) {
			LPOLESTR p = NULL;
			StringFromCLSID(clsid, &p);
			if(p) {
				*pValue = SysAllocString(p);
				CoTaskMemFree(p);
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::get_DisplayAspect(DisplayAspectConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisplayAspectConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		HWND hWndRTB = properties.GetRTBHWnd();
		LONG index = 0;
		if(IsWindow(hWndRTB) && SUCCEEDED(get_Index(&index))) {
			CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);
				hr = pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_NO_INTERFACES);
				*pValue = static_cast<DisplayAspectConstants>(objectDetails.dvaspect);
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::put_DisplayAspect(DisplayAspectConstants newValue)
{
	HRESULT hr = E_FAIL;
	DisplayAspectConstants currentDisplayAspect = static_cast<DisplayAspectConstants>(0);
	if(SUCCEEDED(get_DisplayAspect(&currentDisplayAspect)) && currentDisplayAspect == newValue) {
		// nothing to do
		hr = S_OK;
	} else if(properties.pOleObject && properties.pTextRange) {
		LONG index = 0;
		if(SUCCEEDED(get_Index(&index))) {
			hr = E_FAIL;
			CComPtr<IRichEditOle> pRichEditOle = NULL;
			HWND hWndRTB = properties.GetRTBHWnd();
			if(IsWindow(hWndRTB) && SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				FORMATETC fmt = {0};
				fmt.dwAspect = newValue;
				fmt.lindex = -1;
				CComQIPtr<IOleCache> pCache = properties.pOleObject;
				if(pCache && SUCCEEDED(pCache->Cache(&fmt, newValue == DVASPECT_ICON ? ADVF_NODATA : ADVF_DATAONSTOP, NULL))) {
					if(newValue == DVASPECT_ICON) {
						fmt.cfFormat = CF_METAFILEPICT;
						fmt.dwAspect = newValue;
						fmt.lindex = -1;
						fmt.tymed = TYMED_MFPICT;

						STGMEDIUM stg = {0};
						CLSID clsid = CLSID_NULL;
						properties.pOleObject->GetUserClassID(&clsid);
						stg.hMetaFilePict = OleGetIconOfClass(clsid, NULL, TRUE);
						pCache->SetData(&fmt, &stg, TRUE);
					}
				}
				hr = pRichEditOle->SetDvaspect(index, newValue);
				if(SUCCEEDED(hr)) {
					// update position - unfortunately removing and re-inserting the object seems to be the only working solution
					LONG start = 0;
					LONG end = 0;
					properties.pTextRange->GetStart(&start);
					properties.pTextRange->GetEnd(&end);

					REOBJECT objectDetails = {0};
					objectDetails.cbStruct = sizeof(REOBJECT);
					pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_ALL_INTERFACES);
					objectDetails.dvaspect = newValue;
					objectDetails.sizel.cx = 0;
					objectDetails.sizel.cy = 0;
					properties.pOwnerRTB->EnterSilentObjectDeletionsSection();
					properties.pTextRange->Delete(tomObject, 1, NULL);
					properties.pOwnerRTB->LeaveSilentObjectDeletionsSection();
					properties.pOwnerRTB->EnterSilentObjectInsertionsSection();
					pRichEditOle->InsertObject(&objectDetails);
					properties.pOwnerRTB->LeaveSilentObjectInsertionsSection();

					properties.pTextRange->SetStart(start);
					properties.pTextRange->SetEnd(end);

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
	}
	return hr;
}

STDMETHODIMP OLEObject::get_Flags(OLEObjectFlagsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEObjectFlagsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		HWND hWndRTB = properties.GetRTBHWnd();
		LONG index = 0;
		if(IsWindow(hWndRTB) && SUCCEEDED(get_Index(&index))) {
			CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);
				hr = pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_NO_INTERFACES);
				*pValue = static_cast<OLEObjectFlagsConstants>(objectDetails.dwFlags);
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::put_Flags(OLEObjectFlagsConstants newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pOleObject && properties.pTextRange) {
		LONG index = 0;
		if(SUCCEEDED(get_Index(&index))) {
			HWND hWndRTB = properties.GetRTBHWnd();
			CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(IsWindow(hWndRTB) && SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				DisplayAspectConstants displayAspect = static_cast<DisplayAspectConstants>(0);
				get_DisplayAspect(&displayAspect);
				FORMATETC fmt = {0};
				fmt.dwAspect = displayAspect;
				fmt.lindex = -1;
				CComQIPtr<IOleCache> pCache = properties.pOleObject;
				if(pCache && SUCCEEDED(pCache->Cache(&fmt, displayAspect == DVASPECT_ICON ? ADVF_NODATA : ADVF_DATAONSTOP, NULL))) {
					if(displayAspect == DVASPECT_ICON) {
						fmt.cfFormat = CF_METAFILEPICT;
						fmt.dwAspect = displayAspect;
						fmt.lindex = -1;
						fmt.tymed = TYMED_MFPICT;

						STGMEDIUM stg = {0};
						CLSID clsid = CLSID_NULL;
						properties.pOleObject->GetUserClassID(&clsid);
						stg.hMetaFilePict = OleGetIconOfClass(clsid, NULL, TRUE);
						pCache->SetData(&fmt, &stg, TRUE);
					}
				}
				// update flags - unfortunately removing and re-inserting the object seems to be the only working solution
				LONG start = 0;
				LONG end = 0;
				properties.pTextRange->GetStart(&start);
				properties.pTextRange->GetEnd(&end);

				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);
				pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_ALL_INTERFACES);
				objectDetails.dwFlags = newValue;
				properties.pOwnerRTB->EnterSilentObjectDeletionsSection();
				properties.pTextRange->Delete(tomObject, 1, NULL);
				properties.pOwnerRTB->LeaveSilentObjectDeletionsSection();
				properties.pOwnerRTB->EnterSilentObjectInsertionsSection();
				hr = pRichEditOle->InsertObject(&objectDetails);
				properties.pOwnerRTB->LeaveSilentObjectInsertionsSection();

				properties.pTextRange->SetStart(start);
				properties.pTextRange->SetEnd(end);
			
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

STDMETHODIMP OLEObject::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}
	if(!properties.pTextRange || !properties.pOleObject) {
		return CO_E_RELEASED;
	}

	CComPtr<ITextRange> pTmpRange = NULL;
	properties.pTextRange->GetDuplicate(&pTmpRange);
	pTmpRange->Collapse(tomStart);
	HRESULT hr = pTmpRange->GetIndex(tomObject, pValue);
	// zero-based
	(*pValue)--;
	return hr;
}

STDMETHODIMP OLEObject::get_LinkSource(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		CComQIPtr<IOleLink> pOleLink = properties.pOleObject;
		if(pOleLink) {
			LPOLESTR p = NULL;
			hr = pOleLink->GetSourceDisplayName(&p);
			if(p) {
				*pValue = SysAllocString(p);
				CoTaskMemFree(p);
			}
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::put_LinkSource(BSTR newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		CComQIPtr<IOleLink> pOleLink = properties.pOleObject;
		if(pOleLink) {
			hr = pOleLink->SetSourceDisplayName(newValue);
		} else {
			hr = E_NOINTERFACE;
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::get_ProgID(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		CLSID clsid = CLSID_NULL;
		hr = properties.pOleObject->GetUserClassID(&clsid);
		if(SUCCEEDED(hr)) {
			LPOLESTR p = NULL;
			ProgIDFromCLSID(clsid, &p);
			if(p) {
				*pValue = SysAllocString(p);
				CoTaskMemFree(p);
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::get_TextRange(IRichTextRange** ppTextRange)
{
	ATLASSERT_POINTER(ppTextRange, IRichTextRange*);
	if(!ppTextRange) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pTextRange) {
		CComPtr<ITextRange> pTextRange = NULL;
		hr = properties.pTextRange->GetDuplicate(&pTextRange);
		if(SUCCEEDED(hr)) {
			ClassFactory::InitTextRange(pTextRange, properties.pOwnerRTB, IID_IRichTextRange, reinterpret_cast<LPUNKNOWN*>(ppTextRange));
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::get_TypeName(OLEObjectTypeNameFormatConstants format, BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		LPOLESTR p = NULL;
		hr = properties.pOleObject->GetUserType(format, &p);
		if(p) {
			*pValue = SysAllocString(p);
			CoTaskMemFree(p);
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::get_UserData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		HWND hWndRTB = properties.GetRTBHWnd();
		LONG index = 0;
		if(IsWindow(hWndRTB) && SUCCEEDED(get_Index(&index))) {
			CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);
				hr = pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_NO_INTERFACES);
				*pValue = objectDetails.dwUser;
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::put_UserData(LONG newValue)
{
	HRESULT hr = E_FAIL;
	if(properties.pOleObject && properties.pTextRange) {
		LONG index = 0;
		if(SUCCEEDED(get_Index(&index))) {
			HWND hWndRTB = properties.GetRTBHWnd();
			CComPtr<IRichEditOle> pRichEditOle = NULL;
			if(IsWindow(hWndRTB) && SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
				DisplayAspectConstants displayAspect = static_cast<DisplayAspectConstants>(0);
				get_DisplayAspect(&displayAspect);
				FORMATETC fmt = {0};
				fmt.dwAspect = displayAspect;
				fmt.lindex = -1;
				CComQIPtr<IOleCache> pCache = properties.pOleObject;
				if(pCache && SUCCEEDED(pCache->Cache(&fmt, displayAspect == DVASPECT_ICON ? ADVF_NODATA : ADVF_DATAONSTOP, NULL))) {
					if(displayAspect == DVASPECT_ICON) {
						fmt.cfFormat = CF_METAFILEPICT;
						fmt.dwAspect = displayAspect;
						fmt.lindex = -1;
						fmt.tymed = TYMED_MFPICT;

						STGMEDIUM stg = {0};
						CLSID clsid = CLSID_NULL;
						properties.pOleObject->GetUserClassID(&clsid);
						stg.hMetaFilePict = OleGetIconOfClass(clsid, NULL, TRUE);
						pCache->SetData(&fmt, &stg, TRUE);
					}
				}
				// update data - unfortunately removing and re-inserting the object seems to be the only working solution
				LONG start = 0;
				LONG end = 0;
				properties.pTextRange->GetStart(&start);
				properties.pTextRange->GetEnd(&end);

				REOBJECT objectDetails = {0};
				objectDetails.cbStruct = sizeof(REOBJECT);
				pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_ALL_INTERFACES);
				objectDetails.dwUser = newValue;
				properties.pOwnerRTB->EnterSilentObjectDeletionsSection();
				properties.pTextRange->Delete(tomObject, 1, NULL);
				properties.pOwnerRTB->LeaveSilentObjectDeletionsSection();
				properties.pOwnerRTB->EnterSilentObjectInsertionsSection();
				hr = pRichEditOle->InsertObject(&objectDetails);
				properties.pOwnerRTB->LeaveSilentObjectInsertionsSection();

				properties.pTextRange->SetStart(start);
				properties.pTextRange->SetEnd(end);
			
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


STDMETHODIMP OLEObject::ExecuteVerb(LONG verbID)
{
	LONG index = 0;
	HRESULT hr = E_FAIL;
	if(SUCCEEDED(get_Index(&index))) {
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		HWND hWndRTB = properties.GetRTBHWnd();
		if(IsWindow(hWndRTB) && SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle && properties.pOleObject) {
			REOBJECT objectDetails = {0};
			objectDetails.cbStruct = sizeof(REOBJECT);
			hr = pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_POLESITE);
			if(SUCCEEDED(hr)) {
				if(properties.pTextRange) {
					RECT positionRectangle = {0};
					properties.pTextRange->GetPoint(tomStart | TA_LEFT | TA_TOP | tomClientCoord | tomAllowOffClient, &positionRectangle.left, &positionRectangle.top);
					//properties.pTextRange->GetPoint(tomEnd | TA_RIGHT | TA_BOTTOM | tomClientCoord | tomAllowOffClient, &positionRectangle.right, &positionRectangle.bottom);
					SIZEL pixelSize = {0};
					AtlHiMetricToPixel(&objectDetails.sizel, &pixelSize);
					positionRectangle.right = positionRectangle.left + pixelSize.cx;
					positionRectangle.bottom = positionRectangle.top + pixelSize.cy;

					hr = properties.pOleObject->DoVerb(verbID, NULL, objectDetails.polesite, 0, hWndRTB, &positionRectangle);
				} else {
					hr = properties.pOleObject->DoVerb(verbID, NULL, objectDetails.polesite, 0, NULL, NULL);
				}
				if(objectDetails.polesite) {
					objectDetails.polesite->Release();
					objectDetails.polesite = NULL;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::GetRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	LONG index = 0;
	HRESULT hr = E_FAIL;
	if(SUCCEEDED(get_Index(&index)) && properties.pTextRange) {
		HWND hWndRTB = properties.GetRTBHWnd();
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(IsWindow(hWndRTB) && SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			REOBJECT objectDetails = {0};
			objectDetails.cbStruct = sizeof(REOBJECT);
			hr = pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_NO_INTERFACES);
			if(SUCCEEDED(hr)) {
				RECT boundingRectangle = {0};
				properties.pTextRange->GetPoint(tomStart | TA_LEFT | TA_TOP | tomClientCoord | tomAllowOffClient, &boundingRectangle.left, &boundingRectangle.top);
				//properties.pTextRange->GetPoint(tomEnd | TA_RIGHT | TA_BOTTOM | tomClientCoord | tomAllowOffClient, &boundingRectangle.right, &boundingRectangle.bottom);
				SIZEL pixelSize = {0};
				AtlHiMetricToPixel(&objectDetails.sizel, &pixelSize);
				boundingRectangle.right = boundingRectangle.left + pixelSize.cx;
				boundingRectangle.bottom = boundingRectangle.top + pixelSize.cy;

				if(pXLeft) {
					*pXLeft = boundingRectangle.left;
				}
				if(pYTop) {
					*pYTop = boundingRectangle.top;
				}
				if(pXRight) {
					*pXRight = boundingRectangle.right;
				}
				if(pYBottom) {
					*pYBottom = boundingRectangle.bottom;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::GetSize(OLE_XSIZE_HIMETRIC* pWidth/* = NULL*/, OLE_YSIZE_HIMETRIC* pHeight/* = NULL*/)
{
	LONG index = 0;
	HRESULT hr = E_FAIL;
	if(SUCCEEDED(get_Index(&index))) {
		HWND hWndRTB = properties.GetRTBHWnd();
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(IsWindow(hWndRTB) && SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			REOBJECT objectDetails = {0};
			objectDetails.cbStruct = sizeof(REOBJECT);
			hr = pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_NO_INTERFACES);
			if(SUCCEEDED(hr)) {
				if(pWidth) {
					*pWidth = objectDetails.sizel.cx;
				}
				if(pHeight) {
					*pHeight = objectDetails.sizel.cy;
				}
			}
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::GetVerbs(VARIANT* pVerbs, LONG* pCount)
{
	ATLASSERT_POINTER(pVerbs, VARIANT);
	ATLASSERT_POINTER(pCount, LONG);
	if(!pVerbs) {
		return E_INVALIDARG;
	}
	if(!pCount) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOleObject) {
		VariantClear(pVerbs);
		CComPtr<IEnumOLEVERB> pEnumVerbs = NULL;
		hr = properties.pOleObject->EnumVerbs(&pEnumVerbs);
		if(SUCCEEDED(hr) && pEnumVerbs) {
			pVerbs->vt = VT_ARRAY | VT_RECORD;
			CComPtr<IRecordInfo> pRecordInfo = NULL;
			CLSID clsidOLEVERBDETAILS = {0};
			#ifdef UNICODE
				CLSIDFromString(OLESTR("{DE321E31-614B-4F6B-BC65-BFB2B5E1BE72}"), &clsidOLEVERBDETAILS);
				ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_RichTextBoxCtlLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidOLEVERBDETAILS), &pRecordInfo)));
			#else
				CLSIDFromString(OLESTR("{ACF9C8A9-76EB-4484-84B6-704C4723BC25}"), &clsidOLEVERBDETAILS);
				ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_RichTextBoxCtlLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidOLEVERBDETAILS), &pRecordInfo)));
			#endif
			pVerbs->parray = SafeArrayCreateVectorEx(VT_RECORD, 0, 1, pRecordInfo);
			ATLASSERT_POINTER(pVerbs->parray, SAFEARRAY);

			*pCount = 0;
			ULONG fetched = 0;
			OLEVERB verb = {0};
			OLEVERBDETAILS element = {0};
			for(*pCount = 0; pEnumVerbs->Next(1, &verb, &fetched) == S_OK && fetched == 1; (*pCount)++) {
				element.ID = verb.lVerb;
				element.Name = SysAllocString(verb.lpszVerbName);
				CoTaskMemFree(verb.lpszVerbName);
				element.MenuFlags = verb.fuFlags;
				element.Attributes = verb.grfAttribs;

				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(pVerbs->parray, 1, &lBound);
				SafeArrayGetUBound(pVerbs->parray, 1, &uBound);
				if((*pCount) >= uBound - lBound + 1) {
					SAFEARRAYBOUND bounds = {*pCount + 1, lBound};
					SafeArrayRedim(pVerbs->parray, &bounds);
				}
				SafeArrayPutElement(pVerbs->parray, pCount, &element);
			}
		} else {
			// e.g. REGDB_E_CLASSNOTREG
			hr = OLEOBJ_E_NOVERBS;
		}
		if(hr == OLEOBJ_E_NOVERBS) {
			*pCount = 0;
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP OLEObject::SetSize(OLE_XSIZE_HIMETRIC width, OLE_YSIZE_HIMETRIC height)
{
	// NOTE: properties.pOleObject->SetExtent() fails
	/*if(properties.pOleObject) {
		DisplayAspectConstants displayAspect = static_cast<DisplayAspectConstants>(0);
		if(SUCCEEDED(get_DisplayAspect(&displayAspect))) {
			SIZEL size = {width, height};
			hr = properties.pOleObject->SetExtent(displayAspect, &size);
		}
	}*/

	LONG index = 0;
	HRESULT hr = E_FAIL;
	if(SUCCEEDED(get_Index(&index))) {
		hr = E_FAIL;
		HWND hWndRTB = properties.GetRTBHWnd();
		CComPtr<IRichEditOle> pRichEditOle = NULL;
		if(IsWindow(hWndRTB) && SendMessage(hWndRTB, EM_GETOLEINTERFACE, 0, reinterpret_cast<LPARAM>(&pRichEditOle)) && pRichEditOle) {
			DisplayAspectConstants displayAspect = static_cast<DisplayAspectConstants>(0);
			get_DisplayAspect(&displayAspect);
			FORMATETC fmt = {0};
			fmt.dwAspect = displayAspect;
			fmt.lindex = -1;
			CComQIPtr<IOleCache> pCache = properties.pOleObject;
			if(pCache && SUCCEEDED(pCache->Cache(&fmt, displayAspect == DVASPECT_ICON ? ADVF_NODATA : ADVF_DATAONSTOP, NULL))) {
				if(displayAspect == DVASPECT_ICON) {
					fmt.cfFormat = CF_METAFILEPICT;
					fmt.dwAspect = displayAspect;
					fmt.lindex = -1;
					fmt.tymed = TYMED_MFPICT;

					STGMEDIUM stg = {0};
					CLSID clsid = CLSID_NULL;
					properties.pOleObject->GetUserClassID(&clsid);
					stg.hMetaFilePict = OleGetIconOfClass(clsid, NULL, TRUE);
					pCache->SetData(&fmt, &stg, TRUE);
				}
			}
			// update extent - unfortunately removing and re-inserting the object seems to be the only working solution
			LONG start = 0;
			LONG end = 0;
			properties.pTextRange->GetStart(&start);
			properties.pTextRange->GetEnd(&end);

			REOBJECT objectDetails = {0};
			objectDetails.cbStruct = sizeof(REOBJECT);
			pRichEditOle->GetObject(index, &objectDetails, REO_GETOBJ_ALL_INTERFACES);
			objectDetails.sizel.cx = width;
			objectDetails.sizel.cy = height;
			properties.pOwnerRTB->EnterSilentObjectDeletionsSection();
			properties.pTextRange->Delete(tomObject, 1, NULL);
			properties.pOwnerRTB->LeaveSilentObjectDeletionsSection();
			properties.pOwnerRTB->EnterSilentObjectInsertionsSection();
			hr = pRichEditOle->InsertObject(&objectDetails);
			properties.pOwnerRTB->LeaveSilentObjectInsertionsSection();

			properties.pTextRange->SetStart(start);
			properties.pTextRange->SetEnd(end);
			
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