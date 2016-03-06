//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Resource.h"
#include "Debug.h"
#include "..\EventLog.h"
#include "..\Interfaces.h"
#include "..\PrivateInterfaces.h"
#ifdef DESIGNER
#include "Globals.h"
#else
#include "..\Globals.h"
#endif
#include "DragDrop.h"

extern HINSTANCE g_hInstance;

//
// CToolDataObject
//

CToolDataObject::CToolDataObject(Source			     eSource, 
								 IActiveBar2*		 pBar, 
								 TypedArray<ITool*>& aTools, 
								 BSTR				 bstrBand, 
								 BSTR				 bstrChildBand)
	: m_pBar(pBar),
	  m_pBand(NULL),
	  m_aTools(aTools),
	  m_bstrBand(NULL),
	  m_bstrChildBand(NULL),
	  m_bstrCategory(NULL)
{
	if (eLibraryDragDropId == eSource)
		m_eType = eTool;
	else if (bstrBand)
	{
		m_bstrBand = SysAllocString(bstrBand);
		if (bstrChildBand)
		{
			m_bstrChildBand = SysAllocString(bstrChildBand);
			m_eType = eBandChildBandToolId;
		}
		else
			m_eType = eBandToolId;
	}
	else
		m_eType = eToolId;
	m_cRef = 1;
	m_eSource = eSource;
}

CToolDataObject::CToolDataObject(Source eSource, IActiveBar2* pBar, BSTR bstrCategory, TypedArray<ITool*>& aTools)
	: m_pBar(pBar),
	  m_pBand(NULL),
	  m_aTools(aTools),
	  m_bstrBand(NULL),
	  m_bstrChildBand(NULL)
{
	m_eType = eCategory;
	m_cRef = 1;
	m_eSource = eSource;
    m_bstrCategory = SysAllocString(bstrCategory);
}

CToolDataObject::CToolDataObject(Source eSource, IActiveBar2* pBar, IBand* pBand, TypedArray<ITool*>& aTools)
	: m_pBar(pBar),
	  m_pBand(pBand),
	  m_bstrBand(NULL),
	  m_bstrChildBand(NULL),
	  m_bstrCategory(NULL),
	  m_aTools(aTools)
{
	m_eType = eBand;
	m_cRef = 1;
	m_eSource = eSource;
}

CToolDataObject::~CToolDataObject()
{
	SysFreeString(m_bstrChildBand);
	SysFreeString(m_bstrCategory);
	SysFreeString(m_bstrBand);
}

//
// QueryInterface
//

HRESULT __stdcall CToolDataObject::QueryInterface(REFIID riid, void** ppvObject)
{
	if (riid == IID_IDataObject || riid == IID_IUnknown)
	{
		*ppvObject = static_cast<IDataObject*>(this);
		AddRef();
		return S_OK;
	}
	*ppvObject = 0;
	return E_NOINTERFACE;
}
        
//
// AddRef
//

ULONG __stdcall CToolDataObject::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

//
// Release
//

ULONG __stdcall CToolDataObject::Release()
{
	if (0 == InterlockedDecrement(&m_cRef))
	{
		delete this;
		return 0;
	}
	return m_cRef;
}

//
// GetData
//

HRESULT __stdcall CToolDataObject::GetData(FORMATETC* pFormatEtcIn, STGMEDIUM* pMedium)
{
	if (!(DVASPECT_CONTENT & pFormatEtcIn->dwAspect))
		return DATA_E_FORMATETC;

	if (!(TYMED_ISTREAM & pFormatEtcIn->tymed))
		return DATA_E_FORMATETC;

	IStream* pStream;
	HRESULT hResult = CreateStreamOnHGlobal(NULL, TRUE, &pStream);
	if (FAILED(hResult) || NULL == pStream)
		return STG_E_MEDIUMFULL;

	pMedium->pstm = pStream;
	pMedium->tymed = TYMED_ISTREAM;
	pMedium->pUnkForRelease = NULL;

	long nId = m_eSource;
	hResult = pStream->Write(&nId, sizeof(nId), NULL);
	if (FAILED(hResult))
		return hResult;

	hResult = pStream->Write(&m_pBar, sizeof(m_pBar), NULL);
	if (FAILED(hResult))
		return hResult;

	switch (m_eType)
	{
	case eTool:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolFormat)
		{
			int nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(nCount), NULL);
			if (FAILED(hResult))
				return hResult;

			for (int nTool = 0; nTool < nCount; nTool++)
			{
				hResult = m_aTools.GetAt(nTool)->DragDropExchange(pStream, VARIANT_TRUE);
				if (FAILED(hResult))
					return hResult;
			}

			return hResult;
		}
		break;

	case eToolId:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolIdFormat)
		{
			IBarPrivate* pPrivateBar;
			hResult = m_pBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivateBar);
			if (FAILED(hResult))
				return hResult;

			ITool* pTool;
			int nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(int), NULL);
			for (int nTool = 0; nTool < nCount; nTool++)
			{
				pTool = m_aTools.GetAt(nTool);
				hResult = pPrivateBar->ExchangeToolByIdentity((LPDISPATCH)pStream, 
															  VARIANT_TRUE, 
															  (IDispatch**)&pTool);
			}
			pPrivateBar->Release();
			return hResult;
		}
		else if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolFormat)
		{
			int nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(nCount), NULL);
			if (FAILED(hResult))
				return hResult;

			for (int nTool = 0; nTool < nCount; nTool++)
			{
				hResult = m_aTools.GetAt(nTool)->DragDropExchange(pStream, VARIANT_TRUE);
				if (FAILED(hResult))
					return hResult;
			}
			return hResult;
		}
		break;

	case eBandToolId:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipBandToolIdFormat)
		{
			IBarPrivate* pPrivateBar;
			hResult = m_pBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivateBar);
			if (FAILED(hResult))
				return hResult;

			ITool* pTool;
			int nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(int), NULL);
			for (int nTool = 0; nTool < nCount; nTool++)
			{
				pTool = m_aTools.GetAt(nTool);
				hResult = pPrivateBar->ExchangeToolByBandChildBandToolIdentity((LPDISPATCH)pStream, 
																			   m_bstrBand, 
																			   m_bstrChildBand, 
																			   VARIANT_TRUE, 
																			   (IDispatch**)&pTool);
			}
			pPrivateBar->Release();
			return hResult;
		}
		else if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolFormat)
		{
			int nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(nCount), NULL);
			if (FAILED(hResult))
				return hResult;

			for (int nTool = 0; nTool < nCount; nTool++)
			{
				hResult = m_aTools.GetAt(nTool)->DragDropExchange(pStream, VARIANT_TRUE);
				if (FAILED(hResult))
					return hResult;
			}
			return hResult;
		}
		break;

	case eBandChildBandToolId:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipBandChildBandToolIdFormat)
		{
			IBarPrivate* pPrivateBar;
			hResult = m_pBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivateBar);
			if (FAILED(hResult))
				return hResult;

			ITool* pTool;
			int nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(int), NULL);
			for (int nTool = 0; nTool < nCount; nTool++)
			{
				pTool = m_aTools.GetAt(nTool);
				hResult = pPrivateBar->ExchangeToolByBandChildBandToolIdentity((LPDISPATCH)pStream, 
																			   m_bstrBand, 
																			   m_bstrChildBand, 
																			   VARIANT_TRUE, 
																			   (IDispatch**)&pTool);
			}
			pPrivateBar->Release();
			return hResult;
		}
		else if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolFormat)
		{
			int nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(nCount), NULL);
			if (FAILED(hResult))
				return hResult;

			for (int nTool = 0; nTool < nCount; nTool++)
			{
				hResult = m_aTools.GetAt(nTool)->DragDropExchange(pStream, VARIANT_TRUE);
				if (FAILED(hResult))
					return hResult;
			}
			return hResult;
		}
		break;

	case eBand:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipBandFormat)
		{
			hResult = m_pBand->DragDropExchange(pStream, VARIANT_TRUE);
			if (FAILED(hResult))
				return hResult;
			return hResult;
		}
		break;

	case eCategory:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipCategoryFormat)
		{
			hResult = StWriteBSTR(pStream, m_bstrCategory);
			if (FAILED(hResult))
				return hResult;

			int nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(int), NULL);
			if (FAILED(hResult))
				return hResult;

			for (int nTool = 0; nTool < nCount; nTool++)
			{
				hResult = m_aTools.GetAt(nTool)->DragDropExchange(pStream, VARIANT_TRUE);
				if (FAILED(hResult))
				{	
					assert(FALSE);
					continue;
				}
			}
			return NOERROR;
		}
		break;

	default:
		assert(FALSE);
		break;
	}
	return DATA_E_FORMATETC;
}
        
//
// GetDataHere
//

HRESULT __stdcall CToolDataObject::GetDataHere(FORMATETC* pFormatEtc,
											   STGMEDIUM* pMedium)
{
	return E_NOTIMPL;
}
        
//
// SetData
//

HRESULT __stdcall CToolDataObject::SetData(FORMATETC* pFormatEtc,
										   STGMEDIUM* pMedium,
										   BOOL       bRelease)
{
	return DATA_E_FORMATETC;
}
        

//
// QueryGetData
//

HRESULT __stdcall CToolDataObject::QueryGetData(FORMATETC* pFormatEtcIn)
{
	if (!(DVASPECT_CONTENT & pFormatEtcIn->dwAspect))
		return DATA_E_FORMATETC;

	if (!(TYMED_ISTREAM & pFormatEtcIn->tymed))
		return DATA_E_FORMATETC;

	switch (m_eType)
	{
	case eTool:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolFormat)
			return NOERROR;
		break;

	case eToolId:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolIdFormat)
			return NOERROR;
		else if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolFormat)
			return NOERROR;
		break;

	case eBandToolId:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipBandToolIdFormat)
			return NOERROR;
		else if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolFormat)
			return NOERROR;
		break;

	case eBandChildBandToolId:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipBandChildBandToolIdFormat)
			return NOERROR;
		else if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipToolFormat)
			return NOERROR;
		break;

	case eBand:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipBandFormat)
			return NOERROR;
		break;

	case eCategory:
		if (pFormatEtcIn->cfFormat == GetGlobals().m_nIDClipCategoryFormat)
			return NOERROR;
		break;
	}
	return S_FALSE;
}
        
//
// GetCanonicalFormatEtc
//

HRESULT __stdcall CToolDataObject::GetCanonicalFormatEtc(FORMATETC* pFormatEtctIn,
														 FORMATETC* pFormatEtcOut)
{
	if (NULL == pFormatEtcOut)
		return E_INVALIDARG;

	pFormatEtcOut->ptd = NULL;
	return S_FALSE;
}
        
//
// EnumFormatEtc
//

HRESULT __stdcall CToolDataObject::EnumFormatEtc(DWORD		      dwDirection,
												 IEnumFORMATETC** ppEnumFormatEtc)
{
	*ppEnumFormatEtc = new CToolEnumFORMATETC(this);
	(*ppEnumFormatEtc)->AddRef();
	return S_OK;
}
        
//
// DAdvise
//

HRESULT __stdcall CToolDataObject::DAdvise(FORMATETC*   pFormatEtc, 
										   DWORD		advf, 
										   IAdviseSink* pAdvSink,
										   DWORD*       pdwConnection)
{
	return OLE_E_ADVISENOTSUPPORTED;
}
        
//
// DUnadvise
//

HRESULT __stdcall CToolDataObject::DUnadvise(DWORD dwConnection)
{
	return OLE_E_NOCONNECTION;
}
        
//
// EnumDAdvise
//

HRESULT __stdcall CToolDataObject::EnumDAdvise(IEnumSTATDATA** ppEnumAdvise)
{
	return OLE_E_ADVISENOTSUPPORTED;
}

//
// CToolEnumFORMATETC
//

CToolEnumFORMATETC::CToolEnumFORMATETC(CToolDataObject* pToolDataObject)
	: m_pToolDataObject(pToolDataObject)
{
	m_cRef = 1;
	m_nIndex = 0;

	m_FmEtc[0].cfFormat = GetGlobals().m_nIDClipToolFormat; 
	m_FmEtc[0].ptd = 0; 
	m_FmEtc[0].dwAspect = DVASPECT_CONTENT; 
	m_FmEtc[0].lindex = -1; 
	m_FmEtc[0].tymed = TYMED_ISTREAM; 
	
	m_FmEtc[1].cfFormat = GetGlobals().m_nIDClipToolIdFormat; 
	m_FmEtc[1].ptd = 0; 
	m_FmEtc[1].dwAspect = DVASPECT_CONTENT; 
	m_FmEtc[1].lindex = -1; 
	m_FmEtc[1].tymed = TYMED_ISTREAM; 

	m_FmEtc[2].cfFormat = GetGlobals().m_nIDClipBandToolIdFormat; 
	m_FmEtc[2].ptd = 0; 
	m_FmEtc[2].dwAspect = DVASPECT_CONTENT; 
	m_FmEtc[2].lindex = -1; 
	m_FmEtc[2].tymed = TYMED_ISTREAM; 

	m_FmEtc[3].cfFormat = GetGlobals().m_nIDClipBandChildBandToolIdFormat; 
	m_FmEtc[3].ptd = 0; 
	m_FmEtc[3].dwAspect = DVASPECT_CONTENT; 
	m_FmEtc[3].lindex = -1; 
	m_FmEtc[3].tymed = TYMED_ISTREAM; 
	
	m_FmEtc[4].cfFormat = GetGlobals().m_nIDClipBandFormat; 
	m_FmEtc[4].ptd = 0; 
	m_FmEtc[4].dwAspect = DVASPECT_CONTENT; 
	m_FmEtc[4].lindex = -1; 
	m_FmEtc[4].tymed = TYMED_ISTREAM; 
	
	m_FmEtc[5].cfFormat = GetGlobals().m_nIDClipCategoryFormat; 
	m_FmEtc[5].ptd = 0; 
	m_FmEtc[5].dwAspect = DVASPECT_CONTENT; 
	m_FmEtc[5].lindex = -1; 
	m_FmEtc[5].tymed = TYMED_ISTREAM; 
}
	
//
// QueryTool
//

HRESULT __stdcall CToolEnumFORMATETC::QueryInterface(REFIID riid, void** ppvObject)
{
	if (IID_IEnumFORMATETC == riid || IID_IUnknown == riid)
	{
		*ppvObject = static_cast<IEnumFORMATETC*>(this);
		AddRef();
		return S_OK;
	}
	*ppvObject = NULL;
	return E_NOINTERFACE;
}
        
//
// AddRef
//

ULONG __stdcall CToolEnumFORMATETC::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

//
// Release
//

ULONG __stdcall CToolEnumFORMATETC::Release()
{
	if (0 == InterlockedDecrement(&m_cRef))
	{
		delete this;
		return 0;
	}
	return m_cRef;
}

//
// Next
//

HRESULT __stdcall CToolEnumFORMATETC::Next(ULONG	  celt, 
										   FORMATETC* pFormatetc, 
										   ULONG*	  plFetched)
{
	if (m_nIndex >= 6)
		return S_FALSE;

	if (celt >= 6  || NULL == pFormatetc)
		return E_UNEXPECTED;

	m_nIndex++;
	*pFormatetc = m_FmEtc[m_nIndex];

	if (NULL != plFetched) 
		*plFetched = 6;

	return S_OK;
}

//
// Skip
//

HRESULT __stdcall CToolEnumFORMATETC::Skip(ULONG celt)
{
	UINT nFmts = 6;
	nFmts = (nFmts < celt) ? nFmts : celt;
	if (nFmts != celt)
		return S_FALSE;
	return S_OK;
}
	
//
// Reset
//

HRESULT __stdcall CToolEnumFORMATETC::Reset()
{
	m_nIndex = 0;
	return S_OK;
}
	
//
// Clone
//

HRESULT __stdcall CToolEnumFORMATETC::Clone(IEnumFORMATETC** ppEnum)
{
	*ppEnum = new CToolEnumFORMATETC(m_pToolDataObject);
	(*ppEnum)->AddRef();
	return S_OK;
}

