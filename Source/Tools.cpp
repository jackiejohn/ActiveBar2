//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <olectl.h>
#include "IPServer.h"
#include <stddef.h>       // for offsetof()
#include "Support.h"
#include "Resource.h"
#include "Debug.h"
#include "ImageMgr.h"
#include "Tool.h"
#include "Tools.h"
#include "Band.h"
#include "Bands.h"
#include "Localizer.h"
#include "ChildBands.h"
#include "Errors.h"
#include "CustomProxy.h" 

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

struct MenuTool
{
	UINT		 nToolId;
	UINT		 nNameId;
	UINT		 nCaptionId;
	UINT		 nSubBandId;
	UINT		 nBitmapId;
	UINT         ttTool;
	VARIANT_BOOL vbVisible;
};
MenuTool MainMenu[] =
{
	{CBar::eToolIdSysCommand, IDS_SYSCOMMAND,  0,          IDS_POPUPSYSCOMMAND, 0, (ToolTypes)CTool::ddTTChildSysMenu, VARIANT_FALSE},
	{CBar::eToolIdWindow,     IDS_MIWINDOW,    IDS_WINDOW, IDS_POPUPWINDOW,     0, ddTTButton,						   VARIANT_TRUE},
	{CBar::eToolIdMDIButtons, IDS_MIMDIBUTTON, 0,          0,                   0, (ToolTypes)CTool::ddTTMDIButtons,   VARIANT_FALSE}
};
MenuTool PopupWindow[] =
{
	{(UINT)CBar::eToolIdCascade,   IDS_MICASCADE,  IDS_CASCADE,  0, IDB_CASCADE,  ddTTButton, VARIANT_TRUE},
	{(UINT)CBar::eToolIdTileHorz,  IDS_MITILEHORZ, IDS_TILEHORZ, 0, IDB_TITEHORZ, ddTTButton, VARIANT_TRUE},
	{(UINT)CBar::eToolIdTileVert,  IDS_MITILEVERT, IDS_TILEVERT, 0, IDB_TITEVERT, ddTTButton, VARIANT_TRUE}
};
MenuTool MoreTools[] =
{
	{CBar::eToolMoreTools, IDS_MORETOOLS, IDS_MORETOOLS, IDS_POPUPMORETOOLS, 0, (ToolTypes)CTool::ddTTMoreTools, VARIANT_TRUE},
};

static void SetTool(CTool* pTool, MenuTool& theToolInfo)
{
	DDString strTemp;
	HBITMAP hBitmap;
	HRESULT hResult;
	BSTR bstrTemp;

	pTool->tpV1.m_nToolId =   theToolInfo.nToolId;
	pTool->tpV1.m_ttTools =   (ToolTypes)theToolInfo.ttTool;
	pTool->tpV1.m_vbVisible = theToolInfo.vbVisible;
	if (0 != theToolInfo.nNameId)
	{
		strTemp.LoadString(theToolInfo.nNameId);
		bstrTemp = strTemp.AllocSysString();
		if (bstrTemp)
		{
			hResult = pTool->put_Name(bstrTemp);
			SysFreeString(bstrTemp);
		}
	}
	if (0 != theToolInfo.nCaptionId)
	{
		strTemp.LoadString(theToolInfo.nCaptionId);
		bstrTemp = strTemp.AllocSysString();
		if (bstrTemp)
		{
			hResult = pTool->put_Caption(bstrTemp);
			hResult = pTool->put_Description(bstrTemp);
			SysFreeString(bstrTemp);
		}
	}
	if (0 != theToolInfo.nSubBandId)
	{
		strTemp.LoadString(theToolInfo.nSubBandId);
		bstrTemp = strTemp.AllocSysString();
		if (bstrTemp)
		{
			hResult = pTool->put_SubBand(bstrTemp);
			SysFreeString(bstrTemp);
		}
	}
	if (0 != theToolInfo.nBitmapId)
	{
		hBitmap = LoadBitmap(g_hInstance, MAKEINTRESOURCE(theToolInfo.nBitmapId));
		assert(hBitmap);
		if (hBitmap)
		{
			hResult = pTool->put_Bitmap(ddITNormal, (OLE_HANDLE)hBitmap);
			BOOL bResult = DeleteBitmap(hBitmap);
			assert(bResult);
		}
	}
}

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CTools(IUnknown *pUnkOuter)
{
	return (IUnknown *)(ITools *)new CTools();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_Tools,Tools,_T("Tools Class"),CreateFN_CTools,2,0,&IID_ITools,_T(""));
void *CTools::objectDef=&ToolsObject;
CTools *CTools::CreateInstance(IUnknown *pUnkOuter)
{
	return new CTools();
}
//{OBJECT CREATEFN}
CTools::CTools()
	: m_refCount(1)
	,m_pBar(NULL),
	 m_pBand(NULL),
	 m_pMoreTools(NULL)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
	m_pMDITools[0] = NULL;
	m_pMDITools[1] = NULL;
}

CTools::~CTools()
{
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
// {END CLEANUP}
	CleanupVisible();
	RemoveAll();
	if (m_pMDITools[0])
		m_pMDITools[0]->Release();
	if (m_pMDITools[1])
		m_pMDITools[1]->Release();
	if (m_pMoreTools)
		m_pMoreTools->Release();
}

#ifdef _DEBUG
void CTools::Dump(DumpContext& dc)
{
	int nElem = m_aTools.GetSize();
	for (int cnt = 0; cnt < nElem; ++cnt)
		m_aTools.GetAt(cnt)->Dump(dc);
}
#endif

void CTools::SetBar(CBar* pBar)
{
	m_pBar = pBar;
}

void CTools::SetBand(CBand* pBand)
{
	m_pBand = pBand;
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CTools::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IUnknown))
	{
		AddRef();
		*ppvObjOut=this;
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IDispatch))
	{
		*ppvObjOut=(void *)(IDispatch *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ITools))
	{
		*ppvObjOut=(void *)(ITools *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISupportErrorInfo))
	{
		*ppvObjOut=(void *)(ISupportErrorInfo *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ICategorizeProperties))
	{
		*ppvObjOut=(void *)(ICategorizeProperties *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
STDMETHODIMP_(ULONG) CTools::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CTools::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
STDMETHODIMP CTools::MapPropertyToCategory( DISPID dispid,  PROPCAT*ppropcat)
{
	return E_FAIL;
}
STDMETHODIMP CTools::GetCategoryName( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName)
{
	return E_FAIL;
}
STDMETHODIMP CTools::_NewEnum( IUnknown **retval)
{
	*retval = 0;
	VARIANT* pColl;
	IEnumX* pEnum;
	int cItems=m_aTools.GetSize();
	if (cItems!=0)
	{
		if (!(pColl = (VARIANT*)HeapAlloc(g_hHeap, 0, cItems * sizeof(VARIANT))))
			return E_OUTOFMEMORY;
	}
	else
		pColl=0;

	for (int cnt = 0; cnt < cItems; cnt++)
	{
		(pColl+cnt)->vt=VT_DISPATCH;
		(pColl+cnt)->pdispVal=(ITool*)(CTool*)m_aTools.GetAt(cnt);
	}

	pEnum = (IEnumX*)new CEnumX(IID_IEnumVARIANT,
								cItems, 
								sizeof(VARIANT), 
								pColl, 
								CopyDISPVARIANT);
	if (!pEnum)
        return E_OUTOFMEMORY;
	*retval=pEnum;
	return NOERROR;
}

STDMETHODIMP CTools::Insert( int index, Tool *tool)
{
	if (NULL == tool)
		return E_INVALIDARG;

	if (index<-1 || index>m_aTools.GetSize())
	{
		m_pBar->m_theErrorObject.SendError(IDERR_BADTOOLSINDEX,NULL);
		return CUSTOM_CTL_SCODE(IDERR_BADTOOLSINDEX);
	}
	return InsertTool(index,(ITool*)tool,VARIANT_TRUE);
}

STDMETHODIMP CTools::InterfaceSupportsErrorInfo( REFIID riid)
{
       return S_OK;
}
STDMETHODIMP CTools::DeleteTool(ITool *pTool)
{
	int nElem=m_aTools.GetSize();
	for (int cnt=0;cnt<nElem;++cnt)
	{
		if (((ITool*)m_aTools.GetAt(cnt))==pTool)
		{
			m_aTools.RemoveAt(cnt);
			DeleteVisibleTool(pTool);
			pTool->Release();
			return S_OK;
		}
	}
	return S_FALSE;
}
// IDispatch members

STDMETHODIMP CTools::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CTools::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if (0 == (*pptinfo))
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CTools::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	if (!(riid==IID_NULL))
        return E_INVALIDARG;

	ITypeInfo* pTypeInfo = GetObjectTypeInfo(lcid,objectDef);
	if (0 == pTypeInfo)
		return E_FAIL;

	HRESULT hr = pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgdispid);
	pTypeInfo->Release();
	return hr;
}
STDMETHODIMP CTools::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
{
#ifdef _DEBUG
	if (dispidMember<0)
		TRACE1(1, "IDispatch::Invoke -%X\n",-dispidMember)
	else
		TRACE1(1, "IDispatch::Invoke %X\n",dispidMember)
#endif
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;
	#ifdef ERRORINFO_SUPPORT
	SetErrorInfo(0L,NULL); // should be called if ISupportErrorInfo is used
	#endif
	hr = pTypeInfo->Invoke((IDispatch *)this, dispidMember, wFlags,
		pdispparams, pvarResult,
        pexcepinfo, puArgErr);
    pTypeInfo->Release();
	return hr;
}

// ITools members
int CTools::GetPosOfItem(VARIANT *Index)
{
	CTool *pTool;
	int pos=-1;
	switch(Index->vt)
	{
	case VT_I2:
		pos=Index->iVal;
		break;
	case VT_I2|VT_BYREF:
		pos=*(Index->piVal);
		break;
	case VT_I4:
		pos=Index->lVal;
		break;
	case VT_I4|VT_BYREF:
		pos=*(Index->plVal);
		break;
	default:
		{
			int cnt = 0;
			BSTR searchName;
			VARIANT vstr;
			VariantInit(&vstr);

			if (Index->vt==VT_BSTR)
				searchName=Index->bstrVal;
			else
			{
			
				if (FAILED(VariantChangeType(&vstr,Index,VARIANT_NOVALUEPROP,VT_BSTR)))
					return -1;
				searchName=vstr.bstrVal;
			}

			// accessing by name
			// if name is numeric check for toolid else tool name
			
			if (0 == searchName || 0 == *searchName)
			{
				VariantClear(&vstr);
				return -1;
			}

			int nElem = m_aTools.GetSize();
			for (cnt = 0; cnt < nElem; ++cnt)
			{
				pTool = m_aTools.GetAt(cnt);
				if (pTool->m_bstrName && 0 == wcscmp(searchName,pTool->m_bstrName))
				{
					VariantClear(&vstr);
					return cnt;
				}
			}
			VariantClear(&vstr);
		
			// Try to numeric
			VARIANT v;
			VariantInit(&v);
			
			HRESULT hRes = VariantChangeType(&v,Index,VARIANT_NOVALUEPROP,VT_I4);
			if (FAILED(hRes))
				return hRes;

			pos = v.lVal;
			nElem = m_aTools.GetSize();
			for (cnt = 0; cnt < nElem; cnt++)
			{
				pTool = m_aTools.GetAt(cnt);
				if (pTool->tpV1.m_nToolId == (ULONG)v.lVal)
					return cnt;
			}
		}
	}
	return pos;
}

//
// Item
//

STDMETHODIMP CTools::Item(VARIANT* Index, Tool** retval)
{
	int pos = GetPosOfItem(Index);
	if (pos < 0 || pos > (m_aTools.GetSize()-1))
	{
		m_pBar->m_theErrorObject.SendError(IDERR_BADTOOLSINDEX, 0);
		return CUSTOM_CTL_SCODE(IDERR_BADTOOLSINDEX);
	}

	*retval = (Tool*)(CTool*)m_aTools.GetAt(pos);
	((IUnknown*)(*retval))->AddRef();
	return NOERROR;
}

//
// Count
//

STDMETHODIMP CTools::Count(short* retval)
{
	*retval = m_aTools.GetSize();
	return NOERROR;
}

//
// Add
//

STDMETHODIMP CTools::Add(long nToolId, BSTR bstrName, Tool **retval)
{
	//
	// Create a new tool and insert into the collection
	//
	
	CTool* pTool = CTool::CreateInstance(NULL);
	if (NULL == pTool)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_BADTOOLSINDEX,NULL);
		return CUSTOM_CTL_SCODE(IDERR_BADTOOLSINDEX);
	}

	HRESULT hResult = pTool->put_Name(bstrName);
	if (FAILED(hResult))
	{
		pTool->Release();
		m_pBar->m_theErrorObject.SendError(IDERR_BADTOOLSINDEX,NULL);
		return CUSTOM_CTL_SCODE(IDERR_BADTOOLSINDEX);
	}

	if (m_pBar)
	{
		pTool->tpV1.m_dwIdentity = m_pBar->bpV1.m_dwToolIdentity++;

		if (m_pBar->m_pDesigner)
			static_cast<CTool*>(pTool)->tpV1.m_vbDesignerCreated = TRUE;
	}

	pTool->SetBar(m_pBar);
	pTool->SetBand(m_pBand);
	pTool->m_pTools = this;

	hResult = m_aTools.Add(pTool);
	if (FAILED(hResult))
	{
		pTool->Release();
		return hResult;
	}
	*retval = (Tool*)pTool;
	pTool->AddRef();
	pTool->tpV1.m_nToolId = nToolId;
	if (m_pBand && ddBTMenuBar == m_pBand->bpV1.m_btBands)
	{
		if (CBar::eMDIForm == m_pBar->m_eAppType || CBar::eSDIForm == m_pBar->m_eAppType)
			m_pBar->BuildAccelators(m_pBar->GetMenuBand());
		else if (m_pBar->m_pControlSite)
			m_pBar->m_pControlSite->OnControlInfoChanged();
	}
	return NOERROR;
}

//
// InsertTool
//
// Modified to Insert a Clone
//

STDMETHODIMP CTools::InsertTool(int index, ITool* pSourceTool, VARIANT_BOOL vbClone) 
{
	HRESULT hResult;
	CTool* pTool;
	if (VARIANT_TRUE == vbClone)
	{
		hResult = pSourceTool->Clone((ITool**)&pTool);
		if (FAILED(hResult))
			return hResult;

		pTool->m_bClone = FALSE;

		if (pTool->m_pBar && pTool->m_pBar != m_pBar)
		{
			OLE_HANDLE hBitmap;
			pTool->ReleaseImages();
			pTool->SetBar(m_pBar);
			pTool->SetBand(m_pBand);
			pTool->m_pTools = this;
			for (int nImage = 0; nImage < CTool::eNumOfImageTypes; nImage++)
			{
				hResult = ((CTool*)pSourceTool)->get_Bitmap((ImageTypes)nImage, &hBitmap);
				if (hBitmap)
				{
					hResult = pTool->put_Bitmap((ImageTypes)nImage, hBitmap);
					BOOL bResult = DeleteBitmap((HBITMAP)hBitmap);
					assert(bResult);
				}
				hResult = ((CTool*)pSourceTool)->get_MaskBitmap((ImageTypes)nImage, &hBitmap);
				if (hBitmap)
				{
					hResult = pTool->put_Bitmap((ImageTypes)nImage, hBitmap);
					BOOL bResult = DeleteBitmap((HBITMAP)hBitmap);
					assert(bResult);
				}
			}
		}
	}
	else
	{
		pTool = (CTool*)pSourceTool;
		pSourceTool->AddRef();
	}

	if (m_pBar && m_pBar->m_pDesigner)
		static_cast<CTool*>(pTool)->tpV1.m_vbDesignerCreated = TRUE;

	pTool->SetBar(m_pBar);
	pTool->SetBand(m_pBand);
	pTool->m_pTools = this;
	
	if (m_pBar && this != m_pBar->m_pTools && !m_pBar->m_bToolCreateLock)
	{
		//
		// This tool is being inserted into a different tools collection than the main tools collection
		// We need to check if the tool is in the main tools collection and if it isn't insert it into it.
		//

		if (NULL == m_pBar->m_pTools->FindToolId(pTool->tpV1.m_nToolId))
		{
			//
			// Insert it into the main tools collection.
			//

			hResult = m_pBar->m_pTools->InsertTool(-1, (ITool*)pTool, VARIANT_TRUE);
		}
	}

	if (index == -1)
		hResult = m_aTools.Add(pTool);
	else
		hResult = m_aTools.InsertAt(index, pTool);
	
	if (FAILED(hResult))
	{
		pTool->Release();
		return hResult;
	}
	
	if (m_pBand && ddBTMenuBar == m_pBand->bpV1.m_btBands)
	{
		if (CBar::eMDIForm == m_pBar->m_eAppType || CBar::eSDIForm == m_pBar->m_eAppType)
			m_pBar->BuildAccelators(m_pBar->GetMenuBand());
		else if (m_pBar->m_pControlSite)
			m_pBar->m_pControlSite->OnControlInfoChanged();
	}
	return NOERROR;
}

//
// Remove
//

STDMETHODIMP CTools::Remove( VARIANT *Index)
{
	int pos = GetPosOfItem(Index);
	if (pos < 0 || pos > (m_aTools.GetSize()-1))
	{
		m_pBar->m_theErrorObject.SendError(IDERR_BADTOOLSINDEX,NULL);
		return CUSTOM_CTL_SCODE(IDERR_BADTOOLSINDEX);
	}
	CTool* pTool = m_aTools.GetAt(pos);
	DeleteVisibleTool(pTool);
	pTool->Release();
	m_aTools.RemoveAt(pos);
	if (m_pBand && ddBTMenuBar == m_pBand->bpV1.m_btBands)
	{
		if (CBar::eMDIForm == m_pBar->m_eAppType || CBar::eSDIForm == m_pBar->m_eAppType)
			m_pBar->BuildAccelators(m_pBar->GetMenuBand());
		else if (m_pBar->m_pControlSite)
			m_pBar->m_pControlSite->OnControlInfoChanged();
	}
	return NOERROR;
}

extern IUnknown *CreateFN_CTool(IUnknown *pUnkOuter);

//
// Exchange
//

HRESULT CTools::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
		BOOL bMDIMenu;
		HRESULT hResult;
		CTool*  pTool;
		int		nCount;

		if (VARIANT_TRUE == vbSave)
		{
			try
			{
				//
				// Saving just images
				//

				if (m_pBar && m_pBar->m_bSaveImages)
				{
					nCount = m_aTools.GetSize();
					for (int nTool = 0; nTool < nCount; nTool++)
					{
						hResult = m_aTools.GetAt(nTool)->Exchange(pStream, vbSave);
						if (FAILED(hResult))
							return hResult;
					}
				}
				else
				{
					//
					// Saving everything
					//

					nCount = m_aTools.GetSize();

					hResult = pStream->Write(&nCount, sizeof(DWORD), NULL);
					if (FAILED(hResult))
						return hResult;
					
					for (int nTool = 0; nTool < nCount; nTool++)
					{
						hResult = m_aTools.GetAt(nTool)->Exchange(pStream, vbSave);
						if (FAILED(hResult))
							return hResult;
					}

					hResult = pStream->Write(&bMDIMenu, sizeof(bMDIMenu), NULL);
					if (FAILED(hResult))
						return hResult;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return E_FAIL;
			}
		}
		else
		{
			//
			// Loading
			//
			
			try
			{
				hResult = pStream->Read(&nCount, sizeof(DWORD), 0);
				if (FAILED(hResult))
					return hResult;

				for (int nTool = 0; nTool < nCount; nTool++)
				{
					pTool = CTool::CreateInstance(NULL);
					if (NULL == pTool)
						return E_OUTOFMEMORY;
					
					pTool->SetBar(m_pBar);
					pTool->SetBand(m_pBand);
					pTool->m_pTools = this;
		
					hResult = pTool->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
					
					if (m_pBar && 0 == pTool->tpV1.m_dwIdentity)
						pTool->tpV1.m_dwIdentity = m_pBar->bpV1.m_dwToolIdentity++;

					hResult = m_aTools.Add(pTool);
					if (FAILED(hResult))
					{
						pTool->Release();
						continue;
					}
				}

				hResult = pStream->Read(&bMDIMenu, sizeof(bMDIMenu), NULL);
				if (FAILED(hResult))
					return hResult;
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return E_FAIL;
			}
		}
		return NOERROR;
}

//
// ExchangeConfig
//

HRESULT CTools::ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		BOOL bMDIMenu;
		CTool* pTool;
		int nCount;
		int nTool;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving 
			//

			hResult = pStream->Write(&bMDIMenu, sizeof(bMDIMenu), NULL);
			if (FAILED(hResult))
				return hResult;

			nCount = m_aTools.GetSize();
			hResult = pStream->Write(&nCount, sizeof(nCount), NULL);
			if (FAILED(hResult))
				return hResult;
			
			for (nTool = 0; nTool < nCount; nTool++)
			{
				hResult = m_aTools.GetAt(nTool)->ExchangeConfig(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
			//
			// Loading
			//

			hResult = pStream->Read(&bMDIMenu, sizeof(bMDIMenu), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Read(&nCount, sizeof(nCount), 0);
			if (FAILED(hResult))
				return hResult;

			TypedArray<CTool*> aTools;
			aTools.Copy(m_aTools);
			m_aTools.RemoveAll();
			m_aVisibleTools.RemoveAll();
			for (nTool = 0; nTool < nCount; nTool++)
			{
				//
				// Adding the new Tools
				//

				pTool = CTool::CreateInstance(NULL);
				if (NULL == pTool)
					return E_OUTOFMEMORY;
				
				pTool->SetBar(m_pBar);
				pTool->SetBand(m_pBand);
				pTool->m_pTools = this;

				hResult = pTool->ExchangeConfig(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;

				hResult = m_aTools.Add(pTool);
				if (FAILED(hResult))
					return hResult;
			}

			CTool* pToolOld;
			int nToolOld;
			int nCountOld = aTools.GetSize();
			for (nToolOld = 0; nToolOld < nCountOld; nToolOld++)
			{
				pToolOld = aTools.GetAt(nToolOld);
				if (pToolOld->m_pDispCustom)
				{
					nCount = m_aTools.GetSize();
					for (nTool = 0; nTool < nCount; nTool++)
					{
						//
						// Adding the new Tools
						//

						pTool = m_aTools.GetAt(nTool);
						if (0 == wcscmp(pTool->m_bstrName, pToolOld->m_bstrName))
						{
							LPDISPATCH pDisp = pToolOld->m_pDispCustom;
							pToolOld->put_Custom(NULL);
							pTool->put_Custom(pDisp);
							break;
						}
					}
				}
				pToolOld->Release();
			}
		}
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// DragDropExchange
//

HRESULT CTools::DragDropExchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		HRESULT hResult;
		CTool*  pTool;
		int		nCount;

		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving everything
			//

			nCount = m_aTools.GetSize();

			hResult = pStream->Write(&nCount, sizeof(DWORD), NULL);
			if (FAILED(hResult))
				return hResult;
			
			for (int nTool = 0; nTool < nCount; nTool++)
			{
				hResult = m_aTools.GetAt(nTool)->DragDropExchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
			//
			// Loading
			//

			hResult = pStream->Read(&nCount, sizeof(DWORD), 0);
			if (FAILED(hResult))
				return hResult;

			for (int nTool = 0; nTool < nCount; nTool++)
			{
				pTool = CTool::CreateInstance(NULL);
				if (NULL == pTool)
					return E_OUTOFMEMORY;
				
				pTool->SetBar(m_pBar);
				pTool->SetBand(m_pBand);
				pTool->m_pTools = this;
	
				hResult = pTool->DragDropExchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;

				hResult = m_aTools.Add(pTool);
				if (FAILED(hResult))
				{
					pTool->Release();
					continue;
				}
			}
		}
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
//
// CreateTool
//

STDMETHODIMP CTools::CreateTool( ITool **pTool)
{
	CTool* pNewTool = CTool::CreateInstance(NULL);
	*pTool = pNewTool;
	if (NULL == pNewTool)
		return E_OUTOFMEMORY;
	pNewTool->SetBar(m_pBar);
	pNewTool->SetBand(m_pBand);
	return NOERROR;
}

//
// CopyTo
//

STDMETHODIMP CTools::CopyTo(ITools* xpDest)
{
	CTools* pDest = (CTools*)xpDest;
	if (pDest == this)
		return NOERROR;

	// Copy tool objects and bitmapmanager info
	HRESULT hResult;
	pDest->RemoveAll();
	CTool* pSourceTool,*pDestTool;
	int nCount = m_aTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pSourceTool = m_aTools.GetAt(nIndex);
		pSourceTool->Clone((ITool**)&pDestTool);
		pDestTool->m_bClone = FALSE;
		hResult = pDest->m_aTools.Add(pDestTool);
		if (FAILED(hResult))
		{
			pDestTool->Release();
			return hResult;
		}
	}
	// Now assign 
	return NOERROR;
}

//
// GetVisibleTools
//

int CTools::GetVisibleTools(TypedArray<CTool*>& aTools)
{
	HRESULT hResult;
	CTool* pTool;
	BOOL bFirstTool = TRUE;

	aTools.SetSize(0, m_aTools.GetSize());

	if (ddBTMenuBar == m_pBand->bpV1.m_btBands || ddBTChildMenuBar == m_pBand->bpV1.m_btBands)
	{
		if (NULL == m_pMDITools[0])
		{
			CreateTool((ITool**)&pTool);
			if (pTool)
			{
				SetTool(pTool, MainMenu[0]);
				pTool->SetBar(m_pBar);
				pTool->SetBand(m_pBand);
				pTool->m_pTools = this;
				m_pMDITools[0] = pTool;
			}
		}
		if (m_pMDITools[0] && CBar::eSysMenu & m_pBar->m_dwMdiButtons)
			aTools.Add(m_pMDITools[0]);
	}

	ToolTypes ttPrev = (ToolTypes)-1;
	int nCount = m_aTools.GetSize();
	for (int nTool = 0; nTool < nCount; nTool++)
	{
		pTool = m_aTools.GetAt(nTool);

		if (VARIANT_TRUE == pTool->tpV1.m_vbVisible  && !(ddTTSeparator == pTool->tpV1.m_ttTools && (NULL == pTool->m_bstrCaption || NULL == *pTool->m_bstrCaption) && bFirstTool))
		{
			if (bFirstTool)
				bFirstTool = FALSE;

			if (ddTTSeparator == ttPrev && ddTTSeparator == pTool->tpV1.m_ttTools)
				continue;

			if (ddTTWindowList == pTool->tpV1.m_ttTools)
				m_pBar->AddWindowList(pTool, m_pBand->bpV1.m_btBands, aTools);
			else
				hResult = aTools.Add(pTool);
			ttPrev = pTool->tpV1.m_ttTools;
		}
		else
			pTool->HidehWnd();
	}

	if (ddBTMenuBar == m_pBand->bpV1.m_btBands || ddBTChildMenuBar == m_pBand->bpV1.m_btBands)
	{
		if (NULL == m_pMDITools[1])
		{
			CreateTool((ITool**)&pTool);
			if (pTool)
			{
				SetTool(pTool, MainMenu[2]);
				pTool->SetBar(m_pBar);
				pTool->SetBand(m_pBand);
				pTool->m_pTools = this;
				m_pMDITools[1] = pTool;
			}
		}

		if (m_pMDITools[1] && ((CBar::eRestoreWindow | CBar::eMinimizeWindow | CBar::eCloseWindowPosition) & m_pBar->m_dwMdiButtons))
			aTools.Add(m_pMDITools[1]);
	}

	if (nCount > 0 && m_pBand->IsMoreTools())
	{
		CreateMoreTools();
		aTools.Add(m_pMoreTools);
	}

	return aTools.GetSize();
}

//
// CreateMoreTools
//

CTool* CTools::CreateMoreTools()
{
	if (NULL == m_pMoreTools)
	{
		CreateTool((ITool**)&m_pMoreTools);
		if (m_pMoreTools)
		{
			DDString strTemp;
			SetTool(m_pMoreTools, MoreTools[0]);
			m_pMoreTools->SetBar(m_pBar);
			m_pMoreTools->SetBand(m_pBand);
			m_pMoreTools->m_pTools = this;
		}
	}
	LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTMoreButton);
	DDString strTemp;
	if (szLocaleString)
		strTemp = szLocaleString;
	else
		strTemp.LoadString(IDS_RESETTOOLBAR);
	BSTR bstr = strTemp.AllocSysString();
	if (bstr)
	{
		m_pMoreTools->put_TooltipText(bstr);
		SysFreeString(bstr);
	}
	return m_pMoreTools;
}
//
// ItemById
//

STDMETHODIMP CTools::ItemById( LONG id, Tool** ppTool)
{
	int nIndex = 0;
	int nCount = m_aTools.GetSize();
	for (nIndex = 0; nIndex < nCount; nIndex++)
		if (m_aTools.GetAt(nIndex)->tpV1.m_nToolId == (ULONG)id)
			break;

	if (nCount == nIndex)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_BADTOOLSINDEX, NULL);
		return CUSTOM_CTL_SCODE(IDERR_BADTOOLSINDEX);
	}

	if (ppTool)
	{
		*ppTool = (Tool*)m_aTools.GetAt(nIndex);
		((IUnknown*)(*ppTool))->AddRef();
	}
	return NOERROR;
}

//
// GetVisibleToolIndex
//

int CTools::GetVisibleToolIndex(const CTool* pTool)
{
	int nCount = m_aVisibleTools.GetSize();
	for (int nTool = 0; nTool < nCount; nTool++)
	{
		if (pTool == m_aVisibleTools.GetAt(nTool))
			return nTool;
	}
	return -1;
}

//
// RemoveAll
//

STDMETHODIMP CTools::RemoveAll()
{
	CleanupVisible();
	long nRefCnt;
	CTool* pTool;
	int nElem = m_aTools.GetSize();
	for (int nTool = 0; nTool < nElem; nTool++)
	{
		pTool = m_aTools.GetAt(0);
		if (m_pBar)
			m_pBar->m_theFormsAndControls.Remove(pTool);
		if (pTool->m_pDispCustom && (ddTTForm != pTool->tpV1.m_tsStyle || ddTTControl != pTool->tpV1.m_tsStyle))
		{
			try
			{
				CCustomTool custTool(pTool->m_pDispCustom);
				custTool.SetHost(NULL);
			}
			catch (...)
			{
				assert(FALSE);
			}
		}
		m_aTools.RemoveAt(0);
		nRefCnt = pTool->Release();
	}
	return NOERROR;
}

//
// FindToolId
//

CTool* CTools::FindToolId(ULONG nId)
{
	int nSize = m_aTools.GetSize();
	for (int nTool = 0; nTool < nSize; nTool++)
	{
		if (nId == m_aTools.GetAt(nTool)->tpV1.m_nToolId)
			return m_aTools.GetAt(nTool);
	}
	return NULL;
}

//
// ParentWindowedTools
//

void CTools::ParentWindowedTools(HWND hWndParent)
{
	CTool* pTool;
	int nToolCount = GetToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = GetTool(nTool);
		if (NULL == pTool)
			continue;

		switch (pTool->tpV1.m_ttTools)
		{
		case ddTTControl:
		case ddTTForm:
			pTool->SetParent(hWndParent);
			break;
		}
	}
}

//
// WindowedToolsHide
//

void CTools::HideWindowedTools()
{
	CTool* pTool;
	int nToolCount = GetToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = GetTool(nTool);
		if (NULL == pTool)
			continue;

		switch (pTool->tpV1.m_ttTools)
		{
		case ddTTControl:
		case ddTTForm:
			pTool->HidehWnd();
			break;
		}
	}
}

//
// TabbedWindowedTools
//

void CTools::TabbedWindowedTools(BOOL bShow)
{
	CTool* pTool;
	int nToolCount = GetToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = GetTool(nTool);
		if (NULL == pTool)
			continue;

		switch (pTool->tpV1.m_ttTools)
		{
		case ddTTControl:
		case ddTTForm:
			pTool->m_bShowTool = bShow;
			if (!bShow)
				pTool->HidehWnd();
			break;
		}
	}
}

//
// TabbedWindowedTools
//

void CTools::ShowTabbedWindowedTools()
{
	CTool* pTool;
	int nToolCount = GetToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = GetTool(nTool);
		if (NULL == pTool)
			continue;

		switch (pTool->tpV1.m_ttTools)
		{
		case ddTTControl:
		case ddTTForm:
			ShowWindow(pTool->m_hWndActive, SW_SHOW);
			break;
		}
	}
}

//
// GetToolIndex
//

int CTools::GetToolIndex(const CTool* pTool)
{
	int nCount = m_aVisibleTools.GetSize();
	for (int nTool = 0; nTool < nCount; nTool++)
	{
		if (pTool == m_aVisibleTools.GetAt(nTool))
			return nTool;
	}
	return -1;
}

//
// DeleteVisibleTool
//

void CTools::DeleteVisibleTool(ITool *pTool)
{
	int nCount = m_aVisibleTools.GetSize();
	for (int nTool = 0; nTool < nCount; nTool++)
	{
		if (((ITool*)m_aVisibleTools.GetAt(nTool)) == pTool)
		{
			m_aVisibleTools.RemoveAt(nTool);
			break;
		}
	}
}

//
// CleanupVisible
//

void CTools::CleanupVisible()
{
	int nCount = m_aVisibleTools.GetSize();
	CTool* pTool;
	for (int nTool = 0; nTool < nCount; nTool++)
	{
		pTool = m_aVisibleTools.GetAt(nTool);
		try
		{
			if (pTool->m_bCreatedInternally)
				pTool->Release();
		}
		catch (...)
		{
			assert(FALSE);
		}
	}
	m_aVisibleTools.RemoveAll();
}

//
// CommitVisibleTools
//

BOOL CTools::CommitVisibleTools(TypedArray<CTool*>& aTools)
{
	int nSize = aTools.GetSize();
	if (nSize > 0 && aTools.GetAt(nSize - 1)->tpV1.m_ttTools == ddTTSeparator)
		aTools.RemoveAt(nSize - 1);
	CleanupVisible();
	m_aVisibleTools.Copy(aTools);
	return TRUE;
}

//
// ItemByIdentity
//

HRESULT CTools::ItemByIdentity(DWORD Identity, CTool** ppTool)
{
	int nIndex = 0;
	int nCount = m_aTools.GetSize();
	for (nIndex = 0; nIndex < nCount; nIndex++)
		if (m_aTools.GetAt(nIndex)->tpV1.m_dwIdentity == (ULONG)Identity)
			break;

	if (nCount == nIndex)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_BADTOOLSINDEX, NULL);
		return CUSTOM_CTL_SCODE(IDERR_BADTOOLSINDEX);
	}

	if (ppTool)
	{
		*ppTool = m_aTools.GetAt(nIndex);
		(*ppTool)->AddRef();
	}
	return NOERROR;
}
