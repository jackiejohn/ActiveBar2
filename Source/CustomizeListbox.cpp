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
#include "Interfaces.h"
#include "ipserver.h"
#include <stddef.h>       // for offsetof()
#include "support.h"
#include "resource.h"
#include "Debug.h"
#include "CategoryMgr.h"
#include "XEvents.h"
#include "Flicker.h"
#include "CustomProxy.h"
#include "Globals.h"
#include "Dispids.h"
#include "Bar.h"
#include "Band.h"
#include "Bands.h"
#include "Resource.h"
#include "Returnbool.h"
#include "Returnstring.h"
#include "Localizer.h"
#include "FDialog.h"
#include "DropSource.h"
#include ".\Designer\DragDrop.h"
#include ".\Designer\DragDropMgr.h"
#include "CustomizeListbox.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

void CustomizeListboxErrorObject::AsyncError(short		    nNumber, 
										     IReturnString* pDescription, 
										     long		    nSCode, 
										     BSTR		    bstrSource, 
										     BSTR		    bstrHelpFile, 
										     long		    nHelpContext, 
										     IReturnBool*   pCancelDisplay)
{
#ifdef _UNICODE
	((CCustomizeListbox*)m_pEvents)->FireError(nNumber, 
											   pDescription, 
											   nSCode, 
											   bstrSource, 
											   bstrHelpFile, 
											   nHelpContext, 
											   pCancelDisplay);
#else
	WCHAR Source[256];
	MultiByteToWideChar(CP_ACP, 0, NAMEOFOBJECT(0), -1, Source, 256);
	MAKE_WIDEPTR_FROMANSI(wszHelpFile, HELPFILEOFOBJECT(0));
	((CCustomizeListbox*)m_pEvents)->FireError(nNumber, 
											   (ReturnString*)pDescription, 
											   nSCode, 
											   bstrSource, 
											   bstrHelpFile, 
											   nHelpContext, 
											   (ReturnBool*)pCancelDisplay);
#endif
	BSTR bstrDescription;
	pDescription->get_Value(&bstrDescription);
	VARIANT_BOOL vbCancelDisplay;
	pCancelDisplay->get_Value(&vbCancelDisplay);
	if (VARIANT_FALSE == vbCancelDisplay && !((CCustomizeListbox*)m_pEvents)->AmbientUserMode())
	{
		MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDescription);
		DDString strError;
		strError.Format(_T("Error %d : %s"), (int)nNumber, szDesc);
		ErrorDialog dlgError(strError);
		dlgError.DoModal(((CCustomizeListbox*)m_pEvents)->m_hWnd);
	}
	SysFreeString(bstrDescription);
}

struct CLItemData
{
	CLItemData() 
		: m_pBand(NULL)
	{
		m_bChecked = FALSE;
	}
	
	~CLItemData() 
	{
		m_pBand->Release();
	};
	BOOL   m_bChecked;
	CBand* m_pBand;
};


//{OLE VERBS}
OLEVERB CCustomizeListboxVerbs[]={
	{OLEIVERB_SHOW,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_HIDE,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_INPLACEACTIVATE,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_PRIMARY,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_PROPERTIES,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_UIACTIVATE,(LPWSTR)NULL,0x0,0x0},
};
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CCustomizeListbox(IUnknown *pUnkOuter)
{
	CCustomizeListbox *pObject;
	pObject=new CCustomizeListbox(pUnkOuter);
	if (!pObject)
		return NULL;
	else
		return pObject->PrivateUnknown();
}

DEFINE_CONTROLOBJECT(&CLSID_CustomizeListbox,CustomizeListbox,_T("CustomizeListbox Class"),CreateFN_CCustomizeListbox,2,0,&IID_ICustomizeListbox,_T(""),
	&DIID_ICustomizeListboxEvents,
	OLEMISC_RECOMPOSEONRESIZE|
	OLEMISC_CANTLINKINSIDE|
	OLEMISC_INSIDEOUT|
	OLEMISC_ACTIVATEWHENVISIBLE|
	OLEMISC_SETCLIENTSITEFIRST,
0,
	FALSE,
	FALSE,
	2,	0,NULL,0,NULL
	);
void *CCustomizeListbox::objectDef=&CustomizeListboxObject;
CCustomizeListbox *CCustomizeListbox::CreateInstance(IUnknown *pUnkOuter)
{
	return new CCustomizeListbox(pUnkOuter);
}
//{OBJECT CREATEFN}
ErrorTable s_etErrors[] =
{
	{0x0, IDS_ERR_UNKNOWN,-1} // this entry needs to be here all the time
};

CCustomizeListbox::CCustomizeListbox(IUnknown *pUnkOuter)
	: m_pUnkOuter(pUnkOuter==NULL ? &m_UnkPrivate : pUnkOuter)
	, m_pActiveBar(NULL),
	  m_pCategoryMgr(NULL),
	  m_bstrCategory(NULL)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
	m_pClientSite=NULL;
	m_pControlSite=NULL;
	m_pInPlaceSiteWndless=NULL;
	m_pSimpleFrameSite=NULL;
	m_pOleAdviseHolder=NULL;
	m_pInPlaceSite=NULL;
	m_fDirty=FALSE;
	m_hWnd=NULL;
	m_hWndParent=NULL;
	m_hWndReflect=NULL;
	m_pInPlaceFrame=NULL;
    m_pInPlaceUIWindow=NULL;
	m_pDispAmbient=NULL;
	m_Size.cx=32;
	m_Size.cy=32;
	m_fInPlaceActive=FALSE;
	m_fUIActive=FALSE;
	m_hRgn=NULL;
	m_fUsingWindowRgn=FALSE;
    m_fViewAdvisePrimeFirst = FALSE;
    m_fViewAdviseOnlyOnce = FALSE;
	m_pViewAdviseSink=NULL;
	m_fSaveSucceeded = FALSE;
	m_cpEvents=new CConnectionPoint(this,SINK_TYPE_EVENT);
	m_cpPropNotify=new CConnectionPoint(this,SINK_TYPE_PROPNOTIFY);
// {END INIT}
	m_Font.SetPropertyNotifySink(&m_xPropNotify);
	m_Font.InitializeFont(&GetGlobals()._fdDefault);
	m_nTextHeight = 0;
	m_nItemHeight = 0;
	m_bDirty = FALSE;
	m_clbType = ddCTBands;
	m_nDropIndex = -1;
	m_bDropTarget = FALSE;	
	m_theErrorObject.m_pErrorTable = s_etErrors;
	m_theErrorObject.m_pEvents = this;
	m_theErrorObject.m_bstrHelpFile = SysAllocString(L"ActiveBar.hlp");
	m_theErrorObject.objectDef = objectDef;
	m_bDontRecreate = TRUE;
	m_vbToolDragDrop = VARIANT_TRUE;
    
	m_pFF = new CFlickerFree;
	assert(m_pFF);

	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		m_crButtonFace = GetSysColor(COLOR_BTNFACE);
		m_crXPMenuBackground = NewColor(hDC, m_crButtonFace, 86);
		m_crXPBandBackground = NewColor(hDC, m_crButtonFace, 20);
		m_crXPSelected = NewColor(hDC, GetSysColor(COLOR_HIGHLIGHT), 68);
		m_crXPSelectedBorder = GetSysColor(COLOR_3DDKSHADOW);
		ReleaseDC(NULL, hDC);
	}
}

CCustomizeListbox::~CCustomizeListbox()
{
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
	if (m_pClientSite) m_pClientSite->Release();
	if (m_pControlSite) m_pControlSite->Release();
	if (m_pSimpleFrameSite) m_pSimpleFrameSite->Release();
	if (m_pOleAdviseHolder) m_pOleAdviseHolder->Release();
	if (m_pInPlaceSite) m_pInPlaceSite->Release();
	if (m_pInPlaceFrame) m_pInPlaceFrame->Release();
    if (m_pInPlaceUIWindow) m_pInPlaceUIWindow->Release();
    if (m_pDispAmbient) m_pDispAmbient->Release();
	if (m_pInPlaceSiteWndless) m_pInPlaceSiteWndless->Release();
	if (m_hRgn) DeleteObject(m_hRgn);
	if (m_cpEvents)
		m_cpEvents->Release();
	if (m_cpPropNotify)
		m_cpPropNotify->Release();
// {END CLEANUP}
	if (m_pActiveBar)
		m_pActiveBar->Release();

	if (m_pCategoryMgr)
		delete m_pCategoryMgr;

	if (m_bstrCategory)
		SysFreeString(m_bstrCategory);

	if (m_pFF)
		delete m_pFF;
}

static CCustomizeListbox *s_pLastControlCreated;
LPTSTR szWndClass_CCustomizeListbox=_T("WndClassCCustomizeListbox");

//
//
//

HWND CCustomizeListbox::InternalCreateWindow(HWND hWnd, RECT& rc)
{
	DWORD dwWindowStyle = 0;
	DWORD dwExWindowStyle = 0;
	TCHAR szWindowTitle[256];
	*szWindowTitle = NULL;

#ifdef DEF_CRITSECTION
    EnterCriticalSection(&g_CriticalSection);
#endif

    if (0 == ((CONTROLOBJECTDESC*)objectDef)->szWndClass) 
        RegisterClassData(); 

#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif

    if (((CONTROLOBJECTDESC*)objectDef)->pfnSubClass && (hWnd != GetParkingWindow())) 
	{
        SIZE size;
        size.cx = rc.right-rc.left;
		size.cy = rc.bottom-rc.top;
        m_hWndReflect = CreateReflectWindow(TRUE, hWnd, rc.left, rc.top, &size, ReflectWindowProc);
		assert(m_hWndReflect);
        GetGlobals().m_pControls->SetAt((LPVOID)m_hWndReflect, (LPVOID)this);
        dwWindowStyle |= WS_VISIBLE;
    } 

	BeforeCreateWindow(&dwWindowStyle, &dwExWindowStyle, szWindowTitle);

    s_pLastControlCreated = this;
    m_fCreatingWindow = TRUE;

    m_hWnd = CreateWindowEx(dwExWindowStyle,
                            ((CONTROLOBJECTDESC*)objectDef)->szWndClass,
                            szWindowTitle,
                            dwWindowStyle,
                            m_hWndReflect ? 0 : rc.left,
                            m_hWndReflect ? 0 : rc.top,
                            rc.right-rc.left, 
							rc.bottom-rc.top,
                            m_hWndReflect ? m_hWndReflect : hWnd,
                            (HMENU)m_clbType+100, 
							g_hInstance, 
							0);
	assert(m_hWnd);
	m_fCreatingWindow = FALSE;
    s_pLastControlCreated = 0;

	AfterCreateWindow();
	assert(m_hWnd);
	return m_hWnd;
}

//{DEF IUNKNOWN MEMBERS}
//=--------------------------------------------------------------------------=
// CCustomizeListbox::CPrivateUnknownObject::QueryInterface
//=--------------------------------------------------------------------------=
inline CCustomizeListbox *CCustomizeListbox::CPrivateUnknownObject::m_pMainUnknown(void)
{
    return (CCustomizeListbox *)((LPBYTE)this - offsetof(CCustomizeListbox, m_UnkPrivate));
}

STDMETHODIMP CCustomizeListbox::CPrivateUnknownObject::QueryInterface(REFIID riid,void **ppvObjOut)
{
    CHECK_POINTER(ppvObjOut);
    if (DO_GUIDS_MATCH(riid,IID_IUnknown)) 
	{
        m_cRef++;
        *ppvObjOut = (IUnknown *)this;
        return S_OK;
    } 
	else
        return m_pMainUnknown()->InternalQueryInterface(riid, ppvObjOut);
}

ULONG CCustomizeListbox::AddRef(void)
{
    return 0;
}
ULONG CCustomizeListbox::Release(void)
{
	return 0;
}
STDMETHODIMP CCustomizeListbox::QueryInterface(REFIID riid,void **ppvObjOut)
{
	return NULL;
}
ULONG CCustomizeListbox::CPrivateUnknownObject::AddRef(void)
{
    return ++m_cRef;
}

ULONG CCustomizeListbox::CPrivateUnknownObject::Release(void)
{
    ULONG cRef = --m_cRef;
    if (!m_cRef)
        delete m_pMainUnknown();
    return cRef;
}
HRESULT CCustomizeListbox::InternalQueryInterface(REFIID riid,void  **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IDispatch))
	{
		*ppvObjOut=(void *)(IDispatch *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ICustomizeListbox))
	{
		*ppvObjOut=(void *)(ICustomizeListbox *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleObject))
	{
		*ppvObjOut=(void *)(IOleObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleWindow))
	{
		*ppvObjOut=(void *)(IOleWindow *)(IOleInPlaceObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleInPlaceObject))
	{
		*ppvObjOut=(void *)(IOleInPlaceObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleInPlaceActiveObject))
	{
		*ppvObjOut=(void *)(IOleInPlaceActiveObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IViewObject))
	{
		*ppvObjOut=(void *)(IViewObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IViewObject2))
	{
		*ppvObjOut=(void *)(IViewObject2 *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPersist))
	{
		*ppvObjOut=(void *)(IPersist *)(IPersistPropertyBag *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPersistPropertyBag))
	{
		*ppvObjOut=(void *)(IPersistPropertyBag *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPersistStreamInit))
	{
		*ppvObjOut=(void *)(IPersistStreamInit *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPersistStorage))
	{
		*ppvObjOut=(void *)(IPersistStorage *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IConnectionPointContainer))
	{
		*ppvObjOut=(void *)(IConnectionPointContainer *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISpecifyPropertyPages))
	{
		*ppvObjOut=(void *)(ISpecifyPropertyPages *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IProvideClassInfo))
	{
		*ppvObjOut=(void *)(IProvideClassInfo *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleControl))
	{
		*ppvObjOut=(void *)(IOleControl *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISupportErrorInfo))
	{
		*ppvObjOut=(void *)(ISupportErrorInfo *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPerPropertyBrowsing))
	{
		*ppvObjOut=(void *)(IPerPropertyBrowsing *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ICategorizeProperties))
	{
		*ppvObjOut=(void *)(ICategorizeProperties *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	#ifdef _DEBUG
	TCHAR tmp[141];
	wsprintf(tmp,_T("QI failed for {0x%08X,0x%X,0x%X,{0x%hX,0x%hX,0x%hX,0x%hX,0x%hX,0x%hX,0x%hX,0x%hX}}\n"),
		(unsigned int)riid.Data1,(unsigned int)riid.Data2,(unsigned int)riid.Data3,(unsigned int)riid.Data4[0],(unsigned int)riid.Data4[1],
		(unsigned int)riid.Data4[2],(unsigned int)riid.Data4[3],(unsigned int)riid.Data4[4],(unsigned int)riid.Data4[5],
		(unsigned int)riid.Data4[6],(unsigned int)riid.Data4[7]);
	TRACET(tmp);
	#endif
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
//{ENDDEF IUNKNOWN MEMBERS}










STDMETHODIMP CCustomizeListbox::get_ToolDragDrop(VARIANT_BOOL *retval)
{
	*retval = m_vbToolDragDrop;
	return NOERROR;
}
STDMETHODIMP CCustomizeListbox::put_ToolDragDrop(VARIANT_BOOL val)
{
	m_vbToolDragDrop = val;
	return NOERROR;
}

STDMETHODIMP CCustomizeListbox::get_Font(IFontDisp** ppRetval)
{
	*ppRetval = m_Font.GetFontDispatch();
	return NOERROR;
}

STDMETHODIMP CCustomizeListbox::put_Font(IFontDisp* pVal)
{
	m_Font.InitializeFont(NULL,pVal);
	ViewChanged();
	return NOERROR;
}
STDMETHODIMP CCustomizeListbox::putref_Font(IFontDisp ** val)
{
	return put_Font(*val);
}

//
// put_ActiveBar
//

STDMETHODIMP CCustomizeListbox::put_ActiveBar(LPDISPATCH val)
{
	if (m_pActiveBar)
	{
		m_pActiveBar->Release();
		m_pActiveBar = NULL;
	}
	
	LPDISPATCH pDispatch;
	HRESULT hResult = val->QueryInterface(IID_IActiveBar2, (void**)&pDispatch);
	if (FAILED(hResult))
		return hResult;
	
	m_pActiveBar = (IActiveBar2*)pDispatch;
	CustomizeListboxTypes clbCurrent = m_clbType;
	m_clbType = (CustomizeListboxTypes) - 1;
	put_Type(clbCurrent);
	return NOERROR;
}

//
// Add
//

STDMETHODIMP CCustomizeListbox::Add(BSTR bstrName, BandTypes btType)
{
	if (ddCTBands != m_clbType)
		return E_FAIL;

	BOOL bAllocatedString = FALSE;
	HRESULT hResult;
	int nResult;
	if (NULL == bstrName || NULL == *bstrName)
	{
		ToolBarName dlgToolBarName((CBar*)m_pActiveBar, ToolBarName::eNew);
		GetNewBandName(bstrName);
		MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
		dlgToolBarName.Name(szName);
		nResult = dlgToolBarName.DoModal(m_hWnd);
		if (IDOK == nResult)
		{
			LPTSTR szName = dlgToolBarName.Name();
			if (lstrlen(szName) > 0)
			{
				SysFreeString(bstrName);
				MAKE_WIDEPTR_FROMTCHAR(wBuffer, szName);
				bAllocatedString = TRUE;
				bstrName = SysAllocString(wBuffer);
			}
		}
		else
		{
			SysFreeString(bstrName);
			bstrName = NULL;
		}
	}
	if (bstrName)
	{
		CBand* pNewBand = CBand::CreateInstance(NULL);
		if (NULL == pNewBand)
			return E_FAIL;

		CBar* pBar = (CBar*)m_pActiveBar;
		if (NULL == pBar)
		{
			pNewBand->Release();
			return E_FAIL;
		}

		pNewBand->SetOwner(pBar, TRUE);

		hResult = pNewBand->put_Caption(bstrName);
		hResult = pNewBand->put_Name(bstrName);
		if (FAILED(hResult))
		{
			pNewBand->Release();
			return hResult;
		}

		pNewBand->bpV1.m_rcFloat.left = pBar->m_ptNewCustomizedFloatBandPos.x;
		pNewBand->bpV1.m_rcFloat.top = pBar->m_ptNewCustomizedFloatBandPos.y;
		pNewBand->bpV1.m_rcFloat.right = pNewBand->bpV1.m_rcFloat.left+32;
		pNewBand->bpV1.m_rcFloat.bottom = pNewBand->bpV1.m_rcFloat.top+32;
		pNewBand->put_DockingArea(ddDAFloat);
		hResult = pNewBand->put_CreatedBy(ddCBUser);

		pBar->m_nNewCustomizedFloatBandCounter++;
		if (7 == (7 & pBar->m_nNewCustomizedFloatBandCounter))
		{
			pBar->m_ptNewCustomizedFloatBandPos.y -= 7 * 32;
			pBar->m_ptNewCustomizedFloatBandPos.x -= 6 * 32;
		}
		else
		{
			pBar->m_ptNewCustomizedFloatBandPos.y += 32;
			pBar->m_ptNewCustomizedFloatBandPos.x += 32;
		}

		IBands* pBands;
		hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
		if (FAILED(hResult))
		{
			pNewBand->Release();
			if (bAllocatedString)
				SysFreeString(bstrName);
			return hResult;
		}

		hResult = pBands->InsertBand(-1, pNewBand);
		pBands->Release();
		if (FAILED(hResult))
		{
			if (bAllocatedString)
				SysFreeString(bstrName);
			return hResult;
		}

		CLItemData* pData = new CLItemData;
		if (pData)
		{
			pData->m_pBand = pNewBand;
			SysFreeString(pData->m_pBand->m_bstrCaption);
			pData->m_pBand->m_bstrCaption = SysAllocString(bstrName);
			pData->m_bChecked = VARIANT_TRUE == pNewBand->bpV1.m_vbVisible ? TRUE : FALSE;
			int nIndex = AddString((LPCTSTR)pData);
			pBar->RecalcLayout();
		}
	}

	if (bAllocatedString)
		SysFreeString(bstrName);
	return NOERROR;
}

//
// Delete
//

STDMETHODIMP CCustomizeListbox::Delete()
{
	if (m_clbType != ddCTBands)
		return E_FAIL;

	int nIndex = GetCurSel();
	if (LB_ERR == nIndex)
		return E_FAIL;

	CLItemData* pData = (CLItemData*)GetItemData(nIndex);
	if (NULL == pData)
		return E_FAIL;

	CBar* pBar = (CBar*)m_pActiveBar;
	if (NULL == pBar)
		return E_FAIL;

	TCHAR szBuffer[MAX_PATH];
	LPCTSTR szLocaleString = pBar->Localizer()->GetString(ddLTDeleteToolString);
	if (NULL == szLocaleString)
		lstrcpyn (szBuffer,LoadStringRes(IDS_DELETETOOLBAR), MAX_PATH);
	else
		lstrcpyn(szBuffer,szLocaleString, MAX_PATH);

	CBand* pBand = (CBand*)pData->m_pBand;
	if (NULL == pBand)
		return E_FAIL;

	MAKE_TCHARPTR_FROMWIDE(szCaption, pBand->m_bstrCaption);
	TCHAR szConfirmString[MAX_PATH];
	_sntprintf(szConfirmString, MAX_PATH, szBuffer, szCaption, MAX_PATH);

	szLocaleString = pBar->Localizer()->GetString(ddLTDeleteToolbarCaption);
	if (NULL == szLocaleString)
		lstrcpyn (szBuffer, LoadStringRes(IDS_DELETETOOLBARTITLE), MAX_PATH);
	else
		lstrcpyn (szBuffer, szLocaleString, MAX_PATH);

	if (IDYES == MessageBox(m_hWnd, szConfirmString, szBuffer, MB_ICONEXCLAMATION|MB_YESNO))
	{
		HRESULT hResult;
		CBand* pPrevselBand = pBar->m_diCustSelection.pBand;
		
		memset(&pBar->m_diCustSelection, 0, sizeof(pBar->m_diCustSelection));

		if (pPrevselBand)
			pPrevselBand->Refresh();

		LRESULT lResult = SendMessage(m_hWnd, LB_DELETESTRING, nIndex, 0);

		IBands* pBands;
		hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
		if (FAILED(hResult))
			return hResult;
		
		pBand->put_Visible(VARIANT_FALSE);

		hResult = ((CBands*)pBands)->RemoveEx(pBand);

		int nCount = SendMessage(m_hWnd, LB_GETCOUNT, 0, 0);
		if (nIndex > 0 && nIndex < nCount)
			SendMessage(m_hWnd, LB_SETCURSEL, nIndex, 0);
		else if (nIndex == nCount)
			SendMessage(m_hWnd, LB_SETCURSEL, nIndex-1, 0);
		pBands->Release();
		pBar->Refresh();
	}

	return NOERROR;
}

//
// Rename
//

STDMETHODIMP CCustomizeListbox::Rename(BSTR bstrNew)
{
	CBar* pBar = (CBar*)m_pActiveBar;
	if (NULL == pBar)
		return E_FAIL;

	if (m_clbType != ddCTBands)
		return E_FAIL;

	int nIndex = GetCurSel();
	if (LB_ERR == nIndex)
		return E_FAIL;

	CLItemData* pData = (CLItemData*)GetItemData(nIndex);
	if (NULL == pData)
		return E_FAIL;

	CBand* pBand = (CBand*)pData->m_pBand;
	if (NULL == pBand)
		return E_FAIL;

	HRESULT hResult = pBand->put_Caption(bstrNew);
	if (SUCCEEDED(hResult))
	{
		SysFreeString(pData->m_pBand->m_bstrCaption);
		pData->m_pBand->m_bstrCaption = SysAllocString(bstrNew);
		assert(pData->m_pBand->m_bstrCaption);
		InvalidateItem(nIndex);
	}
	return NOERROR;
}

//
// Reset
//

STDMETHODIMP CCustomizeListbox::Reset()
{
	if (m_clbType != ddCTBands)
		return E_FAIL;

	int nIndex = GetCurSel();
	if (LB_ERR == nIndex)
		return E_FAIL;

	CLItemData* pData = (CLItemData*)GetItemData(nIndex);
	if (NULL == pData)
		return E_FAIL;

	CBar* pBar = (CBar*)m_pActiveBar;
	if (NULL == pBar)
		return E_FAIL;
	
	CBand* pPrevSelBand = pBar->m_diCustSelection.pBand;
	
	memset(&pBar->m_diCustSelection, 0, sizeof(pBar->m_diCustSelection));

	if (pPrevSelBand)
		pPrevSelBand->Refresh();

	CBand* pBand = (CBand*)pData->m_pBand;
	if (NULL == pBand)
		return E_FAIL;

	pBar->FireReset(pBand->m_bstrName);
	ViewChanged();
	InvalidateRect(m_hWnd, NULL, FALSE);
	return NOERROR;
}

//
// get_Category
//

STDMETHODIMP CCustomizeListbox::get_Category(BSTR *retval)
{
	*retval = NULL;
	if (m_clbType == ddCTBands)
	{
		*retval = NULL;
		return E_FAIL;
	}

	if (ddCTCategories == m_clbType && m_hWnd)
	{
		int nIndex = SendMessage(m_hWnd, LB_GETCURSEL, 0, 0);
		if (LB_ERR == nIndex)
			return E_FAIL;

		CatEntry* pCatEntry = (CatEntry*)GetItemData(nIndex);
		if (NULL == pCatEntry)
			return E_FAIL;

		*retval = SysAllocString(pCatEntry->Category());
		return NOERROR;
	}
	*retval = SysAllocString(m_bstrCategory);
	return NOERROR;
}

//
// put_Category
//

STDMETHODIMP CCustomizeListbox::put_Category(BSTR val)
{
	if (m_clbType == ddCTBands)
		return E_FAIL;

	SysFreeString(m_bstrCategory);
	m_bstrCategory = SysAllocString(val);
	if (ddCTTools == m_clbType)
	{
		if (m_pCategoryMgr && m_bstrCategory && *m_bstrCategory)
		{
			m_pCategoryMgr->FillToolListbox(m_hWnd, m_bstrCategory);
			SendMessage(m_hWnd, LB_SETSEL, (WPARAM)TRUE, 0);
			ITool* pTool = (ITool*)GetItemData(0);
			if (pTool)
				FireSelChangeTool((Tool*)pTool);
		}
	}
	return NOERROR;
}

//
// get_Type
//

STDMETHODIMP CCustomizeListbox::get_Type(CustomizeListboxTypes *retval)
{
	*retval = m_clbType;
	return NOERROR;
}

STDMETHODIMP CCustomizeListbox::put_Type(CustomizeListboxTypes val)
{
	if (m_clbType != val)
	{
		m_clbType = val;
		if (AmbientUserMode())
		{
			if (m_bDontRecreate)
				RecreateControlWindow(); 

			switch (m_clbType)
			{
			case ddCTBands:
				{
					const int nCheckHeight = 13;
					
					// Set the Item Height
					HDC hDC = GetDC(NULL);
					if (hDC)
					{
						HFONT hFontOld = SelectFont(hDC, m_Font.GetFontHandle());

						TEXTMETRIC tm;
						GetTextMetrics(hDC, &tm);
						
						m_nTextHeight = m_nItemHeight = tm.tmHeight;
						if (nCheckHeight > m_nItemHeight)
							m_nItemHeight = nCheckHeight;
					
						SelectFont(hDC, hFontOld);
						ReleaseDC(NULL, hDC);
					}

					SendMessage(m_hWnd, LB_SETITEMHEIGHT, 0, MAKELPARAM(m_nItemHeight, 0));
					
					if (m_pCategoryMgr)
					{
						delete m_pCategoryMgr;
						m_pCategoryMgr = NULL;
					}
					if (NULL == m_pCategoryMgr)
						CreateCategoryManager();
					FillBands();
				}
				break;

			case ddCTCategories:
				{
					if (NULL == m_pCategoryMgr)
						CreateCategoryManager();
					if (m_pCategoryMgr)
					{
						m_pCategoryMgr->FillCategoryListbox(m_hWnd, TRUE);
						SendMessage(m_hWnd, LB_SETSEL, (WPARAM)TRUE, 0);
					}
				}
				break;

			case ddCTTools:
				{
					const int nToolButtonHeight = 26;
					// Set the Item Height
					HDC hDC = GetDC(NULL);
					if (hDC)
					{
						HFONT hFontOld = SelectFont(hDC, m_Font.GetFontHandle());

						TEXTMETRIC tm;
						GetTextMetrics(hDC, &tm);
						
						m_nTextHeight = m_nItemHeight = tm.tmHeight;
						if (nToolButtonHeight > m_nItemHeight)
							m_nItemHeight = nToolButtonHeight;
						
						SelectFont(hDC, hFontOld);
						ReleaseDC(NULL, hDC);
					}
					if (IsWindow(m_hWnd))
						SendMessage(m_hWnd, LB_SETITEMHEIGHT, 0, MAKELPARAM(m_nItemHeight, 0));

					if (NULL == m_pCategoryMgr)
						CreateCategoryManager();

					if (m_pCategoryMgr && m_bstrCategory && *m_bstrCategory)
					{
						m_pCategoryMgr->FillToolListbox(m_hWnd, m_bstrCategory);
						SendMessage(m_hWnd, LB_SETSEL, (WPARAM)TRUE, 0);
						ITool* pTool = (ITool*)GetItemData(0);
						if (pTool && (int)pTool != LB_ERR)
							FireSelChangeTool((Tool*)pTool);
					}
				}
				break;
			}
		}
	}
	return NOERROR;
}

// IDispatch members

STDMETHODIMP CCustomizeListbox::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CCustomizeListbox::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CCustomizeListbox::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;
	hr=pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgdispid);
	pTypeInfo->Release();
	return hr;
}
STDMETHODIMP CCustomizeListbox::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
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

// ICustomizeListbox members


// IOleObject members

STDMETHODIMP CCustomizeListbox::SetClientSite( IOleClientSite *pClientSite)
{
	TRACE(1, "IOleObject::SetClientSite\n");
	HRESULT hResult;
	if (m_fInPlaceActive && pClientSite==NULL)
	{
		hResult = InPlaceDeactivate();
	}
	if (m_pClientSite)
	{
		m_pClientSite->Release();
		m_pClientSite=NULL;
	}
	if (m_pControlSite)
	{
		m_pControlSite->Release();
		m_pControlSite=NULL;
	}
	if (m_pSimpleFrameSite)
	{
		m_pSimpleFrameSite->Release();
		m_pSimpleFrameSite=NULL;
	}
	if (m_pOleAdviseHolder)
	{
		m_pOleAdviseHolder->Release();
		m_pOleAdviseHolder=NULL;
	}
	if (m_pInPlaceSite)
	{
		m_pInPlaceSite->Release();
		m_pInPlaceSite=NULL;
	}
	if (m_pInPlaceFrame)
	{
		m_pInPlaceFrame->Release();
		m_pInPlaceFrame=NULL;
	}
	if (m_pInPlaceUIWindow)
	{
		m_pInPlaceUIWindow->Release();
		m_pInPlaceUIWindow=NULL;
	}
	if (m_pDispAmbient)
	{
		m_pDispAmbient->Release();
		m_pDispAmbient=NULL;
	}
	if (m_pInPlaceSiteWndless)
	{
		m_pInPlaceSiteWndless->Release();
		m_pInPlaceSiteWndless=NULL;
	}

	m_pClientSite=pClientSite;
	if (m_pClientSite)
	{
		m_pClientSite->AddRef();
		m_pClientSite->QueryInterface(IID_IOleControlSite, (void**)&m_pControlSite);
	}
#ifdef DEF_IOLEINPLACEOBJECT
	if (m_fInPlaceActive)
        hResult = InPlaceDeactivate();
#endif
	return S_OK;
}
STDMETHODIMP CCustomizeListbox::GetClientSite( IOleClientSite **ppClientSite)
{
	TRACE(1, "IOleObject::GetClientSite\n");
	*ppClientSite=m_pClientSite;
	if (m_pClientSite)
		m_pClientSite->AddRef();
	return S_OK;
}
STDMETHODIMP CCustomizeListbox::SetHostNames( LPCOLESTR szContainerApp, LPCOLESTR szContainerObj)
{
	return S_OK;
}
STDMETHODIMP CCustomizeListbox::Close( DWORD dwSaveOption)
{
	TRACE(1, "IOleObject::Close\n");
	#ifdef DEF_IOLEINPLACEOBJECT
    HRESULT hr;
    if (m_fInPlaceActive) {
        hr = InPlaceDeactivate();
        RETURN_ON_FAILURE(hr);
    }
	#endif

	if (m_pClientSite)
	{
		m_pClientSite->Release();
		m_pClientSite=NULL;
	} 
	if (m_pControlSite)
	{
		m_pControlSite->Release();
		m_pControlSite=NULL;
	}
	if (m_pSimpleFrameSite)
	{
		m_pSimpleFrameSite->Release();
		m_pSimpleFrameSite=NULL;
	}
	if (m_pInPlaceSite)
	{
		m_pInPlaceSite->Release();
		m_pInPlaceSite=NULL;
	}
    

    // handle the save flag.
    if ((dwSaveOption == OLECLOSE_SAVEIFDIRTY || dwSaveOption == OLECLOSE_PROMPTSAVE) && m_fDirty) 
	{
        if (m_pClientSite) m_pClientSite->SaveObject();
        if (m_pOleAdviseHolder) m_pOleAdviseHolder->SendOnSave();
    }
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::SetMoniker( DWORD dwWhichMoniker, IMoniker *pmk)
{
	return E_NOTIMPL;
}
STDMETHODIMP CCustomizeListbox::GetMoniker( DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
	return E_NOTIMPL;
}
STDMETHODIMP CCustomizeListbox::InitFromData( IDataObject *pDataObject, BOOL fCreation, DWORD reserved)
{
	return E_NOTIMPL;
}
STDMETHODIMP CCustomizeListbox::GetClipboardData( DWORD dwReserved, IDataObject **ppDataObject)
{
	*ppDataObject=NULL;
	return E_NOTIMPL;
}
STDMETHODIMP CCustomizeListbox::DoVerb( LONG lVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect)
{
	TRACE1(1, "IOleObject::DoVerb %d\n",lVerb);
	HRESULT hr;
	switch(lVerb)
	{
		case OLEIVERB_SHOW:
		case OLEIVERB_INPLACEACTIVATE:
		case OLEIVERB_UIACTIVATE:
#ifdef DEF_IOLEINPLACEOBJECT
			hr = InPlaceActivate(lVerb);
#else
			hr=NOERROR;
#endif
			OnVerb(lVerb);
			return (hr);
		case OLEIVERB_HIDE:
			#ifdef DEF_IOLEINPLACEOBJECT
			if (m_fUIActive)
				UIDeactivate();
			#endif
			if (m_fInPlaceVisible) 
				SetInPlaceVisible(FALSE);
			OnVerb(lVerb);
			return S_OK;
		case OLEIVERB_PRIMARY:
		case CTLIVERB_PROPERTIES:
		case OLEIVERB_PROPERTIES:
			{
				// show the frame ourselves if the hose can't.
				if (m_pControlSite) 
				{
					hr = m_pControlSite->ShowPropertyFrame();
					if (hr != E_NOTIMPL)
		                return hr;
				}
				IUnknown *pUnk = (IUnknown *)(IOleObject *)this;
				#ifndef UNICODE
				MAKE_WIDEPTR_FROMANSI(pwsz, NAMEOFOBJECT(m_objectIndex));
				#else
				LPCWSTR pwsz=NAMEOFOBJECT(m_objectIndex);
				#endif
	
				ModalDialog(TRUE);
				hr = OleCreatePropertyFrame(GetActiveWindow(),
											GetSystemMetrics(SM_CXSCREEN) / 2,
											GetSystemMetrics(SM_CYSCREEN) / 2,
											pwsz,
											1,
											&pUnk,
											PROPPAGECOUNT(objectDef),
											(LPCLSID)*(PROPPAGES(objectDef)),
											g_lcidLocale,
											NULL, 
											NULL);
				ModalDialog(FALSE);
				return hr;
			}

	default:
        // if it's a derived-control defined verb, pass it on to them
        if (lVerb > 0) 
		{
            hr = OnVerb(lVerb);
            if (hr==OLEOBJ_S_INVALIDVERB) 
			{
                // unrecognised verb -- just do the primary verb and activate
                hr = InPlaceActivate(OLEIVERB_PRIMARY);
                return (FAILED(hr)) ? hr : OLEOBJ_S_INVALIDVERB;
            } 
			else
				return hr;
        } 
		else 
		{
			TRACE(1, "Unrecognized negative verb in DoVerb().\n");
			return E_NOTIMPL;
        }
        break;
	}
	// dead code
}
STDMETHODIMP CCustomizeListbox::EnumVerbs( IEnumOLEVERB **ppEnumOleVerb)
{
	TRACE(1, "IOleObject::EnumVerbs\n");
	int cVerbs;
	OLEVERB *rgVerbs;
	cVerbs=sizeof(CCustomizeListboxVerbs)/sizeof(OLEVERB);
	if (cVerbs == 0)
        return OLEOBJ_E_NOVERBS;
	if (! (rgVerbs = (OLEVERB *)HeapAlloc(g_hHeap, 0, cVerbs * sizeof(OLEVERB))))
		return E_OUTOFMEMORY;
	memcpy(rgVerbs,CCustomizeListboxVerbs,cVerbs * sizeof(OLEVERB));
	*ppEnumOleVerb=(IEnumOLEVERB *)(IEnumX *) new CEnumX(IID_IEnumOLEVERB,cVerbs, sizeof(OLEVERB), rgVerbs, CopyOLEVERB);
	if (!*ppEnumOleVerb)
        return E_OUTOFMEMORY;
	return S_OK;
}
STDMETHODIMP CCustomizeListbox::Update()
{
	return S_OK;
}
STDMETHODIMP CCustomizeListbox::IsUpToDate()
{
	return S_OK;
}
STDMETHODIMP CCustomizeListbox::GetUserClassID( CLSID *pClsid)
{
#ifdef DEF_PERSIST
	return GetClassID(pClsid); // use IPersist impl
#else
	*pClsid=*(((UNKNOWNOBJECTDESC *)(objectDef))->rclsid);
#endif
	return NOERROR;
}
STDMETHODIMP CCustomizeListbox::GetUserType( DWORD dwFormOfType, LPOLESTR *pszUserType)
{
	*pszUserType=OLESTRFROMANSI(((UNKNOWNOBJECTDESC *)(objectDef))->pszObjectName);
	return (*pszUserType) ? S_OK : E_OUTOFMEMORY;
}
STDMETHODIMP CCustomizeListbox::SetExtent( DWORD dwDrawAspect,  SIZEL *psizel)
{
	TRACE(1, "IOleObject::SetExtent\n");

	if (!(dwDrawAspect&DVASPECT_CONTENT))
		return DV_E_DVASPECT;

	SIZEL pixsizel;
	HiMetricToPixel(psizel,&pixsizel);
	BOOL fAcceptSizing;
	fAcceptSizing=OnSetExtent(&pixsizel);
	if (!fAcceptSizing)
		return E_FAIL;
	m_Size=pixsizel;
	RECT rect;
	if (!m_pInPlaceSiteWndless) 
	{
		if (m_fInPlaceActive) 
		{
			// theoretically, one should not need to call OnPosRectChange
			// here, but there appear to be a few host related issues that
            // will make us keep it here.  we won't, however, both with
            // windowless ole controls, since they are all new hosts who
            // should know better
			GetWindowRect(m_hWnd, &rect);
			MapWindowPoints(NULL, m_hWndParent, (LPPOINT)&rect, 2);
			rect.right = rect.left + m_Size.cx;
			rect.bottom = rect.top + m_Size.cy;
			if (m_pInPlaceSite)
				m_pInPlaceSite->OnPosRectChange(&rect);
			if (m_hWnd) 
			{
				// just go and resize
                if (m_hWndReflect)
					SetWindowPos(m_hWndReflect, 0, 0, 0, m_Size.cx, m_Size.cy,
                                     SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                    SetWindowPos(m_hWnd, 0, 0, 0, m_Size.cx, m_Size.cy,
                                 SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
			} 
		}
		else if (m_hWnd) 
			{
				SetWindowPos(m_hWnd, NULL, 0, 0, m_Size.cx, m_Size.cy, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
			} 
		else 
			{
				ViewChanged();
			}
	} 
	else
		if (m_pInPlaceSite) 
			m_pInPlaceSite->OnPosRectChange(&rect);
	return S_OK;
}
STDMETHODIMP CCustomizeListbox::GetExtent( DWORD dwDrawAspect,  SIZEL *psizel)
{
	TRACE(1, "IOleObject::GetExtent\n");
	if (!(dwDrawAspect&DVASPECT_CONTENT))
		return DV_E_DVASPECT;
	PixelToHiMetric((const SIZEL *)&m_Size, psizel);
	return S_OK;
}
STDMETHODIMP CCustomizeListbox::Advise( IAdviseSink *pAdvSink,  DWORD *pdwConnection)
{
	TRACE(1, "IOleObject::Advise\n");
    HRESULT hr;
    if (!m_pOleAdviseHolder) 
	{
        hr = CreateOleAdviseHolder(&m_pOleAdviseHolder);
        RETURN_ON_FAILURE(hr);
    }
	return m_pOleAdviseHolder->Advise(pAdvSink, pdwConnection);
}
STDMETHODIMP CCustomizeListbox::Unadvise(DWORD dwConnection)
{
	TRACE(1, "IOleObject::Unadvise\n");
    if (!m_pOleAdviseHolder) 
	{
		TRACE(1, "IOleObject::Unadvise call without Advise");
		return CONNECT_E_NOCONNECTION;
    }
    return m_pOleAdviseHolder->Unadvise(dwConnection);
}
STDMETHODIMP CCustomizeListbox::EnumAdvise(IEnumSTATDATA **ppenumAdvise)
{
	TRACE(1, "IOleObject::EnumAdvise\n");
	if (!m_pOleAdviseHolder) 
	{
        TRACE(1, "IOleObject::EnumAdvise call without Advise");
        *ppenumAdvise = NULL;
        return E_FAIL;
    }
    return m_pOleAdviseHolder->EnumAdvise(ppenumAdvise);
}
STDMETHODIMP CCustomizeListbox::GetMiscStatus( DWORD dwAspect,  DWORD *pdwStatus)
{
	TRACE(1, "IOleObject::GetMiscStatus\n");
	if (dwAspect!=DVASPECT_CONTENT)
		return DV_E_DVASPECT;
	*pdwStatus=((CONTROLOBJECTDESC *)objectDef)->dwOleMiscFlags;
	return NOERROR;
}
STDMETHODIMP CCustomizeListbox::SetColorScheme( LOGPALETTE *pLogpal)
{
	return S_OK;
}


//{IOleObject Support Code}

void CCustomizeListbox::ViewChanged()
{
    if (m_pViewAdviseSink) 
	{
        m_pViewAdviseSink->OnViewChange(DVASPECT_CONTENT, -1);
        if (m_fViewAdviseOnlyOnce)
            SetAdvise(DVASPECT_CONTENT, 0, NULL);
    }
}

HRESULT CCustomizeListbox::OnVerb(LONG lVerb)
{
	return OLEOBJ_S_INVALIDVERB;
}

BOOL CCustomizeListbox::RegisterClassData(void)
{
	TRACE(1, "RegisterClassData\n");

#ifdef CCustomizeListbox_SUBCLASS
	WNDCLASS wndclass;
	if (!::GetClassInfo(g_hInstance, CCustomizeListbox_SUBCLASS, &wndclass))
        return FALSE;
    
	// this doesn't need a critical section for apartment threading support
    // since it's already in a critical section in CreateInPlaceWindow
    //
    ((CONTROLOBJECTDESC*)objectDef)->pfnSubClass = (WNDPROC)wndclass.lpfnWndProc;
    wndclass.lpfnWndProc    = CCustomizeListbox::ControlWindowProc;
	wndclass.lpszClassName  = szWndClass_CCustomizeListbox;
	((CONTROLOBJECTDESC*)objectDef)->szWndClass=szWndClass_CCustomizeListbox;
	return RegisterClass(&wndclass);
#else
	WNDCLASS wndclass;
	memset(&wndclass, 0, sizeof(WNDCLASS));
    wndclass.style          = CS_VREDRAW | CS_HREDRAW | CS_DBLCLKS;
    wndclass.lpfnWndProc    = CCustomizeListbox::ControlWindowProc;
    wndclass.hInstance      = g_hInstance;
    wndclass.hCursor        = LoadCursor(NULL, IDC_ARROW);
    wndclass.hbrBackground  = (HBRUSH)(COLOR_WINDOW + 1);
    wndclass.lpszClassName  = szWndClass_CCustomizeListbox;
	((CONTROLOBJECTDESC*)objectDef)->szWndClass=szWndClass_CCustomizeListbox;
    return RegisterClass(&wndclass);
#endif
}

LRESULT CALLBACK CCustomizeListbox::ControlWindowProc(HWND hwnd,UINT msg,WPARAM wParam,LPARAM lParam)
{
	CCustomizeListbox* pCtl = NULL;
	if (!ControlFromHwnd(hwnd, pCtl))
	{
        pCtl = s_pLastControlCreated;
		GetGlobals().m_pControls->SetAt((LPVOID)hwnd, (LPVOID&)pCtl);
        pCtl->m_hWnd = hwnd;
	}

    HRESULT hr;
    LRESULT lResult = 0;
    DWORD   dwCookie;


    // if the value isn't a positive value, then it's in some special
    // state [creation or destruction]  this is safe because under win32,
    // the upper 2GB of an address space aren't available.
    //

    // message preprocessing
    //
    if (pCtl->m_pSimpleFrameSite) 
	{
        hr = pCtl->m_pSimpleFrameSite->PreMessageFilter(hwnd, msg, wParam, lParam, &lResult, &dwCookie);
        if (hr == S_FALSE) return lResult;
    }

    // for certain messages, do not call the user window proc. instead,
    // we have something else we'd like to do.
    //
	switch (msg) 
	{
	case WM_PAINT:
        {
			// call the user's OnDraw routine.
			PAINTSTRUCT ps;
			RECT        rc;
			HDC         hdc;
			// if we're given an HDC, then use it
			if (!wParam)
				hdc = BeginPaint(hwnd, &ps);
			else
				hdc = (HDC)wParam;

			GetClientRect(hwnd, &rc);
			pCtl->OnDraw(DVASPECT_CONTENT, hdc, (RECTL *)&rc, NULL, NULL, TRUE);

			if (!wParam)
				EndPaint(hwnd, &ps);
			return 0;
        }
        break;

	default:
        // call the derived-control's window proc
        lResult = pCtl->WindowProc(msg, wParam, lParam);
        break;
    }

    // message postprocessing
    switch (msg) 
	{
    case OCM_DELETEITEM:
		pCtl->OnDeleteItem((LPDELETEITEMSTRUCT) lParam);
		break;

    case WM_NCDESTROY:
        // after this point, the window doesn't exist any more
		if (GetGlobals().m_pControls->RemoveKey((LPVOID)hwnd))
			pCtl->m_hWnd = NULL;
        break;

    case WM_SETFOCUS:
        // give the control site focus notification
        if (pCtl->m_fInPlaceActive && pCtl->m_pControlSite)
            pCtl->m_pControlSite->OnFocus(TRUE);

        pCtl->m_fUIActive = TRUE;
        break;
	case WM_KILLFOCUS:
        // give the control site focus notification
        if (pCtl->m_fInPlaceActive && pCtl->m_pControlSite)
            pCtl->m_pControlSite->OnFocus(FALSE);
        pCtl->m_fUIActive = FALSE;
        break;
	case WM_SIZE:
        // a change in size is a change in view
        if (!pCtl->m_fCreatingWindow)
            pCtl->ViewChanged();
        break;
 
    case WM_DRAWITEM:
    case OCM_DRAWITEM:
		pCtl->DrawItem((LPDRAWITEMSTRUCT)lParam);
		break;

	case WM_COMMAND:
	case OCM_COMMAND:
		pCtl->OnCommand(HIWORD(wParam));
		break;
    }

    // lastly, simple frame postmessage processing
    if (pCtl->m_pSimpleFrameSite)
        pCtl->m_pSimpleFrameSite->PostMessageFilter(hwnd, msg, wParam, lParam, &lResult, dwCookie);

    return lResult;
}

const BYTE g_rgcbDataTypeSize[] = {
    0,                      // VT_EMPTY= 0,
    0,                      // VT_NULL= 1,
    sizeof(short),          // VT_I2= 2,
    sizeof(long),           // VT_I4 = 3,
    sizeof(float),          // VT_R4  = 4,
    sizeof(double),         // VT_R8= 5,
    sizeof(CURRENCY),       // VT_CY= 6,
    sizeof(DATE),           // VT_DATE = 7,
    sizeof(BSTR),           // VT_BSTR = 8,
    sizeof(IDispatch *),    // VT_DISPATCH    = 9,
    sizeof(SCODE),          // VT_ERROR    = 10,
    sizeof(VARIANT_BOOL),   // VT_BOOL    = 11,
    sizeof(VARIANT),        // VT_VARIANT= 12,
    sizeof(IUnknown *),     // VT_UNKNOWN= 13,
};

BOOL CCustomizeListbox::GetAmbientProperty(DISPID  dispid,VARTYPE vt,void *pData)
{
    DISPPARAMS dispparams;
    VARIANT v, v2;
    HRESULT hr;
    v.vt = VT_EMPTY;
    v.lVal = 0;
    v2.vt = VT_EMPTY;
    v2.lVal = 0;
    // get a pointer to the source of ambient properties.
    if (!m_pDispAmbient) 
	{
        if (m_pClientSite)
            m_pClientSite->QueryInterface(IID_IDispatch, (void **)&m_pDispAmbient);
        if (!m_pDispAmbient)
            return FALSE;
    }

    // now go and get the property into a variant.
    memset(&dispparams, 0, sizeof(DISPPARAMS));
    hr = m_pDispAmbient->Invoke(dispid, IID_NULL, 0, DISPATCH_PROPERTYGET, &dispparams,
                                &v, NULL, NULL);
    if (FAILED(hr)) 
		return FALSE;
    // we've got the variant, so now go an coerce it to the type that the user
    // wants.  if the types are the same, then this will copy the stuff to
    // do appropriate ref counting ...
    hr = VariantChangeType(&v2, &v, 0, vt);
    if (FAILED(hr)) 
	{
        VariantClear(&v);
        return FALSE;
    }
    // copy the data to where the user wants it
    CopyMemory(pData, &(v2.lVal), g_rgcbDataTypeSize[vt]);
    VariantClear(&v);
    return TRUE;
}

BOOL CCustomizeListbox::BeforeCreateWindow(DWORD *pdwWindowStyle,DWORD *pdwExWindowStyle,LPTSTR pszWindowTitle)
{
	*pdwExWindowStyle |= WS_EX_CLIENTEDGE;
	*pdwWindowStyle |= WS_VSCROLL|WS_CHILD|WS_CLIPSIBLINGS|WS_VISIBLE|LBS_NOTIFY|LBS_NOINTEGRALHEIGHT;
	switch (m_clbType)
	{
	case ddCTBands:
		*pdwWindowStyle |= LBS_OWNERDRAWFIXED;
		break;

	case ddCTTools:
		*pdwWindowStyle |= LBS_OWNERDRAWFIXED|LBS_EXTENDEDSEL;
		break;
	}	

    return TRUE;
}

BOOL CCustomizeListbox::AfterCreateWindow()
{
	SendMessage(m_hWnd, WM_SETFONT, (WPARAM)m_Font.GetFontHandle(), MAKELPARAM(TRUE, 0));
    return TRUE;
}

void CCustomizeListbox::BeforeDestroyWindow()
{
	// used for subclassed ones
	SendMessage(m_hWnd, LB_RESETCONTENT, 0, 0);
}

LRESULT CALLBACK CCustomizeListbox::ReflectWindowProc(HWND hwnd,UINT msg,WPARAM  wParam,LPARAM  lParam)
{
    switch (msg) {
        case WM_COMMAND:
        case WM_NOTIFY:
        case WM_CTLCOLORBTN:
        case WM_CTLCOLORDLG:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX:
        case WM_CTLCOLORMSGBOX:
        case WM_CTLCOLORSCROLLBAR:
        case WM_CTLCOLORSTATIC:
        case WM_DRAWITEM:
        case WM_MEASUREITEM:
        case WM_DELETEITEM:
        case WM_VKEYTOITEM:
        case WM_CHARTOITEM:
        case WM_COMPAREITEM:
        case WM_HSCROLL:
        case WM_VSCROLL:
        case WM_PARENTNOTIFY:
        case WM_SETFOCUS:
        case WM_SIZE:
			{
			    CCustomizeListbox* pCtl = NULL;
	            if (GetGlobals().m_pControls->Lookup((LPVOID)hwnd, (LPVOID&)pCtl))
				   return SendMessage(pCtl->m_hWnd, OCM__BASE + msg, wParam, lParam);
			}
            break;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

HRESULT CCustomizeListbox::InPlaceActivate(LONG lVerb)
{
    BOOL f;
    SIZEL sizel;
    IOleInPlaceSiteEx *pIPSEx = NULL;
    HRESULT hr;
    BOOL    fNoRedraw = FALSE;

	TRACE1(1, "InPlaceActivate %d\n",lVerb);

    if (!m_pClientSite) // check if client site exists
        return S_OK;
    // get an InPlace site pointer.
    if (!GetInPlaceSite()) 
	{
        // if they want windowless support, then we want IOleInPlaceSiteWindowless
        if (((CONTROLOBJECTDESC *)objectDef)->fWindowless)
            m_pClientSite->QueryInterface(IID_IOleInPlaceSiteWindowless, (void **)&m_pInPlaceSiteWndless);
        // if we're not able to do windowless siting, then we'll just get an
        // IOleInPlaceSite pointer.
        if (!m_pInPlaceSiteWndless) 
		{
            hr = m_pClientSite->QueryInterface(IID_IOleInPlaceSite, (void **)&m_pInPlaceSite);
            RETURN_ON_FAILURE(hr);
        }
    }

    // now, we want an IOleInPlaceSiteEx pointer for windowless and flicker free
    // activation.  if we're windowless, we've already got it, else we need to
    // try and get it
    //
    if (m_pInPlaceSiteWndless) 
	{
        pIPSEx = (IOleInPlaceSiteEx *)m_pInPlaceSiteWndless;
        pIPSEx->AddRef();
    }
	else
        m_pClientSite->QueryInterface(IID_IOleInPlaceSiteEx, (void **)&pIPSEx);

    // if we're not already active, go and do it.
    if (!m_fInPlaceActive) 
	{
        OLEINPLACEFRAMEINFO InPlaceFrameInfo;
        RECT rcPos, rcClip;

        // if we have a windowless site, see if we can go in-place windowless active
        hr = S_FALSE;
        if (m_pInPlaceSiteWndless) 
		{
            hr = m_pInPlaceSiteWndless->CanWindowlessActivate();
            CLEANUP_ON_FAILURE(hr);
		    // if they refused windowless, we'll try windowed
            if (S_OK != hr) 
			{
                RELEASE_OBJECT(m_pInPlaceSiteWndless);
                hr = m_pClientSite->QueryInterface(IID_IOleInPlaceSite, (void **)&m_pInPlaceSite);
                CLEANUP_ON_FAILURE(hr);
            }
        }
        // just try regular windowed in-place activation
        if (hr != S_OK) 
		{
            hr = m_pInPlaceSite->CanInPlaceActivate();
            if (hr != S_OK) 
			{
                hr = (FAILED(hr)) ? E_FAIL : hr;
                goto CleanUp;
            }
        }

        // if we are here, then we have permission to go in-place active.
        // now, announce our intentions to actually go ahead and do this.
        hr = (pIPSEx) ? pIPSEx->OnInPlaceActivateEx(&fNoRedraw, (m_pInPlaceSiteWndless) ? ACTIVATE_WINDOWLESS : 0)
                       : m_pInPlaceSite->OnInPlaceActivate();
        CLEANUP_ON_FAILURE(hr);

        // if we're here, we're ready to go in-place active.  we just need
        // to set up some flags, and then create the window [if we have one]
   
		m_fInPlaceActive = TRUE;

        // we need to get some information about our location in the parent
        // window, as well as some information about the parent
        //
        InPlaceFrameInfo.cb = sizeof(OLEINPLACEFRAMEINFO);
		HWND prevParent=m_hWndParent;
        hr = GetInPlaceSite()->GetWindow(&m_hWndParent);
        if (SUCCEEDED(hr))
            hr = GetInPlaceSite()->GetWindowContext(&m_pInPlaceFrame, &m_pInPlaceUIWindow, &rcPos, &rcClip, &InPlaceFrameInfo);
        CLEANUP_ON_FAILURE(hr);

        // make sure we'll display ourselves in the correct location with the correct size
        sizel.cx = rcPos.right - rcPos.left;
        sizel.cy = rcPos.bottom - rcPos.top;
        f = OnSetExtent(&sizel);
        if (f) m_Size = sizel;
#ifdef DEF_IOLEINPLACEOBJECT
        SetObjectRects(&rcPos, &rcClip);
#endif
        // finally, create our window if we have to!
        if (!m_pInPlaceSiteWndless) 
		{
			// SetInPlaceParent(m_hwndParent);
			ASSERT(!m_pInPlaceSiteWndless, "this should execute for windowed OLE controls only");
			if (prevParent!= m_hWndParent)
			{
				if (m_hWnd)
				    SetParent(GetOuterWindow(),m_hWndParent);
			}
			// end set inplace parent

            // create the window, and display it.  
            if (!CreateInPlaceWindow(rcPos.left, rcPos.top, fNoRedraw)) 
			{
                hr = E_FAIL;
                goto CleanUp;
            }
        }
    }

    RELEASE_OBJECT(pIPSEx);

    // if we're not inplace visible yet, do so now.
    if (!m_fInPlaceVisible)
        SetInPlaceVisible(TRUE);

    // if we weren't asked to UIActivate, then we're done.
    if (lVerb != OLEIVERB_PRIMARY && lVerb != OLEIVERB_UIACTIVATE)
        return S_OK;

    // if we're not already UI active, do sow now.
    if (!m_fUIActive) 
	{
        m_fUIActive = TRUE;
        // inform the container of our intent
        GetInPlaceSite()->OnUIActivate();
        // take the focus  [which is what UI Activation is all about !]
        SetFocus(TRUE);
        // set ourselves up in the host.
        m_pInPlaceFrame->SetActiveObject((IOleInPlaceActiveObject *)this, NULL);
        if (m_pInPlaceUIWindow)
            m_pInPlaceUIWindow->SetActiveObject((IOleInPlaceActiveObject *)this, NULL);
        // we have to explicitly say we don't wany any border space.
        m_pInPlaceFrame->SetBorderSpace(NULL);
        if (m_pInPlaceUIWindow)
            m_pInPlaceUIWindow->SetBorderSpace(NULL);
    }
    return S_OK;

  CleanUp:
    // something bad happened 
    if (pIPSEx) 
		pIPSEx->Release();
    m_fInPlaceActive = FALSE;
    return hr;
}

HWND CCustomizeListbox::CreateInPlaceWindow(int  x,int  y,BOOL fNoRedraw)
{
    BOOL    fVisible;
    HRESULT hr;
    DWORD   dwWindowStyle, dwExWindowStyle;
    TCHAR szWindowTitle[128];

	TRACE(1, "CreateInPlaceWindow\n");

    // if we've already got a window, do nothing.
    if (m_hWnd)
        return m_hWnd;
    // get the user to register the class if it's not already
    // been done.  we have to critical section this since more than one thread
    // can be trying to create this control
    //
#ifdef DEF_CRITSECTION
    EnterCriticalSection(&g_CriticalSection);
#endif

    if (((CONTROLOBJECTDESC *)objectDef)->szWndClass==NULL) 
	{
        if (!RegisterClassData()) 
		{
#ifdef DEF_CRITSECTION
            LeaveCriticalSection(&g_CriticalSection);
#endif
            return NULL;
        } 
    }
#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif

    // let the user set up things like the window title, the
    // style, and anything else they feel interested in fiddling
    // with.
    dwWindowStyle = dwExWindowStyle = 0;
    szWindowTitle[0] = '\0';
    if (!BeforeCreateWindow(&dwWindowStyle, &dwExWindowStyle, szWindowTitle))
        return NULL;

    dwWindowStyle |= (WS_CHILD | WS_CLIPSIBLINGS);

    // create window visible if parent hidden (common case)
    // otherwise, create hidden, then shown.  this is a little subtle, but
    // it makes sense eventually.
    //
    if (!m_hWndParent)
        m_hWndParent = GetParkingWindow();

    fVisible = IsWindowVisible(m_hWndParent);

    // This one kinda sucks -- if a control is subclassed, and we're in
    // a host that doesn't support Message Reflecting, we have to create
    // the user window in another window which will do all the reflecting.
    // VERY blech. [don't however, bother in design mode]
    //
    if (((CONTROLOBJECTDESC *)objectDef)->pfnSubClass
		&& (m_hWndParent != GetParkingWindow())) 
	{
        // determine if the host supports message reflecting.
        if (!m_fCheckedReflecting) 
		{
            VARIANT_BOOL f;
            hr = GetAmbientProperty(DISPID_AMBIENT_MESSAGEREFLECT, VT_BOOL, &f);
            if (FAILED(hr) || !f)
                m_fHostReflects = FALSE;
            m_fCheckedReflecting = TRUE;
        }
        // if the host doesn't support reflecting, then we have to create
        // an extra window around the control window, and then parent it
        // off that.
        if (!m_fHostReflects) 
		{
            ASSERT(m_hWndReflect == NULL, "Where'd this come from?");
            m_hWndReflect = CreateReflectWindow(!fVisible, m_hWndParent, x, y, &m_Size,ReflectWindowProc);
            if (!m_hWndReflect)
                return NULL;
			GetGlobals().m_pControls->SetAt((LPVOID)m_hWndReflect, (LPVOID)this);
            dwWindowStyle |= WS_VISIBLE;
        }
    } 
	else 
	{
        if (!fVisible)
            dwWindowStyle |= WS_VISIBLE;
    }

    // we have to mutex the entire create window process since we need to use
    // the s_pLastControlCreated to pass in the object pointer.  nothing too
    // serious
    //
#ifdef DEF_CRITSECTION
    EnterCriticalSection(&g_CriticalSection);
#endif

    s_pLastControlCreated = this;
    m_fCreatingWindow = TRUE;

    // finally, go create the window, parenting it as appropriate.
    //
    m_hWnd = CreateWindowEx(dwExWindowStyle,
                            ((CONTROLOBJECTDESC *)objectDef)->szWndClass,
                            szWindowTitle,
                            dwWindowStyle,
                            (m_hWndReflect) ? 0 : x,
                            (m_hWndReflect) ? 0 : y,
                            m_Size.cx, m_Size.cy,
                            (m_hWndReflect) ? m_hWndReflect : m_hWndParent,
                            NULL, g_hInstance, NULL);

    // clean up some variables, and leave the critical section
    //
    m_fCreatingWindow = FALSE;
    s_pLastControlCreated = NULL;
#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif
    if (m_hWnd) 
	{
        // let the derived-control do something if they so desire
        if (!AfterCreateWindow()) 
		{
            BeforeDestroyWindow();
            DestroyWindow(m_hWnd);
            m_hWnd = NULL;
            return m_hWnd;
        }

        // if we didn't create the window visible, show it now.
        if (fVisible)
		{
			if (GetParent(m_hWnd) != m_hWndParent)
				// SetWindowPos fails if parent hwnd is passed in so keep
				// this behaviour in those cases without ripping. 
				SetWindowPos(m_hWnd, m_hWndParent, 0, 0, 0, 0,
							 SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_SHOWWINDOW | ((fNoRedraw) ? SWP_NOREDRAW : 0));
		}
    }
    // finally, tell the host of this
    if (m_pClientSite)
        m_pClientSite->ShowObject();
    return m_hWnd;
}


BOOL CCustomizeListbox::SetFocus(BOOL fGrab)
{
    HRESULT hr;
    HWND    hwnd;
	
	TRACE(1, "SetFocus\n");

    // first thing to do is check out UI Activation state, and then set
    // focus [either with windows api, or via the host for windowless
    // controls]
    if (m_pInPlaceSiteWndless) 
	{
        if (!m_fUIActive && fGrab)
            if (FAILED(InPlaceActivate(OLEIVERB_UIACTIVATE))) return FALSE;

        hr = m_pInPlaceSiteWndless->SetFocus(fGrab);
        return (hr == S_OK) ? TRUE : FALSE;
    }
	else 
	{
	    // we've got a window.
        if (m_fInPlaceActive) 
		{
            hwnd= (fGrab) ? m_hWnd : m_hWndParent;
            if (!m_fUIActive && fGrab)
                return SUCCEEDED(InPlaceActivate(OLEIVERB_UIACTIVATE));
            else
			{
				// Notes:
				//    we do this because some controls host non-ole window hierarchies, like
				// the Netscape plugin ocx.  in such cases, the control may need to be UIActive
				// to function properly in the document, but cannot take the windows focus
				// away from one of its child windows.  such controls may override this method
				// and respond as appropriate.
				return (::SetFocus(hwnd) == hwnd);
			}
        }
		else
            return FALSE;
    }
}


void CCustomizeListbox::SetInPlaceVisible(BOOL fShow)
{
	BOOL fVisible;
    m_fInPlaceVisible = fShow;

	TRACE1(1, "SetInPlaceVisible %d\n",(int)fShow);

    // don't do anything if we don't have a window.  otherwise, set it
    if (m_hWnd) 
	{
        fVisible = ((GetWindowLong(GetOuterWindow(), GWL_STYLE) & WS_VISIBLE) != 0);
        if (fVisible && !fShow)
            ShowWindow(GetOuterWindow(), SW_HIDE);
        else if (!fVisible && fShow)
            ShowWindow(GetOuterWindow(), SW_SHOWNA);
    }
}

BOOL CCustomizeListbox::OnSetExtent(const SIZEL *sizel)
{
	return TRUE;
}

void CCustomizeListbox::ModalDialog(BOOL fShow)
{
    if (m_pInPlaceFrame)
        m_pInPlaceFrame->EnableModeless(!fShow);
}

LRESULT CCustomizeListbox::WindowProc(UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_ERASEBKGND:
		{
			CRect rcClient;
			GetClientRect(m_hWnd, &rcClient);
			if (ddCTTools == m_clbType)
			{
				LRESULT lResult = CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, m_hWnd, msg, wParam, lParam);	
				VARIANT_BOOL vbXPLook = VARIANT_FALSE;
				if (m_pActiveBar)
					m_pActiveBar->get_XPLook(&vbXPLook);

				if (VARIANT_FALSE == vbXPLook)
					FillRect((HDC)wParam, &rcClient, (HBRUSH)(1+COLOR_BTNFACE));
				else
				{
					HBRUSH aBrush = CreateSolidBrush(m_crXPMenuBackground);
					if (aBrush)
					{
						FillRect((HDC)wParam, &rcClient, aBrush);
						DeleteBrush(aBrush);
					}
					CRect rcTool = rcClient;
					rcTool.right = rcTool.left + 26;
					FillRect((HDC)wParam, &rcTool, (HBRUSH)(1+COLOR_BTNFACE));
				}
				return lResult;
			}
		}
		break;

	case WM_CTLCOLORLISTBOX:
		if (ddCTTools == m_clbType)
			return (1+COLOR_BTNFACE);
		break;

    case WM_LBUTTONDBLCLK:
    case WM_LBUTTONDOWN:
		{
			::SetFocus(m_hWnd);
			DWORD dwHitValue = (DWORD)SendMessage(m_hWnd, LB_ITEMFROMPOINT, 0, lParam);
			
			int nIndex = LOWORD(dwHitValue);
			if (nIndex >= 0 && nIndex < SendMessage(m_hWnd, LB_GETCOUNT, 0, 0) && 1 != HIWORD(dwHitValue))
			{
				((CBar*)m_pActiveBar)->m_diCustSelection.pTool = NULL;

				//
				// Yes selected or is about to be selected . No control/shift key should not be pressed
				//

				if (0 == (GetKeyState(VK_CONTROL)&0x8000) && 0 == (GetKeyState(VK_SHIFT)&0x8000))
				{
					// check if selected
					if (!SendMessage(m_hWnd, LB_GETSEL, nIndex, 0))
					{ 
						// not selected so do select
						CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, m_hWnd, WM_LBUTTONDOWN, wParam, lParam);	
						CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, m_hWnd, WM_LBUTTONUP, wParam, lParam);	
					}
					POINT pt;
					pt.x = LOWORD(lParam);
					pt.y = HIWORD(lParam);  
					OnLButtonDown(wParam, pt);
					return 0;
				}
			}
		}
		break;

    case WM_RBUTTONDOWN:
		try
		{
			switch (m_clbType)
			{
			case ddCTTools:
				{
					UINT nId = GetWindowLong(m_hWnd, GWL_ID);
					NMHDR nmHdr;
					nmHdr.hwndFrom = m_hWnd;
					nmHdr.idFrom = nId;
					nmHdr.code = NM_RCLICK;

					DWORD nIndex = SendMessage(m_hWnd, LB_ITEMFROMPOINT, 0, lParam);
					if (!SendMessage(m_hWnd, LB_GETSEL, LOWORD(nIndex), 0))
					{
						CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, m_hWnd, WM_LBUTTONDOWN, wParam, lParam);	
						CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, m_hWnd, WM_LBUTTONUP, wParam, lParam);	
					}
					SendMessage(GetParent(m_hWnd), WM_NOTIFY, nId, (LPARAM)&nmHdr);
				}
				break;
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		break;
	}

#ifdef CCustomizeListbox_SUBCLASS
	return CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, m_hWnd, msg, wParam, lParam);	
#else
	return DefWindowProc(m_hWnd,msg,wParam,lParam);
#endif
}
//{IOleObject Support Code}

// IOleWindow members

STDMETHODIMP CCustomizeListbox::GetWindow( HWND *phwnd)
{
    if (m_pInPlaceSiteWndless)
        return E_FAIL;
    *phwnd = GetOuterWindow();
    return (*phwnd) ? S_OK : E_UNEXPECTED;
}
STDMETHODIMP CCustomizeListbox::ContextSensitiveHelp( BOOL fEnterMode)
{
	return E_NOTIMPL;
}

// IOleInPlaceObject members

STDMETHODIMP CCustomizeListbox::InPlaceDeactivate()
{
    if (!m_fInPlaceActive)
        return S_OK;
    // transition from UIActive back to active
    if (m_fUIActive)
        UIDeactivate();
    m_fInPlaceActive = FALSE;
    m_fInPlaceVisible = FALSE;
    // if we have a window, tell it to go away.
    if (m_hWnd) 
	{
        ASSERT(!m_pInPlaceSiteWndless, "internal state really messed up");
        // so our window proc doesn't crash.
        BeforeDestroyWindow();
        DestroyWindow(m_hWnd);
        m_hWnd = NULL;

        if (m_hWndReflect) 
		{
			GetGlobals().m_pControls->RemoveKey((LPVOID)m_hWndReflect);
            DestroyWindow(m_hWndReflect);
            m_hWndReflect = NULL;
        }
    }
    RELEASE_OBJECT(m_pInPlaceFrame);
    RELEASE_OBJECT(m_pInPlaceUIWindow);
    GetInPlaceSite()->OnInPlaceDeactivate();
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::UIDeactivate()
{
    if (!m_fUIActive)
        return S_OK;
    m_fUIActive = FALSE;
    // notify frame windows, if appropriate, that we're no longer ui-active.
    if (m_pInPlaceUIWindow) 
		m_pInPlaceUIWindow->SetActiveObject(NULL, NULL);
    m_pInPlaceFrame->SetActiveObject(NULL, NULL);
    // we don't need to explicitly release the focus here since somebody else grabbing it
    GetInPlaceSite()->OnUIDeactivate(FALSE);
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::SetObjectRects( LPCRECT lprcPosRect, LPCRECT lprcClipRect)
{
    BOOL fRemoveWindowRgn;
    // move our window to the new location and handle clipping. not applicable
    // for windowless controls, since the container will be responsible for all
    // clipping.
    if (m_hWnd) 
	{
        fRemoveWindowRgn = m_fUsingWindowRgn;
        if (lprcClipRect) 
		{
            // the container wants us to clip, so figure out if we really
            // need to
            //
            RECT rcIXect;
            if ( IntersectRect(&rcIXect, lprcPosRect, lprcClipRect) ) {
                if (!EqualRect(&rcIXect, lprcPosRect)) {
                    OffsetRect(&rcIXect, -(lprcPosRect->left), -(lprcPosRect->top));

                    HRGN tempRgn = CreateRectRgnIndirect(&rcIXect);
                    SetWindowRgn(GetOuterWindow(), tempRgn, TRUE);

                    if (m_hRgn != NULL)
                       DeleteObject(m_hRgn);
                    m_hRgn = tempRgn;

                    m_fUsingWindowRgn = TRUE;
                    fRemoveWindowRgn  = FALSE;
                }
            }
        }

        if (fRemoveWindowRgn) 
		{
            SetWindowRgn(GetOuterWindow(), NULL, TRUE);
            if (m_hRgn != NULL)
            {
               DeleteObject(m_hRgn);
               m_hRgn = NULL;
            }
            m_fUsingWindowRgn = FALSE;
        }

        // set our control's location, but don't change it's size at all
        // [people for whom zooming is important should set that up here]
        DWORD dwFlag=0;
//        OnSetObjectRectsChangingWindowPos(&dwFlag);
        int cx, cy;
        cx = lprcPosRect->right - lprcPosRect->left;
        cy = lprcPosRect->bottom - lprcPosRect->top;
        SetWindowPos(GetOuterWindow(), NULL, lprcPosRect->left, lprcPosRect->top, cx, cy, dwFlag | SWP_NOZORDER | SWP_NOACTIVATE);
    }
    // save out our current location.  windowless controls want this 
    m_rcLocation = *lprcPosRect;
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::ReactivateAndUndo()
{
	return E_NOTIMPL;
}

// IOleInPlaceActiveObject members

STDMETHODIMP CCustomizeListbox::TranslateAccelerator( LPMSG lpmsg)
{
	extern short _SpecialKeyState();
    // see if we want it or not.
    if (OnSpecialKey(lpmsg))
        return S_OK;
    // if not, then we want to forward it back to the site for further processing
    if (m_pControlSite)
        return m_pControlSite->TranslateAccelerator(lpmsg, _SpecialKeyState());
    // we didn't want it.
    return S_FALSE;
}
STDMETHODIMP CCustomizeListbox::OnFrameWindowActivate( BOOL fActivate)
{
    return InPlaceActivate(OLEIVERB_UIACTIVATE);
}
STDMETHODIMP CCustomizeListbox::OnDocWindowActivate( BOOL fActivate)
{
    return InPlaceActivate(OLEIVERB_UIACTIVATE);
}
STDMETHODIMP CCustomizeListbox::ResizeBorder( LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow)
{
	// we have no border
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::EnableModeless( BOOL fEnable)
{
    return S_OK;
}


//{IOleInPlaceActiveObject Support Code}
static short _SpecialKeyState()
{
    // don't appear to be able to reduce number of calls to GetKeyState
    //
    BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
    BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
    BOOL bAlt   = (GetKeyState(VK_MENU) < 0);

    return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}


BOOL CCustomizeListbox::OnSpecialKey(LPMSG pmsg)
{
    return FALSE;
}
//{IOleInPlaceActiveObject Support Code}

// IViewObject members

STDMETHODIMP CCustomizeListbox::Draw( DWORD dwDrawAspect, LONG lindex, void *pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDev, HDC hdcDraw, LPCRECTL lprcBounds, LPCRECTL lprcWBounds, BOOL ( STDMETHODCALLTYPE *pfnContinue )(DWORD dwContinue), DWORD dwContinue)
{
    HRESULT hr;
    RECTL rc;
    POINT pVp, pW;
    BOOL  fOptimize = FALSE;
    int iMode;
    BYTE fMetafile = FALSE;
    BYTE fDeleteDC = FALSE;

    // support the aspects required for multi-pass drawing
    switch (dwDrawAspect) 
	{
        case DVASPECT_CONTENT:
        case DVASPECT_OPAQUE:
        case DVASPECT_TRANSPARENT:
            break;
        default:
            return DV_E_DVASPECT;
    }

    // first, have to do a little bit to support printing.
    if (GetDeviceCaps(hdcDraw, TECHNOLOGY) == DT_METAFILE) 
	{
	    // We are dealing with a metafile.
        fMetafile = TRUE;
        // If attributes DC is NULL, create one, based on ptd.
        if (!hicTargetDev) 
		{
		    // Does _CreateOleDC have to return an hDC
            // or can it be flagged to return an hIC 
            // for this particular case?
            hicTargetDev = _CreateOleDC(ptd);
            fDeleteDC = TRUE;
        }
    }
    // check to see if we have any flags passed in the pvAspect parameter.
    if (pvAspect && ((DVASPECTINFO *)pvAspect)->cb == sizeof(DVASPECTINFO))
        fOptimize = (((DVASPECTINFO *)pvAspect)->dwFlags & DVASPECTINFOFLAG_CANOPTIMIZE) ? TRUE : FALSE;
    // if we are windowless, then we just pass this on to the end control code.
    if (m_fInPlaceActive) 
	{
	    // give them a rectangle with which to draw
        //ASSERT(!m_fInPlaceActive || !lprcBounds, "Inplace active and somebody passed in prcBounds!!!");
        if (lprcBounds)
			memcpy(&rc, lprcBounds, sizeof(rc));
		else
			memcpy(&rc, &m_rcLocation, sizeof(rc));
    } 
	else 
	{
	    // first -- convert the DC back to MM_TEXT mapping mode so that the
        // window proc and OnDraw can share the same painting code.  save
        // some information on it, so we can restore it later [without using
        // a SaveDC/RestoreDC]
        rc = *lprcBounds;
        // Don't do anything to hdcDraw if it's a metafile.
        // The control's Draw method must make the appropriate
        // accomodations for drawing to a metafile
        if (!fMetafile) 
		{
            LPtoDP(hdcDraw, (POINT *)&rc, 2);
            SetViewportOrgEx(hdcDraw, 0, 0, &pVp);
            SetWindowOrgEx(hdcDraw, 0, 0, &pW);
            iMode = SetMapMode(hdcDraw, MM_TEXT);
        }
    }

    // prcWBounds is NULL and not used if we are not dealing with a metafile.
    // For metafiles, we pass on rc as *prcBounds, we should also include
    // prcWBounds
    hr = OnDraw(dwDrawAspect, hdcDraw, &rc, lprcWBounds, hicTargetDev, fOptimize);
    // clean up the DC when we're done with it, if appropriate.
    if (!m_fInPlaceActive) 
	{
        SetViewportOrgEx(hdcDraw, pVp.x, pVp.y, NULL);
        SetWindowOrgEx(hdcDraw, pW.x, pW.y, NULL);
        SetMapMode(hdcDraw, iMode);
    }

    if (fDeleteDC) 
		DeleteDC(hicTargetDev);
    return hr;
}
STDMETHODIMP CCustomizeListbox::GetColorSet( DWORD dwDrawAspect, LONG lindex, void *pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDev, LOGPALETTE **ppColorSet)
{
	if (dwDrawAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    *ppColorSet = NULL;
    return (OnGetPalette(hicTargetDev, ppColorSet)) ? 
			((*ppColorSet) ? S_OK : S_FALSE) : E_NOTIMPL;
}
STDMETHODIMP CCustomizeListbox::Freeze( DWORD dwDrawAspect, LONG lindex, void *pvAspect, DWORD *pdwFreeze)
{
	return E_NOTIMPL;
}
STDMETHODIMP CCustomizeListbox::Unfreeze( DWORD dwFreeze)
{
	return E_NOTIMPL;
}
STDMETHODIMP CCustomizeListbox::SetAdvise( DWORD dwAspects, DWORD dwAdviseFlags, IAdviseSink *pAdviseSink)
{
	TRACE(1, "IViewObject::SetAdvise\n");
    if (!(dwAspects & DVASPECT_CONTENT))
        return DV_E_DVASPECT;
    m_fViewAdvisePrimeFirst = (dwAdviseFlags & ADVF_PRIMEFIRST) ? TRUE : FALSE;
    m_fViewAdviseOnlyOnce = (dwAdviseFlags & ADVF_ONLYONCE) ? TRUE : FALSE;

	// cache advise sink interface
    if (m_pViewAdviseSink)
		m_pViewAdviseSink->Release();
    m_pViewAdviseSink = pAdviseSink;
	if (m_pViewAdviseSink)
		m_pViewAdviseSink->AddRef();
    
    // prime them if they want it [we need to store this so they can get flags later]
    if (m_fViewAdvisePrimeFirst)
        ViewChanged();
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::GetAdvise( DWORD *pdwAspects, DWORD *pdwAdviseFlags, IAdviseSink **ppAdviseSink)
{
	TRACE(1, "IViewObject::GetAdvise\n");
    if (pdwAspects)
        *pdwAspects = DVASPECT_CONTENT;
    if (pdwAdviseFlags) 
	{
        *pdwAdviseFlags = 0;
        if (m_fViewAdviseOnlyOnce) 
			*pdwAdviseFlags |= ADVF_ONLYONCE;
        if (m_fViewAdvisePrimeFirst) 
			*pdwAdviseFlags |= ADVF_PRIMEFIRST;
    }
    if (ppAdviseSink) 
	{
        *ppAdviseSink = m_pViewAdviseSink;
        if (*ppAdviseSink)
			(*ppAdviseSink)->AddRef();
    }
    return S_OK;
}


//{IViewObject Support Code}
STDMETHODIMP CCustomizeListbox::OnDraw(DWORD dvAspect, HDC hdcDraw, LPCRECTL prcBounds, LPCRECTL prcWBounds, HDC hicTargetDev, BOOL fOptimize)
{
#ifdef CCustomizeListbox_SUBCLASS
	return DoSuperClassPaint(hdcDraw,prcBounds);
#else
	return NOERROR;
#endif
}

BOOL CCustomizeListbox::OnGetPalette(HDC hicTargetDevice,LOGPALETTE **ppColorSet)
{
	// use if control supports palettes
    return FALSE;
}

#ifdef CCustomizeListbox_SUBCLASS
HRESULT CCustomizeListbox::DoSuperClassPaint(HDC hdc,LPCRECTL prcBounds)
{
    HWND hwnd;
    RECT rcClient;
    int  iMapMode;
    POINT ptWOrg, ptVOrg;
    SIZE  sWOrg, sVOrg;

    // make sure we have a window.
    //
    hwnd = CreateInPlaceWindow(0,0, FALSE);
    if (!hwnd)
        return E_FAIL;

    GetClientRect(hwnd, &rcClient);

    if ((rcClient.right - rcClient.left != prcBounds->right - prcBounds->left) && 
		(rcClient.bottom - rcClient.top != prcBounds->bottom - prcBounds->top)) 
	{

        iMapMode = SetMapMode(hdc, MM_ANISOTROPIC);
        SetWindowExtEx(hdc, rcClient.right, rcClient.bottom, &sWOrg);
        SetViewportExtEx(hdc, prcBounds->right - prcBounds->left, prcBounds->bottom - prcBounds->top, &sVOrg);
    }

    SetWindowOrgEx(hdc, 0, 0, &ptWOrg);
    SetViewportOrgEx(hdc, prcBounds->left, prcBounds->top, &ptVOrg);

#if STRICT
    CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, hwnd, (TRUE) ? WM_PRINT : WM_PAINT, (WPARAM)hdc, (LPARAM)(TRUE ? PRF_CHILDREN | PRF_CLIENT : 0));
#else
    CallWindowProc((FARPROC)SUBCLASSWNDPROCOFCONTROL(0), hwnd, (TRUE) ? WM_PRINT : WM_PAINT, (WPARAM)hdc, (LPARAM)(TRUE ? PRF_CHILDREN | PRF_CLIENT : 0));
#endif // STRICT

    return S_OK;
}
#endif
//{IViewObject Support Code}

// IViewObject2 members

STDMETHODIMP CCustomizeListbox::GetExtent( DWORD dwDrawAspect, LONG lindex, DVTARGETDEVICE *ptd, LPSIZEL lpsizel)
{
	TRACE(1, "IViewObject2::GetExtent\n");
	return GetExtent(dwDrawAspect, lpsizel); // use IOleObject impl.
}



// IPersist members

STDMETHODIMP CCustomizeListbox::GetClassID( LPCLSID lpClassID)
{
	*lpClassID=*(((UNKNOWNOBJECTDESC *)(objectDef))->rclsid);
	return NOERROR;
}

// IPersistPropertyBag members

STDMETHODIMP CCustomizeListbox::Load( IPropertyBag *pPropBag, IErrorLog *pErrorLog)
{
	return PersistText(pPropBag,pErrorLog,FALSE);
}
STDMETHODIMP CCustomizeListbox::Save( IPropertyBag *pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties)
{
	HRESULT hr;
	hr=PersistText(pPropBag,NULL,TRUE);
	if (FAILED(hr))
		return hr;

	if (fClearDirty)
        m_fDirty = FALSE;

    if (m_pOleAdviseHolder)
        m_pOleAdviseHolder->SendOnSave();

	return NOERROR;
}


//{IPersistPropertyBag Support Code}
HRESULT CCustomizeListbox::PersistText(IPropertyBag *pPropBag, IErrorLog *pErrorLog,BOOL save)
{
	VARIANT_BOOL vbSave = save ? VARIANT_TRUE : VARIANT_FALSE;
	HRESULT hr;
	SIZEL slHiMetric;

    // Persist Standard properties
	if (save)
	    PixelToHiMetric(&m_Size, &slHiMetric);
	hr=PersistBagI4(pPropBag,L"_ExtentX",pErrorLog,(LONG *)&slHiMetric.cx,vbSave);
	if (FAILED(hr))
		return hr;
	hr=PersistBagI4(pPropBag,L"_ExtentY",pErrorLog,(LONG *)&slHiMetric.cy,vbSave);
	if (FAILED(hr))
		return hr;
	if (!save)
		HiMetricToPixel(&slHiMetric,&m_Size);

	//{PERSISTBAG}
	//{PERSISTBAG}
	hr=PersistBagFont(pPropBag,L"Font",pErrorLog,&m_Font,&GetGlobals()._fdDefault,vbSave);
	if (FAILED(hr))
		return hr;

	hr=PersistBagI4(pPropBag,L"Type",pErrorLog,(long*)&m_clbType,vbSave);
	if (FAILED(hr))
		return hr;

	hr=PersistBagBSTR(pPropBag,L"Category",pErrorLog,&m_bstrCategory,vbSave);
	if (FAILED(hr))
		return hr;

	return NOERROR;
}
//{IPersistPropertyBag Support Code}

// IPersistStreamInit members

STDMETHODIMP CCustomizeListbox::IsDirty()
{
    return (m_fDirty) ? S_OK : S_FALSE;
}
STDMETHODIMP CCustomizeListbox::Load( LPSTREAM pStm)
{
    HRESULT hr;
    // first thing to do is read in standard properties the user don't
    // persist themselves.
    hr = LoadStandardState(pStm);
    RETURN_ON_FAILURE(hr);
    // load in the user properties.  this method is one they -have- to implement
    // themselves.
    hr = PersistBinaryState(pStm,FALSE);
    return hr;
}
STDMETHODIMP CCustomizeListbox::Save( LPSTREAM pStm, BOOL fClearDirty)
{
    HRESULT hr;
    // use our helper routine that we share with the IStorage persistence
    // code.
    hr = m_SaveToStream(pStm);
    RETURN_ON_FAILURE(hr);
    // clear out dirty flag [if appropriate] and notify that we're done
    // with save.
    if (fClearDirty)
        m_fDirty = FALSE;
    if (m_pOleAdviseHolder)
        m_pOleAdviseHolder->SendOnSave();
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::GetSizeMax( ULARGE_INTEGER *pCbSize)
{
    return E_NOTIMPL;
}
STDMETHODIMP CCustomizeListbox::InitNew()
{
    BOOL f;
    f = InitializeNewState();
    return (f) ? S_OK : E_FAIL;
}


//{IPersistStreamInit Support Code}
#define STREAMHDR_SIGNATURE 0x55441234

typedef struct tagSTREAMHDR {

    DWORD  dwSignature;     // Signature.
    size_t cbWritten;       // Number of bytes written

} STREAMHDR;

HRESULT CCustomizeListbox::LoadStandardState(IStream *pStream)
{
    STREAMHDR stmhdr;
    HRESULT hr;
    SIZEL   slHiMetric;

    hr = pStream->Read(&stmhdr, sizeof(STREAMHDR), NULL);
    RETURN_ON_FAILURE(hr);

    if (stmhdr.dwSignature != STREAMHDR_SIGNATURE)
        return E_UNEXPECTED;

    if (stmhdr.cbWritten != sizeof(m_Size))
        return E_UNEXPECTED;

    hr = pStream->Read(&slHiMetric, sizeof(slHiMetric), NULL);
    RETURN_ON_FAILURE(hr);

    HiMetricToPixel(&slHiMetric, &m_Size);
    return S_OK;
}

HRESULT CCustomizeListbox::PersistBinaryState(IStream *pStream,BOOL save)
{
	//{PERSIST BINARY}
	//{PERSIST BINARY}
	VARIANT_BOOL vbSave = save ? VARIANT_TRUE : VARIANT_FALSE;
	HRESULT hResult = PersistFont(pStream, &m_Font, &GetGlobals()._fdDefault, vbSave);
	if (FAILED(hResult))
		return hResult;
	if (save)
	{
		hResult = pStream->Write(&m_clbType, sizeof(m_clbType), 0);
		if (FAILED(hResult))
			return hResult;
		
		hResult = StWriteBSTR(pStream,m_bstrCategory);
		if (FAILED(hResult))
			return hResult;
	}
	else
	{
		hResult = pStream->Read(&m_clbType, sizeof(m_clbType), 0);
		if (FAILED(hResult))
			return hResult;

		hResult = StReadBSTR(pStream, m_bstrCategory);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}

HRESULT CCustomizeListbox::SaveStandardState(IStream *pStream)
{
    STREAMHDR streamhdr = { STREAMHDR_SIGNATURE, sizeof(SIZEL) };
    HRESULT hr;
    SIZEL   slHiMetric;

    hr = pStream->Write(&streamhdr, sizeof(STREAMHDR), NULL);
    RETURN_ON_FAILURE(hr);

    PixelToHiMetric(&m_Size, &slHiMetric);

    hr = pStream->Write(&slHiMetric, sizeof(slHiMetric), NULL);
    return hr;
}

//{IPersistStreamInit Support Code}

// IPersistStorage members

STDMETHODIMP CCustomizeListbox::InitNew( IStorage *pStg)
{
	TRACE(1, "IPersistStorage::InitNew\n");
	BOOL bInitSuccess;
    // call the overridable function to do this work
    bInitSuccess = InitializeNewState();
    return (bInitSuccess) ? S_OK : E_FAIL;
}
STDMETHODIMP CCustomizeListbox::Load( IStorage *pStg)
{
	TRACE(1, "IPersistStorage::Load\n");
	IStream *pStream;
    HRESULT  hr;
    // we're going to use IPersistStream::Load from the CONTENTS stream.
    hr = pStg->OpenStream(wszCtlSaveStream, 0, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);
    RETURN_ON_FAILURE(hr);
	#ifdef DEF_PERSISTSTREAM
    hr = Load(pStream);
	#endif
    pStream->Release();
    return hr;
}
STDMETHODIMP CCustomizeListbox::Save( IStorage *pStgSave, BOOL fSameAsLoad)
{
	TRACE(1, "IPersistStorage::Save\n");
    IStream *pStream;
    HRESULT  hr;
    // we're just going to save out to the CONTENTES stream.
    hr = pStgSave->CreateStream(wszCtlSaveStream, STGM_WRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE,
                                0, 0, &pStream);
    RETURN_ON_FAILURE(hr);
    hr = m_SaveToStream(pStream);
    m_fSaveSucceeded = (FAILED(hr)) ? FALSE : TRUE;
    pStream->Release();
    return hr;
}
STDMETHODIMP CCustomizeListbox::SaveCompleted( IStorage *pStgNew)
{
	TRACE(1, "IPersistStorage::SaveCompleted\n");
    // if our save succeeded, then we can do our post save work.
    if (m_fSaveSucceeded) 
	{
        m_fDirty = FALSE;
        if (m_pOleAdviseHolder)
            m_pOleAdviseHolder->SendOnSave();
    }
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::HandsOffStorage()
{
	TRACE(1, "IPersistStorage::HandsOffStorage\n");
	return S_OK; // only used when caching IStorage pointer
}


//{IPersistStorage Support Code}
BOOL CCustomizeListbox::InitializeNewState()
{
	//{PERSIST INIT}
	//{PERSIST INIT}
	return TRUE;
}

HRESULT CCustomizeListbox::m_SaveToStream(IStream *pStream)
{
	HRESULT hr;
	hr = SaveStandardState(pStream);
    RETURN_ON_FAILURE(hr);
	hr = PersistBinaryState(pStream,TRUE);
    return hr;
}
//{IPersistStorage Support Code}

// IConnectionPointContainer members

STDMETHODIMP CCustomizeListbox::EnumConnectionPoints( IEnumConnectionPoints **ppEnum)
{
    TRACE(1, "Enum connection points\n");
	IConnectionPoint **rgConnectionPoints;

    // HeapAlloc an array of connection points [since our standard enum
    // assumes this and HeapFree's it later ]
    //
    rgConnectionPoints = (IConnectionPoint **)HeapAlloc(g_hHeap, 0, sizeof(IConnectionPoint *) * 2);
    RETURN_ON_NULLALLOC(rgConnectionPoints);

    // we support the event interface for this dude as well as IPropertyNotifySink
    //
    rgConnectionPoints[0] = (IConnectionPoint*)&m_cpEvents;
    rgConnectionPoints[1] = (IConnectionPoint*)&m_cpPropNotify;

    *ppEnum = (IEnumConnectionPoints *)(IEnumX *) new CEnumX(IID_IEnumConnectionPoints,
                                2, sizeof(IConnectionPoint *), (void *)rgConnectionPoints,
                                CopyAndAddRefObject);
    if (!*ppEnum) 
	{
        HeapFree(g_hHeap, 0, rgConnectionPoints);
        return E_OUTOFMEMORY;
    }

    return S_OK;
}
STDMETHODIMP CCustomizeListbox::FindConnectionPoint( REFIID riid,  IConnectionPoint **ppCP)
{
	TRACE(1, "FindConnectionPoint\n");
    // we support the event interface, and IDispatch for it, and we also
    // support IPropertyNotifySink.
    //
    if ((EVENTREFIIDOFCONTROL(objectDef)!=NULL && DO_GUIDS_MATCH(riid, (*EVENTREFIIDOFCONTROL(objectDef)))) || DO_GUIDS_MATCH(riid, IID_IDispatch))
        *ppCP = (IConnectionPoint*)m_cpEvents;
    else if (DO_GUIDS_MATCH(riid, IID_IPropertyNotifySink))
        *ppCP = (IConnectionPoint*)m_cpPropNotify;
    else
        return E_NOINTERFACE;

    (*ppCP)->AddRef();
    return S_OK;
}


// ISpecifyPropertyPages members

STDMETHODIMP CCustomizeListbox::GetPages( CAUUID *pPages)
{
    const GUID **pElems;
    void *pv;
    WORD  x;

	pPages->cElems=PROPPAGECOUNT(10);
	if (pPages->cElems==0)
	{
		pPages->pElems=NULL;
		return S_OK;
	}

	pv = CoTaskMemAlloc(sizeof(GUID) * (pPages->cElems));
    if (!pv)
		return E_OUTOFMEMORY;

    pPages->pElems = (GUID *)pv;

    pElems = PROPPAGES(10);
    for (x = 0; x < pPages->cElems; x++)
        pPages->pElems[x] = *(pElems[x]);

	return S_OK;
}

// IProvideClassInfo members

STDMETHODIMP CCustomizeListbox::GetClassInfo( ITypeInfo **ppTI)
{
	ITypeLib *pTypeLib;
    HRESULT hr;

	*ppTI = NULL;
    hr = LoadRegTypeLib(LIBID_PROJECT, 
		((CONTROLOBJECTDESC *)objectDef)->automationDesc.versionMajor,
		((CONTROLOBJECTDESC *)objectDef)->automationDesc.versionMinor,
    	LANGIDFROMLCID(g_lcidLocale),
		&pTypeLib);

    RETURN_ON_FAILURE(hr);

    // got the typelib.  get typeinfo for our coclass.
    hr = pTypeLib->GetTypeInfoOfGuid((REFIID)CLSIDOFOBJECT(10), ppTI);
    pTypeLib->Release();
    RETURN_ON_FAILURE(hr);

    return S_OK;
}
//{EVENT WRAPPERS}
void CCustomizeListbox::FireSelChangeBand(Band *pBand)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_SELCHANGEBAND,EVENT_PARAM(VTS_DISPATCH VTS_NONE),pBand);
}

void CCustomizeListbox::FireSelChangeCategory(BSTR strCategory)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_SELCHANGECATEGORY,EVENT_PARAM(VTS_WBSTR VTS_NONE),strCategory);
}

void CCustomizeListbox::FireSelChangeTool(Tool *pTool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_SELCHANGETOOL,EVENT_PARAM(VTS_DISPATCH VTS_NONE),pTool);
}

void CCustomizeListbox::FireError( short Number, ReturnString *Description, long Scode, BSTR Source, BSTR HelpFile, long HelpContext, ReturnBool *CancelDisplay)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_ERROR,EVENT_PARAM(VTS_I2 VTS_DISPATCH VTS_I4 VTS_WBSTR VTS_WBSTR VTS_I4 VTS_DISPATCH VTS_NONE),Number,Description,Scode,Source,HelpFile,HelpContext,CancelDisplay);
}

//{EVENT WRAPPERS}
//
// Utility Functions
//

void CCustomizeListbox::OnCommand(int nNotifyMessage)
{
	try
	{
		switch(nNotifyMessage)
		{
		case LBN_SELCHANGE:
			{
				int nIndex = GetCurSel();
				if (LB_ERR == nIndex)
					break;

				switch (m_clbType)
				{
				case ddCTBands:
					{
						CLItemData* pData = (CLItemData*)GetItemData(nIndex);
						assert(pData);
						if (pData && LB_ERR != (int)pData)
							FireSelChangeBand((Band*)pData->m_pBand);
					}
					break;

				case ddCTCategories:
					{
						int nIndex = GetCurSel();
						CatEntry* pCatEntry = (CatEntry*)GetItemData(nIndex);
						assert(pCatEntry);
						if (pCatEntry && LB_ERR != (int)pCatEntry)
							FireSelChangeCategory(pCatEntry->Category());
					}
					break;

				case ddCTTools:
					{
						int nIndex = GetCurSel();
						ITool* pTool = (ITool*)GetItemData(nIndex);
						assert(pTool && LB_ERR != (int)pTool);
						if (pTool && LB_ERR != (int)pTool)
							FireSelChangeTool((Tool*)pTool);
					}
					break;
				}
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// DrawItem
//

void CCustomizeListbox::DrawItem(LPDRAWITEMSTRUCT pDrawItemStruct)
{
	switch (m_clbType)
	{
	case ddCTBands:
		DrawCheckItem(pDrawItemStruct);
		break;

	case ddCTTools:
		DrawToolItem(pDrawItemStruct);
		break;
	}
}

//
// DrawCheckItem
//

void CCustomizeListbox::DrawCheckItem(LPDRAWITEMSTRUCT pDrawItemStruct)
{
	HDC hDC = pDrawItemStruct->hDC;
	HFONT oldFont = SelectFont(hDC, m_Font.GetFontHandle());
	BOOL bSelected;
	BSTR bstr;
	CLItemData* pItemData = (CLItemData*)pDrawItemStruct->itemData;

	if ((pDrawItemStruct->itemID >= 0) &&
		(pDrawItemStruct->itemAction & (ODA_DRAWENTIRE | ODA_SELECT)))
	{
		BOOL bDisabled = !IsWindowEnabled(m_hWnd);

		COLORREF crBkColorOld = SetBkColor(hDC, GetSysColor(COLOR_WINDOW));
		COLORREF crTextColorOld = SetTextColor(hDC, 0);

		if (!bDisabled && ((pDrawItemStruct->itemState & ODS_SELECTED) != 0))
		{
			SetTextColor(hDC, GetSysColor(COLOR_HIGHLIGHTTEXT));
			SetBkColor(hDC, GetSysColor(COLOR_HIGHLIGHT));
		}

		if (pItemData)
		{
			bstr = pItemData->m_pBand->m_bstrCaption;
			bSelected = pItemData->m_bChecked;
		}

		CRect rcText = pDrawItemStruct->rcItem;
		rcText.left += eCheckboxWidth + 4;
		MAKE_TCHARPTR_FROMWIDE(szText,bstr);

		ExtTextOut(hDC,
				   rcText.left,
				   rcText.top + (m_nItemHeight - m_nTextHeight + 1) / 2,
				   ETO_OPAQUE, 
				   &(pDrawItemStruct->rcItem), 
				   szText,
				   lstrlen(szText), 
				   0);

		CRect rc = pDrawItemStruct->rcItem;
		rc.left += 1;
		rc.right = rc.left + eCheckboxWidth;
		rc.top += (rc.Height() - eCheckboxWidth) / 2;
		rc.bottom = rc.top + eCheckboxWidth;

		DrawFrameControl(hDC, 
						 &rc,
						 DFC_BUTTON,
						 bSelected ? DFCS_BUTTONCHECK | DFCS_CHECKED : DFCS_BUTTONCHECK);

		SetTextColor(hDC, crTextColorOld);
		SetBkColor(hDC, crBkColorOld);
	}

	if (0 != (pDrawItemStruct->itemState & ODS_FOCUS))
		DrawFocusRect(hDC, &(pDrawItemStruct->rcItem));

	SelectFont(hDC, oldFont);
}

//
// DrawToolItem
//

void CCustomizeListbox::DrawToolItem(LPDRAWITEMSTRUCT pDrawItemStruct)
{
	HPALETTE hPal = NULL, hPalOld;
	HRESULT  hResult;
	CRect    rc;
	BOOL     bSelected = pDrawItemStruct->itemState & ODS_SELECTED != 0;
	BSTR	 bstr = NULL;
	HDC		 hDCWin = pDrawItemStruct->hDC;
	int		 nLineHeight;

	VARIANT_BOOL vbXPLook;
	m_pActiveBar->get_XPLook(&vbXPLook);
	
	// TO DO I need to use the palette here

	HDC hDC = hDCWin;
	CRect rcOutput = pDrawItemStruct->rcItem;

	ITool* pTool = (ITool*)(pDrawItemStruct->itemData);
	if (NULL == pTool)
		return;

	if (hPal)
	{
		hPalOld = SelectPalette(hDCWin, hPal, FALSE);
		RealizePalette(hDCWin);
	}

	if (pDrawItemStruct->itemID >= 0)
		//&&
		//(pDrawItemStruct->itemAction & (ODA_DRAWENTIRE | ODA_SELECT)))
	{
		if (m_pFF)
		{
			hDC = m_pFF->RequestDC(hDCWin,
								   rcOutput.Width(),
								   rcOutput.Height());
			if (NULL == hDC)
				hDC = hDCWin;
			else
				rcOutput.Offset(-rcOutput.left, -rcOutput.top);
		}

		HFONT hFontOld = SelectFont(hDC, m_Font.GetFontHandle());

		BOOL bDisabled = !IsWindowEnabled(m_hWnd);

		BOOL vbDrawHighlighted = VARIANT_FALSE;
		if (!bDisabled && bSelected)
			vbDrawHighlighted = VARIANT_TRUE;

		ToolTypes ttTool;
		hResult = pTool->get_ControlType(&ttTool);
		switch (ttTool)
		{
		case ddTTCombobox:
		case ddTTEdit:
			{

				FillRect(hDC,&rcOutput,(HBRUSH)(1+COLOR_BTNFACE));
				pTool->ExtDraw((OLE_HANDLE)hDC,
							   rcOutput.left,
							   rcOutput.top,
							   rcOutput.right,
							   rcOutput.bottom,
							   ddBTPopup,
							   vbDrawHighlighted);
			}
			break;

		default:
			{
				pTool->get_Caption(&bstr);

				// remove \n
				if (bstr && *bstr)
				{
					WCHAR* ptr = (WCHAR*)bstr;
					while (*ptr)
					{
						if (*ptr == 0xd || *ptr == 0xa)
						{
							*ptr = NULL;
							break;
						}
						++ptr;
					}
				}
				
				CRect rcText;
				CRect rcTool = rcOutput;
				rcTool.right = rcTool.left + 26;
				
				FillRect(hDC,&rcTool,(HBRUSH)(1+COLOR_BTNFACE));

				rcTool.Inflate(-1,-1);
				if (bSelected && pDrawItemStruct->itemState & ODS_FOCUS && VARIANT_FALSE == vbXPLook)
					DrawEdge(hDC, &rcTool, EDGE_RAISED, BF_RECT);

				rcTool.Inflate(-1, -1);
				
				nLineHeight = rcOutput.Height();	
				rcText = rcOutput;
				rcText.left += 26;

				SetBkMode(hDC,TRANSPARENT);

				if (bSelected && pDrawItemStruct->itemState & ODS_FOCUS)
				{
					if (VARIANT_FALSE == vbXPLook)
					{
						FillRect(hDC,&rcText,(HBRUSH)(1+COLOR_BTNFACE));
						rcText.Inflate(0, -1);
						FillRect(hDC,&rcText,(HBRUSH)(1+COLOR_HIGHLIGHT));
						SetTextColor(hDC,GetSysColor(COLOR_HIGHLIGHTTEXT));
					}
					else
					{
						CRect rcTemp = rcOutput;
						rcTemp.Inflate(-2, -2);
						
						HPEN aPen = CreatePen(PS_SOLID, 1, m_crXPSelectedBorder);
						if (aPen)
						{
							HPEN penOld = SelectPen(hDC, aPen);

							HBRUSH aBrush = CreateSolidBrush(m_crXPSelected);
							if (aBrush)
							{
								HBRUSH brushOld = SelectBrush(hDC, aBrush);
								
								Rectangle(hDC, rcTemp.left, rcTemp.top, rcTemp.right, rcTemp.bottom);

								SelectBrush(hDC, brushOld);
								DeleteBrush(aBrush);
							}
							SelectPen(hDC, penOld);
							DeletePen(aPen);
						}
						SetTextColor(hDC,GetSysColor(COLOR_BTNTEXT));
					}
				}
				else
				{
					if (VARIANT_FALSE == vbXPLook)
						FillRect(hDC,&rcText,(HBRUSH)(1+COLOR_BTNFACE));
					else
					{
						HBRUSH aBrush = CreateSolidBrush(m_crXPMenuBackground);
						if (aBrush)
						{
							FillRect(hDC, &rcText, aBrush);
							DeleteBrush(aBrush);
						}
					}
					SetTextColor(hDC, GetSysColor(COLOR_BTNTEXT));
				}

				LPDISPATCH pToolDispatch;
				hResult = pTool->get_Custom(&pToolDispatch);
				if (FAILED(hResult) || NULL == pToolDispatch)
				{
					if (ddTTSeparator == ttTool)
					{
						rcTool.left = rcTool.left + ((rcTool.Width() - CTool::eSeparatorThickness) / 2);
						rcTool.right = rcTool.left + CTool::eSeparatorThickness;
						rcTool.Inflate(-2, -2);
						DrawEdge(hDC, &rcTool, BDR_SUNKENOUTER, BF_RECT);
					}
					else
					{
						pTool->DrawPict((OLE_HANDLE)hDC,
										rcTool.left + (rcTool.Width()-16)/2,
										rcTool.top + (rcTool.Height()-16)/2,
										16,
										16,
										VARIANT_TRUE);
					}
				}
				else if (pToolDispatch)
				{
					CCustomTool custTool(pToolDispatch);
					custTool.SetHost((ICustomHost*)pTool);
					VARIANT_BOOL bSaveEnabled;
					pTool->get_Enabled(&bSaveEnabled);
					pTool->put_Enabled(VARIANT_TRUE);
					custTool.OnDraw((OLE_HANDLE)hDC,
									rcTool.left + (rcTool.Width()-16)/2,
									rcTool.top + (rcTool.Height()-16)/2,
									16,
									16);
					pTool->put_Enabled(bSaveEnabled);
					pToolDispatch->Release();
				}

				rcText.left+=2;
				MAKE_TCHARPTR_FROMWIDE(szText,bstr);
				DrawText(hDC,szText,lstrlen(szText),&rcText,DT_LEFT|DT_VCENTER|DT_SINGLELINE);
				SysFreeString(bstr);

				BSTR bstrSubBand = NULL;
				hResult = pTool->get_SubBand(&bstrSubBand);
				if (SUCCEEDED(hResult))
				{
					if (bstrSubBand && *bstrSubBand)
					{
						HDC hMemDC = CreateCompatibleDC(hDC);
						HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapSubMenu());
						SetTextColor(hDC, RGB(0,0,0));	
						SetBkColor(hDC, RGB(255,255,255));
						BitBlt(hDC, 
							   rcText.right-6, 
							   ((rcText.bottom+rcText.top)/2)-3,
							   4,
							   7,
							   hMemDC,
							   0,
							   0,
							   SRCAND);
			
						SetBkColor(hDC, RGB(0,0,0));
						if (VARIANT_TRUE == vbDrawHighlighted && VARIANT_FALSE == vbXPLook)
							SetTextColor(hDC, GetSysColor(COLOR_HIGHLIGHTTEXT));
						else
							SetTextColor(hDC, GetSysColor(COLOR_BTNTEXT));	
						
						BitBlt(hDC, 
							   rcText.right-6,
							   ((rcText.bottom+rcText.top)/2)-3,
							   4,
							   7,
							   hMemDC,
							   0,
							   0,
							   SRCPAINT);
			
						SelectBitmap(hMemDC, hBitmapOld);
						DeleteDC(hMemDC);
					}
					SysFreeString(bstrSubBand);
				}
			}
			break;
		}
		
		SelectFont(hDC, hFontOld);

		if (hDC != hDCWin)
		{
			m_pFF->Paint(hDCWin,
						 pDrawItemStruct->rcItem.left,
						 pDrawItemStruct->rcItem.top);
		}
	}
	
	if (pDrawItemStruct->itemState & ODS_FOCUS && VARIANT_FALSE == vbXPLook)
	{
		CRect rc = pDrawItemStruct->rcItem;
		rc.Inflate(-3, -3);
		DrawFocusRect(hDCWin, &rc);
		TRACE(5, "Focus\n");
	}

	if (hPal)
		SelectPalette(hDCWin, hPalOld, FALSE);
}

//
// OnLButtonDown
//

void CCustomizeListbox::OnLButtonDown(UINT nFlags, POINT pt) 
{
	try
	{
		switch (m_clbType)
		{
		case ddCTBands:
			if (pt.x < eCheckboxWidth)
			{
				CRect rcItem;
				int nIndex = GetCurSel();
				if (LB_ERR == nIndex)
					return;

				GetItemRect(nIndex, &rcItem);
				rcItem.left++;
				rcItem.right = rcItem.left + eCheckboxWidth ;
				rcItem.top += (m_nItemHeight - eCheckboxHeight) / 2;
				rcItem.bottom = rcItem.top + eCheckboxHeight;
				if (!PtInRect(&rcItem, pt))
					return;

				CLItemData* pData;
				BOOL bCtrlPressed = (0x8000 & GetKeyState(VK_CONTROL));
				if (bCtrlPressed)
				{
					BOOL bDoCheck = ((CLItemData*)GetItemData(nIndex))->m_bChecked ^ 1;
					int nCount = GetCount();
					for (nIndex = 0; nIndex < nCount; nIndex++)
					{
						pData = (CLItemData*)GetItemData(nIndex);
						if (pData && ddBTMenuBar != ((CBand*)pData->m_pBand)->bpV1.m_btBands)
						{
							pData->m_bChecked = bDoCheck;
						
							IBand* pBand = (IBand*)pData->m_pBand;
							if (pBand)
								pBand->put_Visible(pData->m_bChecked ? VARIANT_TRUE : VARIANT_FALSE);

							((CBar*)m_pActiveBar)->m_vbCustomizeModified = VARIANT_TRUE;
						}
						else
							MessageBeep(0xFFFFFFFF);
					}
					ViewChanged();
					InvalidateRect(m_hWnd, NULL, FALSE);
				}
				else
				{
					pData = (CLItemData*)GetItemData(nIndex);
					if (pData && ddBTMenuBar != ((CBand*)pData->m_pBand)->bpV1.m_btBands)
					{
						pData->m_bChecked ^= 1;
						
						IBand* pBand = (IBand*)pData->m_pBand;
						if (pBand)
							pBand->put_Visible(pData->m_bChecked ? VARIANT_TRUE : VARIANT_FALSE);

						((CBar*)m_pActiveBar)->m_vbCustomizeModified = VARIANT_TRUE;
						ViewChanged();
						InvalidateRect(m_hWnd, NULL, FALSE);
					}
					else
						MessageBeep(0xFFFFFFFF);
				}
				((CBar*)m_pActiveBar)->RecalcLayout();
				m_bDirty = TRUE;
			}
			break;

		case ddCTTools:
			{
				if (VARIANT_FALSE == m_vbToolDragDrop)
					return;
				
				((CBar*)m_pActiveBar)->m_diCustSelection.pTool = NULL;
				if (((CBar*)m_pActiveBar)->m_diCustSelection.pBand)
					((CBar*)m_pActiveBar)->m_diCustSelection.pBand->Refresh();

				POINT ptDragStart = pt;
				ClientToScreen(m_hWnd, &ptDragStart);

				BOOL bProcessing = TRUE;
				SetCapture(m_hWnd);
				MSG msg;
				DWORD nDragStartTick = GetTickCount();
				while (GetCapture() == m_hWnd && bProcessing)
				{
					GetMessage(&msg, NULL, 0, 0);
					switch (msg.message)
					{
					case WM_KEYDOWN:
						if (VK_ESCAPE == msg.wParam)
							bProcessing = FALSE;
						break;

					case WM_LBUTTONUP:
						bProcessing = FALSE;
						break;

					case WM_MOUSEMOVE:
						{
							try
							{
								DWORD nDeltaTime = GetTickCount() - nDragStartTick;
								long nDeltaX = abs(msg.pt.x - ptDragStart.x);
								long nDeltaY = abs(msg.pt.y - ptDragStart.y);
								if (nDeltaTime < GetGlobals().m_nDragDelay && (nDeltaX < GetGlobals().m_nDragDist || nDeltaY < GetGlobals().m_nDragDist))
									break;

								int nSelCount = SendMessage(m_hWnd, LB_GETSELCOUNT, 0, 0);
								if (0 == nSelCount)
									return;

								if (LB_ERR == nSelCount)
									nSelCount = 1;

								int* pnSelArray = new int[nSelCount];
								if (NULL == pnSelArray)
									return;

								SendMessage(m_hWnd, LB_GETSELITEMS, nSelCount, (LPARAM)pnSelArray);

								TypedArray<ITool*> aToolArray;
								HRESULT hResult;
								CTool* pTool;
								for  (int nSelection = 0; nSelection < nSelCount; nSelection++)
								{
									pTool = (CTool*)SendMessage(m_hWnd, LB_GETITEMDATA, pnSelArray[nSelection], 0);
									assert(pTool);
									if (pTool)
									{
										hResult = aToolArray.Add(pTool);
										if (FAILED(hResult))
											continue;
									}
								}
								delete [] pnSelArray;

								CToolDropSource* pSource = new CToolDropSource;
 								assert(pSource);
								if (NULL == pSource)
									return;

								CToolDataObject* pDataObject = new CToolDataObject(CToolDataObject::eDesignerDragDropId, 
																				   m_pActiveBar, 
																				   aToolArray);
								assert(pDataObject);
								if (NULL == pDataObject)
								{
									pSource->Release();
									return;
								}
								((CBar*)m_pActiveBar)->m_pCustomizeDragLock = (CTool*)aToolArray.GetAt(0); 


								DWORD dwEffect;
								hResult = GetGlobals().m_pDragDropMgr->DoDragDrop((LPUNKNOWN)pDataObject, 
																				  (LPUNKNOWN)pSource, 
																				  DROPEFFECT_COPY, 
																				  &dwEffect);
								((CBar*)m_pActiveBar)->m_pCustomizeDragLock = NULL; 
								if (DRAGDROP_S_DROP != hResult)
								{
									// Simulate simple LButtonDown/Up
									CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, m_hWnd, WM_LBUTTONDOWN, nFlags, MAKELPARAM(pt.x, pt.y));
									CallWindowProc((WNDPROC)((CONTROLOBJECTDESC*)objectDef)->pfnSubClass, m_hWnd, WM_LBUTTONUP, nFlags, MAKELPARAM(pt.x, pt.y));
								}
								else
								{
									((CBar*)m_pActiveBar)->m_vbCustomizeModified = VARIANT_TRUE;
									((CBar*)m_pActiveBar)->RecalcLayout();
								}
								pDataObject->Release();
								pSource->Release();
							}
							CATCH
							{
								assert(FALSE);
								REPORTEXCEPTION(__FILE__, __LINE__)
							}			
						}
						break;

					default:
						// Just dispatch rest of the messages
						DispatchMessage(&msg);
						TRACE(1, _T("Default\n"));
						break;
					}
				}
				if (GetCapture() == m_hWnd)
					ReleaseCapture();
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnDeleteItem
//

void CCustomizeListbox::OnDeleteItem(LPDELETEITEMSTRUCT pDeleteStruct)
{
	switch (m_clbType)
	{
	case ddCTBands:
		{
			CLItemData* pData = (CLItemData*)pDeleteStruct->itemData;
			delete pData;
		}
		break;

	case ddCTCategories:
		break;

	case ddCTTools:
		break;
	}
}

//
// EraseBackground
//

LRESULT CCustomizeListbox::EraseBackground(HDC hDC)
{
	if (ddCTTools == m_clbType)
	{
		CRect rcClient;
		GetClientRect(m_hWnd, &rcClient);
		FillRect(hDC, &rcClient, (HBRUSH)(1+COLOR_BTNFACE));
		return TRUE;
	}
	else
		return FALSE;
}

//
// RecreateControlWindow
//

void CCustomizeListbox::RecreateControlWindow()
{
	try
	{		

		if (m_fInPlaceActive)
		{
			BOOL bUIActive = m_fUIActive;
			InPlaceDeactivate();
            InPlaceActivate(OLEIVERB_PRIMARY);
			if (!bUIActive)
				UIDeactivate();
		}
		else if (m_hWndParent)
		{
			DestroyWindow(m_hWnd);
			CRect rcClient;
			GetClientRect(m_hWndParent, &rcClient);
			CreateInPlaceWindow(rcClient.Width(), 
								rcClient.Height(), 
								TRUE);
		}
		else
		{
			HWND hWndParent = GetParkingWindow();
			if (hWndParent)
			{
				DestroyWindow(m_hWnd);
				CreateInPlaceWindow(m_Size.cx, m_Size.cy, TRUE);
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CreateCategoryManager
//
// m_pCategoryMgr is checked everywhere function does not return an error code
//

BOOL CCustomizeListbox::CreateCategoryManager()
{
	try
	{
		if (NULL == m_pActiveBar)
			return FALSE;

		if (NULL == m_pCategoryMgr)
			m_pCategoryMgr = new CategoryMgr;

		assert(m_pCategoryMgr);
		if (NULL == m_pCategoryMgr)
			return FALSE;

		m_pCategoryMgr->ActiveBar(m_pActiveBar);
		m_pCategoryMgr->Init();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

//
// FillBands
//

void CCustomizeListbox::FillBands()
{
	try
	{
		if (NULL == m_hWnd)
			return;

		SendMessage(m_hWnd, LB_RESETCONTENT, 0, 0);
		IBands* pBands;
		HRESULT hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
		if (FAILED(hResult))
		{
			// TO DO Log event
			return;
		}

		short nCount;
		VARIANT vIndex;
		vIndex.vt = VT_I2;
		IBand* pBand;
		VARIANT_BOOL vbVisible;
		hResult = pBands->Count(&nCount);

		for (vIndex.iVal = 0; vIndex.iVal < nCount; vIndex.iVal++)
		{
			hResult = pBands->Item(&vIndex, (Band**)&pBand);
			if (FAILED(hResult))
			{
				pBand->Release();
				continue;
			}

			switch (((CBand*)pBand)->bpV1.m_btBands)
			{
			case ddBTChildMenuBar:
			case ddBTStatusBar:
			case ddBTPopup:
				pBand->Release();
				continue;
			}
		
			if (ddDAPopup == ((CBand*)pBand)->bpV1.m_daDockingArea)
			{
				pBand->Release();
				continue;
			}

			if (VARIANT_TRUE == ((CBand*)pBand)->bpV1.m_vbDetached)
			{
				pBand->Release();
				continue;
			}
		
			if (!(((CBand*)pBand)->bpV1.m_dwFlags & ddBFHide))
			{
				pBand->Release();
				continue;
			}

			CLItemData* pData = new CLItemData;
			assert(pData);
			if (NULL == pData)
			{
				// TO DO Log event
				pBand->Release();
				continue;
			}

			hResult = pBand->get_Visible(&vbVisible);
			if (FAILED(hResult))
			{
				delete pData;
				pBand->Release();
				continue;
			}

			pData->m_bChecked = vbVisible == VARIANT_TRUE ? TRUE : FALSE;
			pData->m_pBand = (CBand*)pBand;

			int nIndex = AddString((LPCTSTR)pData);
			if (LB_ERR == nIndex)
			{
				delete pData;
				pBand->Release();
			}
		}
		SendMessage(m_hWnd, LB_SETCURSEL, 0, 0);
		pBands->Release();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

STDMETHODIMP CCustomizeListbox::GetControlInfo( CONTROLINFO *pCI)
{
	// NOTE: control writers should override this routine if they want to
    // return accelerator information in their control.
    //
    pCI->hAccel = NULL;
    pCI->cAccel = NULL;

    return S_OK;
}
STDMETHODIMP CCustomizeListbox::OnMnemonic(MSG *pMsg)
{
	// default action. override to do interesting stuff
	return InPlaceActivate(OLEIVERB_UIACTIVATE);
}
STDMETHODIMP CCustomizeListbox::OnAmbientPropertyChange( DISPID dispID)
{
	if (dispID == DISPID_AMBIENT_USERMODE || dispID == DISPID_UNKNOWN)
        m_fModeFlagValid = FALSE;

    AmbientPropertyChanged(dispID);
    return S_OK;
}
STDMETHODIMP CCustomizeListbox::FreezeEvents( BOOL bFreeze)
{
   return S_OK;
}

//{IOleControl Support Code}
void CCustomizeListbox::AmbientPropertyChanged(DISPID dispID)
{
    // do nothing
}
//{IOleControl Support Code}

//{PROPERTY NOTIFICATION SINK}

CCustomizeListbox *CCustomizeListbox::CPropertyNotifySink::m_pMainUnknown() 
{
	return (CCustomizeListbox*)((LPBYTE)this - offsetof(CCustomizeListbox, m_xPropNotify));
}

STDMETHODIMP CCustomizeListbox::CPropertyNotifySink::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (riid==IID_IUnknown || riid==IID_IPropertyNotifySink)
	{
		*ppvObjOut=this;
		AddRef();
		return NOERROR;
	}
	return E_NOTIMPL;
}

STDMETHODIMP CCustomizeListbox::OnRequestEdit(DISPID dispID)
{
	return NOERROR;
}

STDMETHODIMP CCustomizeListbox::OnChanged(DISPID dispID)
{
	ViewChanged();
	return NOERROR;
}

//{PROPERTY NOTIFICATION SINK}

int CCustomizeListbox::GetCheck(int index)
{
	CLItemData* pData = (CLItemData*)GetItemData(index);
	if (pData)
		return pData->m_bChecked;
	return 0;
}

void CCustomizeListbox::SetCheck(int index, BOOL val)
{
	CLItemData* pData = (CLItemData*)GetItemData(index);
	if (pData)
		pData->m_bChecked = val;
}

//
// IToolDrop Interface
//

void DrawDropSign(HWND hWnd, int nIndex)
{
	if (-1 == nIndex)
		return;

	const int nDropSignHeight = 10;
	CRect rcBound;

	int nCount = SendMessage(hWnd, LB_GETCOUNT, 0, 0);
	if (nIndex >= nCount)
	{
		if (0 == nCount)
		{
			CRect rcClient;
			GetClientRect(hWnd, &rcClient);
			rcBound = rcClient;
			rcBound.bottom = rcBound.top;
		}
		else
		{
			SendMessage(hWnd, LB_GETITEMRECT, nIndex-1, (LPARAM)&rcBound);
			rcBound.top = rcBound.bottom;
		}
	}
	else
	{
		SendMessage(hWnd, LB_GETITEMRECT, nIndex, (LPARAM)&rcBound);
		rcBound.bottom = rcBound.top;
	}

	InflateRect(&rcBound, 0, nDropSignHeight/2);

	HDC hDC = GetDC(hWnd);
	if (hDC)
	{
		if (0 == nCount)
			OffsetRect(&rcBound, 0, nDropSignHeight/2);

		CRect rc = rcBound;
		rc.top = rc.Height()/2 - 1;
		rc.bottom = rc.top + 2;
		rc.left += 2;
		rc.right -= 2;

		HBRUSH hBrush = CreateHalftoneBrush(TRUE);
		if (hBrush)
		{
			HBRUSH hBrushOld = SelectBrush(hDC, hBrush);
			SetTextColor(hDC, GetSysColor(COLOR_BTNFACE));
			SetBkColor(hDC, GetSysColor(COLOR_BTNFACE));

			PatBlt(hDC, rc.left, rc.top, rc.Width(), rc.Height(), PATINVERT);

			rc = rcBound;
			rc.right = rc.left + 2;
			
			PatBlt(hDC, rc.left, rc.top, rc.Width(), rc.Height(), PATINVERT);

			rc = rcBound;
			rc.left = rc.right - 2;

			PatBlt(hDC, rc.left, rc.top, rc.Width(), rc.Height(), PATINVERT);

			SelectBrush(hDC, hBrushOld);
			BOOL bResult = DeleteBrush(hBrush);
			assert(bResult);
		}
		ReleaseDC(hWnd, hDC);
	}
}

void CCustomizeListbox::GetNewBandName(BSTR& bstr)
{
	VARIANT vIndex;
	vIndex.vt = VT_I2;
	IBand* pBand;
	short nCount;
	BSTR bstrBandName;
	int nMaxNameIndex = 0;
	int nNewIndex = 0;
	UINT nTestLen;

	LPTSTR szBaseName = LoadStringRes(IDS_CUSTOM_TOOLBARNAME);
	MAKE_WIDEPTR_FROMTCHAR(wBaseName, szBaseName);
	BSTR bstrBaseName = SysAllocString(wBaseName);
	if (NULL == bstrBaseName)
		return;

	IBands* pBands = NULL;
	HRESULT hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
	if (FAILED(hResult))
		goto Cleanup;

	nTestLen = lstrlen(szBaseName);
	hResult = pBands->Count(&nCount);
	if (FAILED(hResult))
		goto Cleanup;

	for (vIndex.iVal = 0; vIndex.iVal < nCount; vIndex.iVal++)
	{
		if (FAILED(pBands->Item(&vIndex, (Band**)&pBand)))
			continue;

		if (FAILED(pBand->get_Name(&bstrBandName)))
		{
			pBand->Release();
			continue;
		}

		if (NULL != bstrBandName && wcslen(bstrBandName) >= nTestLen && 0 == memcmp(bstrBandName, bstrBaseName, nTestLen * sizeof(WCHAR)))
		{
			MAKE_ANSIPTR_FROMWIDE(szNumber, bstrBandName + nTestLen);
			nNewIndex = Fatoi(szNumber);
			if (nNewIndex > nMaxNameIndex)
				nMaxNameIndex = nNewIndex;
		}
		SysFreeString(bstrBandName);
		pBand->Release();
	}
	++nMaxNameIndex;
	bstr = SysAllocStringByteLen(0,(nTestLen+4) * sizeof(WCHAR));
	if (bstr)
	{
		memcpy(bstr, bstrBaseName, nTestLen*sizeof(WCHAR));
		TCHAR szBuffer[20];
		wsprintf(szBuffer, _T("%d"), nMaxNameIndex);
		MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
		wcscpy(bstr + nTestLen, wBuffer);
	}
Cleanup:
	if (pBands)
		pBands->Release();
	SysFreeString(bstrBaseName);
}

STDMETHODIMP CCustomizeListbox::InterfaceSupportsErrorInfo( REFIID riid)
{
    if (riid == (REFIID)INTERFACEOFOBJECT(10))
        return S_OK;

    return S_FALSE;
}

STDMETHODIMP CCustomizeListbox::GetDisplayString( DISPID dispID, BSTR __RPC_FAR *pBstr)
{
	return S_FALSE;
}
STDMETHODIMP CCustomizeListbox::MapPropertyToPage( DISPID dispID, CLSID __RPC_FAR *pClsid)
{
	return PERPROP_E_NOPAGEAVAILABLE;
}
STDMETHODIMP CCustomizeListbox::GetPredefinedStrings(  DISPID dispID,  CALPOLESTR __RPC_FAR *pCaStringsOut,  CADWORD __RPC_FAR *pCaCookiesOut)
{
	return E_FAIL;
}
STDMETHODIMP CCustomizeListbox::GetPredefinedValue( DISPID dispID, DWORD dwCookie, VARIANT __RPC_FAR *pVarOut)
{
	return E_FAIL;
}

VARIANT_BOOL CCustomizeListbox::AmbientUserMode()
{
	BOOL bUserMode;
	if (!GetAmbientProperty(DISPID_AMBIENT_USERMODE, VT_I4, &bUserMode))
		bUserMode = TRUE;
	return bUserMode;
}	

//
// CCustomizeListboxHost
//

STDMETHODIMP CCustomizeListboxHost::QueryInterface(REFIID riid, void **ppvObj)
{
	if (riid==IID_IUnknown || riid==IID_IDispatch)
	{
		*ppvObj=this;
		AddRef();
		return NOERROR;
	}
	return E_NOTIMPL;
}

STDMETHODIMP_(ULONG) CCustomizeListboxHost::AddRef(void)
{
	return 1;
}

STDMETHODIMP_(ULONG) CCustomizeListboxHost::Release(void)
{
	return 1;
}

// IDispatch members
STDMETHODIMP CCustomizeListboxHost::GetTypeInfoCount( UINT *pctinfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CCustomizeListboxHost::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CCustomizeListboxHost::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	return E_NOTIMPL;
}

STDMETHODIMP CCustomizeListboxHost::Invoke( long dispid, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
{
	switch(dispid)
	{
	case DISPID_SELCHANGEBAND:
		if (VT_DISPATCH == pdispparams->rgvarg->vt)
			FireSelChangeBand((IBand*)pdispparams->rgvarg->pdispVal);
		break;

	case DISPID_SELCHANGETOOL:
		if (VT_DISPATCH == pdispparams->rgvarg->vt)
			FireSelChangeTool((ITool*)pdispparams->rgvarg->pdispVal);
		break;

	case DISPID_SELCHANGECATEGORY:
		if (VT_BSTR == pdispparams->rgvarg->vt)
			FireSelChangeCategory((BSTR)pdispparams->rgvarg->pdispVal);
		break;
	}
	return NOERROR;
}
//{IConnectionPointContainer Support Code}
STDMETHODIMP CCustomizeListbox::CConnectionPoint::QueryInterface
(
    REFIID riid,
    void **ppvObjOut
)
{
    if (DO_GUIDS_MATCH(riid, IID_IConnectionPoint) || DO_GUIDS_MATCH(riid, IID_IUnknown)) {
        *ppvObjOut = (IConnectionPoint *)this;
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG CCustomizeListbox::CConnectionPoint::AddRef
(
    void
)
{
	return ++m_refCount;
}

ULONG CCustomizeListbox::CConnectionPoint::Release
(
    void
)
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}

STDMETHODIMP CCustomizeListbox::CConnectionPoint::GetConnectionInterface
(
    IID *piid
)
{
	TRACE(1, "GetConnectionInterface\n");
    if (m_bType == SINK_TYPE_EVENT)
        *piid = *EVENTREFIIDOFCONTROL(10);
    else
        *piid = IID_IPropertyNotifySink;

    return S_OK;
}

STDMETHODIMP CCustomizeListbox::CConnectionPoint::GetConnectionPointContainer
(
    IConnectionPointContainer **ppCPC
)
{
    return m_pControl->ExternalQueryInterface(IID_IConnectionPointContainer, (void **)ppCPC);
}

STDMETHODIMP CCustomizeListbox::CConnectionPoint::Advise
(
    IUnknown *pUnk,
    DWORD    *pdwCookie
)
{
	TRACE(1, "IConnectionPoint::Advise\n");
    HRESULT    hr;
    void      *pv;

    CHECK_POINTER(pdwCookie);

    // first, make sure everybody's got what they thinks they got
    //
    if (m_bType == SINK_TYPE_EVENT) {

        // CONSIDER: 12.95 -- this theoretically is broken -- if they do a find
        // connection point on IDispatch, and they just happened to also support
        // the Event IID, we'd advise on that.  this is not awesome, but will
        // prove entirely acceptable short term.
        //
        hr = pUnk->QueryInterface(*EVENTREFIIDOFCONTROL(10), &pv);
        if (FAILED(hr))
            hr = pUnk->QueryInterface(IID_IDispatch, &pv);
    }
    else
        hr = pUnk->QueryInterface(IID_IPropertyNotifySink, &pv);
    RETURN_ON_FAILURE(hr);

    // finally, add the sink.  it's now been cast to the correct type and has
    // been AddRef'd.
    //
    return AddSink(pv, pdwCookie);
}

HRESULT CCustomizeListbox::CConnectionPoint::AddSink
(
    void  *pv,
    DWORD *pdwCookie
)
{
	TRACE(1, "IConnectionPoint::AddSink\n");
    IUnknown **rgUnkNew;
    int        i = 0;

    // we optimize the case where there is only one sink to not allocate
    // any storage.  turns out very rarely is there more than one.
    //
    switch (m_cSinks) {

        case 0:
            ASSERT(!m_rgSinks, "this should be null when there are no sinks");
            m_rgSinks = (IUnknown **)pv;
            break;

        case 1:
            // go ahead and do the initial allocation.  we'll get 8 at a time
            //
            rgUnkNew = (IUnknown **)HeapAlloc(g_hHeap, 0, 8 * sizeof(IUnknown *));
            RETURN_ON_NULLALLOC(rgUnkNew);
            rgUnkNew[0] = (IUnknown *)m_rgSinks;
            rgUnkNew[1] = (IUnknown *)pv;
            m_rgSinks = rgUnkNew;
            break;

        default:
            // if we're out of sinks, then we have to increase the size
            // of the array
            //
            if (!(m_cSinks & 0x7)) {
                rgUnkNew = (IUnknown **)HeapReAlloc(g_hHeap, 0, m_rgSinks, (m_cSinks + 8) * sizeof(IUnknown *));
                RETURN_ON_NULLALLOC(rgUnkNew);
                m_rgSinks = rgUnkNew;
            } else
                rgUnkNew = m_rgSinks;

            rgUnkNew[m_cSinks] = (IUnknown *)pv;
            break;
    }

    *pdwCookie = (DWORD)pv;
    m_cSinks++;
    return S_OK;
}


STDMETHODIMP CCustomizeListbox::CConnectionPoint::Unadvise
(
    DWORD dwCookie
)
{
	TRACE(1, "IConnectionPoint::Unadvise\n");
    IUnknown *pUnk;
    int       x;

    if (!dwCookie)
        return S_OK;

    // see how many sinks we've currently got, and deal with things based
    // on that.
    //
    switch (m_cSinks) {
        case 1:
            // it's the only sink.  make sure the ptrs are the same, and
            // then free things up
            //
            if ((DWORD)m_rgSinks != dwCookie)
                return CONNECT_E_NOCONNECTION;
            m_rgSinks = NULL;
            break;

        case 2:
            // there are two sinks.  go back down to one sink scenario
            //
            if ((DWORD)m_rgSinks[0] != dwCookie && (DWORD)m_rgSinks[1] != dwCookie)
                return CONNECT_E_NOCONNECTION;

            pUnk = ((DWORD)m_rgSinks[0] == dwCookie)
                   ? m_rgSinks[1]
                   : ((DWORD)m_rgSinks[1] == dwCookie) ? m_rgSinks[0] : NULL;

            if (!pUnk) return CONNECT_E_NOCONNECTION;

            HeapFree(g_hHeap, 0, m_rgSinks);
            m_rgSinks = (IUnknown **)pUnk;
            break;

        default:
            // there are more than two sinks.  just clean up the hole we've
            // got in our array now.
            //
            for (x = 0; x < m_cSinks; x++) {
                if ((DWORD)m_rgSinks[x] == dwCookie)
                    break;
            }
            if (x == m_cSinks) return CONNECT_E_NOCONNECTION;
            if (x < m_cSinks - 1) 
                memcpy(&(m_rgSinks[x]), &(m_rgSinks[x + 1]), (m_cSinks -1 - x) * sizeof(IUnknown *));
            else
                m_rgSinks[x] = NULL;
            break;
    }


    // we're happy
    //
    m_cSinks--;
    ((IUnknown *)dwCookie)->Release();
    return S_OK;
}

STDMETHODIMP CCustomizeListbox::CConnectionPoint::EnumConnections
(
    IEnumConnections **ppEnumOut
)
{
	TRACE(1, "IConnectionPoint::EnumConnections\n");
    CONNECTDATA *rgConnectData = NULL;
    int i;

    if (m_cSinks) {
        // allocate some memory big enough to hold all of the sinks.
        //
        rgConnectData = (CONNECTDATA *)HeapAlloc(g_hHeap, 0, m_cSinks * sizeof(CONNECTDATA));
        RETURN_ON_NULLALLOC(rgConnectData);

        // fill in the array
        //
        if (m_cSinks == 1) {
            rgConnectData[0].pUnk = (IUnknown *)m_rgSinks;
            rgConnectData[0].dwCookie = (DWORD)m_rgSinks;
        } else {
            // loop through all available sinks.
            //
            for (i = 0; i < m_cSinks; i++) {
                rgConnectData[i].pUnk = m_rgSinks[i];
                rgConnectData[i].dwCookie = (DWORD)m_rgSinks[i];
            }
        }
    }

    // create yon enumerator object.
    //
    *ppEnumOut = (IEnumConnections *)(IEnumX *)new CEnumX(IID_IEnumConnections,
                        m_cSinks, sizeof(CONNECTDATA), rgConnectData, CopyAndAddRefObject);
    if (!*ppEnumOut) 
	{
        HeapFree(g_hHeap, 0, rgConnectData);
        return E_OUTOFMEMORY;
    }

    return S_OK;
}

CCustomizeListbox::CConnectionPoint::~CConnectionPoint ()
{
	DisconnectAll();
}

void CCustomizeListbox::CConnectionPoint::DisconnectAll()
{
    int x;
    // clean up some memory stuff
    //
    if (!m_cSinks)
        return;
    else if (m_cSinks == 1)
        ((IUnknown *)m_rgSinks)->Release();
    else {
        for (x = 0; x < m_cSinks; x++)
            QUICK_RELEASE(m_rgSinks[x]);
        HeapFree(g_hHeap, 0, m_rgSinks);
    }
	m_cSinks=0;
}

void CCustomizeListbox::CConnectionPoint::DoInvoke(DISPID dispid,BYTE* pbParams,...)
{
	va_list argList;
	va_start(argList, pbParams);
	++inFireEvent;
	DoInvokeV(dispid,pbParams,argList);
	--inFireEvent;
	va_end(argList);
}

void CCustomizeListbox::CConnectionPoint::DoInvokeV(DISPID dispid,BYTE* pbParams,va_list argList)
{
	TRACE(1, "IConnectionPoint::DoInvoke\n");
	
    int iConnection;
    if (m_cSinks == 0)
        return;
    else if (m_cSinks == 1)
        DispatchHelper(((IDispatch *)m_rgSinks),dispid,DISPATCH_METHOD,VT_EMPTY,NULL,pbParams,argList); 
    else
        for (iConnection = 0; iConnection < m_cSinks; iConnection++)
            DispatchHelper(((IDispatch *)m_rgSinks[iConnection]),dispid,DISPATCH_METHOD,VT_EMPTY,NULL,pbParams,argList); 
	
}

void CCustomizeListbox::CConnectionPoint::DoOnChanged
(
    DISPID dispid
)
{
	TRACE(1, "IConnectionPoint::DoOnChanged\n");
    int iConnection;

    // if we don't have any sinks, then there's nothing to do.
    //
    if (m_cSinks == 0)
        return;
    else if (m_cSinks == 1)
        ((IPropertyNotifySink *)m_rgSinks)->OnChanged(dispid);
    else
        for (iConnection = 0; iConnection < m_cSinks; iConnection++)
            ((IPropertyNotifySink *)m_rgSinks[iConnection])->OnChanged(dispid);
}

BOOL CCustomizeListbox::CConnectionPoint::DoOnRequestEdit
(
    DISPID dispid
)
{
	TRACE(1, "IConnectionPoint::DoOnRequestEdit\n");
    HRESULT hr;
    int     iConnection;

    // if we don't have any sinks, then there's nothing to do.
    //
    if (m_cSinks == 0)
        hr = S_OK;
    else if (m_cSinks == 1)
        hr =((IPropertyNotifySink *)m_rgSinks)->OnRequestEdit(dispid);
    else {
        for (iConnection = 0; iConnection < m_cSinks; iConnection++) {
            hr = ((IPropertyNotifySink *)m_rgSinks[iConnection])->OnRequestEdit(dispid);
            if (hr != S_OK) break;
        }
    }

    return (hr == S_OK) ? TRUE : FALSE;
}

void CCustomizeListbox::BoundPropertyChanged(DISPID dispid)
{
	m_cpPropNotify->DoOnChanged(dispid);
}
//{IConnectionPointContainer Support Code}
STDMETHODIMP CCustomizeListbox::MapPropertyToCategory( DISPID dispid,  PROPCAT*ppropcat)
{
	if (NULL == ppropcat) 
		return E_POINTER;

	switch (dispid) 
	{
	case DISPID_FORECOLOR:   
	case DISPID_BACKCOLOR:
		*ppropcat = PROPCAT_Appearance;
		return S_OK;

	case DISPID_ENABLED:
		*ppropcat = PROPCAT_Behavior;
		return S_OK;

	default:
		return E_FAIL;
	}
}
STDMETHODIMP CCustomizeListbox::GetCategoryName( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName)
{
/*	if (PROPCAT_Scoring == propcat) 
	{
		*pbstrName = ::SysAllocString(L"Scoring"); 
        return S_OK;
    }*/
	return E_FAIL;
}

void CCustomizeListbox::Clear()
{
	delete m_pCategoryMgr;
	m_pCategoryMgr = NULL;
	if (IsWindow(m_hWnd))
		SendMessage(m_hWnd, LB_RESETCONTENT, 0, 0);
}

//
// ToolBarName
//

ToolBarName::ToolBarName(CBar* pBar, Type eType)
	: FDialog(IDD_NEWTOOLBAR),
	  m_pBar(pBar)
{
	m_eType = eType;
	m_szBuffer[0] = NULL;
}

//
// Customize
//

void ToolBarName::Customize(UINT nId, LocalizationTypes eType)
{
	TCHAR szBuffer[MAX_PATH];
	HWND hWnd = GetDlgItem(m_hWnd, nId);
	if (IsWindow(hWnd))
	{
		LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(eType);
		if (szLocaleString)
		{
			lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
			SetWindowText(hWnd, szBuffer);
		}
	}
}

//
// DialogProc
//

BOOL ToolBarName::DialogProc(UINT message,WPARAM wParam,LPARAM lParam)
{
	try
	{
		switch (message)
		{
		case WM_INITDIALOG:
			{
				SetWindowPos(m_hWnd, HWND_TOPMOST, 0,0,0,0, SWP_NOMOVE|SWP_NOSIZE);
				CenterDialog(GetParent(m_hWnd));
				LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(m_eType == eNew ? ddLTNewToolbarCaption : ddLTRenameCaption);
				if (szLocaleString)
				{
					TCHAR szBuffer[MAX_PATH];
					lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
					SetWindowText(m_hWnd, szBuffer);
				}
				else if (m_eType == eRename)
				{
					DDString strTemp;
					strTemp.LoadString(IDS_RENAMECAPTION);
					SetWindowText(m_hWnd, strTemp);
				}
				SetDlgItemText(m_hWnd, IDC_NAME, m_szBuffer);
				Customize(IDC_NAMESTATIC, ddLTToolbarName);
				Customize(IDOK, ddLTOkButton);
				Customize(IDCANCEL, ddLTCancelButton);
			}
			break;
		}
		return FDialog::DialogProc(message,wParam,lParam);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// OnOK
//

void ToolBarName::OnOK()
{
	GetDlgItemText(m_hWnd, IDC_NAME, m_szBuffer, 256);

	MAKE_WIDEPTR_FROMTCHAR(wBuffer, m_szBuffer);

	if (NULL == m_pBar->FindBand(wBuffer))
	{
		LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ToolBarName::eRename);
		FDialog::OnOK();
		return;
	}
	
	TCHAR szTitle[256];
	GetWindowText(m_hWnd, szTitle, 256);
	TCHAR* szError = LoadStringRes(IDS_TOOLBARNAMEEXISTS);
	MessageBox(m_hWnd, szError, szTitle, MB_OK);
}












































































































































































































































































































































































































































































































































