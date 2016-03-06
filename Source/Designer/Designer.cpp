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
#include "ipserver.h"
#include <stddef.h>       // for offsetof()
#include "support.h"
#include "resource.h"
#include "debug.h"
#include "Dialogs.h"
#include "DesignerInterfaces.h"
#include "..\Interfaces.h"
#include "..\EventLog.h"
#include "Designer.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

HWND g_hDlgCurrent = NULL;

DWORD WINAPI DesignerThread(LPVOID pData)
{
	try
	{
		HRESULT hResult = CoInitialize(NULL);
		if (FAILED(hResult))
			MessageBox(NULL, _T("Failed to Initialize COM."), _T("ActiveBar"), MB_OK);

		CDesigner* pDesigner = (CDesigner*)pData;

		pDesigner->AddRef();

		pDesigner->m_pDlgDesigner = new CDlgDesigner;
		if (NULL == pDesigner->m_pDlgDesigner)
			return E_OUTOFMEMORY;

		pDesigner->m_pDlgDesigner->Init(pDesigner->m_pActiveBar, pDesigner->m_bStandALone);

		pDesigner->m_pDlgDesigner->DoModeless(pDesigner->m_hWndMainParent);
		g_hDlgCurrent = pDesigner->m_pDlgDesigner->hWnd();
		if (NULL == g_hDlgCurrent)
		{
			delete pDesigner->m_pDlgDesigner;
			pDesigner->m_pDlgDesigner = NULL;
			return E_FAIL;
		}

		BOOL bResult;
		BOOL bAttached = FALSE;
		MSG msg;
		while (GetMessage(&msg, NULL, 0, 0))
		{
			if (bAttached)
			{
				bResult = AttachThreadInput(pDesigner->m_dwThreadId, pDesigner->m_dwCurrentThreadId, TRUE); 
				assert(bResult);
				if (!bResult)
					return E_FAIL;
				bAttached = TRUE;
			}

			if (NULL == g_hDlgCurrent)
				break;

			if (NULL == g_hDlgCurrent || !IsWindow(g_hDlgCurrent) || !IsDialogMessage(g_hDlgCurrent, &msg))
			{
				if (!TranslateAccelerator(pDesigner->m_pDlgDesigner->hWnd(), NULL, &msg))
					TranslateMessage(&msg);		
				DispatchMessage(&msg);
			}
		}

		IBarPrivate* pPrivate;
		if (SUCCEEDED(pDesigner->m_pActiveBar->QueryInterface(IID_IBarPrivate,(LPVOID*)&pPrivate)))
		{
			pPrivate->DesignerShutdown();
			pPrivate->Release();
		}
		
		while (::PeekMessage(&msg, NULL, WM_DESTROY, WM_DESTROY, PM_NOREMOVE))
		{
			if (!GetMessage(&msg, NULL, WM_DESTROY, WM_DESTROY))
				break;

			DispatchMessage(&msg);
		}

		bResult = AttachThreadInput(pDesigner->m_dwThreadId, pDesigner->m_dwCurrentThreadId, FALSE); 

		pDesigner->Release();

		CoUninitialize();
		ExitThread(0);
	}
	catch (...)
	{
		assert(FALSE);
	}
	return 0;
}

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CDesigner(IUnknown *pUnkOuter)
{
	CDesigner *pObject;
	pObject=new CDesigner(pUnkOuter);
	if (!pObject)
		return NULL;
	else
		return pObject->PrivateUnknown();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_Designer,Designer,_T("Designer Class"),CreateFN_CDesigner,1,0,&IID_IDesigner,_T(""));
void *CDesigner::objectDef=&DesignerObject;
CDesigner *CDesigner::CreateInstance(IUnknown *pUnkOuter)
{
	return new CDesigner(pUnkOuter);
}
//{OBJECT CREATEFN}
CDesigner::CDesigner(IUnknown *pUnkOuter)
	: m_pUnkOuter(pUnkOuter==NULL ? &m_UnkPrivate : pUnkOuter)
	, m_pActiveBar(NULL),
	  m_hDesignThread(NULL),
	  m_hWndMainParent(NULL),
	  m_pDlgDesigner(NULL)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
	m_dwThreadId = 0;
	m_bStandALone = FALSE;
}

CDesigner::~CDesigner()
{
	InterlockedDecrement(&g_cLocks);
	if (m_pActiveBar)
		m_pActiveBar->Release();
// {BEGIN CLEANUP}
// {END CLEANUP}
}

//{DEF IUNKNOWN MEMBERS}
//=--------------------------------------------------------------------------=
// CDesigner::CPrivateUnknownObject::QueryInterface
//=--------------------------------------------------------------------------=
inline CDesigner *CDesigner::CPrivateUnknownObject::m_pMainUnknown(void)
{
    return (CDesigner *)((LPBYTE)this - offsetof(CDesigner, m_UnkPrivate));
}

STDMETHODIMP CDesigner::CPrivateUnknownObject::QueryInterface(REFIID riid,void **ppvObjOut)
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

ULONG CDesigner::CPrivateUnknownObject::AddRef(void)
{
    return ++m_cRef;
}

ULONG CDesigner::CPrivateUnknownObject::Release(void)
{
    ULONG cRef = --m_cRef;
    if (!m_cRef)
        delete m_pMainUnknown();
    return cRef;
}
HRESULT CDesigner::InternalQueryInterface(REFIID riid,void  **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IDispatch))
	{
		*ppvObjOut=(void *)(IDispatch *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IDesigner))
	{
		*ppvObjOut=(void *)(IDesigner *)this;
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
// IDispatch members

STDMETHODIMP CDesigner::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CDesigner::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CDesigner::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
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
STDMETHODIMP CDesigner::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
{
#ifdef _DEBUG
	if (dispidMember<0)
		TRACE1(1,"IDispatch::Invoke -%X\n",-dispidMember)
	else
		TRACE1(1,"IDispatch::Invoke %X\n",dispidMember)
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

// IDesigner members

STDMETHODIMP CDesigner::OpenDesigner( OLE_HANDLE hWndParent,  LPDISPATCH pActiveBar, VARIANT_BOOL vbStandALone)
{
	try
	{
		if (NULL == pActiveBar)
			return E_INVALIDARG;

		m_bStandALone = VARIANT_TRUE == vbStandALone ? TRUE : FALSE;
		m_pActiveBar = (IActiveBar2*)pActiveBar;
		m_pActiveBar->AddRef();
		m_hWndMainParent = (HWND)hWndParent;

		m_dwCurrentThreadId = GetCurrentThreadId();

		m_hDesignThread = CreateThread(NULL, 0, DesignerThread, this, 0, &m_dwThreadId);
		assert(m_hDesignThread);
		if (NULL == m_hDesignThread)
			return E_FAIL;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__);
		return E_FAIL;
	}
	return NOERROR;
}
STDMETHODIMP CDesigner::CloseDesigner()
{
	if (m_pDlgDesigner)
	{
		m_pDlgDesigner->Apply();
		if (IsWindow(m_pDlgDesigner->hWnd()))
			SendMessage(m_pDlgDesigner->hWnd(), WM_CLOSE, 0, 0);
		WaitForSingleObject(m_hDesignThread, INFINITE);
		m_pDlgDesigner = NULL;
	}
	return NOERROR;
}
STDMETHODIMP CDesigner::UIDeactivateCloseDesigner()
{
	if (m_pDlgDesigner)
	{
		m_pDlgDesigner->Apply();
		if (IsWindow(m_pDlgDesigner->hWnd()))
			SendMessage(m_pDlgDesigner->hWnd(), WM_CLOSE, 0, 0);
		WaitForSingleObject(m_hDesignThread, INFINITE);
		m_pDlgDesigner = NULL;
	}
	return NOERROR; 
}
STDMETHODIMP CDesigner::SetFocus()
{
	if (m_pDlgDesigner && IsWindow(m_pDlgDesigner->hWnd()))
		::SetFocus(m_pDlgDesigner->hWnd());
	return NOERROR;
}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
