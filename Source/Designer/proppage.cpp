//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "support.h"
#include "proppage.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern HINSTANCE g_hInstResources;

// this variable is used to pass the pointer to the object to the hwnd.
static CPropertyPage *s_pLastPageCreated;

CPropertyPage::CPropertyPage()
	: m_refCount(1)
{
    m_pPropertyPageSite = NULL;
    m_hwnd = NULL;
    m_fDirty = FALSE;
    m_fActivated = FALSE;
    m_cObjects = 0;
}

CPropertyPage::~CPropertyPage()
{
    // clean up our window.
    if (m_hwnd) 
	{
        SetWindowLong(m_hwnd, GWL_USERDATA, 0xffffffff);
        DestroyWindow(m_hwnd);
    }
    // release all the objects we're holding on to.
    m_ReleaseAllObjects();
    // release the site
    QUICK_RELEASE(m_pPropertyPageSite);
}

HRESULT CPropertyPage::QueryInterface(REFIID riid,void **ppvObjOut)
{
    IUnknown *pUnk;
    *ppvObjOut = NULL;

    if (DO_GUIDS_MATCH(IID_IPropertyPage, riid)) 
	{
        pUnk = (IUnknown *)this;
    } 
	else if (DO_GUIDS_MATCH(IID_IPropertyPage2, riid)) 
	{
        pUnk = (IUnknown *)this;
    } 
	else if (DO_GUIDS_MATCH(IID_IUnknown, riid)) 
	{
        pUnk = (IUnknown *)this;
    }
    pUnk->AddRef();
    *ppvObjOut = (void *)pUnk;
    return S_OK;
}

STDMETHODIMP_(ULONG) CPropertyPage::AddRef()
{
	return ++m_refCount;
}

STDMETHODIMP_(ULONG) CPropertyPage::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}


STDMETHODIMP CPropertyPage::SetPageSite(IPropertyPageSite *pPropertyPageSite)
{
    RELEASE_OBJECT(m_pPropertyPageSite);
    m_pPropertyPageSite = pPropertyPageSite;
    if (pPropertyPageSite)
		pPropertyPageSite->AddRef();
    return S_OK;
}

STDMETHODIMP CPropertyPage::Activate(HWND hwndParent,LPCRECT prcBounds,BOOL fModal)
{
    HRESULT hr;
    // first make sure the dialog window is loaded and created.
    hr = m_EnsureLoaded();
    RETURN_ON_FAILURE(hr);

    // set our parent window if we haven't done so yet.
    if (!m_fActivated) 
	{
        SetParent(m_hwnd, hwndParent);
        m_fActivated = TRUE;
    }

    // now move ourselves to where we're told to be and show ourselves
    Move(prcBounds);
    ShowWindow(m_hwnd, SW_SHOW);
    return S_OK;
}

STDMETHODIMP CPropertyPage::Deactivate(void)
{
    if (m_hwnd)
        DestroyWindow(m_hwnd);
    m_hwnd = NULL;
    m_fActivated = FALSE;
    return S_OK;
}

STDMETHODIMP CPropertyPage::GetPageInfo(PROPPAGEINFO *pPropPageInfo)
{
    RECT rect;

    m_EnsureLoaded();

    memset(pPropPageInfo, 0, sizeof(PROPPAGEINFO));

    pPropPageInfo->pszTitle = OLESTRFROMRESID(TITLEIDOFPROPPAGE(0));
    pPropPageInfo->pszDocString = OLESTRFROMRESID(DOCSTRINGIDOFPROPPAGE(0));
    pPropPageInfo->pszHelpFile = OLESTRFROMANSI(HELPFILEOFPROPPAGE(0));
    pPropPageInfo->dwHelpContext = HELPCONTEXTOFPROPPAGE(0);

    if (!(pPropPageInfo->pszTitle && pPropPageInfo->pszDocString && pPropPageInfo->pszHelpFile))
        goto CleanUp;

    // if we've got a window yet, go and set up the size information they want.
    if (m_hwnd) 
	{
        GetWindowRect(m_hwnd, &rect);

        pPropPageInfo->size.cx = rect.right - rect.left;
        pPropPageInfo->size.cy = rect.bottom - rect.top;
    }

    return S_OK;

  CleanUp:
    if (pPropPageInfo->pszDocString) CoTaskMemFree(pPropPageInfo->pszDocString);
    if (pPropPageInfo->pszHelpFile) CoTaskMemFree(pPropPageInfo->pszHelpFile);
    if (pPropPageInfo->pszTitle) CoTaskMemFree(pPropPageInfo->pszTitle);

    return E_OUTOFMEMORY;
}

STDMETHODIMP CPropertyPage::SetObjects(ULONG cObjects,IUnknown **ppUnkObjects)
{
    HRESULT hr;
    ULONG   x;
    // free up all the old objects first.
    m_ReleaseAllObjects();
    if (!cObjects)
        return S_OK;
    // now go and set up the new ones.
    m_ppUnkObjects = (IUnknown **)HeapAlloc(g_hHeap, 0, cObjects * sizeof(IUnknown *));
    RETURN_ON_NULLALLOC(m_ppUnkObjects);
    // loop through and copy over all the objects.
    for (x = 0; x < cObjects; x++) 
	{
        m_ppUnkObjects[x] = ppUnkObjects[x];
        if (m_ppUnkObjects[x])
			m_ppUnkObjects[x]->AddRef();
    }
    // go and tell the object that there are new objects
    hr = S_OK;
    m_cObjects = cObjects;
    // if we've got a window, go and notify it that we've got new objects.
    if (m_hwnd)
        SendMessage(m_hwnd, PPM_NEWOBJECTS, 0, (LPARAM)&hr);
    if (SUCCEEDED(hr)) 
		m_fDirty = FALSE;
    return hr;
}

STDMETHODIMP CPropertyPage::Show(UINT nCmdShow)
{
    if (m_hwnd)
        ShowWindow(m_hwnd, nCmdShow);
    else
        return E_UNEXPECTED;
    return S_OK;
}

STDMETHODIMP CPropertyPage::Move(LPCRECT prcBounds)
{
    if (m_hwnd)
        SetWindowPos(m_hwnd, NULL, prcBounds->left, prcBounds->top,
                     prcBounds->right - prcBounds->left,
                     prcBounds->bottom - prcBounds->top,
                     SWP_NOZORDER);
    else
        return E_UNEXPECTED;
    return S_OK;
}

STDMETHODIMP CPropertyPage::IsPageDirty(void)
{
    return m_fDirty ? S_OK : S_FALSE;
}

STDMETHODIMP CPropertyPage::Apply(void)
{
    HRESULT hr = S_OK;
    if (m_hwnd) 
	{
        SendMessage(m_hwnd, PPM_APPLY, 0, (LPARAM)&hr);
        RETURN_ON_FAILURE(hr);
        if (m_fDirty) 
		{
            m_fDirty = FALSE;
            if (m_pPropertyPageSite)
                m_pPropertyPageSite->OnStatusChange(PROPPAGESTATUS_DIRTY);
        }
    } 
	else
        return E_UNEXPECTED;
    return S_OK;
}

STDMETHODIMP CPropertyPage::Help(LPCOLESTR pszHelpDir)
{
    BOOL f;
#ifdef UNICODE
	f = WinHelp(m_hwnd, pszHelpDir, HELP_CONTEXT, HELPCONTEXTOFPROPPAGE(0));
#else
    MAKE_ANSIPTR_FROMWIDE(psz, (LPOLESTR)pszHelpDir);
    f = WinHelp(m_hwnd, psz, HELP_CONTEXT, HELPCONTEXTOFPROPPAGE(0));
#endif
    return f ? S_OK : E_FAIL;
}

STDMETHODIMP CPropertyPage::TranslateAccelerator(LPMSG pmsg)
{
    return IsDialogMessage(m_hwnd, pmsg) ? S_OK : S_FALSE;
}

STDMETHODIMP CPropertyPage::EditProperty(DISPID dispid)
{
    HRESULT hr = E_NOTIMPL;
    // send the message on to the control, and see what they want to do with it.
    SendMessage(m_hwnd, PPM_EDITPROPERTY, (WPARAM)dispid, (LPARAM)&hr);
    return hr;
}

HRESULT CPropertyPage::m_EnsureLoaded(void)
{
    HRESULT hr = S_OK;
    if (m_hwnd)
        return S_OK;
    // set up the global variable so that when we're in the dialog proc, we can
    // stuff this in the hwnd
    // crit sect this whole creation process for apartment threading support.
#ifdef DEF_CRITSECTION
    EnterCriticalSection(&g_CriticalSection);
#endif
    s_pLastPageCreated = this;
    // create the dialog window
    CreateDialog(GetResourceHandle(), TEMPLATENAMEOFPROPPAGE(0), GetParkingWindow(),
                          (DLGPROC)CPropertyPage::PropPageDlgProc);
    ASSERT(m_hwnd, "Couldn't load Dialog Resource!!!");
    if (!m_hwnd) 
	{
#ifdef DEF_CRITSECTION
        LeaveCriticalSection(&g_CriticalSection);
#endif
        return HRESULT_FROM_WIN32(GetLastError());
    }
    // clean up variables and leave the critical section
    s_pLastPageCreated = NULL;
#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif
    // go and notify the window that it should pick up any objects that are
    // available
    SendMessage(m_hwnd, PPM_NEWOBJECTS, 0, (LPARAM)&hr);
    return hr;
}

void CPropertyPage::m_ReleaseAllObjects(void)
{
    HRESULT hr;
    UINT x;
    if (!m_cObjects)
        return;
    // some people will want to stash pointers in the PPM_INITOBJECTS case, so
    // we want to tell them to release them now.
	if (IsWindow(m_hwnd))
		SendMessage(m_hwnd, PPM_FREEOBJECTS, 0, (LPARAM)&hr);
    // loop through and blow them all away.
    for (x = 0; x < m_cObjects; x++)
        QUICK_RELEASE(m_ppUnkObjects[x]);
	m_cObjects = 0;
    HeapFree(g_hHeap, 0, m_ppUnkObjects);
    m_ppUnkObjects = NULL;
}

BOOL CALLBACK CPropertyPage::PropPageDlgProc(HWND hwnd,UINT msg,WPARAM wParam,LPARAM lParam)
{
    CPropertyPage *pPropertyPage;
    // get the window long, and see if it's been set to the object this hwnd
    // is operating against.  if not, go and set it now.
    pPropertyPage = (CPropertyPage *)GetWindowLong(hwnd, GWL_USERDATA);
    if ((ULONG)pPropertyPage == 0xffffffff)
        return FALSE;
    if (!pPropertyPage) 
	{
        SetWindowLong(hwnd, GWL_USERDATA, (LONG)s_pLastPageCreated);
        pPropertyPage = s_pLastPageCreated;
        pPropertyPage->m_hwnd = hwnd;
    }

    ASSERT(pPropertyPage, "window, but no associated CPropertyPage");
    // just call the user dialog proc and see if they want to do anything.
    return pPropertyPage->DialogProc(hwnd, msg, wParam, lParam);
}

IUnknown *CPropertyPage::FirstControl(DWORD *pdwCookie)
{
    *pdwCookie = 0;
    return NextControl(pdwCookie);
}

IUnknown *CPropertyPage::NextControl(DWORD *pdwCookie)
{
    UINT      i;
    for (i = *pdwCookie; i < m_cObjects; i++) 
	{
        if (!m_ppUnkObjects[i]) 
			continue;
        *pdwCookie = i + 1;
        return m_ppUnkObjects[i];
    }

    *pdwCookie = 0xffffffff;
    return NULL;
}

void CPropertyPage::MakeDirty(void)
{
    m_fDirty = TRUE;
    if (m_pPropertyPageSite)
        m_pPropertyPageSite->OnStatusChange(PROPPAGESTATUS_DIRTY|PROPPAGESTATUS_VALIDATE);
}



HINSTANCE CPropertyPage::GetResourceHandle(void)
{
//    if (!g_fSatelliteLocalization)
        return g_hInstance;

/*satellite locale impl
    if (g_hInstResources)
        return g_hInstResources;

    // we'll get the ambient localeid from the host, and pass that on to the
    // automation object.
    // enter a critical section for g_lcidLocale and g_fHavelocale
    EnterCriticalSection(&g_CriticalSection);
    if (!g_fHaveLocale) 
	{
        if (m_pPropertyPageSite) 
		{
            m_pPropertyPageSite->GetLocaleID(&g_lcidLocale);
            g_fHaveLocale = TRUE;
        }
    }
    LeaveCriticalSection(&g_CriticalSection);

    return ::GetResourceHandle();
*/
}
