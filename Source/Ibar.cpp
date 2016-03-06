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
#include "Errors.h"
#include "DispIds.h"
#include "Utility.h"
#include "FDialog.h"
#include "Globals.h"
#include "Streams.h"
#include "Localizer.h"
#include "IpServer.h"
#include "ImageMgr.h"
#include "Tool.h"
#include "Band.h"
#include "Bands.h"
#include "Dock.h"
#include "Custom.h"
#include "MiniWin.h"
#include "PopupWin.h"
#include "Returnbool.h"
#include "StaticLink.h"
#include "Bar.h"

extern FONTDESC _fdDefault;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


// IDispatch members

STDMETHODIMP CBar::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CBar::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CBar::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
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
STDMETHODIMP CBar::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
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
// IDynaBar members

STDMETHODIMP CBar::get_Bands(Bands * *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	if (NULL == m_pBands)
		return E_FAIL;
	m_pBands->AddRef();
	*retval = reinterpret_cast<Bands*>(m_pBands);
	return NOERROR;
}

STDMETHODIMP CBar::RecalcLayout()
{
	InternalRecalcLayout();
	Refresh();
	return NOERROR;
}

//
// InternalRecalcLayout
//

STDMETHODIMP CBar::InternalRecalcLayout(BOOL bJustRecalc)
{
	try
	{
		HWND hWndParent = GetParentWindow();
		if (!IsWindow(hWndParent))
			return E_FAIL;

		DWORD dwStyle = GetWindowLong(GetDockWindow(), GWL_STYLE);
		if (WS_MINIMIZE & dwStyle)
			return E_FAIL;

		// starting rect comes from client rect
		SIZEPARENTPARAMS sppLayout;
		sppLayout.sizeTotal.cx = sppLayout.sizeTotal.cy = 0;

		switch (m_eAppType)
		{
		case eMDIForm:
		case eSDIForm:
			GetClientRect(GetDockWindow(), &sppLayout.rc);
			break;

		case eClientArea:
			GetClientRect(hWndParent, &sppLayout.rc);
			break;
		}

		PositionAlignableControls(sppLayout);
		if (!AmbientUserMode() && eMDIForm == m_eAppType)
			sppLayout.rc.Offset(-sppLayout.rc.left, -sppLayout.rc.top);
		
		SIZE sizeParent = sppLayout.rc.Size();

		if (!m_pDockMgr->RecalcLayout(sizeParent, bJustRecalc))
			return E_FAIL;

		if (!bJustRecalc)
		{
			if (!m_bFirstRecalcLayout)
				m_bFirstRecalcLayout = TRUE;

			int nCount = 4; // For the four dock areas
			switch (m_eAppType)
			{
			case eMDIForm:
				if (AmbientUserMode())
					nCount += 2;
				else
					nCount += 1;
				break;

			case eSDIForm:
				nCount += 1;
				break;
			}

			sppLayout.hDWP = ::BeginDeferWindowPos(nCount); 
			if (NULL == sppLayout.hDWP)
			{
				TRACE(1, "Begin Warning: DeferWindowPos failed - low system resources.\n");
				if (GetGlobals().GetEventLog())
					GetGlobals().GetEventLog()->WriteCustom(_T("Begin Warning: DeferWindowPos failed - low system resources."), EVENTLOG_ERROR_TYPE);
			}

			const int nLen = 32;
			TCHAR szClassName[nLen];
			HWND  hWndChild; 
			switch (m_eAppType)
			{
			case eMDIForm:
				try
				{
					//
					// Runtime updating of windows
					//

					if (IsWindow(m_hWnd))
					{
						//
						// Updating the ActiveBar Control Window
						//

						if (AmbientUserMode())
						{
							m_rcFrame = sppLayout.rc;
							m_rcFrame.bottom = m_rcFrame.top;
							RepositionWindow(&sppLayout, m_hWnd, m_rcFrame);
							sppLayout.rc.top += m_rcFrame.Height();
						}
						else
							m_rcFrame = sppLayout.rc;
					}
					
					//
					// Updating the SIZEPARENTPARAMS parameter with the size of the dock windows
					//

					for (hWndChild = ::GetTopWindow(hWndParent); 
						 hWndChild; 
						 hWndChild = ::GetNextWindow(hWndChild, GW_HWNDNEXT))
					{
						GetClassName(hWndChild, szClassName, nLen);
						if (0 == lstrcmpi(szClassName, DD_WNDCLASS("ActiveBarDockWnd")) && IsWindow(hWndChild))
							SendMessage(hWndChild, GetGlobals().WM_SIZEPARENT, (WPARAM)&sppLayout, 0);
					}

					//
					// Sizing the MDIClient Window
					//

					if (IsWindow(m_hWndMDIClient))
					{
						m_rcMDIClient = sppLayout.rc;
						RepositionWindow(&sppLayout, m_hWndMDIClient, sppLayout.rc);
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case eSDIForm:
			case eClientArea:
				try
				{
					//
					// Updating the SIZEPARENTPARAMS parameter with the size of the dock windows
					//

					m_rcFrame = sppLayout.rc;

					for (hWndChild = ::GetTopWindow(hWndParent); 
						 hWndChild; 
						 hWndChild = ::GetNextWindow(hWndChild, GW_HWNDNEXT))
					{
						GetClassName(hWndChild, szClassName, nLen);
						if (0 == lstrcmpi(szClassName, DD_WNDCLASS("ActiveBarDockWnd")) && IsWindow(hWndChild))
							SendMessage(hWndChild, GetGlobals().WM_SIZEPARENT, (WPARAM)&sppLayout, 0);
					}

					//
					// Repositioning the Child Windows of an SDI Form
					//
					
					switch (bpV1.m_asChildren)
					{
					case ddASNone:
						{
							//
							// Resize Event
							//
					
							SIZE size = {sppLayout.rc.left, sppLayout.rc.top};
							PixelToTwips(&size, &size);
							SIZE size2 = {sppLayout.rc.Width(), sppLayout.rc.Height()};
							PixelToTwips(&size2, &size2);
							FireResize(size.cx, size.cy, size2.cx, size2.cy);
							m_rcInsideFrame.Set(size.cx, size.cy, size2.cx, size2.cy);
						}
						break;

					case ddASProportional:
						if (NULL == m_pRelocate)
						{
							m_pRelocate = new CRelocate(this);
							if (NULL == m_pRelocate)
								return E_OUTOFMEMORY;
						}
						if (m_pRelocate->Count() < 1)
							m_pRelocate->GetLayoutInfo(hWndParent, sppLayout.rc);
						if (m_pRelocate->Count() > 0)
							m_pRelocate->LayoutProportional(hWndParent, sppLayout.rc);
						break;

					case ddASClientArea:
						if (NULL == m_pRelocate)
						{
							m_pRelocate = new CRelocate(this);
							if (NULL == m_pRelocate)
								return E_OUTOFMEMORY;
						}
						if (m_pRelocate->Count() > 0)
							m_pRelocate->LayoutClient(hWndParent, sppLayout.rc);
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

			//
			// Move and resize all dock windows at once!
			//

			if (!::EndDeferWindowPos(sppLayout.hDWP))
			{
				assert(FALSE);
				LPTSTR szMsg = GetLastErrorMsg();
				if (GetGlobals().GetEventLog())
					GetGlobals().GetEventLog()->WriteCustom(szMsg, EVENTLOG_ERROR_TYPE);
				LocalFree(szMsg);
				return E_FAIL;
			}

			if (!m_bInitialRecalcLayout)
			{
				//
				// Fire an initial band open event
				//

				m_bInitialRecalcLayout = TRUE;
			}
			UpdateWindow(hWndParent);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

STDMETHODIMP CBar::get_DisplayToolTips(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_vbDisplayToolTips;
	return NOERROR;
}
STDMETHODIMP CBar::put_DisplayToolTips(VARIANT_BOOL val)
{
	bpV1.m_vbDisplayToolTips=val;
	return NOERROR;
}
STDMETHODIMP CBar::get_DisplayKeysInToolTip(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_vbDisplayKeysInToolTip;
	return NOERROR;
}
STDMETHODIMP CBar::put_DisplayKeysInToolTip(VARIANT_BOOL val)
{
	bpV1.m_vbDisplayKeysInToolTip=val;
	return NOERROR;
}
HWND CBar::GetDockWindow()
{
	if (m_hWndDock)
		return m_hWndDock;

	IOleInPlaceSite* m_pInPlaceSite;
	if (NULL == m_pClientSite)
		return NULL;

    HRESULT hResult = m_pClientSite->QueryInterface(IID_IOleInPlaceSite, (void**)&m_pInPlaceSite);
	if (FAILED(hResult))
		return NULL;

	if (NULL == m_pInPlaceSite)
		return NULL;

	HWND hWndParent;
	hResult = m_pInPlaceSite->GetWindow(&hWndParent);
	m_pInPlaceSite->Release();
    if (SUCCEEDED(hResult))
	{
		m_hWndDock = hWndParent;
		return hWndParent;
	}
	return NULL;
}

STDMETHODIMP CBar::get_Tools(Tools** retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = reinterpret_cast<Tools*>(m_pTools);

	m_pTools->AddRef();
	return NOERROR;
}

STDMETHODIMP CBar::get_ActiveBand(Band** retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	if (NULL == m_pActiveBand)
	{
		*retval = NULL;
		return E_FAIL;
	}
	m_pActiveBand->AddRef();
	*retval = (Band*)m_pActiveBand;
	return NOERROR;
}

STDMETHODIMP CBar::Customize(CustomizeTypes ctCustomize)
{
	bpV1.m_ctCustomize = ctCustomize;
	switch (ctCustomize)
	{
	case ddCTCustomizeStart:
		DoCustomization(TRUE);
		if (VARIANT_FALSE == bpV1.m_vbUserDefinedCustomization)
		{
			m_pCustomize = new CCustomize(GetDockWindow(), this);
			if (m_pCustomize)
			{
				HWND hWndPropertySheet = m_pCustomize->ShowSheet();
				MSG msg;
				while (GetMessage(&msg, NULL, 0, 0))
				{
					if (!IsWindow(hWndPropertySheet))
						break;
					if (NULL == m_pCustomize)
						break;
					if (PropSheet_IsDialogMessage(hWndPropertySheet, &msg))
						continue;
					DispatchMessage(&msg);
				}
			}
		}
		break;

	case ddCTCustomizeStop:
		DoCustomization(FALSE);
		break;
	}
	return NOERROR;
}

//
// ReleaseFocus
//

STDMETHODIMP CBar::ReleaseFocus()
{
	HWND hWndDock = GetDockWindow();
	if (hWndDock)
		::SetFocus(hWndDock);
	return NOERROR;
}

//
// Load
//

STDMETHODIMP CBar::Load(BSTR bstrBandName, VARIANT* vData, SaveOptionTypes nOptions)
{
	HRESULT hResult;
	if (ddSOByteArray & nOptions)
	{
		SAFEARRAY* psa;
		if (vData->vt == (VT_ARRAY|VT_UI1))
			psa = vData->parray;
		else if (vData->vt == (VT_BYREF|VT_ARRAY|VT_UI1))
		{
			if (NULL == vData->pparray)
				return E_FAIL;
			psa = *vData->pparray;
		}
		else
			return E_INVALIDARG;
		
		if (NULL == psa)
			return E_FAIL;

		LPBYTE pArg;
		SafeArrayAccessData(psa, (void**)&pArg);

		HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, psa->rgsabound[0].cElements);
		if (NULL == hGlobal)
		{
			SafeArrayUnaccessData(psa);
			return E_FAIL;
		}

		LPVOID pData = GlobalLock(hGlobal);
		if (NULL == pData)
		{
			GlobalFree(hGlobal);
			GlobalUnlock(hGlobal);
			SafeArrayUnaccessData(psa);
			return E_FAIL;
		}

		DWORD dwSize = GlobalSize(hGlobal);

		memcpy(pData, pArg, dwSize);
		SafeArrayUnaccessData(psa);
		GlobalUnlock(hGlobal);

		IStream* pStream;
		hResult = CreateStreamOnHGlobal(hGlobal, TRUE, &pStream);
		if (FAILED(hResult))
			return hResult;
		
		LARGE_INTEGER pos;
		pos.LowPart = pos.HighPart = 0;
		hResult = pStream->Seek(pos, STREAM_SEEK_SET, 0);
		if (FAILED(hResult))
		{
			pStream->Release();
			return hResult;
		}

		hResult = LoadData(pStream, FALSE, bstrBandName);
		pStream->Release();
		if (FAILED(hResult))
		{
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
		}
	}
	else if (ddSOFile & nOptions)
	{
		OFSTRUCT of;
		of.cBytes = sizeof(OFSTRUCT);

		BSTR bstrFile = NULL;
		if (vData->vt == VT_BSTR)
			bstrFile = vData->bstrVal;
		else if (vData->vt == (VT_BYREF|VT_BSTR))
			bstrFile = *vData->pbstrVal;
		else
			return E_INVALIDARG;

		MAKE_TCHARPTR_FROMWIDE(szFileName, bstrFile);
		HANDLE hFile = CreateFile(szFileName, 
								  GENERIC_READ, 
								  FILE_SHARE_READ, 
								  NULL, 
								  OPEN_EXISTING, 
								  FILE_ATTRIBUTE_NORMAL, 
								  NULL);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			DDString strError;
			strError.Format(_T("%s"), GetLastErrorMsg());
			MAKE_WIDEPTR_FROMTCHAR(wError, strError);
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_OPENFILE, wError);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_OPENFILE);
		}

		FileStream* pFileStream = new FileStream;
		if (NULL == pFileStream)
		{
			CloseHandle((HANDLE)hFile);
			return E_OUTOFMEMORY;
		}
		pFileStream->SetHandle(hFile, TRUE);
		
		hResult = LoadData(pFileStream, FALSE, bstrBandName);
		pFileStream->Release();
		if (FAILED(hResult))
		{
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
		}
	}
	else if (ddSOResource & nOptions)
	{
		IStream* pResourceStream = NULL;
		
		if (vData->vt != VT_I4)
			return E_FAIL;
		
		if (NULL == vData->lVal)
			return E_FAIL;

		MAKE_TCHARPTR_FROMWIDE(szBandName, bstrBandName);
		long nResource = atoi(szBandName);
		hResult =  CResourceStream::GetIStream((HINSTANCE)vData->lVal, 
											   MAKEINTRESOURCE(nResource),
											   _T("LAYOUT"),
											   &pResourceStream);
		if (FAILED(hResult))
			return hResult;

		hResult = LoadData(pResourceStream, FALSE, L"");
		pResourceStream->Release();
		if (S_FALSE == hResult)
		{
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
		}
	}
	else if (ddSOStream & nOptions)
	{
		if (VT_UNKNOWN != vData->vt)
			return E_FAIL;

		IStream* pStream = NULL;
		hResult = vData->punkVal->QueryInterface(IID_IStream, (LPVOID*)&pStream);
		if (FAILED(hResult))
			return hResult;

		hResult = LoadData(pStream, FALSE, bstrBandName);
		pStream->Release();
		if (S_FALSE == hResult)
		{
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
		}
	}
	else
		return E_INVALIDARG;
	CacheColors();
	CacheTexture();
	m_theChildMenus.MainBand(GetMenuBand());
	BuildAccelators(m_theChildMenus.MainBand());
	OnFontHeightChanged();
	return hResult;
}

//
// Save
//

STDMETHODIMP CBar::Save(BSTR bstrBandName, BSTR bstrFileName, SaveOptionTypes nOptions, VARIANT* vData)
{
	HRESULT hResult = NOERROR;
	if (ddSOByteArray & nOptions)
	{
		HRESULT hResult = VariantClear(vData);
		if (FAILED(hResult))
			return hResult;
		vData->vt = (VT_ARRAY|VT_UI1);

		IStream* pStream;
		hResult = CreateStreamOnHGlobal(NULL, TRUE, &pStream);
		if (FAILED(hResult))
			return hResult;
		
		HGLOBAL hGlobal;
		hResult = GetHGlobalFromStream(pStream, &hGlobal);
		if (FAILED(hResult))
		{
			pStream->Release();
			return hResult;
		}

		hResult = SaveData(pStream, FALSE, bstrBandName);
		if (FAILED(hResult))
		{
			pStream->Release();
			return hResult;
		}

		LPVOID pData = GlobalLock(hGlobal);
		if (NULL == pData)
		{
			pStream->Release();
			return hResult;
		}

		DWORD dwSize = GlobalSize(hGlobal);

		hResult = SafeArrayAllocDescriptor(1, &vData->parray);
		if (FAILED(hResult))
		{
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}
		vData->parray->fFeatures = 0;
		vData->parray->cbElements = sizeof(BYTE);
		vData->parray->rgsabound[0].cElements = dwSize;
		vData->parray->rgsabound[0].lLbound = 0;
		hResult = SafeArrayAllocData(vData->parray);
		if (FAILED(hResult))
		{
			SafeArrayDestroyDescriptor(vData->parray);
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}
		LPBYTE pArg;
		hResult = SafeArrayAccessData(vData->parray, (void**)&pArg);
		if (FAILED(hResult))
		{
			SafeArrayUnaccessData(vData->parray);
			SafeArrayDestroyDescriptor(vData->parray);
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}
		memcpy (pArg, pData, dwSize);
		SafeArrayUnaccessData(vData->parray);
		GlobalUnlock(hGlobal);
		pStream->Release();
	}
	else if (ddSOFile & nOptions)
	{
		TCHAR szTempFilePath[MAX_PATH];
		UINT nResult = GetTempPath(MAX_PATH, szTempFilePath);
		if (0 == nResult)
			return E_FAIL;

		TCHAR szTempFile[MAX_PATH];
		nResult = GetTempFileName(szTempFilePath, _T("ab2"), 0, szTempFile);
		if (0 == nResult)
			return E_FAIL;

		OFSTRUCT of;
		of.cBytes = sizeof(OFSTRUCT);

		HANDLE hFile = CreateFile(szTempFile, 
								  GENERIC_WRITE, 
								  FILE_SHARE_READ, 
								  NULL, 
								  OPEN_ALWAYS, 
								  FILE_ATTRIBUTE_NORMAL, 
								  NULL);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			DDString strError;
			strError.Format(_T("%s"), GetLastErrorMsg());
			MAKE_WIDEPTR_FROMTCHAR(wError, strError);
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_CREATEFILE, wError);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_CREATEFILE);
		}

		FileStream* pFileStream = new FileStream;
		if (NULL == pFileStream)
		{
			CloseHandle((HANDLE)hFile);
			DeleteFile(szTempFile);
			return E_OUTOFMEMORY;
		}

		pFileStream->SetHandle(hFile, TRUE);

		hResult = SaveData(pFileStream, FALSE, bstrBandName);
		pFileStream->Release();
		if (SUCCEEDED(hResult))
		{
			MAKE_TCHARPTR_FROMWIDE(szFileName, bstrFileName);
			DWORD dwResult = CopyFile(szTempFile, szFileName, FALSE);
		}
		DeleteFile(szTempFile);
	}
	else if (ddSOStream & nOptions)
	{
		if (VT_UNKNOWN != vData->vt)
			return E_FAIL;

		IStream* pStream;
		hResult = vData->punkVal->QueryInterface(IID_IStream, (LPVOID*)&pStream);
		if (FAILED(hResult))
			return hResult;

		hResult = SaveData(pStream, FALSE, bstrBandName);
		pStream->Release();
		if (S_FALSE == hResult)
		{
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
		}
	}
	else
		return E_INVALIDARG;
	return hResult;
}

//
// SaveLayoutChanges
//

STDMETHODIMP CBar::SaveLayoutChanges(BSTR bstrFileName, SaveOptionTypes nOptions, VARIANT* vData)
{
	HRESULT hResult = NOERROR;
	// This is dispatch so the VB doesn't complain about data type
	if (ddSOByteArray & nOptions)
	{
		HRESULT hResult = VariantClear(vData);
		if (FAILED(hResult))
			return hResult;
		vData->vt = (VT_ARRAY|VT_UI1);

		IStream* pStream;
		hResult = CreateStreamOnHGlobal(NULL, TRUE, &pStream);
		if (FAILED(hResult))
			return hResult;
		
		HGLOBAL hGlobal;
		hResult = GetHGlobalFromStream(pStream, &hGlobal);
		if (FAILED(hResult))
		{
			pStream->Release();
			return hResult;
		}

		hResult = SaveData(pStream, TRUE, NULL);

		LPVOID pData = GlobalLock(hGlobal);
		if (NULL == pData)
		{
			pStream->Release();
			return hResult;
		}

		DWORD dwSize = GlobalSize(hGlobal);

		hResult = SafeArrayAllocDescriptor(1, &vData->parray);
		if (FAILED(hResult))
		{
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}

		vData->parray->fFeatures = 0;
		vData->parray->cbElements = sizeof(BYTE);
		vData->parray->rgsabound[0].cElements = dwSize;
		vData->parray->rgsabound[0].lLbound = 0;
		hResult = SafeArrayAllocData(vData->parray);
		if (FAILED(hResult))
		{
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}

		LPBYTE pArg;
		hResult = SafeArrayAccessData(vData->parray, (void**)&pArg);
		if (FAILED(hResult))
		{
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}

		memcpy (pArg, pData, dwSize);
		hResult = SafeArrayUnaccessData(vData->parray);
		GlobalUnlock(hGlobal);
		pStream->Release();
	}
	else if (ddSOFile & nOptions)
	{
		TCHAR szTempFilePath[MAX_PATH];
		UINT nResult = GetTempPath(MAX_PATH, szTempFilePath);
		if (0 == nResult)
			return E_FAIL;

		TCHAR szTempFile[MAX_PATH];
		nResult = GetTempFileName(szTempFilePath, _T("ab2"), 0, szTempFile);
		if (0 == nResult)
			return E_FAIL;

		OFSTRUCT of;
		of.cBytes = sizeof(OFSTRUCT);

		HANDLE hFile = CreateFile(szTempFile, 
								  GENERIC_WRITE, 
								  FILE_SHARE_READ, 
								  NULL, 
								  OPEN_ALWAYS, 
								  FILE_ATTRIBUTE_NORMAL, 
								  NULL);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			DDString strError;
			strError.Format(_T("%s"), GetLastErrorMsg());
			MAKE_WIDEPTR_FROMTCHAR(wError, strError);
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_CREATEFILE, wError);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_CREATEFILE);
		}

		FileStream* pFileStream = new FileStream;
		if (NULL == pFileStream)
		{
			CloseHandle((HANDLE)hFile);
			DeleteFile(szTempFile);
			return E_OUTOFMEMORY;
		}

		pFileStream->SetHandle(hFile, TRUE);

		hResult = SaveData(pFileStream, TRUE, NULL);
		pFileStream->Release();
		if (SUCCEEDED(hResult))
		{
			MAKE_TCHARPTR_FROMWIDE(szFileName, bstrFileName);
			DWORD dwResult = CopyFile(szTempFile, szFileName, FALSE);
		}
		DeleteFile(szTempFile);
	}
	else
		return E_INVALIDARG;
	return hResult;
}

//
// LoadLayoutChanges
//

STDMETHODIMP CBar::LoadLayoutChanges(VARIANT* vData, SaveOptionTypes nOptions)
{
	HRESULT hResult;
	if (ddSOByteArray & nOptions)
	{
		SAFEARRAY* psa;
		if (vData->vt == (VT_ARRAY|VT_UI1))
			psa = vData->parray;
		else if (vData->vt == (VT_BYREF|VT_ARRAY|VT_UI1))
		{
			if (NULL == vData->pparray)
				return E_FAIL;
			psa = *vData->pparray;
		}
		else
			return E_FAIL;
		
		if (NULL == psa)
			return E_FAIL;

		LPBYTE pArg;
		SafeArrayAccessData(vData->parray, (void**)&pArg);

		HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, psa->rgsabound[0].cElements);
		if (NULL == hGlobal)
		{
			SafeArrayUnaccessData(psa);
			return E_FAIL;
		}

		LPVOID pData = GlobalLock(hGlobal);
		if (NULL == pData)
		{
			GlobalFree(hGlobal);
			GlobalUnlock(hGlobal);
			SafeArrayUnaccessData(psa);
			return E_FAIL;
		}

		DWORD dwSize = GlobalSize(hGlobal);

		memcpy(pData, pArg, dwSize);
		SafeArrayUnaccessData(psa);
		GlobalUnlock(hGlobal);

		IStream* pStream;
		hResult = CreateStreamOnHGlobal(hGlobal, TRUE, &pStream);
		if (FAILED(hResult))
			return hResult;
		
		LARGE_INTEGER pos;
		pos.LowPart = pos.HighPart = 0;
		pStream->Seek(pos, STREAM_SEEK_SET, 0);

		hResult = LoadData(pStream, TRUE, NULL);
		
		pStream->Release();
		if (FAILED(hResult))
		{
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
		}
	}
	else if (ddSOFile & nOptions)
	{
		OFSTRUCT of;
		of.cBytes = sizeof(OFSTRUCT);

		BSTR bstrFile = NULL;
		if (vData->vt == VT_BSTR)
			bstrFile = vData->bstrVal;
		else if (vData->vt == (VT_BYREF|VT_BSTR))
			bstrFile = *vData->pbstrVal;
		else
			return E_INVALIDARG;

		MAKE_TCHARPTR_FROMWIDE(szFileName, bstrFile);
		HANDLE hFile = CreateFile(szFileName, 
								  GENERIC_READ, 
								  FILE_SHARE_READ, 
								  NULL, 
								  OPEN_EXISTING, 
								  FILE_ATTRIBUTE_NORMAL, 
								  NULL);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			DDString strError;
			strError.Format(_T("%s"), GetLastErrorMsg());
			MAKE_WIDEPTR_FROMTCHAR(wError, strError);
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_OPENFILE, wError);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_OPENFILE);
		}

		FileStream* pFileStream = new FileStream;
		if (NULL == pFileStream)
		{
			CloseHandle((HANDLE)hFile);
			return E_OUTOFMEMORY;
		}
		pFileStream->SetHandle(hFile, TRUE);

		hResult = LoadData(pFileStream, TRUE, NULL);

		pFileStream->Release();
		if (S_FALSE == hResult)
		{
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
		}
	}
	else
		return E_INVALIDARG;
	return hResult;
}

STDMETHODIMP CBar::get_Font(IFontDisp** ppRetval)
{
	*ppRetval = m_fhFont.GetFontDispatch();
	return NOERROR;
}

STDMETHODIMP CBar::put_Font(IFontDisp* pVal)
{
	BOOL bResult;
	m_fhFont.InitializeFont(NULL, pVal);
	m_cached_fhFont = NULL;
	if (m_hFontMenuVert)
	{
		bResult = DeleteFont(m_hFontMenuVert);
		assert(bResult);
		m_hFontMenuVert = NULL;
	}
	if (m_hFontMenuHorz)
	{
		bResult = DeleteFont(m_hFontMenuHorz);
		assert(bResult);
		m_hFontMenuHorz = NULL;
	}
	OnFontHeightChanged();
	ViewChanged();
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::putref_Font(IFontDisp ** val)
{
	return put_Font(*val);
}

//
// AboutDlg
//

class AboutDlg: public FDialog
{
public:
	AboutDlg::AboutDlg()
		: FDialog(IDD_ABOUT),
		  m_theLink(_T(""), FALSE),
		  m_theLinkIcon(_T("http://www.datadynamics.com"), FALSE),
		  m_viFileVersion(g_hInstance, _T("actbar2.ocx"))
	{
	}
	
	BOOL AboutDlg::DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
	{
		switch (nMsg)
		{
		case WM_INITDIALOG:
			CenterDialog(GetParent(m_hWnd));
			m_theLink.SubclassDlgItem(IDC_WEBADDRESS, m_hWnd);
			m_theLinkIcon.SubclassDlgItem(IDC_ABICON, m_hWnd);
			m_viFileVersion.SetLabel(m_hWnd, IDC_VERSION, _T("ActiveBar, Version "), _T("FileVersion"));
			break;

 		case WM_CTLCOLORSTATIC:
			return SendMessage((HWND)lParam, WM_CTLCOLORSTATIC, wParam, lParam);
		}
		return FDialog::DialogProc(nMsg, wParam, lParam);
	}

	CVersionInfo m_viFileVersion;
	CStaticLink  m_theLinkIcon;
	CStaticLink  m_theLink;
};

STDMETHODIMP CBar::About()
{
	AboutDlg dlgAbout;
	dlgAbout.DoModal(GetFocus());
	return NOERROR;
}

STDMETHODIMP CBar::get_DataPath(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = SysAllocString(m_bstrDataPath);
	return NOERROR;
}
STDMETHODIMP CBar::put_DataPath(BSTR val)
{
	SysFreeString(m_bstrDataPath);
	m_bstrDataPath = SysAllocString(val);
	if (AmbientUserMode())
		PostMessage(m_hWnd, WM_COMMAND, MAKELONG(eStartDownload,0), 0);
	
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::get_MenuFontStyle(MenuFontStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = bpV1.m_mfsMenuFont;
	return NOERROR;
}
STDMETHODIMP CBar::put_MenuFontStyle(MenuFontStyles val)
{
	bpV1.m_mfsMenuFont = val;

	if (m_hFontMenuVert)
		DeleteFont(m_hFontMenuVert);
	m_hFontMenuVert = NULL;

	if (m_hFontMenuHorz)
		DeleteFont(m_hFontMenuHorz);
	m_hFontMenuHorz = NULL;

	OnFontHeightChanged();
	ViewChanged();
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::get_Picture(LPPICTUREDISP *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = m_phPict.GetPictureDispatch();
	return NOERROR;
}
STDMETHODIMP CBar::put_Picture(LPPICTUREDISP val)
{
	short nType;
	HRESULT hResult;
	IPicture* pPicture;
	if (val && SUCCEEDED(val->QueryInterface(IID_IPicture, (LPVOID*)&pPicture)))
	{
		hResult = pPicture->get_Type(&nType);
		pPicture->Release();
		if (SUCCEEDED(hResult))
		{
			if (PICTYPE_BITMAP != nType)
			{
				m_theErrorObject.SendError(IDERR_INVALIDPICTURETYPE, NULL);
				return CUSTOM_CTL_SCODE(IDERR_INVALIDPICTURETYPE);
			}
		}
	}
	m_phPict.SetPictureDispatch(val);
	CacheTexture();
	SetModified();
	return NOERROR;
}

STDMETHODIMP CBar::putref_Picture(LPPICTUREDISP *val)
{
	return put_Picture(*val);
}

STDMETHODIMP CBar::get_BackColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_ocBackground;
	return NOERROR;
}
STDMETHODIMP CBar::put_BackColor(OLE_COLOR val)
{
	bpV1.m_ocBackground=val;
	CacheColors();
	if (!AmbientUserMode())
		Refresh();
	ViewChanged();
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::get_ForeColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_ocForeground;
	return NOERROR;
}
STDMETHODIMP CBar::put_ForeColor(OLE_COLOR val)
{
	bpV1.m_ocForeground=val;
	CacheColors();
	ViewChanged();
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::get_HighlightColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_ocHighLight;
	return NOERROR;
}
STDMETHODIMP CBar::put_HighlightColor(OLE_COLOR val)
{
	bpV1.m_ocHighLight=val;
	CacheColors();
	ViewChanged();
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::get_ShadowColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_ocShadow;
	return NOERROR;
}
STDMETHODIMP CBar::put_ShadowColor(OLE_COLOR val)
{
	bpV1.m_ocShadow=val;
	CacheColors();
	ViewChanged();
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::Refresh()
{
	switch (m_eAppType)
	{
	case eMDIForm:
		break;

	default:
		if (!IsWindow(m_hWnd))
			return E_FAIL;
		InvalidateRect(m_hWnd, NULL, FALSE);
		break;
	}
	if (m_pDockMgr)
		m_pDockMgr->Invalidate(FALSE);
	if (m_pBands)
		m_pBands->RefreshFloating();
	return NOERROR;
}
STDMETHODIMP CBar::get_ControlFont(LPFONTDISP *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=m_fhControl.GetFontDispatch();
	return NOERROR;
}

STDMETHODIMP CBar::putref_ControlFont(LPFONTDISP *val)
{
	return put_ControlFont(*val);
}

STDMETHODIMP CBar::put_ControlFont(IFontDisp * val)
{
	m_fhControl.InitializeFont(NULL, val);
	if (m_cachedControlFont)
		m_cachedControlFont=NULL;
	OnFontHeightChanged();
	ViewChanged();
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::OnSysColorChanged()
{
	CacheColors();
	GetGlobals().Reset();
	// update font too and call recalclayout
	ResetMenuFonts();
	HDC hDCTest = GetDC(NULL);
	if (hDCTest)
	{
		GetGlobals().m_nBitDepth = GetDeviceCaps(hDCTest, BITSPIXEL);
		ReleaseDC(NULL, hDCTest);
	}
	if (m_hFontToolTip)
	{
		BOOL bResult = DeleteFont(m_hFontToolTip);
		assert(bResult);
		m_hFontToolTip = NULL;
	}
	InternalRecalcLayout();
	Refresh();
	return NOERROR;
}
STDMETHODIMP CBar::get_ChildBandFont(LPFONTDISP* retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = m_fhChildBand.GetFontDispatch();
	return NOERROR;
}
STDMETHODIMP CBar::putref_ChildBandFont(LPFONTDISP* val)
{
	return put_ChildBandFont(*val);
}
STDMETHODIMP CBar::put_ChildBandFont(IFontDisp* val)
{
	m_fhChildBand.InitializeFont(NULL, val);
	if (m_hFontChildBandVert)
	{
		BOOL bResult = DeleteFont(m_hFontChildBandVert);
		assert(bResult);
		m_hFontChildBandVert = NULL;
	}
	if (m_hFontChildBand)
		m_hFontChildBand = NULL;
	OnFontHeightChanged();
	ViewChanged();
	SetModified();
	return NOERROR;
}

STDMETHODIMP CBar::PlaySound(BSTR bstrSound, SoundTypes stType)
{
	PLAYSOUNDFUNTION pFunction = GetGlobals().GetPlaySound();
	if (NULL == pFunction)
		return E_FAIL;

	DWORD dwSoundType;
	if (ddSTSystem == stType)
		dwSoundType = SND_ASYNC | SND_NODEFAULT;
	else
		dwSoundType = SND_FILENAME | SND_ASYNC | SND_NODEFAULT;

	MAKE_TCHARPTR_FROMWIDE(szSound, bstrSound);
	BOOL bResult = (pFunction)(szSound, NULL, dwSoundType);
	if (!bResult)
	{
		if (ddSTSystem == stType)
			return NOERROR;
		else
			return E_FAIL;
	}
	return NOERROR;
}
STDMETHODIMP CBar::get_MenuAnimation(MenuStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = bpV1.m_msMenuStyle;
	return NOERROR;
}
STDMETHODIMP CBar::put_MenuAnimation(MenuStyles val)
{
	switch (val)
	{
	case ddMSAnimateUnfold:
	case ddMSAnimateSlide:
	case ddMSAnimateNone:
	case ddMSAnimateRandom:
		bpV1.m_msMenuStyle = val;
		break;
	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}
STDMETHODIMP CBar::get_LargeIcons(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_vbLargeIcons;
	return NOERROR;
}
STDMETHODIMP CBar::put_LargeIcons(VARIANT_BOOL val)
{
	if (bpV1.m_vbLargeIcons != val)
	{
		bpV1.m_vbLargeIcons = val;
		RecalcLayout();
	}
	return NOERROR;
}
STDMETHODIMP CBar::get_ImageManager(VARIANT *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	if (!AmbientUserMode())
		return E_FAIL;
	m_pImageMgr->AddRef();
	retval->vt = VT_UNKNOWN;
	retval->punkVal = m_pImageMgr;
	return NOERROR;
}
STDMETHODIMP CBar::put_ImageManager(VARIANT val)
{
	if (VT_UNKNOWN != val.vt)
		return E_FAIL;	

	if (m_pImageMgr)
		m_pImageMgr->Release();

	m_pImageMgr = (CImageMgr*)val.pdispVal;
	return NOERROR;
}
STDMETHODIMP CBar::get_WhatsThisHelpMode(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	if (!AmbientUserMode())
		return E_FAIL;
	*retval = bpV1.m_vbWhatsThisHelpMode;
	return NOERROR;
}
STDMETHODIMP CBar::put_WhatsThisHelpMode(VARIANT_BOOL val)
{
	if (val != bpV1.m_vbWhatsThisHelpMode)
	{
		bpV1.m_vbWhatsThisHelpMode = val;
		if (VARIANT_TRUE == bpV1.m_vbWhatsThisHelpMode)
		{
			EnterWhatsThisHelpMode(GetDockWindow());
			bpV1.m_vbWhatsThisHelpMode = VARIANT_FALSE;
		}
	}
	return NOERROR;
}
STDMETHODIMP CBar::get_AutoSizeChildren(AutoSizeChildrenTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (eMDIForm == m_eAppType)
	{
		if (!AmbientUserMode())
			return E_FAIL;
	
		m_theErrorObject.SendError(IDERR_ONLYFORUSEWITHSDIAPPLICATIONS, NULL);
		return CUSTOM_CTL_SCODE(IDERR_ONLYFORUSEWITHSDIAPPLICATIONS);
	}

	*retval = bpV1.m_asChildren;
	return NOERROR;
}
STDMETHODIMP CBar::put_AutoSizeChildren(AutoSizeChildrenTypes val)
{
    switch (val)
	{
	case ddASNone:
	case ddASProportional:
	case ddASClientArea:
		bpV1.m_asChildren = val;
		if (ddASNone != bpV1.m_asChildren && NULL == m_pRelocate)
			m_pRelocate = new CRelocate(this);
		if (IsWindow(m_hWnd) && IsWindowVisible(GetParentWindow()))
			::SendMessage(GetParent(m_hWnd), WM_SIZE, 0, 0);
		break;

	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}
STDMETHODIMP CBar::get_AlignToForm(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	if (eMDIForm == m_eAppType && !AmbientUserMode())
		return E_FAIL;
	*retval = bpV1.m_vbAlignToForm;
	return NOERROR;
}
STDMETHODIMP CBar::put_AlignToForm(VARIANT_BOOL val)
{
	if (eMDIForm == m_eAppType)
	{
		if (!AmbientUserMode())
			return E_FAIL;
	
		m_theErrorObject.SendError(IDERR_ONLYFORUSEWITHSDIAPPLICATIONS, NULL);
		return CUSTOM_CTL_SCODE(IDERR_ONLYFORUSEWITHSDIAPPLICATIONS);
	}
	Detach();
	bpV1.m_vbAlignToForm = val;
	Attach(NULL);
	return NOERROR;
}

//
// LocaleString
//

STDMETHODIMP CBar::Localize(LocalizationTypes Index, BSTR LocaleString)
{
	m_pLocalizer->Localize(Index,LocaleString);
	return NOERROR;
}
STDMETHODIMP CBar::RegisterChildMenu( OLE_HANDLE hWndChild,  BSTR strChildMenuName)
{
	if (!m_theChildMenus.RegisterChildWindow((HWND)hWndChild, strChildMenuName))
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CBar::get_Version(BSTR *retval)
{
	CVersionInfo viFileVersion(g_hInstance, _T("actbar2.ocx"));
	LPTSTR szVersion = (LPTSTR)viFileVersion.GetValue(_T("FileVersion"));
	if (NULL == szVersion)
		return E_FAIL;
	// remove space, convert comma to period
	TCHAR *ptr = szVersion;
	while (*ptr)
	{
		if (*ptr==',') 
			*ptr='.';
		else if (*ptr==32)
			lstrcpy(ptr,ptr+1);
		else
			++ptr;
	}
	DDString strVersion = szVersion;
	*retval = strVersion.AllocSysString();
	return NOERROR;
}
STDMETHODIMP CBar::put_ActiveBand(Band * val)
{
	m_pActiveBand = (CBand*)val;
	return NOERROR;
}
STDMETHODIMP CBar::get_AutoUpdateStatusBar(VARIANT_BOOL *retval)
{
	*retval = bpV1.m_vbAutoUpdateStatusbar;
	return NOERROR;
}
STDMETHODIMP CBar::put_AutoUpdateStatusBar(VARIANT_BOOL val)
{
	bpV1.m_vbAutoUpdateStatusbar = val;
	return NOERROR;
}

//
// SaveMenuUsageData
// 

STDMETHODIMP CBar::SaveMenuUsageData( BSTR bstrFileName,  SaveOptionTypes nOptions,  VARIANT *vData)
{
	HRESULT hResult = NOERROR;
	if (ddSOByteArray & nOptions)
	{
		VariantInit(vData);
		vData->vt = (VT_ARRAY|VT_UI1);

		IStream* pStream;
		HRESULT hResult = CreateStreamOnHGlobal(NULL, TRUE, &pStream);
		if (FAILED(hResult))
			return hResult;
		
		HGLOBAL hGlobal;
		hResult = GetHGlobalFromStream(pStream, &hGlobal);
		if (FAILED(hResult))
		{
			pStream->Release();
			return hResult;
		}

		hResult = m_pBands->ExchangeMenuUsageData(pStream, VARIANT_TRUE);
		
		LPVOID pData = GlobalLock(hGlobal);
		if (NULL == pData)
		{
			pStream->Release();
			return hResult;
		}

		DWORD dwSize = GlobalSize(hGlobal);

		hResult = SafeArrayAllocDescriptor(1, &vData->parray);
		if (FAILED(hResult))
		{
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}
		vData->parray->fFeatures = 0;
		vData->parray->cbElements = sizeof(BYTE);
		vData->parray->rgsabound[0].cElements = dwSize;
		vData->parray->rgsabound[0].lLbound = 0;
		hResult = SafeArrayAllocData(vData->parray);
		if (FAILED(hResult))
		{
			SafeArrayDestroyDescriptor(vData->parray);
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}
		LPBYTE pArg;
		hResult = SafeArrayAccessData(vData->parray, (void**)&pArg);
		if (FAILED(hResult))
		{
			SafeArrayUnaccessData(vData->parray);
			SafeArrayDestroyDescriptor(vData->parray);
			GlobalUnlock(hGlobal);
			pStream->Release();
			return hResult;
		}
		memcpy (pArg, pData, dwSize);
		SafeArrayUnaccessData(vData->parray);
		GlobalUnlock(hGlobal);
		pStream->Release();
	}
	else if (ddSOFile & nOptions)
	{
		TCHAR szTempFilePath[MAX_PATH];
		UINT nResult = GetTempPath(MAX_PATH, szTempFilePath);
		if (0 == nResult)
			return E_FAIL;

		TCHAR szTempFile[MAX_PATH];
		nResult = GetTempFileName(szTempFilePath, _T("ab2"), 0, szTempFile);
		if (0 == nResult)
			return E_FAIL;

		OFSTRUCT of;
		of.cBytes = sizeof(OFSTRUCT);

		HANDLE hFile = CreateFile(szTempFile, 
								  GENERIC_WRITE, 
								  FILE_SHARE_READ, 
								  NULL, 
								  OPEN_ALWAYS, 
								  FILE_ATTRIBUTE_NORMAL, 
								  NULL);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			DDString strError;
			strError.Format(_T("%s"), GetLastErrorMsg());
			MAKE_WIDEPTR_FROMTCHAR(wError, strError);
			m_theErrorObject.SendError(IDERR_CREATEFILE, wError);
			return CUSTOM_CTL_SCODE(IDERR_CREATEFILE);
		}

		FileStream* pFileStream = new FileStream;
		if (NULL == pFileStream)
		{
			CloseHandle((HANDLE)hFile);
			DeleteFile(szTempFile);
			return E_OUTOFMEMORY;
		}

		pFileStream->SetHandle(hFile, TRUE);

		hResult = m_pBands->ExchangeMenuUsageData(pFileStream, VARIANT_TRUE);
		if (SUCCEEDED(hResult))
		{
			MAKE_TCHARPTR_FROMWIDE(szFileName, bstrFileName);
			DWORD dwResult = CopyFile(szTempFile, szFileName, FALSE);
		}
		pFileStream->Release();
		DeleteFile(szTempFile);
	}
	else
		return E_FAIL;
	return hResult;
}

//
// LoadMenuUsageData
// 

STDMETHODIMP CBar::LoadMenuUsageData( VARIANT *vData,  SaveOptionTypes nOptions)
{
	HRESULT hResult;
	if (ddSOByteArray & nOptions)
	{
		SAFEARRAY* psa;
		if (vData->vt == (VT_ARRAY|VT_UI1))
			psa = vData->parray;
		else if (vData->vt == (VT_BYREF|VT_ARRAY|VT_UI1))
		{
			if (NULL == vData->pparray)
				return E_FAIL;
			psa = *vData->pparray;
		}
		else
			return E_FAIL;
		
		if (NULL == psa)
			return E_FAIL;

		LPBYTE pArg;
		SafeArrayAccessData(vData->parray, (void**)&pArg);

		HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, psa->rgsabound[0].cElements);
		if (NULL == hGlobal)
		{
			SafeArrayUnaccessData(psa);
			return E_FAIL;
		}

		LPVOID pData = GlobalLock(hGlobal);
		if (NULL == pData)
		{
			GlobalFree(hGlobal);
			GlobalUnlock(hGlobal);
			SafeArrayUnaccessData(psa);
			return E_FAIL;
		}

		DWORD dwSize = GlobalSize(hGlobal);

		memcpy(pData, pArg, dwSize);
		SafeArrayUnaccessData(psa);
		GlobalUnlock(hGlobal);

		IStream* pStream;
		hResult = CreateStreamOnHGlobal(hGlobal, TRUE, &pStream);
		if (FAILED(hResult))
			return hResult;
		
		LARGE_INTEGER pos;
		pos.LowPart = pos.HighPart = 0;
		pStream->Seek(pos, STREAM_SEEK_SET, 0);

		hResult = m_pBands->ExchangeMenuUsageData(pStream, VARIANT_FALSE);
		
		pStream->Release();
	}
	else if (ddSOFile & nOptions)
	{
		OFSTRUCT of;
		of.cBytes = sizeof(OFSTRUCT);

		BSTR bstrFile = NULL;
		if (vData->vt == VT_BSTR)
			bstrFile = vData->bstrVal;
		else if (vData->vt == (VT_BYREF|VT_BSTR))
			bstrFile = *vData->pbstrVal;
		else
			return E_INVALIDARG;

		MAKE_TCHARPTR_FROMWIDE(szFileName, bstrFile);
		HANDLE hFile = CreateFile(szFileName, 
								  GENERIC_READ, 
								  FILE_SHARE_READ, 
								  NULL, 
								  OPEN_EXISTING, 
								  FILE_ATTRIBUTE_NORMAL, 
								  NULL);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			DDString strError;
			strError.Format(_T("%s"), GetLastErrorMsg());
			MAKE_WIDEPTR_FROMTCHAR(wError, strError);
			m_theErrorObject.SendError(IDERR_OPENFILE, wError);
			return CUSTOM_CTL_SCODE(IDERR_OPENFILE);
		}

		FileStream* pFileStream = new FileStream;
		if (NULL == pFileStream)
		{
			CloseHandle((HANDLE)hFile);
			return E_OUTOFMEMORY;
		}
		pFileStream->SetHandle(hFile, TRUE);

		hResult = m_pBands->ExchangeMenuUsageData(pFileStream, VARIANT_FALSE);

		pFileStream->Release();
		if (S_FALSE == hResult)
		{
			m_theErrorObject.m_nAsyncError++;
			m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
			m_theErrorObject.m_nAsyncError--;
			return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
		}
	}
	else
		return E_FAIL;
	return hResult;
}

//
// ClearMenuUsageData
// 

STDMETHODIMP CBar::ClearMenuUsageData()
{
	return m_pBands->ClearMenuUsageData();
}
STDMETHODIMP CBar::get_PersonalizedMenus(PersonalizedMenuTypes *retval)
{
	*retval = bpV1.m_pmMenus;
	return NOERROR;
}
STDMETHODIMP CBar::put_PersonalizedMenus(PersonalizedMenuTypes val)
{
	bpV1.m_pmMenus = val;
	return NOERROR;
}
STDMETHODIMP CBar::put_ClientAreaControl(LPDISPATCH val)
{
	if (NULL == val)
		return E_INVALIDARG;

	if (eMDIForm == m_eAppType)
	{
		if (!AmbientUserMode())
			return E_FAIL;
	
		m_theErrorObject.SendError(IDERR_ONLYFORUSEWITHSDIAPPLICATIONS, NULL);
		return CUSTOM_CTL_SCODE(IDERR_ONLYFORUSEWITHSDIAPPLICATIONS);
	}

	if (NULL == m_pRelocate)
	{
		m_pRelocate = new CRelocate(this);
		if (NULL == m_pRelocate)
			return E_OUTOFMEMORY;
	}
	HRESULT hResult = m_pRelocate->AddControl(val);
	if (IsWindow(m_hWnd))
		::SendMessage(GetParent(m_hWnd), WM_SIZE, 0, 0);
	return hResult;

}
STDMETHODIMP CBar::putref_ClientAreaControl(LPDISPATCH *val)
{
	if (NULL == val)
		return E_INVALIDARG;
	return put_ClientAreaControl(*val);
}
STDMETHODIMP CBar::get_UserDefinedCustomization(VARIANT_BOOL *retval)
{
	*retval = bpV1.m_vbUserDefinedCustomization;
	return NOERROR;
}
STDMETHODIMP CBar::put_UserDefinedCustomization(VARIANT_BOOL val)
{
	bpV1.m_vbUserDefinedCustomization = val;
	return NOERROR;
}
STDMETHODIMP CBar::GetToolFromPosition( long x,  long y,  Tool** tool)
{
	*tool = NULL;
	CBand* pBand;
	HRESULT hResult = GetBandFromPosition(x, y, (Band**)&pBand);
	if (SUCCEEDED(hResult))
	{
		SIZE size = {x, y};
		TwipsToPixel(&size, &size);
		POINT pt = {size.cx, size.cy};

		CRect rcBand;
		HWND hWnd;
		if (!pBand->GetBandRect(hWnd, rcBand))
		{
			pBand->Release();
			return NOERROR;
		}

		if (!ScreenToClient(hWnd, &pt))
		{
			pBand->Release();
			return NOERROR;
		}

		pt.x -= rcBand.left;
		pt.y -= rcBand.top;

		int nToolIndex;
		CBand* pBandHitTest;
		*tool = (Tool*)pBand->HitTestTools(pBandHitTest, pt, nToolIndex);
		pBand->Release();
		if (*tool)
		{
			((CTool*)(*tool))->AddRef();
			return NOERROR;
		}
	}
	return hResult;
}
STDMETHODIMP CBar::GetBandFromPosition( long x,  long y,  Band** band)
{
	*band = NULL;
	SIZE size = {x, y};
	TwipsToPixel(&size, &size);
	POINT pt = {size.cx, size.cy};
	CDock* pDock = m_pDockMgr->HitTest(pt);
	if (pDock)
	{
		ScreenToClient(pDock->hWnd(), &pt);
		*band = (Band*)pDock->HitTest(pt);
		if (*band)
			((CBand*)(*band))->AddRef();
	}
	else
	{
		CRect rcWindow;
		CBand* pBand;
		int nCount = m_pBands->GetBandCount();
		for (int nBand = 0; nBand < nCount; nBand++)
		{
			pBand = m_pBands->GetBand(nBand);
			if (ddDAFloat == pBand->bpV1.m_daDockingArea && pBand->m_pFloat && pBand->m_pFloat->IsWindow())
			{
				if (pBand->m_pFloat->GetWindowRect(rcWindow) && PtInRect(&rcWindow, pt))
				{
					*band = (Band*)pBand;
					pBand->AddRef();
					break;
				}
			}
			else if (ddDAPopup == pBand->bpV1.m_daDockingArea && pBand->m_pPopupWin && pBand->m_pPopupWin->IsWindow())
			{
				if (pBand->m_pPopupWin->GetWindowRect(rcWindow) && PtInRect(&rcWindow, pt))
				{
					*band = (Band*)pBand;
					pBand->AddRef();
					break;
				}
			}
		}
	}
	return NOERROR;
}
STDMETHODIMP CBar::get_FireDblClickEvent(VARIANT_BOOL *retval)
{
	*retval = bpV1.m_vbFireDblClickEvent;
	return NOERROR;
}
STDMETHODIMP CBar::put_FireDblClickEvent(VARIANT_BOOL val)
{
	bpV1.m_vbFireDblClickEvent = val;
	return NOERROR;
}
STDMETHODIMP CBar::get_ThreeDLight(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = bpV1.m_oc3DLight;
	return NOERROR;
}
STDMETHODIMP CBar::put_ThreeDLight(OLE_COLOR val)
{
	bpV1.m_oc3DLight=val;
	CacheColors();
	SetModified();
	return NOERROR;
}
STDMETHODIMP CBar::get_ThreeDDarkShadow(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_oc3DDarkShadow;
	return NOERROR;
}
STDMETHODIMP CBar::put_ThreeDDarkShadow(OLE_COLOR val)
{
	bpV1.m_oc3DDarkShadow = val;
	CacheColors();
	SetModified();
	return NOERROR;
}

//
// ApplyAll
//

struct CApplyAll
{
	CTool* pTool;
	static void Change(CBand* pBand, CTool* pTool, void* pData);
};

//
// Change
//

void CApplyAll::Change(CBand* pBand, CTool* pTool, void* pData)
{
	CApplyAll* pChangeTool = (CApplyAll*)pData;
	if (pChangeTool)
		pChangeTool->pTool->CopyTo((ITool**)&pTool);
}

STDMETHODIMP CBar::ApplyAll( Tool *tool)
{
	CApplyAll theAll;
	theAll.pTool = (CTool*)tool;
	ModifyTool(theAll.pTool->tpV1.m_nToolId, &theAll, CApplyAll::Change);
	return NOERROR;
}

STDMETHODIMP CBar::get_BackgroundOption(short *retval)
{
	*retval = (BackgroundOptionsTypes)bpV1.m_dwBackgroundOptions;
	return NOERROR;
}
STDMETHODIMP CBar::put_BackgroundOption(short val)
{
	bpV1.m_dwBackgroundOptions = val;
	return NOERROR;
}
STDMETHODIMP CBar::get_ClientAreaLeft(long *retval)
{
	if (!IsWindow(m_hWnd) && !(eSDIForm == m_eAppType || eClientArea == m_eAppType))
		return E_FAIL;

	CRect rcClient;
	if (GetClientRect(m_hWnd, &rcClient))
	{
		SIZE size = {rcClient.left, 0};
		if (m_pDockMgr)
			size.cx += m_pDockMgr->GetDockRect(ddDALeft).Width();
		PixelToTwips(&size, &size);
		*retval = size.cx;
		return NOERROR;
	}
	return E_FAIL;
}
STDMETHODIMP CBar::get_ClientAreaTop(long *retval)
{
	if (!IsWindow(m_hWnd) && !(eSDIForm == m_eAppType || eClientArea == m_eAppType))
		return E_FAIL;

	CRect rcClient;
	if (GetClientRect(m_hWnd, &rcClient))
	{
		SIZE size = {rcClient.top, 0};
		if (m_pDockMgr)
			size.cx += m_pDockMgr->GetDockRect(ddDATop).Height();
		PixelToTwips(&size, &size);
		*retval = size.cx;
		return NOERROR;
	}
	return E_FAIL;
}
STDMETHODIMP CBar::get_ClientAreaWidth(long *retval)
{
	if (!IsWindow(m_hWnd) && !(eSDIForm == m_eAppType || eClientArea == m_eAppType))
		return E_FAIL;

	CRect rcClient;
	if (GetClientRect(m_hWnd, &rcClient))
	{
		SIZE size = {rcClient.Width(), 0};
		if (m_pDockMgr)
			size.cx -= m_pDockMgr->GetDockRect(ddDALeft).Width() + m_pDockMgr->GetDockRect(ddDARight).Width();
		PixelToTwips(&size, &size);
		*retval = size.cx;
		return NOERROR;
	}
	return E_FAIL;
}
STDMETHODIMP CBar::get_ClientAreaHeight(long *retval)
{
	if (!IsWindow(m_hWnd) && !(eSDIForm == m_eAppType || eClientArea == m_eAppType))
		return E_FAIL;

	CRect rcClient;
	if (GetClientRect(m_hWnd, &rcClient))
	{
		SIZE size = {rcClient.Height(), 0};
		if (m_pDockMgr)
			size.cx -= m_pDockMgr->GetDockRect(ddDATop).Height() + m_pDockMgr->GetDockRect(ddDABottom).Height();
		PixelToTwips(&size, &size);
		*retval = size.cx;
		return NOERROR;
	}
	return E_FAIL;
}
STDMETHODIMP CBar::get_SDIChildWindowClass(BSTR *retval)
{
	if (!AmbientUserMode())
		return E_FAIL;
	*retval = SysAllocString(m_bstrWindowClass);
	return NOERROR;
}
STDMETHODIMP CBar::put_SDIChildWindowClass(BSTR val)
{
	SysFreeString(m_bstrWindowClass);
	m_bstrWindowClass = SysAllocString(val);
	return NOERROR;
}
STDMETHODIMP CBar::put_ClientAreaHWnd(OLE_HANDLE val)
{
	if (NULL == val)
		return E_INVALIDARG;

	if (eMDIForm == m_eAppType)
	{
		if (!AmbientUserMode())
			return E_FAIL;
	
		m_theErrorObject.SendError(IDERR_ONLYFORUSEWITHSDIAPPLICATIONS, NULL);
		return CUSTOM_CTL_SCODE(IDERR_ONLYFORUSEWITHSDIAPPLICATIONS);
	}

	if (NULL == m_pRelocate)
	{
		m_pRelocate = new CRelocate(this);
		if (NULL == m_pRelocate)
			return E_OUTOFMEMORY;
	}
	return m_pRelocate->AddControl((HWND)val);
}
STDMETHODIMP CBar::get_UseUnicode(VARIANT_BOOL *retval)
{
	*retval = bpV1.m_vbUseUnicode;
	return NOERROR;
}
STDMETHODIMP CBar::put_UseUnicode(VARIANT_BOOL val)
{
	bpV1.m_vbUseUnicode = val;
	return NOERROR;
}
STDMETHODIMP CBar::get_XPLook(VARIANT_BOOL *retval)
{
	*retval = bpV1.m_vbXPLook;
	return NOERROR;
}
STDMETHODIMP CBar::put_XPLook(VARIANT_BOOL val)
{
	bpV1.m_vbXPLook = val;
	if (!AmbientUserMode())
		RecalcLayout();
	return NOERROR;
}
