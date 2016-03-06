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
#include "Dock.h"
#include "Tool.h"
#include "Bands.h"
#include "MiniWin.h"
#include "Band.h"
#include "Designer\DesignerInterfaces.h"
#include "WindowProc.h"
#include "PopupWin.h"
#include "Bar.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// IBarPrivate members
//

STDMETHODIMP CBar::RegisterBandChange(IDesignerNotifications* pDesignerNotify)
{
	m_pDesignerNotify = pDesignerNotify;
	return NOERROR;
}
STDMETHODIMP CBar::RevokeBandChange()
{
	m_pDesignerNotify = NULL;
	return NOERROR;
}
STDMETHODIMP CBar::DesignerInitialize(IDragDropManager* pDragDropManager, IDesignerNotify* pDesigner)
{
	HRESULT hResult = E_FAIL;
	try
	{
		if (NULL == pDragDropManager && NULL == pDesigner)
			return E_FAIL;

		//
		// Set up ActiveBar so it can interact with the designer
		//

		assert(pDesigner);
		m_pDesigner = pDesigner;

		assert(pDragDropManager);
		m_pDragDropManager = pDragDropManager;

		assert(m_pDockMgr);
		m_pDockMgr->RegisterDragDrop(m_pDragDropManager);

		hResult = InPlaceActivate(OLEIVERB_UIACTIVATE);
		if (SUCCEEDED(hResult))
		{
			if (!m_bCustomization)
				DoCustomization(TRUE);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return hResult;
}
STDMETHODIMP CBar::DesignerShutdown()
{
	try
	{
		int nBand = 0;
		m_bShutdownLock = TRUE;

		//
		// Destroy Float Windows
		//

		HANDLE* pHandles = NULL;
		CBand* pBand;
		TypedArray<HWND> aHandles;
		int nCount = m_pBands->GetBandCount();
		for (nBand = 0; nBand < nCount; nBand++)
		{
			try
			{
				pBand = m_pBands->GetBand(nBand);
				if (pBand && pBand->m_pFloat)
				{
					if (pBand->m_pFloat->IsWindow())
					{
						pBand->m_pFloat->DestroyWindow();
						aHandles.Add(pBand->m_pFloat->hWnd());
					}
					else
						delete pBand->m_pFloat;
				}
			}
			catch (...)
			{
			}
		}

		nCount = aHandles.GetSize();
		if (nCount > 0)
		{
			pHandles = new HANDLE[nCount];
			if (pHandles)
			{
				for (nBand = 0; nBand < nCount; nBand++)
					pHandles[nBand] = aHandles.GetAt(nBand);
				WaitForMultipleObjects(nCount, pHandles, TRUE, INFINITE);
				delete [] pHandles;
			}
		}

		DestroyPopups();

		if (m_pToolShadow && m_pToolShadow->IsWindow())
		{
			::PostMessage(m_hWnd, GetGlobals().WM_KILLWINDOW, (WPARAM)m_pToolShadow->hWnd(), 0);
			::PostMessage(m_hWnd, GetGlobals().WM_KILLWINDOW, (WPARAM)m_pToolShadow->GetShadowWindow(), 0);
		}

		if (IsWindow(m_hWndDesignerShutDownAdvise))
			PostMessage(m_hWndDesignerShutDownAdvise, WM_COMMAND, MAKELONG(1234,0), 0);

		if (m_bCustomization)
		{
			//
			// Shutdown Customization
			//

			DoCustomization(FALSE, FALSE);
			UIDeactivate();
		}

		//
		// Destroy Float Windows
		//

		assert(m_pDockMgr);
		if (m_pDragDropManager && m_pDockMgr)
		{
			m_pDockMgr->RevokeDragDrop(m_pDragDropManager);
			m_pDragDropManager = NULL;
		}
		
		if (m_pDesigner)
			m_pDesigner = NULL;

		if (m_pDesignerImpl)
		{
			m_pDesignerImpl->Release();
			m_pDesignerImpl = NULL;
		}
		HWND hWndDock=GetDockWindow();
		if (hWndDock)
		{
			EnableWindow(hWndDock, TRUE);
			::SetFocus(hWndDock);
		}
		if (IsWindow(m_hWnd))
		{
			PostMessage(m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
			InvalidateRect(m_hWnd, NULL, FALSE);
		}
   	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	m_bShutdownLock = FALSE;
	return NOERROR;
}
STDMETHODIMP CBar::SetDesignerModified()
{
	SetModified();
	RecalcLayout();
	return NOERROR;
}
STDMETHODIMP CBar::ExchangeToolById(IUnknown* pStreamIn, VARIANT_BOOL vbSave, IDispatch** ppToolIn)
{
	ITool** ppTool = (ITool**)ppToolIn;
	IStream* pStream = (IStream*)pStreamIn;
	HRESULT hResult = NOERROR;
	long nId;
	if (VARIANT_TRUE == vbSave)
	{
		hResult = (*ppTool)->get_ID(&nId);
		if (FAILED(hResult))
			return hResult;
		hResult = pStream->Write(&nId, sizeof(nId), NULL);
		if (FAILED(hResult))
			return hResult;
	}
	else
	{
		*ppTool = NULL;
		hResult = pStream->Read(&nId, sizeof(nId), NULL);
		if (FAILED(hResult))
			return hResult;
		hResult = m_pTools->ItemById(nId, (Tool**)ppTool);
		if (FAILED(hResult))
			return hResult;
	}
	return hResult;
}
STDMETHODIMP CBar::ExchangeToolByBandChildBandToolId(IUnknown* pStreamIn, BSTR bstrBand, BSTR bstrChildBand, VARIANT_BOOL vbSave, IDispatch** ppToolIn)
{
	ITool** ppTool = (ITool**)ppToolIn;
	HRESULT hResult = NOERROR;
	long nId;
	IStream* pStream = (IStream*)pStreamIn;
	if (VARIANT_TRUE == vbSave)
	{
		hResult = StWriteBSTR(pStream, bstrBand);
		if (FAILED(hResult))
			return hResult;
		hResult = StWriteBSTR(pStream, bstrChildBand);
		if (FAILED(hResult))
			return hResult;
		hResult = (*ppTool)->get_ID(&nId);
		if (FAILED(hResult))
			return hResult;
		hResult = pStream->Write(&nId, sizeof(nId), NULL);
		if (FAILED(hResult))
			return hResult;
	}
	else
	{
		*ppTool = NULL;
		hResult = StReadBSTR(pStream, bstrBand);
		if (FAILED(hResult))
			return hResult;

		hResult = StReadBSTR(pStream, bstrChildBand);
		if (FAILED(hResult))
			return hResult;

		hResult = pStream->Read(&nId, sizeof(nId), NULL);
		if (FAILED(hResult))
			return hResult;

		IChildBands* pChildBands = NULL;
		IBand*  pBand = NULL;
		IBand*  pChildBand = NULL;
		ITools* pTools = NULL;
		ITool*  pTool = NULL;
		VARIANT vIndex;
		vIndex.vt = VT_BSTR;

		vIndex.bstrVal = bstrBand;
		hResult = m_pBands->Item(&vIndex, (Band**)&pBand);
		if (FAILED(hResult))
			goto Cleanup;

		if (NULL == bstrChildBand || NULL == *bstrChildBand)
		{
			hResult = pBand->get_Tools((Tools**)&pTools);
			if (FAILED(hResult))
				goto Cleanup;
		}
		else
		{
			hResult = pBand->get_ChildBands((ChildBands**)&pChildBands);
			if (FAILED(hResult))
				goto Cleanup;

			vIndex.bstrVal = bstrChildBand;
			hResult = pChildBands->Item(&vIndex, (Band**)&pChildBand);
			if (FAILED(hResult))
				goto Cleanup;

			hResult = pChildBand->get_Tools((Tools**)&pTools);
			if (FAILED(hResult))
				goto Cleanup;
		}
		hResult = pTools->ItemById(nId, (Tool**)ppTool);
Cleanup:
		if (pBand)
			pBand->Release();
		if (pChildBands)
			pChildBands->Release();
		if (pChildBand)
			pChildBand->Release();
		if (pTools)
			pTools->Release();
	}
	return hResult;
}
STDMETHODIMP CBar::ShutDownAdvise( OLE_HANDLE hWndDesignerShutDownAdvise)
{
	m_hWndDesignerShutDownAdvise = (HWND)hWndDesignerShutDownAdvise;
	return NOERROR;
}
STDMETHODIMP CBar::get_Library(VARIANT_BOOL *retval)
{
	*retval = m_vbLibrary;
	return NOERROR;
}
STDMETHODIMP CBar::put_Library(VARIANT_BOOL val)
{
	m_vbLibrary = val;
	return NOERROR;
}
STDMETHODIMP CBar::put_CustomizeDragLock(LPDISPATCH pTool)
{
	m_pCustomizeDragLock = (CTool*)(Tool*)pTool;
	return NOERROR;
}

//
// VerifyImages
//

struct VerifyImages
{
	FMap m_mapImages;
	static void DoIt(CBand* pBand, CTool* pTool, void* pData);
};

void VerifyImages::DoIt(CBand* pBand, CTool* pTool, void* pData)
{
	if (pTool)
	{
		VerifyImages* pImages = (VerifyImages*)pData;
		for (int nImage = 0; nImage < CTool::eNumOfImageTypes; nImage++)
		{
			if (-1 == pTool->tpV1.m_nImageIds[nImage])
				continue;

			long* pnRefCount = (long*)1;
			if (pImages->m_mapImages.Lookup((LPVOID)pTool->tpV1.m_nImageIds[nImage], (LPVOID&)pnRefCount))
			{
				int nTemp = (int)pnRefCount;
				pnRefCount = (long*)++nTemp;
				pImages->m_mapImages.SetAt((LPVOID)pTool->tpV1.m_nImageIds[nImage], (LPVOID)pnRefCount);
			}
			else
				pImages->m_mapImages.SetAt((LPVOID)pTool->tpV1.m_nImageIds[nImage], (LPVOID)pnRefCount);
		}
	}
}

//
// VerifyAndCorrectImages
//

STDMETHODIMP CBar::VerifyAndCorrectImages()
{
	VerifyImages theImages;
	ModifyTool(-1, &theImages, VerifyImages::DoIt);

	if (m_pImageMgr)
	{
		long nImageId;
		ULONG nRefCount;
		ImageEntry* pImageEntry;
		FPOSITION posMap = m_pImageMgr->m_mapImages.GetStartPosition();
		while (posMap)
		{
			m_pImageMgr->m_mapImages.GetNextAssoc(posMap, (LPVOID&)nImageId, (LPVOID&)pImageEntry);
			if (theImages.m_mapImages.Lookup((LPVOID)pImageEntry->ImageId(), (LPVOID&)nRefCount) && pImageEntry->RefCount() != nRefCount)
				pImageEntry->RefCount() = nRefCount;
		}
	}
	return NOERROR;
}
//
// Attach
//

STDMETHODIMP CBar::Attach(OLE_HANDLE hWndParentIn)
{
	if (m_bStartedToolBars)
		return NOERROR;

	HWND hWndParent = (HWND)hWndParentIn;
	if (!IsWindow(hWndParent))
		hWndParent = GetDockWindow();
	else
	{
		m_fInPlaceActive = TRUE;
		m_bUserMode = 1;
		m_hWndDock = m_hWndParent = hWndParent;
		// Create an ActiveBar Window
		CRect rc;
		GetClientRect(hWndParent, &rc);
		m_Size = rc.Size();
		if (NULL == m_hWnd)
			CreateInPlaceWindow(0, 0, FALSE);
		else
			SetWindowPos(m_hWnd, NULL, rc.left, rc.top, rc.Width(), rc.Height(), SWP_SHOWWINDOW|SWP_NOACTIVATE);
	}

	if (!hWndParent)
		return E_FAIL;

	if (AmbientUserMode() && IsWindow(hWndParent))
		DragAcceptFiles(hWndParent, TRUE);

	if (GetAmbientProperty(DISPID_AMBIENT_BACKCOLOR, VT_I4, &m_ocAmbientBackColor))
		OleTranslateColor(m_ocAmbientBackColor, NULL, &m_crAmbientBackColor);

	void* pPrevBar;
	HWND hWndMdi;
	TCHAR szBuffer[_MAX_PATH];
	DWORD dwResult = GetModuleFileName(NULL, szBuffer, 255);

	if (dwResult > 0 && lstrlen(szBuffer) > 12 && 0 == lstrcmpi(szBuffer + lstrlen(szBuffer) - 12, _T("iexplore.exe")))
		m_bIE = TRUE;

	const int nLen = 15;
	TCHAR szClassName[nLen];
	if (AmbientUserMode())
	{
		hWndMdi = ::GetWindow(hWndParent, GW_CHILD);
		if (hWndMdi)
		{
			MAKE_TCHARPTR_FROMWIDE(szWindowClass, m_bstrWindowClass);
			GetClassName(hWndMdi, szClassName, nLen);
			if (0 == lstrcmpi(szClassName, szWindowClass))
			{
				//
				// Hook into the MDIClient Form
				//

				m_eAppType = eMDIForm;
				m_hWndMDIClient = hWndMdi;

				IOleControlSite* pControlSite;
				HRESULT hResult = m_pClientSite->QueryInterface(IID_IOleControlSite, (void**)&pControlSite);
				if (SUCCEEDED(hResult))
				{
					LPDISPATCH pDispatch;
					hResult = pControlSite->GetExtendedControl(&pDispatch);
					pControlSite->Release();
					if (SUCCEEDED(hResult))
					{
						VARIANT vWidth;
						vWidth.vt = VT_I4;
						vWidth.lVal = 0;
						hResult = PropertyPut(pDispatch, L"Width", vWidth);
						pDispatch->Release();
					}
				}

				if (AmbientUserMode())
					DragAcceptFiles(m_hWndMDIClient, TRUE);

				if (GetGlobals().m_pmapBar->Lookup((LPVOID)m_hWndMDIClient, pPrevBar))
				{
					TRACE2(1, _T("Window handle already entered Bar: %X hWnd: %X\n"), 
						   this, 
						   m_hWndMDIClient);
					return NOERROR;
				}
				GetGlobals().m_pmapBar->SetAt((LPVOID)m_hWndMDIClient,(LPVOID)this);

				m_pMDIClientProc = (WNDPROC)GetWindowLong(m_hWndMDIClient, GWL_WNDPROC);
				SetWindowLong(m_hWndMDIClient, GWL_WNDPROC, (LONG)MDIClientWindowProc);
				
				CacheSmButtonSize();
				
				CRect rcWin;
				::GetWindowRect(m_hWndMDIClient, &rcWin);
				m_nNewCustomizedFloatBandCounter = 0;
				m_ptNewCustomizedFloatBandPos.x = rcWin.left + 16;
				m_ptNewCustomizedFloatBandPos.y = rcWin.top + 16;
			}
		}
	}
	else
	{
		//
		// If we are in design mode check if we are on an MDI Form or not
		//

		GetClassName(hWndParent, szClassName, nLen);
		if (0 == lstrcmpi(szClassName, _T("ThunderMDIForm")))
			m_eAppType = eMDIForm;
	}

	if (eMDIForm == m_eAppType)
	{
		//
		// Check if two activebars are present
		//

		if (GetGlobals().m_pmapBar->Lookup((LPVOID)hWndParent, pPrevBar))
		{
			TRACE2(1, _T("Window handle already entered Bar: %X hWnd: %X\n"), 
				   this, 
				   hWndParent);
			return NOERROR;
		}

		//
		// Set the map and subclass the window procedure 
		//

		GetGlobals().m_pmapBar->SetAt((LPVOID)hWndParent,(LPVOID)this);
		TRACE2(1, _T("Window handle entered Bar: %X hWnd: %X\n"), 
			   this, 
			   hWndParent);

		m_pMainProc = (WNDPROC)GetWindowLong(hWndParent, GWL_WNDPROC);
		SetWindowLong (hWndParent, GWL_WNDPROC, (LONG)FrameWindowProc);
	}
	else
	{
		if (VARIANT_TRUE == bpV1.m_vbAlignToForm)
		{
			m_eAppType = eSDIForm;

			//
			// Check if two activebars are present
			//

			if (GetGlobals().m_pmapBar->Lookup((LPVOID)hWndParent, pPrevBar))
			{
				TRACE2(1, _T("Window handle already entered Bar: %X hWnd: %X\n"), 
					   this, 
					   hWndParent);
				return NOERROR;
			}

			//
			// Set the map and subclass the window procedure 
			//

			GetGlobals().m_pmapBar->SetAt((LPVOID)hWndParent,(LPVOID)this);
			TRACE2(1, _T("Window handle entered Bar: %X hWnd: %X\n"), 
				   this, 
				   hWndParent);

			m_pMainProc = (WNDPROC)GetWindowLong(hWndParent, GWL_WNDPROC);
			SetWindowLong (hWndParent, GWL_WNDPROC, (LONG)FormWindowProc);

			CRect rcClient;
			if (GetClientRect(hWndParent, &rcClient) && m_pInPlaceSite)
				m_pInPlaceSite->OnPosRectChange(&rcClient);
		}
		else
			m_eAppType = eClientArea;

		CRect rcWin;
		GetWindowRect(hWndParent, &rcWin);
		m_nNewCustomizedFloatBandCounter = 0;
		m_ptNewCustomizedFloatBandPos.x = rcWin.left + 64;
		m_ptNewCustomizedFloatBandPos.y = rcWin.top + 64;
	}
	LPVOID pTmp;
	if (!GetGlobals().m_pmapAccelator->Lookup((LPVOID)hWndParent, (void*&)pTmp))
		GetGlobals().m_pmapAccelator->SetAt(hWndParent, (LPVOID)this);
	m_bStartedToolBars = TRUE;
	return NOERROR;
}

//
// Detach
//

STDMETHODIMP CBar::Detach()
{
	if (!m_bStartedToolBars)
		return NOERROR;
	
	HWND hWndParent = GetDockWindow();
	if (NULL == hWndParent)
		return NOERROR;

	if (AmbientUserMode() && IsWindow(hWndParent))
		DragAcceptFiles(hWndParent, FALSE);

	BOOL bResult;
	if (m_pMainProc)
	{

		bResult = GetGlobals().m_pmapBar->RemoveKey(hWndParent);
		assert(bResult);
		SetWindowLong (hWndParent, GWL_WNDPROC, (long)m_pMainProc);
		m_pMainProc = NULL;
	}

	if (m_pMDIClientProc)
	{
		if (AmbientUserMode() && IsWindow(m_hWndMDIClient))
			DragAcceptFiles(m_hWndMDIClient, FALSE);

		bResult = GetGlobals().m_pmapBar->RemoveKey((LPVOID)m_hWndMDIClient);
		assert(bResult);

		SetWindowLong (m_hWndMDIClient, GWL_WNDPROC, (long)m_pMDIClientProc);
		m_pMDIClientProc = NULL;
	}
	bResult = GetGlobals().m_pmapAccelator->RemoveKey(GetDockWindow());
	m_bStartedToolBars = FALSE;
	return NOERROR;
}

STDMETHODIMP CBar::DeactivateWindow()
{
	BeforeDestroyWindow();
	return InPlaceDeactivate();
}
STDMETHODIMP CBar::get_PrivateHwnd(OLE_HANDLE *retval)
{
	*retval = (OLE_HANDLE)m_hWnd;
	return NOERROR;
}
STDMETHODIMP CBar::ExchangeToolByIdentity( IUnknown* pStreamIn,  VARIANT_BOOL vbSave, IDispatch**ppToolIn)
{
	CTool* pTool;
	IStream* pStream = (IStream*)pStreamIn;
	HRESULT hResult = NOERROR;
	long nId;
	if (VARIANT_TRUE == vbSave)
	{
		pTool = (CTool*)(ITool*)(*ppToolIn);
		hResult = pStream->Write(&pTool->tpV1.m_dwIdentity, sizeof(nId), NULL);
		if (FAILED(hResult))
			return hResult;
	}
	else
	{
		hResult = pStream->Read(&nId, sizeof(nId), NULL);
		if (FAILED(hResult))
			return hResult;
		hResult = m_pTools->ItemByIdentity(nId, &pTool);
		if (FAILED(hResult))
			return hResult;
		*ppToolIn = (ITool*)(pTool);
	}
	return hResult;
}
STDMETHODIMP CBar::ExchangeToolByBandChildBandToolIdentity(IUnknown* pStreamIn,  BSTR bstrBand, BSTR bstrChildBand,  VARIANT_BOOL vbSave, IDispatch** ppToolIn)
{
	CTool* pTool;
	HRESULT hResult = NOERROR;
	DWORD dwId;
	IStream* pStream = (IStream*)pStreamIn;
	if (VARIANT_TRUE == vbSave)
	{
		pTool = (CTool*)(ITool*)(*ppToolIn);
		hResult = StWriteBSTR(pStream, bstrBand);
		if (FAILED(hResult))
			return hResult;
		hResult = StWriteBSTR(pStream, bstrChildBand);
		if (FAILED(hResult))
			return hResult;
		hResult = pStream->Write(&pTool->tpV1.m_dwIdentity, sizeof(DWORD), NULL);
		if (FAILED(hResult))
			return hResult;
	}
	else
	{
		hResult = StReadBSTR(pStream, bstrBand);
		if (FAILED(hResult))
			return hResult;

		hResult = StReadBSTR(pStream, bstrChildBand);
		if (FAILED(hResult))
			return hResult;

		hResult = pStream->Read(&dwId, sizeof(dwId), NULL);
		if (FAILED(hResult))
			return hResult;

		IChildBands* pChildBands = NULL;
		IBand*  pBand = NULL;
		IBand*  pChildBand = NULL;
		ITools* pTools = NULL;
		VARIANT vIndex;
		vIndex.vt = VT_BSTR;

		vIndex.bstrVal = bstrBand;
		hResult = m_pBands->Item(&vIndex, (Band**)&pBand);
		if (FAILED(hResult))
			goto Cleanup;

		if (NULL == bstrChildBand || NULL == *bstrChildBand)
		{
			hResult = pBand->get_Tools((Tools**)&pTools);
			if (FAILED(hResult))
				goto Cleanup;
		}
		else
		{
			hResult = pBand->get_ChildBands((ChildBands**)&pChildBands);
			if (FAILED(hResult))
				goto Cleanup;

			vIndex.bstrVal = bstrChildBand;
			hResult = pChildBands->Item(&vIndex, (Band**)&pChildBand);
			if (FAILED(hResult))
				goto Cleanup;

			hResult = pChildBand->get_Tools((Tools**)&pTools);
			if (FAILED(hResult))
				goto Cleanup;
		}
		hResult = ((CTools*)pTools)->ItemByIdentity(dwId, &pTool);
		*ppToolIn = (ITool*)pTool;
Cleanup:
		if (pBand)
			pBand->Release();
		if (pChildBands)
			pChildBands->Release();
		if (pChildBand)
			pChildBand->Release();
		if (pTools)
			pTools->Release();
	}
	return hResult;
}
