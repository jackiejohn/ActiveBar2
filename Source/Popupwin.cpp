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
#include <LIMITS.H>
#include <MultiMon.h>
#include <winuser.h>
#include "Errors.h"
#include "IPServer.h"
#include "Support.h"
#include "Globals.h"
#include "Resource.h"
#include "Designer\DesignerInterfaces.h"
#include "Designer\DragDropMgr.h"
#include "Designer\DragDrop.h"
#include "Tool.h"
#include "ChildBands.h"
#include "Bands.h"
#include "Band.h"
#include "MiniWin.h"
#include "TPPopup.h"
#include "PopupWin.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// CPopupWinToolAndShadow
//

void CPopupWinToolAndShadow::Show(CTool* pTool, CRect& rcWindow)
{
	m_pTool = pTool;
	m_bVertical = m_pTool->m_pBand->IsVertical();
	if (!IsWindow())
	{
		if (!CreateWin(NULL, rcWindow))
			return;
	}
	BOOL bFlipVert = FALSE;
	BOOL bFlipHorz = FALSE;
	if (m_pTool->SubBand())
	{
		bFlipHorz = m_pTool->SubBand()->FlipHorz();
		bFlipVert = m_pTool->SubBand()->FlipVert();
	}
	CRect rcShadow = rcWindow;
	if (m_bVertical)
	{
		if (!bFlipHorz)
			rcShadow.right -= 3;
		rcShadow.Offset(3, 5);
	}
	else
	{
		if (!bFlipVert)
			rcShadow.bottom -= 2;
		rcShadow.Offset(4, 3);
	}
	m_theShadow.SetWindowPos(HWND_TOPMOST, rcShadow.left, rcShadow.top, rcShadow.Width(), rcShadow.Height(), SWP_SHOWWINDOW | SWP_NOACTIVATE);
	SetWindowPos(HWND_TOPMOST, rcWindow.left, rcWindow.top, rcWindow.Width(), rcWindow.Height(), SWP_SHOWWINDOW | SWP_NOACTIVATE);
}

BOOL CPopupWinToolAndShadow::CreateWin(HWND hWndParent, const CRect& rcWindow)
{
	m_bVertical = m_pTool->m_pBand->IsVertical();
	BOOL bFlipVert = FALSE;
	BOOL bFlipHorz = FALSE;
	if (m_pTool->SubBand())
	{
		bFlipHorz = m_pTool->SubBand()->FlipHorz();
		bFlipVert = m_pTool->SubBand()->FlipVert();
	}
	CRect rcShadow = rcWindow;
	if (m_bVertical)
	{
		if (!bFlipHorz)
			rcShadow.right -= 3;
		rcShadow.Offset(3, 5);
	}
	else
	{
		if (!bFlipVert)
			rcShadow.bottom -= 2;
		rcShadow.Offset(4, 3);
	}

	m_theShadow.CreateWin(hWndParent, rcShadow);

	CreateEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
			 _T(""),
			 WS_POPUP | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
			 rcWindow.left,
			 rcWindow.top,
			 rcWindow.Width(),
			 rcWindow.Height(),
			 NULL);
	if (!IsWindow())
		return FALSE;
	return TRUE;
}

void CPopupWinToolAndShadow::Hide(BOOL bSetNotPressed)
{
	TRACE(5, "CPopupWinToolAndShadow::Hide()\n");
	CRect rcWindow;
	if (IsWindow())
	{
		//ShowWindow(SW_HIDE);
		DestroyWindow();
	}
	//m_theShadow.Hide();
	m_theShadow.DestroyWindow();
	if (bSetNotPressed)
		m_pTool->SubBand()->InvalidateRect(NULL, TRUE);
}

LRESULT CPopupWinToolAndShadow::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
		case WM_MOUSEACTIVATE:
			return MA_NOACTIVATE;

		case WM_ACTIVATE:
			return FALSE;

		case WM_PAINT:
		{
			try
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcClient;
					GetClientRect(rcClient);

					if (NULL != m_pTool && NULL != m_pTool->m_pBand)
					{
						m_pTool->m_pBand->XPLookPopupWinInvalidateToolRect(m_hWnd, m_pTool, rcClient);
//						FillSolidRect(hDC, rcClient, m_pBar->m_crXPBandBackground);
//						m_pTool->Draw(hDC, rcClient, m_pTool->m_pBand->bpV1.m_btBands, m_bVertical, m_bVertical);
					}

					EndPaint(m_hWnd,&ps);
					return 0;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			break;
		}
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// CPopupWinShadow
//

BOOL CPopupWinShadow::CreateWin(HWND hWndParent, const CRect& rcWindow)
{
	m_pSetLayeredWindowAttributes = GetGlobals().GetSetLayeredWindowAttributes();

	#define WS_EX_LAYERED 0x00080000
	CreateEx(WS_EX_TOOLWINDOW | (m_pSetLayeredWindowAttributes == NULL ? 0 : WS_EX_LAYERED | WS_EX_TRANSPARENT),
			 _T(""),
			 WS_POPUP | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
			 rcWindow.left,
			 rcWindow.top,
			 rcWindow.Width(),
			 rcWindow.Height(),
			 NULL);
	if (!IsWindow())
		return FALSE;

	#define LWA_ALPHA     0x00000002
	if (m_pSetLayeredWindowAttributes)
		(m_pSetLayeredWindowAttributes)(m_hWnd, 0, (255 * 25) / 100, LWA_ALPHA);
	return TRUE;
}

LRESULT CPopupWinShadow::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
		case WM_MOUSEACTIVATE:
			return MA_NOACTIVATE;

		case WM_ACTIVATE:
			return FALSE;

		case WM_PAINT:
		{
			try
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcClient;
					GetClientRect(rcClient);

					if (m_pSetLayeredWindowAttributes)
						FillSolidRect(hDC, rcClient, GetSysColor(COLOR_3DDKSHADOW));
					else
						FillSolidRect(hDC, rcClient, m_pBar->m_crXPMenuBorderShadow);
					
					EndPaint(m_hWnd,&ps);
					return 0;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			break;
		}
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}


#define POPUPWINSTYLE WS_POPUP|WS_CLIPSIBLINGS|WS_CLIPCHILDREN
#define POPUPWINSTYLEEX WS_EX_TOPMOST|WS_EX_TOOLWINDOW|WS_EX_CONTROLPARENT

//
// PaintMenuScroll
//

void CPopupWin::PaintMenuScroll(HDC hDC, const CRect &rcOut, CPopupWin::Scroller eTopBottom)
{
	HDC hMemDC = GetMemDC();
	if (NULL == hMemDC)
		return;

	COLORREF crTextOld = SetTextColor(hDC, RGB(0,0,0));	
	COLORREF crBackOld = SetBkColor(hDC, RGB(255,255,255));
	HBITMAP  hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMenuScroll());
		
	COLORREF crColor = m_pBand->m_pBar->m_crForeground;

	BitBltMasked(hDC,
			     rcOut.left + (rcOut.Width() - CPopupWin::eBitmapWidth) / 2,
				 rcOut.top + (rcOut.Height() - CPopupWin::eBitmapHeight) / 2,
				 CPopupWin::eBitmapWidth,
				 CPopupWin::eBitmapHeight,
			     hMemDC,
			     eTopBottom * CPopupWin::eBitmapWidth,
			     0,
				 CPopupWin::eBitmapWidth,
				 CPopupWin::eBitmapHeight,
			     crColor);

	SetTextColor(hDC, crTextOld);
	SetBkColor(hDC, crBackOld);
	SelectBitmap(hMemDC, hBitmapOld);
}

//
// CPopupDragDrop
//

CBand* CPopupDragDrop::GetBand(POINT pt)
{
	return ((CPopupWin*)m_pWnd)->m_pBand;
}

//
// CPopupWin
//

CPopupWin::CPopupWin(CBand* pBand, BOOL bParentVertical)
	: m_pBand(pBand),
	  m_pOwningTool(NULL)
{
	m_bParentVertical = bParentVertical;
    m_pPopupDropTarget = new CBandDropTarget(CToolDataObject::eActiveBarDragDropId);
	assert(m_pPopupDropTarget);

	m_theShadow.Bar(pBand->m_pBar);

	m_nCurSel = eRemoveSelection;
	m_pBand->m_pPopupToolSelected = NULL;

	m_bTimerStarted = FALSE;
	pBand->m_pPopupWin = this;

	CTool* pTool;
	int nToolCount = m_pBand->m_pTools->GetVisibleToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = pBand->m_pTools->GetVisibleTool(nTool);
		if (NULL == pTool)
			continue;

		pTool->m_bPressed = FALSE;
	}
	
	m_pBand->bpV1.m_daDockingArea = ddDAPopup;
	m_dwWaitTime = 25;
	
	m_bFirstTimePaint = TRUE;

	BOOL bResult = RegisterWindow(DD_WNDCLASS("xPopup"),
								  CS_SAVEBITS,
								  NULL,
								  ::LoadCursor(NULL, IDC_ARROW));
	assert(bResult);
	m_thePopupDragDrop.m_pWnd = this;
	m_pPopupDropTarget->Init(&m_thePopupDragDrop, m_pBand->m_pBar, CToolDataObject::eActiveBarDragDropId);
	m_bMFU = FALSE; // eNormal;
	m_bTrackTimerStarted = FALSE;
	m_nCurPopup = -1;
	m_nParentIndex = -1;
	m_bPopupMenu = FALSE;
	if (m_pOwningTool)
	{
		m_pOwningTool->SubBand(NULL);
		m_pOwningTool->Release();
	}
	m_pFF = new CFlickerFree;
	assert(m_pFF);
}

CPopupWin::~CPopupWin()
{
	TRACE(1, _T("~CPopupWin\n"));
	try
	{
		delete m_pFF;
	}
	catch (...)
	{
		assert(FALSE);
	}
	m_pPopupDropTarget->Release();
}

//
// GetOptimalRect
//

BOOL CPopupWin::GetOptimalRect(CRect& rcPopup)
{
	BOOL bResult;
	rcPopup.SetEmpty();
	if (ddBTPopup == m_pBand->bpV1.m_btBands)
	{
		m_pBand->m_bPopupWinLock = TRUE;
		int w,h;
		bResult = m_pBand->CalcPopupLayout(w, h, TRUE, m_bPopupMenu);
		assert(bResult);
		m_pBand->m_bPopupWinLock = FALSE;
		w += 6;
		h += 6;
		rcPopup.right = w;
		rcPopup.bottom = h;
		bResult = AdjustWindowRectEx(&rcPopup, POPUPWINSTYLE, FALSE, POPUPWINSTYLEEX);
		assert(bResult);
	}
	else
	{
		CRect rcInit = m_pBand->bpV1.m_rcFloat;
		if (0 == m_pBand->bpV1.m_rcFloat.Width() || 0 == m_pBand->bpV1.m_rcFloat.Height())
		{
			rcInit.SetEmpty();
			if (-1 == m_pBand->bpV1.m_rcDimension.right /*|| -1 == m_pBand->bpV1.m_rcDimension.bottom*/)
				rcInit.right = 32767;
			else
			{
				rcInit.right = m_pBand->bpV1.m_rcDimension.right;
				rcInit.bottom = m_pBand->bpV1.m_rcDimension.bottom;
				rcPopup.SetEmpty();
				bResult = AdjustWindowRectEx(&rcPopup, POPUPWINSTYLE, FALSE, POPUPWINSTYLEEX);
				assert(bResult);
				rcInit.top -= rcPopup.top;
				rcInit.left -= rcPopup.left;
				rcInit.right += rcPopup.right;
				rcInit.bottom += rcPopup.bottom;
			}
			bResult = m_pBand->CalcLayout(rcInit, CBand::eLayoutFloat | CBand::eLayoutHorz, rcPopup, TRUE);
			assert(bResult);
		}
		else
		{
			rcPopup.SetEmpty();
			bResult = AdjustWindowRectEx(&rcPopup, POPUPWINSTYLE, FALSE, POPUPWINSTYLEEX);
			assert(bResult);
			rcInit.top -= rcPopup.top;
			rcInit.left -= rcPopup.left;
			rcInit.right += rcPopup.right;
			rcInit.bottom += rcPopup.bottom;
			bResult = m_pBand->CalcLayout(rcInit, CBand::eLayoutFloat | CBand::eLayoutVert, rcPopup, TRUE);
			assert(bResult);
		}
		rcPopup.right += 6;
		rcPopup.bottom += 6;
		bResult = AdjustWindowRectEx(&rcPopup, POPUPWINSTYLE, FALSE, POPUPWINSTYLEEX);
		assert(bResult);
	}
	return bResult;
}

BOOL CPopupWin::PreCreateWin(CTool*					  pOwningTool,
							 const CRect&			  rcBase, 
							 CPopupWin::FlipDirection eSubPopupFlipDir, 
							 BOOL					  bMFU,
							 BOOL					  bForceVFlip,
							 int 					  nHorzFlip)
{
	m_pOwningTool = pOwningTool;
	if (m_pOwningTool)
	{
		m_pOwningTool->AddRef();
		m_pOwningTool->SubBand(this);
	}
	m_nHorzFlip = nHorzFlip;
	m_nfdDirection = eSubPopupFlipDir;
	m_rcBase = rcBase;
	PushLayout();
	m_bMFU = bMFU;
	m_pBand->bpV1.m_vbVisible = VARIANT_TRUE;

	DockingAreaTypes daPrevDock = m_pBand->bpV1.m_daDockingArea;

	try
	{
		IReturnBool* pReturnBool = CRetBool::CreateInstance(NULL);
		if (pReturnBool)
		{
			HRESULT hResult = pReturnBool->put_Value(VARIANT_FALSE);
			if (SUCCEEDED(hResult))
			{
				++m_pBand->m_nInBandOpen;
				m_pBand->m_pBar->FireBandOpen((Band*)m_pBand, (ReturnBool*)pReturnBool);
				--m_pBand->m_nInBandOpen;
				
				VARIANT_BOOL vbCancel;
				hResult = pReturnBool->get_Value(&vbCancel);
				if (SUCCEEDED(hResult))
				{
					if (VARIANT_TRUE == vbCancel)
					{
						pReturnBool->Release();
						return FALSE;
					}
				}
			}
			pReturnBool->Release();
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}

	//
	// Prevent programmer from screwing the system
	//
	// Users change the DockingArea in the FireBandOpen Event causes this function to fail, 
	// This recovers from that
	//
	
	switch (daPrevDock)
	{
	case ddDAFloat:
	case ddDAPopup:
		m_pBand->bpV1.m_daDockingArea = daPrevDock; 
		break;
	}

	if (VARIANT_FALSE == m_pBand->bpV1.m_vbVisible)
	{
		PopLayout();
		return FALSE;
	}

	CRect rcPopup;
	GetOptimalRect(rcPopup);
	rcPopup.Offset(-rcPopup.left, -rcPopup.top);
	m_rcClientOrginal = rcPopup;
	m_rcClientOrginal.Inflate(1, 1);
	// Now check if this popup fits on the screen
	HMONITOR hMon = MonitorFromRect(&rcBase, MONITOR_DEFAULTTONEAREST);
	MONITORINFO mi;
	mi.cbSize = sizeof(MONITORINFO);
	BOOL bResult = GetMonitorInfo(hMon, &mi);
	CRect rcWin = mi.rcWork;

	CRect rcNew = rcPopup;
	int nWidth = rcPopup.Width();
	int nHeight = rcPopup.Height();

	if (eDirectionFlipVert == eSubPopupFlipDir)
		rcNew.Offset(rcBase.left - 1, rcBase.bottom + 1);
	else
		rcNew.Offset(rcBase.right, rcBase.top);
	
	m_nScrollPos = 0;
	m_bScrollActive = FALSE;
	m_bFlipVert = FALSE;
	m_bFlipHorz = FALSE;

	if (bForceVFlip)
	{
		if (eDirectionFlipVert == eSubPopupFlipDir) 
			rcNew.top = rcBase.top - nHeight;
		else
			rcNew.top = rcBase.bottom - nHeight;
		rcNew.bottom = rcNew.top + nHeight;
	}

	switch (nHorzFlip)
	{
	case ddPopupMenuLeftAlign:
		// Do nothing
		break;

	case ddPopupMenuCenterAlign:
		{
			int nOffset = rcNew.Width() / 2;
			if ((rcNew.right - nOffset) > rcWin.right)
				nOffset += rcNew.right - nOffset - rcWin.right;
			else if ((rcNew.left - nOffset) < rcWin.left)
				nOffset += rcNew.left - nOffset - rcWin.left;
			rcNew.Offset(-nOffset, 0);
		}
		break;

	case ddPopupMenuRightAlign:
		int nOffset = rcNew.Width();
		if ((rcNew.right - nOffset) > rcWin.right)
			nOffset += rcNew.right - nOffset - rcWin.right;
		else if ((rcNew.left - nOffset) < rcWin.left)
			nOffset += rcNew.left - nOffset - rcWin.left;
		rcNew.Offset(-nOffset, 0);
		break;
	}

	// flip vertcally
	if (rcNew.bottom > rcWin.bottom)
		m_bFlipVert = TRUE;

	// flip horizontal
	if (rcNew.right > rcWin.right) 
		m_bFlipHorz = TRUE;
	
	if (m_bFlipHorz)
	{
		if (eDirectionFlipVert == eSubPopupFlipDir)
		{
			if (rcBase.right < rcWin.right)
				rcNew.right = rcBase.right;
			else
				rcNew.right = rcWin.right;
		}
		else
		{
			if (rcBase.left < rcWin.right)
				rcNew.right = rcBase.left;
			else
				rcNew.right = rcWin.right;
		}
		rcNew.left = rcNew.right - nWidth;
		if (rcNew.left < 0)
			rcNew.Offset(-rcNew.left, 0);
	}
	else
	{
		rcNew.right = rcNew.left + nWidth;
		if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook && m_pBand->m_pParentPopupWin)
			rcNew.Offset(6, 0);
	}

	if (m_bFlipVert)
	{
		int nLineHeight = GetLineHeight();
		// Check for scrolling req.
		if (rcBase.top - nHeight < 0)
		{
			// Doesn't fit either way so choose the larger one
			int nMaxHeight = 0;
			if (rcBase.top - rcWin.top > rcWin.bottom - rcBase.top) 
			{
				//
				// Flipping upwards
				//

				if (eDirectionFlipVert == eSubPopupFlipDir)
				{
					rcNew.bottom = rcBase.top;
					nMaxHeight = rcBase.top - rcWin.top;
					if (nHeight > nMaxHeight)
						nHeight = nMaxHeight;
				}
				else
				{
					rcNew.bottom = rcBase.bottom;
					nMaxHeight = rcBase.top - rcWin.top;
					if (nHeight > nMaxHeight)
						nHeight = nMaxHeight;
				}
				rcNew.top = rcNew.bottom - nHeight;

			}
			else
			{
				//
				// Alright needs scrolling
				// 

				m_bFlipVert = FALSE;
				if (eDirectionFlipVert == eSubPopupFlipDir)
				{
					rcNew.top = rcBase.bottom;
					nMaxHeight = rcBase.bottom - rcWin.bottom;
					if (nHeight > nMaxHeight)
						nHeight = nMaxHeight;
				}
				else
				{
					rcNew.top = rcBase.top;
					nMaxHeight = rcBase.top - rcWin.bottom;
					if (nHeight > nMaxHeight)
						nHeight = nMaxHeight;
				}

				rcNew.bottom = rcNew.top - nHeight;
			}
			m_bScrollActive = TRUE;
		}
		else 
		{
			// 
			// Fits if flipped upwards
			// 

			if (eDirectionFlipVert == eSubPopupFlipDir)
				rcNew.top = rcBase.top - nHeight;
			else
				rcNew.top = rcBase.bottom - nHeight;
			rcNew.bottom = rcNew.top + nHeight;
		}
	}

	// Done with flipping
	m_rcWindow = rcNew;
	return TRUE;
}

//
// CreateWin
//

BOOL CPopupWin::CreateWin(HWND hWndParent)
{
	try
	{
		// Add Non Client Space 

		if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
		{
			CRect rcWindow = m_rcWindow;
			rcWindow.Offset(4, 4);
			m_theShadow.CreateWin(hWndParent, rcWindow);
		}

		CreateEx(POPUPWINSTYLEEX,
				 _T(""),
				 POPUPWINSTYLE,
				 m_rcWindow.left,
				 m_rcWindow.top,
				 m_rcWindow.Width(),
				 m_rcWindow.Height(),
				 NULL);
			
		if (!IsWindow())
			return FALSE;
		if (m_bScrollActive)
		{
			//
			// Create a timer
			//

			SetTimer(m_hWnd, ePopupScrollTimer, m_pBand->bpV1.m_nScrollingSpeed, NULL);
		}

		if (m_pBand->m_pBar->m_pDragDropManager && S_OK != m_pBand->m_pBar->m_pDragDropManager->RegisterDragDrop((OLE_HANDLE)m_hWnd, m_pPopupDropTarget))
		{
			assert(FALSE);
			TRACE(1, "Failed to register drag and drop target\n");
		}
		else if (S_OK != GetGlobals().m_pDragDropMgr->RegisterDragDrop((OLE_HANDLE)m_hWnd, m_pPopupDropTarget))
		{
			assert(FALSE);
			TRACE(1, "Failed to register drag and drop target\n");
		}
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// WindowProc
//

LRESULT CPopupWin::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (GetGlobals().WM_POPUPWINMSG == nMsg)
		{
			try
			{
				// 0 means testing for window availability
				if (0 == wParam) 
					*(BOOL*)lParam = TRUE;
				else if (1 == wParam)
					(*(CPopupWin**)lParam) = this;
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
			return 0;
		}

		switch(nMsg)
		{
		case WM_ERASEBKGND:
			return 1;

		case WM_MOUSEACTIVATE:
			return MA_NOACTIVATE;

		case WM_ACTIVATE:
			return FALSE;

		case WM_SETFOCUS:
			break;

		case WM_PAINT:
			try
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcClient;
					GetClientRect(rcClient);
					Draw(hDC, rcClient);
					EndPaint(m_hWnd,&ps);
					return 0;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			break;

		case WM_MOUSEMOVE:
		case WM_NCMOUSEMOVE:
			try
			{
				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				if (ddBTPopup == m_pBand->bpV1.m_btBands)
					OnMouseMove(pt);
				else
				{
					if (HitTestDetach(pt))
					{
						m_nCurSel = eDetachBandCaptionSelected;
						HDC hDC = GetDC(m_hWnd);
						if (hDC)
						{
							PaintDetachBar(hDC);
							ReleaseDC(m_hWnd, hDC);
						}
					}
					else
					{
						if (eDetachBandCaptionSelected == m_nCurSel)
						{
							m_nCurSel = eRemoveSelection;
							HDC hDC = GetDC(m_hWnd);
							if (hDC)
							{
								PaintDetachBar(hDC);
								ReleaseDC(m_hWnd, hDC);
							}
						}
						OnMouseMove(pt);
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			return 0;

		case WM_TIMER:
			try
			{
				switch (wParam)
				{
				case eCloseSubTimer:
					{
						KillTimer(m_hWnd, eCloseSubTimer);
						POINT pt;
						HWND hWndPopup;
						GetCursorPos(&pt);
						m_pBand->QueryPopupWin(hWndPopup, pt);
						if (m_pBand->m_pChildPopupWin && hWndPopup != m_pBand->m_pChildPopupWin->m_hWnd)
						{
							m_pBand->m_pChildPopupWin->ReparentEditAndOrCombo();
							m_pBand->m_pChildPopupWin->DestroyWindow();
							m_pBand->m_pChildPopupWin = NULL;
							m_nCurPopup = -1;
						}
					}
					break;

				case eOpenSubTimer:
					KillTimer(m_hWnd, eOpenSubTimer);
					m_bTimerStarted = FALSE;
					OpenSubPopup();
					break;

				case ePopupScrollTimer:
					{
						CRect rcClient;
						GetClientRect(rcClient);

						POINT pt;
						GetCursorPos(&pt);
						::ScreenToClient(m_hWnd, &pt);

						CRect rc;
						if (m_bScrollActive)
						{
							if (m_bTopScrollActive && PtInRect(&m_rcTopScroller, pt))
							{
								if (::IsWindow(m_pBand->m_pBar->m_hWndActiveCombo))
									::SendMessage(m_pBand->m_pBar->m_hWndActiveCombo, WM_CLOSE, 0, 0);

								if (0 != m_nScrollPos)
								{
									SetCurSel(eRemoveSelection);
									--m_nScrollPos;
								}
								InvalidateRect(NULL, FALSE);
								UpdateWindow();
								break;
							}
							else if (m_bBottomScrollActive && PtInRect(&m_rcBottomScroller, pt))
							{
								if (::IsWindow(m_pBand->m_pBar->m_hWndActiveCombo))
									::SendMessage(m_pBand->m_pBar->m_hWndActiveCombo, WM_CLOSE, 0, 0);

								SetCurSel(eRemoveSelection);
								m_nScrollPos++;
								InvalidateRect(NULL, FALSE);
								UpdateWindow();
								break;
							}
						}
					}
					break;

				case eExpandTimer:
					{
						KillTimer(m_hWnd, eExpandTimer);
						m_bTrackTimerStarted = FALSE;
						if (-1 != m_nCurSel)
						{
							m_pBand->m_pBar->m_bPopupMenuExpanded = TRUE;
							CTool* pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
							if (pTool) 
							{
								pTool->tpV1.m_vbVisible = VARIANT_FALSE;
								m_nCurSel = -1;
								m_pBand->m_pPopupToolSelected = NULL;
								CRect rcWindow;
								GetWindowRect(rcWindow);
								if (CalcFlipingAndScrolling(rcWindow))
								{
									if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
									{
										BOOL bVertical = m_pBand->IsVertical();
										BOOL bFlipVert = FALSE;
										BOOL bFlipHorz = FALSE;
										if (pTool->SubBand())
										{
											bFlipHorz = pTool->SubBand()->FlipHorz();
											bFlipVert = pTool->SubBand()->FlipVert();
										}
										CRect rcShadow = rcWindow;
										if (bVertical)
											rcShadow.Offset(3, 5);
										else
											rcShadow.Offset(4, 3);
										m_theShadow.Show(rcShadow);
									}
									SetWindowPos(HWND_TOPMOST, 
												 rcWindow.left, 
												 rcWindow.top, 
												 rcWindow.Width(), 
												 rcWindow.Height(), 
												 SWP_NOACTIVATE);
									InvalidateRect(NULL, FALSE);
								}
							}
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
			break;

		case WM_CANCELMODE:
			try
			{
				ShowWindow(SW_HIDE);
				::PostMessage(m_pBand->m_pBar->m_hWnd, GetGlobals().WM_KILLWINDOW, (WPARAM)m_hWnd, 0);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			break;

		case WM_DESTROY:
			{
				try
				{
					m_pBand->m_pBar->m_theToolStack.Pop();

					KillTimer(m_hWnd, eExpandTimer);
					KillTimer(m_hWnd, eCloseSubTimer);
					if (m_bScrollActive)
						KillTimer(m_hWnd, ePopupScrollTimer);

					CBand* pBand = m_pBand;
					assert(pBand);

					//
					// We are doing this because m_pBand will soon be undefined when this 
					// object gets destroyed
					//

					if (pBand)
					{
						if (pBand == pBand->m_pBar->m_pMoreTools) 
						{
							pBand->m_pBar->m_pBands->RemoveEx(pBand);
							pBand->m_pBar->m_pMoreTools->m_pTools->RemoveAll();
							if (ddBFCustomize & pBand->bpV1.m_dwFlags)
							{
								HRESULT hResult = pBand->m_pBar->m_pBands->RemoveEx(pBand->m_pBar->m_pAllTools);
								pBand->m_pBar->m_pAllTools->m_pTools->RemoveAll();
							}
						}

						pBand->ParentWindowedTools(NULL);

						if (pBand->m_pChildPopupWin)
						{
							if (pBand->m_pBar && pBand->m_pBar->m_pDesigner)
								pBand->m_pChildPopupWin->PostMessage(WM_CLOSE);
							else
							{
								try
								{
									pBand->m_pChildPopupWin->ReparentEditAndOrCombo();
									pBand->m_pChildPopupWin->DestroyWindow();
								}
								catch (...)
								{
									assert(FALSE);
								}
							}
							pBand->m_pChildPopupWin = NULL;
						}

						//
						// Removing the expand tool from the menu
						//

						if (m_bMFU)
						{
							VARIANT vIndex;
							vIndex.vt = VT_I4;
							vIndex.lVal = pBand->m_pTools->GetToolCount();
							if (vIndex.lVal > 0)
							{
								vIndex.lVal--;
								CTool* pTool = pBand->m_pTools->GetTool(vIndex.lVal);
								if (pTool && CTool::ddTTMenuExpandTool == pTool->tpV1.m_ttTools)
									HRESULT hResult = pBand->m_pTools->Remove(&vIndex);
							}
						}

						//
						// Unregistering the window from the drag drop manager
						//

						if (pBand->m_pBar->m_pDragDropManager)
							pBand->m_pBar->m_pDragDropManager->RevokeDragDrop((OLE_HANDLE)m_hWnd);
						else
						{
							assert(GetGlobals().m_pDragDropMgr);
							GetGlobals().m_pDragDropMgr->RevokeDragDrop((OLE_HANDLE)m_hWnd);
						}

						pBand->SetPopupIndex(eRemoveSelection);
						pBand->bpV1.m_vbVisible = VARIANT_FALSE;
						pBand->m_pPopupWin = NULL;
						try
						{
							pBand->m_pBar->FireBandClose((Band*)pBand);


							if (pBand->m_pBar->m_bCustomization && ::IsWindow(pBand->m_pBar->m_hWndModifySelection) && 0 == wcscmp(L"popEditTool", pBand->m_bstrName))
							{
								//
								// This is ugly but it works, This is to change the state of Modify Selection Button
								// on the customize dialog.
								//

								CRect rcButton;
								::GetWindowRect(pBand->m_pBar->m_hWndModifySelection, &rcButton);
								POINT pt;
								GetCursorPos(&pt);
								if (!PtInRect(&rcButton, pt))
								{
									//
									// If the cursor is not over the ModifySelection Button change it's state to 
									// unchecked.
									//

									if (BST_CHECKED == ::SendMessage(pBand->m_pBar->m_hWndModifySelection, BM_GETCHECK, 0, 0))
										::PostMessage(pBand->m_pBar->m_hWndModifySelection, BM_SETCHECK, BST_UNCHECKED, 0);
								}
							}
						}
						CATCH
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
					}

					if (pBand->m_pParentPopupWin && pBand->m_pParentPopupWin->m_pBand) 
						pBand->m_pParentPopupWin->m_pBand->m_pChildPopupWin = NULL;
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
			break;

		case WM_NCDESTROY:
			{	
				LRESULT lResult = 0;
				try
				{
					if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook && m_theShadow.IsWindow())
						m_theShadow.DestroyWindow();
					
					if (m_pOwningTool)
					{
						m_pOwningTool->SubBand(NULL);
						m_pOwningTool->Release();
						m_pOwningTool = NULL;
					}

					lResult = FWnd::WindowProc(nMsg, wParam, lParam); 
	
					MSG msg;
					while (::PeekMessage(&msg, NULL, WM_PAINT, WM_PAINT, PM_NOREMOVE))
					{
						if (GetMessage(&msg, NULL, WM_PAINT, WM_PAINT))
							DispatchMessage(&msg);
					}

					//
					// remove selection
					//
					
					if (m_nCurSel >= 0 && m_nCurSel < m_pBand->m_pTools->GetVisibleToolCount())
						m_pBand->m_pTools->GetVisibleTool(m_nCurSel)->m_bPressed = FALSE;

					if (m_pBand->m_pBar->m_diCustSelection.pBand == m_pBand && ddCBSystem != m_pBand->bpV1.m_nCreatedBy)
						memset(&m_pBand->m_pBar->m_diCustSelection, 0, sizeof(CBar::DropInfo));

					m_pBand->m_nPopupIndex = -1;
					m_pBand->m_nDropLoc = -1;
					
					PopLayout();

					delete this;
					return lResult;
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
					return lResult;
				}
			}
			break;

		case WM_SYSKEYDOWN:
		case WM_KEYDOWN:

			//
			// TODO
			// To prevent keys from breeding to other controls on the form we need to check  SYSKEYUP KEYUP WM_CHAR
			//
			try
			{
				OnKey(wParam, (BOOL*)lParam);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			return 0;

		case WM_KEYUP:
			break;

		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			try
			{
				if (ddBTPopup == m_pBand->bpV1.m_btBands || ddDAPopup == m_pBand->bpV1.m_daDockingArea)
				{
					if (ddBTPopup == m_pBand->bpV1.m_btBands && m_pBand->m_pBar->m_bWhatsThisHelp)
					{
						CTool* pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
						if (pTool)
						{
							try
							{
								m_pBand->m_pBar->FireWhatsThisHelp((Band*)m_pBand, (Tool*)pTool, pTool->tpV1.m_nHelpContextId);
								m_pBand->m_pBar->SetActiveTool(NULL);
							}
							CATCH
							{
								assert(FALSE);
								REPORTEXCEPTION(__FILE__, __LINE__)
							}
						}
					}
					else
					{
						try
						{
							OnButton(nMsg, FALSE, wParam);
						}
						CATCH
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
						return 0;
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			
			//
			// Fall through		
			//

		case WM_LBUTTONDBLCLK:
		case WM_LBUTTONUP:
		case WM_RBUTTONUP:
			try
			{
				OnButton(nMsg, TRUE, wParam);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			return 0;

		case WM_SYSCOMMAND:
			{
				UINT nID = wParam;
				if ((nID & 0xFFF0) != SC_CLOSE || 
					(GetKeyState(VK_F4) < 0 && GetKeyState(VK_MENU) < 0))
				{
					return 0;
				}
			}
			break;

		case WM_COMMAND:
			try
			{
				switch (LOWORD(wParam))
				{
				case CTool::eEdit:
				case CTool::eCombobox:
					
					//
					// Message reflection for the built in Combobox and Edit Tool.
					//

					switch (HIWORD(wParam))
					{
					case EN_KILLFOCUS:
					case EN_CHANGE:
					case CBN_CLOSEUP:
					case CBN_KILLFOCUS:
					case CBN_DROPDOWN:
					case CBN_EDITCHANGE:
						::SendMessage((HWND)lParam, WM_COMMAND, wParam, lParam);
						break;
					}
					break;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			break;

		case WM_GETMINMAXINFO:
			try
			{
				MINMAXINFO* pMMI = (MINMAXINFO*)lParam;
				// allow Windows to fill in the defaults
				FWnd::WindowProc(nMsg,wParam,lParam);
				// don't allow sizing smaller than the non-client area
				CRect rcWindow;
				GetWindowRect(rcWindow);

				CRect rcClient;
				GetClientRect(rcClient);
				
				pMMI->ptMinTrackSize.x = rcWindow.Width() - rcClient.right;
				pMMI->ptMinTrackSize.y = rcWindow.Height() - rcClient.bottom;
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
			return 0;
		}
		return FWnd::WindowProc(nMsg, wParam, lParam);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// DrawXPBackgroundBorder
//

void CPopupWin::DrawXPBackgroundBorder(HDC hDCOff, CRect& rcClient)
{
	try
	{
		BOOL bResult;
		if (!m_bFlipHorz && ddPopupMenuRightAlign == m_nHorzFlip)
			m_bFlipHorz = ddMSAnimateSlide == m_pBand->m_pBar->bpV1.m_msMenuStyle;
		
		FillSolidRect(hDCOff, rcClient, m_pBand->m_pBar->m_crXPMenuBackground);

		CRect rcOwningTool;
		if (m_pOwningTool)
		{
			m_nOwningToolIndex = m_pOwningTool->m_pBand->m_pTools->GetToolIndex(m_pOwningTool);
			m_pOwningTool->m_pBand->GetToolScreenRect(m_nOwningToolIndex, rcOwningTool);
			ScreenToClient(rcOwningTool);
		}
		if (ddBTPopup == m_pBand->bpV1.m_btBands)
		{
			// This is the stripe down the side of a popup menu
			CRect rcBorder = rcClient;
			rcBorder.Inflate(0, -1);
			rcBorder.bottom--;
			rcBorder.right = rcBorder.left + m_pBand->m_sizeMaxIcon.cx + 8;
			if (m_pBand->m_phPicture.m_pPict)
			{
				SIZEL sizePicture = {0, 0};
				m_pBand->m_phPicture.m_pPict->get_Width(&sizePicture.cx);
				HiMetricToPixel(&sizePicture, &sizePicture);
				rcBorder.right += sizePicture.cx;
			}
			FillSolidRect(hDCOff, rcBorder, m_pBand->m_pBar->m_crXPBandBackground);
		}

		HPEN aPen = CreatePen(PS_SOLID, 1, m_pBand->m_pBar->m_crXPMenuSelectedBorderColor);
		if (aPen)
		{
			HPEN penOld = SelectPen(hDCOff, aPen);

			if (m_pBand->m_pParentPopupWin)
			{
				MoveToEx(hDCOff, rcClient.left, rcClient.top, NULL);
				LineTo(hDCOff, rcClient.right - 1, rcClient.top);
				LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
				LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
				LineTo(hDCOff, rcClient.left, rcClient.top);
			}
			else if (m_bParentVertical)
			{
				try
				{
					if (m_nfdDirection == eDirectionFlipHort)
					{
						if (!m_bFlipVert && !m_bFlipHorz)
						{
							MoveToEx(hDCOff, rcClient.left, rcClient.top, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);
							LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
							LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);

							MoveToEx(hDCOff, rcClient.left, rcClient.bottom - 1, NULL);
							LineTo(hDCOff, rcClient.left, rcOwningTool.bottom - 2);
							
							MoveToEx(hDCOff, rcClient.left, rcOwningTool.top, NULL);
							LineTo(hDCOff, rcClient.left, rcClient.top);
						}
						else if (m_bFlipVert && !m_bFlipHorz)
						{
							MoveToEx(hDCOff, rcClient.left, rcClient.top, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);
							LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
							LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);

							MoveToEx(hDCOff, rcClient.left, rcClient.bottom - 1, NULL);
							LineTo(hDCOff, rcClient.left, rcOwningTool.bottom - 2);
							
							MoveToEx(hDCOff, rcClient.left, rcOwningTool.top - 1, NULL);
							LineTo(hDCOff, rcClient.left, rcClient.top);
						}
						else if (!m_bFlipVert && m_bFlipHorz)
						{
							MoveToEx(hDCOff, rcClient.right, rcClient.top, NULL);
							LineTo(hDCOff, rcClient.left, rcClient.top);
							LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
							LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);

							MoveToEx(hDCOff, rcClient.right - 1, rcClient.bottom - 1, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcOwningTool.bottom - 1);
							
							MoveToEx(hDCOff, rcClient.right - 1, rcOwningTool.top - 1, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						}
						else
						{
							MoveToEx(hDCOff, rcClient.right, rcClient.bottom - 1, NULL);
							LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
							LineTo(hDCOff, rcClient.left, rcClient.top);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);

							MoveToEx(hDCOff, rcClient.right - 1, rcClient.bottom - 1, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcOwningTool.bottom - 1);
							
							MoveToEx(hDCOff, rcClient.right - 1, rcOwningTool.top - 1, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						}
					}
					else if (!m_bFlipVert && !m_bFlipHorz)
					{
						MoveToEx(hDCOff, rcClient.left, rcClient.top, NULL);
						LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.left, rcClient.top + m_rcBase.Height() - 1);
					}
					else if (m_bFlipVert && !m_bFlipHorz)
					{
						MoveToEx(hDCOff, rcClient.left, rcClient.bottom - m_rcBase.Height() - 1, NULL);
						LineTo(hDCOff, rcClient.left, rcClient.top);
						LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.left - 1, rcClient.bottom - 1);
					}
					else if (!m_bFlipVert && m_bFlipHorz)
					{
						MoveToEx(hDCOff, rcClient.right, rcClient.top, NULL);
						LineTo(hDCOff, rcClient.left, rcClient.top);
						LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.right - 1, rcClient.top + m_rcBase.Height() - 1);
					}
					else
					{
						MoveToEx(hDCOff, rcClient.right, rcClient.bottom - 1, NULL);
						LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.left, rcClient.top);
						LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - m_rcBase.Height() - 1);
					}
				}
				catch (...)
				{
					assert(FALSE);
				}
			}
			else
			{
				try
				{
					if (m_nfdDirection == eDirectionFlipHort)
					{
						if (m_bFlipHorz && !m_bFlipVert)
						{
							MoveToEx(hDCOff, rcClient.right, rcClient.top, NULL);
							LineTo(hDCOff, rcClient.left, rcClient.top);
							LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
							LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
						
							MoveToEx(hDCOff, rcClient.right - 1, rcClient.bottom, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcOwningTool.bottom);
							
							MoveToEx(hDCOff, rcClient.right - 1, rcOwningTool.top, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						}
						else if (!m_bFlipHorz && !m_bFlipVert)
						{
							MoveToEx(hDCOff, rcClient.left, rcClient.top, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);
							LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
							LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
							
							LineTo(hDCOff, rcClient.left, rcClient.top + m_rcBase.Height());
						}
						else if (!m_bFlipHorz && m_bFlipVert)
						{
							MoveToEx(hDCOff, rcClient.left, rcClient.bottom - m_rcBase.Height() - 1, NULL);
							
							LineTo(hDCOff, rcClient.left, rcClient.top);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);
							LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
							LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
						}
						else
						{
							MoveToEx(hDCOff, rcClient.left, rcClient.top, NULL);
							LineTo(hDCOff, rcClient.right - 1, rcClient.top);
							LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
							LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
							
							LineTo(hDCOff, rcClient.left, rcClient.top + m_rcBase.Height() - 1);
						}
					}
					else if (!m_bFlipVert && !m_bFlipHorz)
					{
						MoveToEx(hDCOff, rcClient.left, rcClient.top, NULL);
						LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						
						MoveToEx(hDCOff, rcClient.left, rcClient.top, NULL);
						LineTo(hDCOff, rcOwningTool.left, rcClient.top);
						
						MoveToEx(hDCOff, rcOwningTool.right - 1, rcClient.top, NULL);
						LineTo(hDCOff, rcClient.right, rcClient.top);
					}
					else if (m_bFlipVert && !m_bFlipHorz)
					{
						MoveToEx(hDCOff, rcClient.left, rcClient.bottom, NULL);
						LineTo(hDCOff, rcClient.left, rcClient.top);
						LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);

						MoveToEx(hDCOff, rcClient.right, rcClient.bottom - 1, NULL);
						LineTo(hDCOff, rcOwningTool.right - 2, rcClient.bottom - 1);
						
						MoveToEx(hDCOff, rcOwningTool.left - 1, rcClient.bottom - 1, NULL);
						LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
					}
					else if (!m_bFlipVert && m_bFlipHorz)
					{
						MoveToEx(hDCOff, rcClient.right, rcClient.top, NULL);
						LineTo(hDCOff, rcOwningTool.right, rcClient.top);
						
						MoveToEx(hDCOff, rcOwningTool.left - 1, rcClient.top, NULL);
						LineTo(hDCOff, rcClient.left, rcClient.top);

						LineTo(hDCOff, rcClient.left, rcClient.top);
						LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.right - 1, rcClient.bottom - 1);
						LineTo(hDCOff, rcClient.right - 1, rcClient.top - 1);
					}
					else
					{
						MoveToEx(hDCOff, rcClient.right - 1, rcClient.bottom, NULL);
						LineTo(hDCOff, rcClient.right - 1, rcClient.top);
						LineTo(hDCOff, rcClient.left, rcClient.top);
						LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);

						MoveToEx(hDCOff, rcClient.right, rcClient.bottom - 1, NULL);
						LineTo(hDCOff, rcOwningTool.right, rcClient.bottom - 1);
						
						MoveToEx(hDCOff, rcOwningTool.left - 1, rcClient.bottom - 1, NULL);
						LineTo(hDCOff, rcClient.left, rcClient.bottom - 1);
					}
				}
				catch (...)
				{
					assert(FALSE);
				}
			}
	
			SelectPen(hDCOff, penOld);
			bResult = DeletePen(aPen);
			assert(bResult);
		}
	}
	catch (...)
	{
		assert(FALSE);
	}
}

//
// Draw
//

void CPopupWin::Draw(HDC hDC, CRect rcClient)
{
	try
	{
		TRACE(5, "CPopupWin::Draw\n");

		POINT pt = {0, 0};
		int nWidthBitmap = rcClient.Width();
		int nHeightBitmap = rcClient.Height();

		BOOL bXPLook = VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook;

		CFlickerFree* pff = &(m_pBand->m_pBar->m_ffObj);
		HDC hDCOff = pff->RequestDC(hDC, nWidthBitmap, nHeightBitmap);
		if (NULL == hDCOff)
			hDCOff = hDC;
		else
			rcClient.Offset(-rcClient.left, -rcClient.top);
		
		HPALETTE hPalOld = NULL;
		HPALETTE hPal = m_pBand->m_pBar->Palette();
		if (hPal)
		{
			hPalOld = SelectPalette(hDCOff, hPal, FALSE);
			RealizePalette(hDCOff);
		}

		BOOL bHasTexture = m_pBand->m_pBar->HasTexture();

		if (bXPLook)
		{
			DrawXPBackgroundBorder(hDCOff, rcClient);
			rcClient.Inflate(-1, -1);
		}
		else
		{
			m_pBand->m_pBar->DrawEdge(hDCOff, rcClient, BDR_RAISEDOUTER, BF_RECT);
			rcClient.Inflate(-1, -1);

			m_pBand->m_pBar->DrawEdge(hDCOff, rcClient, BDR_RAISEDINNER, BF_RECT);
			rcClient.Inflate(-1, -1);

			if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved)
			{
				if (bHasTexture && ddBOPopups & m_pBand->m_pBar->bpV1.m_dwBackgroundOptions)
				{
					SetBrushOrgEx(hDCOff, 0, 0, NULL);
					m_pBand->m_pBar->FillTexture(hDCOff, rcClient);
				}
				else
					FillSolidRect(hDCOff, rcClient, m_pBand->m_pBar->m_crBackground);
			}
			else
				FillSolidRect(hDCOff, rcClient, m_pBand->m_pBar->m_crMDIMenuBackground);
		}

		// Draw Border
		if (ddBTPopup == m_pBand->bpV1.m_btBands)
			m_pBand->m_bPopupWinLock = TRUE;

		if (m_bScrollActive)
		{
			//
			// Paint the scrolling menu
			//

			rcClient.Inflate(-1, -1); 
			m_rcClient = rcClient;
			m_bBottomScrollActive = FALSE;
			m_bTopScrollActive = FALSE;

			//
			// Now draw menu area
			//

			m_nPhysicalScrollPos = m_pBand->m_prcTools[m_nScrollPos].top + 3;
			
			int nToolCount = m_pBand->m_pTools->GetVisibleToolCount();
			int nLastVisibleTool = nToolCount - 1;
			
			//
			// Are we at the end point?
			//

			if (m_pBand->m_prcTools[nLastVisibleTool].bottom <= rcClient.Height() + m_nPhysicalScrollPos) 
				m_nPhysicalScrollPos = m_pBand->m_prcTools[nLastVisibleTool].bottom - rcClient.Height();
			else
				m_bBottomScrollActive = TRUE;
		
			if (0 != m_nScrollPos)
			{
				//
				// Draw Top Scroller
				//

				m_rcTopScroller = rcClient;
				m_rcTopScroller.bottom = m_rcTopScroller.top + eScrollerHeight;
				
				m_rcClient.top += eScrollerHeight;

				FillSolidRect(hDCOff, m_rcTopScroller, m_pBand->m_pBar->m_crBackground);

				m_bTopScrollActive = TRUE;
			}
			if (m_bBottomScrollActive)
			{
				//
				// Draw Bottom Scroller
				//

				m_rcBottomScroller = rcClient;
				m_rcBottomScroller.top = m_rcBottomScroller.bottom - eScrollerHeight;

				m_rcClient.bottom -= m_rcBottomScroller.Height();

				FillSolidRect(hDCOff, m_rcBottomScroller, m_pBand->m_pBar->m_crBackground);
			}

			HDC hDCOff2 = m_pFF->RequestDC(hDC, m_rcClientOrginal.Width(), m_rcClientOrginal.Height());

			CRect rcTemp = m_rcClientOrginal;
			if (bXPLook)
			{
				DrawXPBackgroundBorder(hDCOff2, rcTemp);
				rcTemp.Inflate(-1, -1);
			}
			else
			{
				if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved)
				{
					if (bHasTexture && ddBOPopups & m_pBand->m_pBar->bpV1.m_dwBackgroundOptions)
					{
						SetBrushOrgEx(hDCOff2, 0, 0, NULL);
						m_pBand->m_pBar->FillTexture(hDCOff2, rcTemp);
					}
					else
						FillSolidRect(hDCOff2, rcTemp, m_pBand->m_pBar->m_crBackground);
				}
				else
					FillSolidRect(hDCOff2, rcTemp, m_pBand->m_pBar->m_crMDIMenuBackground);
			}

			m_pBand->Draw(hDCOff2, rcTemp, ddBTPopup == m_pBand->bpV1.m_btBands ? VARIANT_TRUE : m_pBand->IsWrappable(), pt);

			if (m_bTopScrollActive)
				PaintMenuScroll(hDCOff, m_rcTopScroller, eTop);

			if (m_bBottomScrollActive)
				PaintMenuScroll(hDCOff, m_rcBottomScroller, eBottom);

			HRGN hRgn = NULL; 
			HRGN hRgnPrev = NULL;
			int nResult;
			hRgn = CreateRectRgnIndirect(&m_rcClient); 
			assert(hRgn);
			if (hRgn)
			{
				hRgnPrev = CreateRectRgn(0,0,0,0);
				if (hRgnPrev)
					nResult = GetClipRgn(hDCOff, hRgnPrev);

				nResult = SelectClipRgn(hDCOff, hRgn); 
				assert(ERROR != nResult);
			}

			m_pFF->Paint(hDCOff, 
						 rcClient.left, 
						 rcClient.top, 
						 rcClient.Width(), 
						 rcClient.Height(), 
						 0, 
						 m_nPhysicalScrollPos);
			
			if (hRgn)
			{
				if (-1 == nResult)
					SelectClipRgn(hDCOff, NULL); 
				else
					SelectClipRgn(hDCOff, hRgnPrev);
				
				BOOL bResult = DeleteRgn(hRgn);
				assert(bResult);
			}

			if (hRgnPrev)
			{
				BOOL bResult = DeleteRgn(hRgnPrev);
				assert(bResult);
			}
		}
		else
		{
			//
			// Paint non scrolling menu
			//

			if (ddBTPopup == m_pBand->bpV1.m_btBands && m_nCurSel >= 0 && m_nCurSel < m_pBand->m_pTools->GetVisibleToolCount())
			{
				// Select the first tool on the band
				CTool* pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
				if (pTool)
					pTool->m_bPressed = TRUE;
			}
			// For the detach bar...
			rcClient.Inflate(-1, -1); 
			IntersectClipRect(hDCOff, 
							  rcClient.left,
							  rcClient.top,
							  rcClient.right,
							  rcClient.bottom);
			m_pBand->Draw(hDCOff, rcClient, ddBTPopup == m_pBand->bpV1.m_btBands ? VARIANT_TRUE : m_pBand->IsWrappable(), pt);
		}

		m_pBand->m_bPopupWinLock = FALSE;

		if (hDCOff != hDC)
		{
			if (ddMSAnimateNone == m_pBand->m_pBar->bpV1.m_msMenuStyle || 
				!m_bFirstTimePaint ||
				m_pBand->m_pBar->m_bCustomization || 
				m_pBand->m_pBar->m_bWhatsThisHelp ||
				(m_pBand->m_pBar->m_bFirstPopup && VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook))
			{
				//
				// Paint not animated
				//

				pff->Paint(hDC, 0, 0);
			}
			else 
			{
				//
				// Paint animated
				//
				m_pBand->m_pBar->m_bFirstPopup = TRUE;
				MenuStyles msMenuStyle = m_pBand->m_pBar->bpV1.m_msMenuStyle;

				switch (m_nHorzFlip)
				{
				case ddPopupMenuCenterAlign:
					msMenuStyle = ddMSAnimateUnfold;
					break;

				case ddPopupMenuRightAlign:
					if (ddMSAnimateSlide == msMenuStyle)
						m_bFlipHorz = TRUE;
					break;

				case ddPopupMenuLeftAlign:
				default:
					{
						if (ddMSAnimateRandom == msMenuStyle)
						{
							if (rand() < SHRT_MAX / 2)
								msMenuStyle = ddMSAnimateUnfold;
							else
								msMenuStyle = ddMSAnimateSlide;
						}
					}
					break;
				}

				int nXDelta = nWidthBitmap / 6;
				int nYDelta = nHeightBitmap / 6;
				int nY = 0;
				int nX = 0;
				DWORD dwTickCountStart;
				DWORD dwPaintTime;
				BOOL bSliding = ddMSAnimateSlide == msMenuStyle ? TRUE : FALSE;
				if (!m_bFlipHorz && !m_bFlipVert)
				{
					for (nX = nXDelta; nX < nWidthBitmap; nX += nXDelta)
					{
						dwTickCountStart = GetTickCount(); 
						nY += nYDelta;
						pff->Paint(hDC,
								   0, 
								   0, 
								   bSliding ? nX : nWidthBitmap, 
								   nY, 
								   bSliding ? nWidthBitmap - nX : 0, 
								   nHeightBitmap - nY);
						dwPaintTime = GetTickCount() - dwTickCountStart;
						if (dwPaintTime < m_dwWaitTime)
							Sleep(m_dwWaitTime - dwPaintTime);
					}
				}
				else if (!m_bFlipHorz && m_bFlipVert)
				{
					nY = nHeightBitmap;
					for (nX = nXDelta; nX < nWidthBitmap; nX += nXDelta)
					{
						dwTickCountStart = GetTickCount(); 
						nY -= nYDelta;
						pff->Paint(hDC,
								   0, 
								   nY,
								   bSliding ? nX : nWidthBitmap, 
								   nHeightBitmap, 
								   bSliding ? nWidthBitmap - nX : 0, 
								   nY);
						dwPaintTime = GetTickCount() - dwTickCountStart;
						if (dwPaintTime < m_dwWaitTime)
							Sleep(m_dwWaitTime - dwPaintTime);
					}
				}
				else if (m_bFlipHorz && !m_bFlipVert)
				{
					for (nX = nWidthBitmap - nXDelta; nX > 0; nX -= nXDelta)
					{
						dwTickCountStart = GetTickCount(); 
						nY += nYDelta;
						pff->Paint(hDC,
								   nX, 
								   0, 
								   nWidthBitmap, 
								   bSliding ? nY : nHeightBitmap, 
								   nX, 
								   bSliding ? nHeightBitmap - nY : 0);
						dwPaintTime = GetTickCount() - dwTickCountStart;
						if (dwPaintTime < m_dwWaitTime)
							Sleep(m_dwWaitTime - dwPaintTime);
					}
				}
				else
				{
					nY = nHeightBitmap;
					for (nX = nWidthBitmap - nXDelta; nX > 0; nX -= nXDelta)
					{
						dwTickCountStart = GetTickCount(); 
						nY -= nYDelta;
						pff->Paint(hDC,
								   nX, 
								   nY, 
								   nWidthBitmap, 
								   nHeightBitmap, 
								   nX, 
								   nY);
						dwPaintTime = GetTickCount() - dwTickCountStart;
						if (dwPaintTime < m_dwWaitTime)
							Sleep(m_dwWaitTime - dwPaintTime);
					}
				}

				if (nX != nWidthBitmap || nY != nHeightBitmap)
				{
					pff->Paint(hDC,
							   0, 
							   0, 
							   nWidthBitmap, 
							   nHeightBitmap, 
							   0, 
							   0);
				}
			}
		}

		if (ddBFDetach & m_pBand->bpV1.m_dwFlags && !m_bPopupMenu)
			PaintDetachBar(hDC);

		m_bFirstTimePaint = FALSE;
		SelectClipRgn(hDCOff, NULL);
		m_pBand->ParentWindowedTools(m_hWnd);

		if (hPal)
			SelectPalette(hDCOff, hPalOld, FALSE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// PaintDetachBar
//

void CPopupWin::PaintDetachBar(HDC hDC)
{
	BOOL bResult;
	CRect rc;
	GetClientRect(rc);
	if (VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook)
	{
		rc.Inflate(-eDetachCaptionBorderThickness, -eDetachCaptionBorderThickness);
		rc.bottom = rc.top + CBand::eDetachBarHeight - eDetachCaptionBorderAdjustment;
		rc.left += eDetachCaptionBorderAdjustment;
		rc.right -= eDetachCaptionBorderAdjustment;
		FillRect(hDC, &rc, eDetachBandCaptionSelected == m_nCurSel ? (HBRUSH)(1+COLOR_ACTIVECAPTION) : (HBRUSH)(1+COLOR_INACTIVECAPTION));
	}
	else
	{
		BOOL bSelected = eDetachBandCaptionSelected == m_nCurSel;
		rc.Inflate(-2, -2);
		rc.left++;
		rc.right--;
		rc.bottom = rc.top + CBand::eDetachBarHeight;

		HPEN penBack = CreatePen(PS_SOLID, 1, bSelected ? m_pBand->m_pBar->m_crXPSelectedBorderColor : m_pBand->m_pBar->m_crXPBandBackground);
		if (penBack)
		{
			HPEN penOld = SelectPen(hDC, penBack);
			
			HBRUSH hBrush = CreateSolidBrush(eDetachBandCaptionSelected == m_nCurSel ? m_pBand->m_pBar->m_crXPSelectedColor : m_pBand->m_pBar->m_crXPBandBackground);
			if (hBrush)
			{
				HBRUSH brushOld = SelectBrush(hDC, hBrush);

				RoundRect(hDC, rc.left, rc.top, rc.right, rc.bottom, 1, 1);

				SelectBrush(hDC, brushOld);
				bResult = DeleteBrush(hBrush);
				assert(bResult);

				HPEN penBack2 = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_APPWORKSPACE));
				if (penBack2)
				{
					HPEN penOld2 = SelectPen(hDC, penBack2);
					int nOffset = (rc.Width() - 33) / 2;
					rc.left += nOffset;
					rc.right = rc.left + 33;

					if (bSelected)
					{
						rc.top += 3;
						MoveToEx(hDC, rc.left, rc.top, NULL);
						LineTo(hDC, rc.right, rc.top);
						rc.top += 2;
						MoveToEx(hDC, rc.left, rc.top, NULL);
						LineTo(hDC, rc.right, rc.top);
					}
					else
					{
						rc.top += 2;
						MoveToEx(hDC, rc.left, rc.top, NULL);
						LineTo(hDC, rc.right, rc.top);
						rc.top += 2;
						MoveToEx(hDC, rc.left, rc.top, NULL);
						LineTo(hDC, rc.right, rc.top);
						rc.top += 2;
						MoveToEx(hDC, rc.left, rc.top, NULL);
						LineTo(hDC, rc.right, rc.top);
					}

					SelectPen(hDC, penOld2);
					bResult = DeletePen(penBack2);
					assert(bResult);
				}


			}
			SelectPen(hDC, penOld);
			bResult = DeletePen(penBack);
			assert(bResult);
		}
	}
}

//
// HitTestDetach
//

BOOL CPopupWin::HitTestDetach(POINT pt)
{
	if (m_bPopupMenu)
		return FALSE;

	if (!(ddBFDetach & m_pBand->bpV1.m_dwFlags))
		return FALSE;

	CRect rcClient;
	GetClientRect(rcClient);
	rcClient.bottom = rcClient.top + CBand::eDetachBarHeight;
	if (PtInRect(&rcClient, pt))
		return TRUE;
	return FALSE;
}

//
// OnMouseMove
//

void CPopupWin::OnMouseMove(POINT pt)
{
	if (Customization())
		return;

	if (HitTestDetach(pt))
	{
		if (ddBTPopup == m_pBand->bpV1.m_btBands)
			SetCurSel(eDetachBandCaptionSelected);
		else
			SetCurSelNormalBand(eDetachBandCaptionSelected);
		return;
	}

	if (m_pBand && m_pBand->m_pParentPopupWin && m_pBand->m_pParentPopupWin != this && ddBTPopup == m_pBand->m_pParentPopupWin->m_pBand->bpV1.m_btBands && m_pBand->m_pParentPopupWin->m_nCurSel != m_nParentIndex)
		m_pBand->m_pParentPopupWin->SetCurSel(m_nParentIndex);
	if (NULL == m_pBand->m_prcTools)
		return;

	CTool* pTool = NULL;
	CRect rcTool;
	int nToolCount = 0;
	int nTool = 0;

	CRect rcClient;
	GetClientRect(rcClient);

	BOOL bDetachable = ddBFDetach & m_pBand->bpV1.m_dwFlags && !m_bPopupMenu;
	if (bDetachable)
		rcClient.top += CBand::eDetachBarHeight;

	// ???????
	if (!PtInRect(&rcClient, pt))
		goto OutOfClientArea;

	// ???????
	if (m_bScrollActive)
	{
		if (!PtInRect(&m_rcClient, pt))
			goto OutOfClientArea;
	}

	--pt.x;
	--pt.y;

	nToolCount = m_pBand->m_pTools->GetVisibleToolCount();
	for (nTool = 0; nTool < nToolCount; nTool++)
	{
		rcTool = m_pBand->m_prcTools[nTool];
		if (m_bScrollActive)
			rcTool.Offset(0, -m_nPhysicalScrollPos);

		if (!PtInRect(&rcTool, pt))
			continue;

		pTool = m_pBand->m_pTools->GetVisibleTool(nTool);
		if (pTool)
		{
			if (ddTTCombobox == pTool->tpV1.m_ttTools && !pTool->IsComboReadOnly())
			{
				CRect rcTool;
				if (m_pBand->GetToolRect(nTool, rcTool))
				{
					rcTool.left += pTool->m_nComboNameOffset;
					rcTool.right -= 10;
					if (PtInRect(&rcTool, pt))
						SetCursor(LoadCursor(NULL, IDC_IBEAM));
					else
						SetCursor(LoadCursor(NULL, IDC_ARROW));
				}
				else
					SetCursor(LoadCursor(NULL, IDC_ARROW));

			}
			else if (ddTTEdit == pTool->tpV1.m_ttTools)
			{
				rcTool.left += pTool->m_nComboNameOffset;
				if (PtInRect(&rcTool, pt))
					SetCursor(LoadCursor(NULL, IDC_IBEAM));
				else
					SetCursor(LoadCursor(NULL, IDC_ARROW));
			}
			else
				SetCursor(LoadCursor(NULL, IDC_ARROW));

			if (nTool != m_nCurSel)
			{
				if (VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
				{
					if (ddBTPopup == m_pBand->bpV1.m_btBands)
						SetCurSel(nTool);
					else
						SetCurSelNormalBand(nTool);
					if (!m_bTrackTimerStarted)
					{
						if (CTool::ddTTMenuExpandTool == pTool->tpV1.m_ttTools && ddPMDisplayOnHover == m_pBand->m_pBar->bpV1.m_pmMenus)
						{
							SetTimer(m_hWnd, eExpandTimer, eExpandTimerDuration, NULL);
							m_bTrackTimerStarted = TRUE;
						}
					}
					else if (m_bTrackTimerStarted && CTool::ddTTMenuExpandTool != pTool->tpV1.m_ttTools)
					{
						KillTimer(m_hWnd, eExpandTimer);
						m_bTrackTimerStarted = FALSE;
					}
				}
				else
				{
					if (ddBTPopup == m_pBand->bpV1.m_btBands)
						SetCurSel(eRemoveSelection);
					else
						SetCurSelNormalBand(eRemoveSelection);
					try
					{
						if (pTool && VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
						{
							m_pBand->m_pBar->FireMouseEnter((Tool*)pTool);
							m_pBand->m_pBar->FireMenuItemEnter((Tool*)pTool);
						}
					}
					CATCH
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
					if (m_pBand->m_pBar->m_bMenuLoop)
						m_pBand->m_pBar->StatusBandUpdate(pTool);
				}
			}
			
			//
			// Start timer if not already started
			//

			if (pTool->HasSubBand() && m_nCurSel != m_nCurPopup && VARIANT_TRUE == pTool->tpV1.m_vbEnabled && NULL == m_pBand->m_pChildPopupWin && !m_bTimerStarted)
			{
				m_bTimerStarted = TRUE;
				SetTimer(m_hWnd, eOpenSubTimer, eOpenSubTimerDuration, NULL);
			}
			else if (m_nCurSel == m_nCurPopup)
			{
				KillTimer(m_hWnd, eCloseSubTimer);
			}
			return;
		}
	}

OutOfClientArea:
	if (NULL == pTool)
		SetCursor(LoadCursor(NULL, IDC_ARROW));

	// No SubPopup opened yet
	if (eRemoveSelection != m_nCurSel && NULL == m_pBand->m_pChildPopupWin) 
	{
		if (ddBTPopup == m_pBand->bpV1.m_btBands)
			SetCurSel(eRemoveSelection);
		else
			SetCurSelNormalBand(eRemoveSelection);
	}
}

//
// SetCurSelNormalBand
//

void CPopupWin::SetCurSelNormalBand(int nNewSel)
{
	if (ddBTPopup == m_pBand->bpV1.m_btBands)
	{
		assert(FALSE);
		return;
	}
	
	TRACE(5, _T("SetCurSelNormalBand\n"));
	int nToolCount = m_pBand->m_pTools->GetVisibleToolCount();
	if (nNewSel >= nToolCount) // defensive
		return;

	CTool* pTool;
	if (m_nCurSel >= 0)
	{
		// Send mouse exit event defensive
		if (m_nCurSel < nToolCount) 
		{
			pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
			if (pTool)
			{
				try
				{
					m_pBand->m_pBar->FireMenuItemExit((Tool*)pTool);
					m_pBand->m_pBar->FireMouseExit((Tool*)pTool);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
		}
	}

	if (m_pBand->m_pChildPopupWin)
	{
		try
		{
			SetTimer(m_hWnd, eCloseSubTimer, eCloseSubTimerDuration, NULL);
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__);
		}
	}

	HDC hDC = GetDC(m_hWnd);
	if (hDC)
	{
		HBRUSH hbrTexture = NULL;
		BOOL bHasTexture = m_pBand->m_pBar->HasTexture();
		if (bHasTexture)
			hbrTexture = m_pBand->m_pBar->m_hBrushTexture;
		
		CRect rcTool,rcOut;
		CFlickerFree ff;
		if (m_nCurSel >= 0) 
		{
			//
			// Erase old tool
			//

			pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
			if (pTool && ddTTSeparator != pTool->tpV1.m_ttTools)
			{
				m_pBand->m_pBar->m_pActiveTool = NULL;

				rcTool = m_pBand->m_prcTools[m_nCurSel];
				rcTool.Offset(3, 3);

				if (m_bScrollActive)
					rcTool.Offset(0, -m_nPhysicalScrollPos);

				rcTool.Inflate(2, 2);
				rcOut = rcTool;
				HDC hDCOut = ff.RequestDC(hDC, rcTool.Width(), rcTool.Height());
				if (NULL == hDCOut)
					hDCOut = hDC;
				else
					rcOut.Offset(-rcOut.left, -rcOut.top);

				if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
				{
					if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved || pTool->m_bMFU)
						FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crXPMenuBackground);
					else
						FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crXPBackground);
				}
				else
				{
					if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved || pTool->m_bMFU)
					{
						if (bHasTexture && ddBOPopups & m_pBand->m_pBar->bpV1.m_dwBackgroundOptions)
						{
							if (hbrTexture)
								UnrealizeObject(hbrTexture);
							
							SetBrushOrgEx(hDCOut, -rcTool.left, -rcTool.top, NULL);
							
							m_pBand->m_pBar->FillTexture(hDCOut, rcOut);
						}
						else
							FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crBackground);
					}
					else
						FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crMDIMenuBackground);
				}

				if (ddBTPopup == m_pBand->bpV1.m_btBands)
					m_pBand->m_bPopupWinLock = TRUE;

				if (ddTSColor == m_pBand->bpV1.m_tsMouseTracking)
					pTool->m_bGrayImage = TRUE;

				rcOut.Inflate(-2, -2);

				pTool->Draw(hDCOut, rcOut, m_pBand->bpV1.m_btBands, FALSE, FALSE);

				if (ddTSColor == m_pBand->bpV1.m_tsMouseTracking)
					pTool->m_bGrayImage = FALSE;
				
				if (hDC != hDCOut)
					ff.Paint(hDC, rcTool.left, rcTool.top);

				switch (pTool->tpV1.m_ttTools)
				{
				case ddTTForm:
				case ddTTControl:
					if (::IsWindow(pTool->m_hWndActive))
					{
						::SetWindowPos(pTool->m_hWndActive, 
									   NULL, 
									   rcTool.left, 
									   rcTool.top, 
									   0, 
									   0, 
									   SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
					}
					break;
				}

				m_nCurSel = eRemoveSelection;
				m_pBand->m_pPopupToolSelected = NULL;
			}
		}
		if (eDetachBandCaptionSelected == nNewSel || eDetachBandCaptionSelected == m_nCurSel)
		{
			m_nCurSel = nNewSel;
			PaintDetachBar(hDC);
		}
		m_nCurSel = nNewSel;
			
		if (nNewSel >= 0)
		{
			pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);

			if (pTool && ddTTSeparator != pTool->tpV1.m_ttTools)
			{
				m_pBand->m_pBar->m_pActiveTool = pTool;
				rcTool = m_pBand->m_prcTools[m_nCurSel];
				rcTool.Offset(3, 3);

				if (m_bScrollActive)
					rcTool.Offset(0, -m_nPhysicalScrollPos);

				rcTool.Inflate(2, 2);
				rcOut = rcTool;
				HDC hDCOut = ff.RequestDC(hDC, rcTool.Width(), rcTool.Height());
				if (NULL == hDCOut)
					hDCOut = hDC;
				else
					rcOut.Offset(-rcOut.left, -rcOut.top);

				if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
				{
					if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved || pTool->m_bMFU)
						FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crXPMenuBackground);
					else
						FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crXPBackground);
				}
				else
				{
					if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved || pTool->m_bMFU)
					{
						if (bHasTexture && ddBOPopups & m_pBand->m_pBar->bpV1.m_dwBackgroundOptions)
						{
							if (hbrTexture)
								UnrealizeObject(hbrTexture);
							
							SetBrushOrgEx(hDCOut, -rcTool.left, -rcTool.top, NULL);
							
							m_pBand->m_pBar->FillTexture(hDCOut, rcOut);
						}
						else
							FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crBackground);
					}
					else
						FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crMDIMenuBackground);
				}

				rcOut.Inflate(-2, -2);

				pTool->Draw(hDCOut, rcOut, m_pBand->bpV1.m_btBands, FALSE, FALSE);
				if (hDC != hDCOut)
					ff.Paint(hDC, rcTool.left, rcTool.top);

				switch (pTool->tpV1.m_ttTools)
				{
				case ddTTForm:
				case ddTTControl:
					if (::IsWindow(pTool->m_hWndActive))
					{
						::SetWindowPos(pTool->m_hWndActive, 
									   NULL, 
									   rcTool.left, 
									   rcTool.top, 
									   0, 
									   0, 
									   SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
					}
					break;
				}
			}
		}
		ReleaseDC(m_hWnd, hDC);
	}
	m_pBand->m_bPopupWinLock = FALSE;

	if (m_nCurSel >= 0 && m_nCurSel < m_pBand->m_pTools->GetVisibleToolCount())
	{
		// Send mouse exit event defensive
		pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
		if (pTool)
		{
			try
			{
				m_pBand->m_pBar->FireMouseEnter((Tool*)pTool);
				m_pBand->m_pBar->FireMenuItemEnter((Tool*)pTool);
				if (m_pBand->m_pBar->m_bMenuLoop)
					m_pBand->m_pBar->StatusBandUpdate(pTool);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
	}
}

//
// SetCurSel
//

void CPopupWin::SetCurSel(int nNewSel)
{
	if (ddBTPopup != m_pBand->bpV1.m_btBands)
	{
		assert(FALSE);
		return;
	}
	
	int nToolCount = m_pBand->m_pTools->GetVisibleToolCount();
	if (nNewSel >= nToolCount) // defensive
		return;

	CTool* pTool;
	if (m_nCurSel >= 0)
	{
		// Send mouse exit event defensive
		if (m_nCurSel < nToolCount) 
		{
			pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
			if (pTool && pTool)
			{
				try
				{
					m_pBand->m_pBar->FireMenuItemExit((Tool*)pTool);
					m_pBand->m_pBar->FireMouseExit((Tool*)pTool);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
		}
	}

	if (m_pBand->m_pChildPopupWin)
	{
		try
		{
			SetTimer(m_hWnd, eCloseSubTimer, eCloseSubTimerDuration, NULL);
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__);
		}
	}

	HDC hDC = GetDC(m_hWnd);
	if (hDC)
	{
		HRGN hRgn = NULL; 
		HRGN hRgnPrev = NULL;
		int nResult;
		if (m_bScrollActive)
		{
			hRgn = CreateRectRgnIndirect(&m_rcClient); 
			assert(hRgn);
			if (hRgn)
			{
				hRgnPrev = CreateRectRgn(0,0,0,0);
				if (hRgnPrev)
					nResult = GetClipRgn(hDC, hRgnPrev);

				nResult = SelectClipRgn(hDC, hRgn); 
				assert(ERROR != nResult);
			}
		}

		BOOL bHasTexture = m_pBand->m_pBar->HasTexture();
		HBRUSH hbrTexture = m_pBand->m_pBar->m_hBrushTexture;
		
		CRect rcTool,rcOut;
		CFlickerFree ff;
		if (m_nCurSel >= 0) 
		{
			//
			// Erase old tool
			//

			pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
			if (pTool && ddTTSeparator != pTool->tpV1.m_ttTools)
			{
				pTool->m_bPressed = FALSE;
				m_pBand->m_pPopupToolSelected = NULL;
				rcTool = m_pBand->m_prcTools[m_nCurSel];
				rcTool.Offset(3, 3);
				if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
					rcTool.Inflate(1, 1);

				if (m_bScrollActive)
					rcTool.Offset(0, -m_nPhysicalScrollPos);

				rcOut = rcTool;
				HDC hDCOut = ff.RequestDC(hDC, rcTool.Width(), rcTool.Height());
				if (NULL == hDCOut)
					hDCOut = hDC;
				else
					rcOut.Offset(-rcOut.left, -rcOut.top);

				if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
				{
					FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crXPMenuBackground);
					CRect rcBorder = rcOut;
					rcBorder.right = rcBorder.left + m_pBand->m_sizeMaxIcon.cx + 5;
					if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved || pTool->m_bMFU)
						FillSolidRect(hDCOut, rcBorder, m_pBand->m_pBar->m_crXPBandBackground);
					else
						FillSolidRect(hDCOut, rcBorder, m_pBand->m_pBar->m_crXPMRUBackground);
				}
				else
				{
					if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved || pTool->m_bMFU)
					{
						if (bHasTexture && ddBOPopups & m_pBand->m_pBar->bpV1.m_dwBackgroundOptions)
						{
							if (hbrTexture)
								UnrealizeObject(hbrTexture);
							
							SetBrushOrgEx(hDCOut, -rcTool.left, -rcTool.top, NULL);
							
							m_pBand->m_pBar->FillTexture(hDCOut, rcOut);
						}
						else
							FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crBackground);
					}
					else
						FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crMDIMenuBackground);
				}

				if (ddBTPopup == m_pBand->bpV1.m_btBands)
					m_pBand->m_bPopupWinLock = TRUE;

				if (ddTSColor == m_pBand->bpV1.m_tsMouseTracking)
					pTool->m_bGrayImage = TRUE;

				pTool->Draw(hDCOut, rcOut, m_pBand->bpV1.m_btBands, FALSE, FALSE);

				if (ddTSColor == m_pBand->bpV1.m_tsMouseTracking)
					pTool->m_bGrayImage = FALSE;
				
				if (hDC != hDCOut)
					ff.Paint(hDC, rcTool.left, rcTool.top);

				switch (pTool->tpV1.m_ttTools)
				{
				case ddTTForm:
				case ddTTControl:
					if (::IsWindow(pTool->m_hWndActive))
					{
						::SetWindowPos(pTool->m_hWndActive, 
									   NULL, 
									   rcTool.left, 
									   rcTool.top, 
									   0, 
									   0, 
									   SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
					}
					break;
				}

				m_nCurSel = eRemoveSelection;
			}
		}
		if (eDetachBandCaptionSelected == nNewSel || eDetachBandCaptionSelected == m_nCurSel)
		{
			m_nCurSel = nNewSel;
			PaintDetachBar(hDC);
		}
		m_nCurSel = nNewSel;
			
		if (nNewSel >= 0)
		{
			pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
			if (pTool && ddTTSeparator != pTool->tpV1.m_ttTools)
			{
				m_pBand->m_pPopupToolSelected = pTool;
				pTool->m_bPressed = TRUE;
				rcTool = m_pBand->m_prcTools[m_nCurSel];
				rcTool.Offset(3, 3);
				if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
					rcTool.Inflate(1, 1);

				if (m_bScrollActive)
					rcTool.Offset(0, -m_nPhysicalScrollPos);

				rcOut = rcTool;
				HDC hDCOut = ff.RequestDC(hDC, rcTool.Width(), rcTool.Height());
				if (NULL == hDCOut)
					hDCOut = hDC;
				else
					rcOut.Offset(-rcOut.left, -rcOut.top);

				if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
				{
					FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crXPMenuBackground);
					CRect rcBorder = rcOut;
					rcBorder.right = rcBorder.left + m_pBand->m_sizeMaxIcon.cx + 5;
					if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved || pTool->m_bMFU)
						FillSolidRect(hDCOut, rcBorder, m_pBand->m_pBar->m_crXPBandBackground);
					else
						FillSolidRect(hDCOut, rcBorder, m_pBand->m_pBar->m_crXPBackground);
				}
				else 
				{
					if (!m_bMFU || !m_pBand->m_pBar->m_bPopupMenuExpanded || !m_pBand->m_bToolRemoved || pTool->m_bMFU)
					{
						if (bHasTexture && ddBOPopups & m_pBand->m_pBar->bpV1.m_dwBackgroundOptions)
						{
							if (hbrTexture)
								UnrealizeObject(hbrTexture);
							
							SetBrushOrgEx(hDCOut, -rcTool.left, -rcTool.top, NULL);
							
							m_pBand->m_pBar->FillTexture(hDCOut, rcOut);
						}
						else
							FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crBackground);
					}
					else
						FillSolidRect(hDCOut, rcOut, m_pBand->m_pBar->m_crMDIMenuBackground);
				}

				pTool->Draw(hDCOut, rcOut, m_pBand->bpV1.m_btBands, FALSE, FALSE);
				if (hDC != hDCOut)
					ff.Paint(hDC, rcTool.left, rcTool.top);

				switch (pTool->tpV1.m_ttTools)
				{
				case ddTTForm:
				case ddTTControl:
					if (::IsWindow(pTool->m_hWndActive))
					{
						::SetWindowPos(pTool->m_hWndActive, 
									   NULL, 
									   rcTool.left, 
									   rcTool.top, 
									   0, 
									   0, 
									   SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
					}
					break;
				}
			}
		}
		if (hRgn)
		{
			if (-1 == nResult)
				SelectClipRgn(hDC, NULL); 
			else
				SelectClipRgn(hDC, hRgnPrev);
			
			BOOL bResult = DeleteRgn(hRgn);
			assert(bResult);
		}

		if (hRgnPrev)
		{
			BOOL bResult = DeleteRgn(hRgnPrev);
			assert(bResult);
		}
		ReleaseDC(m_hWnd, hDC);
	}
	m_pBand->m_bPopupWinLock = FALSE;
	if (m_nCurSel >= 0 && m_nCurSel < m_pBand->m_pTools->GetVisibleToolCount())
	{
		// Send mouse exit event defensive
		pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
		if (pTool)
		{
			try
			{
				m_pBand->m_pBar->FireMouseEnter((Tool*)pTool);
				m_pBand->m_pBar->FireMenuItemEnter((Tool*)pTool);
				if (m_pBand->m_pBar->m_bMenuLoop)
					m_pBand->m_pBar->StatusBandUpdate(pTool);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
	}
}

//
// OnKey
//

void CPopupWin::OnKey(UINT nKey, BOOL* pbUsed)
{
	if (Customization())
		return;

	assert (m_pBand);

	if (m_pBand->m_pChildPopupWin)
	{
		m_pBand->AddRef();
		if (!m_pBand->m_pChildPopupWin->IsWindow())
			return;

		LRESULT lResult = m_pBand->m_pChildPopupWin->SendMessage(WM_KEYDOWN, nKey, (LPARAM)pbUsed);
		if (*pbUsed)
		{
			if (VK_RETURN == nKey && NULL == m_pBand->m_pChildPopupWin && IsWindow())
			{
				ReparentEditAndOrCombo();
				DestroyWindow();
			}
			m_pBand->Release();
			return;
		}
		else if (VK_LEFT == nKey)
		{
			if (m_pBand->m_pChildPopupWin->IsWindow())
			{
				m_pBand->m_pChildPopupWin->ReparentEditAndOrCombo();
				m_pBand->m_pChildPopupWin->DestroyWindow();
			}
			m_pBand->m_pChildPopupWin = NULL;
			*pbUsed = TRUE;
			m_pBand->Release();
			return;
		}
		else if (VK_RIGHT == nKey && eRemoveSelection == m_pBand->m_pChildPopupWin->m_nCurSel) 
		{
			// special handling for sub popups who have nothing selected
			m_pBand->m_pChildPopupWin->SetCurSel(0);
			*pbUsed = TRUE;
			m_pBand->Release();
			return;
		}
		m_pBand->Release();
	}

	if (VK_TAB == nKey)
	{
		if (GetKeyState(VK_SHIFT) & 0x8000)
			nKey = VK_UP;
		else
			nKey = VK_DOWN;
	}

	int nToolCount = m_pBand->m_pTools->GetVisibleToolCount();
	int nNewSel = m_nCurSel;
	try
	{
		CRect rcClient;
		if (m_bScrollActive)
		{
			GetClientRect(rcClient);
			rcClient.Inflate(-3, -3);
		}
		switch (nKey)
		{
		case VK_UP:
			--nNewSel;
			while (nNewSel >= 0 && ddTTSeparator == m_pBand->m_pTools->GetVisibleTool(nNewSel)->tpV1.m_ttTools)
			{
				--nNewSel;
				if (m_bScrollActive && m_bTopScrollActive)
				{
					CRect rcTool = m_pBand->m_prcTools[nNewSel];
					rcTool.Offset(0, -m_nPhysicalScrollPos);
					if (rcTool.top <= rcClient.top)
					{
						m_nScrollPos--;
						m_nPhysicalScrollPos = m_pBand->m_prcTools[m_nScrollPos].top + 3;
					}
				}
			}

			if (nNewSel < 0)
			{
				if (!m_bScrollActive)
				{
					nNewSel = nToolCount - 1; 
					while (nNewSel > 0 && ddTTSeparator == m_pBand->m_pTools->GetVisibleTool(nNewSel)->tpV1.m_ttTools)
						--nNewSel;
				}
				else
					nNewSel = 0;
			}

			*pbUsed = TRUE;
			if (nNewSel != m_nCurSel)
			{
				if (m_bScrollActive && m_bTopScrollActive)
				{
					CRect rcTool = m_pBand->m_prcTools[nNewSel];
					rcTool.Offset(0, -m_nPhysicalScrollPos);
					if (rcTool.top <= rcClient.top)
					{
						SetCurSel(eRemoveSelection);
						if (m_nScrollPos > 0)
							m_nScrollPos--;
						if (IsWindow())
						{
							InvalidateRect(NULL, FALSE);
							UpdateWindow();
						}
						SetCurSel(nNewSel);
					}
					else
						SetCurSel(nNewSel);
				}
				else
					SetCurSel(nNewSel);
			}
			return;

		case VK_DOWN:
			++nNewSel;
			while (nNewSel < nToolCount && ddTTSeparator == m_pBand->m_pTools->GetVisibleTool(nNewSel)->tpV1.m_ttTools)
			{
				++nNewSel;
				if (m_bScrollActive)
				{
					CRect rcTool = m_pBand->m_prcTools[nNewSel];
					rcTool.Offset(0, -m_nPhysicalScrollPos);

					if (rcTool.bottom > rcClient.bottom)
					{
						m_nScrollPos++;
						m_nPhysicalScrollPos = m_pBand->m_prcTools[m_nScrollPos].top + 3;
					}
				}
			}

			// wrap down
			if (nNewSel >= nToolCount) 
			{
				if (!m_bScrollActive)
					nNewSel = 0;
				else
					nNewSel = nToolCount - 1;
			}

			*pbUsed = TRUE;
			if (nNewSel != m_nCurSel)
			{
				if (m_bScrollActive && m_bBottomScrollActive)
				{
					CRect rcTool = m_pBand->m_prcTools[nNewSel];
					rcTool.Offset(0, -m_nPhysicalScrollPos);

					if (rcTool.bottom >= rcClient.bottom)
					{
						if (m_nScrollPos < (nToolCount-1))
						{
							m_nScrollPos++;
							if (rcTool.top >= rcClient.bottom && m_nScrollPos < (nToolCount-1))
								m_nScrollPos++;
						}
						SetCurSel(eRemoveSelection);
						if (IsWindow())
						{
							InvalidateRect(NULL, FALSE);
							UpdateWindow();
						}
						SetCurSel(nNewSel);
					}
					else
						SetCurSel(nNewSel);
				}
				else
				{
					SetCurSel(nNewSel);
					CTool* pTool = m_pBand->m_pTools->GetVisibleTool(nNewSel);
					if (pTool && CTool::ddTTMenuExpandTool == pTool->tpV1.m_ttTools)
					{
						m_pBand->m_pBar->m_bPopupMenuExpanded = TRUE;
						pTool->tpV1.m_vbVisible = VARIANT_FALSE;
						m_nCurSel = -1;
						m_pBand->m_pPopupToolSelected = NULL;
						if (GetOptimalRect(m_rcClientOrginal))
						{
							if (m_bFlipVert)
							{
								CRect rcWindow;
								GetWindowRect(rcWindow);
								SetWindowPos(NULL, 
											 rcWindow.left, 
											 rcWindow.top + 12, 
											 m_rcClientOrginal.Width(), 
											 m_rcClientOrginal.Height(), 
											 SWP_NOACTIVATE);
							}
							else
							{
								SetWindowPos(NULL, 
											 0, 
											 0, 
											 m_rcClientOrginal.Width(), 
											 m_rcClientOrginal.Height(), 
											 SWP_NOREPOSITION|SWP_NOMOVE|SWP_NOACTIVATE);
								InvalidateRect(NULL, FALSE);
							}
						}
					}
				}
			}
			return;

		case VK_RIGHT:
			if (NULL == m_pBand->m_pChildPopupWin && eRemoveSelection != m_nCurSel && 
				m_pBand->m_pTools->GetVisibleTool(m_nCurSel)->HasSubBand() && 
				VARIANT_TRUE == m_pBand->m_pTools->GetVisibleTool(m_nCurSel)->tpV1.m_vbEnabled)
			{
				*pbUsed = TRUE;
				OpenSubPopup();
				if (m_pBand->m_pChildPopupWin)
					m_pBand->m_pChildPopupWin->m_nCurSel = 0;
			}
			return;

		case VK_RETURN:
			if (eRemoveSelection != m_nCurSel)
			{
				CTool* pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
				if (pTool)
				{
					if (pTool->HasSubBand())
					{
						OpenSubPopup();
						if (m_pBand->m_pChildPopupWin)
							m_pBand->m_pChildPopupWin->m_nCurSel = 0;
					}
					else
					{
						FireDelayedClick(pTool);
						ReleaseCapture();
						if (IsWindow())
							PostMessage(WM_CLOSE);
					}
				}
				*pbUsed = TRUE;
			}
			else if (eRemoveSelection == m_nCurSel)
			{
				POINT pt;
				GetCursorPos(&pt);
				CRect rcWindow;
				GetWindowRect(rcWindow);
				if (!PtInRect(&rcWindow, pt))
				{
					if (m_pBand->m_pBar->m_pPopupRoot)
						m_pBand->m_pBar->m_pPopupRoot->SetPopupIndex(eRemoveSelection);
					*pbUsed = TRUE;
					if (IsWindow())
						PostMessage(WM_CLOSE);
				}
			}
			return;

		default:
			if (m_nCurSel > -1 && m_nCurSel < m_pBand->m_pTools->GetVisibleToolCount())
			{
				CTool* pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
				if (pTool)
				{
					switch (pTool->tpV1.m_ttTools)
					{
					case ddTTControl:
					case ddTTForm:
						*pbUsed = TRUE;
						return;
						break;
					}
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

	// Check if menu shortcut key

	int nIndex = m_pBand->CheckMenuKey((WORD)nKey);
	if (eRemoveSelection == nIndex)
	{
		if (0 == (GetKeyState(VK_CONTROL)&0x8000) && 
			0 == (GetKeyState(VK_SHIFT)&0x8000) &&
			0 == (GetKeyState(VK_MENU)&0x8000) && 
			nKey >= 'A' && nKey <= 'Z')
		{
			*pbUsed = TRUE;
		}
		return;
	}
	
	if (GetKeyState(VK_CONTROL)&0x8000 || GetKeyState(VK_SHIFT)&0x8000)
	{
		*pbUsed = TRUE;
		return;
	}

	// Press that tool (might open sub menu)
	CTool* pTool = m_pBand->m_pTools->GetVisibleTool(nIndex);
	if (pTool)
	{
		if (pTool->HasSubBand())
		{
			KillTimer(m_hWnd, eCloseSubTimer);
			SetCurSel(nIndex);
			OpenSubPopup();
			MSG msg;
			while (::PeekMessage(&msg, NULL, WM_PAINT, WM_PAINT, PM_NOREMOVE))
			{
				if (GetMessage(&msg, NULL, WM_PAINT, WM_PAINT))
					DispatchMessage(&msg);
			}
			if (m_pBand->m_pChildPopupWin)
				m_pBand->m_pChildPopupWin->SetCurSel(0);
		}
		else
		{
			ReleaseCapture();
			ReparentEditAndOrCombo();
			FireDelayedClick(pTool);
			if (IsWindow())
				DestroyWindow();
		}
	}
	*pbUsed = TRUE;
}

//
// OnButton
//
// 
// Returns true if detach success
//

BOOL CPopupWin::OnButton(UINT nMsg, BOOL bMouseUp, UINT nFlags) 
{
	POINT ptScreen;
	GetCursorPos(&ptScreen);
	POINT pt = ptScreen;
	::ScreenToClient(m_hWnd, &pt);

	if (m_pBand != m_pBand->m_pBar->m_pEditToolPopup && (m_pBand->m_pBar->m_bCustomization || m_pBand->m_pBar->m_bWhatsThisHelp))
	{
		if (WM_LBUTTONDOWN == nMsg || WM_RBUTTONDOWN == nMsg)
		{
			m_pBand->OnCustomMouseDown(nMsg, pt);
			return TRUE;
		}
		return FALSE;
	}

	CRect rcClient, rcTool;
	GetClientRect(rcClient);

	if (ddBFDetach & m_pBand->bpV1.m_dwFlags && !m_bPopupMenu)
	{
		rcTool = rcClient;
		rcTool.bottom = rcTool.top + CBand::eDetachBarHeight;
		if (PtInRect(&rcTool, pt))
		{
			POINT ptOffset = {0,0};
			ClientToScreen(ptOffset);
			rcTool.Offset(ptOffset.x, ptOffset.y);

			if (DetachTrack(rcTool))
			{
				m_pBand->m_pBar->m_nPopupExitAction = PEA_DETACH;
				m_pBand->m_pBar->m_ptPea = ptScreen;
				PopLayout();
				return TRUE;
			}
			return FALSE;
		}
		rcClient.top += CBand::eDetachBarHeight;
	}

	if (!PtInRect(&rcClient, pt))
	{
		ReparentEditAndOrCombo();
		DestroyWindow();
		return FALSE;
	}

	if (m_nCurSel >= 0)
	{
		CTool* pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
		if (pTool)
		{
			switch (pTool->tpV1.m_ttTools)
			{
			case ddTTSeparator:
				break;

			case CTool::ddTTMenuExpandTool:
				if (WM_LBUTTONUP == nMsg)
				{
					m_pBand->m_pBar->m_bPopupMenuExpanded = TRUE;
					pTool->tpV1.m_vbVisible = VARIANT_FALSE;
					m_nCurSel = -1;
					m_pBand->m_pPopupToolSelected = NULL;
					CRect rcWindow;
					GetWindowRect(rcWindow);
					if (CalcFlipingAndScrolling(rcWindow))
					{
						SetWindowPos(NULL, 
									 rcWindow.left, 
									 rcWindow.top, 
									 rcWindow.Width(), 
									 rcWindow.Height(), 
									 SWP_NOACTIVATE);
						InvalidateRect(NULL, FALSE);
					}
				}
				else
					TrackMenuExpandButton(pTool);
				break;

			default:
				if (!pTool->HasSubBand())
				{
					if (ddTTCombobox == pTool->tpV1.m_ttTools || ddTTEdit == pTool->tpV1.m_ttTools)
					{
						if (!bMouseUp)
						{
							rcTool = m_pBand->m_prcTools[m_nCurSel];
							m_pBand->m_nCurrentTool	= m_nCurSel;
							try
							{
								pTool->OnLButtonDown(0, pt);
							}
							catch (...)
							{
								assert(FALSE);
							}
							m_pBand->m_nCurrentTool	= eRemoveSelection;
						}
					}
					else
					{
						if (bMouseUp)
						{
							switch (pTool->tpV1.m_ttTools)
							{
							case CTool::ddTTAllTools:
								break;

							case ddTTControl:
							case ddTTForm:
								if (!m_pBand->m_pBar->m_bWhatsThisHelp)
									::SetFocus(pTool->m_hWndActive);
								break;

							default:
								int nCurSel = m_nCurSel;
								if (!m_pBand->m_bAllTools || CBar::eToolIdResetToolbar == pTool->tpV1.m_nToolId || CBar::eToolIdCustomize == pTool->tpV1.m_nToolId)
								{
									CTool* pTool2 = m_pBand->m_pBar->m_theToolStack.Pop();
									while (pTool2)
									{
										pTool2->tpV1.m_nUsage++;
										pTool2 = m_pBand->m_pBar->m_theToolStack.Pop();
									}
									if (!m_pBand->m_pBar->m_bWhatsThisHelp)
										FireDelayedClick(pTool);
									ReleaseCapture();
									if (IsWindow()) 
									{
										ReparentEditAndOrCombo();
										DestroyWindow();
									}
								}
								else
								{
									pTool->AddRef();
									::PostMessage(m_pBand->m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)pTool);
								}
								break;
							}
						}
					}
				}
				else if (NULL == m_pBand->m_pChildPopupWin)
				{
					KillTimer(m_hWnd, eOpenSubTimer);
					OpenSubPopup();
				}
				else if (ddTTButtonDropDown == pTool->tpV1.m_ttTools)
				{
					if (bMouseUp)
					{
						FireDelayedClick(pTool);
						ReleaseCapture();
						if (IsWindow()) 
							PostMessage(WM_CLOSE);
					}
					else
					{
						pTool->m_bPressed = TRUE;
						InvalidateRect(NULL, FALSE);
					}
				}
			}
		}
	}
	return FALSE;
}

//
// OpenSubPopup
//

void CPopupWin::OpenSubPopup()
{
	m_bTimerStarted = FALSE;
	if (m_pBand->m_pChildPopupWin)
	{
		if (m_pBand->m_pBar && m_pBand->m_pBar->m_pDesigner)
			::PostMessage(m_pBand->m_pBar->m_hWnd, GetGlobals().WM_KILLWINDOW, (WPARAM)m_pBand->m_pChildPopupWin->hWnd(), 0);
		else
		{
			m_pBand->m_pChildPopupWin->ReparentEditAndOrCombo();
			m_pBand->m_pChildPopupWin->DestroyWindow();
		}
		m_pBand->m_pChildPopupWin = NULL;
		m_nCurPopup = -1;
	}

	if (m_nCurSel >= 0)
	{
		CTool* pTool = m_pBand->m_pTools->GetVisibleTool(m_nCurSel);
		if (NULL == pTool)
			return;

		CRect rcTool = m_pBand->m_prcTools[m_nCurSel];
		POINT pt = {0, 0};
		ClientToScreen(pt);
		rcTool.Offset(pt.x, pt.y);

		if (m_bScrollActive)
			rcTool.Offset(0, -m_nPhysicalScrollPos);

		if (pTool->HasSubBand() && VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
		{
			CBand* pSubBand = m_pBand->m_pBar->FindSubBand(pTool->m_bstrSubBand);
			if (NULL == pSubBand)
				return;

			if (!pSubBand->CheckPopupOpen())
				return;
			
			//
			// Cycle detection. This should be improved to detect multilevel cycles
			//

			if (pSubBand == m_pBand)
				return;

			FlipDirection eFlipDirection = (FlipDirection)(m_pBand->m_pPopupWin || ddDALeft == m_pBand->bpV1.m_daDockingArea || ddDARight == m_pBand->bpV1.m_daDockingArea || ddDAFloat == m_pBand->bpV1.m_daDockingArea);

			m_pBand->m_pChildPopupWin = new CPopupWin(pSubBand, FALSE);
			assert(m_pBand->m_pChildPopupWin);
			if (NULL == m_pBand->m_pChildPopupWin)
				return;

			pSubBand->m_pParentPopupWin = this;

			int nDirection = ddPopupMenuLeftAlign;
			if (CTool::ddTTMoreTools == pTool->tpV1.m_ttTools)
			{
				nDirection = ddPopupMenuRightAlign;
				if (m_pBand->IsVertical())
					rcTool.Offset(-rcTool.Width() + CBand::eBevelBorder2, 0);
				else
					rcTool.Offset(rcTool.Width() + CBand::eBevelBorder2, 0);
			}

			if (!m_pBand->m_pChildPopupWin->PreCreateWin(pTool,
													     rcTool,
													     eFlipDirection,
													     ddPMDisabled != m_pBand->m_pBar->bpV1.m_pmMenus,
													     ddBFPopupFlipUp & m_pBand->bpV1.m_dwFlags,
													     nDirection))
			{
				delete m_pBand->m_pChildPopupWin;
				return;
			}

			if (m_pBand->m_pChildPopupWin->CreateWin(m_pBand->m_pBar->GetDockWindow()))
			{
				if (VARIANT_TRUE == m_pBand->m_pBar->bpV1.m_vbXPLook)
				{
					m_pBand->m_pChildPopupWin->Shadow().SetWindowPos(NULL, 
																	 0, 
																	 0, 
																	 0, 
																	 0,
																	 SWP_SHOWWINDOW|SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
				}
				m_pBand->m_pChildPopupWin->SetWindowPos(NULL, 
														0, 
														0, 
														0, 
														0,
														SWP_SHOWWINDOW|SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
				m_pBand->m_pBar->m_theToolStack.Push(pTool);
				m_nCurPopup = m_nCurSel;
				m_pBand->m_pChildPopupWin->m_nParentIndex = m_nCurSel;
			}
			else
			{
				delete m_pBand->m_pChildPopupWin;
				m_pBand->m_pChildPopupWin = NULL;
			}
		}
	}
}

//
// DetachTrack
//

BOOL CPopupWin::DetachTrack(const CRect& rcBound)
{
	POINT pt;
	BOOL  bDetach = FALSE;
	HWND  hWndPrevCapture = GetCapture();
	MSG   msg;

	SetCapture(m_hWnd);
	while (GetCapture() == m_hWnd)
	{
		GetMessage(&msg, NULL, 0, 0);
		switch (msg.message)
		{
			case WM_KEYDOWN:
			case WM_KEYUP:
			case WM_LBUTTONDOWN:
			case WM_LBUTTONUP:
			case WM_RBUTTONUP:
			case WM_RBUTTONDOWN:
				goto ExitLoop;

			case WM_MOUSEMOVE:
				GetCursorPos(&pt);
				if (!PtInRect(&rcBound, pt))
				{
					bDetach = TRUE;
					goto ExitLoop;
				}
				break;

			default:
				DispatchMessage(&msg);
				break;
		}
	}

ExitLoop:
	ReleaseCapture();
	return bDetach;
}

void CPopupWin::PushLayout()
{
	m_rcPaintCache = m_pBand->m_rcCurrentPaint;
}

void CPopupWin::PopLayout()
{
	m_pBand->m_rcCurrentPaint = m_rcPaintCache;
}

BOOL CPopupWin::Customization()
{
	return m_pBand->m_pBar->m_bCustomization && (NULL == m_pBand || ddCBSystem != m_pBand->bpV1.m_nCreatedBy);
}

//
// RefreshTool
//

void CPopupWin::RefreshTool(CTool* pTool)
{
	int nCount = m_pBand->m_pTools->GetVisibleToolCount();
	for (int nTool = 0; nTool < nCount; ++nTool)
	{
		if (m_pBand->m_pTools->GetVisibleTool(nTool) == pTool)
		{
			HDC hDC = GetDC(m_hWnd);
			if (hDC)
			{
				CRect rcClient;
				GetClientRect(rcClient);
				Draw(hDC, rcClient);
				ReleaseDC(m_hWnd, hDC);
			}
			break;
		}
	}
}

//
// FireDelayedClick
//

void CPopupWin::FireDelayedClick(CTool* pTool)
{
	if (NULL == pTool)
		return;

	CBand* pBand = m_pBand;
	assert (pBand);
	CBar* pBar = pBand->m_pBar;
	assert (pBar);

	pTool->AddRef();

	if (pBar->m_pPopupRoot)
		pBar->m_pPopupRoot->SetPopupIndex(eRemoveSelection);

	if (::IsWindow(pBar->m_hWnd))
	{
		::PostMessage(pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)pTool);
		if (pBar->m_bIsVisualFoxPro)
		{
			pTool->AddRef();
			::PostMessage(pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)pTool);
		}
	}
	else
	{
		pBar->FireToolClick(reinterpret_cast<Tool*>(pTool));
		pTool->Release();
	}
}

//
// GetLineHeight
//

int CPopupWin::GetLineHeight()
{
	int nToolCount = m_pBand->m_pTools->GetVisibleToolCount();
	
	if (nToolCount < 2 || NULL == m_pBand->m_prcTools)
		return eDefaultLineHeight;

	int nFirstVisible = -1;
	int nNextVisible = -1;
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		if (VARIANT_TRUE == m_pBand->m_pTools->GetVisibleTool(nTool)->tpV1.m_vbVisible)
		{
			if (-1 == nFirstVisible)
				nFirstVisible = nTool;
			else if (-1 == nNextVisible)
			{
				nNextVisible = nTool;
				break;
			}
		}
	}
	if (-1 == nFirstVisible || -1 == nNextVisible)
		return eDefaultLineHeight;

	return m_pBand->m_prcTools[nNextVisible].top - m_pBand->m_prcTools[nFirstVisible].top;
}

//
// TrackMenuExpandButton
//

BOOL CPopupWin::TrackMenuExpandButton(CTool* pTool)
{
	BOOL bStateChanged = FALSE;

	HDC hDC = GetDC(m_hWnd);
	HWND hWndPrev = NULL;
	CRect rcTool;
	if (m_pBand->GetToolScreenRect(m_nCurSel, rcTool))
	{
		CRect rcClientTool = rcTool;
		ScreenToClient(rcClientTool);
		pTool->m_bDropDownPressed = TRUE;
		pTool->Draw(hDC, rcClientTool, ddBTPopup, FALSE, FALSE);

		MSG msg;
		hWndPrev = SetCapture(m_hWnd);
  		while (GetCapture() == m_hWnd)
		{
			GetMessage(&msg, NULL, 0, 0);

			switch (msg.message)
			{
			case WM_CANCELMODE:
				{
					if (hWndPrev)
						SetCapture(hWndPrev);
					else
						ReleaseCapture();
				}
				goto ExitLoop;

			case WM_LBUTTONUP:
				if (hWndPrev)
					SetCapture(hWndPrev);

				if (PtInRect(&rcTool, msg.pt))
				{
					m_pBand->m_pBar->m_bPopupMenuExpanded = TRUE;
					pTool->tpV1.m_vbVisible = VARIANT_FALSE;
					m_nCurSel = -1;
					m_pBand->m_pPopupToolSelected = NULL;
					CRect rcWindow;
					GetWindowRect(rcWindow);
					if (CalcFlipingAndScrolling(rcWindow))
					{
						SetWindowPos(NULL, 
									 rcWindow.left, 
									 rcWindow.top, 
									 rcWindow.Width(), 
									 rcWindow.Height(), 
									 SWP_NOACTIVATE);
						InvalidateRect(NULL, FALSE);
					}
				}
				else
				{
					pTool->m_bDropDownPressed = FALSE;
					pTool->m_bPressed = FALSE;
					if (m_pBand->m_pBar->HasTexture())
					{
						SetBrushOrgEx(hDC, 0, 0, NULL);
						m_pBand->m_pBar->FillTexture(hDC, rcClientTool);
					}
					else
						FillSolidRect(hDC, rcClientTool, m_pBand->m_pBar->m_crBackground);
					pTool->Draw(hDC, rcClientTool, ddBTPopup, FALSE, FALSE);
				}
				goto ExitLoop;

			case WM_MOUSEMOVE:
				if (PtInRect(&rcTool, msg.pt))
				{
					if (!pTool->m_bDropDownPressed)
					{
						pTool->m_bDropDownPressed = TRUE;
						bStateChanged = TRUE;
					}
				}
				else if (pTool->m_bDropDownPressed)
				{
					pTool->m_bDropDownPressed = FALSE;
					bStateChanged = TRUE;
				}
				if (bStateChanged)
				{
					pTool->Draw(hDC, rcClientTool, ddBTPopup, FALSE, FALSE);
					bStateChanged = FALSE;
				}
				break;

			default:
				DispatchMessage(&msg);
				break;
			}
		}
	}

ExitLoop:
	pTool->m_bDropDownPressed = FALSE;
	if (GetCapture() == m_hWnd)
		ReleaseCapture();
	ReleaseDC(m_hWnd, hDC);
	return TRUE;
}

//
// CalcFlipingAndScrolling
//

BOOL CPopupWin::CalcFlipingAndScrolling(CRect& rcPopup)
{
	CRect rcBase = rcPopup;
	GetOptimalRect(rcPopup);
	rcPopup.Offset(-rcPopup.left, -rcPopup.top);
	m_rcClientOrginal = rcPopup;
	m_rcClientOrginal.Inflate(1, 1);

	// Now check if this popup fits on the screen
	CRect rcWin;
	rcWin.right = GetSystemMetrics(SM_CXVIRTUALSCREEN);
	rcWin.bottom = GetSystemMetrics(SM_CYVIRTUALSCREEN);

	CRect rcNew = rcPopup;
	int nWidth = rcPopup.Width();
	int nHeight = rcPopup.Height();

	rcNew.Offset(rcBase.left, rcBase.top);

	m_nScrollPos = 0;
	m_bScrollActive = FALSE;
	m_bFlipVert = FALSE;
	m_bFlipHorz = FALSE;

	if (rcNew.bottom > rcWin.bottom)
	{
		m_bFlipVert = TRUE;
		if (rcNew.right > rcWin.right)
			m_bFlipHorz = TRUE;
	}
	else if (rcNew.right > rcWin.right) 
	{
		// flip horizontal
		m_bFlipHorz = TRUE;
	}
	
	if (m_bFlipHorz)
	{
		if (rcBase.left - nWidth < 0)
		{
			rcNew.left = 0;
			rcNew.right = nWidth;
		}
		else
			rcNew.left = rcBase.left - nWidth;
	}

	if (m_bFlipVert)
	{
		int nLineHeight = GetLineHeight();
		// Check for scrolling req.
		if (rcBase.top - nHeight < 0)
		{
			// Doesn't fit either way so choose the larger one
			if (rcBase.top - rcWin.top > rcWin.bottom - rcBase.top) 
			{
				//
				// Flipping upwards
				//

				rcNew.top = rcWin.top;

				int nDiff = nHeight % nLineHeight;

				nHeight = rcBase.bottom - rcWin.top;

				nHeight -= (nHeight % nLineHeight) + nDiff + nLineHeight;
				
				rcNew.top = rcBase.bottom - nHeight;
			}
			else
			{
				//
				// Alright needs scrolling
				// 
				rcNew.top = rcBase.top;

				int nDiff = nHeight % nLineHeight;

				//
				// Alright needs scrolling
				//

				nHeight = rcWin.bottom - rcBase.top; 

				//
				// Correction (integral height)
				//
				
				nHeight -= (nHeight % nLineHeight) + nLineHeight - 4; 
			}
			m_bScrollActive = TRUE;
		}
		else 
		{
			// 
			// Fits if flipped upwards
			// 

			rcNew.top = rcBase.bottom - nHeight;
		}
	}
	
	rcNew.right = rcNew.left + nWidth;
	rcNew.bottom = rcNew.top + nHeight;

	// Done with flipping
	m_rcWindow = rcPopup = rcNew;
	if (m_bScrollActive)
		SetTimer(m_hWnd, ePopupScrollTimer, m_pBand->bpV1.m_nScrollingSpeed, NULL);
	return TRUE;
}

void CPopupWin::ReparentEditAndOrCombo()
{
	ActiveCombobox* pActiveCombo = m_pBand->m_pBar->GetActiveCombo();
	if (pActiveCombo && pActiveCombo->IsWindow())
	{
		HWND hWndParent = ::GetParent(pActiveCombo->hWnd());
		if (hWndParent == m_hWnd)
		{
			pActiveCombo->ShowWindow(SW_HIDE);
			::SetParent(pActiveCombo->hWnd(), m_pBand->m_pBar->m_hWnd);
		}
	}
	
	ActiveEdit*	pActiveEdit = m_pBand->m_pBar->GetActiveEdit();
	if (pActiveEdit && pActiveEdit->IsWindow())
	{
		HWND hWndParent = ::GetParent(pActiveEdit->hWnd());
		if (hWndParent == m_hWnd)
		{
			pActiveEdit->ShowWindow(SW_HIDE);
			::SetParent(pActiveEdit->hWnd(), m_pBand->m_pBar->m_hWnd);
		}
	}
}
