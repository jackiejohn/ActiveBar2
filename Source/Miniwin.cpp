//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//

#include "precomp.h"
#include <olectl.h>
#include "IpServer.h"
#include "Support.h"
#include "Globals.h"
#include "Debug.h"
#include "Designer\DesignerInterfaces.h"
#include "Designer\DragDrop.h"
#include "ChildBands.h"
#include "Band.h"
#include "Dock.h"
#include "Bar.h"
#include "MiniWin.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define DC_GRADIENT 0x0020

//
// HandleFloatingSysCommand
//

static BOOL HandleFloatingSysCommand(HWND hWnd, UINT nID, LPARAM lParam)
{
	HWND hWndParent = GetTopLevelParent(hWnd);
	switch (nID & 0xfff0)
	{
	case SC_PREVWINDOW:
	case SC_NEXTWINDOW:
		if (VK_F6 == LOWORD(lParam) && hWndParent)
		{
			SetFocus(hWndParent);
			return TRUE;
		}
		break;

	case SC_CLOSE:
	case SC_KEYMENU:

		// 
		// Check lParam.  If it is 0L, then the user may have done
		// an Alt+Tab, so just ignore it.  This breaks the ability to
		// just press the Alt-key and have the first menu selected,
		// but this is minor compared to what happens in the Alt+Tab case.
		//

		if ((nID & 0xfff0) == SC_CLOSE || lParam != 0L)
		{
			if (hWndParent)
			{
				//
				// Sending the above WM_SYSCOMMAND may destroy the app,
				// so we have to be careful about restoring activation
				// and focus after sending it.
				//

				HWND hWndSave = hWnd;
				HWND hWndFocus = ::GetFocus();
				SetActiveWindow(hWndParent);
				SendMessage(hWndParent, WM_SYSCOMMAND, nID, lParam);

				if (::IsWindow(hWndSave))
					::SetActiveWindow(hWndSave);

				if (::IsWindow(hWndFocus))
					::SetFocus(hWndFocus);
			}
		}
		return TRUE;
	}
	return FALSE;
}

//
// DrawDragRect2
//

static BOOL DrawDragRect2(HDC          hDC,
						  const CRect& rc,
						  int          nThickness,
						  CRect*       prcLast,
						  int          nLastThickness)
{
	HBRUSH hBrushOld = NULL;
	HBRUSH hBrush = NULL;
	CRect  rcInside;
	HRGN   hRgnOutside = NULL;
	HRGN   hRgnInside = NULL;
	HRGN   hRgnLast = NULL;
	HRGN   hRgnNew = NULL;
	BOOL   bResult;
	BOOL   bReturn = TRUE;
	int    nResult;

	hBrush = CreateHalftoneBrush(TRUE);
	if (NULL == hBrush)
		return FALSE;
	
	hRgnOutside = CreateRectRgnIndirect(&rc);
	if (NULL == hRgnOutside)
	{
		bReturn = FALSE;
		goto Cleanup;
	}

	rcInside = rc;
	rcInside.Inflate(-nThickness, -nThickness);
	rcInside.Intersect(rc, rcInside);

	hRgnInside = CreateRectRgnIndirect(&rcInside);
	if (NULL == hRgnInside)
	{
		bReturn = FALSE;
		goto Cleanup;
	}

	hRgnNew = CreateRectRgn(0, 0, 0, 0);
	if (NULL == hRgnNew)
	{
		bReturn = FALSE;
		goto Cleanup;
	}

	nResult = CombineRgn(hRgnNew, hRgnOutside, hRgnInside, RGN_XOR);				  
	assert(ERROR != nResult);
	if (ERROR == nResult)
	{
		bReturn = FALSE;
		goto Cleanup;
	}

	if (prcLast)
	{
		//
		// Erase the old drag rectangle
		//

		hRgnLast = CreateRectRgn(0, 0, 0, 0);
		assert(hRgnLast);
		if (NULL == hRgnLast)
		{
			bReturn = FALSE;
			goto Cleanup;
		}

		rcInside = *prcLast;
		rcInside.Inflate(-nLastThickness, -nLastThickness);
		rcInside.Intersect(*prcLast, rcInside);

		if (hRgnOutside)
			SetRectRgn(hRgnOutside, prcLast->left, prcLast->top, prcLast->right, prcLast->bottom);

		if (hRgnInside)
			SetRectRgn(hRgnInside, rcInside.left, rcInside.top, rcInside.right, rcInside.bottom);

		nResult = CombineRgn(hRgnLast, hRgnOutside, hRgnInside, RGN_XOR);				  
		assert(ERROR != nResult);
		if (ERROR == nResult)
		{
			bReturn = FALSE;
			goto Cleanup;
		}

		nResult = SelectClipRgn(hDC, hRgnLast);
		assert(ERROR != nResult);
		if (ERROR == nResult)
		{
			bReturn = FALSE;
			goto Cleanup;
		}

		nResult = GetClipBox(hDC, &rcInside);
		assert(ERROR != nResult);
		if (ERROR == nResult)
		{
			bReturn = FALSE;
			goto Cleanup;
		}
		
		hBrushOld = SelectBrush(hDC, hBrush);

		bResult = PatBlt(hDC, rcInside.left, rcInside.top, rcInside.Width(), rcInside.Height(), PATINVERT);
		assert(bResult);
		
		SelectBrush(hDC, hBrushOld);
		
		hBrushOld = NULL;
	}

	// Draw into the new region
	nResult = SelectClipRgn(hDC, hRgnNew);
	assert(ERROR != nResult);
	if (ERROR == nResult)
	{
		bReturn = FALSE;
		goto Cleanup;
	}

	nResult = GetClipBox(hDC, &rcInside);
	assert(ERROR != nResult);
	if (ERROR == nResult)
	{
		bReturn = FALSE;
		goto Cleanup;
	}

	hBrushOld = SelectBrush(hDC, hBrush);

	bResult = PatBlt(hDC, rcInside.left, rcInside.top, rcInside.Width(), rcInside.Height(), PATINVERT);
	assert(bResult);
	
	SelectBrush(hDC, hBrushOld);
	
	// Cleanup DC
	nResult = SelectClipRgn(hDC, NULL);
	assert(ERROR != nResult);

Cleanup:
	if (hBrush)
	{
		bResult = DeleteBrush(hBrush);
		assert(bResult);
	}

	if (hRgnOutside)
	{
		bResult = DeleteRgn(hRgnOutside);
		assert(bResult);
	}

	if (hRgnInside)
	{ 
		bResult = DeleteRgn(hRgnInside);
		assert(bResult);
	}

	if (hRgnNew)
	{
		bResult = DeleteRgn(hRgnNew);
		assert(bResult);
	}

	if (hRgnLast)
	{
		bResult = DeleteRgn(hRgnLast);
		assert(bResult);
	}
	return bReturn;
}

//
// CMiniDragDrop
//

CBand* CMiniDragDrop::GetBand(POINT pt)
{
	return ((CMiniWin*)m_pWnd)->m_pBand;
}

DWORD CMiniWin::m_dwStyle = WS_CAPTION|WS_THICKFRAME|WS_POPUP|WS_SYSMENU|WS_CLIPCHILDREN;
DWORD CMiniWin::m_dwStyleEx = WS_EX_TOOLWINDOW;

//
// CMiniWin
//

CMiniWin::CMiniWin(CBand* pBand)
	: m_pBand(pBand),
	  m_pBar(NULL),
	  m_pDockMgr(NULL),
	  m_theMiniDropTarget(CToolDataObject::eActiveBarDragDropId)
{
	m_bBandClosing = FALSE;
	m_bActive = TRUE;
	assert(m_pBand);
	if (NULL == m_pBand)
		return;

	assert(m_pBand->m_pBar);
	if (NULL == m_pBand->m_pBar)
		return;

	m_pBar = m_pBand->m_pBar;

	assert(m_pBar->m_pDockMgr);
	m_pDockMgr = m_pBar->m_pDockMgr;

	if (!g_fSysWin95Shell)
	{
		m_dwStyle = WS_THICKFRAME|WS_POPUP|WS_SYSMENU;
		m_dwStyleEx = 0;
	}
	BOOL bResult = RegisterWindow(DD_WNDCLASS("ABFloatWin"),
								  CS_DBLCLKS|CS_SAVEBITS|CS_VREDRAW|CS_HREDRAW,
								  NULL,
								  ::LoadCursor(NULL, IDC_ARROW));
	assert(bResult);
	m_theBandDragDrop.m_pWnd = this;
	m_theMiniDropTarget.Init (&m_theBandDragDrop, m_pBar, CToolDataObject::eActiveBarDragDropId);
	m_bHover = FALSE;
}

CMiniWin::~CMiniWin()
{
	TRACE(1, _T("~CMiniWin\n"));
}
	
//
// Show
//

void CMiniWin::Show(BOOL bActive)
{
	SetWindowPos(NULL, 
				 0, 
				 0, 
				 0, 
				 0,
				 (bActive ? SWP_SHOWWINDOW:SWP_HIDEWINDOW)|SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOSIZE|SWP_NOMOVE);
}

//
// InCustomization
//

BOOL CMiniWin::InCustomization()
{
	assert (m_pBar);
	return m_pBar->m_bCustomization;
}

//
// Resize
//

BOOL CMiniWin::FreeResize()
{
	assert(m_pBand);
	return (BOOL)((ddCBSNone != m_pBand->bpV1.m_cbsChildStyle) || (ddBFSizer & m_pBand->bpV1.m_dwFlags));
}

//
// CreateWin
//

BOOL CMiniWin::CreateWin(HWND hWndParent)
{
	m_pBand->bpV1.m_vbVisible = VARIANT_TRUE;
	DockingAreaTypes daPrevDock = m_pBand->bpV1.m_daDockingArea;

	// Prevent programmer from screwing the system
	if (ddDAFloat == daPrevDock)
		m_pBand->bpV1.m_daDockingArea = daPrevDock; 

	if (VARIANT_FALSE == m_pBand->bpV1.m_vbVisible)
		return FALSE;

	MAKE_TCHARPTR_FROMWIDE(szTitle, m_pBand->m_bstrCaption);
	if (ddBTChildMenuBar == m_pBand->m_pRootBand->bpV1.m_btBands || ddBTMenuBar == m_pBand->m_pRootBand->bpV1.m_btBands || (!(ddBFHide & m_pBand->m_pRootBand->bpV1.m_dwFlags) && !(ddBFClose & m_pBand->m_pRootBand->bpV1.m_dwFlags)))
		m_dwStyle &= ~WS_SYSMENU;
	else
		m_dwStyle |= WS_SYSMENU;

	SIZE size = m_pBand->bpV1.m_rcFloat.Size();
	if (0 == size.cx || 0 == size.cy)
	{
		CRect rcTemp = m_pBand->GetOptimalFloatRect(FALSE);
		m_pBand->bpV1.m_rcFloat.right = m_pBand->bpV1.m_rcFloat.left + rcTemp.Width();
		m_pBand->bpV1.m_rcFloat.bottom = m_pBand->bpV1.m_rcFloat.top + rcTemp.Height();
		size = m_pBand->bpV1.m_rcFloat.Size();
	}

	CreateEx(m_dwStyleEx,
			 szTitle,
			 m_dwStyle,
			 m_pBand->bpV1.m_rcFloat.left,
			 m_pBand->bpV1.m_rcFloat.top,
			 m_pBand->bpV1.m_rcFloat.Width(),
			 m_pBand->bpV1.m_rcFloat.Height(),
			 hWndParent);

	if (IsWindow())
	{
		m_pBand->ParentWindowedTools(m_hWnd);

		if (m_pBar->m_pDragDropManager && S_OK != m_pBar->m_pDragDropManager->RegisterDragDrop((OLE_HANDLE)m_hWnd, &m_theMiniDropTarget))
		{
			assert(FALSE);
			TRACE(1, "Failed to register drag and drop target\n");
		}
		else if (S_OK != GetGlobals().m_pDragDropMgr->RegisterDragDrop((OLE_HANDLE)m_hWnd, &m_theMiniDropTarget))
		{
			assert(FALSE);
			TRACE(1, "Failed to register drag and drop target\n");
		}
//		UpdateWindow();
	}

	return (NULL == m_hWnd ? FALSE : TRUE);
}

//
// WindowProc
//

LRESULT CMiniWin::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	if (GetGlobals().WM_TOOLTIPSUPPORT == nMsg)
	{
		*reinterpret_cast<BOOL*>(lParam) = TRUE;
		return 0;
	}

	if (GetGlobals().WM_KILLWINDOW == nMsg)
	{
		SendMessage(WM_CLOSE);
		return 0;
	}

	if (GetGlobals().WM_FLOATSTATUS == nMsg)
	{
		LRESULT lResult = 0;
		if (wParam & FS_SYNCACTIVE)
		{
			lResult = 1;
		}
		if (wParam & (FS_SHOW|FS_HIDE))
		{
			SetWindowPos(NULL, 0, 0, 0, 0, ((wParam & FS_SHOW) ? SWP_SHOWWINDOW : SWP_HIDEWINDOW) | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
		}
		if (wParam & (FS_ENABLE|FS_DISABLE))
		{
			EnableWindow(m_hWnd, (wParam & FS_ENABLE) != 0);
		}
		if (wParam & (FS_ACTIVATE|FS_DEACTIVATE))
		{
			SendMessage(WM_NCACTIVATE, (wParam & FS_ACTIVATE) != 0);
		}
		return lResult;
	}

	BOOL bForeground = NULL != GetActiveWindow();

	switch (nMsg)
	{
	case WM_NCCREATE:
		{
			if (CBar::eClientArea != m_pBar->m_eAppType)
			{
				LRESULT lResult =  FWnd::WindowProc(nMsg, wParam, lParam);
				if (!lResult)
					return lResult;

				HWND hParentWnd = ::GetWindow(m_hWnd, GW_OWNER);
				assert(hParentWnd != NULL);
				HWND hActiveWnd = GetForegroundWindow();
				BOOL bActive = (hParentWnd == hActiveWnd) || (::SendMessage(hActiveWnd, GetGlobals().WM_FLOATSTATUS, FS_SYNCACTIVE, 0) != 0);
				SendMessage(GetGlobals().WM_FLOATSTATUS, bActive ? FS_ACTIVATE : FS_DEACTIVATE);
				return TRUE;
			}
		}
		break;

	case WM_NCACTIVATE:
		{
			if (CBar::eSDIForm == m_pBar->m_eAppType)
			{
				static BOOL bGuard = FALSE;
				if (bGuard)
					return FALSE;
				bGuard = TRUE;

				HWND hWnd = GetFocus();
				if (NULL != hWnd && m_pBar->m_hWnd != hWnd && IsDescendant(m_hWnd, hWnd))
					::SetFocus(m_pBar->m_hWnd);
				
				if (m_bActive != (BOOL)wParam)
				{
					m_bActive = (BOOL)wParam;
					SendMessage(WM_NCPAINT);
				}
				bGuard = FALSE;
			}
			else if (m_bActive != (BOOL)wParam)
			{
				m_bActive = (BOOL)wParam;
				SendMessage(WM_NCPAINT);
			}
		}
		return TRUE;

	case WM_MOUSEACTIVATE:
		{
			if (m_pBar && m_pBar->AmbientUserMode())
			{
				switch (m_pBar->m_eAppType)
				{
				case CBar::eMDIForm:
					{
						HWND hWndActive = GetForegroundWindow();
						BOOL bResult;
						if (NULL == hWndActive || (m_hWnd != hWndActive && !IsDescendant(hWndActive, m_hWnd)))
							bResult = SetForegroundWindow(m_pBand->m_pBar->GetDockWindow());
						return MA_NOACTIVATE;
					}
					break;

				case CBar::eClientArea:
				case CBar::eSDIForm:
					{
						LRESULT lResult = FWnd::WindowProc(nMsg, wParam, lParam);
						POINT pt;
						GetCursorPos(&pt);
						HWND hWndPoint = ::WindowFromPoint(pt);
						if (hWndPoint == GetFocus())
							return lResult;
						HWND hWnd = m_pBar->m_theFormsAndControls.FindChildFromPoint(pt);
						if (::IsWindow(hWnd))
							::SetFocus(m_pBar->m_hWnd);
						return MA_NOACTIVATE;
					}
					break;
				}
			}
		}
		break;

	case WM_SETFOCUS:
		TRACE(6, "WM_SETFOCUS\n")
		break;

	case WM_KILLFOCUS:
		TRACE(6, "WM_KILLFOCUS\n")
		break;

	case WM_GETMINMAXINFO:
		{
			// Allow Windows to fill in the defaults
			FWnd::WindowProc(nMsg, wParam, lParam);
			
			MINMAXINFO* pMMI = (MINMAXINFO*)lParam;
			if (pMMI)
			{
				// Don't allow sizing smaller than the non-client area
				CRect rcWindow;
				GetWindowRect(rcWindow);
				
				CRect rcClient;
				GetClientRect(rcClient);

				pMMI->ptMinTrackSize.x = rcWindow.Width() - rcClient.right;
				pMMI->ptMinTrackSize.y = rcWindow.Height() - rcClient.bottom;
			}
		}
		return 0;

	case WM_DESTROY:
		try
		{
			m_pBand->ParentWindowedTools(NULL);

			LRESULT lResult =  FWnd::WindowProc(nMsg,wParam,lParam);

			if (m_pBar->m_pDragDropManager)
				m_pBar->m_pDragDropManager->RevokeDragDrop((OLE_HANDLE)m_hWnd);
			else
				GetGlobals().m_pDragDropMgr->RevokeDragDrop((OLE_HANDLE)m_hWnd);

			GetWindowRect(m_pBand->bpV1.m_rcFloat);
		
			if (!m_bBandClosing)
			{
				m_pBand->m_bMiniWinShowLock = FALSE;
				m_bBandClosing = TRUE;
				if (ddDAFloat == m_pBand->bpV1.m_daDockingArea && NULL == m_pBand->m_pBar->m_pDesigner) 
					m_pBand->put_Visible(VARIANT_FALSE);
			}
			m_pBand->m_pFloat = NULL;
			return lResult;
		}
		CATCH
		{
			m_pBand->m_pFloat = NULL;
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		break;

	case WM_NCDESTROY:
		{
			LRESULT lResult =  FWnd::WindowProc(nMsg,wParam,lParam);
			delete this;
			return lResult;
		}
		break;

	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(m_hWnd,&ps);
			if (hDC)
			{
				CRect rcBound;
				GetClientRect(rcBound);
				CFlickerFree& ffObject = m_pBar->m_ffObj;

				HDC hDCOff = ffObject.RequestDC(hDC, rcBound.Width(), rcBound.Height());
				if (NULL == hDCOff)
					hDCOff = hDC;
				else
					rcBound.Offset(-rcBound.left, -rcBound.top);
				
				if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				{
					if (m_pBand->IsMenuBand())
						FillSolidRect(hDCOff, rcBound, m_pBar->m_crXPBackground);
					else
						FillSolidRect(hDCOff, rcBound, m_pBar->m_crXPBandBackground);
				}
				else if (m_pBar->HasTexture() && ddBOFloat & m_pBar->bpV1.m_dwBackgroundOptions)
				{
					SetBrushOrgEx(hDCOff, 0, 0, NULL);
					m_pBand->m_pBar->FillTexture(hDCOff, rcBound);
				}
				else
					FillSolidRect(hDCOff, rcBound, m_pBar->m_crBackground);

				POINT pt = {0,0};
				m_pBand->Draw(hDCOff, rcBound, m_pBand->IsWrappable(), pt);

				if (ddCBSSlidingTabs == m_pBand->bpV1.m_cbsChildStyle)
					m_pBand->DrawAnimatedSlidingTabs(ffObject, hDC, rcBound);

				if (hDCOff != hDC)
					ffObject.Paint(hDC, 0, 0);

				EndPaint(m_hWnd, &ps);
				return 0;
			}
		}
		break;

	case WM_NCMOUSEMOVE:
		{
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				POINTS pts = MAKEPOINTS(lParam);   
				POINT  pt = {pts.x, pts.y};
				CRect rc;
				GetWindowRect(rc);
				CRect rcClose = m_rcClose;
				rcClose.Offset(rc.left, rc.top);
				CRect rcTemp = m_rcClose;
				if (PtInRect(&rcClose, pt))
				{
					if (!m_bHover)
					{
						m_bHover = TRUE;
						HDC hDC = GetWindowDC(m_hWnd);
						if (hDC)
						{
							rcTemp.Inflate(3, 3);
							FillSolidRect(hDC, rcTemp, m_pBar->m_crXPSelectedColor);

							DrawCross(hDC, 
									  m_rcClose.left, 
									  m_rcClose.top, 
									  m_rcClose.right, 
									  m_rcClose.bottom, 
									  m_nPenThickness,
									  GetSysColor(COLOR_BTNTEXT));
							ReleaseDC(m_hWnd, hDC);
						}
					}
				}
				else if (m_bHover)
				{
					m_bHover = FALSE;
					HDC hDC = GetWindowDC(m_hWnd);
					if (hDC)
					{
						rcTemp.Inflate(4, 4);
						FillSolidRect(hDC, rcTemp, m_pBar->m_crXPFloatCaptionBackground);

						DrawCross(hDC, 
								  m_rcClose.left, 
								  m_rcClose.top, 
								  m_rcClose.right, 
								  m_rcClose.bottom, 
								  m_nPenThickness,
								  GetSysColor(COLOR_BTNTEXT));
						ReleaseDC(m_hWnd, hDC);
					}
				}
			}
		}
		return 0;

	case WM_NCLBUTTONDOWN:
		{
			// Position of cursor 
			POINTS pts = MAKEPOINTS(lParam);   
			POINT  pt;
			pt.x = pts.x;
			pt.y = pts.y;
			
			if (!m_pBand->m_pBar->IsForeground())
				ActivateTopParent();

			if (OnNcLButtonDown((INT)wParam, pt))
				return 0;
		}
		break;
	
	case WM_NCLBUTTONDBLCLK:
		{	
			if (!bForeground)
				break;
			
			//
			// We are going to dock the window on the previous dock area
			//

			if (m_pBand)
			{
				switch (m_pBand->m_daPrevDockingArea)
				{
				case ddDATop:
				case ddDABottom:
				case ddDALeft:
				case ddDARight:
					{
						DWORD dwDockFlags[] = {ddBFDockTop, ddBFDockBottom, ddBFDockLeft, ddBFDockRight};
						if (!(dwDockFlags[m_pBand->m_daPrevDockingArea] & m_pBand->bpV1.m_dwFlags))
							return 0;
						m_pBand->ParentWindowedTools(m_pBand->m_pDockMgr->hWnd(m_pBand->m_daPrevDockingArea));
					}
					break;

				default:
					m_pBand->HideWindowedTools();
					break;
				}

				m_pBand->bpV1.m_daDockingArea = m_pBand->m_daPrevDockingArea;
				m_pBand->bpV1.m_nDockOffset = m_pBand->m_nPrevDockOffset;
				m_pBand->bpV1.m_nDockLine = m_pBand->m_nPrevDockLine;

				//
				// Put on stack since Recalclayout will Destroy this guy
				//

				try
				{
					CBand* pBand = m_pBand;
					pBand->AddRef();
					m_pBar->FireBandMove((Band*)pBand);
					m_pBar->FireBandDock((Band*)pBand);
					pBand->Release();
					m_pBar->RecalcLayout();
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
		}
		return 0;

	case WM_LBUTTONDOWN:
		{
			//
			// Make it the top most window
			//

			SetWindowPos(HWND_TOP,
						 0,
						 0,
						 0,
						 0,
						 SWP_NOSIZE|SWP_NOMOVE|SWP_NOACTIVATE);
			POINT pt;
			GetCursorPos(&pt);
			ScreenToClient(pt);
			m_pBand->OnLButtonDown(wParam, pt);
		}
		break;

	case WM_LBUTTONDBLCLK:
		{
			if (m_pBar->m_bMenuLoop)
				return TRUE;

			POINT pt = {LOWORD(lParam), HIWORD(lParam)};
			m_pBand->OnLButtonDblClk(wParam, pt);
		}
		return TRUE;

	case WM_RBUTTONDOWN:
		{
			POINT pt = {LOWORD(lParam), HIWORD(lParam)};
			m_pBand->OnRButtonDown(wParam, pt);
		}
		break;
	
	case WM_RBUTTONUP:
		{
			if (!m_pBar->m_bCustomization)
			{
				BOOL bResult = m_pBar->OnBarContextMenu();
				return TRUE;
			}
			POINT pt = {LOWORD(lParam), HIWORD(lParam)};
			m_pBand->OnRButtonUp(wParam, pt);
		}
		break;

	case WM_LBUTTONUP:
		{
			POINT pt = {LOWORD(lParam), HIWORD(lParam)};
			m_pBand->OnLButtonUp(wParam, pt);
		}
		break;

	case WM_MOUSEMOVE:
		{
			HWND hWndActive = GetActiveWindow();
			HWND hWndTemp = m_pBar->GetDockWindow();
			if (hWndTemp != hWndActive)
			{
				while (hWndTemp)
				{
					hWndTemp = GetParent(hWndTemp);
					if (hWndTemp == hWndActive)
						break;
				}
				if (NULL == hWndTemp)
					break;
			}

			if (!bForeground)
				break;

			POINT pt = {LOWORD(lParam), HIWORD(lParam)};
			m_pBand->OnMouseMove(wParam, pt);
		}
		return 0;
	
	case WM_SHOWWINDOW:
		switch (lParam) 
		{
		case SW_PARENTCLOSING:
			m_pBand->m_bMiniWinShowLock = TRUE;
			break;
		
		case SW_PARENTOPENING:
			m_pBand->m_bMiniWinShowLock = FALSE;
			break;
		}
		break;

	case WM_SETCURSOR:
		switch(LOWORD(lParam))
		{
		case HTLEFT:
		case HTRIGHT:
			SetCursor(LoadCursor(NULL, IDC_SIZEWE));
			break;

		case HTTOP:
		case HTBOTTOM:
			SetCursor(LoadCursor(NULL, IDC_SIZENS));
			break;

		case HTTOPLEFT:
		case HTBOTTOMRIGHT:
			SetCursor(LoadCursor(NULL, IDC_SIZENWSE));
			break;

		case HTTOPRIGHT:
		case HTBOTTOMLEFT:
			SetCursor(LoadCursor(NULL, IDC_SIZENESW));
			break;

		default:
			SetCursor(LoadCursor(NULL, IDC_ARROW));
			SetCursor(GetCursor());
			break;
		}
		return 0;

	case WM_SETTEXT:
		FWnd::WindowProc(nMsg, wParam, lParam);
		SendMessage(WM_NCPAINT, 0, 0);
		return 0;

    case WM_PARENTNOTIFY:
		switch (LOWORD(wParam))
		{
		case WM_CREATE:
		case WM_DESTROY:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			break;

		case WM_LBUTTONDOWN:
		case WM_MBUTTONDOWN:
		case WM_RBUTTONDOWN:
			{
				POINT pt;
				GetCursorPos(&pt);
				HWND hWnd = m_pBar->m_theFormsAndControls.FindChildFromPoint(pt);
				if (::IsWindow(hWnd))
					return ::SendMessage(hWnd, nMsg, wParam, lParam);
			}
			break;
		}
		break;

    case WM_NOTIFY:
    case WM_NOTIFYFORMAT:
    case WM_CTLCOLORBTN:
    case WM_CTLCOLORDLG:
    case WM_CTLCOLOREDIT:
	case WM_CTLCOLORLISTBOX:
    case WM_CTLCOLORMSGBOX:
    case WM_CTLCOLORSTATIC:
    case WM_DRAWITEM:
    case WM_MEASUREITEM:
    case WM_DELETEITEM:
    case WM_VKEYTOITEM:
    case WM_CHARTOITEM:
	case WM_COMPAREITEM:
    case WM_HSCROLL:
    case WM_VSCROLL:
    case WM_SIZE:
		{
			try
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			CATCH
			{
				m_pBand->m_pFloat = NULL;
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
		break;

	case WM_COMMAND:
		{
			try
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						::SendMessage(hWnd, nMsg, wParam, lParam);
				}
				else
				{
					switch (LOWORD(wParam))
					{
					case eFloatResize:
						{
							// external resize request
							CRect rcFloat = m_pBand->bpV1.m_rcFloat;
							if (FreeResize())
							{
								MoveWindow(rcFloat, TRUE);
								m_pBand->m_pBar->RecalcLayout();
							}
							else
							{
								CRect rcAdjust;
								AdjustWindowRectEx(rcAdjust, FALSE);

								CRect rcBound = rcFloat;
								rcBound.left   -= rcAdjust.left;
								rcBound.top    -= rcAdjust.top;
								rcBound.right  -= rcAdjust.right;
								rcBound.bottom -= rcAdjust.bottom;
								
								// Make sure layout is correct
								CRect rcResult;
								if (m_pBand->CalcLayout(rcBound, CBand::eLayoutFloat | CBand::eLayoutHorz, rcResult, TRUE))
								{
									AdjustWindowRectEx(rcResult, FALSE);
									rcFloat.right = rcFloat.left + rcResult.Width();
									rcFloat.bottom = rcFloat.top + rcResult.Height();
									MoveWindow(rcFloat, TRUE);
								}
							}
						}
						break;

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
			}
			CATCH
			{
				m_pBand->m_pFloat = NULL;
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
		break;

	case WM_SYSCOMMAND:
		{
			UINT nID = wParam;
			if ((nID & 0xFFF0) != SC_CLOSE || (GetKeyState(VK_F4) < 0 && GetKeyState(VK_MENU) < 0))
			{
				if (HandleFloatingSysCommand(m_hWnd, nID, lParam))
					return 0;
			}
			if (SC_CLOSE == nID)
				m_pBand->HideWindowedTools();
		}
		break;

	case WM_NCPAINT:
		{
			//
			// Painting the windows caption area
			//

			BOOL bResult;
			CRect rc;
			GetWindowRect(rc);
			HDC hDC = GetWindowDC(m_hWnd);
			if (hDC)
			{
				SIZE sizeBorder;
				sizeBorder.cx = GetSystemMetrics(SM_CXFRAME);
				sizeBorder.cy = GetSystemMetrics(SM_CYFRAME);
				rc.Offset(-rc.left, -rc.top);

				if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				{
					CRect rcBorder = rc;
					rcBorder.Inflate(-sizeBorder.cx + 2, -sizeBorder.cy + 2);
					HPEN penBack = CreatePen(PS_SOLID, sizeBorder.cx, m_pBar->m_crXPBandBackground);
					if (penBack)
					{
						HPEN penOld = SelectPen(hDC, penBack);

						HBRUSH brushOld = SelectBrush(hDC, (HBRUSH)GetStockObject(HOLLOW_BRUSH));

						Rectangle(hDC, rcBorder.left, rcBorder.top, rcBorder.right, rcBorder.bottom);
					
						SelectBrush(hDC, brushOld); // Don't have to delete this one
						
						SelectPen(hDC, penOld);
						
						bResult = DeletePen(penBack);
						assert(bResult);
					}

					penBack = CreatePen(PS_SOLID, sizeBorder.cx, GetSysColor(COLOR_APPWORKSPACE));
					if (penBack)
					{
						HPEN penOld = SelectPen(hDC, penBack);
					
						MoveToEx(hDC, rc.left, rc.top, NULL);
						LineTo(hDC, rc.right, rc.top);
						LineTo(hDC, rc.right, rc.bottom);
						LineTo(hDC, rc.left, rc.bottom);
						LineTo(hDC, rc.left, rc.top);

						SelectPen(hDC, penOld);
						bResult = DeletePen(penBack);
						assert(bResult);
					}
				}
				else
					DrawFrameRect(m_pBar, hDC, rc, sizeBorder.cx, sizeBorder.cy);

				rc.Inflate(-sizeBorder.cx, -sizeBorder.cy);
				
				int nCaptionHeight;
				if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				{
					nCaptionHeight = eMiniWinCaptionHeight ;
					rc.bottom = rc.top + eMiniWinCaptionHeight - 1;

					HBRUSH hBrush = CreateSolidBrush(m_pBar->m_crXPFloatCaptionBackground);
					if (hBrush)
					{
						HBRUSH brushOld = SelectBrush(hDC, hBrush);
						FillRect(hDC, &rc, hBrush);
						SelectBrush(hDC, brushOld);
						BOOL bResult = DeleteBrush(hBrush);
						assert(bResult);
					}
					
					HFONT hFontSmall = m_pBar->GetSmallFont(FALSE);
					HFONT hFontOld = SelectFont(hDC, hFontSmall);
					
					SetBkMode(hDC, TRANSPARENT);
					SetTextColor(hDC, GetSysColor(COLOR_CAPTIONTEXT));

					TCHAR szCaption[128];
					SendMessage(WM_GETTEXT, 128, (LPARAM)szCaption);
					TextOut(hDC, rc.left+1, rc.top+1, szCaption, lstrlen(szCaption));
					SelectFont(hDC, hFontOld);
				}
				else if (g_fSysWin95Shell)
				{
					nCaptionHeight = GetSystemMetrics(SM_CYSMCAPTION);
					rc.bottom = rc.top + nCaptionHeight - 1;
					UINT uFlags = DC_TEXT | DC_SMALLCAP;
					if (m_bActive)
						uFlags |= DC_ACTIVE;
					if (GetGlobals().m_nBitDepth > 8)
						uFlags |= DC_GRADIENT;
					DrawCaption(m_hWnd, hDC, &rc, uFlags);
				}
				else
				{
					//
					// Old Window so we have to draw the caption
					//

					nCaptionHeight = eMiniWinCaptionHeight ;
					rc.bottom = rc.top + eMiniWinCaptionHeight - 1;
					
					FillRect(hDC, &rc, (HBRUSH)(1+COLOR_ACTIVECAPTION));
					
					HFONT hFontSmall = m_pBar->GetSmallFont(FALSE);
					HFONT hFontOld = SelectFont(hDC, hFontSmall);
					
					SetBkMode(hDC, TRANSPARENT);
					SetTextColor(hDC, GetSysColor(COLOR_CAPTIONTEXT));

					TCHAR szCaption[128];
					SendMessage(WM_GETTEXT, 128, (LPARAM)szCaption);
					TextOut(hDC, rc.left+1, rc.top+1, szCaption, lstrlen(szCaption));
					SelectFont(hDC, hFontOld);
				}

				// Space below caption
				CRect rcTemp = rc; 
				rcTemp.top = rcTemp.bottom;
				rcTemp.bottom = rcTemp.top + 1;
				
				FillSolidRect(hDC, rcTemp, GetSysColor(COLOR_BTNFACE));

				if (ddBTChildMenuBar != m_pBand->bpV1.m_btBands && 
					ddBTMenuBar != m_pBand->bpV1.m_btBands && 
					(ddBFHide & m_pBand->bpV1.m_dwFlags || ddBFClose & m_pBand->bpV1.m_dwFlags))
				{
					//
					// Draw close box
					//

					rc.left = rc.right - nCaptionHeight - 1;
					if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
					{
						rc.Offset(1, 0);
						rc.Inflate(-4, -3);

						m_rcClose = rc;
						m_rcClose.left++;
						m_rcClose.top++;
						m_rcClose.right -= 3; 
						m_rcClose.bottom -= 3; 
						// Now draw the X
						m_nPenThickness = 1 + ((rc.Height() + 1) / 10);
						DrawCross(hDC, 
								  m_rcClose.left, 
								  m_rcClose.top, 
								  m_rcClose.right, 
								  m_rcClose.bottom, 
								  m_nPenThickness,
								  GetSysColor(COLOR_BTNTEXT));
					}
					else if (g_fSysWin95Shell)
					{
						rc.Inflate(-2, -2);
						
						// Still off a little so adjust it
						rc.left += 2;
						
						// Draw the close button
						DrawFrameControl(hDC, &rc, DFC_CAPTION, DFCS_CAPTIONCLOSE);
					}
					else
					{
						//
						// Old Window so we have to draw the close button
						//

						rc.Inflate(-2, -1);
						rc.Offset(1, 0);
						DrawEdge(hDC, &rc, EDGE_RAISED, BF_RECT);
						
						rc.Inflate(-2, -2);
						FillRect(hDC, &rc,(HBRUSH)(1+COLOR_BTNFACE));

						// Now draw the X
						int nPenThickness = 1 + ((rc.Height()+1) / 10);
						DrawCross(hDC, 
								  rc.left + 1, 
								  rc.top + 2, 
								  rc.right - 3, 
								  rc.bottom - 2, 
								  nPenThickness,
								  GetSysColor(COLOR_BTNTEXT));
					}
				}
				ReleaseDC(m_hWnd, hDC);
			}
			return 0;
		}

	case WM_NCCALCSIZE:
		if (!g_fSysWin95Shell)
		{
			//
			// This part needs to be handled if Win16 (send NCPaint message...)
			// prepare ncAdjustRect for future use
			//
			
			LPNCCALCSIZE_PARAMS pParams = (LPNCCALCSIZE_PARAMS)lParam;
			pParams->rgrc[0].top += eMiniWinCaptionHeight;
		}
		break;

	case WM_NCHITTEST:
		{
			UINT nHit = FWnd::WindowProc(nMsg,wParam,lParam);
			if (!g_fSysWin95Shell && HTNOWHERE == nHit)
			{
				RECT rc;
				GetWindowRect(rc);
				if ((LOWORD(lParam) > rc.right - 18) && HIWORD(lParam) < (rc.top + 18))
					return HTCLOSE;

				return HTCAPTION;
			}

			if ((nHit < HTSIZEFIRST || nHit > HTSIZELAST) && nHit != HTGROWBOX)
				return nHit;

			SIZE sizeFrame;
			sizeFrame.cx = GetSystemMetrics(SM_CXFRAME);
			sizeFrame.cy = GetSystemMetrics(SM_CYFRAME);

			CRect rcWindow;
			GetWindowRect(rcWindow);
			rcWindow.Inflate(-sizeFrame.cx, -sizeFrame.cy);

			if (!FreeResize())
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				switch (nHit)
				{
				case HTTOPLEFT:
					return pt.y < rcWindow.top ? HTTOP : HTLEFT;

				case HTTOPRIGHT:
					return pt.y < rcWindow.top ? HTTOP : HTRIGHT;

				case HTBOTTOMLEFT:
					return pt.y > rcWindow.bottom ? HTBOTTOM : HTLEFT;

				case HTGROWBOX:
				case HTBOTTOMRIGHT:
					return pt.y > rcWindow.bottom ? HTBOTTOM : HTRIGHT;
				}	
			}
			return nHit;
		}
	}
	return FWnd::WindowProc(nMsg,wParam,lParam);
}

//
// ActivateTopParent
//

void CMiniWin::ActivateTopParent()
{
	HWND hWndActive = GetForegroundWindow();
	if (::IsWindow(hWndActive) || !(hWndActive == m_hWnd || ::IsChild(hWndActive, m_hWnd)))
	{
		HWND hWndTopLevel = GetTopLevelParent(m_hWnd);
		SetForegroundWindow(hWndTopLevel);
	}
}

//
// OnNcLButtonDown
//

BOOL CMiniWin::OnNcLButtonDown(UINT nHitTest, POINT pt)
{
	BOOL bXPLook = (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook);
	SetWindowPos(HWND_TOP,
				 0,
				 0,
				 0,
				 0,
				 SWP_NOSIZE|SWP_NOMOVE|SWP_NOACTIVATE);

	if (!m_pBar->IsForeground())
		ActivateTopParent();

	//
	// Test for close box
	//

	if (!bXPLook && HTCLOSE == nHitTest)
		return FALSE;

	// Size Window?
	if (nHitTest >= HTSIZEFIRST && nHitTest <= HTSIZELAST) 
		m_pDockMgr->StartResize(m_pBand, nHitTest, pt);
	else
	{
		try
		{
			HWND hWndCapture;
			CRect rc;
			GetWindowRect(rc);
			CRect rcClose = m_rcClose;
			rcClose.Offset(rc.left, rc.top);
			rcClose.Inflate(3, 3);
			if (bXPLook && PtInRect(&rcClose, pt))
			{
				BOOL bOver = TRUE;
				HDC hDC = GetWindowDC(m_hWnd);
				if (hDC)
				{
					CRect rcTemp = m_rcClose;
					rcTemp.Inflate(3, 3);
					FillSolidRect(hDC, rcTemp, m_pBar->m_crXPPressed);

					DrawCross(hDC, 
							  m_rcClose.left, 
							  m_rcClose.top, 
							  m_rcClose.right, 
							  m_rcClose.bottom, 
							  m_nPenThickness,
							  GetSysColor(COLOR_BTNTEXT));

					SetCapture(m_hWnd);
					MSG msg;
					BOOL bProcessing = TRUE;
					DWORD nDragStartTick = GetTickCount();
					while (GetCapture() == m_hWnd && bProcessing)
					{
						GetMessage(&msg, NULL, 0, 0);

						switch (msg.message)
						{
						case WM_RBUTTONUP:
							bProcessing = FALSE;
							break;

						case WM_LBUTTONUP:
							if (PtInRect(&rcClose, msg.pt))
								PostMessage(WM_CLOSE);
							bProcessing = FALSE;
							break;

						case WM_MOUSEMOVE:
							{
								if (PtInRect(&rcClose, msg.pt))
								{
									if (!bOver)
									{
										bOver = TRUE;
										FillSolidRect(hDC, rcTemp, m_pBar->m_crXPPressed);

										DrawCross(hDC, 
												  m_rcClose.left, 
												  m_rcClose.top, 
												  m_rcClose.right, 
												  m_rcClose.bottom, 
												  m_nPenThickness,
												  GetSysColor(COLOR_BTNTEXT));
									}
								}
								else if (bOver)
								{
									bOver = FALSE;
									FillSolidRect(hDC, rcTemp, m_pBar->m_crXPSelectedColor);

									DrawCross(hDC, 
											  m_rcClose.left, 
											  m_rcClose.top, 
											  m_rcClose.right, 
											  m_rcClose.bottom, 
											  m_nPenThickness,
											  GetSysColor(COLOR_BTNTEXT));
								}
							}
							break;

						default:
							// Just dispatch rest of the messages
							DispatchMessage(&msg);
							break;
						}
					}
					ReleaseDC(m_hWnd, hDC);
					if (GetCapture() == m_hWnd)
						ReleaseCapture();

					FillSolidRect(hDC, rcTemp, m_pBar->m_crXPFloatCaptionBackground);

					DrawCross(hDC, 
							  m_rcClose.left, 
							  m_rcClose.top, 
							  m_rcClose.right, 
							  m_rcClose.bottom, 
							  m_nPenThickness,
							  GetSysColor(COLOR_BTNTEXT));
				}
			}
			else
			{
				ScreenToClient(pt);
				CRect rcBand;
				if (m_pBand->GetBandRect(hWndCapture, rcBand))
				{
					POINT ptDragStart = pt;

					BOOL bStartDrag = FALSE;
					BOOL bProcessing = TRUE;
					SetCapture(hWndCapture);
					if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
						SetCursor(LoadCursor(NULL, IDC_SIZEALL));
					MSG msg;
					DWORD nDragStartTick = GetTickCount();
					while (GetCapture() == hWndCapture && bProcessing)
					{
						GetMessage(&msg, NULL, 0, 0);
						switch (msg.message)
						{
						case WM_KEYDOWN:
							if (VK_ESCAPE == msg.wParam)
								bProcessing = FALSE;
							break;

						case WM_RBUTTONUP:
						case WM_LBUTTONUP:
							bProcessing = FALSE;
							break;

						case WM_MOUSEMOVE:
							{
								POINT pt = {LOWORD(msg.lParam), HIWORD(msg.lParam)};

								if (GetTickCount()-nDragStartTick < GetGlobals().m_nDragDelay &&
									abs(pt.x-ptDragStart.x) < GetGlobals().m_nDragDist &&
									abs(pt.y-ptDragStart.y) < GetGlobals().m_nDragDist)
								{
									break;
								}
								bStartDrag = TRUE;
								bProcessing = FALSE;
							}
							break;

						default:
							// Just dispatch rest of the messages
							DispatchMessage(&msg);
							break;
						}
					}
					if (GetCapture() == hWndCapture)
						ReleaseCapture();
					if (bStartDrag)
						m_pBand->m_pDockMgr->StartDrag(m_pBand, pt);
				}
			}
		}
		catch (...)
		{
			assert(FALSE);
		}
	}
	return TRUE;
}

//
// DragWindow
//

void CMiniWin::DragWindow(const CRect& rc)
{
	SetWindowPos(NULL, 
				 rc.left, 
				 rc.top, 
				 rc.Width(), 
				 rc.Height(), 
				 SWP_NOZORDER | SWP_NOACTIVATE);
}

//
// DrawDragRect
//

void CMiniWin::DrawDragRect(const CRect& rc, 
							int          nThickness,
							CRect*       prcLast,
							int          nLastThickness)
{
	HWND hWndDesktop = GetDesktopWindow();
	HDC hDC = GetWindowDC(hWndDesktop);
	if (hDC)
	{
		DrawDragRect2(hDC, rc, nThickness, prcLast, nLastThickness);
		ReleaseDC (hWndDesktop, hDC);
	}
}

