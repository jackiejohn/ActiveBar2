//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <stdio.h>
#include "Bar.h"
#include "Dock.h"
#include "Debug.h"
#include "IPServer.h"
#include "Globals.h"
#include "ShortCuts.h"
#include "Utility.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// RemoveWindowFromWindowList
//

BOOL CBar::RemoveWindowFromWindowList(HWND hWndChild)
{
	try
	{
		int nCount = m_aChildWindows.GetSize();
		for (int nWnd = 0; nWnd < nCount; nWnd++)
		{
			if (hWndChild == m_aChildWindows.GetAt(nWnd))
			{
				m_aChildWindows.RemoveAt(nWnd);
				return TRUE;
			}
		}
		return FALSE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// NotifyFloatingWindows
//

void CBar::NotifyFloatingWindows(DWORD dwFlags)
{
	try
	{
		// get top level parent frame window first unless this is a child window
		HWND hWndParent = GetDockWindow();
		if (dwFlags & (FS_DEACTIVATE|FS_ACTIVATE) && IsWindow(hWndParent))
		{
			// update parent window activation state
			BOOL bActivate = !(dwFlags & FS_DEACTIVATE);
			BOOL bEnabled = ::IsWindowEnabled(hWndParent);

			if (bActivate && bEnabled)
			{
				// Excel will try to Activate itself when it receives a
				// WM_NCACTIVATE so we need to keep it from doing that here.
				m_nFlags |= WF_KEEPMINIACTIVE;
				SendMessage(hWndParent, WM_NCACTIVATE, TRUE, 0);
				m_nFlags &= ~WF_KEEPMINIACTIVE;
			}
			else
				SendMessage(hWndParent, WM_NCACTIVATE, FALSE, 0);
		}

		// then update the state of all floating windows owned by the parent
		HWND hWnd2 = ::GetWindow(::GetDesktopWindow(), GW_CHILD);
		while (hWnd2)
		{
			if (IsDescendant(hWndParent, hWnd2) && hWndParent != hWnd2)
				::SendMessage(hWnd2, GetGlobals().WM_FLOATSTATUS, dwFlags, 0);
			hWnd2 = ::GetWindow(hWnd2, GW_HWNDNEXT);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// FrameWindowProc
//

LRESULT CALLBACK FrameWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		LRESULT lResult;
		CBar* pBar = NULL;
		if (!GetGlobals().m_pmapBar->Lookup((LPVOID)hWnd, (void*&)pBar))
			return DefFrameProc(hWnd, NULL, nMsg, wParam, lParam);

		WNDPROC pWndProc = pBar->FrameProc();
		if (NULL == pWndProc)
			return DefWindowProc(hWnd, nMsg, wParam, lParam);

		switch (nMsg)
		{
		case WM_DROPFILES:
			try
			{
				// handle of internal drop structure 
				HDROP hDrop = (HDROP)wParam;

				UINT nCount = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0); 
				if (nCount < 1)
					return FALSE;

				int nResult;
				TCHAR szBuffer[_MAX_PATH];
				for (UINT nIndex = 0; nIndex < nCount; nIndex++)
				{
					nResult = DragQueryFile(hDrop, nIndex, szBuffer, _MAX_PATH); 
					if (nResult > _MAX_PATH)
						continue;

					MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
					BSTR bstrBuffer = SysAllocString(wBuffer);
					if (bstrBuffer)
					{
						pBar->FireFileDrop((Band*)NULL, bstrBuffer);
						SysFreeString(bstrBuffer);
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}	
			break;

		case WM_ACTIVATEAPP:
			{
				pBar->m_bApplicationActive = (BOOL)wParam;
				CBand* pMenuBand = pBar->GetMenuBand();
				if (pMenuBand)
					pMenuBand->Refresh();
				if (!wParam)
					pBar->DestroyPopups();
			}
			break;

		case WM_ACTIVATE:
			{
				HWND hWndOther = (HWND)lParam;
				int nState = LOWORD(wParam);
				BOOL bMinimized = HIWORD(wParam);
				if (bMinimized)
				{
					pBar->HideMiniWins();
				}

				lResult = CallWindowProc(pWndProc, hWnd, nMsg, wParam, lParam);
				// get top level frame unless this is a child window
				// determine if window should be active or not
				HWND hWndTopLevel = hWnd;

				HWND hWndActive = (nState == WA_INACTIVE ? hWndOther : hWnd);
				BOOL bStayActive = (hWndTopLevel == hWndActive);
				if (!bStayActive && IsWindow(hWndActive))
					bStayActive = (IsWindow(hWndTopLevel) && IsChild(hWndTopLevel, hWndActive) || (0 != SendMessage(hWndActive, GetGlobals().WM_FLOATSTATUS, FS_SYNCACTIVE, 0)));

				pBar->m_nFlags &= ~WF_STAYACTIVE;
				if (bStayActive)
				{
					pBar->m_nFlags |= WF_STAYACTIVE;
				}

				// sync floating windows to the new state
				pBar->NotifyFloatingWindows(bStayActive ? FS_ACTIVATE : FS_DEACTIVATE);

				return lResult;
			}
			break;

		case WM_NCACTIVATE:
			{
				// stay active if WF_STAYACTIVE bit is on
				BOOL bActive = (BOOL)wParam;
				if (pBar->m_nFlags & WF_STAYACTIVE)
				{
					bActive = TRUE;
				}

				// but do not stay active if the window is disabled
				if (!IsWindowEnabled(hWnd))
				{
					bActive = FALSE;
				}

				// do not call the base class because it will call Default()
				//  and we may have changed bActive.
				return CallWindowProc(pWndProc, hWnd, nMsg, bActive, 0L);
			}
			break;

		case WM_ENABLE:
			{
				BOOL bEnable = (BOOL)wParam;  // activation and minimization options
				if (bEnable && (pBar->m_nFlags & WF_STAYDISABLED))
				{
					// Work around for MAPI support. This makes sure the main window
					// remains disabled even when the mail system is booting.
					EnableWindow(hWnd, FALSE);
					::SetFocus(NULL);
					return 0;
				}

				// only for top-level (and non-owned) windows
				if (GetParent(hWnd))
				{
					return 0;
				}

				// force WM_NCACTIVATE because Windows may think it is unecessary
				if (bEnable && (pBar->m_nFlags & WF_STAYACTIVE))
				{
					try
					{
						// CR 2251 changed from a send to a post
						PostMessage(hWnd, WM_NCACTIVATE, TRUE, 0);
					}
					catch (...)
					{
					}
				}
				// force WM_NCACTIVATE for floating windows too
				pBar->NotifyFloatingWindows(bEnable ? FS_ENABLE : FS_DISABLE);

				pBar->m_theFormsAndControls.EnableForms(bEnable);
			}
			break;

		case WM_MENUCHAR:
			{
				if (!pBar->AmbientUserMode())
					break;

				if (!(pBar->m_bMenuLoop || pBar->m_bClickLoop))
				{
					int nCode = HIWORD(wParam);
					if (MF_SYSMENU == nCode)
					{
						CTool* pTool;
						TCHAR cAccel = _totupper((TCHAR)LOWORD(wParam));
						if (pBar->m_mapMenuBar.Lookup(cAccel, pTool))
						{
							PostMessage(hWnd, WM_SYSCOMMAND, SC_KEYMENU, cAccel);
							return MAKELONG(2, MNC_SELECT);
						}
					}
				}
			}
			break;

		case WM_SYSCOMMAND:
			try
			{
				switch (wParam)
				{
				case SC_MAXIMIZE:
					lResult = CallWindowProc(pWndProc, hWnd, nMsg, wParam, lParam);
					pBar->RecalcLayout();
					return lResult;

				case SC_CLOSE:
					{
						short nCancel = 0;
						pBar->FireQueryUnload(&nCancel);
						if (0 == nCancel)
						{
							pBar->DestroyMiniWins();
							pBar->m_theFormsAndControls.ShutDown();
						}
					}
					break;

				case SC_KEYMENU:
					{
						if (!pBar->AmbientUserMode() || (WS_DISABLED & GetWindowLong(hWnd, GWL_STYLE)))
							break;

						if (GetKeyState(VK_SHIFT) < 0 || GetKeyState(VK_CONTROL) < 0)
							return 0;

						if (!(pBar->m_bMenuLoop || pBar->m_bClickLoop))
						{
							if (0 == lParam)
							{
								CBand* pMenuBand = pBar->GetMenuBand();
								if (pMenuBand)
								{
									pMenuBand->EnterMenuLoop(0, NULL, CBand::ENTERMENULOOP_KEY, VK_MENU);
									return 0;
								}
							}
							else
							{
								TCHAR theChar = _totupper(lParam);
								if (pBar->DoMenuAccelator((unsigned short)theChar))
									return 0;
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

		case WM_SIZE:
			if (SIZE_RESTORED == wParam || SIZE_MAXIMIZED == wParam)
			{
				if (pBar->AmbientUserMode())
				{
					if (pBar->m_pDockMgr)
						pBar->m_pDockMgr->ResizeDockableForms();
					lResult = CallWindowProc(pWndProc, hWnd, nMsg, wParam, lParam);
					if (pBar->m_pDockMgr)
						pBar->m_pDockMgr->ResizeDockableForms();
					pBar->RecalcLayout();
					return lResult;
				}
				else
				{
					pBar->RecalcLayout();
					pBar->ViewChanged();
					SetWindowPos(pBar->m_hWnd, HWND_TOP, pBar->m_rcFrame.left, pBar->m_rcFrame.top, pBar->m_rcFrame.Width(), pBar->m_rcFrame.Height(), SWP_SHOWWINDOW|SWP_NOACTIVATE);
				}
			}
			break;

		case WM_SYSCOLORCHANGE:
		case WM_SETTINGCHANGE:
			pBar->OnSysColorChanged();
			break;

		case WM_WINDOWPOSCHANGED:
			{
				if (pBar->m_bWindowPosChangeMessage)
					break;

				pBar->m_bWindowPosChangeMessage = TRUE;
				pBar->m_theErrorObject.m_nAsyncError++;

				LPWINDOWPOS pWP = (LPWINDOWPOS)lParam;	
				if (pWP && ~(SWP_FRAMECHANGED & pWP->flags))
					pBar->UpdateChildButtonState();

				pBar->m_theErrorObject.m_nAsyncError--;
				pBar->m_bWindowPosChangeMessage = FALSE;
			}
			break;
		}
		try
		{
			return CallWindowProc(pWndProc, hWnd, nMsg, wParam, lParam);
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
			return DefWindowProc(hWnd, nMsg, wParam, lParam);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return DefWindowProc(hWnd, nMsg, wParam, lParam);
	}
}

//
// MDIClientWindowProc
//

LRESULT CALLBACK MDIClientWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		CBar* pBar = NULL;
		if (!GetGlobals().m_pmapBar->Lookup((LPVOID)hWnd, (void*&)pBar))
			return DefWindowProc(hWnd, nMsg, wParam, lParam);

		WNDPROC pWndProc = pBar->MDIClientProc();
		if (NULL == pWndProc)
			return DefWindowProc(hWnd, nMsg, wParam, lParam);

		if (pBar->m_bBandDestroyLock)
			return NULL;

		switch (nMsg)
		{
		case WM_LBUTTONDOWN:
			if (pBar->m_bCustomization || pBar->m_bWhatsThisHelp)
			{
				MessageBeep(1000);
				return 0;
			}
			break;

		case WM_MDICREATE:
			{
				HWND hWndChild = (HWND)CallWindowProc(pWndProc, hWnd, nMsg, wParam, lParam);
				if (IsWindow(hWndChild))
					pBar->m_aChildWindows.Add(hWndChild);
				return (LRESULT)hWndChild;
			}
			break;

		case WM_MDIACTIVATE:
			{
				HWND hWndChild = (HWND)wParam;
				if (hWndChild)
				{
					pBar->m_hWndMenuChild = hWndChild;
					pBar->m_theChildMenus.SwitchMenus(hWndChild);
				}
			}
			break;

		case WM_MDIDESTROY:
			{
				try
				{
					LRESULT lResult;
					HWND hWndChild = (HWND)wParam;
					if (hWndChild)
					{
						pBar->RemoveWindowFromWindowList(hWndChild);
						pBar->m_theChildMenus.UnregisterChildWindow(hWndChild);

						lResult = CallWindowProc(pWndProc, hWnd, nMsg, wParam, lParam);

						HWND hWndActive = (HWND)SendMessage(pBar->m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
						pBar->m_theChildMenus.SwitchMenus(hWndActive);
					}
					PostMessage(pBar->m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
					return lResult;
				}
				CATCH
				{
					REPORTEXCEPTION(__FILE__, __LINE__)
					return DefWindowProc(hWnd, nMsg, wParam, lParam);
				}
			}
			break;

		case WM_MDISETMENU:
			{
				CBand* pBand = pBar->GetMenuBand();
				if (pBand && VARIANT_TRUE == pBand->bpV1.m_vbVisible)
				{
					SetMenu(pBar->GetDockWindow(), NULL);
					HWND hWndActive = (HWND)SendMessage(pBar->m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
					if (pBar->m_hWndMenuChild != hWndActive)
					{
						pBar->m_hWndMenuChild = hWndActive;
						pBar->m_theChildMenus.SwitchMenus(hWndActive);
					}
					else
						PostMessage(pBar->m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
					return 0;
				}
				PostMessage(pBar->m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
			}
			break;
			
		case WM_MDIREFRESHMENU:
			{
				CBand* pBand = pBar->GetMenuBand();
				if (pBand && VARIANT_TRUE == pBand->bpV1.m_vbVisible)
					return 0;
			}
			break;

		case WM_WINDOWPOSCHANGING:
			{
				if (pBar->m_bFirstRecalcLayout)
				{
					LPWINDOWPOS pWP = (LPWINDOWPOS)lParam;
					CRect& rc = pBar->MDIClientRect();
					pWP->x = rc.left;
					pWP->y = rc.top;
					pWP->cx = rc.Width();
					pWP->cy = rc.Height();
				}
			}
			break;

		case WM_DROPFILES:
			try
			{
				// handle of internal drop structure 
				HDROP hDrop = (HDROP)wParam;

				UINT nCount = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0); 
				if (nCount < 1)
					return FALSE;

				int nResult;
				TCHAR szBuffer[_MAX_PATH];
				for (UINT nIndex = 0; nIndex < nCount; nIndex++)
				{
					nResult = DragQueryFile(hDrop, nIndex, szBuffer, _MAX_PATH); 
					if (nResult > _MAX_PATH)
						continue;

					MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
					BSTR bstrBuffer = SysAllocString(wBuffer);
					if (bstrBuffer)
					{
						pBar->FireFileDrop((Band*)NULL, bstrBuffer);
						SysFreeString(bstrBuffer);
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}	
			break;
		}
		return CallWindowProc(pWndProc, hWnd, nMsg, wParam, lParam);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return DefWindowProc(hWnd, nMsg, wParam, lParam);
	}
}

//
// FormWindowProc
//

LRESULT CALLBACK FormWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	CBar* pBar = NULL;
	try
	{
		if (!GetGlobals().m_pmapBar->Lookup((LPVOID)hWnd, (void*&)pBar))
			return DefWindowProc(hWnd, nMsg, wParam, lParam);

		switch (nMsg)
		{
		case WM_ACTIVATEAPP:
			{
				pBar->m_bApplicationActive = (BOOL)wParam;
				CBand* pMenuBand = pBar->GetMenuBand();
				if (pMenuBand)
					pMenuBand->Refresh();

				if (!wParam)
					pBar->DestroyPopups();
			}
			break;

		case WM_ACTIVATE:
			{
				HWND hWndOther = (HWND)lParam;
				int nState = LOWORD(wParam);
				BOOL bMinimized = HIWORD(wParam);
				if (bMinimized)
				{
					pBar->HideMiniWins();
				}


				LRESULT lResult = CallWindowProc((WNDPROC)pBar->FrameProc(), hWnd, nMsg, wParam, lParam);
				// get top level frame unless this is a child window
				// determine if window should be active or not
				HWND hWndTopLevel = hWnd;

				HWND hWndActive = (nState == WA_INACTIVE ? hWndOther : hWnd);
				BOOL bStayActive = (nState == WA_INACTIVE ? FALSE : TRUE);
				if (!bStayActive && IsWindow(hWndActive))
					bStayActive = (IsChild(hWndTopLevel, hWndActive) || (0 != SendMessage(hWndActive, GetGlobals().WM_FLOATSTATUS, FS_SYNCACTIVE, 0)));

				if (IsWindowVisible(GetParent(hWnd)))
				{
					pBar->m_nFlags &= ~WF_STAYACTIVE;
					if (bStayActive )
						pBar->m_nFlags |= WF_STAYACTIVE;
				}

				pBar->NotifyFloatingWindows(bStayActive ? FS_ACTIVATE : FS_DEACTIVATE);

				return lResult;
			}
			break;

		case WM_NCACTIVATE:
			{
				BOOL bActive = (BOOL)wParam;
				// stay active if WF_STAYACTIVE bit is on
				if (pBar->m_nFlags & WF_STAYACTIVE)
					bActive = TRUE;

				// but do not stay active if the window is disabled
				if (!IsWindowEnabled(hWnd))
					bActive = FALSE;

				// do not call the base class because it will call Default()
				//  and we may have changed bActive.
				return CallWindowProc((WNDPROC)pBar->FrameProc(), hWnd, nMsg, bActive, lParam);
			}
			break;

		case WM_ENABLE:
			{
				BOOL bEnable = (BOOL)wParam;  // activation and minimization options
				if (bEnable && (pBar->m_nFlags & WF_STAYDISABLED))
				{
					// Work around for MAPI support. This makes sure the main window
					// remains disabled even when the mail system is booting.
					EnableWindow(hWnd, FALSE);
					::SetFocus(NULL);
					return 0;
				}

				// only for top-level (and non-owned) windows
				if (GetParent(hWnd))
					return 0;

				// force WM_NCACTIVATE because Windows may think it is unecessary
				if (bEnable && (pBar->m_nFlags & WF_STAYACTIVE))
					SendMessage(hWnd, WM_NCACTIVATE, TRUE, 0);

				// force WM_NCACTIVATE for floating windows too
				pBar->NotifyFloatingWindows(bEnable ? FS_ENABLE : FS_DISABLE);

				pBar->m_theFormsAndControls.EnableForms(bEnable);
			}
			break;

		case WM_DROPFILES:
			try
			{
				// handle of internal drop structure 
				HDROP hDrop = (HDROP)wParam;

				UINT nCount = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0); 
				if (nCount < 1)
					return FALSE;

				int nResult;
				TCHAR szBuffer[_MAX_PATH];
				for (UINT nIndex = 0; nIndex < nCount; nIndex++)
				{
					nResult = DragQueryFile(hDrop, nIndex, szBuffer, _MAX_PATH); 
					if (nResult > _MAX_PATH)
						continue;

					MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
					BSTR bstrBuffer = SysAllocString(wBuffer);
					if (bstrBuffer)
					{
						pBar->FireFileDrop((Band*)NULL, bstrBuffer);
						SysFreeString(bstrBuffer);
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}	
			break;

		case WM_SYSCOLORCHANGE:
		case WM_SETTINGCHANGE:
			try
			{
				pBar->OnSysColorChanged();
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
			break;

		case WM_MENUCHAR:
			{
				if (!pBar->AmbientUserMode())
					break;

				if (!(pBar->m_bMenuLoop || pBar->m_bClickLoop))
				{
					int nCode = HIWORD(wParam);
					if (MF_SYSMENU == nCode)
					{
						CTool* pTool;
						TCHAR cAccel = _totupper((TCHAR)LOWORD(wParam));
						if (pBar->m_mapMenuBar.Lookup(cAccel, pTool))
						{
							PostMessage(hWnd, WM_SYSCOMMAND, SC_KEYMENU, cAccel);
							return MAKELONG(2, MNC_SELECT);
						}
					}
				}
			}
			break;

		case WM_SYSCOMMAND:
			try
			{
				switch (wParam)
				{
				case SC_CONTEXTHELP:
					pBar->EnterWhatsThisHelpMode(hWnd);
					break;

				case SC_CLOSE:
					{
						short nCancel = 0;
						pBar->FireQueryUnload(&nCancel);
						if (0 == nCancel)
						{
							pBar->DestroyMiniWins();
							pBar->m_theFormsAndControls.ShutDown();
						}
					}
					break;

				case SC_KEYMENU:
					{
						if (!pBar->AmbientUserMode() || (WS_DISABLED & GetWindowLong(hWnd, GWL_STYLE)))
							break;

						if (GetKeyState(VK_SHIFT) < 0 || GetKeyState(VK_CONTROL) < 0)
							return 0;

						if (!(pBar->m_bMenuLoop || pBar->m_bClickLoop))
						{
							if (0 == lParam)
							{
								CBand* pMenuBand = pBar->GetMenuBand();
								if (pMenuBand)
								{
									pMenuBand->EnterMenuLoop(0, NULL, CBand::ENTERMENULOOP_KEY, VK_MENU);
									return 0;
								}
							}
							else
							{
								TCHAR theChar = _totupper(lParam);
								if (pBar->DoMenuAccelator((unsigned short)theChar))
									return 0;
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

		case WM_SIZE:
			try
			{
				LRESULT lResult = CallWindowProc((WNDPROC)pBar->FrameProc(), hWnd, nMsg, wParam, lParam);
				if (pBar->m_pDockMgr)
					pBar->m_pDockMgr->ResizeDockableForms();
				pBar->RecalcLayout();
				if (pBar->m_pInPlaceSite) 
					pBar->m_pInPlaceSite->OnPosRectChange(&pBar->m_rcFrame);
				else 
					SetWindowPos(pBar->m_hWnd, HWND_TOP, pBar->m_rcFrame.left, pBar->m_rcFrame.top, pBar->m_rcFrame.Width(), pBar->m_rcFrame.Height(), SWP_SHOWWINDOW|SWP_NOACTIVATE);
				return lResult;
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
			break;
		}
		return CallWindowProc((WNDPROC)pBar->FrameProc(), hWnd, nMsg, wParam, lParam);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return DefWindowProc(hWnd, nMsg, wParam, lParam);
	}
}

//
// CacheSmButtonSize
//

void CBar::CacheSmButtonSize()
{
	CBand* pMenuBand = GetMenuBand();
	if (pMenuBand)
	{
		m_sizeSMButton.cy = pMenuBand->Height() + 2;
		m_sizeSMButton.cx = m_sizeSMButton.cy;
	}
/*	if (VARIANT_TRUE == bpV1.m_vbXPLook)
	{
		CBand* pMenuBand = GetMenuBand();
		if (pMenuBand)
		{
			m_sizeSMButton.cy = pMenuBand->Height();
			m_sizeSMButton.cx = m_sizeSMButton.cy;
		}
	}
	else
	{
			m_sizeSMButton.cy = pMenuBand->Height();
			m_sizeSMButton.cx = m_sizeSMButton.cy;
//		int cyCaption = GetSystemMetrics(SM_CYCAPTION);
//		m_sizeSMButton.cy = cyCaption - 5;
//		m_sizeSMButton.cx = m_sizeSMButton.cy + 2;
	}*/
}

//
// UpdateChildButtonState
// 

void CBar::UpdateChildButtonState()
{
	if (m_hWndMDIClient)
	{
		DWORD dwNewButtonState = 0;
		BOOL bMaximized = FALSE;
		HWND hWndActiveMDIChild = (HWND)SendMessage(m_hWndMDIClient, WM_MDIGETACTIVE, 0, (LPARAM)&bMaximized);
		if (hWndActiveMDIChild && bMaximized)
		{
			DWORD dwStyleChild = GetWindowLong(hWndActiveMDIChild, GWL_STYLE);
			if (dwStyleChild & WS_SYSMENU)
			{
				if (WS_MINIMIZEBOX & dwStyleChild)
					dwNewButtonState |= 4;

				if (WS_MAXIMIZEBOX & dwStyleChild)
					dwNewButtonState |= 2;

				if (WS_SYSMENU & dwStyleChild)
					dwNewButtonState |= 8 | 1; // Close box and system menu
			}
		}
		if (m_dwMdiButtons != dwNewButtonState)
		{
			m_dwMdiButtons = dwNewButtonState;
			m_bNeedRecalcLayout = TRUE;
			PostMessage(m_hWnd, GetGlobals().WM_REFRESHMENUBAND, 0, 0);
		}
	}
}

//
// HandleSpecialKeys
//

BOOL HandleSpecialKeys(CBar* pBar, WPARAM wParam, LPARAM lParam)
{
	static UpdateTabTool theUpdateTabTool;
	if (!(HIWORD(lParam) & KF_UP))
	{
		switch (wParam)
		{
		case VK_TAB:
			{
				CBand* pBand;
				CTool* pTool;
				HWND hWndBand;
				int nTool; 
				if (0x8000 & GetKeyState(VK_SHIFT))
				{
					//
					// Go in reverse
					//

					if (pBar->TabBackward(pBand, hWndBand, pTool, nTool))
					{
						theUpdateTabTool.pBand = pBand;
						theUpdateTabTool.pTool = pTool;
						theUpdateTabTool.nTool = nTool;
						theUpdateTabTool.hWndBand = hWndBand;
						PostMessage(pBar->m_hWnd, GetGlobals().WM_UPDATETABTOOL, (WPARAM)&theUpdateTabTool, 0);
						return TRUE;
					}
				}
				else
				{
					if (pBar->TabForward(pBand, hWndBand, pTool, nTool))
					{
						theUpdateTabTool.pBand = pBand;
						theUpdateTabTool.pTool = pTool;
						theUpdateTabTool.nTool = nTool;
						theUpdateTabTool.hWndBand = hWndBand;
						PostMessage(pBar->m_hWnd, GetGlobals().WM_UPDATETABTOOL, (WPARAM)&theUpdateTabTool, 0);
						return TRUE;
					}
				}
			}
			break;
		}
	}
	return FALSE;
}

//
// HandleShortCutsKeys
//

BOOL HandleShortCutsKeys(CBar* pBar, HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	switch (wParam)
	{
	case VK_CAPITAL:
	case VK_INSERT:
	case VK_NUMLOCK:
	case VK_SCROLL:
		{
			if (pBar->StatusBand())
				pBar->StatusBand()->Refresh();
		}
		break;
	}

	MSG msg = {hWnd, WM_KEYDOWN, wParam, lParam, 0, 0, 0};
	if (!pBar->m_bMenuLoop && pBar->m_pShortCuts->TranslateAccelerator(&msg))
		return TRUE;
	return FALSE;
}

//
// KeyboardProc
//

LRESULT CALLBACK KeyBoardHook(int nCode, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (VK_F4 == wParam && (HIWORD(lParam) & KF_ALTDOWN))
		{
			TRACE(5, "VK_F4 == wParam && (HIWORD(lParam) & KF_ALTDOWN)\n")
			return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);
		}

		CBar* pBar = NULL;
		HWND hWndActive = GetActiveWindow();

		//
		// First check if there is an Active MDI Client
		//

		HWND hWndMdiClient = ::GetWindow(hWndActive, GW_CHILD);
		if (IsWindow(hWndMdiClient))
		{
			TCHAR szClassName[120];
			GetClassName(hWndMdiClient, szClassName, 120);
			if (0 == lstrcmpi(szClassName, _T("MDICLIENT")))
			{
				BOOL bMaximized;
				HWND hWndActiveMDIChild = (HWND)SendMessage(hWndMdiClient, WM_MDIGETACTIVE, 0, (LPARAM)&bMaximized);
				if (hWndActiveMDIChild && GetGlobals().m_pmapAccelator->Lookup((LPVOID)hWndActiveMDIChild, (void*&)pBar) && !(WS_DISABLED & GetWindowLong(hWndActiveMDIChild, GWL_STYLE)))
				{
					if (!pBar->AmbientUserMode())
						return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);

					if (HIWORD(lParam) & KF_REPEAT)
						return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);

					switch (wParam)
					{
					case VK_CONTROL:
					case VK_SHIFT:
						return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);
					}

					if (HandleShortCutsKeys(pBar, hWndActiveMDIChild, wParam, lParam))
					{
						GetGlobals().ClearAccelators();
						return -1;
					}
				}
			}
		}

		//
		// Usally the MDI Frame Window or the SDI Window
		//
		
		if (GetGlobals().m_pmapAccelator->Lookup((LPVOID)hWndActive, (void*&)pBar) && !(WS_DISABLED & GetWindowLong(hWndActive, GWL_STYLE)))
		{
			if (!pBar->AmbientUserMode())
				return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);

			if (HIWORD(lParam) & KF_REPEAT)
				return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);

			switch (wParam)
			{
			case VK_CONTROL:
			case VK_SHIFT:
				return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);
			}

			if (HandleShortCutsKeys(pBar, hWndActive, wParam, lParam))
			{
				GetGlobals().ClearAccelators();
				return -1;
			}
		}

		//
		// How check if any of the Activebars are children of the ActiveWindow
		//

		BOOL bFound = FALSE;
		HWND hWnd;
		FPOSITION posMap = GetGlobals().m_pmapAccelator->GetStartPosition();
		while (posMap)
		{
			try
			{
				GetGlobals().m_pmapAccelator->GetNextAssoc(posMap, (LPVOID&)hWnd, (LPVOID&)pBar);
				if ((IsChild(hWndActive, hWnd) || hWnd == hWndActive) && IsWindowEnabled(hWnd))
				{
					if (!pBar->AmbientUserMode())
						return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);

					if (HIWORD(lParam) & KF_REPEAT)
						return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);

					switch (wParam)
					{
					case VK_CONTROL:
					case VK_SHIFT:
						return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);
					}

					if (HandleShortCutsKeys(pBar, hWnd, wParam, lParam))
					{
						GetGlobals().ClearAccelators();
						return -1;
					}
				}
			}
			catch (...)
			{
				assert(FALSE);
			}
		}
		return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return CallNextHookEx(GetGlobals().m_hHookAccelator, nCode, wParam, lParam);
	}
}

