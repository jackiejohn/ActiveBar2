//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "ipserver.h"
#include "Dock.h"
#include "Band.h"
#include "Bands.h"
#include "Tool.h"
#include "popupwin.h"
#include "map.h"
#include "bar.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// WindowFromPointEx
//

HWND WindowFromPointEx(POINT ptScreen)
{
	POINT pt = ptScreen;
	HWND hWndPoint = WindowFromPoint(pt);
	if (NULL == hWndPoint)
		return NULL;

	ScreenToClient (hWndPoint, &pt);
	
	HWND hWndChild;
	while (TRUE)
	{
		// Search through all child windows at this point. 
		hWndChild = ChildWindowFromPoint(hWndPoint, pt);
 
		if (NULL == hWndChild || hWndChild == hWndPoint || !IsWindowVisible(hWndChild))
			break;

		hWndPoint = hWndChild;
		pt = ptScreen;
		ScreenToClient(hWndPoint, &pt);
	}
	return hWndPoint;
}

//
// ChildWindowFromPointEx
//

HWND ChildWindowFromPointEx(HWND hWnd, POINT& pt)
{
	HWND hWndChild = ChildWindowFromPoint(hWnd, pt);
	HWND hWndPrevChild = NULL;
	while (hWndChild && hWndChild != hWndPrevChild)
	{
		hWndPrevChild = hWndChild;
		hWndChild = ChildWindowFromPoint(hWndChild, pt);
	}
	return hWndPrevChild;
}

//
// EnterWhatsThisHelpMode
//

BOOL CBar::EnterWhatsThisHelpMode(HWND hWnd)
{
	HCURSOR hPrevCursor = SetCursor(LoadCursor(NULL, IDC_HELP));
	
	m_bWhatsThisHelp = TRUE;
	GetGlobals().m_pWhatsThisHelpActiveBar = this;			

	int    nToolIndex;
	MSG    msg;
	BOOL   bHandled = TRUE;
	BOOL   bProcessing = TRUE;
	HWND   hWndTemp = m_hWnd;
	HWND   hChild;
	CDock* pDock = NULL;
	CBand* pBand = NULL;
	CTool* pTool = NULL;
	POINT  pt;

	SetCapture(hWndTemp);
	while (bProcessing && hWndTemp == GetCapture())
	{
		if (0 == GetMessage(&msg, NULL, 0, 0))
		{
			bProcessing = FALSE;
			continue;
		}

		switch (msg.message)
		{
		case WM_KEYDOWN:
			switch (msg.wParam)
			{
			case VK_ESCAPE:
				bProcessing = FALSE;
				break;
			}
			break;

		case WM_LBUTTONDOWN:
			{
				pt = msg.pt;

				//
				// Find which dock area was clicked on
				//

				pDock = m_pDockMgr->HitTest(pt);
				if (NULL == pDock)
				{
					BOOL bFound = FALSE;
					HWND  hWndBand;
					CRect rcBand;
					int nCount = m_pBands->GetBandCount();
					for (int nBand = 0; nBand < nCount; nBand++)
					{
						pBand = m_pBands->GetBand(nBand);
						if (ddDAFloat != pBand->bpV1.m_daDockingArea)
							continue;

						if (!pBand->GetBandRect(hWndBand, rcBand))
							continue;

						if (!GetWindowRect(hWndBand, &rcBand))
							continue;

						if (PtInRect(&rcBand, pt))
						{
							bFound = TRUE;
							ScreenToClient(hWndBand, &pt);
							break;
						}
					}
					if (!bFound)
					{
						ScreenToClient(hWnd, &pt);
						hChild = ChildWindowFromPointEx(hWnd, pt);
						bProcessing = FALSE;
						bHandled = FALSE;
						continue;
					}
				}
				else
				{
					//
					// Make the point relative to the dock area
					//
					
					pDock->ScreenToClient(pt);
					pBand = pDock->HitTest(pt);
					if (NULL == pBand)
 						continue;

					//
					// Make the point relative to the band
					//
					
					pt.x -= pBand->m_rcDock.left;
					pt.y -= pBand->m_rcDock.top;
				}

				CBand* pBandHitTest;
				pTool = pBand->HitTestTools(pBandHitTest, pt, nToolIndex);
				if (NULL == pTool)
 					continue;

				//
				// Clicked on ActiveBar
				//

				if (pTool->HasSubBand() && (ddTTButtonDropDown != pTool->tpV1.m_ttTools || 
											pTool->m_bDropDownPressed))
				{
					if (CTool::ddTTMoreTools == pTool->tpV1.m_ttTools)
					{
						pBand->m_nCurrentTool = nToolIndex;
						pTool->OnLButtonDown(msg.wParam, pt);
						pBand->m_nCurrentTool = -1;
						if (nToolIndex < pBand->GetVisibleToolCount())
							pBand->InvalidateToolRect(nToolIndex);
					}
					else
					{
						//
						// On a menu item
						//

						pBand->EnterMenuLoop(nToolIndex, pTool, CBand::ENTERMENULOOP_CLICK, 0);
					}
					bProcessing = FALSE;
				}
				else
				{
					//
					// On a non menu item
					//

					FireWhatsThisHelp((Band*)pBand, (Tool*)pTool, pTool->tpV1.m_nHelpContextId);
					bProcessing = FALSE;
				}
			}
			break;

		default:
			DispatchMessage(&msg);
		}
	}

	SetCursor(hPrevCursor);
	ReleaseCapture();
	
	if (!bHandled)
	{
		//
		// If it was not clicked on ActiveBar
		//

		HELPINFO helpInfo;
		helpInfo.cbSize = sizeof(HELPINFO); 
		helpInfo.iContextType = HELPINFO_WINDOW;
		GetCursorPos(&helpInfo.MousePos);
		if (hChild)
		{
			//
			// Clicked on a child window
			//

			helpInfo.dwContextId = 0; 
			helpInfo.iCtrlId = GetWindowLong(hChild, GWL_ID); 
			helpInfo.hItemHandle = hChild; 
		}
		else
		{
			//
			// Clicked on the Window
			//

			helpInfo.dwContextId = 0xFFFFFFFF; 
			helpInfo.iCtrlId = -1;
			helpInfo.hItemHandle = hWnd;
		}
		SendMessage(hWnd, WM_HELP, 0, (LPARAM)&helpInfo);
	}
	m_bWhatsThisHelp = FALSE;
	GetGlobals().m_pWhatsThisHelpActiveBar = NULL;
	return bHandled;
}
