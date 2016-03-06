#include "precomp.h"
#include <olectl.h>
#include "ipserver.h"
#include <stddef.h>  
#include "support.h"
#include "resource.h"
#include "debug.h"
#include "bar.h"
#include "xevents.h"
#include "bands.h"
#include "band.h"
#include "dock.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// OCXKeyState
//

short OCXKeyState()
{
    BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
    BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
    BOOL bAlt   = (GetKeyState(VK_MENU) < 0);
    return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}

void CBar::CacheSmButtonSize()
{
	int cyCaption = GetSystemMetrics(SM_CYCAPTION);
	m_sizeSMButton.cy = cyCaption - 5;
	m_sizeSMButton.cx = m_sizeSMButton.cy + 2;
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
		HWND hWndActiveMDIChild = (HWND)SendMessage(m_hWndMDIClient,
												    WM_MDIGETACTIVE,
												    0,
												    (LPARAM)&bMaximized);
		if (hWndActiveMDIChild && bMaximized)
		{
			DWORD dwStyleChild = GetWindowLong(hWndActiveMDIChild, GWL_STYLE);
			if (dwStyleChild & WS_SYSMENU)
			{
				if (dwStyleChild & WS_MINIMIZEBOX)
					dwNewButtonState |= 4;

				if (dwStyleChild & WS_MAXIMIZEBOX)
					dwNewButtonState |= 2;

				if (dwStyleChild & WS_SYSMENU)
					dwNewButtonState |= 1; // close box
			}
		}

		if (m_dwMdiButtons != dwNewButtonState)
		{
			m_dwMdiButtons = dwNewButtonState;
			m_bNeedRecalcLayout = TRUE;
			PostMessage(m_hWnd, WM_REFRESHMENUBAND, 0, 0);
		}
	}
}

