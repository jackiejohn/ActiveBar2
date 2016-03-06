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
#define COMPILE_MULTIMON_STUBS
#include <MultiMon.h>
#include <stddef.h>       // for offsetof()
#include "ipserver.h"
#include "debug.h"
#include "support.h"
#include "resource.h"
#include "bar.h"
#include "fwnd.h"
#include "tooltip.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// CToolTip
//

CToolTip::CToolTip(CBar* pBar)
	: m_pBar(pBar),
	  m_bstrToolTip(NULL)
{
	m_bAnimated = FALSE;
	m_dwPrevCookie = 0;
	m_dwWaitTime = 10;
	m_tttMode = eIdle;
	m_bCreateMode = TRUE;
	m_bTopMost = FALSE;
	m_dwOpenTime = m_dwCloseTime =  GetTickCount() - 5000;
	BOOL bResult = RegisterWindow(DD_WNDCLASS("ABToolTip"),
								  CS_SAVEBITS|CS_VREDRAW|CS_HREDRAW,
								  NULL,
								  ::LoadCursor(NULL, IDC_ARROW));
	assert(bResult);
}

CToolTip::~CToolTip()
{
	SysFreeString(m_bstrToolTip);
}
	
//
// CreateWin
//

BOOL CToolTip::CreateWin()
{
	CreateEx(WS_EX_TOOLWINDOW|WS_EX_TOPMOST,
		     NULL,
			 WS_POPUP,
			 0,
			 0,
			 0,
			 0,
			 NULL);
	assert(m_hWnd);
	return m_hWnd == NULL ? FALSE : TRUE;
}

//
// Show
//

void CToolTip::Show(DWORD dwCookie, BSTR bstrText, const CRect& rcBound, BOOL bTopMost)
{
	try
	{
		m_bTopMost = bTopMost;

		if (eHiddenByDelay == m_tttMode && (CRect)rcBound != m_rcBound)
			m_tttMode = eIdle;
			
		if (dwCookie == m_dwPrevCookie)
		{
			POINT ptNew;
			GetCursorPos(&ptNew);
			if (eTimerStarted == m_tttMode && (m_pt.x != ptNew.x || m_pt.y != ptNew.y))
			{
				m_pt = ptNew;
				KillTimer(m_hWnd, ID_TIMER);
				SetTimer(m_hWnd, ID_TIMER, TIMER_DURATION, 0);
			}
			return;
		}
		else if (eTimerStarted == m_tttMode)
		{
			KillTimer(m_hWnd, ID_TIMER);
			m_tttMode = eIdle;
		}

		if (NULL == bstrText || NULL == *bstrText)
		{
			Hide();
			return;
		}

		GetCursorPos(&m_pt);
		m_rcBound = rcBound;
		m_hWndBound = WindowFromPointEx(m_pt);
		m_dwPrevCookie = dwCookie;
		SysFreeString(m_bstrToolTip);
		m_bstrToolTip = SysAllocString(bstrText);

		if (eIdle == m_tttMode)
		{
			m_tttMode = eTimerStarted;
			if ((GetTickCount() - m_dwCloseTime) < 1000)
				m_tttMode = eTimerCompleted; // fast open if just closed
			else
				SetTimer(m_hWnd, ID_TIMER, TIMER_DURATION, NULL);
		}

		if (eTimerCompleted == m_tttMode)
			Open();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Open
//

void CToolTip::Open()
{
	try
	{
		SIZE sizeText;
		sizeText.cx = sizeText.cy = 0;

		MAKE_TCHARPTR_FROMWIDE(szText, m_bstrToolTip);
		HDC hDC = GetDC(m_hWnd);
		if (NULL != hDC)
		{
			HFONT hFontOld = SelectFont(hDC, m_hFont);
			GetTextExtentPoint32(hDC, szText, lstrlen(szText), &sizeText);
			SelectFont(hDC, hFontOld);
			ReleaseDC(m_hWnd, hDC);
		}
		
		POINT ptCursor;
		GetCursorPos(&ptCursor);

		CRect rc;
		rc.top = ptCursor.y + 20;
		rc.bottom = rc.top + sizeText.cy + 4;
		rc.left = ptCursor.x;
		rc.right = rc.left + sizeText.cx + 8;

		// Make sure its not off screen
		CRect rcWin;
		rcWin.right = GetSystemMetrics(SM_CXVIRTUALSCREEN);
		rcWin.bottom = GetSystemMetrics(SM_CYVIRTUALSCREEN);

		if (rc.right > rcWin.right)
		{
			rc.left = rcWin.right - rc.Width();
			rc.right = rcWin.right;
		}
		if (rc.bottom > rcWin.bottom)
		{
			rc.top = rcWin.bottom - rc.Height();
			rc.bottom = rcWin.bottom;
			POINT pt;
			GetCursorPos(&pt);
			if (pt.y > rc.top && pt.y < rc.bottom)
			{
				rc.top -= 20;
				rc.bottom -= 20;
			}
		}
		m_bCreateMode = TRUE;
		SetWindowPos(NULL,
					 0,
					 0,
					 0,
					 0,
					 SWP_NOSIZE|SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE|SWP_HIDEWINDOW);

		SetWindowPos(m_bTopMost ? HWND_TOPMOST : HWND_TOP,
					 rc.left,
					 rc.top,
					 rc.Width(),
					 rc.Height(),
					 SWP_NOACTIVATE|SWP_SHOWWINDOW);

		InvalidateRect(NULL, TRUE);

		m_bCreateMode = FALSE;
		m_dwOpenTime = GetTickCount();

		//
		// Set Timer
		//

		SetTimer(m_hWnd, ID_TIMER, TIMER_DURATION, NULL);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Hide
//

void CToolTip::Hide()
{
	try
	{
		if (eIdle == m_tttMode)
			return;

		m_bCreateMode = TRUE;
		SetWindowPos(NULL,
					 0,
					 0,
					 0,
					 0,
					 SWP_NOSIZE|SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE|SWP_HIDEWINDOW);
// The is a pointer to the activeTool this is causing problems
		m_dwPrevCookie = 0;  
		m_bCreateMode = FALSE;
		
		if (eTimerCompleted == m_tttMode)
			m_dwCloseTime = GetTickCount();
		
		ResetTimer();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// ResetTimer
//

void CToolTip::ResetTimer()
{
	KillTimer(m_hWnd, ID_TIMER);
	m_tttMode = eIdle;
}

//
// OnDraw
//

void CToolTip::OnDraw(HDC hDC, CRect& rcBound)
{
	try
	{
		int nHeight = rcBound.Height();
		int nWidth = rcBound.Width();
		
		HDC hDCPaint = m_pBar->m_ffObj.RequestDC(hDC, nWidth, nHeight);
		if (NULL == hDCPaint)
			hDCPaint = hDC;

		DrawEdge(hDCPaint, &rcBound, BDR_RAISEDOUTER, BF_RECT);
		rcBound.Inflate(-1, -1);

		FillSolidRect(hDCPaint, rcBound, GetSysColor(COLOR_INFOBK));

		HFONT hFontOld = SelectFont(hDCPaint, m_hFont);

		SetBkMode(hDCPaint, TRANSPARENT);
		SetTextColor(hDCPaint, GetSysColor(COLOR_INFOTEXT));

		MAKE_TCHARPTR_FROMWIDE(szText, m_bstrToolTip);
		BOOL bResult = DrawText(hDCPaint, szText, lstrlen(szText), &rcBound, DT_CENTER | DT_VCENTER);
		assert(bResult);
		
		SelectFont(hDCPaint, hFontOld);

		if (hDCPaint != hDC)
		{
			DWORD dwTickCountStart;
			DWORD dwPaintTime;
			int nDy = nHeight / 8;
			for (int nY = nDy; nY < nHeight && m_bAnimated; nY += nDy)
			{
				dwTickCountStart = GetTickCount(); 
				m_pBar->m_ffObj.Paint(hDC,
									  0, 
									  0, 
									  nWidth, 
									  nY, 
									  0, 
									  0);
				dwPaintTime = GetTickCount() - dwTickCountStart;
				if (dwPaintTime < m_dwWaitTime)
					Sleep(m_dwWaitTime - dwPaintTime);
			}
			if (m_bAnimated)
				m_bAnimated = FALSE;
			m_pBar->m_ffObj.Paint(hDC,
								  0, 
								  0, 
								  nWidth, 
								  nHeight, 
								  0, 
								  0);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// WindowProc
//

LRESULT CToolTip::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_WINDOWPOSCHANGED:
			if (!m_bCreateMode)
			{
				m_bCreateMode = TRUE;
				m_pBar->HideToolTips(TRUE);
				return 0;
			}
			break;

		case WM_NCDESTROY:
			FWnd::WindowProc(nMsg, wParam, lParam); // internal processing
			m_pBar->m_pTip = NULL;
			delete this;
			return 0;

		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd,&ps);
				if (hDC)
				{
					CRect rcBound;
					GetClientRect(rcBound);
					OnDraw(hDC, rcBound);
					EndPaint(m_hWnd, &ps);
					return 0;
				}
			}
			break;

		case WM_TIMER:
			try
			{
				POINT pt;
				GetCursorPos(&pt);
				if (eTimerStarted == m_tttMode)
				{
					HWND hWndUnderMouse = WindowFromPointEx(pt);
					if (PtInRect(&m_rcBound, pt) && 
						m_hWndBound == hWndUnderMouse)
					{
						m_tttMode = eTimerCompleted;
						m_bAnimated = TRUE;
						Open();
					}
					
				}
				if (eTimerCompleted == m_tttMode)
				{
					// check if still in boundrect and this app is still active
					HWND hWndUnderMouse = WindowFromPointEx(pt);
					if (!m_pBar->IsForeground(m_hWnd) || 
						!PtInRect(&m_rcBound, pt) || 
						m_hWndBound != hWndUnderMouse)
					{
						Hide();
					}
				}

				// If the tooltip is displayed more than 5 seconds shutdown
				DWORD dwCurTime = GetTickCount();
				if (eTimerCompleted == m_tttMode && 
					dwCurTime > m_dwOpenTime && 
					dwCurTime - m_dwOpenTime > 5000)
				{
					Hide();
					m_tttMode = eHiddenByDelay; // hidden by delay
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// CBar Members
//

//
// ShowToolTip
//

void CBar::ShowToolTip(DWORD dwCookie, BSTR bstrText, const CRect& rcBound, BOOL bTopMost)
{
	if (VARIANT_FALSE == bpV1.m_vbDisplayToolTips)
		return;

	if (NULL == m_hFontToolTip)
	{
		if (g_fSysWin95Shell)
		{
			NONCLIENTMETRICS nm;
			nm.cbSize = sizeof(NONCLIENTMETRICS);
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS,0,&nm,0);
			m_hFontToolTip = CreateFontIndirect(&nm.lfStatusFont);
		}
		else
		{
			LOGFONT lf;
			memset(&lf, 0, sizeof(LOGFONT));
#ifdef JAPBUILD
			MAKE_ANSIPTR_FROMWIDE(szStr,L"‚l‚r ‚oƒSƒVƒbƒN\0\0");
			lstrcpy(lf.lfFaceName, szStr);
#else
			lstrcpy(lf.lfFaceName, _T("MS Sans Serif"));
#endif
			lf.lfHeight = -8;
			m_hFontToolTip = CreateFontIndirect(&lf);
		}
	}
			
	if (NULL == m_pTip)
	{
		m_pTip = new CToolTip(this);
		assert(m_pTip);
		if (m_pTip)
		{
			m_pTip->CreateWin();
			m_pTip->SetFont(m_hFontToolTip);
		}
		else
			return;
	}
	else
		m_pTip->SetFont(m_hFontToolTip);

	//
	// If parent window is topmost make tooltip topmost
	//

	DWORD dwStyleEx = GetWindowLong(GetDockWindow(), GWL_EXSTYLE); 

	//
	// We can't make it topmost all the time since the tooltip will 
	// bring the window to the front
	//
	
	m_pTip->Show(dwCookie, bstrText, rcBound, bTopMost ? TRUE : (dwStyleEx & WS_EX_TOPMOST) != 0);
}

//
// HideToolTips
//

void CBar::HideToolTips(BOOL bResetTimer)
{
	if (NULL == m_pTip)
		return;

	m_pTip->Hide();
	if (bResetTimer)
		m_pTip->m_dwCloseTime -= 4000;
}
