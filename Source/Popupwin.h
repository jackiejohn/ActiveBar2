#ifndef __POPUPWIN_H__
#define __POPUPWIN_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "DropSource.h"

class CTool;
class CBand;
class CPopupWin;

//
// CPopupDragDrop
//

struct CPopupDragDrop : public BandDragDrop
{
	virtual CBand* GetBand(POINT pt);
};

//
// CPopupWinShadow
//

class CPopupWinShadow : public FWnd
{
public:
	CPopupWinShadow()
	{
		BOOL bResult = RegisterWindow(DD_WNDCLASS("xShadow"),
									  CS_SAVEBITS,
									  NULL,
									  ::LoadCursor(NULL, IDC_ARROW));
	}

	BOOL CreateWin(HWND hWndParent, const CRect& rcWindow);

	void Bar(CBar* pBar);

	void Show(const CRect& rcWindow, HDWP hDWP = NULL);
	void Hide();

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);

	SETLAYEREDWINDOWATTRIBUTES m_pSetLayeredWindowAttributes;
	CBar* m_pBar;
};

inline void CPopupWinShadow::Bar(CBar* pBar)
{
	m_pBar = pBar;
}

inline void CPopupWinShadow::Show(const CRect& rcWindow, HDWP hDWP)
{
	TRACE(5, "CPopupWinShadow::Show()\n");
	if (IsWindow())
	{
		if (hDWP)
		{
			DeferWindowPos(hDWP, m_hWnd, HWND_TOPMOST, rcWindow.left, rcWindow.top, rcWindow.Width(), rcWindow.Height(), SWP_NOACTIVATE);
			DeferWindowPos(hDWP, m_hWnd, HWND_TOPMOST, rcWindow.left, rcWindow.top, rcWindow.Width(), rcWindow.Height(), SWP_SHOWWINDOW | SWP_NOACTIVATE);
		}
		else
			SetWindowPos(HWND_TOPMOST, rcWindow.left, rcWindow.top, rcWindow.Width(), rcWindow.Height(), SWP_SHOWWINDOW | SWP_NOACTIVATE);
	}
}

inline void CPopupWinShadow::Hide()
{
	TRACE(5, "CPopupWinShadow::Hide()\n");
	if (IsWindow())
		ShowWindow(SW_HIDE);
}

//
// CPopupWinToolAndShadow
//

class CPopupWinToolAndShadow : public FWnd
{
public:
	CPopupWinToolAndShadow()
	{
		BOOL bResult = RegisterWindow(DD_WNDCLASS("xToolShadow"),
									  CS_SAVEBITS,
									  NULL,
									  ::LoadCursor(NULL, IDC_ARROW));
	}

	BOOL CreateWin(HWND hWndParent, const CRect& rcWindow);

	void Bar(CBar* pBar);

	void Show(CTool* pTool, CRect& rcWindow);
	void Hide(BOOL bSetNotPressed = FALSE);

	void DestroyWindow();

	HWND GetShadowWindow();

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);

	SETLAYEREDWINDOWATTRIBUTES m_pSetLayeredWindowAttributes;
	CPopupWinShadow m_theShadow;
	CTool* m_pTool;
	CBar* m_pBar;
	BOOL m_bVertical;
};

inline void CPopupWinToolAndShadow::DestroyWindow()
{
	FWnd::DestroyWindow();
	if (m_theShadow.IsWindow())
		m_theShadow.DestroyWindow();
}

inline void CPopupWinToolAndShadow::Bar(CBar* pBar)
{
	m_pBar = pBar;
	m_theShadow.Bar(m_pBar);
}

inline HWND CPopupWinToolAndShadow::GetShadowWindow()
{
	return m_theShadow.hWnd();
}

//
// CPopupWin
//

class CPopupWin : public FWnd
{
public:

	enum FlipDirection
	{
		eDirectionFlipVert = 0,
		eDirectionFlipHort = 1
	};

	enum eToolSelections
	{
		eRemoveSelection = -1,
		eDetachBandCaptionSelected = -2
	};

	enum Scroller
	{
		eBottom,
		eTop,
		ePopupScrollTimer = 300,
		eScrollTimerDuration = 100,
		eBitmapWidth = 9,
		eBitmapHeight = 5,
		eScrollerHeight = 9
	};

	enum Expand
	{
		eExpandTimer = 301,
		eExpandTimerDuration = 500
	};

	enum MiscConstants
	{
		PEA_DONTEXIT = 1,
		PEA_DETACH = 2,
		eOpenSubTimer = 200,
		eOpenSubTimerDuration = 300,
		eCloseSubTimer = 201,
		eCloseSubTimerDuration = 450,
		eDetachCaptionBorderThickness = 3,
		eDetachCaptionBorderAdjustment = 2,
		eDefaultLineHeight = 22
	};
	
	CPopupWin(CBand* pBand, BOOL bParentVertical); 
	~CPopupWin(); 

	BOOL PreCreateWin(CTool* pOwningTool, const CRect& rcBase, FlipDirection eDir, BOOL bMFU, BOOL bForceVFlip, int nHorzFlip = ddPopupMenuLeftAlign);
	BOOL CreateWin(HWND hWndParent);
	
	BOOL Flip();
	BOOL FlipVert();
	BOOL FlipHorz();
	BOOL MRUMenu();
	BOOL GetOptimalRect(CRect& rc);
	void RefreshTool(CTool* pTool);
	void SetCurSel(int nNewSel);
	void SetCurSelNormalBand(int nNewSel);

	BOOL CalcFlipingAndScrolling(CRect& rcPopup);

	void AdjustForScrolling(CRect& rc);

	BOOL m_bPopupMenu;

	void ReparentEditAndOrCombo();

	CRect& WindowRect();

	CPopupWinShadow Shadow();

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	BOOL TrackMenuExpandButton(CTool* pTool);
	void Draw(HDC hDC, CRect rcClient);

	void PaintMenuScroll(HDC hDC, const CRect &rcOut, CPopupWin::Scroller eTopBottom);
	
	void DrawXPBackgroundBorder(HDC hDCOff, CRect& rcClient);

	int GetLineHeight();

	BOOL Customization(); 
	BOOL HitTestDetach(POINT pt);
	BOOL DetachTrack(const CRect& rcBound);

	BOOL OnButton (UINT nMsg, BOOL bMouseUp, UINT nFlags);
	void OnKey(UINT key, BOOL* pbUsed);
	void OnMouseMove(POINT pt);
	
	void FireDelayedClick(CTool* pTool);
	void OpenSubPopup();
	void PaintDetachBar(HDC hDC);
	void PopLayout();
	void PushLayout();

	CBandDropTarget* m_pPopupDropTarget;
	CPopupDragDrop  m_thePopupDragDrop;

	CBand* m_pBand;
	DWORD  m_dwWaitTime;
	CRect  m_rcBase;
	CRect  m_rcPaintCache;
	CRect  m_rcWindow;
	CRect  m_rcTopScroller;
	CRect  m_rcBottomScroller;
	int    m_nPhysicalScrollPos;
	int    m_nScrollPos;
	int    m_nCurSel;
	int    m_nCurPopup;
	short  m_nSubMenuOpenDelay;
	short  m_nSubMenuCloseDelay;
#pragma pack(push)
#pragma pack (1)
	BOOL   m_bBottomScrollActive:1;
	BOOL   m_bTopScrollActive:1;
	BOOL   m_bFirstTimePaint:1;
	BOOL   m_bTimerStarted:1;
	BOOL   m_bScrollActive:1;
	BOOL   m_bFlipVert:1;
	BOOL   m_bFlipHorz:1;
	BOOL   m_bMFU:1;
	BOOL   m_bTrackTimerStarted:1;
	BOOL   m_bParentVertical:1;
	int    m_nHorzFlip;
	int    m_nParentIndex;
#pragma pack(pop)
	FlipDirection m_nfdDirection;
	CRect m_rcClient;
	CRect m_rcClientOrginal;
	CFlickerFree* m_pFF;
	CTool* m_pOwningTool;
	CPopupWinShadow m_theShadow;
	long m_nOwningToolIndex;

	friend CPopupDragDrop;
	friend CBand;
};

inline CPopupWinShadow CPopupWin::Shadow()
{
	return m_theShadow;
}

inline BOOL CPopupWin::MRUMenu()
{
	return m_bMFU;
}

inline BOOL CPopupWin::Flip()
{
	return (m_bFlipHorz || m_bFlipVert);
}

inline void CPopupWin::AdjustForScrolling(CRect& rc)
{
	if (m_bScrollActive)
		rc.Offset(0, -m_nPhysicalScrollPos);
}

inline BOOL CPopupWin::FlipVert()
{
	return m_bFlipVert;
}

inline BOOL CPopupWin::FlipHorz()
{
	return m_bFlipHorz;
}

inline CRect& CPopupWin::WindowRect()
{
	return m_rcWindow;
}

#endif