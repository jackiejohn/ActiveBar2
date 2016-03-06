#ifndef __MINIWIN_H__
#define __MINIWIN_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "FWnd.h"
#include "DropSource.h"
class CMiniWin;
class CDockMgr;
class CBand;

//
// CMiniDragDrop
//

struct CMiniDragDrop : public BandDragDrop
{
	virtual CBand* GetBand(POINT pt);
};

//
// CMiniWin
//

class CMiniWin : public FWnd
{
public:
	enum 
	{
		eMiniWinCaptionHeight = 18,
		eFloatResize = 350
	};

	CMiniWin(CBand* pBand);
	virtual ~CMiniWin();

	BOOL CreateWin(HWND hWndParent);

	void ActivateTopParent();
	void Show(BOOL bActive);

	void DragWindow(const CRect& rc);
	
	static void DrawDragRect(const CRect& rc, 
							 int          nThickness = 1,
							 CRect*       prcLast = NULL,
							 int          nLastThickness = 1);

	static void AdjustWindowRectEx(CRect& rc, BOOL bMenu = FALSE);

	void MoveTo(int x, int y);
	
	void BandClosing(const BOOL& bClosing);
	BOOL BandClosing();

	BOOL FreeResize();

	// In Customization Mode?
	BOOL InCustomization(); 

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	BOOL OnNcLButtonDown(UINT nHitTest, POINT pt);
	void SizeWindow(UINT nHit);

	CBandDropTarget  m_theMiniDropTarget;
	CMiniDragDrop    m_theBandDragDrop;
	CDockMgr*	     m_pDockMgr;
	CBand*		     m_pBand;
	CRect			 m_rcClose;
	CBar*	         m_pBar;
	BOOL		     m_bBandClosing;
	BOOL			 m_bActive;
	BOOL			 m_bHover;
	int				 m_nPenThickness;
	static DWORD m_dwStyle;
	static DWORD m_dwStyleEx;
	friend CMiniDragDrop;
};

//
// MoveTo
//

inline void CMiniWin::MoveTo(int x, int y)
{
	SetWindowPos(NULL, 
				 x, 
				 y, 
				 0, 
				 0,
				 SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
}

//
// BandClosing
//

inline void CMiniWin::BandClosing(const BOOL& bClosing)
{
	m_bBandClosing = bClosing;
}

inline BOOL CMiniWin::BandClosing()
{
	return m_bBandClosing;
}

//
// AdjustWindowRectEx
//

inline void CMiniWin::AdjustWindowRectEx(CRect& rc, BOOL bMenu)
{
	if (!g_fSysWin95Shell)
	{
		::AdjustWindowRectEx(&rc, 
							 WS_THICKFRAME|WS_POPUP|WS_SYSMENU, 
							 bMenu, 
							 0);
		rc.top -= eMiniWinCaptionHeight;
	}
	else
	{
		::AdjustWindowRectEx(&rc, 
							 WS_CAPTION|WS_THICKFRAME|WS_POPUP|WS_SYSMENU, 
							 bMenu, 
							 WS_EX_TOOLWINDOW);
	}
}

//
// Forward Reference for utility functions
//
// DrawDragRect2
//

void DrawDragRect2(HDC          hDC,
				   const CRect& rc, 
				   CRect*       prcLast, 
				   HBRUSH       hBrush, 
				   HBRUSH       hBrushLast);

//
// CalcSizeRect
//

void CalcSizeRect(UINT nHit, CRect& rc, int nDeltaX, int nDeltaY);

#endif