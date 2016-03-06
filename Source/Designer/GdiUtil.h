#ifndef GDIUTIL_INCLUDED
#define GDIUTIL_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

//
// CFlickerFree
//

class CFlickerFree
{
public:
	CFlickerFree();
	~CFlickerFree();

	HDC GetDC();
	HBITMAP GetBitmap();
	HDC RequestDC(HDC hDC, int nWidth, int nHeight);
	void Paint(HDC hdcDest, int x, int y);
	void Paint(HDC hdcDest, int x, int y, int w, int h);
	void Paint(HDC hdcDest, int x, int y, int width, int nHeight, int xs, int ys);

protected:
	HDC m_hMemDC;
	HBITMAP m_hBitmap;
	HBITMAP m_hBitmapOld;
	int m_nWidth;
	int m_nHeight;
	int m_nRequestWidth;
	int m_nRequestHeight;
};

inline HBITMAP CFlickerFree::GetBitmap()
{
	return m_hBitmap;
}

inline HDC CFlickerFree::GetDC()
{
	return m_hMemDC;
}

inline void CFlickerFree::Paint(HDC hDCDest,int x, int y, int nWidth, int nHeight, int xs, int ys)
{
	BitBlt(hDCDest,x, y, nWidth, nHeight, m_hMemDC, xs, ys, SRCCOPY);
}

inline void CFlickerFree::Paint(HDC hDCDest, int x, int y)
{
	BitBlt(hDCDest, x, y, m_nRequestWidth, m_nRequestHeight, m_hMemDC, 0, 0, SRCCOPY);
}

inline void CFlickerFree::Paint(HDC hDCDest, int x, int y, int w, int h)
{
	StretchBlt(hDCDest, x, y, w, h, m_hMemDC, 0, 0, m_nRequestWidth, m_nRequestHeight, SRCCOPY);
}

//
// UIUtilities
//

class UIUtilities
{
public:
	UIUtilities();
	~UIUtilities();

	static UIUtilities& GetUIUtilities();
	static void CleanupUIUtilities();

	static void FastRect(HDC hDC, CRect& rc);

	static HBITMAP CreateDitherBitmap(BOOL bMonochrome);

	static HBRUSH CreateHalftoneBrush(BOOL bColor);

	static void Draw3DRect(HDC      hDC, 
						   CRect&   rc, 
						   COLORREF crTopLeft, 
						   COLORREF crBottomRight);

	static void FillSolidRect(HDC      hDC, 
							  CRect&   rc, 
							  COLORREF cr);

	static void FillSolidRect(HDC      hDC, 
							  int      x, 
							  int      y, 
							  int      cx, 
							  int      cy, 
							  COLORREF cr);

	void InitializeToolTips();

	void UnInitializeToolTips();

	void HideToolTips(BOOL bResetTimer);

	void ShowToolTip(void*		  dwCookie, 
					 BSTR		  bstrText, 
					 const CRect& rcBound, 
					 const POINT& pt);
};

inline void UIUtilities::FastRect(HDC hDC, CRect& rc) 
{
	ExtTextOut(hDC, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);
}

inline void UIUtilities::FillSolidRect(HDC hDC, CRect& rc, COLORREF crColor)
{
	if (GetGlobals().m_nBitDepth > 8)
	{
		SetBkColor(hDC, crColor);
		ExtTextOut(hDC, 
				   0, 
				   0, 
				   ETO_OPAQUE, 
				   &rc, 
				   NULL, 
				   0, 
				   NULL);
	}
	else
	{ 
		HBRUSH hBrush = CreateSolidBrush(crColor);
		FillRect(hDC, &rc, hBrush);
		DeleteBrush(hBrush);
	}
}

inline void UIUtilities::FillSolidRect(HDC      hDC, 
									   int      x, 
									   int      y, 
									   int      cx, 
									   int      cy, 
									   COLORREF crColor)
{
	CRect rc(x, x + cx, y, y + cy);
	if (GetGlobals().m_nBitDepth > 8)
	{
		SetBkColor(hDC, crColor);
		ExtTextOut(hDC, 
				   0, 
				   0, 
				   ETO_OPAQUE, 
				   &rc, 
				   NULL, 
				   0, 
				   NULL);
	}
	else
	{ 
		HBRUSH hBrush = CreateSolidBrush(crColor);
		FillRect(hDC, &rc, hBrush);
		DeleteBrush(hBrush);
	}
}

inline void UIUtilities::Draw3DRect(HDC      hDC, 
									CRect&   rc, 
									COLORREF crTopLeft, 
									COLORREF crBottomRight)
{	   
	FillSolidRect(hDC, rc.left, rc.top, rc.Width() - 1, 1, crTopLeft);
	FillSolidRect(hDC, rc.left, rc.top, 1, rc.Width() - 1, crTopLeft);
	FillSolidRect(hDC, rc.left + rc.Width(), rc.top, -1, rc.Height(), crBottomRight);
	FillSolidRect(hDC, rc.left, rc.top + rc.Height(), rc.Width(), -1, crBottomRight);
}

//
// DDToolBar
//

class DDToolBar : public FWnd
{
public:
	DDToolBar();
	~DDToolBar();

	enum Constants
	{
		eRTToolbar = 241,
		eTimer = 555,
		eVerticalGap = 3,
		eSeparatorGap = 8,
		eMaxToolBitmapSize = 64,
		eToolDefaultBorderCX = 3,
		eToolDefaultBorderCY = 3,
		eToolDefaultBitmapCX = 16,
		eToolDefaultBitmapCY = 15,
	};

	enum StatesEx
	{
		TBSTATE_FLYOUT = 0x80
	};

    enum StyleEx
	{
		TBSTYLE_PLACEHOLDER = 0x10,
		TBSTYLE_KEEPTOGETHER = 0x20
	};

	enum ToolbarStyle
	{
		eNormal = 0x01,
		eIE     = 0x02,
		eSunken = 0x04,
		eFlat   = 0x08,
		eRaised = 0x10
	};

	struct Tool
	{
		Tool();

		TBBUTTON   biButton;
		LPTSTR	   szWindowClass;
		HWND       hWnd;
		DWORD	   dwStyle;
		DWORD	   dwStyleEx;
		CRect	   rcButton;
		SIZE	   sizeButton;
		int		   nMin;
		int        nMax;
	};

	HWND Create(HWND hWndParent, const CRect& rc, int nId);
	
	BOOL AddButton(TBBUTTON& biButton);
	BOOL InsertButton(int nIndex, TBBUTTON& biButton);
	BOOL DeleteButton(int nIndex);

	BOOL SetState(int nId, BYTE fsState);
	BOOL GetState(int nId, BYTE& fsState);

	int CommandIdToIndex(int nId);
	
	void InvalidateButton(Tool& theButton);
	void InvalidateButton(int nCmdId);

	virtual void ResetContent() {};

	void EnableButton(int nId, BOOL bFlag);
	void CheckButton(int nId, BOOL bFlag);

	void SetButtonSize(int cx, int cy);
	void SetBorder(int cx, int cy);
	void SetStyle(int sToolbar);

	BOOL GetTool(int nId, Tool*& pTool);

	BOOL LoadToolBar(LPCTSTR szResourceName);
	BOOL LoadBitmap(LPCTSTR szResourceName);
	BOOL SetButtons(const UINT* pnIdArray, int nIDCount);
	void RecalcLayout();

	HIMAGELIST& GetImageList();

protected:
	virtual void OnDoubleClick(POINT) {};

	void PaintToolButton(HDC   hDC, 
						 Tool& theTool, 
						 int   nX, 
						 int   nY, 
						 int   nDX, 
						 int   nDY, 
						 BYTE  fsState);

	void GetHorzFixedOptimalRect(CRect& rc);
	void GetVertFixedOptimalRect(CRect& rc);

	int HitTest(const POINT& pt);
	void GetButtonRect(Tool& aTool, CRect& rcButton);
	
	virtual LRESULT WindowProc(UINT nMessage, WPARAM wParam, LPARAM lParam);
	LRESULT Draw(HDC hDC, CRect& rcBound);
	LRESULT OnTimer(WPARAM wParam, LPARAM lParam);
	LRESULT OnMouseMove(WPARAM wParam, LPARAM lParam);
	LRESULT OnLButtonDown(WPARAM wParam, LPARAM lParam);
	virtual LRESULT OnRButtonDown(WPARAM wParam, LPARAM lParam) {return FALSE;};
	virtual LRESULT OnWindowPosChanged(WPARAM wParam, LPARAM lParam);

	BOOL Enabled(BYTE fsState);
	BOOL Disabled(BYTE fsState);
	BOOL Pressed(BYTE fsState);
	BOOL Checked(BYTE fsState);
	BOOL Indeterminate(BYTE fsState);

	CFlickerFree m_dbDraw;
	HIMAGELIST   m_ilImages;
	int			 m_nToolbar;

	DWORD m_dwLastClickTime;
	POINT m_ptLastClick;
	Tool* m_pButtons;
	SIZE  m_sizeButton;
	SIZE  m_sizeImage;
	LONG  m_nVerticalGap;
	LONG  m_nHorzBorder;
	LONG  m_nVertBorder;
	int   m_nScrollAmount; 
	int   m_nButtonCount;

	static DDToolBar* m_pActiveToolBar;
	static int m_nActiveButton;
};

inline DDToolBar::Tool::Tool()
{
	memset(&biButton, '\0', sizeof(TBBUTTON));
	hWnd = NULL;
	szWindowClass = NULL;
	dwStyle = dwStyleEx = nMax = nMin = sizeButton.cx = sizeButton.cy = 0;
};

inline void DDToolBar::SetStyle(int nToolbar)
{
	m_nToolbar = nToolbar;
}

inline void DDToolBar::SetButtonSize(int cx, int cy) 
{
	m_sizeButton.cx = cx;
	m_sizeButton.cy = cy;
};

inline void DDToolBar::SetBorder(int cx, int cy) 
{
	m_nHorzBorder = cx;
	m_nVertBorder = cy;
};

inline BOOL DDToolBar::Enabled(BYTE fsState)
{
	return (fsState & TBSTATE_ENABLED);
}

inline BOOL DDToolBar::Disabled(BYTE fsState)
{
	return !(fsState&TBSTATE_ENABLED);
}

inline BOOL DDToolBar::Pressed(BYTE fsState)
{
	 return (fsState & TBSTATE_PRESSED);
}

inline BOOL DDToolBar::Checked(BYTE fsState)
{
	 return (fsState & TBSTATE_CHECKED);
}

inline BOOL DDToolBar::Indeterminate(BYTE fsState) 
{
	return (fsState & TBSTATE_INDETERMINATE);
}

inline BOOL DDToolBar::GetTool(int nId, Tool*& pTool)
{
	for (int nIndex = 0; nIndex < m_nButtonCount; nIndex++)
	{
		if (nId == m_pButtons[nIndex].biButton.idCommand)
		{
			pTool = &m_pButtons[nIndex];
			return TRUE;
		}
	}
	return FALSE;
}

inline HIMAGELIST& DDToolBar::GetImageList()
{
	return m_ilImages;
}

//
// CSplitter
//

class CSplitter
{
	// Construction
public:
	CSplitter();

	void hWnd(HWND hWnd);

	void Tracker(const int& nTop,
				 const int& nLeft,
				 const int& nBottom,
				 const int& nRight);

	void Tracker(const CRect& rcTracker);

	int TrackerLeft();
	int TrackerRight();

	void SetLeftLimit(const int& nLeft);
	void SetRightLimit(const int& nRight);
	CRect& Tracker();

	virtual void OnInvertTracker();
	virtual void SetSplitCursor(const BOOL& bTracking);
	virtual BOOL SplitterHitTest(const POINT& pt);

	virtual void StartTracking();
	void DoTracking(const POINT& pt);
	virtual BOOL StopTracking(const POINT& pt);

// Implementation
private:
	int		m_nRight;
	int		m_nLeft;
	BOOL    m_bTracking;
	CRect   m_rcTracker;
	HCURSOR m_hPrimaryCursor;
	HCURSOR m_hSplitVCursor;
	HWND    m_hWnd;
};

inline void CSplitter::Tracker(const CRect& rcTracker)
{
	m_rcTracker = rcTracker;
}

inline CRect& CSplitter::Tracker()
{
	return m_rcTracker;
}

inline void CSplitter::hWnd(HWND hWnd)
{
	m_hWnd = hWnd;
}

inline void CSplitter::Tracker(const int& nLeft,
							   const int& nTop,
							   const int& nRight,
							   const int& nBottom)
{
	m_rcTracker.Set(nLeft, nRight, nTop, nBottom);
}

inline int CSplitter::TrackerLeft()
{
	return m_rcTracker.left;
}

inline int CSplitter::TrackerRight()
{
	return m_rcTracker.right;
}

inline void CSplitter::SetRightLimit(const int& nRight)
{
	m_nRight = nRight;
}

inline void CSplitter::SetLeftLimit(const int& nLeft)
{
	m_nLeft = nLeft;
}

#endif