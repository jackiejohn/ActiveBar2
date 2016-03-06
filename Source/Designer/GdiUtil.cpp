//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Globals.h"
#include "Resource.h"
#include "debug.h"
#include "..\EventLog.h"
#include "Support.h"

extern HINSTANCE g_hInstance;

//
// CFlickerFree
//

CFlickerFree::CFlickerFree()
	: m_hMemDC(NULL),
	  m_hBitmap(NULL),
	  m_hBitmapOld(NULL)
{
	m_nRequestHeight = m_nRequestWidth = m_nHeight = m_nWidth = 0;
}

CFlickerFree::~CFlickerFree()
{
	BOOL bResult;
	if (m_hBitmap)
	{
		SelectBitmap(m_hMemDC, m_hBitmapOld);
		bResult = DeleteBitmap(m_hBitmap);
		assert(bResult);
	}
	if (m_hMemDC)
	{
		bResult = DeleteDC(m_hMemDC);
		assert(bResult);
	}
}

HDC CFlickerFree::RequestDC(HDC hDC, int nWidth, int nHeight)
{
	if (NULL == m_hMemDC)
	{
		m_hMemDC = CreateCompatibleDC(hDC);
		if (NULL == m_hMemDC)
			return NULL;
	}

	if (nWidth > m_nWidth || nHeight > m_nHeight)
	{
		// old bitmap is not large enough create new one
 		if (m_hBitmap)
		{
			SelectBitmap(m_hMemDC, m_hBitmapOld);
			DeleteBitmap(m_hBitmap);
		}
		m_hBitmap = CreateCompatibleBitmap(hDC, nWidth, nHeight);
		if (NULL == m_hBitmap)
			return NULL;

		m_nWidth = nWidth;
		m_nHeight = nHeight;
		m_hBitmapOld = SelectBitmap(m_hMemDC, m_hBitmap);
		PatBlt(m_hMemDC, 0, 0, nWidth, nHeight, BLACKNESS);
	}
	m_nRequestWidth = nWidth;
	m_nRequestHeight = nHeight;
	return m_hMemDC;
}

//
// DDToolTip
//

class DDToolTip : public FWnd
{
public:
	DDToolTip();
	~DDToolTip();

	BOOL RegisterClass();
	void Hide();
	void ResetTimer();

	enum Constants
	{
		eIdTimer = 400,
		eTimerDuration = 600,
		eCheckTimerDuration = 200
	};

	enum Mode
	{
		eIdle,
		eTimerStarted,
		eTimerCompleted
	};

	void Show(void* pCookie, BSTR bstrText, const CRect& rcBound, const POINT& pt);

private:
	virtual LRESULT WindowProc(UINT mMessage, WPARAM wParam, LPARAM lParam);
	LRESULT OnTimer (WPARAM wParam, LPARAM lParam);
	LRESULT Draw(HDC hDC, CRect& rcBound);

	void Open();

	static UINT  WM_TOOLTIPSUPPORT;
	static HFONT m_hFontToolTip;
	static Mode  m_eMode;

	void* m_pPrevCookie;
	POINT m_pt;
	CRect m_rcBound;
	BSTR  m_bstrText;
};

UINT DDToolTip::WM_TOOLTIPSUPPORT = RegisterWindowMessage(_T("ToolTipSupport"));
HFONT DDToolTip::m_hFontToolTip = 0;
DDToolTip::Mode  DDToolTip::m_eMode = eIdle;
static DDToolTip* s_pTip = NULL;

DDToolTip::DDToolTip()
	: m_bstrText(NULL),
	  m_pPrevCookie(0)
{
	if (NULL == m_hFontToolTip)
	{
		NONCLIENTMETRICS nm;
		nm.cbSize=sizeof(NONCLIENTMETRICS);
		SystemParametersInfo(SPI_GETNONCLIENTMETRICS,0,&nm,0);
		m_hFontToolTip = CreateFontIndirect(&nm.lfStatusFont);
		assert(m_hFontToolTip);
	}
			
	RegisterClass();

	if (NULL == s_pTip)
	{
		s_pTip = this;
		CreateEx(WS_EX_TOOLWINDOW|WS_EX_TOPMOST,
			     _T(""),
				 WS_POPUP|WS_BORDER,
				 0,
				 0,
				 0,	
				 0,
   				 0,
				 0,
				 0);
		SetFont(m_hFontToolTip);
	}
}

DDToolTip::~DDToolTip()
{
	SysFreeString(m_bstrText);
	if (m_hFontToolTip)
	{
		BOOL bResult = DeleteFont(m_hFontToolTip);
		assert(bResult);
		m_hFontToolTip = NULL;
	}
	if (IsWindow(m_hWnd))
	{
		KillTimer(m_hWnd, eIdTimer);
		DestroyWindow();
	}
}

BOOL DDToolTip::RegisterClass()
{
	return RegisterWindow(DD_WNDCLASS("ToolTipWnd"),
						  CS_VREDRAW|CS_HREDRAW,
						  0,
						  LoadCursor(0, IDC_ARROW),
						  NULL,
						  NULL);
}

LRESULT DDToolTip::WindowProc(UINT nMessage, WPARAM wParam, LPARAM lParam)
{
	switch (nMessage)
	{
	case WM_DESTROY:
		TRACE(5, _T("WM_DESTROY\n"));
		break;

	case WM_TIMER:
		return OnTimer(wParam, lParam);

	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(m_hWnd, &ps);
			if (hDC)
			{
				CRect rcBound;
				GetClientRect(rcBound);
				Draw(hDC, rcBound);
				EndPaint(m_hWnd, &ps);
				return 0;
			}
		}
		break;
	}
	return FWnd::WindowProc(nMessage, wParam, lParam);
}

void DDToolTip::Open()
{
	SIZE sizeText;
	HDC hDC = GetDC(m_hWnd);
	if (hDC)
	{
		HFONT hFontOld = (HFONT)SelectObject(hDC, m_hFontToolTip);
		MAKE_TCHARPTR_FROMWIDE(szText, m_bstrText);
		GetTextExtentPoint32(hDC, szText, strlen(szText), &sizeText);
		SelectFont(hDC, hFontOld);
		ReleaseDC(m_hWnd, hDC);
	}
	
	int nTemp = (m_rcBound.right + m_rcBound.left)/2;
	CRect rc(nTemp,
			 nTemp+sizeText.cx + 8,
			 m_pt.y + 20,
			 m_pt.y + sizeText.cy + 24);

	SetWindowPos(HWND_TOP,
				 rc.left,
				 rc.top,
				 rc.Width(),
				 rc.Height(),
				 SWP_NOACTIVATE|SWP_SHOWWINDOW);
	
	InvalidateRect(0, TRUE);
	KillTimer(m_hWnd, eIdTimer);
	SetTimer(m_hWnd, eIdTimer, eCheckTimerDuration, NULL);
}

void DDToolTip::Hide()
{
	SetWindowPos(NULL,
				 0,
				 0,
				 0,
				 0,
				 SWP_NOSIZE|SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE|SWP_HIDEWINDOW);
	m_pPrevCookie = NULL;
}

void DDToolTip::Show(void* pCookie, BSTR bstrText, const CRect& rcBound, const POINT& pt)
{
	if (pCookie == m_pPrevCookie)
	{
		if (eTimerStarted == m_eMode)
		{
			KillTimer(m_hWnd, eIdTimer);
			SetTimer(m_hWnd, eIdTimer, eTimerDuration, NULL);
		}
		return;
	}
	else
	{
		if (eTimerStarted == m_eMode)
		{
			KillTimer(m_hWnd, eIdTimer);
			m_eMode=eIdle;
		}
	}
	SysFreeString(m_bstrText);
	m_bstrText = NULL;
	if (NULL == bstrText || NULL == *bstrText)
	{
		Hide();
		return;
	}

	m_pt = pt;
	m_rcBound = rcBound;
	m_pPrevCookie = pCookie;
	m_bstrText = SysAllocString(bstrText);

	UINT nTimer;
	if (eIdle == m_eMode)
	{
		m_eMode = eTimerStarted;
		nTimer = SetTimer(m_hWnd, eIdTimer, eTimerDuration, 0);
	}

	if (eTimerCompleted == m_eMode)
		Open();
}

void DDToolTip::ResetTimer()
{
	KillTimer(m_hWnd, eIdTimer);
	m_eMode = eIdle;
}

LRESULT DDToolTip::OnTimer (WPARAM wParam, LPARAM lParam)
{
	POINT pt;
	GetCursorPos(&pt);
	if (eTimerStarted == m_eMode)
	{
		if (PtInRect(&m_rcBound, pt))
		{
			m_eMode = eTimerCompleted;
			Open();
		}
		
	}
	else if (eTimerCompleted == m_eMode)
	{
		// check if still in boundrect
		if (!PtInRect(&m_rcBound, pt))
		{
			Hide();
			m_eMode = eIdle;
		}
	}
	return 0;
}

LRESULT DDToolTip::Draw(HDC hDC, CRect& rcBound)
{
	SetBkColor(hDC, GetSysColor(COLOR_INFOBK));
	ExtTextOut(hDC, 0, 0, ETO_OPAQUE, &rcBound, 0, 0, 0);
	HFONT hFontOld = (HFONT)SelectObject(hDC, m_hFontToolTip);
	SetBkMode(hDC, TRANSPARENT);
	SetTextColor(hDC, GetSysColor(COLOR_INFOTEXT));
	MAKE_TCHARPTR_FROMWIDE(szText, m_bstrText);
	TextOut(hDC, 2, 1, szText, lstrlen(szText));
	SelectFont(hDC, hFontOld);
	return 0;
}

//
// UIUtilities
//

static UIUtilities* s_pUIUtilities = NULL;
static HBRUSH       s_hbrDither = NULL;
static HDC          s_hDCMono = NULL;

UIUtilities::UIUtilities()
{
	if (NULL == s_pTip)
	{
		s_pTip = new DDToolTip;
		assert(s_pTip);
		if (NULL == s_pTip)
			return;
	}
	// Mono DC and Bitmap for disabled image
	HBITMAP hbmGray = UIUtilities::CreateDitherBitmap(FALSE);
	if (hbmGray)
	{
		s_hbrDither = ::CreatePatternBrush(hbmGray);
		BOOL bResult = DeleteBrush(hbmGray);
		assert(bResult);
	}
	s_hDCMono = CreateCompatibleDC(NULL);
	assert(s_hDCMono);
}

UIUtilities::~UIUtilities()
{
	delete s_pTip;
	s_pTip = NULL;
}

UIUtilities& UIUtilities::GetUIUtilities()
{
	if (NULL == s_pUIUtilities)
	{
		s_pUIUtilities = new UIUtilities;
		assert(s_pUIUtilities);
		if (NULL == s_pUIUtilities)
			throw -1;
		return *s_pUIUtilities;
	}
	else
		return *s_pUIUtilities;
}

void UIUtilities::CleanupUIUtilities()
{
	BOOL bResult;
	if (s_pUIUtilities)
	{
		delete s_pUIUtilities;
		s_pUIUtilities = NULL;
	}
	if (s_hbrDither)
	{
		bResult = DeleteBrush(s_hbrDither);
		assert(bResult);
		s_hbrDither = NULL;
	}

	if (s_hDCMono)
	{
		bResult = DeleteDC(s_hDCMono);
		assert(bResult);
		s_hDCMono = NULL;
	}
}

HBITMAP UIUtilities::CreateDitherBitmap(BOOL bMonochrome)
{
	HBITMAP hbm;
	if (bMonochrome)
	{
		WORD wPatGray[8];
		for (int nIndex = 0; nIndex < 8; nIndex++)
			wPatGray[nIndex] = (nIndex & 1) ? 0x55 : 0xAA;
		hbm = CreateBitmap(8, 8, 1, 1, wPatGray);
	}
	else 
	{
		struct  
		{
			BITMAPINFOHEADER bmiHeader;
			RGBQUAD          bmiColors[16];
		} bmi;

		long nPatGray[8];
		for (int nIndex = 0; nIndex < 8; nIndex++)
	   		nPatGray[nIndex] = (nIndex & 1) ? 0xAAAA5555L : 0x5555AAAAL;
	
		memset(&bmi, 0, sizeof(bmi));
		bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bmi.bmiHeader.biWidth = 8;
		bmi.bmiHeader.biHeight = 8;
		bmi.bmiHeader.biPlanes = 1;
		bmi.bmiHeader.biBitCount = 1;
		bmi.bmiHeader.biCompression = BI_RGB;
		COLORREF clr = ::GetSysColor(COLOR_BTNFACE);
		bmi.bmiColors[0].rgbBlue = GetBValue(clr);
		bmi.bmiColors[0].rgbGreen = GetGValue(clr);
		bmi.bmiColors[0].rgbRed = GetRValue(clr);
		clr = ::GetSysColor(COLOR_BTNHIGHLIGHT);
		bmi.bmiColors[1].rgbBlue = GetBValue(clr);
		bmi.bmiColors[1].rgbGreen = GetGValue(clr);
		bmi.bmiColors[1].rgbRed = GetRValue(clr);
		HDC hDC = GetDC(NULL);
		if (hDC)
		{
			hbm = CreateDIBitmap(hDC, 
								 &bmi.bmiHeader, 
								 CBM_INIT,
								 (LPBYTE)nPatGray, 
								 (LPBITMAPINFO)&bmi, 
								 DIB_RGB_COLORS);
			ReleaseDC(0, hDC);
		}
	}
	return hbm;
}

HBRUSH UIUtilities::CreateHalftoneBrush(BOOL bColor)
{
	HBRUSH hbrDither;
	HBITMAP hbmGray = CreateDitherBitmap(bColor);
	if (hbmGray)
	{
		// Mono DC and Bitmap for disabled image
		hbrDither = ::CreatePatternBrush(hbmGray);
		BOOL bResult = DeleteBrush(hbmGray);
		assert(bResult);
	}
	return hbrDither;
}

void UIUtilities::HideToolTips(BOOL bResetTimer)
{
	if (NULL == s_pTip)
	{
		s_pTip->ResetTimer();
		s_pTip->Hide();
	}
}

void UIUtilities::ShowToolTip(void*		   pCookie, 
							  BSTR		   bstrText, 
							  const CRect& rcBound, 
							  const POINT& pt)
{
	if (s_pTip)
		s_pTip->Show(pCookie, bstrText, rcBound, pt);
}

static const DWORD s_ROP_DSPDxax = 0x00E20746L;
static const DWORD s_ROP_PSDPxax = 0x00B8074AL;

enum ToolOpcode 
{
	TOOLOP_QUERYDROPSUPPORT,
	TOOLOP_DROP,
	TOOLOP_UPDATEDOCKINGFRAME,
	TOOLOP_QUERYDELETEONDROP,
	TOOLOP_FULLDRAGDROP
};

//
// DDToolBar
//

DDToolBar* DDToolBar::m_pActiveToolBar = NULL;
int        DDToolBar::m_nActiveButton  = -1;

DDToolBar::DDToolBar()
	:	m_ilImages(NULL),
	    m_pButtons(NULL)
{
	RegisterWindow(DD_WNDCLASS("ToolBar"),
				   CS_VREDRAW | CS_HREDRAW,
				   NULL,
				   LoadCursor(0, IDC_ARROW),
				   NULL,
				   NULL);

	m_nHorzBorder = eToolDefaultBorderCX;
	m_nVertBorder = eToolDefaultBorderCY;
	m_nVerticalGap = eVerticalGap;
	m_nButtonCount = 0;
	m_dwLastClickTime = 0;
	m_sizeButton.cx = eToolDefaultBitmapCX + eToolDefaultBorderCX * 2 + 1;
	m_sizeButton.cy = eToolDefaultBitmapCY + eToolDefaultBorderCY * 2 + 2;
	m_nScrollAmount = 0;
	m_nToolbar = eIE;
	UIUtilities::GetUIUtilities();
}

DDToolBar::~DDToolBar()
{
	delete [] m_pButtons;
	if (m_pActiveToolBar == this)
		m_pActiveToolBar = NULL;
	ImageList_Destroy(m_ilImages);
}

HWND DDToolBar::Create(HWND hWndParent, const CRect& rc, int nId)
{
	RecalcLayout();
	FWnd::Create(_T("Tools"),
				 WS_CLIPCHILDREN | WS_CHILD | WS_VISIBLE | WS_OVERLAPPED,
				 rc.left,
				 rc.top,
				 rc.Width(),	
				 rc.Height(),
   				 hWndParent,
				 (HMENU)nId,
				 NULL);
	
	HFONT hFont = GetFont();
	for (int nIndex = 0; nIndex < m_nButtonCount; nIndex++)
	{
		if (m_pButtons[nIndex].biButton.fsStyle&TBSTYLE_PLACEHOLDER)
		{
			m_pButtons[nIndex].hWnd = CreateWindowEx(m_pButtons[nIndex].dwStyleEx,
													 m_pButtons[nIndex].szWindowClass,
												     _T(""),
													 WS_VISIBLE|WS_CHILD|m_pButtons[nIndex].dwStyle,
												     0,
												     0,
									 			     0,
												     0,
												     m_hWnd,
												     (HMENU)m_pButtons[nIndex].biButton.idCommand,
												     g_hInstance,
												     NULL);
			if (m_pButtons[nIndex].szWindowClass &&
				_tcscmp (m_pButtons[nIndex].szWindowClass, _T("STATIC")) == 0)
			{
				DDString strTemp;
				strTemp.LoadString(m_pButtons[nIndex].biButton.idCommand);
				::SetWindowText(m_pButtons[nIndex].hWnd, strTemp);
			}
		}
	}
	return m_hWnd;
}

LRESULT DDToolBar::WindowProc(UINT nMessage, WPARAM wParam, LPARAM lParam)
{
	switch (nMessage)
	{
	case WM_VSCROLL:
	case WM_HSCROLL:
		return ::SendMessage(GetParent(m_hWnd), nMessage, wParam, lParam);

	case WM_PAINT: 
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(m_hWnd, &ps);
			if (hDC)
			{
				CRect rcBound;
				GetClientRect(rcBound);
				Draw(hDC, rcBound);
				EndPaint(m_hWnd, &ps);
				return 0;
			}
		}
		break;

	case WM_SIZE: 
		RecalcLayout();
		break;

	case WM_SETFONT: 
		{
			HWND hWndChild = GetWindow(m_hWnd, GW_CHILD);
			while (hWndChild)
			{
				::SendMessage(hWndChild, nMessage, wParam, lParam);
				hWndChild = GetWindow(hWndChild, GW_HWNDNEXT);
			}
		}
		break;

	case WM_COMMAND: 
		return ::SendMessage(GetParent(m_hWnd), nMessage, wParam, lParam);

	case WM_NOTIFY: 
		return ::SendMessage(GetParent(m_hWnd), nMessage, wParam, lParam);

	case WM_TIMER: 
		return OnTimer(wParam, lParam);

	case WM_MOUSEMOVE: 
		return OnMouseMove(wParam, lParam);

	case WM_LBUTTONDOWN: 
		return OnLButtonDown(wParam, lParam);

	case WM_RBUTTONDOWN: 
		return OnRButtonDown(wParam, lParam);

	case WM_WINDOWPOSCHANGING: 
		return OnWindowPosChanged(wParam, lParam);
	}
	return FWnd::WindowProc(nMessage, wParam, lParam);
}

BOOL DDToolBar::AddButton(TBBUTTON& biButton)
{
	Tool* pNewButtons = new Tool[m_nButtonCount+1];
	if (NULL == pNewButtons)
		return FALSE;

	memcpy (pNewButtons, m_pButtons, sizeof(Tool)*m_nButtonCount);
	memcpy (&pNewButtons[m_nButtonCount].biButton, &biButton, sizeof(biButton));

	if (TBSTYLE_SEP == biButton.fsStyle)
	{
		pNewButtons[m_nButtonCount].sizeButton.cx = eSeparatorGap;
		pNewButtons[m_nButtonCount].sizeButton.cy = m_sizeButton.cy;
	}
	else
		pNewButtons[m_nButtonCount].sizeButton = m_sizeButton;
	
	if (m_pButtons)
		delete [] m_pButtons;
	m_pButtons = pNewButtons;

	m_nButtonCount++;
	RecalcLayout();
	return TRUE;
}

BOOL DDToolBar::InsertButton(int nIndex, TBBUTTON& biButton)
{
	if (nIndex > m_nButtonCount)
		return FALSE;

	Tool* pNewButtons = new Tool[m_nButtonCount+1];
	assert(pNewButtons);
	if (NULL == pNewButtons)
		return FALSE;

	memcpy (&pNewButtons, m_pButtons, sizeof(Tool)*m_nButtonCount);
	memcpy (&pNewButtons[m_nButtonCount].biButton, &biButton, sizeof(biButton));

	if (m_pButtons)
		delete [] m_pButtons;
	m_pButtons = pNewButtons;

	m_nButtonCount++;
	RecalcLayout();
	return TRUE;
}

BOOL DDToolBar::DeleteButton(int nIndex)
{
	if (nIndex > m_nButtonCount)
		return FALSE;

	memcpy(&m_pButtons[nIndex], &m_pButtons[nIndex+1], sizeof(Tool)*(m_nButtonCount - nIndex - 1));
	m_nButtonCount--;
	return TRUE;
}

BOOL DDToolBar::SetState(int nId, BYTE fsState)
{
	int nIndex = CommandIdToIndex(nId);
	if (nIndex != -1)
	{
		m_pButtons[nIndex].biButton.fsState = fsState;
		return TRUE;
	}
	return FALSE;
}

BOOL DDToolBar::GetState(int nId, BYTE& fsState)
{
	int nIndex = CommandIdToIndex(nId);
	if (nIndex != -1)
	{
		fsState = m_pButtons[nIndex].biButton.fsState;
		return TRUE;
	}
	return FALSE;
}

int DDToolBar::CommandIdToIndex(int nId)
{
	for (int nIndex = 0; nIndex < m_nButtonCount; nIndex++)
	{
		if (m_pButtons[nIndex].biButton.idCommand == nId)
			return nIndex;
	}
	return -1;
}

void DDToolBar::InvalidateButton(Tool& theButton)
{
	if (m_hWnd)
	{
		CRect rcButton;
		GetButtonRect(theButton, rcButton);
		InvalidateRect(&rcButton, FALSE);
	}	
}

void DDToolBar::InvalidateButton(int nId)
{
	int nIndex = CommandIdToIndex(nId);
	if (nIndex != -1)
		InvalidateButton(m_pButtons[nIndex]);
}

void DDToolBar::EnableButton(int nId, BOOL bFlag)
{
	BYTE fsState;
	GetState(nId, fsState);
	if (bFlag) 
		fsState |= TBSTATE_ENABLED;
	else 
		fsState &= ~TBSTATE_ENABLED;

	SetState(nId, fsState);
	InvalidateButton(nId);
}

void DDToolBar::CheckButton(int nId, BOOL bFlag)
{
	BYTE fsState;
	GetState(nId, fsState);

	if (bFlag) 
		fsState |= TBSTATE_CHECKED;
	else 
		fsState &= ~TBSTATE_CHECKED;

	SetState(nId, fsState);
	InvalidateButton(nId);
}

void DDToolBar::PaintToolButton(HDC   hDC, 
					            Tool& tTool, 
								int   nX, 
								int   nY, 
								int   nDX, 
								int   nDY, 
								BYTE  fsState)
{
	HBITMAP hbmpMono,hbmpMonoOld;
	int nBitmapX, nBitmapY;

	ImageList_GetIconSize(m_ilImages, &nBitmapX, &nBitmapY);
	
	hbmpMono = CreateBitmap(eMaxToolBitmapSize, eMaxToolBitmapSize, 1, 1, 0);
	hbmpMonoOld = SelectBitmap(s_hDCMono, hbmpMono);

	// determine offset of bitmap (vertically centered within button)
	POINT ptOffset;
	ptOffset.x = (nDX - nBitmapX) / 2;
	ptOffset.y = (nDY - nBitmapY) / 2;
	
	// Draw border part 
	CRect rc(nX, nX + nDX, nY, nY + nDY);

	if (Pressed(fsState) || Checked(fsState)) 
	{
		// pressed in or checked
		
		UIUtilities::Draw3DRect(hDC, 
							    CRect(nX, nX + nDX, nY, nY + nDY), 
							    GetSysColor(COLOR_WINDOWFRAME),
							    GetSysColor(COLOR_BTNHIGHLIGHT));
		UIUtilities::Draw3DRect(hDC,
							    CRect(nX + 1, nX + nDX - 2, nY + 1, nY + nDY - 2),
							    GetSysColor(COLOR_BTNSHADOW), 
							    GetSysColor(COLOR_BTNFACE));

		// for any depressed button, add one to the offsets.
		ptOffset.x += 1;
		ptOffset.y += 1;
	}
	else if (0 != (m_nToolbar & eIE) && (fsState & TBSTATE_FLYOUT)) 
	{
		UIUtilities::Draw3DRect(hDC,
							    CRect(nX, nX + nDX, nY, nY + nDY), 
							    GetSysColor(COLOR_BTNHIGHLIGHT),
							    GetSysColor(COLOR_BTNSHADOW));
	}

	if ((m_nToolbar&eNormal) && (!(Pressed(fsState) || Checked(fsState))))
	{
		// regular button look
		UIUtilities::Draw3DRect(hDC, 
								CRect(nX, nX + nDX, nY, nY + nDY), 
								GetSysColor(COLOR_BTNHIGHLIGHT),
								GetSysColor(COLOR_WINDOWFRAME));
		UIUtilities::Draw3DRect(hDC,
								CRect(nX + 1, nX + nDX - 2, nY + 1, nY + nDY - 2),
								GetSysColor(COLOR_BTNFACE), 
								GetSysColor(COLOR_BTNSHADOW));
	}
	
	// Draw image part
	if (Checked(fsState) || Indeterminate(fsState))
	{
		int cxBorder2 = GetSystemMetrics(SM_CXBORDER);
		int cyBorder2 = GetSystemMetrics(SM_CYBORDER);
		
		HGDIOBJ hbrOld = SelectBrush(hDC, s_hbrDither);
		
		PatBlt(s_hDCMono, 0, 0, nDX, nDY, WHITENESS);
		
		ptOffset.x -= cxBorder2;
		ptOffset.y -= cyBorder2;
		
		SetTextColor(hDC,0L);              // 0 -> 0
		SetBkColor(hDC, (COLORREF)0x00FFFFFFL); // 1 -> 1
		
		int delta = cxBorder2 * 2;
		
		// only draw the dither brush where the mask is 1's
		BitBlt(hDC, 
			   nX + cxBorder2, 
			   nY + cyBorder2, 
			   nDX-delta, 
			   nDY-delta, 
			   s_hDCMono, 
			   0, 
			   0, 
			   s_ROP_DSPDxax);

		SelectBrush(hDC,hbrOld);
	}
	
	if (Enabled(fsState) && !Indeterminate(fsState))
	{
		// normal image version
		ImageList_Draw(m_ilImages, 
					   tTool.biButton.iBitmap,
					   hDC,
					   nX + ptOffset.x,
					   nY + ptOffset.y,
					   ILD_NORMAL);
		if (fsState&TBSTATE_PRESSED)
		{
			SelectBitmap(s_hDCMono,hbmpMonoOld);
			DeleteBitmap((HGDIOBJ*)hbmpMono);
			return;
		}
	}
	if (Disabled(fsState) || Indeterminate(fsState))
	{
		PatBlt(s_hDCMono, 0, 0, nDX, nDY, WHITENESS);
		
		SetBkColor(hDC, RGB(255,255,255));
		SetTextColor(hDC, RGB(0,0,0)); 

		ImageList_Draw(m_ilImages, 
					   tTool.biButton.iBitmap, 
			           s_hDCMono,
					   0,
					   0,
					   ILD_NORMAL);
		CRect rc(0, nDX, 0, nDY);
		InvertRect(s_hDCMono, &rc);

		if (Indeterminate(fsState))
		{
			BitBlt(hDC, 
				   nX + ptOffset.x + 1, 
				   nY + ptOffset.y + 1,
				   nBitmapX, 
				   nBitmapY,
				   s_hDCMono, 
				   0, 
				   0, 
				   SRCPAINT);
		}
		SetBkColor(hDC,RGB(0,0,0)); 
		SetTextColor(hDC,RGB(255,255,255));
		BitBlt(hDC, 
			   nX + ptOffset.x, 
			   nY + ptOffset.y, 
			   nBitmapX, 
			   nBitmapY,
			   s_hDCMono, 
			   0, 
			   0, 
			   SRCAND);
		SetBkColor(hDC,RGB(128,128,128)); 
		SetTextColor(hDC,RGB(0,0,0));
		BitBlt(hDC, 
			   nX + ptOffset.x, 
			   nY + ptOffset.y,
			   nBitmapX, 
			   nBitmapY,
			   s_hDCMono, 
			   0, 
			   0, 
			   SRCPAINT);
	}
	SelectBitmap(s_hDCMono, hbmpMonoOld);
	BOOL bResult = DeleteBitmap((HGDIOBJ*)hbmpMono);
	assert(bResult);
}

void DDToolBar::GetHorzFixedOptimalRect(CRect& rc)
{
	rc.Inflate(-m_nHorzBorder, -m_nVertBorder);	 	
	CRect rcTool = rc;
	long nX = rcTool.left;
	long nY = rcTool.top;
	long nToolMaxHeightOnLine = 0;
	SIZE  sizeTool;
	int nMaxX = 0;
	int nMaxY = 0;
	for (int nTool = 0; nTool < m_nButtonCount; nTool++)
	{
		sizeTool = m_pButtons[nTool].sizeButton;
		if ((nX+sizeTool.cx-1) > rcTool.right)
		{
			nX = rcTool.left;
			nY += m_nVerticalGap;
			nY += nToolMaxHeightOnLine;
			nToolMaxHeightOnLine = 0;
		}
		if (sizeTool.cy > nToolMaxHeightOnLine)
			nToolMaxHeightOnLine = sizeTool.cy;

		nX += sizeTool.cx;
		if (nX > nMaxX)
			nMaxX = nX;

		if ((nY + sizeTool.cy) > nMaxY)
			nMaxY = nY + sizeTool.cy;
	}		
	rc.right = nMaxX;
	rc.bottom = nMaxY;
	rc.Inflate(m_nHorzBorder, m_nVertBorder);	 
}

void DDToolBar::GetVertFixedOptimalRect(CRect& rc)
{
	CRect rcTemp;
	int nToolIdx = m_sizeButton.cx;
	int nWidth = m_nVertBorder + m_nHorzBorder;
	while (nWidth < 1000)
	{
		rcTemp = rc;
		rcTemp.right = rcTemp.left + nWidth;
		GetHorzFixedOptimalRect(rcTemp);
		if ((rcTemp.Height()) < (rc.Height()+m_nVertBorder))
			break;

		nWidth += nToolIdx;
	}
	rc = rcTemp;
}

int DDToolBar::HitTest(const POINT& pt)
{
	POINT ptAdjusted = pt;
	ptAdjusted.y += m_nScrollAmount;
	for (int nIndex = 0; nIndex < m_nButtonCount; nIndex++)
	{
		if (PtInRect(&(m_pButtons[nIndex].rcButton), ptAdjusted))
			return nIndex;
	}
	return -1;
}

void DDToolBar::GetButtonRect(Tool& aButton, CRect& rcButton)
{
	rcButton = aButton.rcButton;
	rcButton.Offset(0, -m_nScrollAmount);
	rcButton.right++;
	rcButton.bottom++;
}

void DDToolBar::RecalcLayout()
{
	if (NULL == m_hWnd)
		return;

	CRect rcTool;
	GetClientRect(rcTool);
	rcTool.Inflate(-m_nHorzBorder,-m_nVertBorder);	 
	long nX = rcTool.left;
	long nY = rcTool.top;
	long nToolMaxHeightOnLine = 0;
	for (int nTool = 0; nTool < m_nButtonCount; nTool++)
	{
		if (m_pButtons[nTool].biButton.fsStyle & TBSTYLE_KEEPTOGETHER && nTool + 1 < m_nButtonCount)
		{
			int nToolIndex = nTool + 1;
			int nWidth = m_pButtons[nTool].sizeButton.cx + m_pButtons[nToolIndex].sizeButton.cx;
			while (nToolIndex + 1 < m_nButtonCount && m_pButtons[nToolIndex].biButton.fsStyle & TBSTYLE_KEEPTOGETHER)
			{
				nToolIndex++;
				nWidth += m_pButtons[nToolIndex].sizeButton.cx;
			}
			if ((nX + nWidth - 1) > rcTool.right)
			{
				//
				// Move down a line
				//

				nX = rcTool.left;
				nY += m_nVerticalGap;
				nY += nToolMaxHeightOnLine;
				nToolMaxHeightOnLine = 0;
			}
		}
		else if ((nX + m_pButtons[nTool].sizeButton.cx - 1) > rcTool.right)
		{
			//
			// Move down a line
			//

			nX = rcTool.left;
			nY += m_nVerticalGap;
			nY += nToolMaxHeightOnLine;
			nToolMaxHeightOnLine = 0;
		}			
		if (m_pButtons[nTool].sizeButton.cy > nToolMaxHeightOnLine)
			nToolMaxHeightOnLine = m_pButtons[nTool].sizeButton.cy;

		m_pButtons[nTool].rcButton.left = nX;
		m_pButtons[nTool].rcButton.top = nY;
		m_pButtons[nTool].rcButton.right = m_pButtons[nTool].rcButton.left + m_pButtons[nTool].sizeButton.cx;
		m_pButtons[nTool].rcButton.bottom = m_pButtons[nTool].rcButton.top + m_pButtons[nTool].sizeButton.cy;

		nX += m_pButtons[nTool].sizeButton.cx;

		if (m_pButtons[nTool].biButton.fsStyle&TBSTYLE_PLACEHOLDER && IsWindow(m_pButtons[nTool].hWnd))
		{
			::MoveWindow(m_pButtons[nTool].hWnd,
						 m_pButtons[nTool].rcButton.left,
						 m_pButtons[nTool].rcButton.top,
						 m_pButtons[nTool].rcButton.Width(),
						 m_pButtons[nTool].rcButton.Height(),
						 TRUE);
		}
	}		
}

LRESULT DDToolBar::OnTimer (WPARAM wParam, LPARAM lParam)
{
	if (NULL == m_pActiveToolBar)
		return FALSE;

	POINT pt;
	GetCursorPos(&pt);
	ScreenToClient(pt);
	int nHitIndex = HitTest(pt);
	if (m_pActiveToolBar != this || m_nActiveButton != nHitIndex)
	{
		KillTimer(m_hWnd, eTimer);
		m_pActiveToolBar->m_pButtons[m_nActiveButton].biButton.fsState &= ~TBSTATE_FLYOUT;
		m_pActiveToolBar->InvalidateButton(m_pActiveToolBar->m_pButtons[m_nActiveButton]);
		m_pActiveToolBar->UpdateWindow();
		m_pActiveToolBar = NULL;
	}
	return FALSE;
}

LRESULT DDToolBar::Draw(HDC hDC, CRect& rcBound)
{
	HDC hDCOff = m_dbDraw.RequestDC(hDC, rcBound.Width(), rcBound.Height());
	if (NULL == hDCOff)
		hDCOff = hDC;
	else
		rcBound.Offset(-rcBound.left, -rcBound.top);		

	CRect rc = rcBound;

	if (m_nToolbar&eRaised)
	{
		// Draw border around control
		SetBkColor(hDCOff, GetSysColor(COLOR_BTNHIGHLIGHT));
		rc.bottom = rc.top + 1;
		UIUtilities::FastRect(hDCOff, rc);
		
		rc = rcBound;
		rc.right = rc.left + 1;
		UIUtilities::FastRect(hDCOff, rc);
		
		rc = rcBound;
		
		SetBkColor(hDCOff, GetSysColor(COLOR_BTNSHADOW));
		rc.top = rc.bottom - 1;
		UIUtilities::FastRect(hDCOff, rc);

		rc = rcBound;

		SetBkColor(hDCOff, GetSysColor(COLOR_BTNSHADOW));
		rc.left = rc.right - 1;
		UIUtilities::FastRect(hDCOff, rc);
		
		rcBound.Inflate(-1, -1);
	}
	else if (m_nToolbar&eSunken)
	{
		SetBkColor(hDCOff, GetSysColor(COLOR_BTNHIGHLIGHT));
		rc = rcBound;
		rc.top = rc.bottom - 1;
		UIUtilities::FastRect(hDCOff, rc);
		
		rc = rcBound;
		rc.left = rc.right - 1;
		UIUtilities::FastRect(hDCOff, rc);
		
		// Draw border around control
		rcBound.Inflate(-1, -1);
		SetBkColor(hDCOff, GetSysColor(COLOR_3DLIGHT));
		rc.bottom = rc.top + 1;
		UIUtilities::FastRect(hDCOff, rc);

		rc = rcBound;
		rc.right = rc.left + 1;
		UIUtilities::FastRect(hDCOff, rc);

		rcBound.Inflate(-1, -1);
	}

	UIUtilities::FillSolidRect(hDCOff, rcBound, GetSysColor(COLOR_BTNFACE));

	int nX, nY, nDX, nDY;
	for (int nTool = 0; nTool < m_nButtonCount; nTool++)
	{
		if (NULL == m_ilImages || 
			TBSTYLE_PLACEHOLDER == m_pButtons[nTool].biButton.fsStyle)
		{
			continue;
		}

		nX = rcBound.left + m_pButtons[nTool].rcButton.left;
		nY = rcBound.top + m_pButtons[nTool].rcButton.top - m_nScrollAmount;
		nDX = m_pButtons[nTool].rcButton.Width();
		nDY = m_pButtons[nTool].rcButton.Height();

		if (TBSTYLE_SEP == m_pButtons[nTool].biButton.fsStyle)
		{
			SetBkColor(hDCOff, GetSysColor(COLOR_BTNSHADOW));
			rc.left = nX + nDX / 2;
			rc.right = rc.left + 1;
			rc.top = nY;
			rc.bottom = nY + nDY;
			UIUtilities::FastRect(hDCOff, rc);
			rc.left++;
			rc.right++;
			SetBkColor(hDCOff, GetSysColor(COLOR_BTNHIGHLIGHT));
			UIUtilities::FastRect(hDCOff, rc);
		}
		else
		{
			PaintToolButton(hDCOff, 
						    m_pButtons[nTool], 
							nX, 
							nY, 
							nDX, 
							nDY, 
							m_pButtons[nTool].biButton.fsState);
		}
	}

	if (hDCOff != hDC)
		m_dbDraw.Paint(hDC, 0, 0);
	return FALSE;
}

LRESULT DDToolBar::OnMouseMove(WPARAM wParam, LPARAM lParam)
{
	if (eIE&m_nToolbar)
	{
		POINT pt;
		pt.x = LOWORD(lParam);
		pt.y = HIWORD(lParam);
		int nHitIndex = HitTest(pt);
		if (m_pActiveToolBar != this || m_nActiveButton != nHitIndex)
		{
			if (m_pActiveToolBar)
			{
				KillTimer(m_hWnd, eTimer);
				m_pActiveToolBar->m_pButtons[m_nActiveButton].biButton.fsState &= ~TBSTATE_FLYOUT;
				m_pActiveToolBar->InvalidateButton(m_pActiveToolBar->m_pButtons[m_nActiveButton]);
				m_pActiveToolBar->UpdateWindow();
				m_pActiveToolBar = NULL;
			}
			if (-1 != nHitIndex)
			{
				if (m_pButtons[nHitIndex].biButton.fsState&TBSTATE_ENABLED)
				{
					m_pActiveToolBar = this;
					m_nActiveButton = nHitIndex;
					m_pActiveToolBar->m_pButtons[m_nActiveButton].biButton.fsState |= TBSTATE_FLYOUT;
					m_pActiveToolBar->InvalidateButton(m_pButtons[m_nActiveButton]);
					m_pActiveToolBar->UpdateWindow();
					KillTimer(m_hWnd, eTimer);
					// Check every 250ms if mouse on target
					SetTimer(m_hWnd, eTimer, 100, NULL);
				}
				
				if (m_pActiveToolBar && 
					m_pActiveToolBar->m_pButtons[nHitIndex].biButton.fsStyle != TBSTYLE_SEP)
				{
					POINT ptOffset;
					ptOffset.y = ptOffset.x = 0;
					CRect rcButton;
					ClientToScreen(ptOffset);
					pt.x += ptOffset.x;
					pt.y += ptOffset.y;
					GetButtonRect(m_pActiveToolBar->m_pButtons[nHitIndex], rcButton);
					rcButton.Offset(ptOffset.x, ptOffset.y);
					DDString strTemp;
					strTemp.LoadString(m_pActiveToolBar->m_pButtons[nHitIndex].biButton.idCommand);
					MAKE_WIDEPTR_FROMANSI(wTip, strTemp);
					UIUtilities::GetUIUtilities().ShowToolTip((void*)&m_pActiveToolBar->m_pButtons[nHitIndex], 
															  wTip, 
															  rcButton, 
															  pt);
				}
			}
		}
	}
	return FALSE;
}

LRESULT DDToolBar::OnLButtonDown(WPARAM wParam, LPARAM lParam)
{
	POINT pt;
	pt.x = LOWORD(lParam);
	pt.y = HIWORD(lParam);

	DWORD dwTime = GetTickCount();
	if ((dwTime-m_dwLastClickTime) < GetDoubleClickTime() && 
		m_ptLastClick.x == pt.x && 
		m_ptLastClick.y == pt.y)
	{
		OnDoubleClick(pt);
	}

	m_dwLastClickTime = dwTime;
	m_ptLastClick = pt;

	int nHitIndex = HitTest(pt);
	if (-1 == nHitIndex)
		return FALSE;

	// check if button is enabled/disabled
	if ((m_pButtons[nHitIndex].biButton.fsState&TBSTATE_ENABLED) == 0) 
		return FALSE;

	m_pButtons[nHitIndex].biButton.fsState |= TBSTATE_PRESSED;
	InvalidateButton(m_pButtons[nHitIndex]);
	
	UIUtilities::GetUIUtilities().HideToolTips(FALSE);

	try
	{
		UpdateWindow();
		SetCapture(m_hWnd);
		BOOL bInitState = 0 != (m_pButtons[nHitIndex].biButton.fsState&TBSTATE_CHECKED);
		BOOL bProcessing = TRUE;
		while (bProcessing)
		{
			MSG msg;
			GetMessage(&msg, NULL, NULL, NULL);

			if (GetCapture() != m_hWnd)
				break;

			switch (msg.message)
			{
				// handle movement/accept messages
				case WM_LBUTTONUP:
				{
					if (0 != (m_pButtons[nHitIndex].biButton.fsState&TBSTATE_PRESSED))
					{
						if (TBSTYLE_CHECK == m_pButtons[nHitIndex].biButton.fsStyle)
						{
							m_pButtons[nHitIndex].biButton.fsState &= ~TBSTATE_PRESSED;
							m_pButtons[nHitIndex].biButton.fsState &= ~TBSTATE_CHECKED;
							if (!bInitState)
								m_pButtons[nHitIndex].biButton.fsState |= TBSTATE_CHECKED;
						}
						m_pButtons[nHitIndex].biButton.fsState &= ~TBSTATE_PRESSED;
						::PostMessage(GetParent(m_hWnd),
									  WM_COMMAND,
									  MAKELONG(m_pButtons[nHitIndex].biButton.idCommand, 0),
									  0);
						if (m_pButtons[nHitIndex].biButton.fsState&TBSTATE_CHECKED || 
							0 == (m_pButtons[nHitIndex].biButton.fsState&TBSTATE_ENABLED))
						{
							m_pButtons[nHitIndex].biButton.fsState &= ~TBSTATE_FLYOUT;
						}
					}
					POINT pt;
					pt.x = LOWORD(msg.lParam);
					pt.y = HIWORD(msg.lParam);
					if (-1 == HitTest(pt))
						m_pButtons[nHitIndex].biButton.fsState &= ~TBSTATE_FLYOUT;
					InvalidateButton(m_pButtons[nHitIndex]);
					UpdateWindow();
					bProcessing = FALSE;
				}
				break;

				case WM_MOUSEMOVE:
				{
					POINT ptNew;
					ptNew.x = LOWORD(msg.lParam);
					ptNew.y = HIWORD(msg.lParam);
					if (PtInRect(&(m_pButtons[nHitIndex].rcButton), ptNew))
					{
						if (0 == (m_pButtons[nHitIndex].biButton.fsState & TBSTATE_PRESSED))
						{
							m_pButtons[nHitIndex].biButton.fsState |= TBSTATE_PRESSED;
							InvalidateButton(m_pButtons[nHitIndex]);
							UpdateWindow();
						}
					}
					else
					{
						if (0 != (m_pButtons[nHitIndex].biButton.fsState & TBSTATE_PRESSED))
							m_pButtons[nHitIndex].biButton.fsState &= ~TBSTATE_PRESSED;
						m_pButtons[nHitIndex].biButton.fsState &= ~TBSTATE_FLYOUT;
						InvalidateButton(m_pButtons[nHitIndex]);
						UpdateWindow();
					}
				}
				break;

				// Just dispatch rest of the messages
				default:
					DispatchMessage(&msg);
					break;
			}
		}
		ReleaseCapture();
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
	return FALSE;
}

LRESULT DDToolBar::OnWindowPosChanged(WPARAM wParam, LPARAM lParam)
{
	return TRUE;
}

BOOL DDToolBar::LoadToolBar(LPCTSTR szResourceName)
{
	assert(szResourceName != NULL);

	struct CToolBarData
	{
		WORD wVersion;
		WORD wWidth;
		WORD wHeight;
		WORD wItemCount;

		WORD* Items()
		{ 
			return (WORD*)(this+1); 
		}
	};

	// determine location of the bitmap in resource fork
	HRSRC hRsrc = ::FindResource(g_hInstance, szResourceName, MAKEINTRESOURCE(eRTToolbar));
	if (NULL == hRsrc)
		return FALSE;

	HGLOBAL hGlobal = LoadResource(g_hInstance, hRsrc);
	if (NULL == hGlobal)
		return FALSE;

	CToolBarData* pData = (CToolBarData*)LockResource(hGlobal);
	assert(pData);
	if (NULL == pData)
		return FALSE;

	m_sizeImage.cx = pData->wWidth;
	m_sizeImage.cy = pData->wHeight;
	m_sizeButton.cx = pData->wWidth + 2 * m_nHorzBorder;
	m_sizeButton.cy = pData->wHeight + 2 * m_nVertBorder;
	m_nButtonCount = pData->wItemCount;
	assert(pData->wVersion == 1);

	UINT* pItems = new UINT[m_nButtonCount];
	assert(pItems);

	for (int i = 0; i < pData->wItemCount; i++)
		pItems[i] = pData->Items()[i];

	BOOL bResult = SetButtons(pItems, m_nButtonCount);
	delete [] pItems;

	// Load bitmap now that sizes are known by the toolbar control
	if (bResult)
		bResult = LoadBitmap(szResourceName);

	UnlockResource(hGlobal);
	FreeResource(hGlobal);
	return bResult;
}

BOOL DDToolBar::SetButtons(const UINT* pIDArray, int nIDCount)
{
	assert(nIDCount >= 1);  // must be at least one of them
	assert(pIDArray != NULL);

	// delete all existing buttons
	ResetContent();

	if (pIDArray)
	{
		m_pButtons = new Tool[m_nButtonCount];
		assert(m_pButtons);
	
		// add new buttons to the common control
		TBBUTTON button; 
		memset(&button, 0, sizeof(TBBUTTON));
		int iImage = 0;
		for (int i = 0; i < nIDCount; i++)
		{
			button.fsState = TBSTATE_ENABLED;
			if (0 == (button.idCommand = *pIDArray++))
			{
				// separator
				button.fsStyle = TBSTYLE_SEP;
				// width of separator includes 8 pixel overlap
				button.iBitmap = 8;
				m_pButtons[i].sizeButton.cx = eSeparatorGap;
			}
			else
			{
				// a command button with image
				button.fsStyle = TBSTYLE_BUTTON;
				button.iBitmap = iImage++;
				m_pButtons[i].sizeButton.cx = m_sizeButton.cx;
			}
			m_pButtons[i].biButton = button;
			m_pButtons[i].sizeButton.cy = m_sizeButton.cy;
			m_pButtons[i].rcButton.SetEmpty();
			m_pButtons[i].nMin = 0;
			m_pButtons[i].nMax = 0;
		}
		RecalcLayout();
	}
	return TRUE;
}

BOOL DDToolBar::LoadBitmap(LPCTSTR szResourceName)
{
	assert(szResourceName);

	//
	// Determine location of the bitmap in resource fork
	//

	HRSRC hRsrcImageWell = ::FindResource(g_hInstance, szResourceName, RT_BITMAP);
	if (NULL == hRsrcImageWell)
		return FALSE;

	m_ilImages = ImageList_LoadBitmap(g_hInstance, szResourceName, m_sizeImage.cx, 0, RGB(192, 192, 192)); 
	assert(m_ilImages);
	return TRUE;
}

//
// CSplitter
//

CSplitter::CSplitter()
 : m_hPrimaryCursor(NULL),
   m_hSplitVCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDC_HSPLITBAR))),
   m_hWnd(NULL)
{
	m_bTracking = FALSE;
	m_nRight = 0;
	m_nLeft = 0;
}

void CSplitter::OnInvertTracker()
{
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		// Invert the brush pattern (looks just like frame window sizing)
		HBRUSH hBrush = UIUtilities::CreateHalftoneBrush(TRUE);
		HBRUSH hBrushOld = SelectBrush(hDC, hBrush);
		CRect rcSplitter = m_rcTracker;
		ClientToScreen(m_hWnd, rcSplitter);
		PatBlt(hDC, 
			   rcSplitter.left, 
			   rcSplitter.top, 
			   rcSplitter.Width(), 
			   rcSplitter.Height(), 
			   PATINVERT);

		SelectBrush(hDC, hBrushOld);
		BOOL bResult = DeleteBrush(hBrush);
		assert(bResult);
		ReleaseDC(NULL, hDC);
	}
}

void CSplitter::SetSplitCursor(const BOOL& bTracking)
{
	if (bTracking)
		m_hPrimaryCursor = ::SetCursor(m_hSplitVCursor);
	else
		::SetCursor(m_hPrimaryCursor);
}

BOOL CSplitter::SplitterHitTest(const POINT& pt)
{
	if (PtInRect(&m_rcTracker, pt))
		return TRUE;
	return FALSE;
}

void CSplitter::StartTracking()
{
	SetCapture(m_hWnd);
	//SetFocus(m_hWnd);
	m_bTracking = TRUE;
	OnInvertTracker();
	SetSplitCursor(TRUE);
}

void CSplitter::DoTracking(const POINT& pt)
{
	POINT ptInternal = pt;

	if (GetCapture() != m_hWnd)
		StopTracking(pt);

	if (m_bTracking)
	{
		if ((short)ptInternal.x < m_nLeft)
			ptInternal.x = m_nLeft;
		if ((short)ptInternal.x > m_nRight)
			ptInternal.x = m_nRight;

		OnInvertTracker();
		m_rcTracker.Offset(ptInternal.x - m_rcTracker.left, 0);
		OnInvertTracker();
	}
	else if (SplitterHitTest(ptInternal))
		SetSplitCursor(TRUE);
}

BOOL CSplitter::StopTracking(const POINT& pt)
{
	if (!m_bTracking)
		return FALSE;

	ReleaseCapture();

	// Erase tracker rectangle
	POINT ptInternal = pt;
	
	OnInvertTracker();
	m_bTracking = FALSE;
	SetSplitCursor(FALSE);
	
	if ((short)ptInternal.x < m_nLeft)
		ptInternal.x = m_nLeft;

	if ((short)ptInternal.x > m_nRight)
		ptInternal.x = m_nRight;

	m_rcTracker.Offset(ptInternal.x - m_rcTracker.left, 0);
	return TRUE;
}

