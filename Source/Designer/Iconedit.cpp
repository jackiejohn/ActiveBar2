//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "DesignerInterfaces.h"
#include "..\Interfaces.h"
#include "IconEdit.h"
#include "IpServer.h"
#include "Globals.h"
#include "Resource.h"
#include "FDialog.h"
#include "..\EventLog.h"
#include "GDIUtil.h"
#include "Dialogs.h"
#include "Support.h"
#include <stdio.h>

//
// IconEdit.cpp
//

HPALETTE g_hPal = NULL;

HBITMAP CreateMaskBitmap(HDC hDC, HBITMAP hBitmap, int nWidth, int nHeight);

//
// Forward declarations
//

typedef VOID CALLBACK CircleDDAProc(int X, int Y, LPARAM pData);
void CircleDDA(int left, 
			   int top, 
			   int right, 
			   int bottom, 
			   CircleDDAProc* pProc, 
			   LPARAM pData);

typedef VOID CALLBACK CircleDDALineProc(int X, int Y, int nLen, LPARAM pData);
void CircleSolidDDA(int left, 
					int top, 
					int right, 
					int bottom, 
					CircleDDALineProc* pProc, 
					LPARAM pData);

//
// Utility Drawing functions
//

static HBITMAP ScaleBitmap(HBITMAP hBitmapSrc, SIZE sizeNew)
{
	HBITMAP hBitmapNew = NULL;
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return NULL;

	BITMAP bmInfo;
	GetObject(hBitmapSrc, sizeof(BITMAP), &bmInfo);
	int nWidth = bmInfo.bmWidth;
	int nHeight = bmInfo.bmHeight;

	hBitmapNew = CreateCompatibleBitmap(hDC, sizeNew.cx, sizeNew.cy);
	if (NULL == hBitmapNew)
		return NULL;

	HDC hDCMem1 = CreateCompatibleDC(hDC);
	if (hDCMem1)
	{
		HDC hDCMem2 = CreateCompatibleDC(hDC);
		if (hDCMem2)
		{
			HBITMAP hBitmapOld1 = SelectBitmap(hDCMem1, hBitmapNew);
			HBITMAP hBitmapOld2 = SelectBitmap(hDCMem2, hBitmapSrc);

			StretchBlt(hDCMem1, 
					   0, 
					   0, 
					   sizeNew.cx, 
					   sizeNew.cy, 
					   hDCMem2, 
					   0, 
					   0, 
					   nWidth, 
					   nHeight, 
					   SRCCOPY);

			SelectBitmap(hDCMem2, hBitmapOld2);
			SelectBitmap(hDCMem1, hBitmapOld1);

			DeleteDC(hDCMem2);
		}
		DeleteDC(hDCMem1);
	}
	
	ReleaseDC(NULL, hDC);

	return hBitmapNew;
}

//
// Drawing prinitives
//

static LPARAM			  s_pData = NULL;
static CircleDDAProc*     s_pEllipsePixCallBack = NULL;
static CircleDDALineProc* s_pEllipseLineCallBack = NULL;

//
// set_pixels
//

static void set_pixels (int x, int y, int xc, int yc, int n)
{
	if (n != 0)
	{
		switch (n)
		{
		case 1: // First Quad
			s_pEllipsePixCallBack(xc-x, yc-y, s_pData); 
			break;

		case 2: // Second Quad
			s_pEllipsePixCallBack(xc+x, yc-y, s_pData); 
			break;

		case 3: // Third Quad
			s_pEllipsePixCallBack(xc-x, yc+y, s_pData); 
			break;

		case 4: // Fourth Quad
			s_pEllipsePixCallBack(xc+x, yc+y, s_pData); 
			break;
		}
	}
	else
	{
		s_pEllipsePixCallBack(xc-x, yc-y, s_pData);
		s_pEllipsePixCallBack(xc+x, yc-y, s_pData);
		s_pEllipsePixCallBack(xc-x, yc+y, s_pData);
		s_pEllipsePixCallBack(xc+x, yc+y, s_pData);
	}
}

//
// set_pixels_solid
//

static void set_pixels_solid (int x2, int y2, int x, int y, int xc, int yc, int n)
{
	if (n != 0)
	{
		switch (n)
		{
		case 1: 
			s_pEllipseLineCallBack(xc-x, yc-y, x+1, s_pData);
			break;

		case 2: 
			s_pEllipseLineCallBack(xc, yc-y, x+1, s_pData);
			break;

		case 3: 
			s_pEllipseLineCallBack(xc-x, yc+y, x+1, s_pData);
			break;

		case 4: 
			s_pEllipseLineCallBack(xc, yc+y, x+1, s_pData);
			break;
		}
	}
	else
	{
		s_pEllipseLineCallBack(xc-x, yc-y, 2*x, s_pData);
		s_pEllipseLineCallBack(xc-x, yc+y, 2*x, s_pData);
	}
}

//
// smooth
//

static void smooth (int xc, int yc, int a0, int b0, int region, int rem1, int rem2)
{
   int x = 0;
   int y = b0;
   long a = a0;
   long b = b0;

   long Asquared = a * a;
   long TwoAsquared = 2 * Asquared;
   long Bsquared = b * b;
   long TwoBsquared = 2 * Bsquared;

   long d = Bsquared - Asquared * b + Asquared/4L;
   long dx = 0;
   long dy = TwoAsquared * b;

	while (dx < dy)
	{
		set_pixels (x, y, xc,      yc,      1);
		set_pixels (x, y, xc+rem1, yc,      2);
		set_pixels (x, y, xc,	   yc+rem2, 3);
		set_pixels (x, y, xc+rem1, yc+rem2, 4);

		if (d > 0L)
		{
			--y;
			dy -= TwoAsquared;
			d -= dy;
		}
		++x;
		dx += TwoBsquared;
		d += Bsquared + dx;
   }

   d += (3L * (Asquared - Bsquared) / 2L - (dx + dy)) / 2L;

   while (y >= 0)
   {
		set_pixels (x, y, xc,      yc,      1);
		set_pixels (x, y, xc+rem1, yc,      2);
		set_pixels (x, y, xc,      yc+rem2, 3);
		set_pixels (x, y, xc+rem1, yc+rem2, 4);

      if (d < 0L)
      {
         ++x;
         dx += TwoBsquared;
         d += dx;
      }

      --y;
      dy -= TwoAsquared;
      d += Asquared - dy;
   }
}

//
// smooth_lx
//

static void smooth_lx (int x2, int y2, int xc, int yc, int a0, int b0, int region, int rem1, int rem2)
{
    int x = 0;
    int y = b0;
    long a = a0;
    long b = b0;

    long Asquared = a * a;
    long TwoAsquared = 2 * Asquared;
    long Bsquared = b * b;
    long TwoBsquared = 2 * Bsquared;

	long d = Bsquared - Asquared * b + Asquared/4L;
	long dx = 0;
	long dy = TwoAsquared * b;

    while (dx < dy)
    {
		set_pixels_solid (x2, y2, x, y, xc,      yc,      1);
		set_pixels_solid (x2, y2, x, y, xc+rem1, yc,      2);
		set_pixels_solid (x2, y2, x, y, xc,		 yc+rem2, 3);
		set_pixels_solid (x2, y2, x, y, xc+rem1, yc+rem2, 4);

		if (d > 0L)
		{
			--y;
			dy -= TwoAsquared;
			d -= dy;
		}
		++x;
		dx += TwoBsquared;
		d += Bsquared + dx;
	}

    d += (3L * (Asquared - Bsquared) / 2L - (dx + dy)) / 2L;

	while (y >= 0)
    {
		set_pixels_solid (x2, y2, x, y, xc,      yc,      1);
		set_pixels_solid (x2, y2, x, y, xc+rem1, yc,      2);
		set_pixels_solid (x2, y2, x, y, xc,      yc+rem2, 3);
		set_pixels_solid (x2, y2, x, y, xc+rem1, yc+rem2, 4);

		if (d < 0L)
		{
			++x;
			dx += TwoBsquared;
			d += dx;
		}

        --y;
        dy -= TwoAsquared;
        d += Asquared - dy;
    }
}

//
// CircleDDA
//

static void CircleDDA(int left, int top, int right, int bottom, CircleDDAProc* pProc, LPARAM pData)
{
	s_pEllipsePixCallBack = pProc;
	s_pData = pData;
	
	int xc = (left + right) / 2;
	int yc = (top + bottom) / 2;
	int x_rad = xc - left;
	int y_rad = yc - top;
	smooth (xc, yc, x_rad, y_rad, 0, (right-left)&1, (bottom-top)&1);
}

//
// CircleSolidDDA
//

static void CircleSolidDDA(int left, int top, int right, int bottom, CircleDDALineProc* pProc, LPARAM pData)
{
	assert(pProc);
	assert(pData);

	s_pEllipseLineCallBack = pProc;
	s_pData = pData;

	int xc = (left+right)/2;
	int yc = (top+bottom)/2;
	int x_rad = xc-left;
	int y_rad = yc-top;
	smooth_lx(xc+x_rad-1, yc+y_rad-1, xc, yc, x_rad, y_rad, 0, (right-left)&1, (bottom-top)&1);
}

//
// Class ColorWell
//

ColorWell::ColorWell()
{
	m_nCurSel = 0;
	m_bSupportSelection = TRUE;
	m_szMsgTitle.LoadString(IDS_ACTIVEBAR);
}

//
// CreateWin
//

BOOL ColorWell::CreateWin(HWND hWndParent, CRect rc, UINT nId)
{
	RegisterWindow(DD_WNDCLASS("ColorPalette"), 
				   CS_VREDRAW|CS_HREDRAW, 
				   0, 
				   ::LoadCursor(NULL, IDC_ARROW));
	if (CreateEx(0,
				 _T(""),
				 WS_TABSTOP|WS_CHILD|WS_VISIBLE,
				 rc.left,
				 rc.top,
				 rc.Width(),
				 rc.Height(),
				 hWndParent,
				 (HMENU)nId,
				 NULL) == 0)
	{
		return FALSE;
	}
	return TRUE;
}

//
// WindowProc
//

LRESULT ColorWell::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcClient;
					GetClientRect(rcClient);
					Draw(hDC, rcClient);
					EndPaint(m_hWnd, &ps);
					return 0;
				}
			}
			break;

		case WM_LBUTTONDOWN:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				OnLButtonDown((DWORD)wParam, pt);
			}
			break;

		case WM_RBUTTONDOWN:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				OnRButtonDown((DWORD)wParam, pt);
			}
			break;

		case WM_SETFOCUS:
		case WM_KILLFOCUS:
			{
				if (m_bSupportSelection)
				{
					int nIndex = 0;
					if (m_nCurSel != -1)
						nIndex = m_nCurSel;
					InvalidateItem(nIndex);
				}
			}
			break;

		case WM_KEYDOWN:
			OnKeyDown((int)wParam,lParam);
			break;

		case WM_GETDLGCODE:
			return DLGC_WANTARROWS;
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
// OnKeyDown
//

void ColorWell::OnKeyDown(int nChar, LPARAM nKeyData)
{
	if (!m_bSupportSelection)
		return;

	int nNewIndex = m_nCurSel;
	if (0 == m_nCurSel) 
		nNewIndex = 0;

	int x = nNewIndex / eNX;
	int y = nNewIndex % eNX;
	switch(nChar)
	{
	case VK_UP:
		if (y != 0)
			--y;
		break;

	case VK_DOWN:
		if (y != (eNX-1))
			++y;
		break;

	case VK_LEFT:
		if (x != 0)
			--x;
		break;

	case VK_RIGHT:
		if (x != (eNY-1))
			++x;
		break;
	}

	nNewIndex = x * eNX + y;
	if (nNewIndex != m_nCurSel)
	{
		InvalidateItem(m_nCurSel);
		m_nCurSel = nNewIndex;
		InvalidateItem(m_nCurSel);
		::SendMessage(GetParent(m_hWnd),
			          WM_COMMAND,
					  MAKELONG(GetWindowLong(m_hWnd, GWL_ID), 0),
					  (LPARAM)m_hWnd);
	}
}

//
// Colors for the color well
//

static DWORD crPaintColorBox[ColorWell::eNunOfColors]=
{
	0x0, 	  0x00FFFF, 0x00FF00, 0xFF0000,
	0xFFFFFF, 0x008080, 0x008000, 0x800000,
	0xC0C0C0, 0x0000FF, 0xFFFF00, 0xFF00FF,
	0x808080, 0x000080, 0x808000, 0x800080
};

//
// Draw
//

void ColorWell::Draw(HDC hDC, CRect& rcClient)
{
	try
	{
		int x,y;
		HBRUSH hColor;
		CRect rc(rcClient);
		CRect rc2;
		BOOL bHasFocus = (GetFocus() == m_hWnd);

		for (x = 0; x < eNY; x++)
		{
			rc.right = rc.left + eBoxSize;
			for (y = 0; y < eNX; y++)
			{
				rc.bottom = rc.top + eBoxSize;
				if ((x * eNX + y) == m_nCurSel && m_bSupportSelection)
					DrawEdge(hDC, &rc, EDGE_RAISED, BF_RECT);
				else 
					DrawEdge(hDC, &rc, EDGE_SUNKEN, BF_RECT);

				rc2 = rc;
				rc2.Inflate(-2, -2);
				
				hColor = CreateSolidBrush(crPaintColorBox[x*eNX+y]);
				FillRect(hDC, &rc2, hColor);
				DeleteBrush(hColor);

				rc.top += eBoxSize;
				if (bHasFocus)
				{
					if ((m_nCurSel == (x * eNX + y) || (-1 == m_nCurSel && 0 == x && 0 == y)) && m_bSupportSelection)
					{
						rc2.Inflate(-1, -1);
						DrawFocusRect(hDC, &rc2);
					}
				}
			}
			rc.left += eBoxSize;
			rc.top = rcClient.top;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnLButtonDown
//

void ColorWell::OnLButtonDown(UINT nFlags, POINT pt)
{
	SetFocus();

	int nIndex;
	if (pt.y > (eNX * eBoxSize))
		nIndex = eNX - 1;
	else
		nIndex = pt.y / eBoxSize;
	
	if (pt.x > (eNY * eBoxSize))
		nIndex += (eNY - 1) * eNX;
	else
		nIndex += (pt.x / eBoxSize) * eNX;
	
	if (m_bSupportSelection)
		InvalidateItem(m_nCurSel);

	m_nCurSel = nIndex;
	if (m_bSupportSelection)
		InvalidateItem(m_nCurSel);

	::SendMessage(GetParent(m_hWnd),
			      WM_COMMAND,
				  MAKELONG(GetWindowLong(m_hWnd, GWL_ID), BN_CLICKED), 
				  Painter::eLeftButton);
}

//
// OnRButtonDown
//

void ColorWell::OnRButtonDown(UINT nFlags, POINT pt)
{
	SetFocus();

	int nIndex;
	if (pt.y > (eNX * eBoxSize))
		nIndex = eNX - 1;
	else
		nIndex = pt.y / eBoxSize;
	
	if (pt.x > eNY * eBoxSize)
		nIndex += (eNY - 1) * eNX;
	else
		nIndex += (pt.x / eBoxSize) * eNX;
	
	if (m_bSupportSelection)
		InvalidateItem(m_nCurSel);

	m_nCurSel = nIndex;
	if (m_bSupportSelection)
		InvalidateItem(m_nCurSel);

	::SendMessage(GetParent(m_hWnd),
			      WM_COMMAND,
				  MAKELONG(GetWindowLong(m_hWnd, GWL_ID), BN_CLICKED), 
				  Painter::eRightButton);
}

//
// SetCurSel
//

void ColorWell::SetCurSel(int nNewIndex)
{
	if (nNewIndex != m_nCurSel)
	{
		if (m_bSupportSelection)
			InvalidateItem(m_nCurSel);

		m_nCurSel = nNewIndex;
		
		if (m_bSupportSelection)
			InvalidateItem(m_nCurSel);
	}
}

//
// InvalidateItem
//

void ColorWell::InvalidateItem(int nIndex)
{
	if (-1 == nIndex)
		return;

	CRect rcClient;
	GetClientRect(rcClient);
	rcClient.left += (nIndex / eNX) * eBoxSize;
	rcClient.top += (nIndex % eNX) * eBoxSize;
	rcClient.right = rcClient.left + eBoxSize;
	rcClient.bottom = rcClient.top + eBoxSize;
	InvalidateRect(&rcClient, FALSE);
}

//
// GetSelIndex
//

int ColorWell::GetSelIndex(COLORREF crColor)
{
	for (int nColor = 0; nColor < eNunOfColors; nColor++)
	{
		if (crPaintColorBox[nColor] == crColor)
			return nColor;
	}
	return -1;
}

//
// GetCurColor
//

COLORREF ColorWell::GetCurColor()
{
	if (-1 == m_nCurSel)
		return 0;
	return crPaintColorBox[m_nCurSel];
}

//
// Class Preview
//

BOOL Preview::CreateWin(HWND hWndParent, PREVIEWTYPES eType, UINT nId)
{
	m_szMsgTitle.LoadString(IDS_ACTIVEBAR);
	m_nType = eType;
	RegisterWindow(DD_WNDCLASS("IconEditorPreview"),
				   CS_VREDRAW|CS_HREDRAW|CS_DBLCLKS,
				   NULL,
				   ::LoadCursor(NULL, IDC_ARROW));
	if (CreateEx(WS_EX_CLIENTEDGE,
				 _T(""), 
		         WS_CHILD|WS_VISIBLE,
			     0,
			     0,
			     0,
			     0,
			     hWndParent,
			     (HMENU)nId))
	{
		return TRUE;
	}
	return FALSE;
}

//
// WindowProc
//

LRESULT Preview::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcClient;
					GetClientRect(rcClient);
					Draw(hDC, rcClient);
					EndPaint(m_hWnd, &ps);
					return 0;
				}
			}
			break;

		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			switch (m_nType)
			{
			case eColor:
				{
					POINT pt;
					pt.x = LOWORD(lParam);
					pt.y = HIWORD(lParam);

					if (PtInRect(&m_rcScreen, pt))
					{
						//
						// Background or foreground color changed
						//

						::SendMessage(GetParent(m_hWnd),
									  WM_COMMAND,
									  MAKELONG(GetWindowLong(m_hWnd, GWL_ID), BN_CLICKED), 
									  WM_LBUTTONDOWN == nMsg ? Painter::eLeftButton : Painter::eRightButton);
					}
					else if (PtInRect(&m_rcCycle, pt))
					{
						//
						// The Cycle Bitmap was hit
						//

						::SendMessage(GetParent(m_hWnd),
									  WM_COMMAND,
									  MAKELONG(GetWindowLong(m_hWnd, GWL_ID), BN_CLICKED), 
									  5);
					}
				}
				break;
			}
			break;

		case WM_LBUTTONDBLCLK:
			switch (m_nType)
			{
			case eColor:
				{
					POINT pt;
					pt.x = LOWORD(lParam);
					pt.y = HIWORD(lParam);

					if (PtInRect(&m_rcScreen, pt))
					{
						//
						// Background or foreground color changed
						//

						::SendMessage(GetParent(m_hWnd),
									  WM_COMMAND,
									  MAKELONG(GetWindowLong(m_hWnd, GWL_ID), BN_CLICKED), 
									  WM_LBUTTONDOWN == nMsg ? Painter::eLeftButton : Painter::eRightButton);
					}
					else if (PtInRect(&m_rcCycle, pt))
					{
						//
						// The Cycle Bitmap was hit
						//

						::SendMessage(GetParent(m_hWnd),
									  WM_COMMAND,
									  MAKELONG(GetWindowLong(m_hWnd, GWL_ID), BN_CLICKED), 
									  5);
					}
					else if (PtInRect(&m_rcBackColor, pt))
					{
						GetGlobals().m_pDefineColor->Color(m_pIconEdit->m_crBackground);
						int nResult = GetGlobals().m_pDefineColor->DoModal(m_hWnd);
						if (ID_ADDCOLOR == nResult)
						{
							m_pIconEdit->m_crBackground = GetGlobals().m_pDefineColor->Color();
							InvalidateRect(&m_rcBackColor, FALSE);
						}
					}
					else if (PtInRect(&m_rcForeColor, pt))
					{
						GetGlobals().m_pDefineColor->Color(m_pIconEdit->m_crForeground);
						int nResult = GetGlobals().m_pDefineColor->DoModal(m_hWnd);
						if (ID_ADDCOLOR == nResult)
						{
							m_pIconEdit->m_crForeground = GetGlobals().m_pDefineColor->Color();
							InvalidateRect(&m_rcForeColor, FALSE);
						}
					}
				}
				break;
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
// UpdateDibSection
//

void Preview::UpdateDibSection()
{
	try
	{
		HDC hDC = GetDC(m_hWnd);
		assert(hDC);
		if (NULL == hDC)
			return;

		HPALETTE hOldPal;
		if (g_hPal)
		{
			hOldPal = SelectPalette(hDC, g_hPal, FALSE);
			RealizePalette(hDC);
		}

		CRect rcClient;
		GetClientRect(rcClient);
		PaintDibSection(hDC, rcClient);
		
		if (g_hPal)
			SelectPalette(hDC, hOldPal, FALSE);	

		ReleaseDC(m_hWnd, hDC);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// PaintDibSection
//

void Preview::PaintDibSection(HDC hDCOff, const CRect& rc)
{
	// Blast the dib section on the screen
	SIZE size = m_pIconEdit->m_pntEdit.m_size;

	// Now we have to transfer COLORREF array to 24 bit dibsection
	COLORREF crBtnFace = GetSysColor(COLOR_BTNFACE);
	COLORREF crBtnShadow = GetSysColor(COLOR_BTNSHADOW);
	COLORREF crBtnHighlight = GetSysColor(COLOR_BTNHIGHLIGHT);
	
	HBITMAP hDibSection;
	void* pDibSectionBits;
	HDC hDC = GetDC(NULL);
	assert(hDC);
	if (hDC)
	{
		hDibSection = CreateDIBSection(hDC,
									   &(m_pIconEdit->m_pntEdit.m_bmInfo),
									   DIB_RGB_COLORS,
									   &pDibSectionBits,
									   0,
									   0);
		ReleaseDC(NULL, hDC);
		if (!hDibSection)
			return;
	}
	
	LPBYTE pLineBits = (LPBYTE)pDibSectionBits;
	BOOL bPressed = m_pIconEdit->m_nImageIndex == 1;
	
	COLORREF* pColArray = m_pIconEdit->m_pntEdit.m_pcrImageData + (size.cx * (size.cy-1));
	int nDibLineWidth = ((size.cx * 3) * 8 + 31) / 32 * 4;

	LPBYTE pBits;
	DWORD dwColumn;
	int x, y;

	for (y = size.cy - 1; y >= 0; y--)
	{
		pBits = pLineBits;
		for (x = 0; x < size.cx; x++)
		{
			dwColumn = *pColArray;
			if (0xFFFFFFFF == dwColumn)
			{
				if (bPressed)
				{
					if ((x+y) & 1)
						dwColumn = crBtnShadow;
					else
						dwColumn = crBtnHighlight;
				}
				else
					dwColumn = crBtnFace;
			}
			*(pBits+2) = (BYTE)(dwColumn&0xFF);
			dwColumn >>= 8;

			*(pBits+1) = (BYTE)(dwColumn&0xFF);
			dwColumn >>= 8;
			
			*pBits = static_cast<BYTE>(dwColumn&0xFF);
			
			pBits += 3;
			++pColArray;
		}
		pLineBits += nDibLineWidth;
		pColArray -= 2 * size.cx; 
	}
	
	HDC hDCMem = CreateCompatibleDC(0);
	int nCenterX = max (0, (rc.Width()-size.cx)/2);
	int nCenterY = max (0, (rc.Height()-size.cy)/2);
	int nWidth = min (rc.Width(), size.cx);
	int nHeight = min (rc.Height(), size.cy);
	HBITMAP hBitmapOld = SelectBitmap(hDCMem, hDibSection);

	StretchDIBits(hDCOff, 
				  nCenterX, 
				  nCenterY, 
				  nWidth, 
				  nHeight, 
				  0, 
				  0, 
				  size.cx, 
				  size.cy, 
				  pDibSectionBits, 
				  &m_pIconEdit->m_pntEdit.m_bmInfo, 
				  DIB_RGB_COLORS, 
				  SRCCOPY); 

	SelectBitmap(hDCMem, hBitmapOld);
	DeleteDC(hDCMem);
	DeleteBitmap(hDibSection);
}

//
// Draw
//

void Preview::Draw(HDC hDC, CRect& rcClient)
{
	try
	{
		HDC hDCOff = GetGlobals().FlickerFree()->RequestDC(hDC, rcClient.Width(), rcClient.Height());
		if (NULL == hDCOff)
			return;

		HBRUSH hBrush;
		HPALETTE hOldPal1, hOldPal2;

		if (g_hPal)
		{
			hOldPal1 = SelectPalette(hDCOff, g_hPal, FALSE);
			hOldPal2 = SelectPalette(hDC, g_hPal, FALSE);
			RealizePalette(hDC);
		}

		CRect rc(rcClient);
		FillRect(hDCOff, &rc, (HBRUSH)(COLOR_BTNFACE+1));
		
		switch (m_nType)
		{
		case ePreview:
			{
				//
				// Paint Icon
				//

				PaintDibSection(hDCOff, rcClient);
				
				//
				// Draw Border
				//

				SIZE size = m_pIconEdit->m_pntEdit.m_size;
				CRect rcBorder;
				rcBorder.left = (rcClient.Width() - size.cx) / 2;
				rcBorder.top = (rcClient.Height() - size.cy) / 2;
				rcBorder.right = rcBorder.left + size.cx;
				rcBorder.bottom = rcBorder.top + size.cy;
				rcBorder.Inflate(3, 3);

				switch(m_pIconEdit->m_nImageIndex)
				{
				case ddISNormal:
				case ddISMask:
				case ddISDisabled:
					DrawEdge(hDCOff, &rcBorder, BDR_RAISEDINNER, BF_RECT);
					break;

				case ddISGray:
					DrawEdge(hDCOff, &rcBorder, BDR_SUNKENOUTER, BF_RECT);
					break;
				}
			}
			break;

		case eColor:
			{
				//
				// BackColor location
				//

				m_rcBackColor = rc;
				m_rcBackColor.bottom -= 6;
				m_rcBackColor.top = m_rcBackColor.bottom - eScreenSize - 2;
				m_rcBackColor.right -= 3;
				m_rcBackColor.left = m_rcBackColor.right - eScreenSize - 2;

				//
				// ForeColor location
				//

				m_rcForeColor = m_rcBackColor;
				m_rcForeColor.Offset(-eCycleBitmapSize, -eCycleBitmapSize);

				//
				// Cycle location
				//

				m_rcCycle.bottom = m_rcBackColor.bottom;
				m_rcCycle.right = m_rcBackColor.left - 1;
				m_rcCycle.left = m_rcCycle.right - eCycleBitmapSize;
				m_rcCycle.top = m_rcCycle.bottom - eCycleBitmapSize;

				//
				// Back Color
				//
				
				DrawEdge(hDCOff, &m_rcBackColor, BDR_RAISEDINNER, BF_RECT);
				m_rcBackColor.Inflate(-1, -1);
				
				if (0xFFFFFFFF == m_pIconEdit->m_crBackground)
				{

					SetBkColor(hDCOff, RGB(255,255,255));
					SetTextColor(hDCOff, RGB(192,192,192));
					FillRect(hDCOff, &m_rcBackColor, GetGlobals().ScreenBrush());
				}
				else
				{
					hBrush = CreateSolidBrush(m_pIconEdit->m_crBackground);
					FillRect(hDCOff, &m_rcBackColor, hBrush);
					DeleteBrush(hBrush);
				}

				//
				// Fore Color
				//
				
				DrawEdge(hDCOff, &m_rcForeColor, BDR_RAISEDINNER, BF_RECT);
				m_rcForeColor.Inflate(-1, -1);
				
				if (0xFFFFFFFF == m_pIconEdit->m_crForeground)
				{
					SetBkColor(hDCOff, RGB(255, 255, 255));
					SetTextColor(hDCOff, RGB(192, 192, 192));
					FillRect(hDCOff, &m_rcForeColor, GetGlobals().ScreenBrush());
				}
				else
				{
					hBrush = CreateSolidBrush(m_pIconEdit->m_crForeground);
					FillRect(hDCOff, &m_rcForeColor, hBrush);
					DeleteBrush(hBrush);
				}

				//
				// Cycle Color
				//

				m_rcScreen = rc;
				m_rcScreen.Inflate(0, -(rc.Height() - eScreenSize)/2);
				m_rcScreen.left += 4;
				m_rcScreen.right = m_rcScreen.left + eScreenSize;

				SetBkColor(hDCOff, RGB(255, 255, 255));
				SetTextColor(hDCOff, RGB(192, 192, 192));

				FillRect(hDCOff, &m_rcScreen, GetGlobals().ScreenBrush());
				
				DrawEdge(hDCOff, &m_rcScreen, EDGE_SUNKEN, BF_RECT);

				HDC hDCMem = CreateCompatibleDC(hDC);
				HBITMAP hBitmapOld = SelectBitmap(hDCMem, GetGlobals().CycleBitmap());
				
				SetBkColor(hDCOff, RGB(255, 255, 255));
				SetTextColor(hDCOff, RGB(0, 0, 0));
				
				BitBlt(hDCOff, 
					   m_rcCycle.left, 
					   m_rcCycle.top, 
					   eCycleBitmapSize, 
					   eCycleBitmapSize, 
					   hDCMem, 
					   0, 
					   0, 
					   SRCAND);

				SetBkColor(hDCOff, RGB(0, 0, 0));
				SetTextColor(hDCOff, GetSysColor(COLOR_BTNTEXT));
				
				BitBlt(hDCOff, 
					   m_rcCycle.left, 
					   m_rcCycle.top, 
					   eCycleBitmapSize, 
					   eCycleBitmapSize, 
					   hDCMem, 
					   0, 
					   0, 
					   SRCPAINT);

				SelectBitmap(hDCMem, hBitmapOld);
				DeleteDC(hDCMem);
			}
			break;
		}

		GetGlobals().FlickerFree()->Paint(hDC, rcClient.left, rcClient.top);

		if (g_hPal)
		{
			SelectPalette(hDCOff, hOldPal1, FALSE);
			SelectPalette(hDC, hOldPal2, FALSE);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Class Painter
//

Painter::Painter()
	: m_pcrImageData(NULL)
{
	for (int nImage = 0; nImage < eNumOfImages; nImage++)
		m_pcrMainData[nImage] = NULL;
	m_size.cx = m_size.cy = eDefaultPainterSize;
	m_nMaxVScrollPos = 0;
	m_nVScrollPos = 0;
	m_nVDisplayBoxes = 0;
	m_nMaxHScrollPos = 0;
	m_nHScrollPos = 0;
	m_nHDisplayBoxes = 0;
	m_bRedraw = TRUE;
	m_szMsgTitle.LoadString(IDS_ACTIVEBAR);
	m_nTool = ID_PAINT_PEN;
}

//
// ~Painter
//

Painter::~Painter()
{
	if (m_size.cx != 0 && m_size.cy != 0)
	{
		for (int nImage = 0; nImage < eNumOfImages; nImage++)
			delete [] m_pcrMainData[nImage];
	}
}

#ifdef _DEBUG
void Painter::Dump(DumpContext& dc)
{
	if (m_pcrImageData)
	{
		for (int nY = 0; nY < m_size.cy; nY++)
		{
			for (int nX = 0; nX < m_size.cx; nX++)
			{
				_stprintf(dc.m_szBuffer, 
						  _T("%i,%i,%i "), 
						  GetRValue(m_pcrImageData[nY * nX]),
						  GetGValue(m_pcrImageData[nY * nX]),
						  GetBValue(m_pcrImageData[nY * nX]));
				dc.Write();
			}
			_stprintf(dc.m_szBuffer, _T("\n"));
			dc.Write();
		}
	}
}
#endif

//
// SetSize
//

void Painter::SetSize(int cx, int cy)
{
	if (!::IsWindow(m_hWnd))
		return;

	if (0 == cx)
		cx = 1;

	if (0 == cy)
		cy = 1;

	COLORREF* pcrMainData[4];
	int nSize = cx * cy;
	for (int nImage = 0; nImage < eNumOfImages; nImage++)
		pcrMainData[nImage] = new COLORREF[nSize];

	m_bmInfo.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
	m_bmInfo.bmiHeader.biWidth=cx;
	m_bmInfo.bmiHeader.biHeight=cy;
	m_bmInfo.bmiHeader.biPlanes=1;
	m_bmInfo.bmiHeader.biBitCount=24;
	m_bmInfo.bmiHeader.biCompression=BI_RGB;
	m_bmInfo.bmiHeader.biSizeImage=0;
	m_bmInfo.bmiHeader.biXPelsPerMeter=0;
	m_bmInfo.bmiHeader.biYPelsPerMeter=0;
	m_bmInfo.bmiHeader.biClrUsed=0;
	m_bmInfo.bmiHeader.biClrImportant=0;

	// Clear data area
	for (nImage = 0; nImage < eNumOfImages; nImage++)
	{
		for (int nIndex = 0; nIndex < nSize; ++nIndex)
			pcrMainData[nImage][nIndex] = 0xFFFFFFFF;
	}

	// Now copy old data
	if (m_pcrMainData[0])
	{
		int nY, nX;
		for (nImage = 0; nImage < eNumOfImages; nImage++)
		{
			for (nY = 0; nY < cy && nY < m_size.cy; nY++)
			{
				for (nX = 0; nX < cx && nX < m_size.cx; nX++)
					pcrMainData[nImage][nX+nY*cx] = m_pcrMainData[nImage][nX+nY*m_size.cx];
			}
		}
		for (int nImage = 0; nImage < eNumOfImages; nImage++)
			delete [] m_pcrMainData[nImage];
	}
	
	m_size.cx = cx;
	m_size.cy = cy;

	for (nImage = 0; nImage < eNumOfImages; nImage++)
		m_pcrMainData[nImage] = pcrMainData[nImage];

	m_pcrImageData = m_pcrMainData[m_pIconEdit->m_nImageIndex];

	const int nXEdge = GetSystemMetrics(SM_CXEDGE);
	const int nYEdge = GetSystemMetrics(SM_CYEDGE);
	m_rcSelection.SetEmpty();

	m_nVDisplayBoxes = 0; 
	m_nMaxVScrollPos = 0;
	m_nVScrollPos = 0;
	m_nHDisplayBoxes = 0;
	m_nMaxHScrollPos = 0;
	m_nHScrollPos = 0;

	SendMessage(WM_SETREDRAW, FALSE); 
	SetScrollInfo();

	CRect rcClient;
	GetClientRect(rcClient);
	int nWidth = rcClient.Width();
	int nHeight = rcClient.Height();
	int nBoxSizeH = (nWidth-(m_size.cx-1))/m_size.cx;
	int nBoxSizeV = (nHeight-(m_size.cy-1))/m_size.cy;
	m_nBoxSize = min (nBoxSizeH, nBoxSizeV);

	if (m_nBoxSize < eMinBoxSize)
		m_nBoxSize = eMinBoxSize;

	m_nHDisplayBoxes = nWidth / (m_nBoxSize+1);
	m_nMaxHScrollPos = nWidth == m_nBoxSize ? 0 : max(0, m_size.cx);
	m_nHScrollPos = min (m_nHScrollPos, m_nMaxHScrollPos);

	m_nVDisplayBoxes = nHeight / (m_nBoxSize+1); 
	m_nMaxVScrollPos = nHeight == m_nBoxSize ? 0 : max(0, m_size.cy);
	m_nVScrollPos = min (m_nVScrollPos, m_nMaxVScrollPos);

	SetScrollInfo();
	SendMessage(WM_SETREDRAW, TRUE);
	InvalidateRect(NULL, TRUE);
	m_pIconEdit->m_pvWin.InvalidateRect(NULL, TRUE);
}

//
// ClearCanvas
//

void Painter::ClearCanvas()
{
	int nSize = m_size.cx * m_size.cy;
	for (int nIndex = 0; nIndex < nSize; nIndex++)
		m_pcrImageData[nIndex] = 0xFFFFFFFF;
}

//
// CreateWin
//

BOOL Painter::CreateWin(HWND hWndParent)
{
	RegisterWindow(DD_WNDCLASS("Painter"), 
				   CS_VREDRAW|CS_HREDRAW|CS_DBLCLKS, 
				   NULL, 
				   0);
	if (CreateEx(WS_EX_CLIENTEDGE, 
				 _T(""), 
				 WS_CHILD|WS_VISIBLE|WS_HSCROLL|WS_VSCROLL, 
				 0, 
				 0, 
				 0, 
				 0, 
				 hWndParent))
	{
		return TRUE;
	}
	return FALSE;
}

//
// SetScrollInfo
//

void Painter::SetScrollInfo()
{
	SCROLLINFO si;
	si.cbSize = sizeof(si); 
	si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;         
	si.nMin   = 0; 
	si.nMax   = m_nMaxVScrollPos;         
	si.nPage  = m_nVDisplayBoxes+1; 
	si.nPos   = m_nVScrollPos;         
	::SetScrollInfo(m_hWnd, SB_VERT, &si, TRUE);  
	si.nMax   = m_nMaxHScrollPos;         
	si.nPage  = m_nHDisplayBoxes+1; 
	si.nPos   = m_nHScrollPos;         
	::SetScrollInfo(m_hWnd, SB_HORZ, &si, TRUE);  
}

//
// WindowProc
//

LRESULT Painter::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch(nMsg)
		{
		case WM_PAINT:
			if (m_bRedraw)
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcClient;
					GetClientRect(rcClient);
					Draw(hDC, rcClient);
					EndPaint(m_hWnd, &ps);
					return 0;
				}
			}
			break;

		case WM_SETREDRAW:
			m_bRedraw = wParam;
			break;

		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			OnButtonDown(nMsg == WM_LBUTTONDOWN ? eLeftButton : eRightButton);
			break;

		case WM_MOUSEMOVE:
			if (ID_PAINT_LASSO == m_nTool)
			{
				POINT pt;
				pt.x = LOWORD(lParam) / m_nBoxSize;
				pt.y = HIWORD(lParam) / m_nBoxSize;
				if (pt.x >= m_rcSelection.left && 
					pt.x < (m_rcSelection.left+m_rcSelection.right) &&
					pt.y >= m_rcSelection.top && 
					pt.y < (m_rcSelection.top+m_rcSelection.bottom))
				{
					SetCursor(LoadCursor(NULL, IDC_SIZEALL));
					break;
				}
			}
			SetCursor(m_hCursor);
			break;

		case WM_VSCROLL:
			OnVScroll(LOWORD(wParam), HIWORD(wParam));
			break;

		case WM_HSCROLL:
			OnHScroll(LOWORD(wParam), HIWORD(wParam));
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return FWnd::WindowProc(nMsg,wParam,lParam);
}

//
// OnVScroll
//

void Painter::OnVScroll(int nScrollCode, int nPos)
{
	CRect rcClient;
	GetClientRect(rcClient);
	int nYClient = rcClient.Height();
	int nYInc = 0;
	switch (nScrollCode)
	{
	case SB_PAGEUP:
        nYInc = min(-1, -(nYClient/(m_nBoxSize+1))); 
		break;

	case SB_PAGEDOWN:
		nYInc = max(1, (nYClient/(m_nBoxSize+1))); 
		break;

	case SB_LINEUP:
		nYInc = -1;
		break;

	case SB_LINEDOWN:
         nYInc = 1;                  
		break;

	case SB_THUMBTRACK:
        nYInc = nPos - m_nVScrollPos;                  
		break;
	}                                          
	if (nYInc = max(-m_nVScrollPos, min(nYInc, m_nMaxVScrollPos - (m_nVDisplayBoxes+m_nVScrollPos))))
	{ 
		m_nVScrollPos += nYInc;
		ScrollWindowEx(m_hWnd, 
					   0, 
					   -(m_nBoxSize+1) * nYInc,
					   (const RECT*)NULL, 
					   (const RECT*)NULL, 
					   (HRGN)NULL,
					   (LPRECT)NULL,
					   SW_INVALIDATE|SW_ERASE); 
		SCROLLINFO si;
		si.cbSize = sizeof(si);             
		si.fMask  = SIF_POS; 
		si.nPos   = m_nVScrollPos; 
		::SetScrollInfo (m_hWnd, SB_VERT, &si, TRUE); 
		InvalidateRect(NULL, TRUE);
	} 
}

//
// OnHScroll
//

void Painter::OnHScroll(int nScrollCode, int nPos)
{
	CRect rcClient;
	GetClientRect(rcClient);
	int nXClient = rcClient.Width();
	int nXInc = 0;
	switch (nScrollCode)
	{
	case SB_PAGEUP:
        nXInc = min(-1, -(nXClient/(m_nBoxSize+1)));
		break;

	case SB_PAGEDOWN:
		nXInc = max(1, (nXClient/(m_nBoxSize+1))); 
		break;

	case SB_LINEUP:
		nXInc = -1;
		break;

	case SB_LINEDOWN:
        nXInc = 1;                  
		break;

	case SB_THUMBTRACK:
        nXInc = nPos - m_nHScrollPos;
		break;
	}
	if (nXInc = max(-m_nHScrollPos, min(nXInc, m_nMaxHScrollPos - (m_nHDisplayBoxes+m_nHScrollPos))))
	{ 
		m_nHScrollPos += nXInc;
		ScrollWindowEx(m_hWnd, 
					   -(m_nBoxSize+1) * nXInc, 
					   0,
					   (const RECT*)NULL, 
					   (const RECT*)NULL, 
					   (HRGN)NULL,
					   (LPRECT)NULL,
					   SW_INVALIDATE|SW_ERASE); 
		SCROLLINFO si;
		si.cbSize = sizeof(si);             
		si.fMask  = SIF_POS; 
		si.nPos   = m_nHScrollPos; 
		::SetScrollInfo (m_hWnd, SB_HORZ, &si, TRUE); 
		InvalidateRect(NULL, TRUE);
	} 
}

//
// CreateHalftoneBrush
//

static HBRUSH CreateHalftoneBrush(BOOL bColor)
{
	HBRUSH hbrDither;
	// Mono DC and Bitmap for disabled image
	HBITMAP hbmGray = UIUtilities::CreateDitherBitmap(bColor);
	if (hbmGray)
	{
		hbrDither = ::CreatePatternBrush(hbmGray);
		DeleteObject(hbmGray);
	}
	return hbrDither;
}

//
// DrawDragRect
//

static void DrawDragRect(HDC     hDC,
					     CRect&  rc, 
						 SIZE    size,
						 CRect*  prcLast, 
						 SIZE    sizeLast, 
						 HBRUSH  hBrush, 
						 HBRUSH  hBrushLast)
{
	// First, determine the update region and select it
	HRGN hRgnOutside = CreateRectRgnIndirect(&rc);
	assert(hRgnOutside);

	CRect rcCopy(rc);
	rcCopy.Inflate(-size.cx, -size.cy);
	CRect rcTemp = rcCopy;

	rcCopy.Intersect(rcTemp, rc);

	HRGN hRgnInside = CreateRectRgnIndirect(&rcCopy);
	assert(hRgnInside);
	
	HRGN hRgnNew = CreateRectRgn(0, 0, 0, 0);
	assert(hRgnNew);

	int nResult = CombineRgn(hRgnNew, hRgnOutside, hRgnInside, RGN_XOR);
	assert(ERROR != nResult);

	BOOL bDelBrush = FALSE;
	if (NULL == hBrush)
	{
		hBrush = CreateHalftoneBrush(TRUE);
		bDelBrush = TRUE;
	}

	if (NULL == hBrushLast)
		hBrushLast = hBrush;

	HRGN hRgnLast = NULL, hRgnUpdate = NULL;
	if (prcLast)
	{
		// find difference between new region and old region
		hRgnLast = CreateRectRgn(0, 0, 0, 0);
		SetRectRgn(hRgnOutside, 
					prcLast->left, 
					prcLast->top, 
					prcLast->right, 
					prcLast->bottom);
		rcCopy = *prcLast;
		rcCopy.Inflate(-sizeLast.cx,  -sizeLast.cy);
		rcTemp = rcCopy;
		rcCopy.Intersect(rcTemp, *prcLast);
		SetRectRgn(hRgnInside, rcCopy.left, rcCopy.top, rcCopy.right, rcCopy.bottom);
		CombineRgn(hRgnLast, hRgnOutside, hRgnInside, RGN_XOR);				  

		// Only diff them if brushes are the same
		if (hBrush == hBrushLast)
		{
			hRgnUpdate = CreateRectRgn(0, 0, 0, 0);
			CombineRgn(hRgnUpdate, hRgnLast, hRgnNew, RGN_XOR);				  
		}
	}

	HBRUSH hBrushOld;
	if (hBrush != hBrushLast && prcLast)
	{
		// brushes are different -- erase old region first
		SelectClipRgn(hDC, hRgnLast);
		GetClipBox(hDC, &rcCopy);
		hBrushOld = SelectBrush(hDC, hBrushLast);
		PatBlt(hDC,
			   rcCopy.left, 
			   rcCopy.top, 
			   rcCopy.Width(), 
			   rcCopy.Height(), 
			   PATINVERT);
		SelectBrush(hDC, hBrushOld);
		hBrushOld = NULL;
	}

	// Draw into the update/new region
	SelectClipRgn(hDC, hRgnUpdate ? hRgnUpdate : hRgnNew);
	GetClipBox(hDC, &rcCopy);

	hBrushOld = SelectBrush(hDC, hBrush);
	
	PatBlt(hDC,
		   rcCopy.left, 
		   rcCopy.top, 
		   rcCopy.Width(), 
		   rcCopy.Height(), 
		   PATINVERT);

	if (hBrushOld)
		SelectBrush(hDC, hBrushOld);
	
	// cleanup DC
	BOOL bResult;
	if (bDelBrush)
	{
		bResult = DeleteBrush(hBrush);
		assert(bResult);
	}

	SelectClipRgn(hDC, NULL);
	if (hRgnOutside)
	{
		bResult = DeleteObject(hRgnOutside);
		assert(bResult);
	}
	if (hRgnInside)
	{
		bResult = DeleteObject(hRgnInside);
		assert(bResult);
	}
	if (hRgnNew)
	{
		bResult = DeleteObject(hRgnNew);
		assert(bResult);
	}
	if (hRgnLast)
	{
		bResult = DeleteObject(hRgnLast);
		assert(bResult);
	}
	if (hRgnUpdate)
	{
		bResult = DeleteObject(hRgnUpdate);
		assert(bResult);
	}
}

//
// OnPaint
//

void Painter::Draw(HDC hDC, CRect& rcClient)
{
	try
	{
		CRect rcBound(rcClient);
		HDC hDCOff = NULL;
		if (GetGlobals().FlickerFree())
			hDCOff = GetGlobals().FlickerFree()->RequestDC(hDC, rcClient.Width(), rcClient.Height());

		if (NULL == hDCOff)
			hDCOff = hDC;
		else
			rcBound.Offset(-rcBound.left, -rcBound.top);

		FillRect(hDCOff, &rcBound, (HBRUSH)(1+COLOR_BTNFACE)); 

		// Draw border
		assert(m_pcrImageData);
		if (m_pcrImageData)
		{
			const int nXEdge = GetSystemMetrics(SM_CXEDGE);
			const int nYEdge = GetSystemMetrics(SM_CYEDGE);

			rcBound.Inflate(-nXEdge, -nYEdge);

			COLORREF* pcrColor;
			int x,y,scrX,scrY;
			CRect rcBox;
			if (0 != m_size.cx && 0 != m_size.cy)
			{
				// Paint boxes
				HBRUSH hBrushColor;
				pcrColor = m_pcrImageData + (m_nVScrollPos*m_size.cx);
				scrY = rcBound.top;
				for (y = m_nVScrollPos; y < m_size.cy; ++y, scrY += m_nBoxSize+1)
				{
					pcrColor += m_nHScrollPos;
					scrX = rcBound.left;
					for (x = m_nHScrollPos; x < m_size.cx; ++x, scrX += m_nBoxSize+1)
					{
						rcBox.left = scrX;
						rcBox.top = scrY;
						rcBox.right = scrX + m_nBoxSize;
						rcBox.bottom = scrY + m_nBoxSize;
							
						if (0xFFFFFFFF == *pcrColor)
						{
							SetBkColor(hDCOff, RGB(255,255,255));
							SetTextColor(hDCOff, RGB(192,192,192));
							if (GetGlobals().ScreenBrush())
								FillRect(hDCOff, &rcBox, GetGlobals().ScreenBrush());
						}
						else
						{
							hBrushColor = CreateSolidBrush(*pcrColor);
							if (hBrushColor)
							{
								FillRect(hDCOff, &rcBox, hBrushColor);
								DeleteBrush(hBrushColor);
							}
						}
						pcrColor++;
					}
				}
			}
		}

		if (ID_PAINT_LASSO == m_nTool && 0 != m_rcSelection.right && 0 != m_rcSelection.bottom)
		{
			CRect rcSel;
			rcSel.left = m_rcSelection.left * (m_nBoxSize + 1);
			rcSel.top = m_rcSelection.top * (m_nBoxSize + 1);
			rcSel.right = (m_rcSelection.left + m_rcSelection.right) * (m_nBoxSize + 1);
			rcSel.bottom = (m_rcSelection.top + m_rcSelection.bottom) * (m_nBoxSize + 1);
			rcSel.Inflate(-m_nBoxSize / 2, -m_nBoxSize / 2);
			
			SIZE sizeDrag;
			sizeDrag.cx = m_nBoxSize <= 2 ? 1 : m_nBoxSize / 2;
			if (sizeDrag.cx > 4)
				sizeDrag.cx = 4;

			sizeDrag.cy = sizeDrag.cx;

			SetBkColor(hDCOff, 0);
			SetTextColor(hDCOff, 0xFFFFFF);
			rcSel.right += sizeDrag.cx;
			rcSel.bottom += sizeDrag.cy;
			DrawDragRect(hDCOff, rcSel, sizeDrag, NULL, sizeDrag, NULL, NULL);
		}
		if (hDC != hDCOff)
			GetGlobals().FlickerFree()->Paint(hDC, rcClient.left, rcClient.top);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// SetTool
//

void Painter::SetTool(UINT nTool)
{
	if (ID_PAINT_LASSO == m_nTool)
	{
		if (nTool == m_nTool)
			m_rcSelection.bottom = m_rcSelection.right = 0;
		InvalidateRect(NULL, FALSE);
	}

	m_nTool = nTool;

	switch (m_nTool)
	{
	case ID_PAINT_LASSO:
		m_hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_MARKER));
		if (m_rcSelection.right != 0 && m_rcSelection.bottom != 0)
			InvalidateRect(NULL, FALSE);
		break;

	case ID_PAINT_PEN:
	case ID_PAINT_PATTERN:
		m_hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_PEN));
		break;

	case ID_PAINT_PIPET:
		m_hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_PIPET));
		break;

	case ID_PAINT_FILL:
		m_hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_BUCKET));
		break;

	case ID_PAINT_LINE:
	case ID_PAINT_RECT0:
	case ID_PAINT_RECT1:
	case ID_PAINT_RECT2:
	case ID_PAINT_CIRC0:
	case ID_PAINT_CIRC1:
	case ID_PAINT_CIRC2:
		m_hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_HAIR));
		break;
	}
	
	POINT pt;
	GetCursorPos(&pt);
	CRect rcWin;
	GetWindowRect(rcWin);
	if (PtInRect(&rcWin, pt))
		SetCursor(m_hCursor);
}

//
// ConvertToImageBorders
//

static Painter::IMAGEBORDERS ConvertToImageBorders (int nImageBorders) 
{
	switch (nImageBorders)
	{
	case 0:
		return Painter::eBorder;
	case 1:
		return Painter::eBorderFill;
	case 2:
		return Painter::eFill;
	}
	return Painter::eBorder;
}

//
// DoTimerScrolling
//

void Painter::DoTimerScrolling()
{
	POINT pt;
	GetCursorPos(&pt);
	ScreenToClient(pt);
	CRect rcClient;
	GetClientRect(rcClient);
	if (pt.x > rcClient.right)
		OnHScroll(SB_LINEDOWN, 0);
	else if (pt.x < rcClient.left)
		OnHScroll(SB_LINEUP, 0);
	else if (pt.y < rcClient.top)
		OnVScroll(SB_LINEUP, 0);
	else if (pt.y > rcClient.bottom)
		OnVScroll(SB_LINEDOWN, 0);
}

//
// OnButtonDown
//

void Painter::OnButtonDown(MOUSEBUTTONS eButton)
{
	POINT ptInit;
	POINT ptNew;
	BOOL bCancelled = FALSE;
	
	m_pIconEdit->UndoPush();
	m_pIconEdit->m_bDirty[m_pIconEdit->m_nImageIndex] = TRUE;

	// eliminates reentry due to modal loop
	static BOOL bInside = FALSE; 
	if (bInside)
		return;
	bInside = TRUE;

	int nCurTool = m_nTool;
				
	COLORREF crColor;
	if (Painter::eLeftButton == eButton)
		crColor = m_pIconEdit->m_crForeground;
	else
		crColor = m_pIconEdit->m_crBackground;

	GetCursorPos(&ptInit);
	ScreenToClient(ptInit);

	POINT pt, ptNew2, ptInit2;
	pt.x = ptInit.x / (m_nBoxSize+1);
	pt.y = ptInit.y / (m_nBoxSize+1);
	ptInit2 = pt;

	switch(nCurTool)
	{
	case ID_PAINT_PEN:
		OnPenTrack(pt, crColor);
		break;

	case ID_PAINT_PATTERN:
		OnPatternTrack(pt, crColor);
		break;

	case ID_PAINT_PIPET:
		OnPipetTrack(pt, eButton);
		break;

	case ID_PAINT_LINE:
		OnLineInit(pt, crColor);
		break;

	case ID_PAINT_RECT0:
		OnRectInit(pt, eBorder); 
		break;

	case ID_PAINT_RECT1:
		OnRectInit(pt, eBorderFill); 
		break;
	
	case ID_PAINT_RECT2:
		OnRectInit(pt, eFill); 
		break;

	case ID_PAINT_CIRC0:
		OnCircleInit(pt, eBorder, crColor);
		break;

	case ID_PAINT_CIRC1:
		OnCircleInit(pt, eBorderFill, crColor);
		break;

	case ID_PAINT_CIRC2:
		OnCircleInit(pt, eFill, crColor);
		break;

	case ID_PAINT_FILL:
		FloodFill(pt.x + m_nHScrollPos, 
				  pt.y + m_nVScrollPos, 
				  Painter::eLeftButton == eButton ? m_pIconEdit->m_crForeground : m_pIconEdit->m_crBackground);
		bInside = FALSE;
		return;

	case ID_PAINT_LASSO:
		if ((m_rcSelection.right != 0 && m_rcSelection.bottom != 0) &&
			pt.x >= m_rcSelection.left && 
			pt.y >= m_rcSelection.top &&
			pt.x < m_rcSelection.left + m_rcSelection.right &&
			pt.y < m_rcSelection.top + m_rcSelection.bottom)
		{
			nCurTool = ID_PAINT_MOVER;
			OnMoverInit(pt);
		}
		else
			OnSelectInit(pt);
		break;

	default:
		bInside = FALSE;
		return;
	}

	ptNew2 = pt;
	SetCapture(m_hWnd);

	MSG msg;
	while (1)
	{
		if (GetCapture() != m_hWnd)
			break;

		GetMessage(&msg, NULL, 0, 0);

		switch (msg.message)
		{
		// handle movement/accept messages
		case WM_CANCELMODE:
			bCancelled = TRUE;
			goto ExitLoop;

		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
		case WM_LBUTTONUP:
		case WM_RBUTTONUP:
			goto ExitLoop;

		case WM_KEYDOWN:
		case WM_KEYUP:
			if (VK_ESCAPE == msg.wParam)
			{
				bCancelled = TRUE;
				goto ExitLoop;
			}
			if (VK_CONTROL == msg.wParam || VK_SHIFT == msg.wParam)
			{
				GetCursorPos(&ptNew);
				ScreenToClient(ptNew);
				goto ForceMovePoint;
			}
			break;

		case WM_MOUSEMOVE:
			{
				ptNew.x = (int)(short)LOWORD(msg.lParam);
				ptNew.y = (int)(short)HIWORD(msg.lParam);

ForceMovePoint:
				ptNew2.x = ptNew.x / (m_nBoxSize+1);
				ptNew2.y = ptNew.y / (m_nBoxSize+1);

				// this needed for line vert/horz locking
				BOOL bChange = ptNew2.x != pt.x || ptNew2.y != pt.y; 

				switch(nCurTool)
				{
				case ID_PAINT_PEN:
					if (bChange)
						OnPenTrack(ptNew2, crColor);
					break;

				case ID_PAINT_PATTERN:
					if (bChange)
						OnPatternTrack(pt, crColor);
					break;

				case ID_PAINT_PIPET:
					if (bChange)
						OnPipetTrack(ptNew2, eButton);
					break;

				case ID_PAINT_LINE:
					{
						if (0x8000 & GetKeyState(VK_CONTROL))
							ptNew2.y = ptInit.y;

						if (0x8000 & GetKeyState(VK_SHIFT))
							ptNew2.x = ptInit.x;

						if (ptNew2.x != pt.x || ptNew2.y != pt.y)
							OnLineTrack(ptNew2, crColor);
					}
					break;

				case ID_PAINT_RECT0:
				case ID_PAINT_RECT1:
				case ID_PAINT_RECT2:
					if (0x8000 & GetKeyState(VK_CONTROL))
					{
						if (ptNew2.x - ptInit2.x < ptNew2.y - ptInit.y)
							ptNew2.y = ptInit2.y + ptNew2.x - ptInit.x;
						else
							ptNew2.x = ptInit2.x + ptNew2.y-ptInit.y;
					}
					if (ptNew2.x != pt.x || ptNew2.y != pt.y)
						OnRectTrack(ptNew2, ConvertToImageBorders(nCurTool-ID_PAINT_RECT0)); 
					break;

				case ID_PAINT_CIRC0:
				case ID_PAINT_CIRC1:
				case ID_PAINT_CIRC2:
					if (0x8000 & GetKeyState(VK_CONTROL))
					{
						if (ptNew2.x - ptInit2.x < ptNew2.y - ptInit.y)
							ptNew2.y = ptInit2.y + (ptNew2.x - ptInit.x);
						else
							ptNew2.x = ptInit2.x + (ptNew2.y - ptInit.y);
					}
					if (ptNew2.x != pt.x || ptNew2.y != pt.y)
						OnCircleTrack(ptNew2, ConvertToImageBorders(nCurTool-ID_PAINT_CIRC0), crColor); 
					break;

				case ID_PAINT_LASSO:
					if (ptNew2.x < 0) 
						ptNew2.x = 0;

					if (ptNew2.y < 0) 
						ptNew2.y = 0;

					if (ptNew2.x >= m_size.cx) 
						ptNew2.x = m_size.cx - 1;

					if (ptNew2.y >= m_size.cy) 
						ptNew2.y = m_size.cy - 1;

					if (ptNew2.x != pt.x || ptNew2.y != pt.y)
						OnSelectTrack(ptNew2);
					break;

				case ID_PAINT_MOVER:
					OnMoverTrack(ptNew2);
					break;
				}
				pt=ptNew2;
			}
			break;
		// just diptatch rest of the messages
		default:
			DispatchMessage(&msg);
			break;
		}
	}

ExitLoop:
	ReleaseCapture();

	switch (nCurTool)
	{
	case ID_PAINT_LINE:
		OnLineCleanup(ptNew2, crColor, bCancelled);
		break;
	
	case ID_PAINT_RECT0:
		OnRectCleanup(ptNew2, eBorder, bCancelled); 
		break;

	case ID_PAINT_RECT1:
		OnRectCleanup(ptNew2, eBorderFill, bCancelled); 
		break;
	
	case ID_PAINT_RECT2:
		OnRectCleanup(ptNew2, eFill, bCancelled); 
		break;
	
	case ID_PAINT_CIRC0:
		OnCircleCleanup(ptNew2, eBorder, bCancelled, crColor); 
		break;
	
	case ID_PAINT_CIRC1:
		OnCircleCleanup(ptNew2, eBorderFill, bCancelled, crColor); 
		break;
	
	case ID_PAINT_CIRC2:
		OnCircleCleanup(ptNew2, eFill, bCancelled, crColor); 
		break;

	case ID_PAINT_MOVER:
		OnMoverCleanup(ptNew2, bCancelled);
		break;
	}
	bInside = FALSE;
}

//
// PutPixel
//

void Painter::PutPixel(int x, int y, COLORREF crColor)
{
	assert(m_pcrImageData);
	m_pcrImageData[m_size.cx * (y+m_nVScrollPos) + (x+m_nHScrollPos)] = crColor;
	RECT rcPixel;
	rcPixel.left = CIconEdit::eGridSpacing + x * (m_nBoxSize+1);
	rcPixel.top = CIconEdit::eGridSpacing + y * (m_nBoxSize+1);
	rcPixel.right = rcPixel.left + m_nBoxSize;
	rcPixel.bottom = rcPixel.top + m_nBoxSize;

	HDC hDC = GetDC(m_hWnd);
	assert(hDC);

	if (0xFFFFFFFF == crColor)
	{
		SetBkColor(hDC, RGB(255,255,255));
		SetTextColor(hDC, RGB(192,192,192));
		FillRect(hDC, &rcPixel, GetGlobals().ScreenBrush());
	}
	else
	{
		SetBkColor(hDC, crColor);
		ExtTextOut(hDC, 0, 0, ETO_OPAQUE, &rcPixel, NULL, 0, NULL);
	}

	ReleaseDC(m_hWnd, hDC);
}

//
// ShiftIcon
//

void Painter::ShiftIcon(UINT nMsg)
{
	assert(m_pcrImageData);
	const int nSizeOfColorRef = sizeof(COLORREF);
	COLORREF* pcrTemp;
	int       nIndex;
	int       nTemp;

	switch (nMsg)
	{
	case ID_SHIFT_LEFT:
		nTemp = nSizeOfColorRef * (m_size.cx - 1);
		for (nIndex = 0; nIndex < m_size.cy; nIndex++)
		{
			pcrTemp = &m_pcrImageData[nIndex * m_size.cx];
			memcpy(pcrTemp, pcrTemp+1, nTemp);
		}
		
		pcrTemp = &m_pcrImageData[m_size.cx - 1];

		for (nIndex = 0; nIndex < m_size.cy; nIndex++)
			pcrTemp[nIndex * m_size.cx] = m_pIconEdit->m_crBackground;
		break;

	case ID_SHIFT_RIGHT:
		nTemp = nSizeOfColorRef * (m_size.cx-1);
		
		for (nIndex = 0; nIndex < m_size.cy; nIndex++)
		{
			pcrTemp = &m_pcrImageData[nIndex * m_size.cx];
			memmove(pcrTemp+1, pcrTemp, nTemp);
		}
		
		for (nIndex = 0; nIndex < m_size.cy; nIndex++)
			m_pcrImageData[nIndex * m_size.cx] = m_pIconEdit->m_crBackground;
		break;

	case ID_SHIFT_UP:
		nTemp = m_size.cx * (m_size.cy-1);
		
		memcpy(&m_pcrImageData[0], &m_pcrImageData[m_size.cx], nSizeOfColorRef * nTemp);

		pcrTemp = &m_pcrImageData[nTemp];

		for (nIndex = 0; nIndex < m_size.cx; ++nIndex)
			pcrTemp[nIndex] = m_pIconEdit->m_crBackground;
		break;

	case ID_SHIFT_DOWN:
		memmove(&m_pcrImageData[m_size.cx], &m_pcrImageData[0], nSizeOfColorRef * (m_size.cx * (m_size.cy-1)));

		for (nIndex = 0; nIndex < m_size.cx; nIndex++)
			m_pcrImageData[nIndex] = m_pIconEdit->m_crBackground;
		break;
	}
	m_pIconEdit->SetDirty();
	RefreshImage();
}

// 
// OnPenTrack
//

void Painter::OnPenTrack(POINT p, COLORREF crColor)
{
	if (p.x >= m_size.cx || p.x < 0 || p.y >= m_size.cy || p.y < 0)
		return;

	PutPixel(p.x, p.y, crColor);
	m_pIconEdit->m_pvWin.UpdateDibSection();
}

//
// OnPenTrack
//

void Painter::OnPenTrack(int x, int y, COLORREF crColor)
{
	if (x >= m_size.cx || x < 0 || y >= m_size.cy || y < 0)
		return;
	PutPixel(x, y, crColor);
}

//
// OnPatternTrack
//

void Painter::OnPatternTrack(POINT pt, COLORREF crColor)
{
	if (pt.x >= m_size.cx || pt.x < 0 || pt.y >= m_size.cy || pt.y < 0)
		return;

	BOOL bCtrl = (GetKeyState(VK_CONTROL)&0x8000) != 0;
	if (((pt.x+pt.y)&1) != bCtrl) 
	{
		OnPenTrack(pt.x,   pt.y,   crColor);
		OnPenTrack(pt.x-1, pt.y-1, crColor);
		OnPenTrack(pt.x+1, pt.y-1, crColor);
		OnPenTrack(pt.x-1, pt.y+1, crColor);
		OnPenTrack(pt.x+1, pt.y+1, crColor);
	}
	else
	{
		OnPenTrack(pt. x-1, pt.y,   crColor);
		OnPenTrack(pt. x+1, pt.y,   crColor);
		OnPenTrack(pt. x,   pt.y-1, crColor);
		OnPenTrack(pt. x,   pt.y+1, crColor);
	}
	m_pIconEdit->m_pvWin.UpdateDibSection();
}

//
// OnPipetTrack
//

void Painter::OnPipetTrack(POINT pt, MOUSEBUTTONS eButton)
{
	assert(m_pcrImageData);
	pt.x += m_nHScrollPos;
	pt.x += m_nVScrollPos;

	if (pt.x >= m_size.cx || pt.x < 0 || pt.y >= m_size.cy || pt.y < 0)
		return;

	COLORREF crNew = m_pcrImageData[pt.x + (pt.y * m_size.cx)];
	if (Painter::eLeftButton == eButton)
		m_pIconEdit->m_crForeground = crNew;
	else
		m_pIconEdit->m_crBackground = crNew;
	m_pIconEdit->UpdateColorPreview();
}

//
// StdInit
//

void Painter::StdInit(POINT pt, COLORREF crColor)
{
	int nSize = m_size.cx * m_size.cy;
	m_pcrBackup = new COLORREF[nSize];
	if (NULL == m_pcrBackup)
		return;

	assert(m_pcrImageData);
	memcpy(m_pcrBackup, m_pcrImageData, nSize * sizeof(COLORREF));
	pt.x += m_nHScrollPos;
	pt.y += m_nVScrollPos;
	m_ptInit = pt;
}

//
// OnLineInit
//

void Painter::OnLineInit(POINT pt, COLORREF crColor)
{
	StdInit(pt, crColor);
	OnLineTrack(pt, crColor);
}

//
// PainterLineDDAProc
//
// x-coordinate of point being evaluated  
// y-coordinate of point being evaluated  
// address of application-defined data 
//

void CALLBACK Painter::PainterLineDDAProc(int X, int Y, LPARAM pData) 
{
	Painter* pPainter = reinterpret_cast<Painter*>(pData);
	assert (pPainter);
	POINT pt;
	pt.x=X;
	pt.y=Y;
	pPainter->OnLineTrackPutPoint(pt);
}

//
// OnLineTrackPutPoint
//

void Painter::OnLineTrackPutPoint(POINT pt)
{
	if (pt.x >= m_size.cx || pt.x < 0 || pt.y >= m_size.cy || pt.y < 0)
		return;

	assert(m_pcrImageData);
	m_pcrImageData[pt.x + m_size.cx * pt.y] = m_crTrack;
}

//
// OnLineTrack
//

void Painter::OnLineTrack(POINT pt, COLORREF crColor)
{
	if (0 == m_pcrBackup)
		return;

	assert(m_pcrImageData);
	memcpy(m_pcrImageData, m_pcrBackup, m_size.cx*m_size.cy*sizeof(COLORREF));

	// Now draw line on top
	m_crTrack = crColor;

	pt.x += m_nHScrollPos;
	pt.y += m_nVScrollPos;
	
	LineDDA(m_ptInit.x,	// x-coordinate of line's starting point 
			m_ptInit.y,	// y-coordinate of line's starting point 
			pt.x,	    // x-coordinate of line's ending point 
			pt.y,	    // y-coordinate of line's ending point 
			PainterLineDDAProc, // address of application-defined callback function  
			(LPARAM)this); // address of application-defined data 
	
	// put the last point
	assert(m_pcrImageData);
	if (!(pt.x >= m_size.cx || pt.x < 0 || pt.y >= m_size.cy || pt.y < 0))
		m_pcrImageData[pt.x + m_size.cx * pt.y] = m_crTrack;

	RefreshImage();
}

//
// OnLineCleanup
//

void Painter::OnLineCleanup(POINT pt, COLORREF crColor, BOOL bCancelled)
{
	if (NULL == m_pcrBackup)
		return;

	if (bCancelled)
	{
		assert(m_pcrImageData);
		memcpy(m_pcrImageData, m_pcrBackup, m_size.cx*m_size.cy*sizeof(COLORREF));
		
		//
		// blast on screen
		//

		HDC hDCWin = GetDC(m_hWnd);
		assert(hDCWin);
		if (hDCWin)
		{
			CRect rcClient;
			GetClientRect(rcClient);
			Draw(hDCWin, rcClient);
			ReleaseDC(m_hWnd, hDCWin);
		}
	}
	else
		OnLineTrack(pt, crColor);
	
	delete [] m_pcrBackup;

	m_pIconEdit->m_pvWin.UpdateDibSection();
}

//
// OnRectInit
//

void Painter::OnRectInit(POINT pt, IMAGEBORDERS eMode)
{
	StdInit(pt, 0);
	OnRectTrack(pt, eMode);
}

//
// OnRectTrack
//

void Painter::OnRectTrack(POINT pt, IMAGEBORDERS eMode)
{
	if (NULL == m_pcrBackup)
		return;

	assert(m_pcrImageData);
	memcpy(m_pcrImageData, m_pcrBackup, m_size.cx * m_size.cy * sizeof(COLORREF));

	pt.x += m_nHScrollPos;
	pt.y += m_nVScrollPos;

	CRect rc;
	// normalized rectangle
	if (m_ptInit.x < pt.x)
	{
		rc.left = m_ptInit.x; 
		rc.right = pt.x;
	}
	else
	{
		rc.left = pt.x;
		rc.right = m_ptInit.x;
	}
	if (m_ptInit.y < pt.y)
	{
		rc.top = m_ptInit.y; 
		rc.bottom = pt.y;
	}
	else
	{
		rc.top = pt.y;
		rc.bottom = m_ptInit.y;
	}

	COLORREF crColor;
	int x, y;
	if (eFill != eMode)
	{ 
		// Draw frame
		if (eBorderFill == eMode)
			crColor = m_pIconEdit->m_crBackground;	
		else
			crColor = m_pIconEdit->m_crForeground;	

		// top and bottom lines
		for (x = rc.left; x <= rc.right; ++x) 
		{
			if (!(x >= m_size.cx || x < 0))
			{
				if (!(rc.top >= m_size.cy || rc.top < 0))
					m_pcrImageData[x + m_size.cx * rc.top] = crColor;

				if (!(rc.bottom >= m_size.cy || rc.bottom < 0))
					m_pcrImageData[x + m_size.cx * rc.bottom] = crColor;
			}
		}

		// left and right lines
		for (y = rc.top; y <= rc.bottom; ++y) 
		{
			if (!(y >= m_size.cy || y < 0))
			{
				if (!(rc.left >= m_size.cx || rc.left < 0))
					m_pcrImageData[rc.left + m_size.cx * y] = crColor;
				if (!(rc.right >= m_size.cx || rc.right < 0))
					m_pcrImageData[rc.right + m_size.cx * y] = crColor;
			}
		}
	}

	if (eBorder != eMode)
	{
		// Draw inside
		crColor = m_pIconEdit->m_crForeground;
		if (eBorderFill == eMode)
			rc.Inflate(-1, -1);
		
		// Make sure rect not empty
		if (rc.Height() >= 0 && rc.Width() >= 0) 
		{
			for (y = rc.top; y <= rc.bottom; ++y)
			{
				for (x = rc.left; x <= rc.right; ++x)
				{
					if (!(x >= m_size.cx || x < 0 || y >= m_size.cy || y < 0))
						m_pcrImageData[x + m_size.cx * y] = crColor;
				}
			}
		}
	}
	RefreshImage();
}

//
// OnRectCleanup
//

void Painter::OnRectCleanup(POINT pt, IMAGEBORDERS eMode, BOOL bCancelled)
{
	if (NULL == m_pcrBackup)
		return;

	if (bCancelled)
	{
		assert(m_pcrImageData);
		memcpy(m_pcrImageData, m_pcrBackup, m_size.cx * m_size.cy * sizeof(COLORREF));
		HDC hDCWin = GetDC(m_hWnd);
		assert(hDCWin);
		if (hDCWin)
		{
			CRect rcClient;
			Draw(hDCWin, rcClient);
			ReleaseDC(m_hWnd, hDCWin);
		}
	}
	else
		OnRectTrack(pt, eMode);

	delete [] m_pcrBackup;
	m_pIconEdit->m_pvWin.UpdateDibSection();
}

//
// CircleProc
//

void CALLBACK Painter::CircleProc(int X, int Y, LPARAM pData)
{
	Painter* pPainter = reinterpret_cast<Painter*>(pData);
	assert (pPainter);
	POINT pt;
	pt.x=X;
	pt.y=Y;
	pPainter->OnLineTrackPutPoint(pt);
}

//
// CircleLineProc
//

void CALLBACK Painter::CircleLineProc(int X, int Y, int nLen, LPARAM pData)
{
	Painter* pPainter = reinterpret_cast<Painter*>(pData);
	assert (pPainter);

	POINT pt;
	pt.x = X;
	pt.y = Y;
	for (int nIndex = 0; nIndex < nLen; ++nIndex)
	{
		pPainter->OnLineTrackPutPoint(pt);
		++pt.x;
	}
}

//
// OnCircleTrack
//

void Painter::OnCircleTrack(POINT pt, IMAGEBORDERS eMode, COLORREF crColor)
{
	assert(m_pcrBackup);
	if (NULL == m_pcrBackup)
		return;

	assert(m_pcrImageData);
	memcpy(m_pcrImageData, m_pcrBackup, m_size.cx*m_size.cy*sizeof(COLORREF));

	pt.x += m_nHScrollPos;
	pt.y += m_nVScrollPos;

	CRect rc;
	// Normalized rectangle
	if (m_ptInit.x < pt.x)
	{
		rc.left = m_ptInit.x; 
		rc.right = pt.x;
	}
	else
	{
		rc.left = pt.x;
		rc.right = m_ptInit.x;
	}

	if (m_ptInit.y < pt.y)
	{
		rc.top = m_ptInit.y; 
		rc.bottom = pt.y;
	}
	else
	{
		rc.top = pt.y;
		rc.bottom = m_ptInit.y;
	}
	
	// now use r
	switch (eMode)
	{
	case eBorder:
		m_crTrack = m_pIconEdit->m_crForeground;
		CircleDDA(rc.left, 
			      rc.top, 
				  rc.right, 
				  rc.bottom, 
				  &CircleProc, 
				  (LPARAM)this);
		break;

	case eBorderFill:
		m_crTrack = m_pIconEdit->m_crForeground;
		CircleSolidDDA(rc.left, 
					   rc.top, 
					   rc.right, 
					   rc.bottom, 
					   &CircleLineProc, 
					   (LPARAM)this);
		m_crTrack = m_pIconEdit->m_crBackground;
		CircleDDA(rc.left, 
				  rc.top, 
				  rc.right, 
				  rc.bottom, 
				  &CircleProc, 
				  (LPARAM)this);
		break;

	case eFill:
		m_crTrack = m_pIconEdit->m_crForeground;
		CircleSolidDDA(rc.left, 
					   rc.top, 
					   rc.right, 
					   rc.bottom, 
					   &CircleLineProc, 
					   (LPARAM)this);
		break;
	}
	RefreshImage();
}

//
// OnCircleCleanup
//

void Painter::OnCircleCleanup(POINT		   pt, 
							  IMAGEBORDERS eMode, 
							  BOOL		   bCancelled, 
							  COLORREF	   crColor)
{
	if (NULL == m_pcrBackup)
		return;

	if (bCancelled)
	{
		assert(m_pcrImageData);
		memcpy(m_pcrImageData, m_pcrBackup, m_size.cx * m_size.cy * sizeof(COLORREF));
		HDC hDCWin = GetDC(m_hWnd);
		assert(hDCWin);
		if (hDCWin)
		{
			CRect rcClient;
			Draw(hDCWin, rcClient);
			ReleaseDC(m_hWnd, hDCWin);
		}
	}
	else
		OnCircleTrack(pt, eMode, crColor);

	delete [] m_pcrBackup;
	m_pIconEdit->m_pvWin.UpdateDibSection();
}

//
// Struct SegmentStack 
//

struct SegmentStack
{
	//
	// Segment
	//

	struct Segment 
	{
		short y;
		short xl;
		short xr;
		short dy;
	};

	enum 
	{
		eMaxStackDepth = 500
	};

	SegmentStack(const int& nLimitY)
	{
		m_nLimitY = nLimitY;
		m_psStack = m_sStack;
	}

	void Push(int nY, int nXL, int nXR, int nDY)
	{
		try
		{
			if (m_psStack < m_sStack + eMaxStackDepth && nY + 1 >= 0 && nY + nDY < m_nLimitY)
			{
				m_psStack->y = nY; 
				m_psStack->xl = nXL; 
				m_psStack->xr = nXR; 
				m_psStack->dy = nDY; 
				m_psStack++;
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}

	BOOL Pop(int& nY, int& nXL, int& nXR, int& nDY)
	{
		try
		{
			m_psStack--;
			nY = m_psStack->y + (nDY = m_psStack->dy); 
			nXL = m_psStack->xl; 
			nXR = m_psStack->xr;
			return m_psStack >= m_sStack;
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		return FALSE;
	}

private:
	Segment  m_sStack[eMaxStackDepth];
	Segment* m_psStack;
	int      m_nLimitY;
};

//
// FloodFill
//

void Painter::FloodFill(int x, int y, COLORREF crNew)
{
	try
	{
		assert(m_pcrImageData);
		// same color dont floodfill
		COLORREF crImageData = m_pcrImageData[x+m_size.cx*y];
		if (crImageData == crNew)
			return; 

		SegmentStack theSegmentStack(m_size.cy);
		theSegmentStack.Push(y, x, x, 1);
		theSegmentStack.Push(y+1, x, x, -1);

		int l, x1, x2, dy;
		while (theSegmentStack.Pop(y, x1, x2, dy))
		{
			try
			{
				if (y < 0)
					goto Skip;

				// Pop segment off and fill neighboring scan line
				for (x = x1; x >= 0 && crImageData == m_pcrImageData[x + m_size.cx * y]; x--)
					m_pcrImageData[x + m_size.cx * y] = crNew;

				if (x >= x1) 
					goto Skip;

				l = x + 1;
				// leak on left ?
				if (l < x1) 
					theSegmentStack.Push(y, l, x1 - 1, -dy); 

				x = x1 + 1;
				do
				{
					try
					{
						for (; x < m_size.cx && crImageData == m_pcrImageData[x + m_size.cx * y]; x++)
							m_pcrImageData[x + m_size.cx * y] = crNew;

						theSegmentStack.Push(y, l, x-1, dy);

						// might be a leak on the right
						if (x > x2+1) 
							theSegmentStack.Push(y, x2+1, x-1, -dy); 
Skip:		
						for (++x; x <= x2 && crImageData != m_pcrImageData[x + m_size.cx * y]; x++);

						l = x;
					}
					CATCH
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
				} 
				while (x <= x2);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
		RefreshImage();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnSelectTrack
//

void Painter::OnSelectTrack(POINT pt)
{
	m_rcSelection.left = min(m_ptInit.x, pt.x);
	m_rcSelection.top = min(m_ptInit.y, pt.y);
	m_rcSelection.right = abs(m_ptInit.x - pt.x) + 1;
	m_rcSelection.bottom = abs(m_ptInit.y - pt.y) + 1;
	HDC hDCWin = GetDC(m_hWnd);
	assert(hDCWin);
	if (hDCWin)
	{
		CRect rcClient;
		GetClientRect(rcClient);
		Draw(hDCWin, rcClient);
		ReleaseDC(m_hWnd, hDCWin);
	}
}

//
// OnMoverTrack
//

void Painter::OnMoverTrack(POINT pt)
{
	// move area in selection rect by p-m_ptInit
	if (NULL == m_pcrBackup)
		return;

	assert(m_pcrImageData);
	memcpy(m_pcrImageData, m_pcrBackup, m_size.cx*m_size.cy*sizeof(COLORREF));

	int x = m_rcPrevSelection.left + m_nHScrollPos;
	int y = m_rcPrevSelection.top + m_nVScrollPos;
	const int nHeight = y + m_rcPrevSelection.bottom;
	const int nWidth = x + m_rcPrevSelection.right;
	COLORREF rcBackColor = m_pIconEdit->m_crBackground;
	if (0 == (GetKeyState(VK_CONTROL)&0x8000)) 
	{
		// If Control Key is not pressed set old area to rcBackColor
		for (; y < nHeight; y++)
			for (x = m_rcPrevSelection.left + m_nHScrollPos; x < nWidth; x++)
				m_pcrImageData[x+y*m_size.cx] = rcBackColor;	
	}

	int nDestX, nDestY;
	for (y = m_rcPrevSelection.top + m_nVScrollPos; y < nHeight; y++)
	{
		for (x = m_rcPrevSelection.left + m_nHScrollPos; x < nWidth; x++)
		{
			nDestX = x + pt.x - m_ptInit.x;
			nDestY = y + pt.y - m_ptInit.y;
			if (nDestX >= 0 && nDestX < m_size.cx && nDestY >= 0 && nDestY < m_size.cy)
			{
				if (x >= 0 && x < m_size.cx && y >= 0 && y < m_size.cy)
					m_pcrImageData[nDestX+nDestY*m_size.cx] = m_pcrBackup[x+y*m_size.cx];
			}
		}
	}

	m_rcSelection = m_rcPrevSelection;
	m_rcSelection.left += pt.x - m_ptInit.x;
	m_rcSelection.top += pt.y - m_ptInit.y;
	// output result
	RefreshImage();
}

//
// OnMoverCleanup
//

void Painter::OnMoverCleanup(POINT pt, BOOL bCancelled)
{
	if (bCancelled)
	{
		m_rcSelection = m_rcPrevSelection;
		memcpy(m_pcrImageData, m_pcrBackup, m_size.cx*m_size.cy*sizeof(COLORREF));
		RefreshImage();
	}
	m_rcSelection.right = 0;
	m_rcSelection.bottom = 0;
}

//
// RefreshImage
//

void Painter::RefreshImage()
{
	HDC hDCWin = GetDC(m_hWnd);
	assert(hDCWin);
	if (hDCWin)
	{
		CRect rcClient;
		GetClientRect(rcClient);
		Draw(hDCWin, rcClient);
		ReleaseDC(m_hWnd, hDCWin);
	}
	m_pIconEdit->m_pvWin.UpdateDibSection();	
}

//
// Clear
//

void UndoManager::Clear()
{
	int nElem = m_faUndoList.GetSize();
	for (int nIndex = 0; nIndex < nElem; ++nIndex)
		delete reinterpret_cast<UndoItem*>(m_faUndoList.GetAt(nIndex));
	m_faUndoList.RemoveAll();
}

//
// Push
//

void UndoManager::Push(int cx, int cy, LPBYTE pData, DWORD dwSize)
{
	UndoItem* pNewItem = new UndoItem(pData, dwSize, cx, cy);
	assert(pNewItem);
	m_faUndoList.Add(pNewItem);
}

//
// PopQuery
//

BOOL UndoManager::PopQuery(int& nCx, int& nCy)
{
	int nElem = m_faUndoList.GetSize();
	if (0 == nElem)
		return FALSE;

	UndoItem* pItem = reinterpret_cast<UndoItem*>(m_faUndoList.GetAt(nElem-1));
	assert(pItem);
	pItem->GetSize(nCx, nCy);
	return TRUE;
}

//
// Pop
//

void UndoManager::Pop(LPBYTE pDestData)
{
	int nElem = m_faUndoList.GetSize();
	if (0 == nElem)
		return;
	--nElem;

	UndoItem* pItem = reinterpret_cast<UndoItem*>(m_faUndoList.GetAt(nElem));
	assert(pItem);
 	if (pItem)
		pItem->GetData(pDestData);

	// Now remove from array
	m_faUndoList.RemoveAt(nElem);
}

//
// CIconEdit
//

CIconEdit::CIconEdit()
	: m_hWndSnapShotParent(NULL)
{
	m_szMsgTitle.LoadString(IDS_ACTIVEBAR);
	m_szClassName = DD_WNDCLASS("IconEdit");
	RegisterWindow(m_szClassName,
				   CS_VREDRAW|CS_HREDRAW|CS_DBLCLKS,
				   (HBRUSH)(1+COLOR_BTNFACE),
				   ::LoadCursor(NULL, IDC_ARROW));
	Init();
	m_pvColor.m_pIconEdit = this;
	m_pntEdit.m_pIconEdit = this;
	m_pvWin.m_pIconEdit = this;
}

//
// Init
//

void CIconEdit::Init()
{
	// screen
	m_crBackground = 0xFFFFFFFF; 
	m_crForeground = 0;
	m_nImageIndex = 0;
	memset (&m_bDirty[0], 0, sizeof(BOOL)*eNumOfStyles);
}

//
// CreateWin
//

BOOL CIconEdit::CreateWin(HWND hWndParent, HWND hWndSnapShotParent, int nOffsetButtons)
{
	m_hWndSnapShotParent = hWndSnapShotParent;
	CRect rcDefault;
	::GetClientRect(hWndParent, &rcDefault);
	FWnd::Create(_T(""),
				 WS_CHILD|WS_VISIBLE|WS_GROUP|WS_TABSTOP,
				 rcDefault.left,
				 rcDefault.top,
				 rcDefault.Width()-nOffsetButtons,
				 rcDefault.Height(),
   				 hWndParent,
				 NULL,
				 NULL);
	m_pntEdit.SetSize(Painter::eDefaultPainterSize, Painter::eDefaultPainterSize);
	return (m_hWnd ? TRUE : FALSE);
}

//
// WindowProc
//

LRESULT CIconEdit::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch(nMsg)
		{
		case WM_VSCROLL:
			{
				int nScrollCode = (int) LOWORD(wParam);
				int nPos = (short int) HIWORD(wParam);
				int nId = GetWindowLong((HWND)lParam, GWL_ID);
				switch (nId)
				{
				case ID_HEIGHTSPIN:
					if (nPos >= eImageMin && nPos <= eImageMax && m_pntEdit.m_size.cy != nPos)
					{
						m_pntEdit.SetSize(m_pntEdit.m_size.cx, nPos);
						SetDirty();
					}
					break;

				case ID_WIDTHSPIN:
					if (nPos >= eImageMin && nPos <= eImageMax && m_pntEdit.m_size.cx != nPos)
					{
						m_pntEdit.SetSize(nPos, m_pntEdit.m_size.cy);
						SetDirty();
					}
					break;
				}
			}
			break;

		case WM_SETFONT: 
			{
				HWND hWndChild = GetWindow(m_hWnd, GW_CHILD);
				while (hWndChild)
				{
					::SendMessage(hWndChild, nMsg, wParam, lParam);
					hWndChild = GetWindow(hWndChild, GW_HWNDNEXT);
				}
			}
			break;

		case WM_COMMAND:
			OnCommand(wParam, lParam);
			break;

		case WM_CREATE:
			if (!CreateChildren())
				return -1;
			break;

		case WM_SIZE:
			OnSize();
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
// OnUndo
//

void CIconEdit::OnUndo()
{
	int cx,cy;
	if (m_umManager.PopQuery(cx, cy))
	{
		if (cx != m_pntEdit.m_size.cx || cy != m_pntEdit.m_size.cy)
			m_pntEdit.SetSize(cx, cy);

		m_umManager.Pop((LPBYTE)m_pntEdit.m_pcrMainData[m_nImageIndex]);
		m_pntEdit.RefreshImage();
		m_pvWin.InvalidateRect(NULL, TRUE);
	}
}

//
// OnCommand
//

BOOL CIconEdit::OnCommand(WPARAM wParam, LPARAM lParam)
{
	BOOL bReturn = TRUE;
	WORD nId = LOWORD(wParam);
	WORD nNotifyCode = HIWORD(wParam);

	switch (nNotifyCode)
	{
	case BN_CLICKED:

		switch (nId)
		{
		case ID_SNAPSHOT:
			OnSnapShot(m_hWndSnapShotParent);
			break;

		case ID_PAINT_LASSO:
		case ID_PAINT_PEN:
		case ID_PAINT_PIPET:
		case ID_PAINT_FILL:
		case ID_PAINT_LINE:
		case ID_PAINT_PATTERN:
		case ID_PAINT_RECT0:
		case ID_PAINT_RECT1:
		case ID_PAINT_RECT2:
		case ID_PAINT_CIRC0:
		case ID_PAINT_CIRC1:
		case ID_PAINT_CIRC2:
			if (nId != m_nCurrentTool)
			{
				m_tbPaintTool.CheckButton(m_nCurrentTool,FALSE);
				m_nCurrentTool=nId;
				m_tbPaintTool.CheckButton(m_nCurrentTool,TRUE);
			}
			m_pntEdit.SetTool(m_nCurrentTool);
			break;

		case ID_SHIFT_DOWN:
		case ID_SHIFT_UP:
		case ID_SHIFT_LEFT:
		case ID_SHIFT_RIGHT:
			m_pntEdit.ShiftIcon(nId);
			break;

		case ID_NORMAL_IMAGE:
			SetImageId(ID_NORMAL_IMAGE);
			break;

		case ID_PRESSED_IMAGE:
			SetImageId(ID_PRESSED_IMAGE);
			break;

		case ID_HOVER_IMAGE:
			SetImageId(ID_HOVER_IMAGE);
			break;
		
		case ID_DISABLED_IMAGE:
			SetImageId(ID_DISABLED_IMAGE);
			break;

		case IDC_COLORWELL:
			if (Painter::eLeftButton == lParam)
				m_crForeground = m_cwPallette.GetCurColor();
			else
				m_crBackground = m_cwPallette.GetCurColor();
			UpdateColorPreview();
			break;

		case IDC_CLRPREVIEW:
			switch (lParam)
			{
			case Painter::eLeftButton:
				m_crForeground = 0xFFFFFFFF;
				break;

			case Painter::eRightButton:
				m_crBackground = 0xFFFFFFFF;
				break;

			default:
				{
					COLORREF crTemp = m_crForeground;
					m_crForeground = m_crBackground;
					m_crBackground = crTemp;
				}
				break;
			}
			UpdateColorPreview();
			break;

		case ID_FILEOPEN:
			LoadPicture();
			break;

		case ID_FILESAVE:
			SavePicture();
			break;

		case ID_CLEAR:
			ClearPicture();
			break;

		case ID_UNDO:
			OnUndo();
			break;

		case ID_CUT:
			Cut();
			break;

		case ID_COPY:
			Copy();
			break;

		case ID_PASTE:
			Paste();
			break;

		default:
			bReturn = FALSE;
		}
		break;

	case EN_KILLFOCUS:
		switch (nId)
		{
		case ID_WIDTH:
			{
				TCHAR szBuffer[128];
				szBuffer[0] = NULL;
				if (GetWindowText((HWND)lParam, szBuffer, 128) && NULL != szBuffer[0])
				{
					m_pntEdit.SetSize(atoi(szBuffer), m_pntEdit.m_size.cy);
					SetDirty();
				}
			}
			break;

		case ID_HEIGHT:
			{
				TCHAR szBuffer[128];
				szBuffer[0] = NULL;
				if (GetWindowText((HWND)lParam, szBuffer, 128) && NULL != szBuffer[0])
				{
					m_pntEdit.SetSize(m_pntEdit.m_size.cx, atoi(szBuffer));
					SetDirty();
				}
			}
			break;
		}
		break;
	}

	return bReturn;
}

//
// UndoPush
//

void CIconEdit::UndoPush()
{
	m_umManager.Push(m_pntEdit.m_size.cx, 
					 m_pntEdit.m_size.cy, 
					 (LPBYTE)m_pntEdit.m_pcrMainData[m_nImageIndex],
					 m_pntEdit.m_size.cx * m_pntEdit.m_size.cy * sizeof(COLORREF));
}

//
// SetImageIndex
//

void CIconEdit::SetImageId(int nId, BOOL bUpdateView)
{
	switch (nId)
	{
		case ID_NORMAL_IMAGE:
			m_nImageIndex = ddISNormal;
			break;
		case ID_HOVER_IMAGE:
			m_nImageIndex = ddISMask;
			break;
		case ID_PRESSED_IMAGE:
			m_nImageIndex = ddISGray;
			break;
		case ID_DISABLED_IMAGE:
			m_nImageIndex = ddISDisabled;
			break;
	}

	if (NULL == m_hWnd)
		return;

	m_tbShiftBar.CheckButton(m_nPrevId, FALSE);
	m_pntEdit.SetImageIndex(m_nImageIndex);
	if (bUpdateView)
	{
		m_pntEdit.RefreshImage();
		m_pvWin.InvalidateRect(NULL, FALSE);
	}
	m_nPrevId = nId;
	m_tbShiftBar.CheckButton(nId, TRUE);
}

//
// CreateChildren
//

BOOL CIconEdit::CreateChildren()
{
	CRect rc;

	// Paint Toolbar
	m_tbPaintTool.LoadToolBar(MAKEINTRESOURCE(IDB_PAINT_BAR));
	m_tbPaintTool.Create(m_hWnd, rc, ID_TOOLBAR);
	m_nCurrentTool = ID_PAINT_PEN;
	m_tbPaintTool.CheckButton(ID_PAINT_PEN, TRUE);

	// Shift Toolbar
	m_tbShiftBar.LoadToolBar(MAKEINTRESOURCE(IDB_STATEANDSHIFT));
	m_tbShiftBar.Create(m_hWnd, rc, ID_TOOLBAR+1);
	m_nPrevId = ID_NORMAL_IMAGE;
	m_tbShiftBar.CheckButton(ID_NORMAL_IMAGE, TRUE);
	m_tbShiftBar.SetStyle(DDToolBar::eRaised|DDToolBar::eIE);

	// FileDimensions Toolbar
	m_tbFileDimensions.LoadToolBar(MAKEINTRESOURCE(IDR_FILEDEMENSIONS));
	//m_tbFileDimensions.SetStyle(DDToolBar::eRaised|DDToolBar::eIE);
	DDToolBar::Tool* pTool;
	if (m_tbFileDimensions.GetTool(ID_PLACEHOLDERSPACE, pTool))
	{
		pTool->biButton.fsStyle = DDToolBar::TBSTYLE_PLACEHOLDER;
		pTool->sizeButton.cx = 40;
	}
	if (m_tbFileDimensions.GetTool(ID_STATICWIDTH, pTool))
	{
		pTool->biButton.fsStyle = DDToolBar::TBSTYLE_PLACEHOLDER|DDToolBar::TBSTYLE_KEEPTOGETHER;
		pTool->szWindowClass = _T("STATIC");
		pTool->sizeButton.cx = 40;
		pTool->dwStyle = SS_CENTERIMAGE;
	}
	if (m_tbFileDimensions.GetTool(ID_WIDTH, pTool))
	{
		pTool->biButton.fsStyle = DDToolBar::TBSTYLE_PLACEHOLDER|DDToolBar::TBSTYLE_KEEPTOGETHER;
		pTool->szWindowClass = _T("EDIT");
		pTool->sizeButton.cx = 30;
		pTool->dwStyle = ES_AUTOHSCROLL|WS_GROUP|WS_TABSTOP;
		pTool->dwStyleEx = WS_EX_CLIENTEDGE; 
	}
	if (m_tbFileDimensions.GetTool(ID_WIDTHSPIN, pTool))
	{
		pTool->biButton.fsStyle = DDToolBar::TBSTYLE_PLACEHOLDER;
		pTool->szWindowClass = UPDOWN_CLASS;
		pTool->sizeButton.cx = 15;
		pTool->dwStyle = UDS_AUTOBUDDY|UDS_SETBUDDYINT;
	}
	if (m_tbFileDimensions.GetTool(ID_STATICHEIGHT, pTool))
	{
		pTool->biButton.fsStyle = DDToolBar::TBSTYLE_PLACEHOLDER|DDToolBar::TBSTYLE_KEEPTOGETHER;
		pTool->szWindowClass = _T("STATIC");
		pTool->sizeButton.cx = 40;
		pTool->dwStyle = SS_CENTERIMAGE;
	}
	if (m_tbFileDimensions.GetTool(ID_HEIGHT, pTool))
	{
		pTool->biButton.fsStyle = DDToolBar::TBSTYLE_PLACEHOLDER|DDToolBar::TBSTYLE_KEEPTOGETHER;
		pTool->szWindowClass = _T("EDIT");
		pTool->sizeButton.cx = 30;
		pTool->dwStyle = ES_AUTOHSCROLL|WS_TABSTOP;
		pTool->dwStyleEx = WS_EX_CLIENTEDGE; 
	}
	if (m_tbFileDimensions.GetTool(ID_HEIGHTSPIN, pTool))
	{
		pTool->biButton.fsStyle = DDToolBar::TBSTYLE_PLACEHOLDER;
		pTool->szWindowClass = UPDOWN_CLASS;
		pTool->sizeButton.cx = 15;
		pTool->dwStyle = UDS_AUTOBUDDY|UDS_SETBUDDYINT;
	}
	m_tbFileDimensions.Create(m_hWnd, rc, ID_TOOLBAR+2);

	// Create Colorwell
	m_cwPallette.CreateWin(m_hWnd, rc, IDC_COLORWELL);

	// Painter Window
	m_pntEdit.CreateWin(m_hWnd); 
	m_pntEdit.SetTool(m_nCurrentTool);
	
	// Preview Window
	m_pvWin.CreateWin(m_hWnd, 
					  Preview::ePreview, 
					  IDC_IMAGEPREVIEW);

	// Color Review Window
	m_pvColor.CreateWin(m_hWnd, 
						Preview::eColor, 
						IDC_CLRPREVIEW);

	//
	// Initialize the Height and Width
	//
	
	SetToolBarImageDimemsions();
	if (m_tbFileDimensions.GetTool(ID_HEIGHTSPIN, pTool))
		::SendMessage(pTool->hWnd, UDM_SETRANGE, 0, MAKELONG(eImageMax, eImageMin));
	if (m_tbFileDimensions.GetTool(ID_WIDTHSPIN, pTool))
		::SendMessage(pTool->hWnd, UDM_SETRANGE, 0, MAKELONG(eImageMax, eImageMin));
	return TRUE;
}

//
// SetToolBarImageDimemsions
//

void CIconEdit::SetToolBarImageDimemsions()
{
	//
	// Initialize the Height and Width
	//

	DDToolBar::Tool* pTool;
	TCHAR szText[10];
	if (m_tbFileDimensions.GetTool(ID_HEIGHT, pTool))
	{
		_stprintf(szText, _T("%i"), m_pntEdit.m_size.cy);
		::SendMessage(pTool->hWnd, WM_SETTEXT, 0, (LPARAM)szText); 
	}
	if (m_tbFileDimensions.GetTool(ID_WIDTH, pTool))
	{
		_stprintf(szText, _T("%i"), m_pntEdit.m_size.cx);
		::SendMessage(pTool->hWnd, WM_SETTEXT, 0, (LPARAM)szText); 
	}
}

//
// OnSize
//

void CIconEdit::OnSize()
{
	CRect rcClient;
	GetClientRect(rcClient);

	// Preview Pane
	CRect rcPreview(8, 
					rcClient.left + ePreviewWidth, 
					4, 
					rcClient.top + ePreviewHeight);
	m_pvWin.MoveWindow(rcPreview, TRUE);

	// Paint Toolbar 
	CRect rc;
	rc.top = rcPreview.bottom + 4;
	rc.left = 8;
	rc.right = rc.left + ePreviewWidth;
	rc.bottom = rc.top + 81 + 26;
	m_tbPaintTool.MoveWindow(rc, TRUE);

	// Color Well
	rc.left += 2;
	rc.top = rc.bottom + 4;
	rc.bottom = rc.top + 4 * 18;
	rc.right = rc.left + 4 * 18;
	m_cwPallette.m_bSupportSelection = FALSE;
	m_cwPallette.MoveWindow(rc, TRUE);

	// Color Preview
	rc.top = rc.bottom + 1;
	rc.bottom = rc.top + eColorPreviewHeight;
	m_pvColor.MoveWindow(rc, TRUE);

	// Painter
	rc = rcClient;
	rc.left += eBevel + ePreviewWidth + 8;
	rc.top += eBevel * 2 + eFileDimensionHeight;
	rc.bottom -= Painter::eShiftBarHeight + eBevel;
	rc.right = rc.left + min (rc.Height(), rc.Width());

	// check if width > height or height < width
	if (0 == m_pntEdit.m_size.cx || 0 == m_pntEdit.m_size.cy)
	{ 
		m_pntEdit.SetSize(1, 1);
		m_pntEdit.SetBoxSize(1);
	}
	else
	{ 
		const int nXEdge = GetSystemMetrics(SM_CXEDGE);
		const int nYEdge = GetSystemMetrics(SM_CYEDGE);
		int nBoxSizeH = (rc.Width() - (nXEdge * 2) - (m_pntEdit.m_size.cx - 1)) / m_pntEdit.m_size.cx;
		int nBoxSizeV = (rc.Height() - (nYEdge * 2) - (m_pntEdit.m_size.cy - 1)) / m_pntEdit.m_size.cy;
		int nBoxSize = min (nBoxSizeH, nBoxSizeV);
		m_pntEdit.SetBoxSize(nBoxSize);
	}
	m_pntEdit.MoveWindow(rc, TRUE);

	// Shift Toolbar
	rc.top = rc.bottom + 3;
	rc.bottom = rc.top + Painter::eShiftBarHeight;
	if (rc.Width() < (20 + 2) * eStateAndShiftCount + 8)
		rc.right = rc.left + (20 + 2) * eStateAndShiftCount + 8;
	m_tbShiftBar.MoveWindow(rc, TRUE);

	// FileDimensions Toolbar
	rc.left = rcClient.left = eBevel + ePreviewWidth + 8;
	rc.top = rcClient.top + eBevel;
	rc.bottom = rc.top + eFileDimensionHeight;
	m_tbFileDimensions.MoveWindow(rc, TRUE);
}

//
// SetBitmaps
//

BOOL CIconEdit::SetBitmaps(HBITMAP hBitmap, HBITMAP hBitmapMask, COLORREF crMask)
{
	if (NULL == hBitmap || NULL == hBitmapMask)
	{
		if (0 == m_nImageIndex)
		{
			m_pntEdit.SetSize(Painter::eDefaultPainterSize, Painter::eDefaultPainterSize);
			SetToolBarImageDimemsions();
		}
		return TRUE;
	}

	BITMAP bmInfo;
	GetObject(hBitmap,sizeof(BITMAP),&bmInfo);
	int nWidth = bmInfo.bmWidth;
	int nHeight = bmInfo.bmHeight;

	if (0 == m_nImageIndex)
	{
		m_pntEdit.SetSize(nWidth, nHeight);
		SetToolBarImageDimemsions();
	}

	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
	{
		ReleaseDC(NULL, hDC);
		return FALSE;
	}

	HDC hDCMem = CreateCompatibleDC(hDC);
	if (NULL == hDCMem)
		return FALSE;

	HBITMAP hBitmapOld = SelectBitmap(hDCMem, hBitmap);

	HPALETTE hPalOld;
	if (g_hPal)
		hPalOld = SelectPalette(hDCMem, g_hPal, FALSE);
	
	COLORREF* pData = m_pntEdit.m_pcrImageData;
	int x,y;
	for (y = 0; y < nHeight; y++)
	{
		for (x = 0; x < nWidth; x++)
		{
			*pData = GetPixel(hDCMem, x, y);
			pData++;
		}
	}
	
	if (g_hPal)
		SelectPalette(hDCMem, hPalOld, FALSE);
	
	SelectBitmap(hDCMem, hBitmapOld);

	if (NULL == hBitmapMask)
	{
		int nSize = m_pntEdit.m_size.cx * m_pntEdit.m_size.cy;
		for (int nIndex = 0; nIndex < nSize; nIndex++)
			if (m_pntEdit.m_pcrImageData[nIndex] == crMask)
				m_pntEdit.m_pcrImageData[nIndex] = 0xFFFFFFFF;
	}
	else
	{
		pData = m_pntEdit.m_pcrImageData;
		hBitmapOld = SelectBitmap(hDCMem, hBitmapMask);
		for (y = 0; y < nHeight; y++)
		{
			for (x = 0; x < nWidth; x++)
			{
				if (GetPixel(hDCMem, x, y) != 0)
					*pData = 0xFFFFFFFF;
				pData++;
			}
		}
		SelectBitmap(hDCMem, hBitmapOld);
	}

	DeleteDC(hDCMem);
	ReleaseDC(NULL, hDC);
	return TRUE;
}

//
// GetBitmaps
//

BOOL CIconEdit::GetBitmaps(HBITMAP& hBitmap, HBITMAP& hMaskBitmap)
{
	SIZE sizeBitmap = m_pntEdit.m_size;
	COLORREF* pcrImageData = m_pntEdit.m_pcrImageData;

	// first check if bitmap is fully empty
	const int nSizeBitmap = sizeBitmap.cx * sizeBitmap.cy;
	for (int nIndex = 0; nIndex < nSizeBitmap; ++nIndex)
	{
		if (pcrImageData[nIndex] != 0xFFFFFFFF)
			break;
	}

	HBITMAP bmp, bmpMask;
	if (nIndex == nSizeBitmap)
	{
		hBitmap = NULL;
		hMaskBitmap = NULL;
		return FALSE;
	}

	HDC hDC = GetDC(NULL);
	assert(hDC);
	if (NULL == hDC)
		return FALSE;

	HDC hDCMem = CreateCompatibleDC(hDC);
	assert(hDCMem);
	if (NULL == hDCMem)
	{
		ReleaseDC(NULL, hDC);
		return FALSE;
	}

	HPALETTE oldPal,oldPal2;
	if (g_hPal)
	{
		oldPal=SelectPalette(hDCMem, g_hPal, FALSE);
		oldPal2=SelectPalette(hDC, g_hPal, FALSE);
		RealizePalette(hDCMem);
		RealizePalette(hDC);
	}

	bmp = CreateCompatibleBitmap(hDC, sizeBitmap.cx, sizeBitmap.cy); 
	bmpMask = CreateBitmap(sizeBitmap.cx, sizeBitmap.cy, 1, 1, 0);
	
	HBITMAP oldBitmap = SelectBitmap(hDCMem, bmp);

	CRect r;
	SetRect(&r, 0, 0, sizeBitmap.cx, sizeBitmap.cy);
	FillRect(hDCMem, &r, (HBRUSH)(1+COLOR_BTNFACE));

	CRect rx;
	SetRect(&rx, 0, 0, sizeBitmap.cx, sizeBitmap.cy);
	m_pvWin.PaintDibSection(hDCMem, rx);

	SelectObject(hDCMem,oldBitmap);

	int x,y;
	oldBitmap = SelectBitmap(hDCMem, bmpMask);
	for (y=0; y < sizeBitmap.cy; ++y)
	{
		for (x=0; x < sizeBitmap.cx; ++x)
		{
			if (pcrImageData[x+sizeBitmap.cx*y] != 0xFFFFFFFF)
				SetPixel(hDCMem, x, y, 0); // White
			else
				SetPixel(hDCMem, x, y, 0xFFFFFF); // Black
		}
	}
	SelectObject(hDCMem, oldBitmap);
	
	if (g_hPal)
	{
		SelectPalette(hDCMem, oldPal, FALSE);
		SelectPalette(hDC, oldPal2, FALSE);
	}

	DeleteDC(hDCMem);
	ReleaseDC(NULL, hDC);
	hBitmap = bmp;
	hMaskBitmap = bmpMask;
	return TRUE;
}

static CIconEdit* s_pSnapShotWnd = NULL;

//
// SnapShotWndProc
//

LRESULT CALLBACK CIconEdit::SnapShotWndProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(hWnd,&ps);
			if (hDC)
			{
				s_pSnapShotWnd->DrawSnapShotZoom(hWnd,
												 hDC,
												 s_pSnapShotWnd->m_ptZCaretPos.x,
												 s_pSnapShotWnd->m_ptZCaretPos.y);
				EndPaint(hWnd, &ps);
				return 0;
			}
		}
		break;
	}
	return DefWindowProc(hWnd, nMsg, wParam, lParam);
}

//
// OnSnapShot
//

void CIconEdit::OnSnapShot(HWND hWndParent)
{
	HWND hWnd = NULL;
	try
	{
		::ShowWindow(hWndParent, SW_SHOWMINIMIZED);

		// TopLeft
		m_nZArea = 0; 
		s_pSnapShotWnd = this;

		HCURSOR hCursorPrev;
		POINT   ptCurrent, ptNew;
		BOOL    bCancelled = FALSE;
		int		nKeyboardOffsetX = 0;
		int		nKeyboardOffsetY = 0;
		int		nVisibleCnt = 0;
		int		nMaxX = GetSystemMetrics(SM_CXSCREEN) - m_pntEdit.m_size.cx;
		int		nMaxY = GetSystemMetrics(SM_CYSCREEN) - m_pntEdit.m_size.cy;

		WNDCLASS wndClass;
		wndClass.style = CS_HREDRAW|CS_VREDRAW;
		wndClass.lpfnWndProc=SnapShotWndProc;
		wndClass.cbClsExtra=0;
		wndClass.cbWndExtra=0;
		wndClass.hInstance = g_hInstance;
		wndClass.hIcon = NULL;
		wndClass.hCursor = NULL;
		wndClass.hbrBackground = NULL;
		wndClass.lpszMenuName=NULL;
		wndClass.lpszClassName = _T("DDSnapShot");
		RegisterClass(&wndClass);

		hWnd = CreateWindowEx(WS_EX_TRANSPARENT | WS_EX_TOPMOST,
							  wndClass.lpszClassName,
							  _T(""),
							  WS_POPUP,
							  0,
							  0,
							  GetSystemMetrics(SM_CXSCREEN),
							  GetSystemMetrics(SM_CYSCREEN),
							  NULL,
							  NULL,
							  g_hInstance,
							  NULL);

		if (NULL == hWnd)
			return;

		m_hWndSnap = hWnd;
		GetCursorPos(&ptCurrent);
		m_ptZCaretPos = ptCurrent;

		::ShowWindow(hWnd, SW_SHOWNORMAL);
		
		hCursorPrev = ::SetCursor(NULL);
		SetCapture(hWnd);
		UpdateWindow();
		UpdateSnapShotZoom(hWnd, ptCurrent.x, ptCurrent.y);
		
		SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDC_SNAPSHOT)));
		
		nKeyboardOffsetX = nKeyboardOffsetY = 0;
		
		MSG msg;
		while (TRUE)
		{
			if (GetCapture() != hWnd)
				break;

			GetMessage(&msg, NULL, 0, 0);
			switch (msg.message)
			{
			// handle movement/accept messages
			case WM_CANCELMODE:
				bCancelled=TRUE;
				goto ExitLoop;
			
			case WM_LBUTTONDOWN:
			case WM_RBUTTONDOWN:
			case WM_LBUTTONUP:
			case WM_RBUTTONUP:
				goto ExitLoop;

			case WM_KEYDOWN:
				switch (msg.wParam)
				{
				case VK_LEFT:
					--nKeyboardOffsetX;
					goto MovePointKeyboard;

				case VK_RIGHT:
					++nKeyboardOffsetX;
					goto MovePointKeyboard;

				case VK_UP:
					--nKeyboardOffsetY;
					goto MovePointKeyboard;

				case VK_DOWN:
					++nKeyboardOffsetY;
					goto MovePointKeyboard;

				case VK_SPACE:
					break;
				}
				break;

			case WM_KEYUP:
				switch (msg.wParam)
				{
					case VK_ESCAPE:
					{
						bCancelled = TRUE;
						goto ExitLoop;
					}
					case VK_RETURN:
						goto ExitLoop;
				}
				break;

			case WM_MOUSEMOVE:
			{
	MovePointKeyboard:
				GetCursorPos(&ptNew);
				ptNew.x += nKeyboardOffsetX;
				ptNew.y += nKeyboardOffsetY;

				if (ptNew.x > nMaxX)
					ptNew.x = nMaxX;
				if (ptNew.y > nMaxY)
					ptNew.y = nMaxY;

				ptCurrent = ptNew;
				m_ptZCaretPos = ptCurrent;
				UpdateSnapShotZoom(hWnd, ptCurrent.x, ptCurrent.y);
			}
			break;

			// just dispatch rest of the messages
			default:	
				DispatchMessage(&msg);
				break;
			}
		}

ExitLoop:
		if (GetCapture() == hWnd)
			ReleaseCapture();

		if (!bCancelled)
		{
			// Transfer screen to pArray
			COLORREF* pcrData = m_pntEdit.m_pcrImageData;
			assert(pcrData);

			COLORREF  crBtnFace = GetSysColor(COLOR_BTNFACE);
			COLORREF  crTemp;

			int nWidth = m_pntEdit.m_size.cx;
			int nHeight = m_pntEdit.m_size.cy;
			HDC hDCScreen = GetDC(hWnd);
			if (hDCScreen)
			{
				int nX;
				int nY;
				for (nY = 0; nY < nHeight; nY++)
				{
					for (nX = 0; nX < nWidth; nX++)
					{
						crTemp = GetPixel(hDCScreen, ptCurrent.x + nX, ptCurrent.y + nY);
						if ((crTemp & 0xF0F0F0) == crBtnFace || crTemp == crBtnFace)
							crTemp = 0xFFFFFFFF;
						*pcrData++ = crTemp;
					}
				}

				m_bDirty[m_nImageIndex] = TRUE;
				ReleaseDC(hWnd, hDCScreen);
			}
		}
		::DestroyWindow(hWnd);
		SetCursor(hCursorPrev);
	}
	CATCH
	{
		assert(FALSE);
		if (GetCapture() == hWnd)
			ReleaseCapture();
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	::ShowWindow(hWndParent, SW_RESTORE);
}

//
// UpdateSnapShotZoom
//
// Grab image at x,y. zoom image by 8x + 1 line and blast 
// to flicker free dc then screen

BOOL CIconEdit::UpdateSnapShotZoom(HWND hWnd, int x, int y)
{
	BOOL bResult = FALSE;
	try
	{
		int w = m_pntEdit.m_size.cx;
		int h = m_pntEdit.m_size.cy;
		POINT pt = {x, y};

		CRect rcSnapShotArea;
		if (0 == m_nZArea)
		{
			rcSnapShotArea;
			rcSnapShotArea.right = w * 8 + 12;
			rcSnapShotArea.bottom = h * 8 + 12;
		}
		else
		{
			::GetWindowRect(hWnd,&rcSnapShotArea);
			rcSnapShotArea.right = rcSnapShotArea.left + w * 8 + 12;
			rcSnapShotArea.top = rcSnapShotArea.bottom - (h * 8 + 12);
		}

		if (PtInRect(&rcSnapShotArea, pt))
		{ 
			// switch area
			m_nZArea ^= 1;
			::InvalidateRect(NULL, NULL, TRUE);
			::UpdateWindow(hWnd);
		}
		
		HDC hDCScreen = GetDC(hWnd);
		if (hDCScreen)
		{
			bResult = DrawSnapShotZoom(hWnd, hDCScreen, x, y);
			ReleaseDC(hWnd, hDCScreen);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return bResult;
}

//
// DrawSnapShotZoom
//

BOOL CIconEdit::DrawSnapShotZoom(HWND hWnd, HDC hDCScreen, int x, int y)
{
	BOOL bResult = FALSE;
	CRect rcSnapShotArea;
	int w = m_pntEdit.m_size.cx;
	int h = m_pntEdit.m_size.cy;
	
	try
	{
		if (0 == m_nZArea)	
		{
			rcSnapShotArea.right = w * 8 + 12;
			rcSnapShotArea.bottom = h * 8 + 12;
		}
		else
		{
			::GetWindowRect(hWnd, &rcSnapShotArea);
			rcSnapShotArea.right = rcSnapShotArea.left + w * 8 + 12;
			rcSnapShotArea.top = rcSnapShotArea.bottom - (h * 8 + 12);
		}

		if (NULL == GetGlobals().FlickerFree())
			return FALSE;

		// +6 is for border
		HDC hDCOff = GetGlobals().FlickerFree()->RequestDC(hDCScreen, w * 8 + 12, h * 8 + 12); 
		if (NULL == hDCOff)
			return FALSE;

		CRect rc;
		rc.right = w * 8 + 12;
		rc.bottom = h * 8 + 12;
		FillRect(hDCOff, &rc, (HBRUSH)(1+COLOR_BTNFACE));

		CFlickerFree ffSmall;
		HDC hDCSmall = ffSmall.RequestDC(hDCScreen,w,h);
		if (NULL == hDCSmall)
			return FALSE;

		bResult = BitBlt(hDCSmall, 0, 0, w, h, hDCScreen, x, y, SRCCOPY);
		if (!bResult)
			return bResult;

		bResult = StretchBlt(hDCOff, 6, 6, w * 8, h * 8, hDCSmall, 0, 0, w, h, SRCCOPY);
		if (!bResult)
			return bResult;

		DrawEdge(hDCOff, &rc, EDGE_RAISED, BF_RECT);
		
		rc.Inflate(-4, -4);

		DrawEdge(hDCOff, &rc, EDGE_SUNKEN, BF_RECT);

		// Now draw lines
		SetBkColor(hDCOff, 0);
		rc.left = 6;
		rc.right = 6 + w * 8;
		for (y = 1; y < h; ++y)
		{
			rc.top = 6 + y * 8;
			rc.bottom = rc.top + 1;
			ExtTextOut(hDCOff, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);
		}

		rc.top = 6 + 0;
		rc.bottom = 6 + h * 8;
		for (x = 1; x < w; ++x)
		{
			rc.left = 6 + x * 8;
			rc.right = rc.left + 1;
			ExtTextOut(hDCOff, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);
		}

		if (0 == m_nZArea)
			GetGlobals().FlickerFree()->Paint(hDCScreen, 0, 0);
		else
		{
			::GetWindowRect(hWnd, &rcSnapShotArea);
			GetGlobals().FlickerFree()->Paint(hDCScreen, 0, rcSnapShotArea.bottom - (h * 8 + 12));
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return bResult;
}

//
// OlePictureFromHBitmap
//
// HBitmap to IPicture 
//

static IPicture* OlePictureFromHBitmap(HBITMAP hBM)
{
	// now we have the stream
	IPicture* pPict = NULL;
	PICTDESC picDesc;
	picDesc.cbSizeofstruct = sizeof(PICTDESC);
	picDesc.picType = PICTYPE_BITMAP;
    picDesc.bmp.hbitmap = hBM;
	picDesc.bmp.hpal = 0;
	HRESULT hResult = OleCreatePictureIndirect(&picDesc, IID_IPicture, TRUE, (LPVOID*)&pPict);
	if (FAILED(hResult))
		return NULL;
	return pPict;
}

//
// SavePicture
//

void CIconEdit::SavePicture()
{
	int nBitCount = 24;
	DDString strPictureTypes;
	strPictureTypes.LoadString(IDS_SAVEPICTURETYPES);

	CFileDialog dlgSave(IDS_SAVEIMAGE,
						FALSE,
						_T("*.bmp"),
						_T("*.bmp"),
						OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY | OFN_FILEMUSTEXIST,
						strPictureTypes,
						m_hWnd);
	if (IDOK == dlgSave.DoModal())
	{
		LPCTSTR szFileName = dlgSave.GetFileName();
		HANDLE hFile = CreateFile(szFileName, 
								  GENERIC_WRITE, 
								  FILE_SHARE_READ, 
								  0, 
								  CREATE_ALWAYS, 
								  FILE_ATTRIBUTE_NORMAL, 
								  0);
		if (hFile)
		{
			SaveBitmap(hFile, nBitCount);
			CloseHandle(hFile);
		}
		else
			MessageBox (IDS_ERR_COULNDNOTCREATEFILE, MB_OK);
	}
}

//
// IconPreview
//

class IconPreview : public FWnd
{
public:
	IconPreview();
	~IconPreview();

	void ClearPicture();
	BOOL SetPicture(LPCTSTR szPicture);

private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	BOOL Draw(HDC hDC, CRect& rcClient); 
	IPicture* m_pPicture;
};

IconPreview::IconPreview()
	: m_pPicture(NULL)
{
}

IconPreview::~IconPreview()
{
	if (m_pPicture)
		m_pPicture->Release();
}

//
// SetPicture
//

BOOL IconPreview::SetPicture(LPCTSTR szPicture)
{
	if (m_pPicture)
	{
		m_pPicture->Release();
		m_pPicture = NULL;
	}
	m_pPicture = OleLoadPictureHelper(szPicture);
	InvalidateRect(NULL, TRUE);
	if (NULL == m_pPicture)
	{
		MessageBox(IDS_ERR_FAILEDTOLOADPICTURE);
		return FALSE;
	}
	return TRUE;
}

void IconPreview::ClearPicture()
{
	if (m_pPicture)
	{
		m_pPicture->Release();
		m_pPicture = NULL;
		InvalidateRect(NULL, TRUE);
	}
}

//
// WindowProc
//

LRESULT IconPreview::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcClient;
					GetClientRect(rcClient);
					Draw(hDC, rcClient);
					EndPaint(m_hWnd, &ps);
					return 0;
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
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// Draw
//

BOOL IconPreview::Draw(HDC hDC, CRect& rcClient)
{
	try
	{
		FillRect(hDC, &rcClient, (HBRUSH)(COLOR_BTNFACE + 1)); 
		if (NULL == m_pPicture)
			return FALSE;

		SIZEL size;
		HRESULT hResult = m_pPicture->get_Width(&size.cx);
		hResult = m_pPicture->get_Height(&size.cy);

		HiMetricToPixel(&size, &size);
		int nWidth = rcClient.Width();
		int nHeight = rcClient.Height();

		if (size.cx > nWidth)
			size.cx = nWidth;
		
		if (size.cy > nHeight)
			size.cy = nHeight;

		int nWidthPicture = size.cx;
		int nHeightPicture = size.cy;

		PixelToHiMetric(&size, &size);
		hResult = m_pPicture->Render(hDC, 
								     (nWidth - nWidthPicture) / 2, 
								     (nHeight - nHeightPicture) / 2, 
								     nWidthPicture, 
								     nHeightPicture, 
								     0, 
								     size.cy - 1, 
								     size.cx, 
								     -size.cy, 
								     NULL);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return TRUE;
}

//
// OFNOpenHookProc
//

static int s_nCreateMask = 0;
static IconPreview s_thePreview; 

UINT CALLBACK OFNOpenHookProc(HWND hdlg, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_INITDIALOG:
		if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_AUTOCREATEMASKCONFIRM))
		{
			::SendMessage(GetDlgItem(hdlg, IDC_CHK_CREATEMASK), BM_SETCHECK, BST_CHECKED, 0);
			s_nCreateMask = 1;
		}
		else
			s_nCreateMask = 0;

		s_thePreview.SubClassAttach(GetDlgItem(hdlg, IDC_PICTURE));
		break;

	case WM_DESTROY:
		s_thePreview.UnsubClass();
		break;

	case WM_NOTIFY:
		{
			NMHDR* pNotify = (NMHDR*)lParam;
			switch (pNotify->code)
			{
			case CDN_SELCHANGE:
				{
					TCHAR szFile[MAX_PATH];
					int nResult = CommDlg_OpenSave_GetFilePath(GetParent(hdlg), szFile, MAX_PATH);  
					assert(nResult);
					if (nResult)
					{
						DWORD dwFileAttributes = GetFileAttributes(szFile);
						if (!(FILE_ATTRIBUTE_DIRECTORY & dwFileAttributes))
							s_thePreview.SetPicture(szFile);
						else
							s_thePreview.ClearPicture();
					}
				}
				break;
			}
		}
		break;

	case WM_COMMAND:
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				switch (LOWORD(wParam))
				{
				case IDC_CHK_CREATEMASK:
					s_nCreateMask = SendMessage((HWND)lParam, BM_GETCHECK, 0, 0);
					break;
				}
				break;
			}
			break;
		}
		break;
	}
	return 0;
}

//
// LoadPicture
//

void CIconEdit::LoadPicture()
{
	DDString strPictureTypes;
	strPictureTypes.LoadString(IDS_LOADPICTURETYPES);

	CFileDialog dlgFile(IDS_LOADPICT,
						TRUE,
						_T("*.bmp;*.ico;*.cur,*.gif;*.jpg"),
						_T("*.bmp;*.ico;*.cur,*.gif;*.jpg"),
						OFN_FILEMUSTEXIST | OFN_HIDEREADONLY,
						strPictureTypes,
						m_hWnd,
						IDD_OPEN_ICONEDITOR,
						&OFNOpenHookProc);
	
	if (IDOK == dlgFile.DoModal())
	{
		IPicture* pNewPicture = OleLoadPictureHelper(dlgFile.GetFileName());
		if (NULL == pNewPicture)
		{
			MessageBox (IDS_ERR_LOADINGIMAGEFAILED, MB_ICONSTOP);
			return;
		}
		PastePicture(pNewPicture, BST_CHECKED == s_nCreateMask);
		pNewPicture->Release();
	}
}

//
// PastePicture
//

void CIconEdit::PastePicture(IPicture* pPict, BOOL bCreateMask)
{
	UndoPush();

	UINT nType = 0;

	OLE_HANDLE hImage = NULL;
	HRESULT hResult = pPict->get_Handle(&hImage);
	if (FAILED(hResult))
		return;

	hResult = pPict->get_Type((short*)&nType);
	if (FAILED(hResult))
		return;

	switch(nType)
	{
	case PICTYPE_BITMAP:
		{
			BITMAP bmInfo;
			int nPrev;
			GetObject((HGDIOBJ)hImage, sizeof(BITMAP), &bmInfo);
			if (bmInfo.bmWidth > eImageMax || bmInfo.bmHeight > eImageMax)
			{
				MessageBox (IDS_ERR_IMAGETOOBIG, MB_ICONSTOP);
				return;
			}

			PasteBitmap((HBITMAP)hImage, FALSE, nPrev);
			if (bCreateMask)
			{
				HDC hDC = GetDC(NULL);
				if (hDC)
				{
					HBITMAP hBitmapMask = CreateMaskBitmap(hDC, (HBITMAP)hImage, bmInfo.bmWidth, bmInfo.bmHeight);
					if (hBitmapMask)
					{
						PasteBitmap(hBitmapMask, TRUE, nPrev);
						DeleteBitmap(hBitmapMask);
					}
					ReleaseDC(NULL, hDC);
				}
			}
			SetToolBarImageDimemsions();
			m_bDirty[m_nImageIndex] = TRUE;
			m_pntEdit.RefreshImage();
		}
		break;

	case PICTYPE_ICON:
		PasteIcon((HICON)hImage);
		m_bDirty[m_nImageIndex] = TRUE;
		m_pntEdit.RefreshImage();
		break;
	}
}

//
// CBitmapScaleDlg
//

class CBitmapScaleDlg :public FDialog
{
public:
	enum OPTIONS
	{
		eEqualToSource,
		eFitCurrentSize,
		eClipToFit
	};

	CBitmapScaleDlg() 
		: FDialog(IDD_BITMAPSCALE) 
	{
	};

	OPTIONS m_eOptions;

private:
	virtual BOOL DialogProc(UINT nMsg,WPARAM wParam,LPARAM lParam);
	virtual void OnOK();		
};

BOOL CBitmapScaleDlg::DialogProc(UINT nMsg,WPARAM wParam,LPARAM lParam)
{
	switch(nMsg)
	{
	case WM_INITDIALOG:
		SetWindowPos(m_hWnd, HWND_TOPMOST, 0,0,0,0, SWP_NOMOVE|SWP_NOSIZE);
		CenterDialog(GetParent(m_hWnd));
		SendMessage(GetDlgItem(m_hWnd, IDC_OPT_EQUAL), BM_SETCHECK, TRUE, 0);
		break;
	}
	return FDialog::DialogProc(nMsg,wParam,lParam);
}

void CBitmapScaleDlg::OnOK()
{
	if (SendMessage(GetDlgItem(m_hWnd, IDC_OPT_EQUAL), BM_GETCHECK, 0, 0))
		m_eOptions = eEqualToSource;
	else if (SendMessage(GetDlgItem(m_hWnd, IDC_OPT_SCALETOTOOL), BM_GETCHECK, 0, 0))
		m_eOptions = eFitCurrentSize;
	else
		m_eOptions = eClipToFit;
	FDialog::OnOK();
}

//
// PasteBitmap
//

void CIconEdit::PasteBitmap(HBITMAP hBitmapSrc, BOOL bMask, int& nOption)
{
	BITMAP bitmapInfo;
	GetObject(hBitmapSrc, sizeof(BITMAP), &bitmapInfo);

	int nSrcWidth = bitmapInfo.bmWidth;
	int nSrcHeight = bitmapInfo.bmHeight;

	if (!bMask)
		nOption = CBitmapScaleDlg::eEqualToSource;

	if (!bMask && (nSrcWidth > m_pntEdit.m_size.cx || nSrcHeight > m_pntEdit.m_size.cy))
	{
		CBitmapScaleDlg dlgScaleBitmap;
		if (IDOK != dlgScaleBitmap.DoModal())
			return;

		nOption = dlgScaleBitmap.m_eOptions;
	}
	
	HBITMAP hBitmap = hBitmapSrc;
	switch (nOption)
	{
	case CBitmapScaleDlg::eEqualToSource: 
		// Set destintion size == source size 
		m_pntEdit.SetSize(nSrcWidth, nSrcHeight);
		break;

	case CBitmapScaleDlg::eFitCurrentSize:
		// Scale bitmap to fit
		hBitmap = ScaleBitmap(hBitmapSrc, m_pntEdit.m_size);
		nSrcWidth = m_pntEdit.m_size.cx;
		nSrcHeight = m_pntEdit.m_size.cy;
		break;
	}

	COLORREF crColor;
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		HDC hDCMem = CreateCompatibleDC(hDC);
		if (hDCMem)
		{
			HBITMAP hBitmapOld = SelectBitmap(hDCMem, hBitmap);
			int w = min(nSrcWidth, m_pntEdit.m_size.cx);
			int h = min(nSrcHeight, m_pntEdit.m_size.cy);
			int x,y;
			if (bMask)
			{
				SetBkColor(hDCMem, RGB(0,0,0));
				SetTextColor(hDCMem, RGB(255,255,255));
				for (y = 0; y < h; y++)
				{
					for (x = 0; x < w; x++)
					{
						crColor = GetPixel(hDCMem, x, y);
						if (crColor != 0)
							m_pntEdit.PutPixelRaw(x, y, 0xFFFFFFFF);
					}
				}
			}
			else
			{
				for (y = 0; y < h; ++y)
				{
					for (x = 0; x < w; ++x)
					{
						crColor = GetPixel(hDCMem, x, y);
						m_pntEdit.PutPixelRaw(x, y, crColor);
					}
				}
			}
			SelectBitmap(hDCMem, hBitmapOld);
			DeleteDC(hDCMem);
		}
		ReleaseDC(NULL, hDC);
	}
	if (CBitmapScaleDlg::eFitCurrentSize == nOption)
	{
		// Scale
		DeleteBitmap(hBitmap);
	}
}

//
// PasteIcon
//

void CIconEdit::PasteIcon(HICON hIcon)
{
	ICONINFO info;
	memset(&info, 0, sizeof(ICONINFO));
	info.fIcon = TRUE;
	if (GetIconInfo(hIcon, &info))
	{
		int nPrev;
		if (info.hbmColor)
		{
			//
			// We got an icon
			//

			PasteBitmap(info.hbmColor, FALSE, nPrev);
			PasteBitmap(info.hbmMask, TRUE, nPrev);
		}
		else if (info.hbmMask)
		{
			// 
			// We got a cursor
			//

			BITMAP bmInfo;
			int nResult = GetObject(info.hbmMask, sizeof(BITMAP), &bmInfo);
			if (0 == nResult)
				return;

			HDC hDC = GetDC(NULL);
			if (hDC)
			{
				int nHeight = bmInfo.bmHeight / 2;
				HDC hDCSrc = CreateCompatibleDC(hDC);
				HBITMAP hBitmapSrcOld = SelectBitmap(hDCSrc, info.hbmMask);
				
				HDC hDCMask = CreateCompatibleDC(hDCSrc);
				HBITMAP hMask = CreateBitmap(bmInfo.bmWidth, nHeight, 1, 1, 0); 
				HBITMAP hBitmapMaskOld = SelectBitmap(hDCMask, hMask);
				BitBlt(hDCMask, 0, 0, bmInfo.bmWidth, nHeight, hDCSrc, 0, 0, SRCCOPY);

				SelectBitmap(hDCMask, hBitmapMaskOld);
				DeleteDC(hDCMask);

				HDC hDCImage = CreateCompatibleDC(hDCSrc);
				HBITMAP hImage = CreateBitmap(bmInfo.bmWidth, nHeight, 1, 1, 0); 
				HBITMAP hBitmapImageOld = SelectBitmap(hDCImage, hImage);
				BitBlt(hDCImage, 0, 0, bmInfo.bmWidth, nHeight, hDCSrc, 0, nHeight, SRCCOPY);

				SelectBitmap(hDCImage, hBitmapImageOld);
				DeleteDC(hDCImage);

				SelectBitmap(hDCSrc, hBitmapSrcOld);
				DeleteDC(hDCSrc);

				PasteBitmap(hImage, FALSE, nPrev);
				PasteBitmap(hMask, TRUE, nPrev);

				DeleteBitmap(hImage);
				DeleteBitmap(hMask);
				ReleaseDC(NULL, hDC);
			}
		}
		SetToolBarImageDimemsions();
	}
}

//
// SavePicture
//

BOOL CIconEdit::SaveBitmap(HANDLE hFile, int nBitCount)
{
	HBITMAP hBmp;
	HBITMAP hDib;
	LPVOID pBits = 0; 
	BITMAPINFO bmInfo;
	if (!CreateImageBitmap(hBmp, hDib, bmInfo, pBits, nBitCount))
	{
		MessageBox (IDS_ERR_COULDNOTCREATEBITMAP, MB_OK);
		return FALSE;
	}

	int nColorTableEntries = sizeof(bmInfo.bmiColors)  / sizeof(RGBQUAD);
	BITMAPFILEHEADER bmfh;
	bmfh.bfType = 0x4d42; // 'BM'
	int nSizeHdr = sizeof(BITMAPINFOHEADER) + nColorTableEntries;
	bmfh.bfSize = 0;
	bmfh.bfReserved1 = bmfh.bfReserved2 = 0;
	bmfh.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER) + nColorTableEntries;	

	DWORD dwNumberOfBytesWritten = 0;

	BOOL bResult = WriteFile(hFile, (LPVOID)&bmfh, sizeof(BITMAPFILEHEADER), &dwNumberOfBytesWritten, 0); 
	if (!bResult)
	{
		MessageBox (IDS_ERR_COULDNOTSAVEIMAGE);
		return FALSE;
	}

	bResult = WriteFile(hFile, (LPVOID)&bmInfo, nSizeHdr, &dwNumberOfBytesWritten, 0); 
	if (!bResult)
	{
		MessageBox (IDS_ERR_COULDNOTSAVEIMAGE);
		return FALSE;
	}

	DWORD m_dwSizeImage = 0;
	if (bmInfo.bmiHeader.biSizeImage == 0)
	{
		DWORD dwBytes = ((DWORD) bmInfo.bmiHeader.biWidth * bmInfo.bmiHeader.biBitCount) / 32;
		if (((DWORD)bmInfo.bmiHeader.biWidth * bmInfo.bmiHeader.biBitCount) % 32) 
			dwBytes++;
		dwBytes *= 4;
		// no compression
		m_dwSizeImage = dwBytes * bmInfo.bmiHeader.biHeight; 
	}

	bResult = WriteFile(hFile, 
					    (LPVOID)pBits, 
						m_dwSizeImage,
						&dwNumberOfBytesWritten, 
						0); 
	if (!bResult)
	{
		MessageBox (IDS_ERR_COULDNOTSAVEIMAGE);
		return FALSE;
	}
	return TRUE;
}

//
// SaveIcon()
//

BOOL CIconEdit::SaveIcon(HANDLE hFile)
{
	DWORD dwNumberOfBytesWritten;
	WORD  nWriteIt = 0;

	//
	// Writing the Header
	//

	BOOL bResult = WriteFile(hFile, 
							 &nWriteIt, 
							 sizeof(nWriteIt),
							 &dwNumberOfBytesWritten, 
							 0); 
	nWriteIt = 1;
	bResult = WriteFile(hFile, 
						&nWriteIt, 
						sizeof(nWriteIt),
						&dwNumberOfBytesWritten, 
						0); 
	nWriteIt = 1;
	bResult = WriteFile(hFile, 
						&nWriteIt, 
						sizeof(nWriteIt),
						&dwNumberOfBytesWritten, 
						0); 

	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		HBITMAP hImage;
		HBITMAP hMask;
		GetBitmaps(hImage, hMask);

		LPVOID pColorBits;
		HBITMAP hColor = CreateDIBSection(hDC, &m_pntEdit.m_bmInfo, DIB_RGB_COLORS, &pColorBits, 0, 0); 

		LPVOID pMaskBits;
		HBITMAP hMask2 = CreateDIBSection(hDC, &m_pntEdit.m_bmInfo, DIB_RGB_COLORS, &pMaskBits, 0, 0); 

		//
		// Get the Icon information and bits
		//


		//
		// Writing IconDirEntry
		//
	}
	return TRUE;
}

//
// CreateImageBitmap
//

BOOL CIconEdit::CreateImageBitmap(HBITMAP& hBmp, HBITMAP& hDib, BITMAPINFO& bmi, LPVOID& ppvBits, int nBitCount)
{
	HDC hDC = GetDC(m_pvWin.hWnd());
	if (hDC)
	{
		HDC hDCSrc = CreateCompatibleDC(hDC);
		if (hDCSrc)
		{
			HDC hDCDest = CreateCompatibleDC(hDC);
			if (hDCDest)
			{
				hBmp = CreateCompatibleBitmap(hDC, 
		 									  m_pntEdit.m_size.cx, 
											  m_pntEdit.m_size.cy);
				if (hBmp)
				{
					memcpy (&bmi, &(m_pntEdit.m_bmInfo), sizeof(BITMAPINFO));

					if (-1 != nBitCount)
						bmi.bmiHeader.biBitCount = nBitCount;

					hDib = CreateDIBSection(hDC, &bmi, DIB_RGB_COLORS, &ppvBits, 0, 0); 
					if (hDib)
					{
						HBITMAP hBitmapSrcOld = SelectBitmap(hDCSrc, hBmp);
						HBITMAP hBitmapDestOld = SelectBitmap(hDCDest, hDib);

						HPALETTE hPalSrcOld;
						HPALETTE hPalDestOld;
						if (g_hPal)
						{
							hPalSrcOld = SelectPalette(hDCSrc, g_hPal, FALSE);
							RealizePalette(hDCSrc);
							hPalDestOld = SelectPalette(hDCDest, g_hPal, FALSE);
							RealizePalette(hDCDest);
						}

						CRect rcClient(0, 
									   m_pntEdit.m_size.cx,
									   0, 
									   m_pntEdit.m_size.cy);

						m_pvWin.PaintDibSection(hDCSrc, rcClient);
						m_pvWin.PaintDibSection(hDCDest, rcClient);
						
						if (g_hPal)
						{
							SelectPalette(hDCSrc, hPalSrcOld, FALSE);
							SelectPalette(hDCDest, hPalDestOld, FALSE);
						}

						SelectBitmap(hDCSrc, hBitmapSrcOld);
						SelectBitmap(hDCDest, hBitmapDestOld);
					}
				}
				DeleteDC(hDCDest);
			}
			DeleteDC(hDCSrc);
		}
		ReleaseDC(m_pvWin.hWnd(), hDC);
	}
	return TRUE;
}

// 
// Copy
//

BOOL CIconEdit::Copy()
{
	HBITMAP hBmp;
	HBITMAP hDib;
	BITMAPINFO bmi;
	LPVOID ppData;
	if (!CreateImageBitmap(hBmp, hDib, bmi, ppData))
	{
		MessageBox (IDS_ERR_COULDNOTCREATEBITMAP, MB_ICONSTOP);
		return FALSE;
	}

	if (!OpenClipboard(m_hWnd))
	{
		MessageBox (IDS_ERR_COULDNOTOPENCLIPBOARD, MB_ICONSTOP);
		return FALSE;
	}

	EmptyClipboard();

	HANDLE hBM = SetClipboardData(CF_BITMAP, hBmp);
	if (NULL == hBM)
	{
		DeleteBitmap(hBmp);
		MessageBox (IDS_ERR_COULDNOTSETCLIPBOARD, MB_ICONSTOP);
		CloseClipboard();
		return FALSE;
	}

	if (!CloseClipboard())
	{
		MessageBox (IDS_ERR_COULDNOTCLOSECLIPBOARD, MB_ICONSTOP);
		return FALSE;
	}
	return TRUE;
}

//
// Paste
//

void CIconEdit::Paste()
{
	if (!OpenClipboard(m_hWnd))
	{
		MessageBox (IDS_ERR_COULDNOTOPENCLIPBOARD, MB_ICONSTOP);
		return;
	}

	IPicture* pPic = NULL;
	HBITMAP hNew = NULL;
	HBITMAP hSrcOld;
	HBITMAP hDestOld;
	HDC hDCSrc = NULL;
	HDC hDCDest = NULL;
	HDC hDC = NULL;
	BITMAP bmInfo;
	HBITMAP hBMClip = (HBITMAP)GetClipboardData(CF_BITMAP); 
	if (NULL == hBMClip)
	{
		MessageBox (IDS_ERR_COULDNOTGETCLIPBOARDDATA, MB_ICONSTOP);
		goto Cleanup;
	}

	hDC = GetDC(NULL);
	if (NULL == hDC)
		goto Cleanup;

	GetObject(hBMClip, sizeof(BITMAP), &bmInfo);

	hNew = CreateCompatibleBitmap(hDC, bmInfo.bmWidth, bmInfo.bmHeight);
	if (NULL == hNew)
		goto Cleanup;

	hDCSrc = CreateCompatibleDC(hDC);
	if (NULL == hDCSrc)
		goto Cleanup;

	hDCDest = CreateCompatibleDC(hDC);
	if (NULL == hDCDest)
		goto Cleanup;

	hSrcOld = SelectBitmap(hDCSrc, hBMClip);
	hDestOld = SelectBitmap(hDCDest, hNew);

	BitBlt(hDCDest, 0, 0, bmInfo.bmWidth, bmInfo.bmHeight, hDCSrc, 0, 0, SRCCOPY);

	SelectBitmap(hDCSrc, hSrcOld);
	SelectBitmap(hDCDest, hDestOld);

	pPic = OlePictureFromHBitmap(hNew);
	if (NULL == pPic)
	{
		MessageBox (IDS_ERR_COULDNOTCREATEIPICTURE, MB_ICONSTOP);
		return;
	}

	if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_AUTOCREATEMASKCONFIRM))
		PastePicture(pPic, TRUE);
	else
		PastePicture(pPic, FALSE);

Cleanup:
	if (hDCSrc)
		DeleteDC(hDCSrc);

	if (hDCDest)
		DeleteDC(hDCDest);
	
	if (hDC)
		ReleaseDC(NULL, hDC);

	if (pPic)
		pPic->Release();

	if (!CloseClipboard())
		MessageBox (IDS_ERR_COULDNOTCLOSECLIPBOARD, MB_ICONSTOP);
}

//
// ClearPicture
//

BOOL CIconEdit::ClearPicture()
{
	m_bDirty[m_nImageIndex] = TRUE;

	int x, y;
	UndoPush();

	for (y = 0; y < m_pntEdit.m_size.cy; ++y)
	{
		for (x = 0; x < m_pntEdit.m_size.cx; ++x)
			m_pntEdit.m_pcrImageData[x+m_pntEdit.m_size.cx*y] = 0xFFFFFFFF;
	}
	m_pntEdit.RefreshImage();
	return TRUE;
}

//
// CreateMaskBitmap
//

HBITMAP CreateMaskBitmap(HDC hDC, HBITMAP hBitmap, int nWidth, int nHeight)
{
    // Create memory DCs to work with.
    HDC hDCMask = ::CreateCompatibleDC(hDC);
    HDC hDCImage = ::CreateCompatibleDC(hDC);

    // Create a monochrome bitmap for the mask.
	HBITMAP hbmMask = ::CreateBitmap(nWidth, nHeight, 1, 1, 0);
    
	// Select the mono bitmap into its DC.
    HBITMAP hbmMaskOld = SelectBitmap(hDCMask, hbmMask);
    
	// Select the image bitmap into its DC.
    HBITMAP hbmImageOld = SelectBitmap(hDCImage, hBitmap);
    
	// Set the transparency color to be the top-left pixel.
	COLORREF crColor = ::GetPixel(hDCImage, 0, 0);
	int nHit = 1;
	if (crColor == GetPixel(hDCImage, nWidth-1, nHeight-1)) 
		nHit++;
	if (crColor == GetPixel(hDCImage, 0, nHeight-1)) 
		nHit++;
	if (crColor == GetPixel(hDCImage, nWidth-1, nHeight-1)) 
		nHit++;

	if (nHit > 1)
		::SetBkColor(hDCImage, crColor);
	else 
		::SetBkColor(hDCImage, RGB(192,192,192));

	SetTextColor(hDCImage, RGB(255,255,255));

    // Make the mask.
    ::BitBlt(hDCMask, 0, 0, nWidth, nHeight, hDCImage, 0, 0, SRCCOPY);

    // Tidy up.
    SelectBitmap(hDCMask, hbmMaskOld);
    SelectBitmap(hDCImage, hbmImageOld);
    ::DeleteDC(hDCMask);
    ::DeleteDC(hDCImage);
	return hbmMask;
}
