#ifndef UTILITY_INCLUDED
#define UTILITY_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Globals.h"

interface IActiveBar2;
class CPictureHolder;
class CFlickerFree;
class CBar;

//
// CRect
//

class CRect : public RECT
{
public:
	CRect();

	CRect(const	CRect& rhs);
	CRect(const	RECT& rhs);
	CRect(LPCRECTL rhs);

	CRect(const int& left, 
		  const int& right, 
		  const int& top, 
		  const int& bottom);

// Attributes
	int operator==(CRect& rhs);
	int operator!=(const CRect& rhs);
	CRect& operator=(const CRect& rcBound);
	CRect& operator=(RECT& rhs);

	void SetEmpty();

	void Set(const int& left, 
			 const int& right, 
			 const int& top, 
			 const int& bottom);

// Operations
	BOOL IsEmpty();
	int Height() const;
	int Width() const;
	SIZE Size() const;
	void Inflate(int x, int y);
	void Offset(int x, int y);
	BOOL Intersect(const CRect& rc, const CRect& rc2);
	
	POINT& TopLeft();
	POINT& BottomRight();
#ifdef _DEBUG
	void Dump(TCHAR* szError) const;
#endif
};

//
// CRect
//

inline CRect::CRect()
{
	::SetRectEmpty(this);
}

inline CRect::CRect(const int& nLeft, 
				    const int& nRight, 
					const int& nTop, 
					const int& nBottom)
{
	::SetRect(this, nLeft, nTop, nRight, nBottom);
}

inline CRect::CRect(const CRect& rhs)
{
	::SetRect(this, rhs.left, rhs.top, rhs.right, rhs.bottom);
}

inline CRect::CRect(const RECT& rhs)
{
	::SetRect(this, rhs.left, rhs.top, rhs.right, rhs.bottom);
}

inline CRect::CRect(LPCRECTL rhs)
{
	::SetRect(this, rhs->left, rhs->top, rhs->right, rhs->bottom);
}

inline BOOL CRect::IsEmpty()
{
	return IsRectEmpty(this);
}

//
// SetRectEmpty
//

inline void CRect::SetEmpty()
{
	::SetRectEmpty(this);
}

//
// SetRect
//

inline void CRect::Set(const int& nLeft, 
					   const int& nTop, 
					   const int& nRight, 
					   const int& nBottom)
{
	::SetRect(this, nLeft, nTop, nRight, nBottom);
}

//
// operator==
//

inline int CRect::operator==(CRect& rhs)
{
	return memcmp(this, &rhs, sizeof(RECT)) == 0;
}

//
// operator!=
//

inline int CRect::operator!=(const CRect& rhs)
{
	return memcmp(this, &rhs, sizeof(RECT)) != 0;
}

//
// operator=
//

inline CRect& CRect::operator=(const CRect& rhs)
{
	if (&rhs == this)
		return *this;
	memcpy(this, &rhs, sizeof(CRect));
	return *this;
}

//
// operator=
//

inline CRect& CRect::operator=(RECT& rhs)
{
	memcpy(this, &rhs, sizeof(CRect));
	return *this;
}

//
// Height
//

inline int CRect::Height() const
{
	return bottom - top;
}

//
// Width
//

inline int CRect::Width() const 
{
	return right - left;
}

//
// Size
//

inline SIZE CRect::Size() const
{
	SIZE size;
	size.cx = Width();
	size.cy = Height();
	return size;
}

//
// Inflate
//

inline void CRect::Inflate(int x, int y)
{ 
	::InflateRect(this, x, y); 
}

//
// Offset
//

inline void CRect::Offset(int x, int y)
{
	::OffsetRect(this, x, y);
}

//
// Intersect
//

inline BOOL CRect::Intersect(const CRect& rc, const CRect& rc2)
{
	return IntersectRect(this, &rc, &rc2);
}

//
// TopLeft
//

inline POINT& CRect::TopLeft()
{
	return *((POINT*)this);
}

//
// BottomRight
//

inline POINT& CRect::BottomRight()
{
	return *((POINT*)(this+1));
}

#ifdef _DEBUG
inline void CRect::Dump(TCHAR* szError) const
{
	TRACE1(10, _T("%s\n"), szError);
	TRACE2(10, _T("CRect: left: %i top: %i\n"), left, top);
	TRACE2(10, _T("		  right: %i bottom: %i\n"), right, bottom);
}
#endif

//
// DDString
//

class DDString
{
public:
	DDString(TCHAR* szString);
	
	DDString()
		: m_szBuffer(NULL) 
	{
	}

	~DDString();

	BSTR AllocSysString () const;
	int GetLength();
	BOOL LoadString (UINT nResourceId);
	void Format(LPCTSTR szFormat, ...);
	void Format(UINT nFormatId, ...);
	operator LPCTSTR () const;
	operator LPTSTR () const;
	const DDString& operator=(const TCHAR* sz);

private:
	void FormatV(LPCTSTR lpszFormat, va_list argList);
	TCHAR* m_szBuffer;
};

inline int DDString::GetLength()
{
	return _tcslen(m_szBuffer) + 1;
}

inline DDString::operator LPCTSTR() const
{
	return m_szBuffer;
}

inline DDString::operator LPTSTR() const
{
	return m_szBuffer;
}

inline const DDString& DDString::operator=(const TCHAR* sz)
{
	delete [] m_szBuffer;
	m_szBuffer = new TCHAR[_tcslen(sz)+1];
	_tcscpy(m_szBuffer, sz); 
	return *this; 
}

//
// CColor
//

class CColor
{
public:
	CColor ();
	CColor (COLORREF crColor);

	enum
	{
		eLSMAX = 240,
		eHUEMAX = 240,
		eRGBMAX = 255,
		eUNDEFINED = (eLSMAX * 2 / 3)
	};

	void GetHSL(short& nHue, short& nLum, short& nSat);
	void SetHSL(short nHue, short nLum, short nSat);
	
	void GetRGB(short& nRed, short& nGreen, short& nBlue);
	void SetRGB(short nRed, short nGreen, short nBlue);

	COLORREF Color();
	void Color(COLORREF crColor);

	void ModifyColor(double dHuePercentage, double dLumPercentage, double dSatPercentage);

private:
	void RGBToHSL();
	void HSLToRGB();

	double m_nRed;
	double m_nGreen;
	double m_nBlue;
	double m_nHue;
	double m_nLum;
	double m_nSat;
};

//
// Color
// 

inline COLORREF CColor::Color()
{
	return RGB(m_nRed, m_nGreen, m_nBlue);
}

inline void CColor::Color(COLORREF crColor)
{
	m_nRed = GetRValue(crColor);
	m_nGreen = GetGValue(crColor);
	m_nBlue = GetBValue(crColor);
	RGBToHSL();
}

inline void CColor::ModifyColor(double dHuePercentage, double dLumPercentage, double dSatPercentage)
{
	m_nHue = dHuePercentage * m_nHue;
	m_nLum = dLumPercentage * m_nLum;
	m_nSat = dSatPercentage * m_nSat;
	HSLToRGB();
}

//
// GetHLS
//

inline void CColor::GetHSL(short& nHue, short& nLum, short& nSat)
{
	nHue = (short)(m_nHue + 0.5);
	nLum = (short)(m_nLum + 0.5);
	nSat = (short)(m_nSat + 0.5);
}

//
// SetHLS
//

inline void CColor::SetHSL(short nHue, short nLum, short nSat)
{
	m_nHue = (double)nHue;
	m_nLum = (double)nLum;
	m_nSat = (double)nSat;
	HSLToRGB();
}

//
// GetRGB
//

inline void CColor::GetRGB(short& nRed, short& nGreen, short& nBlue)
{
	nRed = (short)m_nRed;
	nGreen = (short)m_nGreen;
	nBlue = (short)m_nBlue;
}

//
// SetRGB
//

inline void CColor::SetRGB(short nRed, short nGreen, short nBlue)
{
	m_nRed = (double)nRed;
	m_nGreen = (double)nGreen;
	m_nBlue = (double)nBlue;
	RGBToHSL();
}

inline void ScreenToClient(HWND hWnd, CRect& rc)
{ 
	assert(::IsWindow(hWnd)); 
	::ScreenToClient(hWnd, (LPPOINT)&rc);
	::ScreenToClient(hWnd, ((LPPOINT)&rc)+1); 
}

inline void ClientToScreen(HWND hWnd, CRect& rc)
{ 
	assert(::IsWindow(hWnd)); 
	::ClientToScreen(hWnd, (LPPOINT)&rc);
	::ClientToScreen(hWnd, ((LPPOINT)&rc)+1); 
}

inline BOOL FillSolidRect(HDC		   hDC, 
						  const CRect& rc, 
						  COLORREF	   crColor)
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
	return TRUE;
}

inline BOOL FillSolidRect(HDC      hDC, 
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
		ExtTextOut(hDC, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);
	}
	else
	{ 
		HBRUSH hBrush = CreateSolidBrush(crColor);
		FillRect(hDC, &rc, hBrush);
		DeleteBrush(hBrush);
	}
	return TRUE;
}

inline LPTSTR GetLastErrorMsg()
{
	LPVOID lpMsgBuf;
	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
				  NULL,
				  GetLastError(),
				  MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
				  (LPTSTR)&lpMsgBuf,    
				  0,    
				  NULL);
	return (LPTSTR)lpMsgBuf;
}
//
// GDI Utilities
//

//
// DrawDitherRect
//

void DrawDitherRect(HDC hDC, CRect rcDither, COLORREF crColor);

HBITMAP CreateBitmapFromPicture(IPicture* pPict, BOOL bMonochrome);

HBRUSH CreateHalftoneBrush(BOOL bMonochrome = TRUE);

HBITMAP CreateDitherBitmap(BOOL bMonochrome = TRUE);

BOOL FillRgnBitmap(HDC hDC, HRGN hRgn, HBITMAP hBmp, const CRect& rc, int nWidth, int nHeight);

HRESULT StWriteBSTR(IStream* pStream, BSTR bstr);

HRESULT StReadBSTR(IStream* pStream, BSTR& bstr);

int GetVariantSize(VARIANT& vProperty);

HRESULT StWriteVariant(IStream* pStream, VARIANT& vProperty);

HRESULT StReadVariant(IStream* pStream, VARIANT& vProperty);

HWND WindowFromPointEx(POINT ptScreen);

HWND GetParentOwner(HWND hWnd);

HWND GetTopLevelParent(HWND hWnd);

BOOL IsDescendant(HWND hWndParent, HWND hWndChild);

LPTSTR LoadStringRes(UINT id);

extern HBITMAP g_lastBitmapSel;
extern HDC hMemDC;

inline void SelectMemDCBitmap(HBITMAP hBmp)
{
	if (hBmp==g_lastBitmapSel)
		return;
	SelectObject(hMemDC,hBmp);
	g_lastBitmapSel=hBmp;
}

inline HDC GetMemDC()
{
	
	
	if (NULL == hMemDC)
	{
		HDC hdc = GetDC(GetDesktopWindow());
		hMemDC = CreateCompatibleDC(hdc);
		ReleaseDC(GetDesktopWindow(),hdc);
	}
	return hMemDC;
}

inline long GetStringSize(BSTR bstr)
{
	if (bstr)
	{
		int nSize = wcslen(bstr);
		if (nSize > 0)
			return (((nSize + 1) * sizeof(WCHAR)) + sizeof(short));
	}
	return sizeof(short);
}

long GetPaletteSize(HPALETTE hPal);

void WINAPI CopyDISPVARIANT(void* pvDest, const void* pvSrc, DWORD dwCopy);

short OCXShiftState();

BOOL CreateGradient(HDC			  hDC, 
					HPALETTE	  hPal, 
					CFlickerFree& ffGradient, 
					COLORREF      crBegColor, 
					COLORREF      crEndColor, 
					BOOL		  bVertical, 
					int			  nWidth, 
					int			  nHeight);

void DrawFrameRect(CBar* pBar, HDC hDC, CRect rc, int nBorderCX, int nBorderCY);

HRESULT PropertyGet(LPDISPATCH pDispatch, OLECHAR* szProperty, VARIANT& vProperty);
HRESULT PropertyPut(LPDISPATCH pDispatch, OLECHAR* szProperty, VARIANT& vProperty);

void DrawCross(HDC hDC, long x1, long y1, long x2, long y2, long nThickness, COLORREF crColor);

void DrawVerticalText(HDC hDC, LPCTSTR szCaption, int nLen, const CRect& rcOut, BSTR bstrCaption = NULL);

void DrawExpandBitmap(HDC hDC, CRect rcBitmap, HBITMAP hBitmap, int nOffset, int nSrcWidth, int nSrcHeight);

BOOL BitBltMasked(HDC hDC, int left, int top, int width, int height, HDC hMemDC, int x, int y, int nSrcWidth, int nSrcHeight, COLORREF rcFore);

//
// FindLastToolId
//

long FindLastToolId(IActiveBar2* pActiveBar);

//
// GetAmbientProperty
//

BOOL GetAmbientProperty(LPDISPATCH pDispAmbient, DISPID dispid, VARTYPE vt, LPVOID pData);

//
// GetExtenderDisp
//

LPDISPATCH GetExtenderDisp(LPDISPATCH pInner);

CHAR WINAPI TranslateKeyCode(DWORD ch);

COLORREF GetShadeColor(HDC hDC, COLORREF crIn, int nValue);

COLORREF NewColor(HDC hDC, COLORREF crIn, int nValue);

COLORREF GrayColor(HDC hDC, COLORREF crIn, int nValue);
#endif