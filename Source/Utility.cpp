//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Flicker.h"
#include "Fontholder.h"
#include "Support.h"
#include "Bar.h"
#ifdef _SCRIPTSTRING
#include "ScriptString.h"
#endif
#include <assert.h>
#include <stdio.h>
#include "Utility.h"
extern HINSTANCE g_hInstance;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

HBITMAP g_lastBitmapSel=NULL;
HDC hMemDC = NULL;

//
// DDString
//

DDString::DDString(TCHAR* szString)
	: m_szBuffer(NULL) 
{
	int nLen = lstrlen(szString);
	if (nLen > 0)
	{
		m_szBuffer = new TCHAR[nLen+1];
		if (m_szBuffer)
			lstrcpy(m_szBuffer, szString);
	}
}
	
//
// ~DDString
//

DDString::~DDString()
{
	delete [] m_szBuffer;
}

//
// AllocSysString
//

BSTR DDString::AllocSysString() const
{
	MAKE_WIDEPTR_FROMTCHAR(wString, m_szBuffer);
	return SysAllocString(wString);
}

//
// LoadString
//

BOOL DDString::LoadString (UINT nResourceId)
{
	delete [] m_szBuffer;
	m_szBuffer = NULL;
	static TCHAR szBuffer[512];
	::LoadString(g_hInstance, nResourceId, szBuffer, 512);
	m_szBuffer = new TCHAR[lstrlen(szBuffer)+1];
	if (NULL == m_szBuffer)
		return FALSE;
	lstrcpy(m_szBuffer, szBuffer);
	return TRUE;
}

#define FORCE_ANSI      0x10000
#define FORCE_UNICODE   0x20000

void DDString::FormatV(LPCTSTR lpszFormat, va_list argList)
{
	va_list argListSave = argList;

	// make a guess at the maximum length of the resulting string
	int nMaxLen = 0;
	for (LPCTSTR lpsz = lpszFormat; *lpsz != '\0'; lpsz = _tcsinc(lpsz))
	{
		// handle '%' character, but watch out for '%%'
		if (*lpsz != '%' || *(lpsz = _tcsinc(lpsz)) == '%')
		{
			nMaxLen += _tclen(lpsz);
			continue;
		}

		int nItemLen = 0;

		// handle '%' character with format
		int nWidth = 0;
		for (; *lpsz != '\0'; lpsz = _tcsinc(lpsz))
		{
			// check for valid flags
			if (*lpsz == '#')
				nMaxLen += 2;   // for '0x'
			else if (*lpsz == '*')
				nWidth = va_arg(argList, int);
			else if (*lpsz == '-' || *lpsz == '+' || *lpsz == '0' ||
				*lpsz == ' ')
				;
			else // hit non-flag character
				break;
		}
		// get width and skip it
		if (nWidth == 0)
		{
			// width indicated by
			nWidth = _ttoi(lpsz);
			for (; *lpsz != '\0' && _istdigit(*lpsz); lpsz = _tcsinc(lpsz))
				;
		}
		assert(nWidth >= 0);

		int nPrecision = 0;
		if (*lpsz == '.')
		{
			// skip past '.' separator (width.precision)
			lpsz = _tcsinc(lpsz);

			// get precision and skip it
			if (*lpsz == '*')
			{
				nPrecision = va_arg(argList, int);
				lpsz = _tcsinc(lpsz);
			}
			else
			{
				nPrecision = _ttoi(lpsz);
				for (; *lpsz != '\0' && _istdigit(*lpsz); lpsz = _tcsinc(lpsz))
					;
			}
			assert(nPrecision >= 0);
		}

		// should be on type modifier or specifier
		int nModifier = 0;
		switch (*lpsz)
		{
		// modifiers that affect size
		case 'h':
			nModifier = FORCE_ANSI;
			lpsz = _tcsinc(lpsz);
			break;
		case 'l':
			nModifier = FORCE_UNICODE;
			lpsz = _tcsinc(lpsz);
			break;

		// modifiers that do not affect size
		case 'F':
		case 'N':
		case 'L':
			lpsz = _tcsinc(lpsz);
			break;
		}

		// now should be on specifier
		switch (*lpsz | nModifier)
		{
		// single characters
		case 'c':
		case 'C':
			nItemLen = 2;
			va_arg(argList, TCHAR);
			break;
		case 'c'|FORCE_ANSI:
		case 'C'|FORCE_ANSI:
			nItemLen = 2;
			va_arg(argList, char);
			break;
		case 'c'|FORCE_UNICODE:
		case 'C'|FORCE_UNICODE:
			nItemLen = 2;
			va_arg(argList, WCHAR);
			break;

		// strings
		case 's':
		{
			LPCTSTR pstrNextArg = va_arg(argList, LPCTSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6;  // "(null)"
			else
			{
			   nItemLen = lstrlen(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
			break;
		}

		case 'S':
		{
#ifndef _UNICODE
			LPWSTR pstrNextArg = va_arg(argList, LPWSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6;  // "(null)"
			else
			{
			   nItemLen = wcslen(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
#else
			LPCSTR pstrNextArg = va_arg(argList, LPCSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6; // "(null)"
			else
			{
			   nItemLen = lstrlenA(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
#endif
			break;
		}

		case 's'|FORCE_ANSI:
		case 'S'|FORCE_ANSI:
		{
			LPCSTR pstrNextArg = va_arg(argList, LPCSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6; // "(null)"
			else
			{
			   nItemLen = lstrlenA(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
			break;
		}

		case 's'|FORCE_UNICODE:
		case 'S'|FORCE_UNICODE:
		{
			LPWSTR pstrNextArg = va_arg(argList, LPWSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6; // "(null)"
			else
			{
			   nItemLen = wcslen(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
			break;
		}
		}

		// adjust nItemLen for strings
		if (nItemLen != 0)
		{
			nItemLen = max(nItemLen, nWidth);
			if (nPrecision != 0)
				nItemLen = min(nItemLen, nPrecision);
		}
		else
		{
			switch (*lpsz)
			{
			// integers
			case 'd':
			case 'i':
			case 'u':
			case 'x':
			case 'X':
			case 'o':
				va_arg(argList, int);
				nItemLen = 32;
				nItemLen = max(nItemLen, nWidth+nPrecision);
				break;

			case 'e':
			case 'f':
			case 'g':
			case 'G':
				va_arg(argList, double);
				nItemLen = 128;
				nItemLen = max(nItemLen, nWidth+nPrecision);
				break;

			case 'p':
				va_arg(argList, void*);
				nItemLen = 32;
				nItemLen = max(nItemLen, nWidth+nPrecision);
				break;

			// no output
			case 'n':
				va_arg(argList, int*);
				break;

			default:
				assert(FALSE);  // unknown formatting option
			}
		}

		// adjust nMaxLen for output nItemLen
		nMaxLen += nItemLen;
	}

	delete [] m_szBuffer;
	m_szBuffer = new TCHAR[nMaxLen+1];
	_vstprintf(m_szBuffer, lpszFormat, argListSave);
	va_end(argListSave);
}

void DDString::Format(LPCTSTR szFormat, ...)
{
	va_list argList;
	va_start(argList, szFormat);
	FormatV(szFormat, argList);
	va_end(argList);
}

void DDString::Format(UINT nFormatId, ...)
{
	DDString strFormat;
	strFormat.LoadString(nFormatId);

	va_list argList;
	va_start(argList, nFormatId);
	FormatV(strFormat, argList);
	va_end(argList);
}

//
// CColor
//

CColor::CColor()
{
	m_nRed = 0;
	m_nGreen = 0;
	m_nBlue = 0;
	RGBToHSL();
}

CColor::CColor(COLORREF crColor)
{
	m_nRed = GetRValue(crColor);
	m_nGreen = GetGValue(crColor);
	m_nBlue = GetBValue(crColor);
	RGBToHSL();
}

void CColor::RGBToHSL()
{
	double cMax = max(max(m_nRed, m_nGreen), m_nBlue);
	double cMin = min(min(m_nRed, m_nGreen), m_nBlue);

	m_nLum = (((cMax + cMin) * eLSMAX) + eRGBMAX ) / (2 * eRGBMAX);

	if (cMax == cMin) 
	{   
		//
		// Achromatic case 
		//
		// r = g = b
		//

		m_nSat = 0;
		m_nHue = eUNDEFINED;
	}
	else 
	{     
		//
		// Chromatic Case
		//
		// Saturation
		//

		if (m_nLum <= (eLSMAX / 2))
			m_nSat = (((cMax - cMin) * eLSMAX) + ((cMax + cMin) / 2)) / (cMax + cMin);
		else
			m_nSat = (((cMax - cMin) * eLSMAX) + ((2 * eRGBMAX - cMax - cMin) / 2)) / (2 * eRGBMAX - cMax - cMin);

		//
		// Hue
		//

		double nRDelta = (((cMax - m_nRed) * (eHUEMAX/6)) + ((cMax - cMin) / 2)) / (cMax - cMin);
		double nGDelta = (((cMax - m_nGreen) * (eHUEMAX/6)) + ((cMax - cMin) / 2)) / (cMax - cMin);
		double nBDelta = (((cMax - m_nBlue) * (eHUEMAX/6)) + ((cMax - cMin) / 2)) / (cMax - cMin);

		if (m_nRed == cMax)
			m_nHue = nBDelta - nGDelta;
		else if (m_nGreen == cMax)
			m_nHue = (eHUEMAX / 3) + nRDelta - nBDelta;
		else 
			// m_nBlue == cMax
			m_nHue = ((2 * eHUEMAX) / 3) + nGDelta - nRDelta;

		if (m_nHue < 0)
			m_nHue += eHUEMAX;

		if (m_nHue > eHUEMAX)
			m_nHue -= eHUEMAX;
	}
}

//
// Utility routine for HueToRGB
//

inline static double HueToRGB(double n1, double n2, double nHue)
{ 
	//
	// range check: note values passed add/subtract thirds of range
	if (nHue < 0)
		nHue += CColor::eHUEMAX;

	if (nHue > CColor::eHUEMAX)
		nHue -= CColor::eHUEMAX;

    //
	// return r, g, or b value from this tridrant
	//

	if (nHue < (CColor::eHUEMAX / 6))
		return (n1 + (((n2 - n1) * nHue + (CColor::eHUEMAX / 12)) / (CColor::eHUEMAX / 6)));

	if (nHue < (CColor::eHUEMAX / 2))
		return n2;

	if (nHue < ((CColor::eHUEMAX * 2) / 3))
		return (n1 + (((n2 - n1) * (((CColor::eHUEMAX * 2) / 3) - nHue) + (CColor::eHUEMAX / 12)) / (CColor::eHUEMAX / 6))); 
	else
		return n1;
} 

void CColor::HSLToRGB()
{
	if (0 == m_nSat) 
	{            
		//
		// Achromatic case
		//

		m_nRed = m_nGreen = m_nBlue = (m_nLum * eRGBMAX) / eLSMAX;
		if (eUNDEFINED != m_nHue) 
		{
			// ERROR
			m_nRed = m_nGreen = m_nBlue = 0;
		}
	}
	else  
	{
		// chromatic case
		// set up magic numbers

		double nMagic2;
		if (m_nLum <= (eLSMAX / 2))
			nMagic2 = (m_nLum * (eLSMAX + m_nSat) + (eLSMAX / 2)) / eLSMAX;
		else
			nMagic2 = m_nLum + m_nSat - ((m_nLum * m_nSat) + (eLSMAX / 2)) / eLSMAX;

		double nMagic1 = 2 * m_nLum - nMagic2;

		// Get RGB, change units from HLSMAX to RGBMAX
		
		m_nRed =   (HueToRGB(nMagic1, nMagic2, m_nHue + (eHUEMAX / 3)) * eRGBMAX + (eHUEMAX / 2)) / eHUEMAX; 
		m_nGreen = (HueToRGB(nMagic1, nMagic2, m_nHue) * eRGBMAX + (eHUEMAX / 2)) / eHUEMAX;
		m_nBlue =  (HueToRGB(nMagic1, nMagic2, m_nHue - (eHUEMAX / 3)) * eRGBMAX + (eHUEMAX / 2)) / eHUEMAX; 
	}
}
//
// CreateBitmapFromPicture
//

HBITMAP CreateBitmapFromPicture(IPicture *ppict,BOOL monochrome)
{
	const int nHimetricPerInch = 2540;
	long hmWidth;
	long hmHeight;

	// icons have a problem with width/height calc. even if the icon is 16x16 
	ppict->get_Width(&hmWidth);
	ppict->get_Height(&hmHeight);

	HBITMAP newBitmap = NULL;
	HDC hdc=GetDC(NULL);
	if (hdc)
	{
		int w = MulDiv(GetDeviceCaps(hdc,LOGPIXELSX),hmWidth,nHimetricPerInch);
		int h = MulDiv(GetDeviceCaps(hdc,LOGPIXELSY),hmHeight,nHimetricPerInch);

		if (monochrome)
			newBitmap = CreateBitmap(w,h,1,1,NULL);
		else
			newBitmap = CreateCompatibleBitmap(hdc,w,h);
		
		// Now render into newBitmap.
		HDC hMemDC = CreateCompatibleDC(hdc);
		if (hMemDC)
		{
			HBITMAP oldBitmap = SelectBitmap(hMemDC, newBitmap);
			RECT r;
			r.left=0;
			r.top=0;
			r.right=w;
			r.bottom=h;
			// make sure icon resources render correctly
			FillRect(hMemDC,&r,(HBRUSH)(1+COLOR_BTNFACE)); 
			ppict->Render(hMemDC, 
						  0,
						  0,
						  w,
						  h,
						  0, 
						  hmHeight-1,
						  hmWidth, 
						  -hmHeight,
						  &r);
			SelectBitmap(hMemDC,oldBitmap);
			DeleteDC(hMemDC);
		}
		ReleaseDC(0,hdc);
	}
	return newBitmap;
}

//
// CreateDitherBitmap
//

HBITMAP CreateDitherBitmap(BOOL bMonochrome)
{
	HBITMAP hbm;
	
	// initialize the brushes
	
	if (bMonochrome)
	{
		WORD patGray[8];
		for (int i = 0; i < 8; i++)
			patGray[i] = (i & 1) ? 0x55 : 0xAA;
	
		hbm = CreateBitmap(8,8,1,1,patGray);
	}
	else 
	{
		struct  // BITMAPINFO with 16 colors
		{
			BITMAPINFOHEADER bmiHeader;
			RGBQUAD          bmiColors[16];
		} bmi;
		long patGray[8];
		for (int i = 0; i < 8; i++)
	   		patGray[i] = (i & 1) ? 0xAAAA5555L : 0x5555AAAAL;
	
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
		hbm = CreateDIBitmap(hDC, &bmi.bmiHeader, CBM_INIT,(LPBYTE)patGray, (LPBITMAPINFO)&bmi, DIB_RGB_COLORS);
		ReleaseDC(NULL, hDC);
	}
	return hbm;
}

//
// CreateHalftoneBrush
//

HBRUSH CreateHalftoneBrush(BOOL bMonochrome)
{
	HBRUSH hbrDither;
	HBITMAP hbmGray = CreateDitherBitmap(bMonochrome);
	// Mono DC and Bitmap for disabled image
	if (hbmGray != NULL)
	{
		hbrDither = ::CreatePatternBrush(hbmGray);
		DeleteObject(hbmGray);
	}
	return hbrDither;
}

//
// FillRgnBitmap
//

BOOL FillRgnBitmap(HDC hDC, HRGN hRgn, HBITMAP hBmp, const CRect& rc, int nWidth, int nHeight)
{
	HBITMAP hBitmapOld;
	HRGN hSave = NULL;
	BOOL bResult;
	HDC hMemDC;
	int nLeftOffset;
	int nTopOffset;
	int nResult;
	int nX, nY;

	if (nWidth <= 0)
		return FALSE;

	if (nHeight <= 0)
		return FALSE;
	
	hSave = CreateRectRgn(0, 0, 1, 1);
	if (hSave)
		nResult = GetClipRgn(hDC, hSave);
		
	if (hRgn)
		ExtSelectClipRgn(hDC, hRgn, RGN_AND);
	else 
		IntersectClipRect(hDC, rc.left, rc.top, rc.right, rc.bottom);

	hMemDC = CreateCompatibleDC(hDC);
	if (NULL == hMemDC)
	{
		bResult = FALSE;
		goto Cleanup; // failed
	}

	hBitmapOld = SelectBitmap(hMemDC, hBmp);

	POINT ptOrg;
	GetBrushOrgEx(hDC, &ptOrg);

	// Offset from the left side of the bitmap to start tiling at
	nLeftOffset = (ptOrg.x - rc.left) % nWidth; 
	if (nLeftOffset < 0)
		nLeftOffset = -nLeftOffset;
	else
		nLeftOffset = nWidth - nLeftOffset;

	nTopOffset = (ptOrg.y - rc.top) % nHeight;
	if (nTopOffset < 0)
		nTopOffset = -nTopOffset;
	else
		nTopOffset = nHeight - nTopOffset;

	nY = rc.top;
	if (0 != nTopOffset)
	{
		nX = rc.left;
		if (0 != nLeftOffset)
		{
			bResult = BitBlt(hDC, nX, nY, nWidth - nLeftOffset, nHeight - nTopOffset, hMemDC, nLeftOffset, nTopOffset, SRCCOPY);
			if (!bResult)
				goto Cleanup;
			nX += nWidth - nLeftOffset;
		}
		
		for (; nX < rc.right; nX += nWidth)
		{
			bResult = BitBlt(hDC, nX, nY, nWidth, nHeight - nTopOffset, hMemDC, 0, nTopOffset, SRCCOPY);
			if (!bResult)
				goto Cleanup;
		}

		nY += nHeight - nTopOffset;
	}
	for (; nY < rc.bottom; nY += nHeight)
	{
		nX = rc.left;
		if (0 != nLeftOffset)
		{
			bResult = BitBlt(hDC, nX, nY, nWidth - nLeftOffset, nHeight, hMemDC, nLeftOffset, 0, SRCCOPY);
			if (!bResult)
				goto Cleanup;

			nX += nWidth - nLeftOffset;
		}

		for (; nX < rc.right; nX += nWidth)
		{
			bResult = BitBlt(hDC, nX, nY, nWidth, nHeight, hMemDC, 0, 0, SRCCOPY);
			if (!bResult)
				goto Cleanup;
		}
	}

	SelectBitmap(hMemDC, hBitmapOld);
	bResult = DeleteDC(hMemDC);
	assert(bResult);

Cleanup:
	if (0 == nResult)
		SelectClipRgn(hDC, NULL);
	else
		SelectClipRgn(hDC, hSave);

	if (hSave)
		bResult = DeleteRgn(hSave);
	assert(bResult);
	return bResult;
}

//
// StWriteBSTR
//

HRESULT StWriteBSTR(IStream* pStream, BSTR bstr)
{
	HRESULT hResult;
	short nSize = 0;
	if (bstr)
	{
		nSize = (wcslen(bstr)+1) * sizeof(WCHAR);
		hResult = pStream->Write(&nSize, sizeof(short), NULL);
		if (FAILED(hResult))
			return hResult;
		hResult = pStream->Write(bstr, nSize, NULL);
		if (FAILED(hResult))
			return hResult;
	}
	else
	{
		hResult = pStream->Write(&nSize, sizeof(short), NULL);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}

//
// StReadBSTR
//

HRESULT StReadBSTR(IStream* pStream, BSTR& bstr)
{
	SysFreeString(bstr);
	bstr = NULL;
	short nSize;
	HRESULT hResult = pStream->Read(&nSize, sizeof(short), NULL);
	if (FAILED(hResult))
		return hResult;

	if (nSize > 0)
	{
		bstr = SysAllocStringByteLen(NULL, nSize);
		if (NULL == bstr)
			return E_OUTOFMEMORY;

		pStream->Read(bstr, nSize, NULL);
		if (FAILED(hResult))
		{
			SysFreeString(bstr);
			bstr = NULL;
			return hResult;
		}
	}
	return NOERROR;
}

int GetVariantSize(VARIANT& vProperty)
{
	HRESULT hr;
	long nSize = 0;
	switch (vProperty.vt)
	{
	case VT_UNKNOWN:
	case VT_DISPATCH:
		{
			IPersistStream* pPersistStream;
			if (vProperty.punkVal)
			{
				hr = vProperty.punkVal->QueryInterface(IID_IPersistStream, (void**)&pPersistStream);
				if (FAILED(hr))
					return hr;
			}
			if (pPersistStream)
			{
				ULARGE_INTEGER lnSize;
				hr = pPersistStream->GetSizeMax(&lnSize);
				nSize += lnSize.LowPart;
			}
			else
				nSize += sizeof(CLSID_NULL);
			pPersistStream->Release();
			return nSize;
		}

	case VT_UI1:
	case VT_I1:
		nSize += sizeof(BYTE);
		break;
	case VT_I2:
	case VT_UI2:
	case VT_BOOL:
		nSize += sizeof(short);
		break;
	case VT_I4:
	case VT_UI4:
	case VT_R4:
	case VT_INT:
	case VT_UINT:
	case VT_ERROR:
		nSize += sizeof(long);
		break;
	case VT_R8:
	case VT_CY:
	case VT_DATE:
		nSize += sizeof(double);
		break;
	
	case VT_EMPTY:
		return sizeof(VARTYPE);

	default:
		break;
	}
	if (nSize != 0)
	{
		nSize += sizeof(VARTYPE);
		return nSize;
	}
	BSTR bstrWrite;
	VARIANT varBSTR;
	varBSTR.vt = VT_EMPTY;
	if (vProperty.vt != VT_BSTR)
	{
		hr = VariantChangeType(&varBSTR, &vProperty, VARIANT_NOVALUEPROP, VT_BSTR);
		if (FAILED(hr))
			return hr;
		bstrWrite = varBSTR.bstrVal;
	}
	else
		bstrWrite = vProperty.bstrVal;

	nSize = GetStringSize(bstrWrite);
	VariantClear(&varBSTR);
	return nSize;
}

//
// StWriteVariant
//

HRESULT StWriteVariant(IStream* pStream, VARIANT& vProperty)
{
	try
	{
		HRESULT hr = pStream->Write(&vProperty.vt, sizeof(VARTYPE), NULL);
		if (FAILED(hr))
			return hr;

		int cbWrite = 0;
		switch (vProperty.vt)
		{
		case VT_UNKNOWN:
		case VT_DISPATCH:
			{
				IPersistStream* pPersistStream;
				if (vProperty.punkVal)
				{
					hr = vProperty.punkVal->QueryInterface(IID_IPersistStream, (void**)&pPersistStream);
					if (FAILED(hr))
						return hr;
				}
				if (pPersistStream)
					hr = OleSaveToStream(pPersistStream, pStream);
				else
					hr = WriteClassStm(pStream, CLSID_NULL);
				pPersistStream->Release();
				return hr;
			}
		case VT_UI1:
		case VT_I1:
			cbWrite = sizeof(BYTE);
			break;
		case VT_I2:
		case VT_UI2:
		case VT_BOOL:
			cbWrite = sizeof(short);
			break;
		case VT_I4:
		case VT_UI4:
		case VT_R4:
		case VT_INT:
		case VT_UINT:
		case VT_ERROR:
			cbWrite = sizeof(long);
			break;
		case VT_R8:
		case VT_CY:
		case VT_DATE:
			cbWrite = sizeof(double);
			break;

		case VT_EMPTY:
			return NOERROR;

		default:
			break;
		}
		if (cbWrite != 0)
			return pStream->Write((void*) &vProperty.bVal, cbWrite, NULL);

		BSTR bstrWrite;
		VARIANT varBSTR;
		varBSTR.vt = VT_EMPTY;
		if (vProperty.vt != VT_BSTR)
		{
			hr = VariantChangeType(&varBSTR, &vProperty, VARIANT_NOVALUEPROP, VT_BSTR);
			if (FAILED(hr))
				return hr;
			bstrWrite = varBSTR.bstrVal;
		}
		else
			bstrWrite = vProperty.bstrVal;

		hr = StWriteBSTR(pStream, bstrWrite);
		VariantClear(&varBSTR);
		return hr;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// StReadVariant
//

HRESULT StReadVariant(IStream* pStream, VARIANT& vProperty)
{
	assert(pStream);
	HRESULT hr;
	hr = VariantClear(&vProperty);
	if (FAILED(hr))
		return hr;
	
	VARTYPE vtRead;
	hr = pStream->Read(&vtRead, sizeof(VARTYPE), NULL);
	if (hr == S_FALSE)
		hr = E_FAIL;

	if (FAILED(hr))
		return hr;

	vProperty.vt = vtRead;
	int cbRead = 0;
	switch (vtRead)
	{
	case VT_UNKNOWN:
	case VT_DISPATCH:
		{
			vProperty.punkVal = NULL;
			hr = OleLoadFromStream(pStream, (vtRead == VT_UNKNOWN) ? IID_IUnknown : IID_IDispatch, (void**)&vProperty.punkVal);
			if (hr == REGDB_E_CLASSNOTREG)
				hr = S_OK;
			return S_OK;
		}
	case VT_UI1:
	case VT_I1:
		cbRead = sizeof(BYTE);
		break;
	case VT_I2:
	case VT_UI2:
	case VT_BOOL:
		cbRead = sizeof(short);
		break;
	case VT_I4:
	case VT_UI4:
	case VT_R4:
	case VT_INT:
	case VT_UINT:
	case VT_ERROR:
		cbRead = sizeof(long);
		break;
	case VT_R8:
	case VT_CY:
	case VT_DATE:
		cbRead = sizeof(double);
		break;

	case VT_EMPTY:
		return NOERROR;

	default:
		break;
	}
	if (cbRead != 0)
	{
		hr = pStream->Read((void*) &vProperty.bVal, cbRead, NULL);
		if (hr == S_FALSE)
			hr = E_FAIL;
		return hr;
	}

	BSTR bstrRead = NULL;
	hr = StReadBSTR(pStream, bstrRead);
	if (FAILED(hr))
		return hr;

	vProperty.vt = VT_BSTR;
	vProperty.bstrVal = bstrRead;
	if (vtRead != VT_BSTR)
		hr = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, vtRead);
	return hr;
}

//
// GetParentOwner
//

HWND GetParentOwner(HWND hWnd)
{
	if (::GetWindowLong(hWnd, GWL_STYLE) & WS_CHILD)
		return ::GetParent(hWnd);
	return ::GetWindow(hWnd, GW_OWNER);
}

//
// GetTopLevelParent
//

HWND GetTopLevelParent(HWND hWnd) 
{
	HWND hWndParent = hWnd;
	HWND hWndT;
	while (hWndT = GetParentOwner(hWndParent))
		hWndParent = hWndT;
	return hWndParent;
}

BOOL IsDescendant(HWND hWndParent, HWND hWndChild)
{
	assert(::IsWindow(hWndParent));
	assert(::IsWindow(hWndChild));

	do
	{
		if (hWndParent == hWndChild)
			return TRUE;

		hWndChild = GetParentOwner(hWndChild);
	} while (hWndChild);

	return FALSE;
}

LPTSTR LoadStringRes(UINT id)
{
	const int MAXRESSTR_BUFFER = 512;
	static TCHAR szResBuffer[MAXRESSTR_BUFFER];
	LoadString(g_hInstance,id,szResBuffer,MAXRESSTR_BUFFER);
	return szResBuffer;
}

void WINAPI CopyDISPVARIANT(void *pvDest,const void *pvSrc,DWORD cbCopy)
{
	memcpy(pvDest,pvSrc,cbCopy);
	if (((VARIANT *)pvDest)->pdispVal)
		((VARIANT *)pvDest)->pdispVal->AddRef();
}

//
// OCXShiftState
//

short OCXShiftState()
{
    BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
    BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
    BOOL bAlt   = (GetKeyState(VK_MENU) < 0);
    return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}

//
// CreateGradient
//

BOOL CreateGradient(HDC hDC, HPALETTE hPal, CFlickerFree& ffGradient, COLORREF crBegColor, COLORREF crEndColor, BOOL bVertical, int nWidth, int nHeight)
{
	try
	{
		CRect rcBounds;
		rcBounds.Set(0, 0, nWidth, nHeight);
		HDC hDCNoFlicker = ffGradient.RequestDC(hDC, nWidth, nHeight);
		
		if (NULL == hDCNoFlicker)
			return FALSE;

		HPALETTE hPalOld;
		if (hPal)
		{
			hPalOld = SelectPalette(hDCNoFlicker, hPal, FALSE);
			RealizePalette(hDCNoFlicker);
		}

		SetBkColor(hDCNoFlicker, crBegColor);
		ExtTextOut(hDCNoFlicker, 0, 0, ETO_OPAQUE, &rcBounds, NULL, 0, NULL);
		
		int nBegRed = GetRValue(crBegColor); // red..
		int nBegGreen = GetGValue(crBegColor); // ..green
		int nBegBlue = GetBValue(crBegColor); // ..blue color vals

		int nEndRed = GetRValue(crEndColor); // red..
		int nEndGreen = GetGValue(crEndColor); // ..green
		int nEndBlue = GetBValue(crEndColor); // ..blue color vals

		int xRedDelta;
		int xGreenDelta;
		int xBlueDelta;

		int nX;
		int xDelta;
		if (bVertical)
		{
			nX = nHeight;
			xDelta= max(nHeight / 80, 1); // width of one shade band
		}
		else
		{
			nX = nWidth;
			xDelta= max(nWidth / 80, 1); // width of one shade band
			xRedDelta = max((nBegRed - nEndRed) / 80, 1);
			xGreenDelta = max((nBegGreen - nEndGreen) / 80, 1);
			xBlueDelta = max((nBegBlue - nEndBlue) / 80, 1);
		}

		int nWmx2;
		int nW2;
		while (nX >= 0) 
		{	
			//
			// Paint bands right to left
			//
			if (bVertical)
			{
				nWmx2 = (nHeight - nX) * (nHeight - nX); // w minus x squared
				nW2  = nHeight * nHeight; // w squared
			}
			else
			{
				nWmx2 = (nWidth - nX) * (nWidth - nX); // w minus x squared
				nW2  = nWidth * nWidth; // w squared
			}
			if (0 == nW2)
				break;

			if (nEndRed > nBegRed)
				xRedDelta = nEndRed - (nEndRed - nBegRed) * nWmx2 / nW2;
			else
				xRedDelta = nEndRed + (nBegRed - nEndRed) * nWmx2 / nW2; 
			
			if (nEndGreen > nBegGreen)
				xGreenDelta = nEndGreen - (nEndGreen - nBegGreen) * nWmx2 / nW2;
			else
				xGreenDelta = nEndGreen + (nBegGreen - nEndGreen) * nWmx2 / nW2; 

			if (nEndBlue > nBegBlue)
				xBlueDelta = nEndBlue - (nEndBlue - nBegBlue) * nWmx2 / nW2;
			else
				xBlueDelta = nEndBlue + (nBegBlue - nEndBlue) * nWmx2 / nW2; 

			HBRUSH hCurrent = CreateSolidBrush(RGB(xRedDelta, xGreenDelta, xBlueDelta));
			if (hCurrent)
			{
				HBRUSH hBrushOld = SelectBrush(hDCNoFlicker, hCurrent);
			
				if (bVertical)
					PatBlt(hDCNoFlicker, 0, nX, nWidth, xDelta, PATCOPY);
				else
					PatBlt(hDCNoFlicker, nX, 0, xDelta, nHeight, PATCOPY);

				SelectBrush(hDCNoFlicker, hBrushOld);
				DeleteBrush(hCurrent);
			}
			nX -= xDelta;
		}
		if (hPal)
			SelectPalette(hDCNoFlicker, hPalOld, FALSE);
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// DrawFrameRect
//

void DrawFrameRect(CBar* pBar, HDC hDC, CRect rc, int nBorderCX, int nBorderCY)
{
	RECT rcLine = rc;
	rcLine.right = rcLine.left + 1;
	FillSolidRect(hDC, rcLine, pBar->m_crBackground);
	rcLine = rc;
	rcLine.bottom = rcLine.top + 1;
	FillSolidRect(hDC, rcLine, pBar->m_crBackground);
			
	// Right and bottom
	SetBkColor(hDC,GetSysColor(COLOR_WINDOWFRAME));
	rcLine = rc;
	rcLine.left = rcLine.right - 1;
	ExtTextOut(hDC, 0, 0, ETO_OPAQUE, &rcLine, NULL, 0, NULL);
	rcLine = rc;
	rcLine.top = rcLine.bottom - 1;
	ExtTextOut(hDC, 0, 0, ETO_OPAQUE, &rcLine, NULL, 0, NULL);
	
	rc.Inflate(-1, -1);

	// Inner left and top
	rcLine = rc;
	rcLine.right = rcLine.left + 1;
	FillSolidRect(hDC, rcLine, pBar->m_crHighLight);
	rcLine = rc;
	rcLine.bottom = rcLine.top + 1;
	FillSolidRect(hDC, rcLine, pBar->m_crHighLight);

	// Right and bottom
	rcLine = rc;
	rcLine.left = rcLine.right - 1;
	FillSolidRect(hDC, rcLine, pBar->m_crShadow);
	rcLine = rc;
	rcLine.top = rcLine.bottom - 1;
	FillSolidRect(hDC, rcLine, pBar->m_crShadow);
	
	rc.Inflate(-1, -1);
	nBorderCX -= 2;	
	
	if (nBorderCX > 0)
	{
		// left and right
		RECT rcLine = rc;
		rcLine.right = rcLine.left + nBorderCX;
		FillSolidRect(hDC, rcLine, pBar->m_crBackground);
		rcLine = rc;
		rcLine.left = rcLine.right - nBorderCX;
		FillSolidRect(hDC, rcLine, pBar->m_crBackground);
	}

	nBorderCY -= 2;
	if (nBorderCY > 0)
	{
		// top and bottom
		rcLine = rc;
		rcLine.bottom = rcLine.top + nBorderCY;
		FillSolidRect(hDC, rcLine, pBar->m_crBackground);
		rcLine = rc;
		rcLine.top = rcLine.bottom - nBorderCY;
		FillSolidRect(hDC, rcLine, pBar->m_crBackground);
	}
}

//
// PropertyGet
//

HRESULT PropertyGet(LPDISPATCH pDispatch, OLECHAR* szProperty, VARIANT& vProperty)
{
	HRESULT hResult;
	try
	{
		DISPID dispidIds[1];
		OLECHAR* pszNames[] = {{szProperty}, {NULL}};

		hResult = pDispatch->GetIDsOfNames(IID_NULL, pszNames, 1, NULL, dispidIds);
		if (FAILED(hResult))
			return hResult;

		if (dispidIds[0] == (DISPID)-1)
			return E_FAIL;

		DISPPARAMS dispparams;
		VARIANT v;
		v.vt = VT_EMPTY;
		v.lVal = 0;

		// now go and get the property into a variant.
		memset(&dispparams, 0, sizeof(DISPPARAMS));
		hResult = pDispatch->Invoke(dispidIds[0], 
									IID_NULL, 
									LOCALE_USER_DEFAULT, 
									DISPATCH_PROPERTYGET, 
									&dispparams,
									&vProperty, 
									NULL, 
									NULL);
	}
	catch (...)
	{
		assert(FALSE);
		return E_FAIL;
	}
	return hResult;
}
//
// PropertyPut
//

HRESULT PropertyPut(LPDISPATCH pDispatch, OLECHAR* szProperty, VARIANT& vProperty)
{
	DISPID dispidIds[1];
	OLECHAR* pszNames[] = {{szProperty}, {NULL}};
	HRESULT hResult = pDispatch->GetIDsOfNames(IID_NULL, pszNames, 1, NULL, dispidIds);
	if (SUCCEEDED(hResult))
	{
		if (dispidIds[0] == -1)
			return E_FAIL;

		EXCEPINFO theExcepInfo;
		DISPPARAMS dispparams;
		memset(&dispparams, 0, sizeof(DISPPARAMS));
		VARIANT vResult;
		vResult.vt = VT_EMPTY;
		vResult.lVal = 0;

		//
		// If we Align to the form set the extended controls Align Property to Align Top
		// If we don't align to the form we set it to none.
		//

		dispparams.rgvarg = &vProperty;
		dispparams.cArgs = 1;

		hResult = pDispatch->Invoke(dispidIds[0], 
									IID_NULL, 
									NULL, 
									DISPATCH_PROPERTYPUT, 
									&dispparams,
									&vResult, 
									&theExcepInfo, 
									NULL);
		if (FAILED(hResult))
			hResult = (*theExcepInfo.pfnDeferredFillIn)(&theExcepInfo);
	}
	return hResult;
}

//
// DrawCross
//

void DrawCross(HDC hDC, long x1, long y1, long  x2, long y2, long nThickness, COLORREF crColor)
{
	HPEN hPenCross = CreatePen(PS_SOLID, 1, crColor);
	HPEN hPenOld = SelectPen(hDC, hPenCross);
	
	//
	// Draw diagonal lines
	//
	
	MoveToEx(hDC, x1, y1, NULL);
	LineTo(hDC, x2, y2);
	MoveToEx(hDC, x2-1, y1, NULL);
	LineTo(hDC, x1-1, y2);

	// 
	// For cross thickness 2, the diagonal expansion 
	// algorithm doesn't work so it is a special case
	// 

	if (2 == nThickness) 
	{
		MoveToEx(hDC, x1+1, y1, NULL);
		LineTo(hDC, x2+1, y2);
		MoveToEx(hDC, x2, y1, NULL);
		LineTo(hDC, x1, y2);
	}
	else
	{
		--nThickness;
		for (int cnt = 0; cnt < nThickness; cnt++)
		{
			// Horizontal step
			MoveToEx(hDC, x1 + cnt, y1, NULL);
			LineTo(hDC, x2, y2 - cnt);
			MoveToEx(hDC, x2 - 1 - cnt, y1, NULL);
			LineTo(hDC, x1-1, y2 - cnt);

			// Vertical step
			MoveToEx(hDC, x1, y1 + cnt, NULL);
			LineTo(hDC, x2 - cnt, y2);
			MoveToEx(hDC, x2-1, y1 + cnt, 0);
			LineTo(hDC, x1 - 1 + cnt, y2);
		}
	}
	SelectPen(hDC, hPenOld);
	DeletePen(hPenCross);
}

//
// DrawDitherRect
//

void DrawDitherRect(HDC hDC, CRect rcDither, COLORREF crColor)
{
	InflateRect(&rcDither, -1, -1);
	HBRUSH hBrushOld = SelectBrush(hDC, GetGlobals().GetBrushDither());
	SelectBrush(hDC, GetStockObject(NULL_PEN));

	SetTextColor(hDC,RGB(0, 0, 0));	
	SetBkColor(hDC,RGB(255, 255, 255));
	int nOpOld = SetROP2(hDC, R2_MASKPEN);
	
	Rectangle(hDC,
			  rcDither.left,
			  rcDither.top,
			  rcDither.right + 1,
			  rcDither.bottom + 1);

	SetBkColor(hDC, RGB(0, 0, 0));
	SetTextColor(hDC, crColor);
	SetROP2(hDC, R2_MERGEPEN);
	
	Rectangle(hDC,
			  rcDither.left,
			  rcDither.top,
			  rcDither.right + 1,
			  rcDither.bottom + 1);

	SetROP2(hDC, nOpOld);
	SelectBrush(hDC, hBrushOld);
}

long GetPaletteSize(HPALETTE hPal)
{
	long nSize = sizeof(WORD);
	if (hPal)
	{
		WORD nColors;
		GetObject(hPal, sizeof(WORD), &nColors);
		nSize += sizeof(PALETTEENTRY) * nColors;
	}
	return nSize;
}

//
// DrawVerticalText
//

void DrawVerticalText(HDC hDC, LPCTSTR szCaption, int nLen, const CRect& rcOut, BSTR bstrCaption)
{
	LPTSTR szBuffer = NULL;
	BSTR bstrBuffer = NULL;
	RECT rc = rcOut;
	int nDestLen = nLen;
	int nAmpPos = 0;
	int nLastAmpPos = -1;

	if (NULL == bstrCaption)
	{
		szBuffer = new TCHAR[nLen+1];
		if (NULL == szBuffer)
			return;
		lstrcpy(szBuffer, szCaption);
		LPTSTR szPtr = szBuffer;
		// Remove & symbols
		while (*szPtr)
		{
			if ('&' == *szPtr && '&' != *(szPtr+1))
			{
				nLastAmpPos = nAmpPos;
				lstrcpy(szPtr, szPtr+1);
				nDestLen--;
			}
			else
			{
				szPtr++;
				nAmpPos++;
			}
		}
	}
	else
	{
		bstrBuffer = SysAllocString(bstrCaption);
		if (NULL == bstrBuffer)
			return;
		WCHAR* wPtr = bstrBuffer;
		// Remove & symbols
		while (*wPtr)
		{
			if ('&' == *wPtr && '&' != *(wPtr+1))
			{
				nLastAmpPos = nAmpPos;
				wcscpy(wPtr, wPtr+1);
				nDestLen--;
			}
			else
			{
				wPtr++;
				nAmpPos++;
			}
		}
	}

#ifdef _SCRIPTSTRING
	if (bstrCaption)
		ScriptTextOut(hDC, rc.right, rc.top, bstrBuffer, nDestLen);
	else
#endif
		TextOut(hDC, rc.right, rc.top, szBuffer, nDestLen);
	if (-1 != nLastAmpPos)
	{   
		// Draw the underscore
		TEXTMETRIC tm;
		GetTextMetrics(hDC, &tm);

		// first character ?
		SIZE sizeText1;
		if (0 == nLastAmpPos) 
			sizeText1.cx = 0;
		else
		{
#ifdef _SCRIPTSTRING
			if (bstrCaption)
				ScriptGetTextExtent(hDC, bstrBuffer, nLastAmpPos, &sizeText1); 
			else
#endif
				GetTextExtentPoint32(hDC, szBuffer, nLastAmpPos, &sizeText1);
		}

		SIZE sizeText2;
#ifdef _SCRIPTSTRING
		if (bstrCaption)
			ScriptGetTextExtent(hDC, bstrBuffer, nLastAmpPos+1, &sizeText2);
		else
#endif
			GetTextExtentPoint32(hDC, szBuffer, nLastAmpPos+1, &sizeText2);
		
		RECT rcLine;
		rcLine.top = rc.top + sizeText1.cx - 1;
		rcLine.bottom = rc.top + sizeText2.cx + 1;
		rcLine.left = rc.right - (tm.tmHeight - tm.tmInternalLeading) - 1;
		rcLine.right = rcLine.left + 1;
		
		FillSolidRect(hDC, rcLine, GetTextColor(hDC));
	}
	delete [] szBuffer;
}

//
// DrawExpandBitmap
//

void DrawExpandBitmap(HDC hDC, CRect rcBitmap, HBITMAP hBitmap, int nOffset, int nSrcWidth, int nSrcHeight)
{
	HDC hMemDC = GetMemDC();
	if (NULL == hMemDC)
		return;

	HBITMAP hBitmapOld = SelectBitmap(hMemDC, hBitmap);
	
	COLORREF crTextOld = SetTextColor(hDC, RGB(0,0,0));
	COLORREF crBackOld = SetBkColor(hDC, RGB(255,255,255));
	StretchBlt(hDC, rcBitmap.left, rcBitmap.top, rcBitmap.Width(), rcBitmap.Height(), hMemDC, nOffset, 0, nSrcWidth, nSrcHeight, SRCAND);

	SetTextColor(hDC, crTextOld);
	SetBkColor(hDC, crBackOld);
	SelectBitmap(hMemDC, hBitmapOld);
}		

//
// FindLastToolId
//

long FindLastToolId(IActiveBar2* pActiveBar)
{
	long nMaxToolId = -1;
	try
	{
		ITools* pTools;

		HRESULT hResult = pActiveBar->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
			return -2;

		ITool* pTool;
		long nToolId;
		short nCount;
		VARIANT vIndex;
		vIndex.vt = VT_I2;
		pTools->Count(&nCount);
		for (vIndex.iVal = 0; vIndex.iVal < nCount; vIndex.iVal++)
		{
			hResult = pTools->Item(&vIndex, (Tool**)&pTool);
			if (FAILED(hResult))
				continue;

			pTool->get_ID(&nToolId);
			if (nToolId > nMaxToolId && !(nToolId & CBar::eSpecialToolId))
				nMaxToolId = nToolId;
			pTool->Release();
		}
		pTools->Release();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return nMaxToolId;
}

//
// BitBltMasked
//

BOOL BitBltMasked(HDC hDC, int left, int top, int width, int height, HDC hMemDC, int x, int y, int nSrcWidth, int nSrcHeight, COLORREF rcFore)
{
	COLORREF crForeOld = SetTextColor(hDC, RGB(0,0,0));	
	COLORREF crBackOld = SetBkColor(hDC, RGB(255,255,255));

	BOOL bResult = StretchBlt(hDC, 
							  left,
							  top,
							  width,
							  height,
							  hMemDC,
							  x,
							  y,
							  nSrcWidth, 
							  nSrcHeight,
							  SRCAND);
	if (!bResult)
		goto Cleanup;

	SetBkColor(hDC, RGB(0,0,0));
	SetTextColor(hDC, rcFore);

	bResult = StretchBlt(hDC,
						 left,
						 top,
						 width,
						 height,
						 hMemDC,
						 x,
						 y,
						 nSrcWidth, 
						 nSrcHeight,
						 SRCPAINT);
Cleanup:
	SetTextColor(hDC, crForeOld);
	SetBkColor(hDC, crBackOld);
	return bResult;
}

const BYTE g_rgcbDataTypeSize[] = {
    0,                      // VT_EMPTY= 0,
    0,                      // VT_NULL= 1,
    sizeof(short),          // VT_I2= 2,
    sizeof(long),           // VT_I4 = 3,
    sizeof(float),          // VT_R4  = 4,
    sizeof(double),         // VT_R8= 5,
    sizeof(CURRENCY),       // VT_CY= 6,
    sizeof(DATE),           // VT_DATE = 7,
    sizeof(BSTR),           // VT_BSTR = 8,
    sizeof(IDispatch *),    // VT_DISPATCH    = 9,
    sizeof(SCODE),          // VT_ERROR    = 10,
    sizeof(VARIANT_BOOL),   // VT_BOOL    = 11,
    sizeof(VARIANT),        // VT_VARIANT= 12,
    sizeof(IUnknown *),     // VT_UNKNOWN= 13,
};

//
// GetAmbientProperty
//

BOOL GetAmbientProperty(LPDISPATCH pDispAmbient, DISPID dispid, VARTYPE vt, LPVOID pData)
{
	DISPPARAMS dispparams;
    VARIANT v, v2;
    HRESULT hr;
    v.vt = VT_EMPTY;
    v.lVal = 0;
    v2.vt = VT_EMPTY;
    v2.lVal = 0;
    
    // now go and get the property into a variant.
    memset(&dispparams, 0, sizeof(DISPPARAMS));
    hr = pDispAmbient->Invoke(dispid, IID_NULL, 0, DISPATCH_PROPERTYGET, &dispparams,
                                &v, NULL, NULL);
    if (FAILED(hr)) 
		return FALSE;
    // we've got the variant, so now go an coerce it to the type that the user
    // wants.  if the types are the same, then this will copy the stuff to
    // do appropriate ref counting ...
    hr = VariantChangeType(&v2, &v, 0, vt);
    if (FAILED(hr)) 
	{
        VariantClear(&v);
        return FALSE;
    }
    // copy the data to where the user wants it
    CopyMemory(pData, &(v2.lVal), g_rgcbDataTypeSize[vt]);
    VariantClear(&v);
    return TRUE;
}

LPDISPATCH GetExtenderDisp(LPDISPATCH pInner)
{
	IOleObject* pOleObject=NULL;
	IOleControlSite* pControlSite=NULL;
	IOleClientSite* pCL=NULL;
	LPDISPATCH pExtDisp=NULL;
	HRESULT hr;

	hr=pInner->QueryInterface(IID_IOleObject,(LPVOID *)&pOleObject);
	if (FAILED(hr) || pOleObject==NULL)
		goto cleanup;
	hr= pOleObject->GetClientSite(&pCL);
	if (FAILED(hr) || pCL==NULL)
		goto cleanup;
	hr=pCL->QueryInterface(IID_IOleControlSite,(LPVOID *)&pControlSite);
	if (FAILED(hr) || pControlSite==NULL)
		goto cleanup;
	hr=pControlSite->GetExtendedControl(&pExtDisp);
	if (FAILED(hr))
		goto cleanup;

cleanup:
	if (pOleObject)
		pOleObject->Release();
	if (pControlSite)
		pControlSite->Release();
	if (pCL)
		pCL->Release();
	return pExtDisp;
}

//
//
//

CHAR WINAPI TranslateKeyCode(DWORD ch)
{
	TCHAR szLayout[KL_NAMELENGTH] = TEXT("");
	HKL hkl = NULL;
	HKL hklUSEng = NULL;
	SHORT code;
	UINT vk;
	UINT sc;
	BYTE keys[256] = {0};
	WORD chTranslated[3] = {0};
	const TCHAR szLayoutUSEnglish[] = TEXT("00000409");

	//
	// Get the thread's current layout and its layout name
	//

	hkl = GetKeyboardLayout(0);
	if (NULL == hkl)
		goto Default;

	if (!GetKeyboardLayoutName(szLayout))
		goto Default;

	//
	// If the current layout is the US English keyboard layout, we 
	// don't need to do anything.
	//

	if (0 != lstrcmpi(szLayout, szLayoutUSEnglish))
	{
		//
		// Load the US English keyboard layout. We will use this to 
		// convert the key code.
		//

		hklUSEng = LoadKeyboardLayout(szLayoutUSEnglish, KLF_NOTELLSHELL | KLF_SUBSTITUTE_OK);
		if (NULL == hklUSEng)
			goto Default;

		//
		// Translate the key code to the correcponding virtual-key code 
		// and shift state based on the current keyboard layout
		//

		code = VkKeyScanEx((CHAR)ch, hkl);
		if ((BYTE)-1 == LOBYTE(code) && (BYTE)-1 == HIBYTE(code))
			goto Default;

		//
		// Using the virtual-key code, get the scan code based on the 
		// current keyboard layout
		//

		vk = (UINT)LOBYTE(code);
		sc = MapVirtualKeyEx(vk, 0, hkl);

		//
		// Get the current keyboard state. We need this in order to 
		// call ToAsciiEx
		//

		GetKeyboardState(keys);

		//
		// Get the key code for the virtual-key/scan code in the US
		// English keyboard layout
		//

		if (1 != ToAsciiEx(vk, sc, keys, chTranslated, 0, hklUSEng))
			goto Default;

		//
		// Return the translated key code
		//

		UnloadKeyboardLayout(hklUSEng);
		return (CHAR)chTranslated[0];
	}

Default:
	//
	// If an error occurs, just return the key code that was passed 
	// to us
	//
	if (hklUSEng)
		UnloadKeyboardLayout(hklUSEng);
	return (CHAR)ch;
}

COLORREF GetShadeColor(HDC hDC, COLORREF crIn, int nValue)
{
	BYTE bRed = GetRValue(crIn);
	BYTE bGreen = GetGValue(crIn);
	BYTE bBlue = GetBValue(crIn);
 
	bRed = (bRed - nValue);
	if (bRed < 0)
		bRed = 0;
	if (bRed > 255)
		bRed = 255;

	bGreen = (bGreen - nValue) + 2;
	if (bGreen < 0)
		bGreen = 0;
	if (bGreen > 255)
		bGreen = 255;

	bBlue = (bBlue - nValue);
	if (bBlue < 0) 
		bBlue = 0;
	if (bBlue > 255) 
		bBlue = 255;
 
  return GetNearestColor(hDC, RGB(bRed, bGreen, bBlue));
}
 
COLORREF NewColor(HDC hDC, COLORREF crIn, int nValue)
{
	if (nValue > 100)
		nValue = 100;

	int bRed = GetRValue(crIn);
	int bGreen = GetGValue(crIn);
	int bBlue = GetBValue(crIn);
 
	bRed = bRed + (((255 - bRed) * nValue) / 100);
	bGreen = bGreen + (((255 - bGreen) * nValue) / 100);
	bBlue = bBlue + (((255 - bBlue) * nValue) / 100);

	COLORREF rcReturn = GetNearestColor(hDC, RGB(bRed, bGreen, bBlue));
	return rcReturn;
}
 
COLORREF GrayColor(HDC hDC, COLORREF crIn, int nValue)
{
	if (nValue > 100)
		nValue = 100;
	BYTE bRed = GetRValue(crIn);
	BYTE bGreen = GetGValue(crIn);
	BYTE bBlue = GetBValue(crIn);
 
	int nAvg = (bRed + bGreen + bBlue) / 3;
	nAvg = nAvg + nValue;
 
	if (nAvg > 240)
		nAvg = 240;
 
	return GetNearestColor (hDC, RGB(nAvg, nAvg, nAvg));
}
