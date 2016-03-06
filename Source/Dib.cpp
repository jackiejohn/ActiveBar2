//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <assert.h>
#include "Debug.h"
#include "Support.h"
#include "EventLog.h"
#include "Dib.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// CDib
//

CDib::CDib()
: m_hBitmap(NULL),
  m_hPalette(NULL),
  m_pBMIH(NULL),
  m_pImage(NULL)
{
	m_nBmihAlloc = m_nImageAlloc = eNoAlloc;
	Empty();
}

CDib::CDib(const SIZE& size, const int& nBitCount)
: m_hBitmap(NULL),
  m_hPalette(NULL),
  m_pImage(NULL)  // no data yet
{
	m_nBmihAlloc = m_nImageAlloc = eNoAlloc;
	Empty();
	ComputePaletteSize(nBitCount);
	m_pBMIH = (LPBITMAPINFOHEADER) new BYTE[sizeof(BITMAPINFOHEADER) + 
											 sizeof(RGBQUAD) * m_nColorTableEntries];
	m_nBmihAlloc = eCreateAlloc;
	m_pBMIH->biSize = sizeof(BITMAPINFOHEADER);
	m_pBMIH->biWidth = size.cx;
	m_pBMIH->biHeight = size.cy;
	m_pBMIH->biPlanes = 1;
	m_pBMIH->biBitCount = nBitCount;
	m_pBMIH->biCompression = BI_RGB;
	m_pBMIH->biSizeImage = 0;
	m_pBMIH->biXPelsPerMeter = 0;
	m_pBMIH->biYPelsPerMeter = 0;
	m_pBMIH->biClrUsed = m_nColorTableEntries;
	m_pBMIH->biClrImportant = m_nColorTableEntries;
	ComputeMetrics();
	memset(m_pvColorTable, 0, sizeof(RGBQUAD) * m_nColorTableEntries);
}

CDib::~CDib()
{
	Empty();
}

SIZE CDib::GetDimensions()
{	
	SIZE size;
	if (0 == m_pBMIH) 
	{
		size.cx = size.cy = 0;
		return size;
	}
	size.cx = m_pBMIH->biWidth;
	size.cy = m_pBMIH->biHeight;
	return size;
}

BOOL CDib::AttachMemory(LPVOID lpvMem, BOOL bMustDelete, HGLOBAL hGlobal)
{
	// assumes contiguous BITMAPINFOHEADER, color table, image
	// color table could be zero length
	Empty();
	m_hGlobal = hGlobal;
	if (bMustDelete == FALSE) 
		m_nBmihAlloc = eNoAlloc;
	else 
		m_nBmihAlloc = ((hGlobal == 0) ? eCreateAlloc : eHeapAlloc);
	
	try 
	{
		m_pBMIH = (LPBITMAPINFOHEADER) lpvMem;

		ComputeMetrics();
		
		ComputePaletteSize(m_pBMIH->biBitCount);
		
		m_pImage = (LPBYTE)m_pvColorTable + sizeof(RGBQUAD) * m_nColorTableEntries;
		MakePalette();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

HGLOBAL CDib::SaveToMemory()
{
	int nHeaderSize = GetSizeHeader();
    int nImageSize = GetSizeImage();
    HGLOBAL hHeader = ::GlobalAlloc(GMEM_SHARE,
									nHeaderSize + nImageSize);
    LPVOID pHeader = ::GlobalLock(hHeader);
    if (NULL == pHeader)
		return NULL;

    LPVOID pImage = (LPBYTE)pHeader + nHeaderSize;
    memcpy(pHeader, m_pBMIH, nHeaderSize); 
    memcpy(pImage, m_pImage, nImageSize);
      
	// Receiver is supposed to free the global memory 
    ::GlobalUnlock(hHeader);
	return hHeader;
}

UINT CDib::UsePalette(HDC hDC, BOOL bBackground /* = FALSE */)
{
	if (NULL == m_hPalette) 
		return NULL;

	::SelectPalette(hDC, m_hPalette, bBackground);
	return ::RealizePalette(hDC);
}

BOOL CDib::Draw(HDC hDC, const POINT& ptOrigin, const SIZE& size)
{
	if (NULL == m_pBMIH) 
		return FALSE;

	if (m_hPalette) 
		::SelectPalette(hDC, m_hPalette, TRUE);

	::SetStretchBltMode(hDC, COLORONCOLOR);
	::StretchDIBits(hDC, 
					ptOrigin.x, 
					ptOrigin.y, 
					size.cx, 
					size.cy,
					0, 
					0, 
					m_pBMIH->biWidth, 
					m_pBMIH->biHeight,
					m_pImage, 
					(LPBITMAPINFO) m_pBMIH, 
					DIB_RGB_COLORS, 
					SRCCOPY);
	return TRUE;
}

HBITMAP CDib::CreateSection(HDC hDC)
{
	if (NULL == m_pBMIH) 
		return NULL;

	if (NULL != m_pImage) 
		return NULL;

	m_hBitmap = ::CreateDIBSection(hDC, 
								   (LPBITMAPINFO) m_pBMIH,
								   DIB_RGB_COLORS,	
								   (LPVOID*) &m_pImage, 
								   0, 
								   0);
	assert(NULL != m_pImage);
	return m_hBitmap;
}

BOOL CDib::MakePalette()
{
	// makes a logical palette (m_hPalette) from the DIB's color table
	// this palette will be selected and realized prior to drawing the DIB
	if (0 == m_nColorTableEntries) 
		return FALSE;

	if (m_hPalette) 
		::DeletePalette(m_hPalette);

	TRACE1(1, "CDib::MakePalette -- m_nColorTableEntries = %d\n", m_nColorTableEntries);
	LPLOGPALETTE pLogPal = (LPLOGPALETTE) new BYTE[2 * 
												   sizeof(WORD) +
												   m_nColorTableEntries * 
												   sizeof(PALETTEENTRY)];
	pLogPal->palVersion = 0x300;
	pLogPal->palNumEntries = m_nColorTableEntries;
	LPRGBQUAD pDibQuad = (LPRGBQUAD) m_pvColorTable;
	for (int nIndex = 0; nIndex < m_nColorTableEntries; nIndex++) 
	{
		pLogPal->palPalEntry[nIndex].peRed = pDibQuad->rgbRed;
		pLogPal->palPalEntry[nIndex].peGreen = pDibQuad->rgbGreen;
		pLogPal->palPalEntry[nIndex].peBlue = pDibQuad->rgbBlue;
		pLogPal->palPalEntry[nIndex].peFlags = 0;
		pDibQuad++;
	}
	m_hPalette = ::CreatePalette(pLogPal);
	delete pLogPal;
	return TRUE;
}	

HBITMAP CDib::MakeGray()
{
	// convert bits to monochrome by changing palette
	int nMidVal = 0;
	LPRGBQUAD pDibQuad = GetColors();
	for (int cnt = 0; cnt < m_nColorTableEntries; ++cnt)
	{
		nMidVal = (pDibQuad->rgbBlue + pDibQuad->rgbBlue + pDibQuad->rgbRed) / 3;
		if (nMidVal > 0x20 && nMidVal <= 0x80)
			nMidVal = 0x80;
		pDibQuad->rgbBlue = nMidVal; 
		pDibQuad->rgbGreen = nMidVal; 
		pDibQuad->rgbRed = nMidVal; 
		pDibQuad++;
	}

	// get a DC 
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return NULL;

	// create bitmap from DIB info. and bits 
    HBITMAP hBitmap = ::CreateDIBitmap(hDC, 
								   	   m_pBMIH,
									   CBM_INIT, 
									   m_pImage, 
									   (LPBITMAPINFO) m_pBMIH, 
									   DIB_RGB_COLORS);
	ReleaseDC(NULL, hDC);
	return hBitmap;
}

BOOL CDib::SetSystemPalette(HDC hDC)
{
	// If the DIB doesn't have a color table, we can use the system's halftone palette
	if (0 != m_nColorTableEntries) 
		return FALSE;

	if (m_hPalette) 
	{
		BOOL bResult = ::DeletePalette(m_hPalette);
		assert(bResult);
	}
	m_hPalette = ::CreateHalftonePalette(hDC);
	return TRUE;
}

HBITMAP CDib::CreateBitmap(HDC hDC)
{
    if (0 == m_dwSizeImage) 
		return NULL;

    HBITMAP hBitmap = ::CreateDIBitmap(hDC, 
								   	   m_pBMIH,
									   CBM_INIT, 
									   m_pImage, 
									   (LPBITMAPINFO) m_pBMIH, 
									   DIB_RGB_COLORS);
    assert(NULL != hBitmap);
    return hBitmap;
}

HBITMAP CDib::CreateMonochromeBitmap(HDC hDC)
{
	HBITMAP hBitmap = NULL;
	if (NULL == m_pBMIH)
		return NULL;

	HDC hDCDest = ::CreateCompatibleDC(hDC);
	if (hDCDest)
	{
		hBitmap = ::CreateBitmap(m_pBMIH->biWidth, m_pBMIH->biHeight, 1, 1, NULL);
		assert(hBitmap);
		if (hBitmap)
		{
			HBITMAP hBitmapDestOld = SelectBitmap(hDCDest, hBitmap);

			SetBkColor(hDCDest, RGB(255, 255, 255));
			
			::SetStretchBltMode(hDCDest, COLORONCOLOR);
			
			::StretchDIBits(hDCDest, 
							0, 
							0, 
							m_pBMIH->biWidth, 
							m_pBMIH->biHeight,
							0, 
							0, 
							m_pBMIH->biWidth, 
							m_pBMIH->biHeight,
							m_pImage, 
							(LPBITMAPINFO)m_pBMIH, 
							DIB_RGB_COLORS, 
							SRCCOPY);

			SelectBitmap(hDCDest, hBitmapDestOld);
		}
		DeleteDC(hDCDest);
	}
	return hBitmap;
}

BOOL CDib::Compress(HDC hDC, BOOL bCompress)
{
	if (NULL == m_pBMIH)
		return FALSE;

	// 1. makes GDI bitmap from existing DIB
	// 2. makes a new DIB from GDI bitmap with compression
	// 3. cleans up the original DIB
	// 4. puts the new DIB in the object
	if ((4 != m_pBMIH->biBitCount) && (8 != m_pBMIH->biBitCount)) 
		return FALSE;
	
	// compression supported only for 4 bpp and 8 bpp DIBs
	// can't compress a DIB Section!
	if (m_hBitmap) 
		return FALSE; 
	
	TRACE1(1, "Compress: original palette size = %d\n", m_nColorTableEntries); 
	HPALETTE hOldPalette = ::SelectPalette(hDC, m_hPalette, FALSE);
	
	HBITMAP hBitmap;  // temporary
	if (NULL == (hBitmap = CreateBitmap(hDC))) 
		return FALSE;

	int nSize = sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * m_nColorTableEntries;
	LPBITMAPINFOHEADER lpBMIH = (LPBITMAPINFOHEADER) new BYTE[nSize];
	memcpy(lpBMIH, m_pBMIH, nSize);  // new header
	if (bCompress) 
	{
		switch (lpBMIH->biBitCount) 
		{
		case 4:
			lpBMIH->biCompression = BI_RLE4;
			break;
		case 8:
			lpBMIH->biCompression = BI_RLE8;
			break;
		default:
			assert(FALSE);
		}

		// calls GetDIBits with 0 data pointer to get size of compressed DIB
		if(!::GetDIBits(hDC, 
						hBitmap, 
						0, 
						(UINT) lpBMIH->biHeight,
						0, 
						(LPBITMAPINFO)lpBMIH, 
						DIB_RGB_COLORS)) 
		{
			// probably a problem with the color table
	 		::DeleteBitmap(hBitmap);
			delete [] lpBMIH;
			::SelectPalette(hDC, hOldPalette, FALSE);
			return FALSE; 
		}
		if (0 == lpBMIH->biSizeImage) 
		{
	 		::DeleteBitmap(hBitmap);
			delete [] lpBMIH;
			::SelectPalette(hDC, hOldPalette, FALSE);
			return FALSE; 
		}
		else 
			m_dwSizeImage = lpBMIH->biSizeImage;
	}
	else 
	{
		lpBMIH->biCompression = BI_RGB; // decompress
		// figure the image size from the bitmap width and height
		DWORD dwBytes = ((DWORD) lpBMIH->biWidth * lpBMIH->biBitCount) / 32;
		if (((DWORD) lpBMIH->biWidth * lpBMIH->biBitCount) % 32) 
			dwBytes++;

		dwBytes *= 4;
		m_dwSizeImage = dwBytes * lpBMIH->biHeight; // no compression
		lpBMIH->biSizeImage = m_dwSizeImage;
	} 
	// second GetDIBits call to make DIB
	LPBYTE lpImage = (LPBYTE) new BYTE[m_dwSizeImage];
	::GetDIBits(hDC, 
			   hBitmap, 
			   0, 
			   (UINT) lpBMIH->biHeight,
    		   lpImage, 
			   (LPBITMAPINFO) lpBMIH, 
			   DIB_RGB_COLORS);
    TRACE1(1, "dib successfully created - height = %d\n", lpBMIH->biHeight);
	::DeleteBitmap(hBitmap);
	Empty();

	m_nBmihAlloc = m_nImageAlloc = eCreateAlloc;
	m_pBMIH = lpBMIH;
	m_pImage = lpImage;
	
	ComputeMetrics();
	
	ComputePaletteSize(m_pBMIH->biBitCount);
	
	MakePalette();
	
	::SelectPalette(hDC, hOldPalette, FALSE);
	
	TRACE1(1, "Compress: new palette size = %d\n", m_nColorTableEntries); 
	return TRUE;
}

BOOL CDib::Read(HANDLE hFile)
{
	// 1. read file header to get size of info hdr + color table
	// 2. read info hdr (to get image size) and color table
	// 3. read image
	// can't use bfSize in file header
	Empty();
	int nCount, nSize;
	BITMAPFILEHEADER bmfh;
	try 
	{
		DWORD dwNumberOfBytesRead; 
		nCount = ReadFile(hFile, (LPVOID) &bmfh, sizeof(BITMAPFILEHEADER), &dwNumberOfBytesRead, NULL);
		if (nCount != sizeof(BITMAPFILEHEADER)) 
			throw 1;
		
		if (bmfh.bfType != 0x4d42) 
			throw 1;
		
		nSize = bmfh.bfOffBits - sizeof(BITMAPFILEHEADER);
		m_pBMIH = (LPBITMAPINFOHEADER) new BYTE[nSize];
		m_nBmihAlloc = m_nImageAlloc = eCreateAlloc;

		nCount = ReadFile(hFile, m_pBMIH, nSize, &dwNumberOfBytesRead, NULL);
		
		ComputeMetrics();
		
		ComputePaletteSize(m_pBMIH->biBitCount);
		
		MakePalette();
		
		m_pImage = (LPBYTE) new BYTE[m_dwSizeImage];
		
		nCount = ReadFile(hFile, m_pImage, m_dwSizeImage, &dwNumberOfBytesRead, NULL);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

BOOL CDib::ReadSection(HANDLE hFile, HDC hDC)
{
	// new function reads BMP from disk and creates a DIB section
	//    allows modification of bitmaps from disk
	// 1. read file header to get size of info hdr + color table
	// 2. read info hdr (to get image size) and color table
	// 3. create DIB section based on header parms
	// 4. read image into memory that CreateDibSection allocates
	Empty();
	int nCount, nSize;
	BITMAPFILEHEADER bmfh;
	try 
	{
		DWORD dwNumberOfBytesRead; 
		nCount = ReadFile(hFile, (LPVOID)&bmfh, sizeof(BITMAPFILEHEADER), &dwNumberOfBytesRead, NULL);
		if (nCount != sizeof(BITMAPFILEHEADER)) 
			throw 1;

		if (bmfh.bfType != 0x4d42)
			throw 1;

		nSize = bmfh.bfOffBits - sizeof(BITMAPFILEHEADER);
		m_pBMIH = (LPBITMAPINFOHEADER) new BYTE[nSize];
		m_nBmihAlloc = eCreateAlloc;
		m_nImageAlloc = eNoAlloc;
		nCount = ReadFile(hFile, m_pBMIH, nSize, &dwNumberOfBytesRead, NULL);

		if (m_pBMIH->biCompression != BI_RGB) 
			throw 1;

		ComputeMetrics();
		ComputePaletteSize(m_pBMIH->biBitCount);
		MakePalette();
		UsePalette(hDC);

		m_hBitmap = ::CreateDIBSection(hDC, 
									   (LPBITMAPINFO) m_pBMIH,
									   DIB_RGB_COLORS,	
									   (LPVOID*)&m_pImage, 
									   0, 
									   0);
		assert(NULL != m_pImage);
		nCount = ReadFile(hFile, m_pImage, m_dwSizeImage, &dwNumberOfBytesRead, NULL);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

int CDib::GetBitmapSize()
{
	int nSize = sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * m_nColorTableEntries;
	nSize += sizeof(BITMAPFILEHEADER);
	nSize += m_dwSizeImage;
	return nSize;
}

BOOL CDib::Write(HANDLE hFile)
{
	BITMAPFILEHEADER bmfh;
	bmfh.bfType = 0x4d42; // 'BM'
	int nSizeHdr = sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * m_nColorTableEntries;
	bmfh.bfSize = 0;
	bmfh.bfReserved1 = bmfh.bfReserved2 = 0;
	bmfh.bfOffBits = sizeof(BITMAPFILEHEADER) + 
					 sizeof(BITMAPINFOHEADER) +
					 sizeof(RGBQUAD) * 
					 m_nColorTableEntries;	
	try 
	{
		DWORD dwNumberOfBytesRead; 
		WriteFile(hFile, (LPVOID)&bmfh, sizeof(BITMAPFILEHEADER), &dwNumberOfBytesRead, NULL);
		WriteFile(hFile, (LPVOID)m_pBMIH,  nSizeHdr, &dwNumberOfBytesRead, NULL);
		WriteFile(hFile, (LPVOID)m_pImage, m_dwSizeImage, &dwNumberOfBytesRead, NULL);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

//
// helper functions
//

void CDib::ComputePaletteSize(const int& nBitCount)
{
	if ((m_pBMIH == 0) || (m_pBMIH->biClrUsed == 0)) 
	{
		switch(nBitCount) 
		{
			case 1:
				m_nColorTableEntries = 2;
				break;
			case 4:
				m_nColorTableEntries = 16;
				break;
			case 8:
				m_nColorTableEntries = 256;
				break;
			case 16:
			case 24:
			case 32:
				m_nColorTableEntries = 0;
				break;
			default:
				assert(FALSE);
		}
	}
	else 
		m_nColorTableEntries = m_pBMIH->biClrUsed;
	assert((m_nColorTableEntries >= 0) && (m_nColorTableEntries <= 256)); 
}

void CDib::ComputeMetrics()
{
	if (m_pBMIH->biSize != sizeof(BITMAPINFOHEADER)) 
	{
		TRACE(1, "Not a valid Windows bitmap -- probably an OS/2 bitmap\n");
		throw 1;
	}

	m_dwSizeImage = m_pBMIH->biSizeImage;
	if (m_dwSizeImage == 0) 
	{
		DWORD dwBytes = ((DWORD) m_pBMIH->biWidth * m_pBMIH->biBitCount) / 32;
		if (((DWORD) m_pBMIH->biWidth * m_pBMIH->biBitCount) % 32) 
			dwBytes++;
		dwBytes *= 4;
		m_dwSizeImage = dwBytes * m_pBMIH->biHeight; // no compression
	}
	m_pvColorTable = (LPBYTE) m_pBMIH + sizeof(BITMAPINFOHEADER);
}

void CDib::Empty()
{
	// this is supposed to clean up whatever is in the DIB
	if (m_nBmihAlloc == eCreateAlloc) 
	{
		delete [] m_pBMIH;
		m_pBMIH = NULL;
	}
	else if(m_nBmihAlloc == eHeapAlloc) 
	{
		::GlobalUnlock(m_hGlobal);
		::GlobalFree(m_hGlobal);
		m_hGlobal = NULL;
	}
	
	if (m_nImageAlloc == eCreateAlloc) 
	{
		delete [] m_pImage;
		m_pImage = NULL;
	}

	
	if (m_hPalette) 
	{
		::DeletePalette(m_hPalette);
		m_hPalette = NULL;
	}
	
	if (m_hBitmap) 
	{
		::DeleteBitmap(m_hBitmap);
		m_hBitmap = NULL;
	}
	
	m_nBmihAlloc = m_nImageAlloc = eNoAlloc;
	m_nColorTableEntries = 0;
	m_dwSizeImage = 0;
	m_pvFile = NULL;
	m_pvColorTable = NULL;
}

HRESULT CDib::Read(IStream* pStream)
{
	// 1. read file header to get size of info hdr + color table
	// 2. read info hdr (to get image size) and color table
	// 3. read image
	// can't use bfSize in file header
	Empty();
	int nSize;
	HRESULT hResult;
	BITMAPFILEHEADER bmfh;
	try 
	{
		DWORD dwNumberOfBytesRead; 
		hResult = pStream->Read((LPVOID) &bmfh, sizeof(BITMAPFILEHEADER), &dwNumberOfBytesRead);
		if (dwNumberOfBytesRead != sizeof(BITMAPFILEHEADER)) 
			return E_FAIL;
		
		if (bmfh.bfType != 0x4d42) 
			return E_FAIL;
		
		nSize = bmfh.bfOffBits - sizeof(BITMAPFILEHEADER);
		m_nBmihAlloc = m_nImageAlloc = eCreateAlloc;
		m_pBMIH = (LPBITMAPINFOHEADER) new BYTE[nSize];
		if (NULL == m_pBMIH)
			return FALSE;

		hResult = pStream->Read(m_pBMIH, nSize, &dwNumberOfBytesRead);
		if (FAILED(hResult))
		{
			delete [] m_pBMIH;
			return E_FAIL;
		}
		
		ComputeMetrics();
		
		ComputePaletteSize(m_pBMIH->biBitCount);
		
		MakePalette();
		
		m_pImage = (LPBYTE) new BYTE[m_dwSizeImage];
		if (NULL == m_pImage)
		{
			delete [] m_pBMIH;
			return E_OUTOFMEMORY;
		}
		
		hResult = pStream->Read(m_pImage, m_dwSizeImage, &dwNumberOfBytesRead);
		if (FAILED(hResult))
		{
			delete [] m_pBMIH;
			delete [] m_pImage;
			return E_FAIL;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

HRESULT CDib::Write(IStream* pStream)
{
	BITMAPFILEHEADER bmfh;
	bmfh.bfType = 0x4d42; // 'BM'
	int nSizeHdr = sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * m_nColorTableEntries;
	bmfh.bfSize = 0;
	bmfh.bfReserved1 = bmfh.bfReserved2 = 0;
	bmfh.bfOffBits = sizeof(BITMAPFILEHEADER) + 
					 sizeof(BITMAPINFOHEADER) +
					 sizeof(RGBQUAD) * 
					 m_nColorTableEntries;	
	try 
	{
		DWORD dwNumberOfBytesRead; 
		HRESULT hResult = pStream->Write((LPVOID)&bmfh, sizeof(BITMAPFILEHEADER), &dwNumberOfBytesRead);
		if (FAILED(hResult))
			return hResult;
		hResult = pStream->Write((LPVOID)m_pBMIH,  nSizeHdr, &dwNumberOfBytesRead);
		if (FAILED(hResult))
			return hResult;
		hResult = pStream->Write((LPVOID)m_pImage, m_dwSizeImage, &dwNumberOfBytesRead);
		if (FAILED(hResult))
			return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

BOOL CDib::FromBitmap (HBITMAP hBitmap, HDC hDC, int nBitCount)
{
    // check if bitmap handle is valid
    if (NULL == hBitmap || NULL == hDC)
		return FALSE;

	Empty();
	SetSystemPalette(hDC);

    // fill in BITMAP structure, return NULL if it didn't work
    BITMAP bm;                   // bitmap structure
    if (!GetObject(hBitmap, sizeof(bm), (LPBYTE)&bm))
		return FALSE;

	if (0 == nBitCount)
		nBitCount = bm.bmPlanes * bm.bmBitsPixel;
	ComputePaletteSize(nBitCount);

	int nSizeBMHeader = sizeof(BITMAPINFOHEADER);
    // Calculate size of memory block required to store BITMAPINFO
    m_pBMIH = (BITMAPINFOHEADER*)new BYTE[nSizeBMHeader + sizeof(RGBQUAD) * m_nColorTableEntries];
    if (NULL == m_pBMIH)
       return FALSE;

	m_pBMIH->biSize = nSizeBMHeader;
	m_pBMIH->biWidth = bm.bmWidth;
	m_pBMIH->biHeight = bm.bmHeight;
	m_pBMIH->biPlanes = 1;
	m_pBMIH->biBitCount = nBitCount;
	m_pBMIH->biCompression = BI_RGB;
	m_pBMIH->biSizeImage = 0;
	m_pBMIH->biXPelsPerMeter = 0;
	m_pBMIH->biYPelsPerMeter = 0;
	m_pBMIH->biClrUsed = 0;
	m_pBMIH->biClrImportant = 0;

	// select and realize our palette
    HPALETTE hPalOld = SelectPalette(hDC, m_hPalette, FALSE);
    RealizePalette(hDC);

    //
    // Call GetDIBits with a NULL lpBits param, so it will calculate the
    // biSizeImage field for us
    //

    int nResult = GetDIBits(hDC, 
							hBitmap, 
							0, 
							(WORD)bm.bmHeight, 
							NULL, 
  							(LPBITMAPINFO)m_pBMIH,
							DIB_RGB_COLORS);
	if (0 == nResult)
	{
		SelectPalette(hDC, hPalOld, TRUE);
		RealizePalette(hDC);
		return FALSE;
	}

	SelectPalette(hDC, hPalOld, TRUE);
	RealizePalette(hDC);

	m_nBmihAlloc = m_nImageAlloc = eCreateAlloc;

	ComputeMetrics();
	
	ComputePaletteSize(m_pBMIH->biBitCount);
	
	MakePalette();

	m_pImage = new BYTE[m_dwSizeImage];
	if (NULL == m_pImage)
	{
		Empty();
		return FALSE;
	}

	//
	// Call GetDIBits with a NON-NULL lpBits param, and actualy get the
	// bits this time
	//

	if (0 == GetDIBits(hDC, 
					   hBitmap, 
					   0, 
					   (WORD)bm.bmHeight, 
					   (LPBYTE)m_pImage, 
					   (LPBITMAPINFO)m_pBMIH,
					   DIB_RGB_COLORS))
	{
		Empty();
		return NULL;
	}
	return TRUE;
}

//
// CColorQuantizer
//

CColorQuantizer::CColorQuantizer(UINT nMaxColors, UINT nColorBits)
	:	m_pTree(NULL)
{
	m_nMaxColors = nMaxColors;
	m_nColorBits = nColorBits;
	m_nLeafCount = 0;
    for (int i=0; i <= (int)m_nColorBits; i++)
        m_pReducibleNodes[i] = NULL;
}

CColorQuantizer::~CColorQuantizer()
{
	if (m_pTree)
		DeleteTree (m_pTree);
}

HPALETTE CColorQuantizer::CreatePalette()
{
	HPALETTE hPalette = NULL;
	if (m_pTree)
	{
		// Create a logical palette from the colors in the octree
		DWORD dwSize = sizeof (LOGPALETTE) + ((m_nLeafCount - 1) * sizeof (PALETTEENTRY));
		LOGPALETTE* plp  = (LOGPALETTE*) HeapAlloc (GetProcessHeap (), 0, dwSize);
		if (plp == NULL) 
			return NULL;
		plp->palVersion = 0x300;
		plp->palNumEntries = (WORD) m_nLeafCount;
		UINT nIndex = 0;
		GetPaletteColors (m_pTree, plp->palPalEntry, nIndex);
    	hPalette = ::CreatePalette (plp);
	    HeapFree (GetProcessHeap(), 0, plp);
	}
	else
	{
		HDC hDC = GetDC(NULL);
		if (hDC)
		{
			hPalette = CreateHalftonePalette (hDC);
			ReleaseDC(NULL, hDC);
		}
	}
    return hPalette;
}

BOOL CColorQuantizer::ProcessImage (CDib& theImage)
{
    int i, j, nPad = 0;
    BYTE* pbBits;
    WORD* pwBits;
    DWORD* pdwBits;
    DWORD rmask, gmask, bmask;
    int rright, gright, bright;
    int rleft, gleft, bleft;
    BYTE r, g, b;
    WORD wColor;
    DWORD dwColor;

    // Initialize octree variables
	if (m_nColorBits > 8) // Just in case
        return NULL;

    switch (theImage.BitmapInfoHeader()->biBitCount) 
	{
    case 16: // One case for 16-bit DIBs
        rmask = 0x7C00;
        gmask = 0x03E0;
        bmask = 0x001F;

        rright = GetRightShiftCount (rmask);
        gright = GetRightShiftCount (gmask);
        bright = GetRightShiftCount (bmask);

        rleft = GetLeftShiftCount (rmask);
        gleft = GetLeftShiftCount (gmask);
        bleft = GetLeftShiftCount (bmask);

        pwBits = (WORD*)theImage.ImageBits();;
        for (i=0; i < theImage.BitmapInfoHeader()->biHeight; i++) 
		{
            for (j=0; j < theImage.BitmapInfoHeader()->biWidth; j++) 
			{
                wColor = *pwBits++;
                b = (BYTE) (((wColor & (WORD) bmask) >> bright) << bleft);
                g = (BYTE) (((wColor & (WORD) gmask) >> gright) << gleft);
                r = (BYTE) (((wColor & (WORD) rmask) >> rright) << rleft);
                
				AddColor (r, g, b);
            }
            pwBits = (WORD*) (((BYTE*) pwBits) + nPad);
        }
        break;

    case 24: // Another for 24-bit DIBs
        pbBits = (BYTE*)theImage.ImageBits();
        for (i=0; i < theImage.BitmapInfoHeader()->biHeight; i++) 
		{
            for (j=0; j < theImage.BitmapInfoHeader()->biWidth; j++) 
			{
                b = *pbBits++;
                g = *pbBits++;
                r = *pbBits++;
                
				AddColor (r, g, b);
            }
            pbBits += nPad;
        }
        break;

    case 32: // And another for 32-bit DIBs
        rmask = 0x00FF0000;
        gmask = 0x0000FF00;
        bmask = 0x000000FF;

        rright = GetRightShiftCount (rmask);
        gright = GetRightShiftCount (gmask);
        bright = GetRightShiftCount (bmask);

        pdwBits = (DWORD*)theImage.ImageBits();
        for (i=0; i < theImage.BitmapInfoHeader()->biHeight; i++) 
		{
            for (j=0; j < theImage.BitmapInfoHeader()->biWidth; j++) 
			{
                dwColor = *pdwBits++;
                b = (BYTE) ((dwColor & bmask) >> bright);
                g = (BYTE) ((dwColor & gmask) >> gright);
                r = (BYTE) ((dwColor & rmask) >> rright);

				AddColor (r, g, b);
            }
            pdwBits = (DWORD*) (((BYTE*) pdwBits) + nPad);
        }
        break;

    default: // DIB must be 16, 24, or 32-bit!
        return FALSE;
    }

    if (m_nLeafCount > m_nMaxColors && m_pTree) 
	{ 
		// Sanity check
		DeleteTree (m_pTree);
        return FALSE;
    }
	return TRUE;
}

void CColorQuantizer::AddColor (BYTE bRed, BYTE bGreen, BYTE bBlue)
{
	AddColor (m_pTree, bRed, bGreen, bBlue, m_nColorBits, 0, m_nLeafCount, m_pReducibleNodes);
    
	while (m_nLeafCount > m_nMaxColors)
        ReduceTree (m_nColorBits, m_nLeafCount, m_pReducibleNodes);
}

void CColorQuantizer::AddColor (TreeNode*& pNode, BYTE r, BYTE g, BYTE b, UINT nColorBits, UINT nLevel, UINT& nLeafCount, TreeNode** pReducibleNodes)
{
    int nIndex, shift;
    static BYTE mask[8] = { 0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01 };

    // If the node doesn't exist, create it
    if (pNode == NULL)
        pNode = CreateNode (nLevel, nColorBits, nLeafCount, pReducibleNodes);

    // Update color information if it's a leaf node
    if (pNode->bIsLeaf) 
	{
        pNode->nPixelCount++;
        pNode->nRedSum += r;
        pNode->nGreenSum += g;
        pNode->nBlueSum += b;
    }
    else 
	{
	    // Recurse a level deeper if the node is not a leaf
        shift = 7 - nLevel;
        nIndex = (((r & mask[nLevel]) >> shift) << 2) | (((g & mask[nLevel]) >> shift) << 1) | ((b & mask[nLevel]) >> shift);
        AddColor (pNode->pChild[nIndex], r, g, b, nColorBits, nLevel + 1, nLeafCount, pReducibleNodes);
    }
}

TreeNode* CColorQuantizer::CreateNode (UINT nLevel, UINT nColorBits, UINT& nLeafCount, TreeNode** pReducibleNodes)
{
    TreeNode* pNode = (TreeNode*)HeapAlloc (GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof (TreeNode));

    if (NULL == pNode)
        return NULL;

    pNode->bIsLeaf = (nLevel == nColorBits) ? TRUE : FALSE;
    if (pNode->bIsLeaf)
        nLeafCount++;
    else 
	{ 
		// Add the node to the reducible list for this level
        pNode->pNext = pReducibleNodes[nLevel];
        pReducibleNodes[nLevel] = pNode;
    }
    return pNode;
}

void CColorQuantizer::ReduceTree (UINT nColorBits, UINT& nLeafCount, TreeNode** pReducibleNodes)
{
    int i;
    TreeNode* pNode;
    UINT nRedSum, nGreenSum, nBlueSum, nChildren;

    // Find the deepest level containing at least one reducible node
    for (i=nColorBits - 1; (i>0) && (pReducibleNodes[i] == NULL); i--);

    // Reduce the node most recently added to the list at level i
    pNode = pReducibleNodes[i];
    pReducibleNodes[i] = pNode->pNext;

    nRedSum = nGreenSum = nBlueSum = nChildren = 0;
    for (i=0; i<8; i++) 
	{
        if (pNode->pChild[i] != NULL) 
		{
            nRedSum += pNode->pChild[i]->nRedSum;
            nGreenSum += pNode->pChild[i]->nGreenSum;
            nBlueSum += pNode->pChild[i]->nBlueSum;
            pNode->nPixelCount += pNode->pChild[i]->nPixelCount;
            HeapFree (GetProcessHeap (), 0, pNode->pChild[i]);
            pNode->pChild[i] = NULL;
            nChildren++;
        }
    }

    pNode->bIsLeaf = TRUE;
    pNode->nRedSum = nRedSum;
    pNode->nGreenSum = nGreenSum;
    pNode->nBlueSum = nBlueSum;
    nLeafCount -= (nChildren - 1);
}

void CColorQuantizer::DeleteTree (TreeNode*& pNode)
{
    int i;

    for (i=0; i<8; i++) 
	{
        if (pNode->pChild[i])
            DeleteTree (pNode->pChild[i]);
    }
    HeapFree (GetProcessHeap (), 0, pNode);
    pNode = NULL;
}

void CColorQuantizer::GetPaletteColors (TreeNode* pTree, PALETTEENTRY* pPalEntries, UINT& nIndex)
{
    int i;

    if (pTree->bIsLeaf) 
	{
        pPalEntries[nIndex].peRed =   (BYTE) ((pTree->nRedSum) / (pTree->nPixelCount));
        pPalEntries[nIndex].peGreen = (BYTE) ((pTree->nGreenSum) / (pTree->nPixelCount));
        pPalEntries[nIndex].peBlue =  (BYTE) ((pTree->nBlueSum) / (pTree->nPixelCount));
		pPalEntries[nIndex].peFlags = 0;
        nIndex++;
    }
    else 
	{
        for (i=0; i<8; i++) 
		{
            if (pTree->pChild[i])
                GetPaletteColors (pTree->pChild[i], pPalEntries, nIndex);
        }
    }
}

int CColorQuantizer::GetRightShiftCount (DWORD dwVal)
{
    int i;

    for (i=0; i < sizeof (DWORD) * 8; i++) 
	{
        if (dwVal & 1)
            return i;
        dwVal >>= 1;
    }
    return -1;
}

int CColorQuantizer::GetLeftShiftCount (DWORD dwVal)
{
    int nCount, i;

    nCount = 0;
    for (i=0; i < sizeof (DWORD) * 8; i++) 
	{
        if (dwVal & 1)
            nCount++;
        dwVal >>= 1;
    }
    return (8 - nCount);
}

