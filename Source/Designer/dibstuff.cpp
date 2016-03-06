//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "dibstuff.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define IS_WIN30_DIB(lpbi)  ((*(LPDWORD)(lpbi)) == sizeof(BITMAPINFOHEADER))
#define DIB_HEADER_MARKER   ((WORD) ('M' << 8) | 'B')
#define WIDTHBYTES(bits)    (((bits) + 31) / 32 * 4)
/* DIB constants */
#define PALVERSION   0x300


WORD WINAPI DIBNumColors(LPBYTE lpbi)
{
	WORD wBitCount;  // DIB bit count

	/*  If this is a Windows-style DIB, the number of colors in the
	 *  color table can be less than the number of bits per pixel
	 *  allows for (i.e. lpbi->biClrUsed can be set to some value).
	 *  If this is the case, return the appropriate value.
	 */

	if (IS_WIN30_DIB(lpbi))
	{
		DWORD dwClrUsed;

		dwClrUsed = ((LPBITMAPINFOHEADER)lpbi)->biClrUsed;
		if (dwClrUsed != 0)
			return (WORD)dwClrUsed;
	}

	/*  Calculate the number of colors in the color table based on
	 *  the number of bits per pixel for the DIB.
	 */
	if (IS_WIN30_DIB(lpbi))
		wBitCount = ((LPBITMAPINFOHEADER)lpbi)->biBitCount;
	else
		wBitCount = ((LPBITMAPCOREHEADER)lpbi)->bcBitCount;

	/* return number of colors based on bits per pixel */
	switch (wBitCount)
	{
		case 1:
			return 2;

		case 4:
			return 16;

		case 8:
			return 256;

		default:
			return 0;
	}
}

WORD WINAPI PaletteSize(LPBYTE lpbi)
{
   if (IS_WIN30_DIB (lpbi))
	  return (WORD)(::DIBNumColors(lpbi) * sizeof(RGBQUAD));
   else
	  return (WORD)(::DIBNumColors(lpbi) * sizeof(RGBTRIPLE));
}

BOOL WINAPI StSaveDIB(IStream *pStream, HDIB hDib)
{
	if (hDib == NULL)
		return FALSE;

	BITMAPFILEHEADER bmfHdr; // Header for Bitmap file
	LPBITMAPINFOHEADER lpBI;   // Pointer to DIB info structure
	DWORD dwDIBSize;

	// Get BITMAPINFO structure
	lpBI = (LPBITMAPINFOHEADER) ::GlobalLock((HGLOBAL) hDib);
	if (lpBI == NULL)
		return FALSE;

	if (!IS_WIN30_DIB(lpBI))
	{
		::GlobalUnlock((HGLOBAL) hDib);
		return FALSE;       // It's an other-style DIB (save not supported)
	}

	// Fill file header
	bmfHdr.bfType = DIB_HEADER_MARKER;  // file type = "BM" packed in word
	// Can't use GlobalSize for size calc so use  header+colortable.....
	dwDIBSize = *(LPDWORD)lpBI + PaletteSize((LPBYTE)lpBI);
	// Now add size of image
	if ((lpBI->biCompression == BI_RLE8) || (lpBI->biCompression == BI_RLE4))
	{
		// It's an RLE bitmap, we can't calculate size, so use  biSizeImage field
		dwDIBSize += lpBI->biSizeImage;
	}
	else
	{
		DWORD dwBmBitsSize;  // Size of Bitmap Bits only
		// It's not RLE, size =width (DWORD aligned) * height
		dwBmBitsSize = WIDTHBYTES((lpBI->biWidth)*((DWORD)lpBI->biBitCount)) * lpBI->biHeight;
		dwDIBSize += dwBmBitsSize;
		lpBI->biSizeImage = dwBmBitsSize;
	}
	//Setup header with size info
	bmfHdr.bfSize = dwDIBSize + sizeof(BITMAPFILEHEADER);
	bmfHdr.bfReserved1 = 0;
	bmfHdr.bfReserved2 = 0;
	// Now calc butmap bits offset = dib header +colortable size
	bmfHdr.bfOffBits = (DWORD)sizeof(BITMAPFILEHEADER) + lpBI->biSize + PaletteSize((LPBYTE)lpBI);
	//try
	//{
		ULONG totalSize;
		totalSize=sizeof(BITMAPFILEHEADER)+dwDIBSize;
		pStream->Write(&totalSize,sizeof(ULONG),NULL);
		pStream->Write(&bmfHdr,sizeof(BITMAPFILEHEADER),NULL);
		pStream->Write(lpBI,dwDIBSize,NULL); // bitmaps above 4MB will fail
	//}
	//catch(...)
	//{
	//	::GlobalUnlock((HGLOBAL) hDib);
	//	return FALSE;
	//}
	
	::GlobalUnlock((HGLOBAL) hDib);
	return TRUE;
}


HRESULT StReadDIB(IStream *pStream,DWORD dataSize,HDIB *pdib)
{
	BITMAPFILEHEADER bmfHeader;
	DWORD dwBitsSize;
	HDIB hDIB;
	LPSTR pDIB;
	ULONG readAmount;

	*pdib=NULL;
	dwBitsSize = dataSize;
	pStream->Read(&bmfHeader,sizeof(BITMAPFILEHEADER),&readAmount);
	if (readAmount!=sizeof(BITMAPFILEHEADER))
		return E_FAIL;
	
	if (bmfHeader.bfType != DIB_HEADER_MARKER)
		return E_FAIL; // invalid format

	// alloc mem for DIB
	hDIB = (HDIB) ::GlobalAlloc(GMEM_MOVEABLE, dwBitsSize);
	if (hDIB == 0)
		return E_OUTOFMEMORY;
	pDIB = (LPSTR) ::GlobalLock((HGLOBAL) hDIB);
	// read bits
	if (FAILED(pStream->Read(pDIB,dwBitsSize - sizeof(BITMAPFILEHEADER),&readAmount)) || readAmount!=(dwBitsSize - sizeof(BITMAPFILEHEADER)))
	{
		::GlobalUnlock((HGLOBAL) hDIB);
		::GlobalFree((HGLOBAL) hDIB);
		return E_FAIL;
	}
	::GlobalUnlock((HGLOBAL) hDIB);
	*pdib=hDIB;
	return NOERROR;
}

int FAR PalEntriesOnDevice(HDC hDC)
{
   int nColors;  // number of colors

   /*  Find out the number of palette entries on this
    *  device.
    */

   nColors = GetDeviceCaps(hDC, SIZEPALETTE);

   /*  For non-palette devices, we'll use the # of system
    *  colors for our palette size.
    */
   if (!nColors)
      nColors = GetDeviceCaps(hDC, NUMCOLORS);
   return nColors;
}


HPALETTE FAR GetSystemPalette(void)
{
   HDC hDC;                // handle to a DC
   static HPALETTE hPal = NULL;   // handle to a palette
   HANDLE hLogPal;         // handle to a logical palette
   LPLOGPALETTE lpLogPal;  // pointer to a logical palette
   int nColors;            // number of colors

   /* Find out how many palette entries we want. */
   hDC = GetDC(NULL);
   if (!hDC)
      return NULL;
   nColors = PalEntriesOnDevice(hDC);   // Number of palette entries
   if (nColors<=0)
   {
	   ReleaseDC(NULL,hDC);
	   return NULL;
   }
   /* Allocate room for the palette and lock it. */
   hLogPal = GlobalAlloc(GHND, sizeof(LOGPALETTE) + nColors * sizeof(PALETTEENTRY));
   /* if we didn't get a logical palette, return NULL */
   if (!hLogPal)
   {
	  ReleaseDC(NULL,hDC);
      return NULL;
	}
   /* get a pointer to the logical palette */
   lpLogPal = (LPLOGPALETTE)GlobalLock(hLogPal);
   /* set some important fields */
   lpLogPal->palVersion = PALVERSION;
   lpLogPal->palNumEntries = nColors;
   /* Copy the current system palette into our logical palette */
   GetSystemPaletteEntries(hDC, 0, nColors, 
                           (LPPALETTEENTRY)(lpLogPal->palPalEntry));
   
	/*  Go ahead and create the palette.  Once it's created,
    *  we no longer need the LOGPALETTE, so free it.
    */

   hPal = CreatePalette(lpLogPal);
   /* clean up */
   GlobalUnlock(hLogPal);
   GlobalFree(hLogPal);
   ReleaseDC(NULL, hDC);
   return hPal;
}

HDIB BitmapToDIB(HBITMAP hBitmap, HPALETTE hPal, int pixperpixel)
{
	if (!hBitmap)
		return NULL;

	BITMAP bm;
	if (!GetObject(hBitmap, sizeof(bm), (LPSTR)&bm))
		return NULL;

	if (hPal == NULL)
		hPal = (HPALETTE)GetStockObject(DEFAULT_PALETTE);

	WORD biBits = bm.bmPlanes * bm.bmBitsPixel;

	if (-1 == pixperpixel)
    {
		if (biBits <= 1)
			biBits = 1;
		else if (biBits <= 4)
			biBits = 4;
		else if (biBits <= 8)
			biBits = 8;
		else 
			biBits = 24;
	}
	else
		biBits = pixperpixel;

	BITMAPINFOHEADER bi;
	bi.biSize = sizeof(BITMAPINFOHEADER);
	bi.biWidth = bm.bmWidth;
	bi.biHeight = bm.bmHeight;
	bi.biPlanes = 1;
	bi.biBitCount = biBits;
	bi.biCompression = BI_RGB;
	bi.biSizeImage = 0;
	bi.biXPelsPerMeter = 0;
	bi.biYPelsPerMeter = 0;
	bi.biClrUsed = 0;
	bi.biClrImportant = 0;

	DWORD dwLen = bi.biSize + PaletteSize((LPBYTE)&bi);

	HDC hDC = GetDC(NULL);
	HPALETTE hPalOld = NULL;
	if (NULL != hPal)
	{
		SelectPalette(hDC, hPal, FALSE);
		RealizePalette(hDC);
	}

	//
	// Alloc memory block to store our bitmap
	//

	HANDLE hDIB = GlobalAlloc(GHND, dwLen);

	if (!hDIB)
	{
		SelectPalette(hDC, hPalOld, TRUE);
		RealizePalette(hDC);
		ReleaseDC(NULL, hDC);
		return NULL;
	}

	BITMAPINFOHEADER FAR *lpbi = (BITMAPINFOHEADER *)GlobalLock(hDIB);

	*lpbi = bi;

	GetDIBits(hDC, 
			 hBitmap, 
			 0, 
			 (WORD)bi.biHeight, 
			 NULL, 
			 (LPBITMAPINFO)lpbi,
			 DIB_RGB_COLORS);

	bi = *lpbi;
	GlobalUnlock(hDIB);

	if (bi.biSizeImage == 0)
	  bi.biSizeImage = WIDTHBYTES((DWORD)bm.bmWidth * biBits) * bm.bmHeight;

	dwLen = bi.biSize + PaletteSize((LPBYTE)&bi) + bi.biSizeImage;
	HANDLE h;
	if (h = GlobalReAlloc(hDIB, dwLen, 0))
		hDIB = h;
	else
	{
		GlobalFree(hDIB);
		hDIB = NULL;
		SelectPalette(hDC, hPalOld, TRUE);
		RealizePalette(hDC);
		ReleaseDC(NULL, hDC);
		return NULL;
	}

	lpbi = (BITMAPINFOHEADER *)GlobalLock(hDIB);

	if (GetDIBits(hDC, 
				 hBitmap, 
				 0, 
				 (WORD)bi.biHeight, 
				 (LPSTR)lpbi + (WORD)lpbi->biSize + PaletteSize((LPBYTE)lpbi), 
				 (LPBITMAPINFO)lpbi,
				 DIB_RGB_COLORS) == 0)
	{
		GlobalUnlock(hDIB);
		GlobalFree(hDIB);
		hDIB = NULL;
		SelectPalette(hDC, hPalOld, TRUE);
		RealizePalette(hDC);
		ReleaseDC(NULL, hDC);
		return NULL;
	}
	bi = *lpbi;

   GlobalUnlock(hDIB);
   SelectPalette(hDC, hPalOld, TRUE);
   RealizePalette(hDC);
   ReleaseDC(NULL, hDC);

   return hDIB;
}

LPBYTE FAR FindDIBBits(LPBYTE lpDIB)
{
   return (lpDIB + *(LPDWORD)lpDIB + PaletteSize(lpDIB));
}

HBITMAP DIBToBitmap(HDIB hDIB, HPALETTE hPal, BOOL bMonochrome)
{
	if (!hDIB)
		return NULL;

	HBITMAP hBitmap = NULL;
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		LPBITMAPINFOHEADER lpDIBHdr = (LPBITMAPINFOHEADER)GlobalLock(hDIB);
		if (lpDIBHdr)
		{
			// 
			// Get a pointer to the DIB bits
			//

			LPBYTE lpDIBBits = FindDIBBits((LPBYTE)lpDIBHdr);
			if (lpDIBBits)
			{
				HPALETTE hOldPal = NULL;
				if (hPal)
				{
					hOldPal = SelectPalette(hDC, hPal, FALSE);
					RealizePalette(hDC);
				}

				//
				// Create bitmap from DIB information and bits
				//

				if (bMonochrome)
				{
					HDC hdcDest = ::CreateCompatibleDC(hDC);
					if (hdcDest)
					{
						hBitmap = CreateBitmap(lpDIBHdr->biWidth, lpDIBHdr->biHeight, 1, 1, NULL);
						if (hBitmap)
						{
							HBITMAP hBitmapDestOld = SelectBitmap(hdcDest, hBitmap);

							SetBkColor(hdcDest, RGB(255, 255, 255));
							::SetStretchBltMode(hdcDest, COLORONCOLOR);
							::StretchDIBits(hdcDest, 
											0, 
											0, 
											lpDIBHdr->biWidth, 
											lpDIBHdr->biHeight,
											0, 
											0, 
											lpDIBHdr->biWidth, 
											lpDIBHdr->biHeight,
											lpDIBBits, 
											(LPBITMAPINFO)lpDIBHdr, 
											DIB_RGB_COLORS, 
											SRCCOPY);

							SelectBitmap(hdcDest, hBitmapDestOld);
						}
						DeleteDC(hdcDest);
					}
				}
				else
				{
					hBitmap = CreateDIBitmap(hDC, 
											 lpDIBHdr, 
											 CBM_INIT,
											 lpDIBBits, 
											 (LPBITMAPINFO)lpDIBHdr, 
											 DIB_RGB_COLORS);
				}

				if (NULL != hOldPal)
					SelectPalette(hDC, hOldPal, FALSE);
			}
			GlobalUnlock(hDIB);
		}
		ReleaseDC(NULL, hDC);
	}
	return hBitmap;
}

HBITMAP FAR DIBToGrayBitmap(HDIB hDIB)
{
	if (!hDIB)
		return NULL;

	LPBITMAPINFOHEADER lpDIBHdr = (LPBITMAPINFOHEADER)GlobalLock(hDIB);

	LPBYTE lpDIBBits = FindDIBBits((LPBYTE)lpDIBHdr);

	int colCount=PaletteSize((LPBYTE)lpDIBHdr)/sizeof(RGBQUAD);

	RGBQUAD *colPtr = (RGBQUAD*)((LPBYTE)lpDIBHdr + sizeof(BITMAPINFOHEADER));
	for (int cnt=0;cnt<colCount;++cnt)
	{
		int midval=(colPtr->rgbBlue+colPtr->rgbBlue+colPtr->rgbRed)/3;
		if (midval>0x20 && midval<=0x80)
			midval=0x80;
		colPtr->rgbBlue=midval; 
		colPtr->rgbGreen=midval; 
		colPtr->rgbRed=midval; 
		++colPtr;
	}

	HDC hDC = GetDC(NULL);
	if (!hDC)
	{
		GlobalUnlock(hDIB);
		return NULL;
	}

	// 
	// Create bitmap from DIB info. and bits
    //

	HBITMAP hBitmap = CreateDIBitmap(hDC, 
									 (LPBITMAPINFOHEADER)lpDIBHdr, 
									 CBM_INIT,
									 lpDIBBits, 
									 (LPBITMAPINFO)lpDIBHdr, 
									 DIB_RGB_COLORS);
   
	ReleaseDC(NULL, hDC);
	GlobalUnlock(hDIB);
	return hBitmap;
}
