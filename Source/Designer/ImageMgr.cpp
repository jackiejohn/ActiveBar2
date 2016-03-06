//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "AB10Format.h"
#include "DibStuff.h"
#include "Globals.h"
#include "Debug.h"
#include "ImageMgr.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;

extern HBITMAP CreateMaskBitmap(HDC hDC, HBITMAP hBitmap, int nWidth, int nHeight);

static void DebugPrintBitmap(HBITMAP hBitmap, int nOffset = 0)
{
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		HDC hDCSrc = CreateCompatibleDC(hDC);
		HBITMAP hBitmapOld = SelectBitmap(hDCSrc, hBitmap);
		BITMAP bmInfo;
		GetObject(hBitmap, sizeof(bmInfo), &bmInfo);
		BitBlt(hDC, 0, nOffset, bmInfo.bmWidth, bmInfo.bmHeight, hDCSrc, 0, 0, SRCCOPY);
		SelectBitmap(hDCSrc, hBitmapOld);
		ReleaseDC(NULL, hDC);
	}
}

#endif

CImageMgr::CImageMgr()
	: m_refCount(1)
	, m_hPal(NULL),
	  m_ppArrayImageSizeEntry(NULL)
{
	m_dwMasterImageId = 0;
	m_nDepth = 16;
	m_nVersion = 0xFFFF;
}

CImageMgr::~CImageMgr()
{
	CleanUp();
	delete [] m_ppArrayImageSizeEntry;
}

//
// CleanUp
//

void CImageMgr::CleanUp()
{
	long nKey;
	ImageSizeEntry* pImageSizeEntry;
	POSITION posMap = m_mapImageSizes.GetStartPosition();
	while (0 != posMap)
	{
		m_mapImageSizes.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageSizeEntry);
		delete pImageSizeEntry;
	}
	m_mapImageSizes.RemoveAll();
	m_mapImages.RemoveAll();
	if (m_hPal)
		DeletePalette(m_hPal);
}

//
// IImageMgr members
//

STDMETHODIMP CImageMgr::get_MaskColor(long nImageId, OLE_COLOR *retval)
{
	if (-1 == nImageId)
		return NOERROR;

	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	pImageEntry->GetMaskColor(retval);
	return NOERROR;
}
STDMETHODIMP CImageMgr::put_MaskColor(long nImageId, OLE_COLOR val)
{
	if (-1 == nImageId)
		return NOERROR;

	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	pImageEntry->SetMaskColor(val);
	return NOERROR;
}

STDMETHODIMP CImageMgr::get_BitDepth(short *retval)
{
	*retval = m_nDepth;
	return S_OK;
}
STDMETHODIMP CImageMgr::put_BitDepth(short val)
{
	switch (val)
	{
	case 4:
	case 8:
	case 16:
	case 24:
		m_nDepth = val;
		return S_OK;

	default:
		return E_FAIL;
	}
	return E_FAIL;
}
STDMETHODIMP CImageMgr::get_Palette(OLE_HANDLE *retval)
{
	*retval = (OLE_HANDLE)m_hPal;
	return S_OK;
}
STDMETHODIMP CImageMgr::put_Palette(OLE_HANDLE val)
{
	m_hPal = (HPALETTE)val;
	return S_OK;
}
STDMETHODIMP CImageMgr::CreateImage(long* nImageId)
{
	SIZE sizeImage;
	sizeImage.cx = sizeImage.cy = eDefaultImageSize; 

	long nKey = MAKELONG(sizeImage.cx, sizeImage.cy);
	ImageSizeEntry* pImageSizeEntry;
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
	{
		pImageSizeEntry = new ImageSizeEntry(sizeImage, *this);
		if (NULL == pImageSizeEntry)
			return E_FAIL;

		m_mapImageSizes.SetAt((LPVOID)nKey, (LPVOID&)pImageSizeEntry);
	}

	++m_dwMasterImageId;
	ImageEntry* pImageEntry = new ImageEntry(pImageSizeEntry, m_dwMasterImageId);
	if (NULL == pImageEntry)
		return E_FAIL;

	m_mapImages.SetAt((LPVOID)m_dwMasterImageId, (LPVOID&)pImageEntry);
	*nImageId = m_dwMasterImageId;
	return pImageSizeEntry->Add(pImageEntry);
}
STDMETHODIMP CImageMgr::CreateImage2(long*nImageId,  int nCx,  int nCy)
{
	SIZE sizeImage;
	sizeImage.cx = nCx;
	sizeImage.cy = nCy;
	long nKey = MAKELONG(sizeImage.cx, sizeImage.cy);

	ImageSizeEntry* pImageSizeEntry;
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
	{
		pImageSizeEntry = new ImageSizeEntry(sizeImage, *this);
		if (NULL == pImageSizeEntry)
			return E_FAIL;

		m_mapImageSizes.SetAt((LPVOID)nKey, (LPVOID&)pImageSizeEntry);
	}

	m_dwMasterImageId++;
	ImageEntry* pImageEntry = new ImageEntry(pImageSizeEntry, m_dwMasterImageId);
	if (NULL == pImageEntry)
		return E_FAIL;

	m_mapImages.SetAt((LPVOID)m_dwMasterImageId, (LPVOID&)pImageEntry);
	*nImageId = m_dwMasterImageId;
	return pImageSizeEntry->Add(pImageEntry);
}
STDMETHODIMP CImageMgr::AddRefImage( long nImageId)
{
	if (-1 == nImageId)
		return S_OK;

	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	pImageEntry->AddRef();
	return S_OK;
}
STDMETHODIMP CImageMgr::RefCntImage( long nImageId, long* pnRefCnt)
{
	*pnRefCnt = 0;
	if (-1 == nImageId)
		return E_FAIL;

	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	*pnRefCnt = pImageEntry->RefCnt();
	return S_OK;
}
STDMETHODIMP CImageMgr::ReleaseImage(long nImageId)
{
	if (-1 == nImageId)
		return S_OK;

	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	if (0 == pImageEntry->Release())
	{	
		if (!m_mapImages.RemoveKey((LPVOID)nImageId))
			return E_FAIL;
	}
	return S_OK;
}
STDMETHODIMP CImageMgr::Size(long nImageId, long* nCx,  long* nCy)
{
	*nCx = *nCy = 0;
	if (-1 == nImageId)
		return S_OK;

	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;
	
	SIZE size;
	HRESULT hResult = pImageEntry->GetSize(size);
	*nCx = size.cx;
	*nCy = size.cy;
	return hResult;
}
STDMETHODIMP CImageMgr::put_ImageBitmap( long nImageId,  OLE_HANDLE hBitmap)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	return pImageEntry->SetImageBitmap((HBITMAP)hBitmap);
}
STDMETHODIMP CImageMgr::get_ImageBitmap( long nImageId,  OLE_HANDLE* hBitmap)
{
	*hBitmap = NULL;
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	return pImageEntry->GetImageBitmap((HBITMAP&)*hBitmap);
}
STDMETHODIMP CImageMgr::put_MaskBitmap(long nImageId,  OLE_HANDLE hBitmap)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	return pImageEntry->SetMaskBitmap((HBITMAP)hBitmap);
}
STDMETHODIMP CImageMgr::get_MaskBitmap( long nImageId,  OLE_HANDLE*hBitmap)
{
	*hBitmap = NULL;
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	return pImageEntry->GetMaskBitmap((HBITMAP&)*hBitmap);
}
STDMETHODIMP CImageMgr::BitBltEx(OLE_HANDLE hDC,  long nImageId,  long nX,  long nY, long nRop, ImageStyles isStyle)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	HRESULT hResult;
	switch (isStyle)
	{
	case ddISNormal:
		hResult = pImageEntry->BitBlt((HDC)hDC, nX, nY, nRop);
		break;
	
	case ddISGray:
		hResult = pImageEntry->BitBltGray((HDC)hDC, nX, nY, nRop);
		break;
	
	case ddISMask:
		hResult = pImageEntry->BitBltMask((HDC)hDC, nX, nY, nRop);
		break;

	case ddISDisabled:
		hResult = pImageEntry->BitBltDisabled((HDC)hDC, nX, nY, nRop);
		break;
	}
	return hResult;
}
STDMETHODIMP CImageMgr::ScaleBlt( OLE_HANDLE hDC,  long nImageId,  long nX,  long nY,  long nWidth,  long nHeight,  long nRop, ImageStyles isStyle)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	HRESULT hResult;
	switch (isStyle)
	{
	case ddISNormal:
		hResult = pImageEntry->BitBlt((HDC)hDC, nX, nY, nRop, TRUE, nWidth, nHeight);
		break;
	
	case ddISGray:
		hResult = pImageEntry->BitBltGray((HDC)hDC, nX, nY, nRop, TRUE, nWidth, nHeight);
		break;
	
	case ddISMask:
		hResult = pImageEntry->BitBltMask((HDC)hDC, nX, nY, nRop, TRUE, nWidth, nHeight);
		break;

	case ddISDisabled:
		hResult = pImageEntry->BitBltDisabled((HDC)hDC, nX, nY, nRop, TRUE, nWidth, nHeight);
		break;
	}
	return hResult;
}
STDMETHODIMP CImageMgr::Exchange(IStream* pStream, BOOL bSave, DWORD dwVersion)
{
	ImageSizeEntry* pImageSizeEntry;
	HRESULT			hResult = S_OK;
	long			nKey = 0;
	int				nCount;
	int				nIndex;

	if (bSave)
	{
		switch (dwVersion)
		{
		case 0xB0B7:
			{
				Compact();
				
				pStream->Write(&m_nVersion, sizeof(unsigned short), 0);

				hResult = pStream->Write(&m_dwMasterImageId, sizeof(m_dwMasterImageId), 0);
				hResult = pStream->Write(&m_nDepth, sizeof(m_nDepth), 0);
				StWritePalette(pStream, m_hPal);

				nCount = m_mapImageSizes.GetCount();
				hResult = pStream->Write(&nCount, sizeof(nCount), 0);
				POSITION posMap = m_mapImageSizes.GetStartPosition();
				while (0 != posMap)
				{
					m_mapImageSizes.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageSizeEntry);
					hResult = pStream->Write(&nKey, sizeof(nKey), 0);
					hResult = pImageSizeEntry->Exchange(pStream, bSave, dwVersion);
				}
			}
			break;

		case 0xB0B8:
		case 0xB0C1:
			{
				POSITION posMap = m_mapImageSizes.GetStartPosition();
				while (0 != posMap)
				{
					m_mapImageSizes.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageSizeEntry);
					hResult = pImageSizeEntry->Exchange(pStream, bSave, dwVersion);
				}
			}
		}
	}
	else
	{
		switch (dwVersion)
		{
		case 0xB0B7:
			CleanUp();
			pStream->Read(&m_nVersion, sizeof(unsigned short), 0);
			hResult = pStream->Read(&m_dwMasterImageId, sizeof(m_dwMasterImageId), 0);
			hResult = pStream->Read(&m_nDepth, sizeof(m_nDepth), 0);

			StReadPalette(pStream, m_hPal);

			SIZE sizeImage;
			hResult = pStream->Read(&m_nCount, sizeof(m_nCount), 0);
			if (m_ppArrayImageSizeEntry)
			{
				delete [] m_ppArrayImageSizeEntry;
				m_ppArrayImageSizeEntry = NULL;
			}
			m_ppArrayImageSizeEntry = new ImageSizeEntry*[m_nCount];
			for (nIndex = 0; nIndex < m_nCount; nIndex++)
			{
				hResult = pStream->Read(&nKey, sizeof(nKey), 0);
				sizeImage.cx = LOWORD(nKey);
				sizeImage.cy = HIWORD(nKey);
				pImageSizeEntry = new ImageSizeEntry(sizeImage, *this);
				if (NULL == pImageSizeEntry)
					return E_FAIL;
				m_mapImageSizes.SetAt((LPVOID)nKey, (LPVOID&)pImageSizeEntry);
				hResult = pImageSizeEntry->Exchange(pStream, bSave, dwVersion);
				m_ppArrayImageSizeEntry[nIndex] = pImageSizeEntry; 
			}
			break;

		case 0xB0B8:
		case 0xB0C1:
			{
				for (nIndex = 0; nIndex < m_nCount; nIndex++)
					m_ppArrayImageSizeEntry[nIndex]->Exchange(pStream, bSave, dwVersion);
			}
			break;
		}
	}
	return hResult;
}
STDMETHODIMP CImageMgr::Compact()
{
	ImageSizeEntry* pImageSizeEntry;
	HRESULT			hResult = S_OK;
	long			nKey;

	POSITION posMap = m_mapImageSizes.GetStartPosition();
	while (0 != posMap)
	{
		m_mapImageSizes.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageSizeEntry);
		hResult = pImageSizeEntry->Compact();
	}
	return hResult;
}
STDMETHODIMP CImageMgr::ImageInitialSize( short nCx,  short nCy,  short nSize)
{
	long nKey = MAKELONG(nCx, nCy);
	
	ImageSizeEntry* pImageSizeEntry = NULL;
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		return E_FAIL;

	pImageSizeEntry->InitialSize(nSize, ImageSizeEntry::eImage);
	return S_OK;
}
STDMETHODIMP CImageMgr::ImageGrowBy( short nCx,  short nCy,  short nGrowBy)
{
	long nKey = MAKELONG(nCx, nCy);

	ImageSizeEntry* pImageSizeEntry = NULL;
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		return E_FAIL;
	
	pImageSizeEntry->GrowBy(nGrowBy, ImageSizeEntry::eImage);
	return S_OK;
}
STDMETHODIMP CImageMgr::MaskInitialSize( short nCx,  short nCy,  short nSize)
{
	long nKey = MAKELONG(nCx, nCy);

	ImageSizeEntry* pImageSizeEntry = NULL;
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		return E_FAIL;

	pImageSizeEntry->InitialSize(nSize, ImageSizeEntry::eMask);
	return S_OK;
}
STDMETHODIMP CImageMgr::MaskGrowBy( short nCx,  short nCy,  short nGrowBy)
{
	long nKey = MAKELONG(nCx, nCy);

	ImageSizeEntry* pImageSizeEntry = NULL;
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		return E_FAIL;

	pImageSizeEntry->GrowBy(nGrowBy, ImageSizeEntry::eMask);
	return S_OK;
}

//
// GetConvertInfo
//

STDMETHODIMP CImageMgr::GetConvertInfo(IStream* pStream, BOOL bSave, BitmapMgr& theBitmapMgr)
{
	if (bSave)
	{
		m_nVersion = 2;
		pStream->Write(&m_nVersion, sizeof(unsigned short), 0);
	}
	else
	{
		unsigned short nCount;
		DWORD    dwBaseCookie = 1;
		BOOL     bFix = FALSE;

		pStream->Read(&m_nVersion, sizeof(unsigned short), 0);
		if (2 == m_nVersion)
			return NOERROR;
		else if ((m_nVersion & 0xF000) == 0)
		{
			nCount = m_nVersion;
			bFix = TRUE;
		}
		else
			pStream->Read(&nCount, sizeof(unsigned short), 0);

		// new stuff ?
		if (!bFix) 
		{ 
			if (m_nVersion <= 0xFFFE)
				pStream->Read(&theBitmapMgr.sizeDefault, sizeof(SIZE), 0);
		}

		ImageInfo* pInfo;
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pInfo = new ImageInfo;
			if (NULL == pInfo)
				return E_FAIL;

			pStream->Read(&(pInfo->nImageIndex), sizeof(int), 0);
			pStream->Read(&(pInfo->bImageValid[0]), sizeof(BOOL)*4, 0);
			
			if (bFix) 
				memset(&(pInfo->bImageValid[1]), 0, sizeof(BOOL)*3);

			pStream->Read(&(pInfo->bMaskValid[0]), sizeof(BOOL)*4, 0);
			pStream->Read(&(pInfo->crMaskColor[0]), sizeof(COLORREF)*4, 0);
			pStream->Read(&(pInfo->dwCookie), sizeof(DWORD), 0);

			if (pInfo->dwCookie >= dwBaseCookie)
				dwBaseCookie = pInfo->dwCookie + 1;
			
			// Used only for special sized images
			if (-1 == pInfo->nImageIndex)
			{
				for (int nIndex2 = 0; nIndex2 < 4; nIndex2++)
				{
					StReadHBITMAP(pStream, &(pInfo->bmp[nIndex2]), 0, FALSE, m_hPal);
					StReadHBITMAP(pStream, &(pInfo->bmpMask[nIndex2]), 0, TRUE, m_hPal);
				}
			}
			theBitmapMgr.faImages.Add(pInfo);
		}
		
		// Load image lists
		for (nIndex = 0; nIndex < 4; nIndex++)
		{
			StReadHBITMAP(pStream, &(theBitmapMgr.hImage[nIndex]), 0, FALSE, m_hPal);
			StReadHBITMAP(pStream, &(theBitmapMgr.hMask[nIndex]), 0, TRUE, m_hPal);
		}
		
		pStream->Read(theBitmapMgr.nImageCount, 4*sizeof(int), 0);
		pStream->Read(theBitmapMgr.nMaxCount, 4*sizeof(int), 0);
	}
	return NOERROR;
}

//
// Convert
//

STDMETHODIMP CImageMgr::Convert(BitmapMgr& theBitmapMgr, FArray& faTools)
{
	ImageInfo* pImageInfo;
	BITMAP     bmInfo;
	AB10Tool*  pTool;
	SIZE	   sizeImage;
	int		   nImageType;
	int		   nImage;

	// If there are no images there is nothing to convert
	int nSizeImages = theBitmapMgr.faImages.GetSize();
	if (0 == nSizeImages)
		return NOERROR;
	
	int nSizeTools = faTools.GetSize();
	for (int nTool = 0; nTool < nSizeTools; nTool++)
	{
		pTool = (AB10Tool*)faTools.GetAt(nTool);
		if (NULL == pTool || 0 == pTool->m_dwImageCookie)
			continue;

		for (nImage = 0; nImage < nSizeImages; nImage++)
		{
			pImageInfo = (ImageInfo*)theBitmapMgr.faImages.GetAt(nImage);
			if (pImageInfo->dwCookie == pTool->m_dwImageCookie)
				break;
			pImageInfo = NULL;
		}

		if (NULL != pImageInfo)
		{
			for (nImageType = 0; nImageType < AB10Tool::eNumOfImageTypes; nImageType++)
			{
				if (pImageInfo->bImageValid[nImageType])
				{
					if (-1 == pImageInfo->nImageIndex)
					{
						GetObject(pImageInfo->bmp[nImageType], sizeof(BITMAP), &bmInfo);
						sizeImage.cx = bmInfo.bmWidth;
						sizeImage.cy = bmInfo.bmHeight;
						CreateImage2(&pTool->m_nImageIds[nImageType], sizeImage.cx, sizeImage.cy);
						put_ImageBitmap(pTool->m_nImageIds[nImageType], (OLE_HANDLE)pImageInfo->bmp[nImageType]);
					}
					else
					{
						CreateImage2(&pTool->m_nImageIds[nImageType], 
									 theBitmapMgr.sizeDefault.cx, 
									 theBitmapMgr.sizeDefault.cy);
						HBITMAP hImage;
						theBitmapMgr.GetImageBitmap(pImageInfo, nImageType, hImage);
						put_ImageBitmap(pTool->m_nImageIds[nImageType], (OLE_HANDLE)hImage);
						DeleteBitmap(hImage);
					}
					if (pImageInfo->bMaskValid[nImageType])
					{
						if (-1 == pImageInfo->nImageIndex)
						{
							put_MaskBitmap(pTool->m_nImageIds[nImageType], 
										   (OLE_HANDLE)pImageInfo->bmpMask[nImageType]);
						}
						else
						{
							HBITMAP hMask;
							theBitmapMgr.GetMaskBitmap(pImageInfo, nImageType, hMask);
							put_MaskBitmap(pTool->m_nImageIds[nImageType], (OLE_HANDLE)hMask);
							DeleteBitmap(hMask);
						}
					}
				}
			}
		}
	}
	return NOERROR;
}

//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
//
// ImageInfo
//

ImageInfo::ImageInfo()
{
	dwCookie = -1;
	nImageIndex = -1; 
	for (int nIndex = 0; nIndex < 4; nIndex++)
		bImageValid[nIndex] = FALSE;

	for (nIndex = 0; nIndex < 4; nIndex++)
		bMaskValid[nIndex] = FALSE;
	
	for (nIndex = 0; nIndex < 4; nIndex++)
		crMaskColor[nIndex] = RGB(255,255,255);

	// Used only for special sized images
	for (nIndex = 0; nIndex < 4; nIndex++)
		bmp[nIndex] = NULL;
	
	for (nIndex = 0; nIndex < 4; nIndex++)
		bmpMask[nIndex] = NULL;
	
	bmpGray = NULL;
}

ImageInfo::~ImageInfo()
{
	for (int nIndex = 0; nIndex < 4; nIndex++)
	{
		if (NULL != bmp[nIndex])
		{
			int nResult = DeleteBitmap(bmp[nIndex]);
			assert(nResult);
		}
	}

	for (nIndex = 0; nIndex < 4; nIndex++)
	{
		if (NULL != bmpMask[nIndex])
		{
			int nResult = DeleteBitmap(bmpMask[nIndex]);
			assert(nResult);
		}
	}
	
	if (NULL != bmpGray)
	{
		int nResult = DeleteBitmap(bmpGray);
		assert(nResult);
	}
}

//
// BitmapMgr
//

BitmapMgr::BitmapMgr()
{
	for (int nIndex = 0; nIndex < 4; nIndex++)
		hImage[nIndex] = NULL;
	
	for (nIndex = 0; nIndex < 4; nIndex++)
		hMask[nIndex] = NULL;
	
	sizeDefault.cx = sizeDefault.cy = 16;
}

BitmapMgr::~BitmapMgr()
{
	for (int nIndex = 0; nIndex < 4; nIndex++)
	{
		if (NULL != hImage[nIndex])
		{
			int nResult = DeleteObject(hImage[nIndex]);
			assert(nResult);
		}
	}
	
	for (nIndex = 0; nIndex < 4; nIndex++)
	{
		if (NULL != hMask[nIndex])
		{
			int nResult = DeleteObject(hMask[nIndex]);
			assert(nResult);
		}
	}
	
	int nSize = faImages.GetSize();
	for (nIndex = 0; nIndex < nSize; nIndex++)
		delete ((ImageInfo*)faImages.GetAt(nIndex));
}

static HBITMAP CreateBitmap(const ImageInfo* pInfo, 
							const int&       nIndex, 
							const HPALETTE	 hPal, 
							const HBITMAP&	 bmSource, 
							const HBITMAP&	 bmMask,
							const SIZE		 sizeBitmap,
							const BOOL&		 bMask)
{
	// now copy sourceBitmap at position imageIndex to new bitmap
	HBITMAP bmNew, bmOld, bmOld2;
	HPALETTE palOld1, palOld2;

	HDC hdc = GetDC(NULL);
	if (bMask)
		bmNew = CreateBitmap(sizeBitmap.cx,sizeBitmap.cy,1,1,0);
	else
		bmNew = CreateCompatibleBitmap(hdc,sizeBitmap.cx,sizeBitmap.cy);

	HDC hMemDC1 = CreateCompatibleDC(hdc);
	HDC hMemDC2 = GetGlobals().GetMemDC();

	if (hPal)
	{
		palOld1 = SelectPalette(hMemDC1,hPal,FALSE);
		palOld2 = SelectPalette(hMemDC2,hPal,FALSE);
		RealizePalette(hMemDC1);
		RealizePalette(hMemDC2);
	}

	bmOld = SelectBitmap(hMemDC1, bmNew);
	
	PatBlt(hMemDC1, 0, 0, sizeBitmap.cx, sizeBitmap.cy, BLACKNESS);

	::SetTextColor(hMemDC1,RGB(0,0,0));
	::SetBkColor(hMemDC1,RGB(255,255,255));
	::SetTextColor(hMemDC2,RGB(0,0,0));
	::SetBkColor(hMemDC2,RGB(255,255,255));
	if (bMask == FALSE && pInfo->bMaskValid[nIndex])
	{
		// mask out the grey around the base image
		bmOld2 = SelectBitmap(hMemDC2,bmSource);
		::BitBlt(hMemDC1,
				 0,
				 0,
				 sizeBitmap.cx,
				 sizeBitmap.cy,
				 hMemDC2,
				 pInfo->nImageIndex*sizeBitmap.cx,
				 0,
				 SRCINVERT);

		SetBkColor(hMemDC2,RGB(0,0,0));
		SetTextColor(hMemDC2,RGB(255,255,255));

		SelectObject(hMemDC2, bmMask);
		::BitBlt(hMemDC1,
				 0,
				 0,
				 sizeBitmap.cx,
				 sizeBitmap.cy,
				 hMemDC2,
				 pInfo->nImageIndex*sizeBitmap.cx,
				 0,
				 SRCAND);

		SelectObject(hMemDC2, bmSource);
		::BitBlt(hMemDC1,
				 0,
				 0,
				 sizeBitmap.cx,
				 sizeBitmap.cy,
				 hMemDC2,
				 pInfo->nImageIndex*sizeBitmap.cx,
				 0,
				 SRCINVERT);
		SelectObject(hMemDC2,bmOld2);
	}
	else if (bmSource != 0)
	{
		bmOld2 = SelectBitmap(hMemDC2, bmSource);
		::BitBlt(hMemDC1,
				 0,
				 0,
				 sizeBitmap.cx,
				 sizeBitmap.cy,
				 hMemDC2,
				 pInfo->nImageIndex*sizeBitmap.cx,
				 0,
				 SRCCOPY);
		SelectObject(hMemDC2, bmOld2);
	}

	if (hPal)
	{
		SelectPalette(hMemDC1,palOld1,FALSE);
		SelectPalette(hMemDC2,palOld2,FALSE);
	}

	SelectObject(hMemDC1, bmOld);
	
	DeleteDC(hMemDC1);
	
	ReleaseDC(0, hdc); 
	return bmNew;
}

void BitmapMgr::GetImageBitmap(ImageInfo* pImageInfo, int nType, HBITMAP& hImageBitmap)
{
	HBITMAP bmSourceMask = NULL;
	HBITMAP bmSource = hImage[nType];
	if (pImageInfo->bMaskValid[nType])
		bmSourceMask = hMask[nType];
	hImageBitmap = CreateBitmap(pImageInfo, nType, 0, bmSource, bmSourceMask, sizeDefault, FALSE);
}

void BitmapMgr::GetMaskBitmap(ImageInfo* pImageInfo, int nType, HBITMAP& hImageBitmap)
{
	HBITMAP bmSourceMask = 0;
	HBITMAP bmSource = hMask[nType];
	hImageBitmap = CreateBitmap(pImageInfo, nType, 0, bmSource, bmSourceMask, sizeDefault, TRUE);
}

//
// ImageEntry
//

ImageEntry::ImageEntry(ImageSizeEntry* pImageSizeEntry, DWORD dwImageId)
	: m_pImageSizeEntry(pImageSizeEntry),
	  m_dwImageId(dwImageId)
{
	m_nImageIndex = -1;
	m_nMaskIndex = -1;
	m_nRefCnt = 1;
	m_ocMaskColor = -1;
}

ImageEntry::ImageEntry()
	: m_pImageSizeEntry(NULL),
	  m_dwImageId(0)
{
	m_nImageIndex = -1;
	m_nMaskIndex = -1;
	m_nRefCnt = 1;
	m_ocMaskColor = -1;
}

//
// ImageEntry
//

ImageEntry::~ImageEntry()
{
	ReleaseImage();
	ReleaseMask();
	m_pImageSizeEntry->Remove(this);
}

ULONG ImageEntry::Release()
{
	if (0 != --m_nRefCnt)
		return m_nRefCnt;
	delete this;
	return 0;
}

HRESULT ImageEntry::GetMaskColor(OLE_COLOR* ocMaskColor)
{
	*ocMaskColor = m_ocMaskColor;
	return NOERROR;
}

HRESULT ImageEntry::SetMaskColor(OLE_COLOR ocMaskColor)
{
	m_ocMaskColor = ocMaskColor;
	return NOERROR;
}

//
// GetSize
//

STDMETHODIMP ImageEntry::GetSize(SIZE& sizeImage)
{
	if (m_pImageSizeEntry)
		m_pImageSizeEntry->GetSize(sizeImage);
	else
		sizeImage.cx = sizeImage.cy = 0;
	return S_OK;
}

//
// GetImageBitmap
//

STDMETHODIMP ImageEntry::GetImageBitmap(HBITMAP& hBitmapImage)
{
	return m_pImageSizeEntry->GetImageBitmap(m_nImageIndex, m_nMaskIndex, hBitmapImage);
}

//
// SetImageBitmap
//

STDMETHODIMP ImageEntry::SetImageBitmap(HBITMAP hBitmapImage)
{
	return m_pImageSizeEntry->SetImageBitmap(m_nImageIndex, hBitmapImage);
}

//
// GetMaskBitmap
//

STDMETHODIMP ImageEntry::GetMaskBitmap(HBITMAP& hBitmapMask)
{
	if (-1 == m_nMaskIndex)
	{
		OLE_COLOR ocMaskColor;
		GetMaskColor(&ocMaskColor);
		return m_pImageSizeEntry->GetMaskBitmap(m_nImageIndex, FALSE, hBitmapMask, ocMaskColor);
	}
	else
		return m_pImageSizeEntry->GetMaskBitmap(m_nMaskIndex, TRUE, hBitmapMask);
}

//
// SetMaskBitmap
//

STDMETHODIMP ImageEntry::SetMaskBitmap(HBITMAP hBitmapMask)
{
	if (NULL == hBitmapMask)
	{
		ReleaseMask();
		return NOERROR;
	}
	return m_pImageSizeEntry->SetMaskBitmap(m_nMaskIndex, hBitmapMask);
}

//
// BitBlt
//

STDMETHODIMP ImageEntry::BitBlt(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	return m_pImageSizeEntry->BitBlt(hDC, m_nImageIndex, x, y, dwRop, bScale, nHeight, nWidth);
}

//
// BitBltGray
//

STDMETHODIMP ImageEntry::BitBltGray(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	return m_pImageSizeEntry->BitBltGray(hDC, m_nImageIndex, x, y, dwRop, bScale, nHeight, nWidth);
}

//
// BitBltMask
//

STDMETHODIMP ImageEntry::BitBltMask(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	HRESULT hResult;
	if (-1 == m_nMaskIndex)
		hResult = m_pImageSizeEntry->BitBltMask(hDC, FALSE, m_nImageIndex, x, y, dwRop, m_ocMaskColor, bScale, nHeight, nWidth);
	else
		hResult = m_pImageSizeEntry->BitBltMask(hDC, TRUE, m_nMaskIndex, x, y, dwRop, m_ocMaskColor, bScale, nHeight, nWidth);

	return hResult;
}

//
// BitBltDisabled
//

STDMETHODIMP ImageEntry::BitBltDisabled(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	return m_pImageSizeEntry->BitBltDisabled(hDC, m_nImageIndex, m_nMaskIndex, x, y, dwRop, m_ocMaskColor, bScale, nHeight, nWidth);
}

//
// ReleaseImage
//

STDMETHODIMP ImageEntry::ReleaseImage()
{
	if (m_nImageIndex == -1)
		return S_OK;
	HRESULT hResult = m_pImageSizeEntry->RemoveImage(m_nImageIndex);
	return hResult;
}

//
// ReleaseMask
//

STDMETHODIMP ImageEntry::ReleaseMask()
{
	if (-1 == m_nMaskIndex)
		return S_OK;
	HRESULT hResult = m_pImageSizeEntry->RemoveMask(m_nMaskIndex);
	m_nMaskIndex = -1;
	return hResult;
}

//
// Exchange
//

STDMETHODIMP ImageEntry::Exchange(IStream* pStream, BOOL bSave, DWORD dwVersion)
{
	HRESULT hResult;
	if (bSave)
	{
		switch (dwVersion)
		{
		case 0xB0B7:
			hResult = pStream->Write(&m_nImageIndex, sizeof(m_nImageIndex), 0);
			hResult = pStream->Write(&m_nMaskIndex, sizeof(m_nMaskIndex), 0);
			hResult = pStream->Write(&m_dwImageId, sizeof(m_dwImageId), 0);
			break;
		
		case 0xB0B8:
			hResult = pStream->Write(&m_nRefCnt, sizeof(m_nRefCnt), 0);
			break;

		case 0xB0C1:
			hResult = pStream->Write(&m_ocMaskColor, sizeof(m_ocMaskColor), 0);
			break;
		}
	}
	else
	{
		switch (dwVersion)
		{
		case 0xB0B7:
			hResult = pStream->Read(&m_nImageIndex, sizeof(m_nImageIndex), 0);
			hResult = pStream->Read(&m_nMaskIndex, sizeof(m_nMaskIndex), 0);
			hResult = pStream->Read(&m_dwImageId, sizeof(m_dwImageId), 0);
			m_nRefCnt = 0;
			break;

		case 0xB0B8:
			hResult = pStream->Read(&m_nRefCnt, sizeof(m_nRefCnt), 0);
			break;

		case 0xB0C1:
			hResult = pStream->Read(&m_ocMaskColor, sizeof(m_ocMaskColor), 0);
			break;
		}
	}
	return hResult;
}

//
// ImageSizeEntry
//

ImageSizeEntry::ImageSizeEntry(const SIZE& sizeImage, CImageMgr& theImageMgr)
	: m_sizeImage(sizeImage),
	  m_theImageMgr(theImageMgr),
	  m_hBitmapGray(NULL),
	  m_ppArrayImageEntry(NULL)
{
	if (sizeImage.cx > CImageMgr::eDefaultImageSize)
	{
		m_nImageInitial =  CImageMgr::eDefaultImageInitialElements * CImageMgr::eDefaultImageSize / sizeImage.cx;
		if (0 == m_nImageInitial)
			m_nImageInitial = 1;
		m_nMaskInitial = m_nImageInitial;
	}
	else
	{
		m_nImageInitial = CImageMgr::eDefaultImageInitialElements;
		m_nMaskInitial = CImageMgr::eDefaultMaskInitialElements;
	}
	m_nImageGrowBy = CImageMgr::eDefaultImageGrowBy;
	m_nMaskGrowBy = CImageMgr::eDefaultMaskGrowBy;
	m_hMask = ::CreateBitmap(sizeImage.cx, sizeImage.cy, 1, 1, 0);
}

//
// ~ImageSizeEntry
//

ImageSizeEntry::~ImageSizeEntry()
{
	if (m_hBitmapGray)
		DeleteBitmap(m_hBitmapGray);

	long nKey;
	ImageEntry* pImageEntry;
	POSITION posMap = m_mapImages.GetStartPosition();
	while (0 != posMap)
	{
		m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
		delete pImageEntry;
	}

	int nSize = m_faImageBitmaps.GetSize();
	for (int nIndex = 0; nIndex < nSize; nIndex++)
		DeleteBitmap(m_faImageBitmaps.GetAt(nIndex));

	nSize = m_faMaskBitmaps.GetSize();
	for (nIndex = 0; nIndex < nSize; nIndex++)
		DeleteBitmap(m_faMaskBitmaps.GetAt(nIndex));

	if (m_hMask)
		DeleteBitmap(m_hMask);

	m_faImageBitmaps.RemoveAll();
	m_faImageFreeList.RemoveAll();
	m_faMaskBitmaps.RemoveAll();
	m_faMaskFreeList.RemoveAll();
	m_mapImages.RemoveAll();
	delete [] m_ppArrayImageEntry;
}

//
// GetPalette
//

void ImageSizeEntry::GetPalette(HPALETTE& hPal)
{
	m_theImageMgr.get_Palette((OLE_HANDLE*)&hPal);
}

//
// GetColorDepth
//

void ImageSizeEntry::GetColorDepth(short& nDepth)
{
	m_theImageMgr.get_BitDepth(&nDepth);
}

//
// CreateMaskBitmap
//

void ImageSizeEntry::CreateMaskBitmap(HDC hDC, HBITMAP hBitmap, int nRelativeIndex, HBITMAP& hBitmapMask, COLORREF crColor)
{
    // Create memory DCs to work with.
    HDC hDCMask = ::CreateCompatibleDC(hDC);
    HDC hDCImage = ::CreateCompatibleDC(hDC);

	// Select the mono bitmap into its DC.
	if (NULL == hBitmapMask)
		hBitmapMask	= ::CreateBitmap(m_sizeImage.cx, m_sizeImage.cy, 1, 1, 0);

	HBITMAP hbmMaskOld = SelectBitmap(hDCMask, hBitmapMask);
    
	// Select the image bitmap into its DC.
    HBITMAP hbmImageOld = SelectBitmap(hDCImage, hBitmap);
    
	if (-1 == crColor)
	{
		// Set the transparency color to be the top-left pixel.
		COLORREF crColor = ::GetPixel(hDCImage, (nRelativeIndex*m_sizeImage.cx), 0);
		int nHit = 1;
		if (crColor == GetPixel(hDCImage, ((nRelativeIndex+1)*m_sizeImage.cx)-1, m_sizeImage.cy-1)) 
			nHit++;
		if (crColor == GetPixel(hDCImage, (nRelativeIndex*m_sizeImage.cx), m_sizeImage.cy-1)) 
			nHit++;
		if (crColor == GetPixel(hDCImage, ((nRelativeIndex+1)*m_sizeImage.cx)-1, m_sizeImage.cy-1)) 
			nHit++;

		if (nHit > 1)
			SetBkColor(hDCImage, crColor);
		else 
			SetBkColor(hDCImage, RGB(192,192,192));
	}
	else
		SetBkColor(hDCImage, crColor);

	SetTextColor(hDCImage, RGB(255,255,255));

    // Make the mask.
	::BitBlt(hDCMask, 
			 0,
			 0,
			 m_sizeImage.cx,
			 m_sizeImage.cy,
			 hDCImage,
			 nRelativeIndex*m_sizeImage.cx,
			 0,
			 SRCCOPY);

    // Tidy up.
    SelectBitmap(hDCMask, hbmMaskOld);
    SelectBitmap(hDCImage, hbmImageOld);
    DeleteDC(hDCMask);
    DeleteDC(hDCImage);
}

//
// Add
//

STDMETHODIMP ImageSizeEntry::Add(ImageEntry* pImageEntry)
{
	DWORD dwImageId;
	pImageEntry->GetImageId(dwImageId);
	m_mapImages.SetAt((LPVOID)dwImageId, (LPVOID)pImageEntry);
	return S_OK;
}

//
// Remove
//

STDMETHODIMP ImageSizeEntry::Remove(ImageEntry* pImageEntry)
{
	DWORD dwImageId;
	pImageEntry->GetImageId(dwImageId);
	m_mapImages.RemoveKey((LPVOID)dwImageId);
	return S_OK;
}

//
// GetImageBitmap
//

STDMETHODIMP ImageSizeEntry::GetImageBitmap(const int& nIndex, const int& nMaskIndex, HBITMAP& hBitmapImage)
{
	if (nIndex < 0)
		return E_FAIL;

	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return E_FAIL;
	
	HPALETTE hPal;
	GetPalette(hPal); 

	int nBitmapIndex = BitmapIndex(nIndex, eImage);
	int nRelativeIndex = RelativeIndex(nIndex, eImage);
	HBITMAP hImage = (HBITMAP)m_faImageBitmaps.GetAt(nBitmapIndex);

	int nMaskBitmapIndex = -1;
	int nMaskRelativeIndex = -1;
	HBITMAP hMask = 0;
	if (-1 != nMaskIndex)
	{
		nMaskBitmapIndex = BitmapIndex(nMaskIndex, eMask);
		nMaskRelativeIndex = RelativeIndex(nMaskIndex, eMask);
		hMask = (HBITMAP)m_faMaskBitmaps.GetAt(nMaskBitmapIndex);
	}

	hBitmapImage = CreateCompatibleBitmap(hDC, 
										  m_sizeImage.cx, 
										  m_sizeImage.cy);
	
	HDC hDCDest = CreateCompatibleDC(hDC);
	HDC hDCImage = GetGlobals().GetMemDC();

	HPALETTE hPalDestOld, hPalImageOld;
	if (hPal)
	{
		hPalDestOld = SelectPalette(hDCDest, hPal, FALSE);
		RealizePalette(hDCDest);

		hPalImageOld = SelectPalette(hDCImage, hPal, FALSE);
		RealizePalette(hDCImage);
	}
	
	HBITMAP hBitmapDestOld = SelectBitmap(hDCDest, hBitmapImage);
	
	PatBlt(hDCDest, 
		   0, 
		   0, 
		   m_sizeImage.cx, 
		   m_sizeImage.cy, 
		   BLACKNESS);

	::SetTextColor(hDCDest,RGB(0,0,0));
	::SetBkColor(hDCDest,RGB(255,255,255));

	HBITMAP hBitmapImageOld = SelectBitmap(hDCImage, hImage);

	// Mask out the grey around the base image
	int nXPosImage = nRelativeIndex*m_sizeImage.cx;
	if (NULL == hMask)
	{
		::BitBlt(hDCDest, 
				 0, 
				 0, 
				 m_sizeImage.cx, 
				 m_sizeImage.cy,
				 hDCImage,
				 nXPosImage,
				 0,
				 SRCCOPY);
	}
	else
	{
		int nXPosMask = nMaskRelativeIndex*m_sizeImage.cx;
		HDC hDCMask = CreateCompatibleDC(hDC);
		HPALETTE hPalMaskOld;
		if (hPal)
		{
			hPalMaskOld = SelectPalette(hDCMask, hPal, FALSE);
			RealizePalette(hDCMask);
		}

		::BitBlt(hDCDest,
				 0,
				 0,
				 m_sizeImage.cx,
				 m_sizeImage.cy,
				 hDCImage,
				 nXPosImage,
				 0,
				 SRCINVERT);

		HBITMAP hBitmapMaskOld = SelectBitmap(hDCMask, hMask);

		::BitBlt(hDCDest,
				 0,
				 0,
				 m_sizeImage.cx,
				 m_sizeImage.cy,
				 hDCMask,
				 nXPosMask,
				 0,
				 SRCAND);
		SelectBitmap(hDCMask, hBitmapMaskOld);

		::BitBlt(hDCDest,
				 0,
				 0,
				 m_sizeImage.cx,
				 m_sizeImage.cy,
				 hDCImage,
				 nXPosImage,
				 0,
				 SRCINVERT);
		
		if (hPal)
			SelectPalette(hDCMask, hPalMaskOld, FALSE);
		DeleteDC(hDCMask);
	}

	if (hPal)
	{
		SelectPalette(hDCDest, hPalDestOld, FALSE);
		SelectPalette(hDCImage, hPalImageOld, FALSE);
	}

	SelectBitmap(hDCDest, hBitmapDestOld);
	SelectBitmap(hDCImage, hBitmapImageOld);

	DeleteDC(hDCDest);
	ReleaseDC(NULL, hDC);
	return S_OK;
}

//
// SetImageBitmap
//

STDMETHODIMP ImageSizeEntry::SetImageBitmap(int& nIndex, HBITMAP hBitmapImage)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return E_FAIL;

	HBITMAP hImages;
	if (-1 == nIndex)
	{
		if (0 == m_faImageBitmaps.GetSize())
		{
			int nWidth = m_sizeImage.cx * m_nImageInitial;
			hImages = CreateCompatibleBitmap(hDC, nWidth, m_sizeImage.cy);
			m_faImageBitmaps.Add(hImages);
			nIndex = 0;
			for (int nIndex2 = 1; nIndex2 < m_nImageInitial; nIndex2++)
				m_faImageFreeList.Add((LPVOID)nIndex2);

		}
		else if (m_faImageFreeList.GetSize() > 0)
		{
			nIndex = (int)m_faImageFreeList.GetAt(0);
			m_faImageFreeList.RemoveAt(0);
			hImages = (HBITMAP)m_faImageBitmaps.GetAt(BitmapIndex(nIndex, eImage));
		}
		else
		{
			int nSize = m_faImageBitmaps.GetSize();
			int nWidth = m_sizeImage.cx * m_nImageGrowBy;
			hImages = CreateCompatibleBitmap(hDC, nWidth, m_sizeImage.cy);
			m_faImageBitmaps.Add(hImages);
			nIndex = m_nImageInitial + ((nSize-1) * m_nImageGrowBy);
			for (int nIndex2 = nIndex+1; nIndex2 < m_nImageGrowBy+nIndex; nIndex2++)
				m_faImageFreeList.Add((LPVOID)nIndex2);
		}
	}
	else
		hImages = (HBITMAP)m_faImageBitmaps.GetAt(BitmapIndex(nIndex, eImage));

	int nRelativeIndex = RelativeIndex(nIndex, eImage);
	HPALETTE hPal;
	GetPalette(hPal); 

	HDC hDCSrc = CreateCompatibleDC(hDC);
	HDC hDCDest = GetGlobals().GetMemDC();

	HBITMAP hBitmapOld1 = SelectBitmap(hDCSrc, hBitmapImage);
	HBITMAP hBitmapOld2 = SelectBitmap(hDCDest, hImages);

	HPALETTE hPalOld1, hPalOld2;
	if (hPal)
	{
		hPalOld1 = SelectPalette(hDCSrc, hPal, FALSE);
		RealizePalette(hDCSrc);

		hPalOld2 = SelectPalette(hDCDest, hPal, FALSE);
		RealizePalette(hDCDest);
	}
	
	::BitBlt(hDCDest, 
		     nRelativeIndex*m_sizeImage.cx, 
		     0, 
		     m_sizeImage.cx, 
		     m_sizeImage.cy, 
		     hDCSrc, 
		     0, 
		     0,
		     SRCCOPY);

	if (hPal)
	{
		SelectPalette(hDCSrc, hPalOld1, FALSE);
		SelectPalette(hDCDest, hPalOld2, FALSE);
	}

	SelectBitmap(hDCSrc, hBitmapOld1);
	SelectBitmap(hDCDest, hBitmapOld2);

	DeleteDC(hDCSrc);
	ReleaseDC(NULL, hDC);	
	return S_OK;
}

//
// GetMaskBitmap
//

STDMETHODIMP ImageSizeEntry::GetMaskBitmap(const int& nIndex, BOOL bMask, HBITMAP& hBitmapMask, OLE_COLOR ocColor)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return E_FAIL;

	if (!bMask)
	{
		COLORREF crColor = -1;
		if (-1 != ocColor)
			OleTranslateColor(ocColor, NULL, &crColor);
		CreateMaskBitmap(hDC, 
						 (HBITMAP)m_faImageBitmaps.GetAt(BitmapIndex(nIndex, eImage)), 
						 RelativeIndex(nIndex, eImage), 
						 hBitmapMask,
						 crColor);
	}
	else
	{
		int nMaskBitmapIndex = BitmapIndex(nIndex, eMask);
		int nMaskRelativeIndex = RelativeIndex(nIndex, eMask);
		HBITMAP	hMask = (HBITMAP)m_faMaskBitmaps.GetAt(nMaskBitmapIndex);

		hBitmapMask = CreateBitmap(m_sizeImage.cx, 
								   m_sizeImage.cy,
								   1,
								   1,
								   0);
		
		HDC hDCDest = CreateCompatibleDC(hDC);
		HDC hDCMask = GetGlobals().GetMemDC();

		HPALETTE hPal;
		GetPalette(hPal); 
		HPALETTE hPalDestOld, hPalMaskOld;
		if (hPal)
		{
			hPalDestOld = SelectPalette(hDCDest, hPal, FALSE);
			RealizePalette(hDCDest);

			hPalMaskOld = SelectPalette(hDCMask, hPal, FALSE);
			RealizePalette(hDCMask);
		}
		
		HBITMAP hBitmapDestOld = SelectBitmap(hDCDest, hBitmapMask);
		
		::SetTextColor(hDCDest,RGB(0,0,0));
		::SetBkColor(hDCDest,RGB(255,255,255));
		
		HBITMAP hBitmapMaskOld = SelectBitmap(hDCMask, hMask);
		
		::BitBlt(hDCDest, 
				 0, 
				 0, 
				 m_sizeImage.cx, 
				 m_sizeImage.cy,
				 hDCMask,
				 nMaskRelativeIndex*m_sizeImage.cx,
				 0,
				 SRCCOPY);

		SelectBitmap(hDCMask, hBitmapMaskOld);

		if (hPal)
		{
			SelectPalette(hDCDest, hPalDestOld, FALSE);
			SelectPalette(hDCMask, hPalMaskOld, FALSE);
		}

		SelectBitmap(hDCDest, hBitmapDestOld);

		DeleteDC(hDCDest);
	}
	ReleaseDC(NULL, hDC);
	return S_OK;
}

//
// SetMaskBitmap
//

STDMETHODIMP ImageSizeEntry::SetMaskBitmap(int& nIndex, HBITMAP hBitmapMask)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return E_FAIL;

	HBITMAP hMask;
	if (-1 == nIndex)
	{
		if (0 == m_faMaskBitmaps.GetSize())
		{
			// Initial allocation
			int nWidth = m_sizeImage.cx * m_nMaskInitial;
			hMask = CreateBitmap(nWidth, m_sizeImage.cy, 1, 1, 0);
			m_faMaskBitmaps.Add(hMask);
			nIndex = 0;
			for (int nIndex2 = 1; nIndex2 < m_nMaskInitial; nIndex2++)
				m_faMaskFreeList.Add((LPVOID)nIndex2);
		}
		else if (m_faMaskFreeList.GetSize() > 0)
		{
			// Check the free list
			nIndex = (int)m_faMaskFreeList.GetAt(0);
			m_faMaskFreeList.RemoveAt(0);
			hMask = (HBITMAP)m_faMaskBitmaps.GetAt(BitmapIndex(nIndex, eMask));
		}
		else
		{
			int nSize = m_faMaskBitmaps.GetSize();
			int nWidth = m_sizeImage.cx * m_nMaskGrowBy;
			hMask = CreateBitmap(nWidth, m_sizeImage.cy, 1, 1, 0);
			m_faMaskBitmaps.Add(hMask);
			nIndex = m_nMaskInitial + ((nSize-1) *  m_nMaskGrowBy);
			for (int nIndex2 = nIndex+1; nIndex2 < m_nMaskGrowBy+nIndex; nIndex2++)
				m_faMaskFreeList.Add((LPVOID)nIndex2);
		}
	}
	else
		hMask = (HBITMAP)m_faMaskBitmaps.GetAt(BitmapIndex(nIndex, eMask));

	int nRelativeIndex = RelativeIndex(nIndex, eMask);

	HDC hDCSrc = CreateCompatibleDC(hDC);
	HDC hDCDest = CreateCompatibleDC(hDC);

	HBITMAP hBitmapOld1 = SelectBitmap(hDCSrc, hBitmapMask);
	HBITMAP hBitmapOld2 = SelectBitmap(hDCDest, hMask);

	HPALETTE hPal;
	GetPalette(hPal); 
	HPALETTE hPalOld1, hPalOld2;
	if (hPal)
	{
		hPalOld1 = SelectPalette(hDCSrc, hPal, FALSE);
		RealizePalette(hDCSrc);

		hPalOld2 = SelectPalette(hDCDest, hPal, FALSE);
		RealizePalette(hDCDest);
	}
	
	int nXPos = nRelativeIndex*m_sizeImage.cx;
	::BitBlt(hDCDest, 
		     nXPos, 
		     0, 
		     m_sizeImage.cx, 
		     m_sizeImage.cy, 
		     hDCSrc, 
		     0, 
		     0,
		     SRCCOPY);

	if (hPal)
	{
		SelectPalette(hDCSrc, hPalOld1, FALSE);
		SelectPalette(hDCDest, hPalOld2, FALSE);
	}

	SelectBitmap(hDCSrc, hBitmapOld1);
	SelectBitmap(hDCDest, hBitmapOld2);

	DeleteDC(hDCSrc);
	DeleteDC(hDCDest);
	ReleaseDC(NULL, hDC);
	return S_OK;
}

//
// BitBlt
//

STDMETHODIMP ImageSizeEntry::BitBlt (HDC hDC, int nIndex, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	HDC hDCMem = GetGlobals().GetMemDC();
	if (NULL == hDCMem)
		return E_FAIL;

	int nBitmapIndex = BitmapIndex(nIndex, eImage);
	int nRelativeIndex = RelativeIndex(nIndex, eImage);
	HBITMAP hImage = (HBITMAP)m_faImageBitmaps.GetAt(nBitmapIndex);

	HPALETTE hPal, hPalOld;
	GetPalette(hPal); 
	if (hPal)
	{
		hPalOld = SelectPalette(hDC, hPal, FALSE);
		RealizePalette(hDC);
	}

	HBITMAP hBitmapOld = SelectBitmap(hDCMem, hImage);
	
	if (bScale)
	{
		::StretchBlt(hDC, 
					 x, 
					 y,
					 nHeight,
					 nWidth,
					 hDCMem,
					 nRelativeIndex*m_sizeImage.cx,
					 0,
					 m_sizeImage.cx,
					 m_sizeImage.cy,
					 dwRop);
	}
	else
	{
		::BitBlt(hDC, 
				 x, 
				 y, 
				 m_sizeImage.cx, 
				 m_sizeImage.cy, 
				 hDCMem, 
				 nRelativeIndex*m_sizeImage.cx, 
				 0,
				 dwRop);
	}
	SelectBitmap(hDCMem, hBitmapOld);
	
	if (hPal)
		SelectPalette(hDC, hPalOld, FALSE);

	return S_OK;
}

//
// BitBltGray
//

STDMETHODIMP ImageSizeEntry::BitBltGray (HDC hDC, int nIndex, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	int nBitmapIndex = BitmapIndex(nIndex, eImage);
	int nRelativeIndex = RelativeIndex(nIndex, eImage);
	HBITMAP hImage = (HBITMAP)m_faImageBitmaps.GetAt(nBitmapIndex);

	HPALETTE hPal;
	GetPalette(hPal); 

	HDC hDCMem = GetGlobals().GetMemDC();
	if (NULL == m_hBitmapGray)
	{
		HDIB hDib = BitmapToDIB(hImage, hPal, 8);
		m_hBitmapGray = DIBToGrayBitmap(hDib);
		GlobalFree(hDib);
	}
	HBITMAP hBitmapOld = SelectBitmap(hDCMem, m_hBitmapGray);
	if (bScale)
	{
		::StretchBlt(hDC, 
					 x, 
					 y,
					 nHeight,
					 nWidth,
					 hDCMem,
					 nRelativeIndex*m_sizeImage.cx,
					 0,
					 m_sizeImage.cx,
					 m_sizeImage.cy,
					 dwRop);
	}
	else
	{
		::BitBlt(hDC, 
				 x, 
				 y,
				 m_sizeImage.cx,
				 m_sizeImage.cy,
				 hDCMem,
				 nRelativeIndex*m_sizeImage.cx,
				 0,
				 dwRop);
	}
	SelectBitmap(hDCMem, hBitmapOld);
	return S_OK;
}

//
// BitBltMask
//

STDMETHODIMP ImageSizeEntry::BitBltMask (HDC hDC, BOOL bMask, int nIndex, int x, int y, DWORD dwRop, OLE_COLOR ocColor, BOOL bScale, int nHeight, int nWidth)
{
	HBITMAP hMask;
	int nRelativeIndex;
	if (!bMask)
	{
		COLORREF crColor = -1;
		if (-1 != ocColor)
			OleTranslateColor(ocColor, NULL, &crColor);
		CreateMaskBitmap(hDC, 
					     (HBITMAP)m_faImageBitmaps.GetAt(BitmapIndex(nIndex, eImage)), 
						 RelativeIndex(nIndex, eImage), 
						 m_hMask,
						 crColor);
		hMask = m_hMask;
	}
	else
	{
		nRelativeIndex = RelativeIndex(nIndex, eMask);
		hMask = (HBITMAP)m_faMaskBitmaps.GetAt(BitmapIndex(nIndex, eMask));
	}

	HPALETTE hPal;
	GetPalette(hPal); 

	HDC hDCMem = GetGlobals().GetMemDC();
	HBITMAP hBitmapOld = SelectBitmap(hDCMem, hMask);
	if (bScale)
	{
		::StretchBlt(hDC, 
					 x, 
					 y,
					 nHeight,
					 nWidth,
					 hDCMem,
					 !bMask ? 0 : nRelativeIndex*m_sizeImage.cx,
					 0,
					 m_sizeImage.cx,
					 m_sizeImage.cy,
					 dwRop);
	}
	else
	{
		::BitBlt(hDC, 
				 x, 
				 y,
				 m_sizeImage.cx,
				 m_sizeImage.cy,
				 hDCMem,
				 !bMask ? 0 : nRelativeIndex*m_sizeImage.cx,
				 0,
				 dwRop);
	}

	SelectBitmap(hDCMem, hBitmapOld);
	return S_OK;
}

//
// BitBltDisabled
//

STDMETHODIMP ImageSizeEntry::BitBltDisabled (HDC hDC, int nIndex, int nMaskIndex, int x, int y, DWORD dwRop, OLE_COLOR ocColor, BOOL bScale, int nHeight, int nWidth)
{
	HDC hDCScreen = GetDC(NULL);
	if (NULL == hDCScreen)
		return E_FAIL;

	int nBitmapIndex = BitmapIndex(nIndex, eImage);
	int nRelativeIndex = RelativeIndex(nIndex, eImage);
	HBITMAP hImage = (HBITMAP)m_faImageBitmaps.GetAt(nBitmapIndex);

	HPALETTE hPal;
	GetPalette(hPal); 

	HDC hDCMask = ::CreateCompatibleDC(hDCScreen);
	HBITMAP hMask;
	if (bScale)
		hMask = ::CreateBitmap(nHeight, nWidth, 1, 1, 0);
	else
		hMask = ::CreateBitmap(m_sizeImage.cx, m_sizeImage.cy, 1, 1, 0);

	HBITMAP hBitmapOld = SelectBitmap(hDCMask, hMask);

	::PatBlt(hDCMask, 0, 0, m_sizeImage.cx, m_sizeImage.cy, BLACKNESS);

	// Fix the white portion
	HDC hDCColorBase = GetGlobals().GetMemDC();
	::SetBkColor(hDCColorBase, RGB(255, 255, 255));
	::SetTextColor(hDCColorBase, RGB(0, 0, 0));
	BitBlt(hDCMask, nIndex, 0, 0, SRCPAINT);

	::SetBkColor(hDCColorBase, RGB(192, 192, 192));
	::SetTextColor(hDCColorBase, RGB(0, 0, 0));
	BitBlt(hDCMask, nIndex, 0, 0, SRCPAINT);

	::SetBkColor(hDCColorBase, RGB(255, 255, 255));
	::SetTextColor(hDCColorBase, RGB(0, 0, 0));

	if (-1 == nMaskIndex)
	{
		COLORREF crColor = -1;
		if (-1 != ocColor)
			OleTranslateColor(ocColor, NULL, &crColor);
		BitBltMask(hDCMask, FALSE, nIndex, 0, 0, SRCPAINT);
	}
	else
		BitBltMask(hDCMask, TRUE, nMaskIndex, 0, 0, SRCPAINT);

	// HighLight
	COLORREF crBackOld = SetBkColor(hDC, RGB(255,255,255));
	COLORREF crForeOld = SetTextColor(hDC, RGB(0,0,0));
	if (bScale)
		::StretchBlt(hDC, x+1, y+1, nHeight, nWidth, hDCMask, 0, 0, m_sizeImage.cx, m_sizeImage.cy, SRCAND);
	else
		::BitBlt(hDC, x+1, y+1, m_sizeImage.cx, m_sizeImage.cy, hDCMask, 0, 0, SRCAND);

	SetBkColor(hDC, RGB(0,0,0));
	SetTextColor(hDC, GetSysColor(COLOR_BTNHILIGHT));
	if (bScale)
		::StretchBlt(hDC, x+1, y+1, nHeight, nWidth, hDCMask, 0, 0, m_sizeImage.cx, m_sizeImage.cy, SRCPAINT);
	else
		::BitBlt(hDC, x+1, y+1, m_sizeImage.cx, m_sizeImage.cy, hDCMask, 0, 0, SRCPAINT);

	// Shadow
	SetBkColor(hDC, RGB(255,255,255));
	SetTextColor(hDC, RGB(0,0,0));
	if (bScale)
		::StretchBlt(hDC, x, y, nHeight, nWidth, hDCMask, 0, 0, m_sizeImage.cx, m_sizeImage.cy, SRCAND);
	else
		::BitBlt(hDC, x, y, m_sizeImage.cx, m_sizeImage.cy, hDCMask, 0, 0, SRCAND);

	SetBkColor(hDC, RGB(0,0,0));
	SetTextColor(hDC, GetSysColor(COLOR_BTNSHADOW));
	if (bScale)
		::StretchBlt(hDC, x, y, nHeight, nWidth, hDCMask, 0, 0, m_sizeImage.cx, m_sizeImage.cy, SRCPAINT);
	else
		::BitBlt(hDC, x, y, m_sizeImage.cx, m_sizeImage.cy, hDCMask, 0, 0, SRCPAINT);

	// Clean Up
	SetBkColor(hDC, crBackOld);
	SetTextColor(hDC, crForeOld);

	SelectBitmap(hDCMask, hBitmapOld);
	DeleteBitmap(hMask);

	DeleteDC(hDCMask);
	ReleaseDC(NULL, hDCScreen);
	return S_OK;
}

//
// RemoveImage
//

STDMETHODIMP ImageSizeEntry::RemoveImage(int nIndex)
{
	int nSize = m_faImageFreeList.GetSize();
	for (int nFreeIndex = 0; nFreeIndex < nSize; nFreeIndex++)
	{
		if (nIndex == (int)m_faImageFreeList.GetAt(nFreeIndex))
			return E_FAIL;
	}
	m_faImageFreeList.Add((LPVOID)nIndex);
	return S_OK;
}

//
// RemoveMask
//

STDMETHODIMP ImageSizeEntry::RemoveMask(int nIndex)
{
	int nSize = m_faMaskFreeList.GetSize();
	for (int nFreeIndex = 0; nFreeIndex < nSize; nFreeIndex++)
	{
		if (nIndex == (int)m_faMaskFreeList.GetAt(nFreeIndex))
			return E_FAIL;
	}
	m_faMaskFreeList.Add((LPVOID)nIndex);
	return S_OK;
}

//
// GetSize
//

STDMETHODIMP ImageSizeEntry::GetSize(SIZE& sizeImage)
{
	sizeImage = m_sizeImage;
	return S_OK;
}

//
// Exchange
//

STDMETHODIMP ImageSizeEntry::Exchange(IStream* pStream, BOOL bSave, DWORD dwVersion)
{
	ImageEntry* pImageEntry;
	HPALETTE hPal;
	HRESULT  hResult;
	HBITMAP  hBitmap = NULL;
	long     nKey;
	int		 nSize;
	int		 nIndex;
	int      nTemp;

	GetPalette(hPal);
	
	if (bSave)
	{
		short nDepth;
		GetColorDepth(nDepth);

		switch (dwVersion)
		{
		case 0xB0B7:
			{
				nSize = m_faImageBitmaps.GetSize();
				hResult = pStream->Write(&nSize, sizeof(nSize), 0);
				for (nIndex = 0; nIndex < nSize; nIndex++)
					StWriteHBITMAP(pStream, (HBITMAP)m_faImageBitmaps.GetAt(nIndex), hPal, m_theImageMgr.m_nDepth);

				nSize = m_faImageFreeList.GetSize();
				hResult = pStream->Write(&nSize, sizeof(nSize), 0);
				for (nIndex = 0; nIndex < nSize; nIndex++)
				{
					nTemp = (int)m_faImageFreeList.GetAt(nIndex);
					hResult = pStream->Write(&nTemp, sizeof(int), 0);
				}

				nSize = m_faMaskBitmaps.GetSize();
				hResult = pStream->Write(&nSize, sizeof(nSize), 0);
				for (nIndex = 0; nIndex < nSize; nIndex++)
					StWriteHBITMAP(pStream, (HBITMAP)m_faMaskBitmaps.GetAt(nIndex), hPal, 1);
				
				nSize = m_faMaskFreeList.GetSize();
				hResult = pStream->Write(&nSize, sizeof(nSize), 0);
				for (nIndex = 0; nIndex < nSize; nIndex++)
				{
					nTemp = (int)m_faMaskFreeList.GetAt(nIndex);
					hResult = pStream->Write(&nTemp, sizeof(int), 0);
				}

				nSize = m_mapImages.GetCount();
				hResult = pStream->Write(&nSize, sizeof(nSize), 0);
				POSITION posMap = m_mapImages.GetStartPosition();
				while (0 != posMap)
				{
					m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
					pImageEntry->Exchange(pStream, bSave, dwVersion);
				}

				hResult = pStream->Write(&m_sizeImage, sizeof(m_sizeImage), 0);
				hResult = pStream->Write(&m_nImageInitial, sizeof(m_nImageInitial), 0);
				hResult = pStream->Write(&m_nImageGrowBy, sizeof(m_nImageGrowBy), 0);
				hResult = pStream->Write(&m_nMaskInitial, sizeof(m_nMaskInitial), 0);
				hResult = pStream->Write(&m_nMaskGrowBy, sizeof(m_nMaskGrowBy), 0);
			}
			break;

		case 0xB0B8:
		case 0xB0C1:
			{
				POSITION posMap = m_mapImages.GetStartPosition();
				while (0 != posMap)
				{
					m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
					pImageEntry->Exchange(pStream, bSave, dwVersion);
				}
			}
			break;
		}
	}
	else
	{
		switch (dwVersion)
		{
		case 0xB0B7:
			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				StReadHBITMAP(pStream, &hBitmap, 0, FALSE, hPal);
				m_faImageBitmaps.Add(hBitmap);
			}
			
			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				hResult = pStream->Read(&nTemp, sizeof(nTemp), 0);
				m_faImageFreeList.Add((LPVOID)nTemp);
			}

			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				StReadHBITMAP(pStream, &hBitmap, 0, TRUE, hPal);
				m_faMaskBitmaps.Add(hBitmap);
			}
			
			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				hResult = pStream->Read(&nTemp, sizeof(nTemp), 0);
				m_faMaskFreeList.Add((LPVOID)nTemp);
			}

			if (m_ppArrayImageEntry)
			{
				delete [] m_ppArrayImageEntry;
				m_ppArrayImageEntry = NULL;
			}
			DWORD dwImageId;
			hResult = pStream->Read(&m_nImageCount, sizeof(nSize), 0);
			m_ppArrayImageEntry = new ImageEntry*[m_nImageCount];
			for (nIndex = 0; nIndex < m_nImageCount; nIndex++)
			{
				pImageEntry = new ImageEntry(this);
				if (NULL == pImageEntry)
					return E_FAIL;
				pImageEntry->Exchange(pStream, bSave, dwVersion);
				pImageEntry->GetImageId(dwImageId);
				m_mapImages.SetAt((LPVOID)dwImageId, (LPVOID)pImageEntry);
				m_theImageMgr.SetImageEntry(dwImageId, pImageEntry);
				m_ppArrayImageEntry[nIndex] = pImageEntry;
			}
			hResult = pStream->Read(&m_sizeImage, sizeof(m_sizeImage), 0);
			hResult = pStream->Read(&m_nImageInitial, sizeof(m_nImageInitial), 0);
			hResult = pStream->Read(&m_nImageGrowBy, sizeof(m_nImageGrowBy), 0);
			hResult = pStream->Read(&m_nMaskInitial, sizeof(m_nMaskInitial), 0);
			hResult = pStream->Read(&m_nMaskGrowBy, sizeof(m_nMaskGrowBy), 0);
			break;

		case 0xB0B8:
		case 0xB0C1:
			{
				for (nIndex = 0; nIndex < m_nImageCount; nIndex++)
					m_ppArrayImageEntry[nIndex]->Exchange(pStream, bSave, dwVersion);
			}
			break;
		}

	}
	return hResult;
}

//
// DumpArray
//

static void DumpArray(FArray& faIntegers)
{
	int nElements = faIntegers.GetSize();
	for (int nIndex = 0; nIndex < nElements; nIndex++)
		TRACE2(1, _T("Index: %i, Value: %i\n"), nIndex, faIntegers.GetAt(nIndex));
}

//
// FindImageEntry
//

static ImageEntry* FindImageEntry(FMap& mapImageEntries, int nIndex)
{
	ImageEntry* pImageEntry;
	long		nKey;
	POSITION    posMap = mapImageEntries.GetStartPosition();
	while (0 != posMap)
	{
		mapImageEntries.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
		if (pImageEntry->ImageIndex() == nIndex)
			return pImageEntry;
	}
	return 0;
}

//
// FindImageEntry
//

static ImageEntry* FindMaskEntry(FMap& mapImageEntries, int nIndex)
{
	ImageEntry* pImageEntry;
	long		nKey;
	POSITION    posMap = mapImageEntries.GetStartPosition();
	while (0 != posMap)
	{
		mapImageEntries.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
		if (pImageEntry->MaskIndex() == nIndex)
			return pImageEntry;
	}
	return 0;
}

//
// CompactBitmap
//

BOOL ImageSizeEntry::CompactBitmap(FArray&   faBitmaps, 
								   FArray&   faFreeList, 
								   int&      nInitialBitmapSize, 
								   int       nGrowBy, 
								   SIZE	     sizeImage, 
								   ImageType eType)
{
	int nNumOfBitmaps = faBitmaps.GetSize();
	int nNumOfAvailableSlots = faFreeList.GetSize();
	if ((0 == nNumOfBitmaps && 0 == nNumOfAvailableSlots) || 0 == nInitialBitmapSize)
	{
		nInitialBitmapSize = nGrowBy;
		return TRUE;
	}

	int nMaxNumOfImages = nInitialBitmapSize;
	for (int nImage = 1; nImage < nNumOfBitmaps; nImage++)
		nMaxNumOfImages += nGrowBy;

	int nNumOfImages = nMaxNumOfImages - nNumOfAvailableSlots;
	if (nNumOfImages <= 0)
	{
		nInitialBitmapSize = nGrowBy;
		for (nImage = 0; nImage < nNumOfBitmaps; nImage++)
			DeleteBitmap((HBITMAP)faBitmaps.GetAt(nImage));
		faBitmaps.RemoveAll();
		faFreeList.RemoveAll();
		return TRUE;
	}

	int nPercentEmpty = (nNumOfAvailableSlots * 100) / nMaxNumOfImages;
	if (nPercentEmpty < 30)
		return TRUE;

	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return FALSE;

	HBITMAP hBitmapNew = NULL;
	if (eMask == eType)
		hBitmapNew = CreateBitmap(nNumOfImages * sizeImage.cx, m_sizeImage.cy, 1, 1, 0);
	else
		hBitmapNew = CreateCompatibleBitmap(hDC, nNumOfImages * sizeImage.cx, sizeImage.cy);

	if (NULL == hBitmapNew)
	{
		ReleaseDC(NULL, hDC);
		return FALSE;
	}

	HDC hDCDest = CreateCompatibleDC(hDC);
	HBITMAP hBitmapNewOld = SelectBitmap(hDCDest, hBitmapNew);
	
	HPALETTE hPal;
	m_theImageMgr.get_Palette((OLE_HANDLE*)&hPal);
	HPALETTE hPalOld1, hPalOld2;
	if (hPal)
	{
		hPalOld2 = SelectPalette(hDCDest, hPal, FALSE);
		RealizePalette(hDCDest);
	}

	long nKey;
	ImageEntry* pImageEntry;
	HBITMAP hBitmapSrc;
	HBITMAP hBitmapSrcOld;
	HDC hDCSrc = CreateCompatibleDC(hDC);
	int nBitmapIndex;
	int nRelativeIndex;
	int nNewImageIndex = 0;
	POSITION posMap = m_mapImages.GetStartPosition();
	while (0 != posMap)
	{
		m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
		
		switch (eType)
		{
		case eMask:
			nImage = pImageEntry->MaskIndex();
			break;

		case eImage:
			nImage = pImageEntry->ImageIndex();
			break;
		}

		if (-1 != nImage)
		{
			nBitmapIndex = BitmapIndex(nImage, eType);
			hBitmapSrc = (HBITMAP)faBitmaps.GetAt(nBitmapIndex);
			hBitmapSrcOld = SelectBitmap(hDCSrc, hBitmapSrc);
			
			nRelativeIndex = RelativeIndex(nImage, eType);

			if (hPal)
			{
				hPalOld1 = SelectPalette(hDCSrc, hPal, FALSE);
				RealizePalette(hDCSrc);
			}

			::BitBlt(hDCDest, 
					 nNewImageIndex * sizeImage.cx, 
					 0, 
					 sizeImage.cx, 
					 sizeImage.cy, 
					 hDCSrc, 
					 nRelativeIndex * m_sizeImage.cx, 
					 0,
					 SRCCOPY);

			if (hPal)
				SelectPalette(hDCSrc, hPalOld1, FALSE);

			SelectBitmap(hDCSrc, hBitmapSrcOld);

			switch (eType)
			{
			case eMask:
				pImageEntry->MaskIndex(nNewImageIndex);
				break;

			case eImage:
				pImageEntry->ImageIndex(nNewImageIndex);
				break;
			}
			nNewImageIndex++;
		}
	}
	nInitialBitmapSize = nNumOfImages;
	SelectBitmap(hDCDest, hBitmapNewOld);

	if (hPal)
		SelectPalette(hDCDest, hPalOld2, FALSE);

	DeleteDC(hDCSrc);
	DeleteDC(hDCDest);

	ReleaseDC(NULL, hDC);

	for (nImage = 0; nImage < nNumOfBitmaps; nImage++)
		DeleteBitmap((HBITMAP)faBitmaps.GetAt(nImage));

	faBitmaps.RemoveAll();
	faFreeList.RemoveAll();
	faBitmaps.Add(hBitmapNew);
	return TRUE;
}

//
// Compact
//

STDMETHODIMP ImageSizeEntry::Compact()
{
	CompactBitmap(m_faImageBitmaps, 
				  m_faImageFreeList, 
				  m_nImageInitial, 
				  m_nImageGrowBy, 
				  m_sizeImage, 
				  eImage);
	CompactBitmap(m_faMaskBitmaps, 
			      m_faMaskFreeList, 
				  m_nMaskInitial, 
				  m_nMaskGrowBy, 
				  m_sizeImage, 
				  eMask);
	if (m_hBitmapGray)
	{
		DeleteBitmap(m_hBitmapGray);
		m_hBitmapGray = NULL;
	}
	return S_OK;
}

//
// StReadPalette
//

HRESULT StReadPalette(IStream* pStream, HPALETTE& hPal)
{
	WORD nColors;
	pStream->Read(&nColors, sizeof(WORD), 0);
	if (nColors != 0)
	{
		LOGPALETTE* pPalette = new LOGPALETTE[nColors-1];
		if (pPalette)
		{
			pStream->Read(pPalette->palPalEntry, sizeof(PALETTEENTRY)*nColors, 0);
			pPalette->palVersion = 0x300; 
			pPalette->palNumEntries = nColors;
			hPal = CreatePalette(pPalette);
			delete [] pPalette;
		}
		else
			E_FAIL;
	}
	else
	{
#ifdef _DEBUG
		assert(NULL == hPal);
#endif
		return E_FAIL;
	}
	return NOERROR;
}
/*
//
// StReadPalette
//

HRESULT StReadPalette(IStream* pStream, HPALETTE& hPal)
{
	WORD nColors;
	HRESULT hResult = pStream->Read(&nColors, sizeof(WORD), 0);
	if (FAILED(hResult))
		return hResult;

	if (nColors > 0)
	{
		LOGPALETTE* pPalette = (LOGPALETTE*)malloc(sizeof(LOGPALETTE) + sizeof(PALETTEENTRY) * (nColors - 1));
		if (pPalette)
		{
			hResult = pStream->Read(pPalette->palPalEntry, sizeof(PALETTEENTRY) * nColors, NULL);
			if (FAILED(hResult))
				return hResult;

			pPalette->palVersion = 0x300; 
			pPalette->palNumEntries = nColors;

			hPal = CreatePalette(pPalette);
			free (pPalette);
			if (NULL == hPal)
				return E_FAIL;
		}
		else
			return E_FAIL;
	}
	else
	{
#ifdef _DEBUG
		assert(NULL == hPal);
#endif
	}
	return NOERROR;
}
*/
//
// StWritePalette
//

HRESULT StWritePalette(IStream* pStream, HPALETTE hPal)
{
	if (hPal == NULL)
		goto recover;

	WORD nColors;
	GetObject(hPal, sizeof(WORD), &nColors);
	if (nColors != 0)
	{
		PALETTEENTRY* pPalEntries = new PALETTEENTRY[nColors];
		if (pPalEntries)
		{
			int nResult = GetPaletteEntries(hPal, 0, nColors, pPalEntries);
			pStream->Write(&nColors, sizeof(WORD), 0);
			pStream->Write(pPalEntries, sizeof(PALETTEENTRY)*nColors, 0);
			delete [] pPalEntries;
			return NOERROR;
		}
		else
			return E_FAIL;
	}
recover:
	nColors = 0;
	pStream->Write(&nColors, sizeof(WORD), 0);
	return NOERROR;
}

//
//
// StWriteHBITMAP
//
// These function need to return error codes
//

HRESULT StWriteHBITMAP(IStream* pStream, HBITMAP hBitmap, HPALETTE hPal, int nBitsPerPixel)
{
	ULONG nSize = 0;
	if (NULL == hBitmap)
	{
		pStream->Write(&nSize, sizeof(ULONG), 0);
		return E_FAIL;
	}

	HDIB hDib = BitmapToDIB(hBitmap, hPal, nBitsPerPixel);
	if (NULL == hDib)
	{
		pStream->Write(&nSize, sizeof(ULONG), 0);
		return E_FAIL;
	}

	if (!StSaveDIB(pStream, hDib))
	{
		pStream->Write(&nSize, sizeof(ULONG), 0);
		return E_FAIL;
	}

	GlobalFree(hDib);
	return NOERROR;
}

//
// StReadHBITMAP
//

HRESULT StReadHBITMAP(IStream* pStream, HBITMAP* phBitmap, HBITMAP* phGrayBmp, BOOL bMonoChrome, HPALETTE hPal)
{
	ULONG nSize = 0;
	pStream->Read(&nSize, sizeof(ULONG), 0);
	if (0 == nSize)
	{
		*phBitmap = NULL;
		return E_FAIL;
	}

	HDIB hDib;
	StReadDIB(pStream, nSize, &hDib);
	if (NULL == hDib)
	{
		*phBitmap = NULL;
		return E_FAIL;
	}

	*phBitmap = DIBToBitmap(hDib, hPal, bMonoChrome);
	if (NULL != phGrayBmp)
		*phGrayBmp = DIBToGrayBitmap(hDib);

	GlobalFree(hDib);
	return NOERROR;
}

