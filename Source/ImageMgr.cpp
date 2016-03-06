//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Dib.h"
#include "Debug.h"
#include "Support.h"
#include "IpServer.h"
#include "Resource.h"
#include "Bar.h"
#include "Tool.h"
#include "ImageMgr.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
static void DebugPrintBitmap(HBITMAP hBitmap, int nOffset = 0)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return;

	HDC hDCSrc = CreateCompatibleDC(hDC);
	HBITMAP hBitmapOld = SelectBitmap(hDCSrc, hBitmap);
	BITMAP bmInfo;
	GetObject(hBitmap, sizeof(bmInfo), &bmInfo);
	BitBlt(hDC, 
		   0, 
		   nOffset, 
		   bmInfo.bmWidth, 
		   bmInfo.bmHeight, 
		   hDCSrc, 
		   0, 
		   0, 
		   SRCCOPY);
	SelectBitmap(hDCSrc, hBitmapOld);
	DeleteDC(hDCSrc);
	ReleaseDC(NULL, hDC);
}

#endif
extern HDC GetMemDC();

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CImageMgr(IUnknown *pUnkOuter)
{
	return (IUnknown *)new CImageMgr();
}

DEFINE_UNKNOWNOBJECT(&CLSID_ImageMgr,ImageMgr,_T("ImageMgr Class"),CreateFN_CImageMgr);
void *CImageMgr::objectDef=&ImageMgrObject;
CImageMgr *CImageMgr::CreateInstance(IUnknown *pUnkOuter)
{
	return new CImageMgr();
}
//{OBJECT CREATEFN}
CImageMgr::CImageMgr()
	: m_refCount(1)
	, m_hPal(NULL)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
	m_nImageLoad = 0;
}

CImageMgr::ImageManagerV1::ImageManagerV1()
{
	m_dwMasterImageId = 0;
}

CImageMgr::~CImageMgr()
{
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
// {END CLEANUP}
	CleanUp();
}

#ifdef _DEBUG
void CImageMgr::Dump(DumpContext& dc)
{
	ImageSizeEntry* pImageSizeEntry;
	ImageEntry* pImageEntry;
	long nKey;

	_stprintf(dc.m_szBuffer,
			  _T("Master Image Id: %i\n"), 
			  imV1.m_dwMasterImageId);
	dc.Write();

	_stprintf(dc.m_szBuffer,
			  _T("Image Manager - Image Size Entries - Count: %i\n"), 
			  m_mapImageSizes.GetCount());
	dc.Write();
	
	FPOSITION posMap = m_mapImageSizes.GetStartPosition();
	while (posMap)
	{
		m_mapImageSizes.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageSizeEntry);
		pImageSizeEntry->Dump(dc);
	}

	_stprintf(dc.m_szBuffer,
			  _T("Image Manager - Image Entries - Count: %i\n"), 
			  m_mapImages.GetCount());
	dc.Write();

	posMap = m_mapImages.GetStartPosition();
	while (0 != posMap)
	{
		m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
		pImageEntry->Dump(dc);
	}
}
#endif

//
// CleanUp
//

void CImageMgr::CleanUp()
{
	int nResult;
	long nKey;
	ImageSizeEntry* pImageSizeEntry;
	FPOSITION posMap = m_mapImageSizes.GetStartPosition();
	while (posMap)
	{
		m_mapImageSizes.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageSizeEntry);
		delete pImageSizeEntry;
	}
	m_mapImageSizes.RemoveAll();
	m_mapImages.RemoveAll();
	if (m_hPal)
	{
		nResult = DeletePalette(m_hPal);
		m_hPal= NULL;
		assert(nResult);
	}
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CImageMgr::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IUnknown))
	{
		AddRef();
		*ppvObjOut=this;
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IImageMgr))
	{
		*ppvObjOut=(void *)(IImageMgr *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
STDMETHODIMP_(ULONG) CImageMgr::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CImageMgr::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
STDMETHODIMP CImageMgr::get_MaskColor(long nImageId, OLE_COLOR *retval)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;
	
	pImageEntry->GetMaskColor(retval);
	return NOERROR;
}
STDMETHODIMP CImageMgr::put_MaskColor(long nImageId, OLE_COLOR val)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	pImageEntry->SetMaskColor(val);
	return NOERROR;
}
STDMETHODIMP CImageMgr::get_Palette(OLE_HANDLE *retval)
{
	*retval = (OLE_HANDLE)m_hPal;
	return NOERROR;
}
STDMETHODIMP CImageMgr::put_Palette(OLE_HANDLE val)
{
	int nResult;
	if (m_hPal)
	{
		nResult = DeletePalette(m_hPal);
		assert(nResult);
	}
	m_hPal = (HPALETTE)val;
	return NOERROR;
}
STDMETHODIMP CImageMgr::CreateImage(long* pnImageId, VARIANT_BOOL vbDesignerCreated)
{
	return CreateImageEx(pnImageId, eDefaultImageSize, eDefaultImageSize, vbDesignerCreated);
}
STDMETHODIMP CImageMgr::CreateImageEx(long* pnImageId,  int nCx,  int nCy, VARIANT_BOOL vbDesignerCreated)
{
	try
	{
		assert(pnImageId);
		if (NULL == pnImageId)
			return E_FAIL;

		*pnImageId = -1;
		SIZE sizeImage;
		sizeImage.cx = nCx;
		sizeImage.cy = nCy;
		long nKey = MAKELONG(sizeImage.cx, sizeImage.cy);

		ImageSizeEntry* pImageSizeEntry;
		if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		{
			pImageSizeEntry = new ImageSizeEntry(sizeImage, *this);
			if (NULL == pImageSizeEntry)
				return E_OUTOFMEMORY;

			m_mapImageSizes.SetAt((LPVOID)nKey, (LPVOID&)pImageSizeEntry);
		}

		++imV1.m_dwMasterImageId;
		ImageEntry* pImageEntry = new ImageEntry(pImageSizeEntry, imV1.m_dwMasterImageId, vbDesignerCreated);
		if (NULL == pImageEntry)
			return E_OUTOFMEMORY;

		m_mapImages.SetAt((LPVOID)imV1.m_dwMasterImageId, (LPVOID&)pImageEntry);
		*pnImageId = imV1.m_dwMasterImageId;
		return pImageSizeEntry->Add(pImageEntry);
	}
	CATCH
	{
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}
STDMETHODIMP CImageMgr::AddRefImage( long nImageId)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	pImageEntry->AddRef();
	return S_OK;
}
STDMETHODIMP CImageMgr::RefCntImage(long nImageId, long* pnRefCnt)
{
	assert(pnRefCnt);
	if (NULL == pnRefCnt)
		return E_FAIL;

	*pnRefCnt = 0;
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	*pnRefCnt = pImageEntry->RefCnt();
	return S_OK;
}
STDMETHODIMP CImageMgr::ReleaseImage(long* pnImageId)
{
	assert(pnImageId);
	if (NULL == pnImageId)
		return E_FAIL;

	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)*pnImageId, (LPVOID&)pImageEntry))
	{
		*pnImageId = -1;
		return E_FAIL;
	}

	if (0 == pImageEntry->Release())
	{	
		if (!m_mapImages.RemoveKey((LPVOID)*pnImageId))
		{
			*pnImageId = -1;
			return E_FAIL;
		}
	}
	*pnImageId = -1;
	return S_OK;
}
STDMETHODIMP CImageMgr::Size(long nImageId, long* nCx,  long* nCy)
{
	assert(nCx);
	assert(nCy);

	if (NULL == nCx || NULL == nCy)
		return E_FAIL;

	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;
	
	SIZE size;
	HRESULT hResult = pImageEntry->GetSize(size);
	*nCx = size.cx;
	*nCy = size.cy;
	return hResult;
}
STDMETHODIMP CImageMgr::PutImageBitmap( long nImageId, VARIANT_BOOL vbDesignerUpdated, OLE_HANDLE hBitmap)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	return pImageEntry->SetImageBitmap((HBITMAP)hBitmap, vbDesignerUpdated);
}
STDMETHODIMP CImageMgr::GetImageBitmap( long nImageId, VARIANT_BOOL vbUseMask, OLE_HANDLE* phBitmap)
{
	assert(phBitmap);
	if (NULL == phBitmap)
		return E_FAIL;

	*phBitmap = NULL;
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	return pImageEntry->GetImageBitmap((HBITMAP&)*phBitmap, vbUseMask);
}
STDMETHODIMP CImageMgr::PutMaskBitmap(long nImageId, VARIANT_BOOL vbDesignerUpdated, OLE_HANDLE hBitmap)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	return pImageEntry->SetMaskBitmap((HBITMAP)hBitmap, vbDesignerUpdated);
}

//
// get_MaskBitmap
//

STDMETHODIMP CImageMgr::GetMaskBitmap( long nImageId,  OLE_HANDLE* phBitmap)
{
	assert(phBitmap);
	if (NULL == phBitmap)
		return E_FAIL;

	*phBitmap = NULL;
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	return pImageEntry->GetMaskBitmap((HBITMAP&)*phBitmap);
}

//
// BitBltEx
//

STDMETHODIMP CImageMgr::BitBltEx(OLE_HANDLE hDC,  long nImageId,  long nX,  long nY, long nRop, ImageStyles isStyle)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	assert(pImageEntry);
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

//
// ScaleBlt
//

STDMETHODIMP CImageMgr::ScaleBlt( OLE_HANDLE hDC,  long nImageId,  long nX,  long nY,  long nWidth,  long nHeight,  long nRop, ImageStyles isStyle)
{
	ImageEntry* pImageEntry;
	if (!m_mapImages.Lookup((LPVOID)nImageId, (LPVOID&)pImageEntry))
		return E_FAIL;

	assert(pImageEntry);
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

//
// Exchange
//

HRESULT CImageMgr::Exchange(CBar* pBar, IStream* pStream, VARIANT_BOOL vbSave)
{
	long nStreamSize;
	short nSize, nSize2;
	ImageSizeEntry* pImageSizeEntry;
	HRESULT			hResult;
	long			nKey;
	int				nCount;
	int				nIndex;

	if (VARIANT_TRUE == vbSave)
	{
		try
		{
			//
			// Saving
			//

			BOOL bResult = Compact();

			nStreamSize = GetPaletteSize(m_hPal) + sizeof(nSize) + sizeof(imV1);
			hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWritePalette(pStream, m_hPal);
			if (FAILED(hResult))
				return hResult;

			nSize = sizeof(imV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&imV1, nSize, 0);
			if (FAILED(hResult))
				return hResult;

			nCount = m_mapImageSizes.GetCount();
			hResult = pStream->Write(&nCount, sizeof(nCount), 0);
			if (FAILED(hResult))
				return hResult;

			FPOSITION posMap = m_mapImageSizes.GetStartPosition();
			while (posMap)
			{
				m_mapImageSizes.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageSizeEntry);
				
				hResult = pStream->Write(&nKey, sizeof(nKey), 0);
				if (FAILED(hResult))
					return hResult;

				hResult = pImageSizeEntry->Exchange(pBar, pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		CATCH
		{
			REPORTEXCEPTION(__FILE__, __LINE__)
			return E_FAIL;
		}
	}
	else
	{
		try
		{
			//
			// Loading
			//

			m_nImageLoad++;

			CleanUp();

			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = StReadPalette(pStream, m_hPal);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= GetPaletteSize(m_hPal);
			if (nStreamSize <= 0)
				goto FinishedReading;

			hResult = pStream->Read(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= sizeof(nSize);
			if (nStreamSize <= 0)
				goto FinishedReading;

			nSize2 = sizeof(imV1);
			hResult = pStream->Read(&imV1, nSize < nSize2 ? nSize : nSize2, 0);
			if (FAILED(hResult))
				return hResult;

			if (nSize2 < nSize)
			{
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}

			nStreamSize -= nSize;
			if (nStreamSize <= 0)
				goto FinishedReading;

			nOffset.HighPart = 0;
			nOffset.LowPart = nStreamSize;
			hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
			if (FAILED(hResult))
				return hResult;

FinishedReading:
			SIZE sizeImage;
			hResult = pStream->Read(&nCount, sizeof(nCount), 0);
			for (nIndex = 0; nIndex < nCount; nIndex++)
			{
				hResult = pStream->Read(&nKey, sizeof(nKey), 0);
				if (FAILED(hResult))
					return hResult;

				sizeImage.cx = LOWORD(nKey);
				sizeImage.cy = HIWORD(nKey);
				
				pImageSizeEntry = new ImageSizeEntry(sizeImage, *this);
				if (NULL == pImageSizeEntry)
					return E_OUTOFMEMORY;

				m_mapImageSizes.SetAt((LPVOID)nKey, (LPVOID&)pImageSizeEntry);
				hResult = pImageSizeEntry->Exchange(pBar, pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		CATCH
		{
			REPORTEXCEPTION(__FILE__, __LINE__)
			return E_FAIL;
		}
	}
	return hResult;
}

//
// PersistConfig
//

HRESULT CImageMgr::ExchangeConfig(CBar* pBar, IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		TypedArray<ImageEntry*> aNotDesignerUpdated;
		TypedArray<ImageEntry*> aNotDesignerCreated;
		ImageEntry*				pImageEntry;
		HRESULT					hResult;
		HBITMAP				    hImage;
		HBITMAP				    hMask;
		BOOL					bResult;
		long					nKey;
		CDib				    theDib;
		int						nCount;
		int						nEntry;
		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving
			//

			FPOSITION posMap = m_mapImages.GetStartPosition();
			while (posMap)
			{
				m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
				if (VARIANT_FALSE == pImageEntry->iepV1.m_vbDesignerCreated)
					hResult = aNotDesignerCreated.Add(pImageEntry);
				else if (VARIANT_FALSE == pImageEntry->iepV1.m_vbDesignerUpdated)
					hResult = aNotDesignerUpdated.Add(pImageEntry);
			}
			
			//
			// Save non Designer created entries
			//

			HDC hDC = GetDC(NULL);
			if (NULL == hDC)
				return E_FAIL;

			nCount = aNotDesignerCreated.GetSize();
			
			hResult = pStream->Write(&nCount, sizeof(nCount), 0);
			if (FAILED(hResult))
				goto CleanUp;
			
			for (nEntry = 0; nEntry < nCount; nEntry++)
			{
				pImageEntry = aNotDesignerCreated.GetAt(nEntry);

				hResult = pImageEntry->GetImageBitmap(hImage, VARIANT_FALSE);
				if (FAILED(hResult))
					goto CleanUp;

				bResult = theDib.FromBitmap(hImage, hDC);
				assert(bResult);
				
				bResult = DeleteBitmap(hImage);
				assert(bResult);
				hImage = NULL;

				hResult = theDib.Write(pStream);
				if (FAILED(hResult))
					goto CleanUp;

				theDib.Empty();

				hResult = pImageEntry->GetMaskBitmap(hMask);
				if (FAILED(hResult))
					goto CleanUp;

				bResult = theDib.FromBitmap(hMask, hDC);
				assert(bResult);
				
				bResult = DeleteBitmap(hMask);
				assert(bResult);
				hMask = NULL;

				hResult = theDib.Write(pStream);
				if (FAILED(hResult))
					goto CleanUp;

				theDib.Empty();

				hResult = pImageEntry->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					goto CleanUp;
			}

			//
			// Save non Designer updated entries
			//

			nCount = aNotDesignerUpdated.GetSize();
			
			hResult = pStream->Write(&nCount, sizeof(nCount), 0);
			if (FAILED(hResult))
				goto CleanUp;

			for (nEntry = 0; nEntry < nCount; nEntry++)
			{
				pImageEntry = aNotDesignerUpdated.GetAt(nEntry);

				hResult = pStream->Write(&pImageEntry->iepV1.m_dwImageId, sizeof(pImageEntry->iepV1.m_dwImageId), 0);
				if (FAILED(hResult))
					goto CleanUp;

				hResult = pImageEntry->GetImageBitmap(hImage, VARIANT_FALSE);
				if (FAILED(hResult))
					goto CleanUp;

				bResult = theDib.FromBitmap(hImage, hDC);
				assert(bResult);

				hResult = theDib.Write(pStream);
				if (FAILED(hResult))
					goto CleanUp;

				theDib.Empty();

				bResult = DeleteBitmap(hImage);
				assert(bResult);
				hImage = NULL;

				hResult = pImageEntry->GetMaskBitmap(hMask);
				assert(bResult);

				bResult = theDib.FromBitmap(hMask, hDC);
				assert(bResult);

				bResult = DeleteBitmap(hMask);
				assert(bResult);
				hMask = NULL;

				hResult = theDib.Write(pStream);
				if (FAILED(hResult))
					goto CleanUp;

				theDib.Empty();
			}
CleanUp:
			ReleaseDC(NULL, hDC);
			return hResult;
		}
		else
		{
			//
			// Loading
			//

			//
			// Non designer created images
			//

			hResult = pStream->Read(&nCount, sizeof(nCount), 0);
			if (FAILED(hResult))
				return hResult;

			HDC hDC = GetDC(NULL);
			if (NULL == hDC)
				return E_FAIL;

			BOOL bResult;
			BITMAP bmInfo;
			for (nEntry = 0; nEntry < nCount; nEntry++)
			{
				theDib.Read(pStream);
				hImage = theDib.CreateBitmap(hDC);
				if (NULL == hImage)
				{
					hResult = E_FAIL;
					goto CleanUp2;
				}

				if (8 == GetGlobals().m_nBitDepth)
					pBar->m_pColorQuantizer->ProcessImage(theDib);
				
				theDib.Empty();

				GetObject(hImage, sizeof(BITMAP), &bmInfo);

				theDib.Read(pStream);
				hMask = theDib.CreateMonochromeBitmap(hDC);
				theDib.Empty();

				SIZE sizeImage;
				sizeImage.cx = bmInfo.bmWidth;
				sizeImage.cy = bmInfo.bmHeight;
				long nKey = MAKELONG(sizeImage.cx, sizeImage.cy);

				ImageSizeEntry* pImageSizeEntry;
				if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
				{
					pImageSizeEntry = new ImageSizeEntry(sizeImage, *this);
					if (NULL == pImageSizeEntry)
					{
						hResult = E_OUTOFMEMORY;
						goto CleanUp2;
					}

					m_mapImageSizes.SetAt((LPVOID)nKey, (LPVOID&)pImageSizeEntry);
				}

				ImageEntry* pImageEntry = new ImageEntry(pImageSizeEntry);
				if (NULL == pImageEntry)
				{
					hResult = E_OUTOFMEMORY;
					goto CleanUp2;
				}

				hResult = pImageEntry->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					goto CleanUp2;

				m_mapImages.SetAt((LPVOID)pImageEntry->iepV1.m_dwImageId, (LPVOID&)pImageEntry);
				hResult = pImageSizeEntry->Add(pImageEntry);

				// The tools will add the reference count
				pImageEntry->iepV1.m_nRefCnt = 0;

				if (pImageEntry->iepV1.m_dwImageId > imV1.m_dwMasterImageId)
					imV1.m_dwMasterImageId = pImageEntry->iepV1.m_dwImageId;

				pImageEntry->iepV1.m_nImageIndex = -1;
				pImageEntry->iepV1.m_nMaskIndex = -1;
				
				pImageEntry->SetImageBitmap(hImage, VARIANT_FALSE);
				bResult = DeleteBitmap(hImage);
				assert(bResult);
				hImage = NULL;

				if (-1 == pImageEntry->iepV1.m_ocMaskColor)
					pImageEntry->SetMaskBitmap(hMask, VARIANT_FALSE);
				bResult = DeleteBitmap(hMask);
				assert(bResult);
				hMask = NULL;
			}

			//
			// Non designer updated images
			//

			hResult = pStream->Read(&nCount, sizeof(nCount), 0);
			if (FAILED(hResult))
				goto CleanUp2;

			DWORD dwImageId;
			for (nEntry = 0; nEntry < nCount; nEntry++)
			{
				hResult = pStream->Read(&dwImageId, sizeof(dwImageId), 0);
				if (FAILED(hResult))
					goto CleanUp2;

				if (!m_mapImages.Lookup((LPVOID)dwImageId, (LPVOID&)pImageEntry))
				{
					hResult = E_FAIL;
					goto CleanUp2;
				}

				theDib.Read(pStream);
				hImage = theDib.CreateBitmap(hDC);
				pImageEntry->SetImageBitmap(hImage, VARIANT_FALSE);

				bResult = DeleteBitmap(hImage);
				assert(bResult);
				theDib.Empty();

				theDib.Read(pStream);
				hMask = theDib.CreateMonochromeBitmap(hDC);
				pImageEntry->SetMaskBitmap(hMask, VARIANT_FALSE);

				bResult = DeleteBitmap(hMask);
				assert(bResult);
				theDib.Empty();
			}
CleanUp2:
			ReleaseDC(NULL, hDC);
			return hResult;
		}
	}
	CATCH
	{
		assert(TRUE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

//
// Compact
//

STDMETHODIMP CImageMgr::Compact()
{
	ImageSizeEntry* pImageSizeEntry;
	HRESULT			hResult = S_OK;
	long			nKey;

	FPOSITION posMap = m_mapImageSizes.GetStartPosition();
	while (posMap)
	{
		m_mapImageSizes.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageSizeEntry);
		assert(pImageSizeEntry);
		if (pImageSizeEntry)
		{
			hResult = pImageSizeEntry->Compact();
			if (FAILED(hResult))
				return hResult;
		}
	}
	return hResult;
}

//
// BitmapInitialSize
//

STDMETHODIMP CImageMgr::BitmapInitialSize( short nCx,  short nCy,  short nSize)
{
	ImageSizeEntry* pImageSizeEntry;
	long nKey = MAKELONG(nCx, nCy);
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		return E_FAIL;
	
	assert(pImageSizeEntry);
	pImageSizeEntry->InitialSize(nSize, ImageSizeEntry::eImage);
	return S_OK;
}

//
// BitmapGrowBy
//

STDMETHODIMP CImageMgr::BitmapGrowBy( short nCx,  short nCy,  short nGrowBy)
{
	ImageSizeEntry* pImageSizeEntry;
	long nKey = MAKELONG(nCx, nCy);
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		return E_FAIL;
	
	assert(pImageSizeEntry);
	pImageSizeEntry->GrowBy(nGrowBy, ImageSizeEntry::eImage);
	return S_OK;
}

//
// MaskInitialSize
//

STDMETHODIMP CImageMgr::MaskBitmapInitialSize( short nCx,  short nCy,  short nSize)
{
	ImageSizeEntry* pImageSizeEntry;
	long nKey = MAKELONG(nCx, nCy);
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		return E_FAIL;
	
	pImageSizeEntry->InitialSize(nSize, ImageSizeEntry::eMask);
	return S_OK;
}

//
// MaskGrowBy
//

STDMETHODIMP CImageMgr::MaskBitmapGrowBy( short nCx,  short nCy,  short nGrowBy)
{
	ImageSizeEntry* pImageSizeEntry;
	long nKey = MAKELONG(nCx, nCy);
	if (!m_mapImageSizes.Lookup((LPVOID)nKey, (LPVOID&)pImageSizeEntry))
		return E_FAIL;
	
	assert(pImageSizeEntry);
	pImageSizeEntry->GrowBy(nGrowBy, ImageSizeEntry::eMask);
	return S_OK;
}

//
// ImageEntry
//

ImageEntry::ImageEntry(ImageSizeEntry* pImageSizeEntry, DWORD dwImageId, VARIANT_BOOL vbDesignerCreated)
	: m_pImageSizeEntry(pImageSizeEntry)
{
	assert(m_pImageSizeEntry);
    iepV1.m_dwImageId = dwImageId;
	iepV1.m_vbDesignerCreated = vbDesignerCreated;
}

ImageEntry::ImageEntryV1::ImageEntryV1()
{
	m_ocMaskColor = -1;
	m_nImageIndex = -1;
	m_nMaskIndex = -1;
	m_nRefCnt = 1;
	m_vbDesignerCreated = VARIANT_FALSE;
	m_vbDesignerUpdated = VARIANT_FALSE;
}

//
// ~ImageEntry
//

ImageEntry::~ImageEntry()
{
	ReleaseImage();
	ReleaseMask();
	assert(m_pImageSizeEntry);
	m_pImageSizeEntry->Remove(this);
	assert(!m_pImageSizeEntry->GetImageMgr().FindImageId(iepV1.m_nImageIndex));
}

ULONG ImageEntry::Release()
{
	if (0 != --iepV1.m_nRefCnt)
		return iepV1.m_nRefCnt;
	delete this;
	return 0;
}

//
// Special Debugging Functions
//

#ifdef _DEBUG
void ImageEntry::Dump(DumpContext& dc)
{
	_stprintf(dc.m_szBuffer,
			  _T("m_dwImageId: %li m_RefCnt: %li\n"), 
			  iepV1.m_dwImageId,
			  iepV1.m_nRefCnt);
	dc.Write();
	_stprintf(dc.m_szBuffer,
			  _T("Image Index: %li Mask Index: %li\n"), 
			  iepV1.m_nImageIndex,
			  iepV1.m_nMaskIndex);
	dc.Write();
}
#endif

//
// GetSize
//

HRESULT ImageEntry::GetSize(SIZE& sizeImage)
{
	if (NULL == m_pImageSizeEntry)
	{
		sizeImage.cx = sizeImage.cy = -1;
		return E_FAIL;
	}
	m_pImageSizeEntry->GetSize(sizeImage);
	return S_OK;
}

//
// GetImageBitmap
//

HRESULT ImageEntry::GetImageBitmap(HBITMAP& hBitmapImage, VARIANT_BOOL vbUseMask)
{
	hBitmapImage = NULL;
	assert(m_pImageSizeEntry);
	return m_pImageSizeEntry->GetImageBitmap(iepV1.m_nImageIndex, iepV1.m_nMaskIndex, hBitmapImage, vbUseMask);
}

//
// SetImageBitmap
//

HRESULT ImageEntry::SetImageBitmap(HBITMAP hBitmapImage, VARIANT_BOOL vbDesignerUpdated)
{
	if (NULL == hBitmapImage)
	{
		ReleaseImage();
		return NOERROR;
	}
	assert(m_pImageSizeEntry);
	iepV1.m_vbDesignerUpdated = vbDesignerUpdated;
	ReleaseMask();
	return m_pImageSizeEntry->SetImageBitmap(iepV1.m_nImageIndex, hBitmapImage);
}

//
// GetMaskBitmap
//

HRESULT ImageEntry::GetMaskBitmap(HBITMAP& hBitmapMask)
{
	hBitmapMask = NULL;
	assert(m_pImageSizeEntry);
	if (-1 == iepV1.m_nMaskIndex)
		return m_pImageSizeEntry->GetMaskBitmap(iepV1.m_nImageIndex, FALSE, hBitmapMask, iepV1.m_ocMaskColor);
	else
		return m_pImageSizeEntry->GetMaskBitmap(iepV1.m_nMaskIndex, TRUE, hBitmapMask);
}

//
// SetMaskBitmap
//

HRESULT ImageEntry::SetMaskBitmap(HBITMAP hBitmapMask, VARIANT_BOOL vbDesignerUpdated)
{
	if (NULL == hBitmapMask)
	{
		ReleaseMask();
		return NOERROR;
	}
	assert(m_pImageSizeEntry);
	iepV1.m_vbDesignerUpdated = vbDesignerUpdated;
	return m_pImageSizeEntry->SetMaskBitmap(iepV1.m_nMaskIndex, hBitmapMask);
}

//
// BitBlt
//

HRESULT ImageEntry::BitBlt(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	assert(m_pImageSizeEntry);
	return m_pImageSizeEntry->BitBlt(hDC, iepV1.m_nImageIndex, x, y, dwRop, bScale, nHeight, nWidth);
}

//
// BitBltGray
//

HRESULT ImageEntry::BitBltGray(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	assert(m_pImageSizeEntry);
	return m_pImageSizeEntry->BitBltGray(hDC, iepV1.m_nImageIndex, x, y, dwRop, bScale, nHeight, nWidth);
}

//
// BitBltMask
//

HRESULT ImageEntry::BitBltMask(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	HRESULT hResult;
	assert(m_pImageSizeEntry);
	if (-1 == iepV1.m_nMaskIndex)
		hResult = m_pImageSizeEntry->BitBltMask(hDC, FALSE, iepV1.m_nImageIndex, x, y, dwRop, iepV1.m_ocMaskColor, bScale, nHeight, nWidth);
	else
		hResult = m_pImageSizeEntry->BitBltMask(hDC, TRUE, iepV1.m_nMaskIndex, x, y, dwRop, iepV1.m_ocMaskColor, bScale, nHeight, nWidth);
	return hResult;
}

//
// BitBltDisabled
//

HRESULT ImageEntry::BitBltDisabled(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	return m_pImageSizeEntry->BitBltDisabled(hDC, iepV1.m_nImageIndex, iepV1.m_nMaskIndex, x, y, dwRop, iepV1.m_ocMaskColor, bScale, nHeight, nWidth);
}

//
// ReleaseImage
//

HRESULT ImageEntry::ReleaseImage()
{
	if (-1 == iepV1.m_nImageIndex)
		return S_OK;
	assert(m_pImageSizeEntry);
	HRESULT hResult = m_pImageSizeEntry->RemoveImage(iepV1.m_nImageIndex);
	iepV1.m_nImageIndex = -1;
	return hResult;
}

//
// ReleaseMask
//

HRESULT ImageEntry::ReleaseMask()
{
	if (-1 == iepV1.m_nMaskIndex)
		return S_OK;
	assert(m_pImageSizeEntry);
	HRESULT hResult = m_pImageSizeEntry->RemoveMask(iepV1.m_nMaskIndex);
	iepV1.m_nMaskIndex = -1;
	return hResult;
}

//
// Exchange
//

HRESULT ImageEntry::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		long nStreamSize;
		short nSize;
		short nSize2;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			nStreamSize = sizeof(nSize) + sizeof(iepV1);
			hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nSize = sizeof(iepV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&iepV1, sizeof(iepV1), 0);
			if (FAILED(hResult))
				return hResult;
		}
		else
		{
			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Read(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= sizeof(nSize);
			if (nStreamSize <= 0)
				goto FinishedReading;

			nSize2 = sizeof(iepV1);
			hResult = pStream->Read(&iepV1, nSize < nSize2 ? nSize : nSize2, 0);
			if (FAILED(hResult))
				return hResult;

			if (nSize2 < nSize)
			{
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}
			nStreamSize -= nSize;
			if (nStreamSize <= 0)
				goto FinishedReading;

			nOffset.HighPart = 0;
			nOffset.LowPart = nStreamSize;
			hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
			if (FAILED(hResult))
				return hResult;
		}
FinishedReading:
		return hResult;
	}
	CATCH
	{
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// ImageSizeEntry
//

ImageSizeEntry::ImageSizeEntry(const SIZE& sizeImage, CImageMgr& theImageMgr)
	: m_theImageMgr(theImageMgr),
	  m_hMask(NULL)
{
	iseV1.m_sizeImage = sizeImage;

	if (sizeImage.cx > CImageMgr::eDefaultImageSize)
	{
		iseV1.m_nImageInitial =  CImageMgr::eDefaultImageInitialElements * CImageMgr::eDefaultImageSize / sizeImage.cx;
		if (0 == iseV1.m_nImageInitial)
			iseV1.m_nImageInitial = 1;
		iseV1.m_nMaskInitial = iseV1.m_nImageInitial;
	}
	else
	{
		iseV1.m_nImageInitial = CImageMgr::eDefaultImageInitialElements;
		iseV1.m_nMaskInitial = CImageMgr::eDefaultMaskInitialElements;
	}
	iseV1.m_nImageGrowBy = CImageMgr::eDefaultImageGrowBy;
	iseV1.m_nMaskGrowBy = CImageMgr::eDefaultMaskGrowBy;
}

//
// ~ImageSizeEntry
//

ImageSizeEntry::~ImageSizeEntry()
{
	HBITMAP hTemp;
	int nIndex = 0;
	int nResult;
	long nKey;
	ImageEntry* pImageEntry;
	
	FPOSITION posMap = m_mapImages.GetStartPosition();
	while (posMap)
	{
		m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
		assert(pImageEntry);
		pImageEntry->Release();
	}

	int nSize = m_aImageBitmaps.GetSize();
	for (nIndex = 0; nIndex < nSize; nIndex++)
	{
		hTemp = m_aImageBitmaps.GetAt(nIndex);
		if (hTemp)
		{
			nResult = DeleteBitmap(hTemp);
			assert(0 != nResult);
		}
	}
	m_aImageBitmaps.RemoveAll();
	m_aImageFreeList.RemoveAll();

	nSize = m_aGrayBitmaps.GetSize();
	for (nIndex = 0; nIndex < nSize; nIndex++)
	{
		hTemp = m_aGrayBitmaps.GetAt(nIndex);
		if (hTemp)
		{
			nResult = DeleteBitmap(hTemp);
			assert(0 != nResult);
		}
	}
	m_aGrayBitmaps.RemoveAll();

	nSize = m_aMaskBitmaps.GetSize();
	for (nIndex = 0; nIndex < nSize; nIndex++)
	{
		hTemp = m_aMaskBitmaps.GetAt(nIndex);
		if (hTemp)
		{
			nResult = DeleteBitmap(hTemp);
			assert(0 != nResult);
		}
	}
	m_aMaskBitmaps.RemoveAll();
	m_aMaskFreeList.RemoveAll();

	if (m_hMask)
	{
		nResult = DeleteBitmap(m_hMask);
		m_hMask = NULL;
		assert(0 != nResult);
	}
	m_mapImages.RemoveAll();
}

ImageSizeEntry::ImageSizeEntryV1::ImageSizeEntryV1()
{
	m_sizeImage.cx = m_sizeImage.cy = 0;
	m_nImageInitial = m_nImageGrowBy = m_nMaskInitial = m_nMaskGrowBy = 0;
}

//
// Special Debugging Functions
//

#ifdef _DEBUG
void ImageSizeEntry::Dump(DumpContext& dc)
{
	int nIndex = 0; 
	_stprintf(dc.m_szBuffer,
			  _T("Image Size Entry - nImage size x: %i y: %i, m_nInitial: %i, m_nGrowBy: %i\n"), 
			  iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, iseV1.m_nImageInitial, iseV1.m_nImageGrowBy);
	dc.Write();

	dc.Write(_T("Image Free List\n\t"));
	int nCount = m_aImageFreeList.GetSize();
	for (nIndex = 0; nIndex < nCount; nIndex++)
	{
		_stprintf(dc.m_szBuffer,
				  _T("%i, "),
				  m_aImageFreeList.GetAt(nIndex));
		dc.Write();
	}

	dc.Write(_T("\nMask Free List\n\t"));
	nCount = m_aMaskFreeList.GetSize();
	for (nIndex = 0; nIndex < nCount; nIndex++)
	{
		_stprintf(dc.m_szBuffer,
				  _T("%i, "),
				  m_aMaskFreeList.GetAt(nIndex));
		dc.Write();
	}

	_stprintf(dc.m_szBuffer,
			  _T("\nImage Entries - Count %i\n"), 
			  m_mapImages.GetCount());
	dc.Write();

	ImageEntry* pImageEntry;
	long nKey;
	FPOSITION posMap = m_mapImages.GetStartPosition();
	while (posMap)
	{
		m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
		pImageEntry->Dump(dc);
	}
}
#endif

//
// GetPalette
//

void ImageSizeEntry::GetPalette(HPALETTE& hPal)
{
	m_theImageMgr.get_Palette((OLE_HANDLE*)&hPal);
}

//
// CreateMaskBitmap
//

BOOL ImageSizeEntry::CreateMaskBitmap(HDC hDC, HBITMAP hBitmap, int nRelativeIndex, HBITMAP& hBitmapMask, OLE_COLOR ocColor)
{
	BOOL bResult = TRUE;
	if (NULL == hBitmapMask)
	{
		//
		// If a mask bitmap is not being stored in the image manager we need to create one here so we can return it.
		//

		hBitmapMask = CreateBitmap(iseV1.m_sizeImage.cx, 
								   iseV1.m_sizeImage.cy,
								   1,
								   1,
								   0);
		if (NULL == hBitmapMask)
			return FALSE;
	}

    //
	// Create memory DCs to work with.
	//
	
	HBITMAP hbmMaskOld;
	HBITMAP hbmImageOld;
	HDC hDCMask = NULL;
	HDC hDCImage = NULL;

    hDCMask = ::CreateCompatibleDC(hDC);
	if (NULL == hDCMask)
	{
		bResult = FALSE;
		goto CleanUp;
	}

	hDCImage = ::CreateCompatibleDC(hDC);
	if (NULL == hDCImage)
	{
		bResult = FALSE;
		goto CleanUp;
	}

/*	// Select the mono bitmap into its DC.
	if (NULL == m_hMask)
	{
		m_hMask = ::CreateBitmap(iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, 1, 1, 0);
		if (NULL == m_hMask)
		{
			bResult = FALSE;
			goto CleanUp;
		}
	}
*/
	hbmMaskOld = SelectBitmap(hDCMask, hBitmapMask);

	// Select the image bitmap into its DC.
	hbmImageOld = SelectBitmap(hDCImage, hBitmap);

	if (-1 == ocColor)
	{
		//
		// Set the transparency color to be the top-left pixel.
		//
		
		COLORREF crColor = ::GetPixel(hDCImage, (nRelativeIndex*iseV1.m_sizeImage.cx), 0);
		int nHit = 1;
		if (crColor == GetPixel(hDCImage, ((nRelativeIndex+1)*iseV1.m_sizeImage.cx)-1, iseV1.m_sizeImage.cy-1)) 
			nHit++;
		if (crColor == GetPixel(hDCImage, (nRelativeIndex*iseV1.m_sizeImage.cx), iseV1.m_sizeImage.cy-1)) 
			nHit++;
		if (crColor == GetPixel(hDCImage, ((nRelativeIndex+1)*iseV1.m_sizeImage.cx)-1, iseV1.m_sizeImage.cy-1)) 
			nHit++;

		if (nHit > 1)
			::SetBkColor(hDCImage, crColor);
		else 
		{
			//
			// Default to a gray background
			//

			::SetBkColor(hDCImage, RGB(192, 192, 192)); 
		}
	}
	else
	{
		//
		// The mask color was set so use it
		//

		COLORREF crColor;
		OleTranslateColor(ocColor, NULL, &crColor);
		::SetBkColor(hDCImage, crColor);
	}

	SetTextColor(hDCImage, RGB(255,255,255));

	bResult = ::BitBlt(hDCMask, 
					   0,
					   0,
					   iseV1.m_sizeImage.cx,
					   iseV1.m_sizeImage.cy,
					   hDCImage,
					   nRelativeIndex * iseV1.m_sizeImage.cx,
					   0,
					   SRCCOPY);

	SelectBitmap(hDCMask, hbmMaskOld);
	SelectBitmap(hDCImage, hbmImageOld);

CleanUp:
	if (hDCImage)
		DeleteDC(hDCImage);

	if (hDCMask)
		DeleteDC(hDCMask);
	return bResult;
}

//
// Add
//

HRESULT ImageSizeEntry::Add(ImageEntry* pImageEntry)
{
	assert(pImageEntry);
	if (NULL == pImageEntry)
		return E_FAIL;
	DWORD dwImageId;
	pImageEntry->GetImageId(dwImageId);
	m_mapImages.SetAt((LPVOID)dwImageId, (LPVOID)pImageEntry);
	return S_OK;
}

//
// Remove
//

HRESULT ImageSizeEntry::Remove(ImageEntry* pImageEntry)
{
	assert(pImageEntry);
	if (NULL == pImageEntry)
		return E_FAIL;
	
	DWORD dwImageId;
	pImageEntry->GetImageId(dwImageId);
	
	m_mapImages.RemoveKey((LPVOID)dwImageId);
	return S_OK;
}

//
// GetImageBitmap
//

HRESULT ImageSizeEntry::GetImageBitmap(const int& nIndex, const int& nMaskIndex, HBITMAP& hBitmapImage, VARIANT_BOOL vbUseMask)
{
	if (nIndex < 0)
		return E_FAIL;

	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return E_FAIL;
	
	HPALETTE hPalDestOld, hPalImageOld;
	HPALETTE hPalMaskOld = NULL;
	HPALETTE hPal = NULL;
	HBITMAP hBitmapImageOld = NULL;
	HBITMAP hBitmapDestOld = NULL;
	HBITMAP hBitmapMaskOld = NULL;
	HBITMAP hImage = NULL;
	HBITMAP hMask = NULL;
	HDC hDCImage = NULL;
	HDC hDCDest = NULL;
	HDC hDCMask = NULL;
	int nMaskRelativeIndex = -1;
	int nMaskBitmapIndex = -1;
	int nRelativeIndex;
	int nBitmapIndex;
	int nXPosImage;
	int nXPosMask;

	GetPalette(hPal); 

	nBitmapIndex = BitmapIndex(nIndex, eImage);
	hImage = m_aImageBitmaps.GetAt(nBitmapIndex);
	if (NULL == hImage)
		goto CleanUp;

	hBitmapImage = CreateCompatibleBitmap(hDC, 
										  iseV1.m_sizeImage.cx, 
										  iseV1.m_sizeImage.cy);
	if (NULL == hBitmapImage)
		goto CleanUp;

	nRelativeIndex = RelativeIndex(nIndex, eImage);

	hDCDest = CreateCompatibleDC(hDC);
	if (NULL == hDCDest)
		goto CleanUp;

	hDCImage = GetMemDC();
	if (NULL == hDCImage)
		goto CleanUp;

	if (hPal)
	{
		hPalDestOld = SelectPalette(hDCDest, hPal, FALSE);
		RealizePalette(hDCDest);

		hPalImageOld = SelectPalette(hDCImage, hPal, FALSE);
		RealizePalette(hDCImage);
	}

	hBitmapDestOld = SelectBitmap(hDCDest, hBitmapImage);
	hBitmapImageOld = SelectBitmap(hDCImage, hImage);

	nXPosImage = nRelativeIndex * iseV1.m_sizeImage.cx;

	if (-1 != nMaskIndex && VARIANT_TRUE == vbUseMask)
	{
		nMaskBitmapIndex = BitmapIndex(nMaskIndex, eMask);
		nMaskRelativeIndex = RelativeIndex(nMaskIndex, eMask);
		hMask = m_aMaskBitmaps.GetAt(nMaskBitmapIndex);
	
		PatBlt(hDCDest, 
			   0, 
			   0, 
			   iseV1.m_sizeImage.cx, 
			   iseV1.m_sizeImage.cy, 
			   BLACKNESS);

		::SetTextColor(hDCDest, RGB(0,0,0));
		::SetBkColor(hDCDest, RGB(255,255,255));

		// Mask out the grey around the base image
		nXPosMask = nMaskRelativeIndex * iseV1.m_sizeImage.cx;
		hDCMask = CreateCompatibleDC(hDC);
		if (NULL == hDCMask)
			goto CleanUp;

		if (hPal)
		{
			hPalMaskOld = SelectPalette(hDCMask, hPal, FALSE);
			RealizePalette(hDCMask);
		}

		::BitBlt(hDCDest,
				 0,
				 0,
				 iseV1.m_sizeImage.cx,
				 iseV1.m_sizeImage.cy,
				 hDCImage,
				 nXPosImage,
				 0,
				 SRCINVERT);

		hBitmapMaskOld = SelectBitmap(hDCMask, hMask);

		::BitBlt(hDCDest,
				 0,
				 0,
				 iseV1.m_sizeImage.cx,
				 iseV1.m_sizeImage.cy,
				 hDCMask,
				 nXPosMask,
				 0,
				 SRCAND);

		SelectBitmap(hDCMask, hBitmapMaskOld);

		::BitBlt(hDCDest,
				 0,
				 0,
				 iseV1.m_sizeImage.cx,
				 iseV1.m_sizeImage.cy,
				 hDCImage,
				 nXPosImage,
				 0,
				 SRCINVERT);
		
		if (hPal)
			SelectPalette(hDCMask, hPalMaskOld, FALSE);

		DeleteDC(hDCMask);
	}
	else
	{
		::BitBlt(hDCDest, 
				 0, 
				 0, 
				 iseV1.m_sizeImage.cx, 
				 iseV1.m_sizeImage.cy,
				 hDCImage,
				 nXPosImage,
				 0,
				 SRCCOPY);
	}

	if (hPal)
	{
		SelectPalette(hDCDest, hPalDestOld, FALSE);
		SelectPalette(hDCImage, hPalImageOld, FALSE);
	}

CleanUp:
	if (hBitmapDestOld)
		SelectBitmap(hDCDest, hBitmapDestOld);

	if (hBitmapImageOld)
		SelectBitmap(hDCImage, hBitmapImageOld);

	if (hDCDest)
		DeleteDC(hDCDest);

	if (hDC)
		ReleaseDC(NULL, hDC);
	
	return S_OK;
}

//
// SetImageBitmap
//

HRESULT ImageSizeEntry::SetImageBitmap(int& nIndex, HBITMAP hBitmapImage)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return E_FAIL;

	HRESULT hResult;
	HBITMAP hImages;
	int nRelativeIndex = 0;
	if (-1 == nIndex)
	{
		if (0 == m_aImageBitmaps.GetSize())
		{
			//
			// Initial
			//

			int nWidth = iseV1.m_sizeImage.cx * iseV1.m_nImageInitial;
			hImages = CreateCompatibleBitmap(hDC, nWidth, iseV1.m_sizeImage.cy);
			hResult = m_aImageBitmaps.Add(hImages);
			nIndex = 0;
			
			for (int nIndex2 = 1; nIndex2 < iseV1.m_nImageInitial; nIndex2++)
				hResult = m_aImageFreeList.Add(nIndex2);

			hResult = m_aGrayBitmaps.Add(NULL);
		}
		else if (m_aImageFreeList.GetSize() > 0)
		{
			//
			// Are there free positions?
			//

			nIndex = m_aImageFreeList.GetAt(0);
			m_aImageFreeList.RemoveAt(0);
			hImages = m_aImageBitmaps.GetAt(BitmapIndex(nIndex, eImage));
			m_aGrayBitmaps.SetAt(BitmapIndex(nIndex, eImage), NULL);
		}
		else
		{
			//
			// No free positions so create an additional storage bitmap
			//

			int nSize = m_aImageBitmaps.GetSize();
			int nWidth = iseV1.m_sizeImage.cx * iseV1.m_nImageGrowBy;
			
			hImages = CreateCompatibleBitmap(hDC, nWidth, iseV1.m_sizeImage.cy);
			if (NULL == hImages)
				return E_FAIL;

			hResult = m_aImageBitmaps.Add(hImages);
			hResult = m_aGrayBitmaps.Add(NULL);
			nIndex = iseV1.m_nImageInitial + ((nSize-1) * iseV1.m_nImageGrowBy);
			for (int nIndex2 = nIndex+1; nIndex2 < iseV1.m_nImageGrowBy+nIndex; nIndex2++)
				hResult = m_aImageFreeList.Add(nIndex2);
		}
	}
	else
		hImages = m_aImageBitmaps.GetAt(BitmapIndex(nIndex, eImage));

	nRelativeIndex = RelativeIndex(nIndex, eImage);
	HPALETTE hPal;
	GetPalette(hPal); 

	HDC hDCSrc = CreateCompatibleDC(NULL);
	HDC hDCDest = GetMemDC();

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
		     nRelativeIndex*iseV1.m_sizeImage.cx, 
		     0, 
		     iseV1.m_sizeImage.cx, 
		     iseV1.m_sizeImage.cy, 
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

HRESULT ImageSizeEntry::GetMaskBitmap(const int& nIndex, BOOL bMask, HBITMAP& hBitmapMask, OLE_COLOR ocColor)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return E_FAIL;

	BOOL bResult = FALSE;
	if (!bMask)
	{
		bResult = CreateMaskBitmap(hDC, 
								   (HBITMAP)m_aImageBitmaps.GetAt(BitmapIndex(nIndex, eImage)), 
								   RelativeIndex(nIndex, eImage), 
								   hBitmapMask);
	}
	else
	{
		HPALETTE hPalDestOld, hPalMaskOld;
		HPALETTE hPal = NULL;
		HBITMAP hBitmapDestOld;
		HBITMAP hBitmapMaskOld;
		HDC hDCDest = NULL;
		HDC hDCMask = NULL;
		
		int nMaskBitmapIndex = BitmapIndex(nIndex, eMask);
		int nMaskRelativeIndex = RelativeIndex(nIndex, eMask);

		HBITMAP	hMask = (HBITMAP)m_aMaskBitmaps.GetAt(nMaskBitmapIndex);
		if (NULL == hMask)
			goto CleanUp;

		hDCDest = CreateCompatibleDC(hDC);
		if (NULL == hDCDest)
			goto CleanUp;

		hDCMask = GetMemDC();
		if (NULL == hDCMask)
			goto CleanUp;

		GetPalette(hPal); 
		if (hPal)
		{
			hPalDestOld = SelectPalette(hDCDest, hPal, FALSE);
			RealizePalette(hDCDest);

			hPalMaskOld = SelectPalette(hDCMask, hPal, FALSE);
			RealizePalette(hDCMask);
		}
		
		hBitmapMask = CreateBitmap(iseV1.m_sizeImage.cx, 
							       iseV1.m_sizeImage.cy,
								   1,
								   1,
								   0);

		if (NULL == hBitmapMask)
			goto CleanUp;

		hBitmapDestOld = SelectBitmap(hDCDest, hBitmapMask);
		
		::SetTextColor(hDCDest,RGB(0,0,0));
		::SetBkColor(hDCDest,RGB(255,255,255));
		
		hBitmapMaskOld = SelectBitmap(hDCMask, hMask);
		
		bResult = ::BitBlt(hDCDest, 
						   0, 
						   0, 
						   iseV1.m_sizeImage.cx, 
						   iseV1.m_sizeImage.cy,
						   hDCMask,
						   nMaskRelativeIndex*iseV1.m_sizeImage.cx,
						   0,
						   SRCCOPY);

		SelectBitmap(hDCMask, hBitmapMaskOld);

		if (hPal)
		{
			SelectPalette(hDCDest, hPalDestOld, FALSE);
			SelectPalette(hDCMask, hPalMaskOld, FALSE);
		}

		SelectBitmap(hDCDest, hBitmapDestOld);

CleanUp:
		if (hDCDest)
			DeleteDC(hDCDest);
	}	
	ReleaseDC(NULL, hDC);

	if (!bResult)
		return E_FAIL;
	return S_OK;
}

//
// SetMaskBitmap
//

HRESULT ImageSizeEntry::SetMaskBitmap(int& nIndex, HBITMAP hBitmapMask)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return E_FAIL;

	HPALETTE hPalOld1 = NULL;
	HPALETTE hPalOld2 = NULL;
	HPALETTE hPal = NULL;
	HRESULT hResult;
	HBITMAP hBitmapOld1 = NULL;
	HBITMAP hBitmapOld2 = NULL;
	HBITMAP hMask = NULL;
	BOOL bResult = FALSE;
	HDC hDCDest = NULL;
	HDC hDCSrc = NULL;
	int nRelativeIndex = 0;
	int nXPos;
	
	if (-1 == nIndex)
	{
		if (0 == m_aMaskBitmaps.GetSize())
		{
			// Initial allocation
			int nWidth = iseV1.m_sizeImage.cx * iseV1.m_nMaskInitial;
			hMask = CreateCompatibleBitmap(hDC, nWidth, iseV1.m_sizeImage.cy);
			if (NULL == hMask)
				goto CleanUp;
			hResult = m_aMaskBitmaps.Add(hMask);
			nIndex = 0;
			for (int nIndex2 = 1; nIndex2 < iseV1.m_nMaskInitial; nIndex2++)
				hResult = m_aMaskFreeList.Add(nIndex2);
		}
		else if (m_aMaskFreeList.GetSize() > 0)
		{
			// Check the free list
			nIndex = m_aMaskFreeList.GetAt(0);
			m_aMaskFreeList.RemoveAt(0);
			hMask = m_aMaskBitmaps.GetAt(BitmapIndex(nIndex, eMask));
		}
		else
		{
			int nSize = m_aMaskBitmaps.GetSize();
			int nWidth = iseV1.m_sizeImage.cx * iseV1.m_nMaskGrowBy;
			hMask = CreateCompatibleBitmap(hDC, nWidth, iseV1.m_sizeImage.cy);
			if (NULL == hMask)
				goto CleanUp;

			hResult = m_aMaskBitmaps.Add(hMask);
			nIndex = iseV1.m_nMaskInitial + ((nSize-1) *  iseV1.m_nMaskGrowBy);
			for (int nIndex2 = nIndex+1; nIndex2 < iseV1.m_nMaskGrowBy+nIndex; nIndex2++)
				hResult = m_aMaskFreeList.Add(nIndex2);
		}
	}
	else
		hMask = m_aMaskBitmaps.GetAt(BitmapIndex(nIndex, eMask));

	nRelativeIndex = RelativeIndex(nIndex, eMask);
	GetPalette(hPal); 

	hDCSrc = CreateCompatibleDC(NULL);
	if (NULL == hDCSrc)
		goto CleanUp;

	hDCDest = GetMemDC();
	if (NULL == hDCDest)
		goto CleanUp;

	hBitmapOld1 = SelectBitmap(hDCSrc, hBitmapMask);
	hBitmapOld2 = SelectBitmap(hDCDest, hMask);

	if (hPal)
	{
		hPalOld1 = SelectPalette(hDCSrc, hPal, FALSE);
		RealizePalette(hDCSrc);

		hPalOld2 = SelectPalette(hDCDest, hPal, FALSE);
		RealizePalette(hDCDest);
	}
	
	nXPos = nRelativeIndex*iseV1.m_sizeImage.cx;
	bResult = ::BitBlt(hDCDest, 
					   nXPos, 
					   0, 
					   iseV1.m_sizeImage.cx, 
					   iseV1.m_sizeImage.cy, 
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

CleanUp:
	if (hDCSrc)
		DeleteDC(hDCSrc);

	ReleaseDC(NULL, hDC);
	if (!bResult)
		return E_FAIL;
	return S_OK;
}

//
// BitBlt
//

HRESULT ImageSizeEntry::BitBlt (HDC hDC, int nIndex, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	if (-1 == nIndex)
		return E_FAIL;

	int nBitmapIndex = BitmapIndex(nIndex, eImage);
	int nRelativeIndex = RelativeIndex(nIndex, eImage);
	HBITMAP hImage = m_aImageBitmaps.GetAt(nBitmapIndex);
	assert(hImage);
	if (NULL == hImage)
		return E_FAIL;

	HPALETTE hPal, hPalOld;
	HBITMAP hBitmapOld;
	BOOL bResult = FALSE;
	HDC hDCMem;
	GetPalette(hPal); 
	if (hPal)
	{
		hPalOld = SelectPalette(hDC, hPal, FALSE);
		RealizePalette(hDC);
	}

	hDCMem = GetMemDC();
	if (NULL == hDCMem)
		goto CleanUp;

	hBitmapOld = SelectBitmap(hDCMem, hImage);
	
	if (bScale)
	{
		bResult = ::StretchBlt(hDC, 
							   x, 
							   y,
							   nHeight,
							   nWidth,
							   hDCMem,
							   nRelativeIndex*iseV1.m_sizeImage.cx,
							   0,
							   iseV1.m_sizeImage.cx,
							   iseV1.m_sizeImage.cy,
							   dwRop);
	}
	else
	{
		bResult = ::BitBlt(hDC, 
						   x, 
						   y, 
						   iseV1.m_sizeImage.cx, 
						   iseV1.m_sizeImage.cy, 
						   hDCMem, 
						   nRelativeIndex*iseV1.m_sizeImage.cx, 
						   0,
						   dwRop);
	}
	SelectBitmap(hDCMem, hBitmapOld);
	if (hPal)
		SelectPalette(hDC, hPalOld, FALSE);

CleanUp:
	if (!bResult)
		return E_FAIL;
	return S_OK;
}

//
// BitBltGray
//

HRESULT ImageSizeEntry::BitBltGray (HDC hDC, int nIndex, int x, int y, DWORD dwRop, BOOL bScale, int nHeight, int nWidth)
{
	if (-1 == nIndex)
		return E_FAIL;

	int nBitmapIndex = BitmapIndex(nIndex, eImage);
	int nRelativeIndex = RelativeIndex(nIndex, eImage);
	HBITMAP hImage = m_aImageBitmaps.GetAt(nBitmapIndex);
	assert(hImage);
	if (NULL == hImage)
		return E_FAIL;

	HPALETTE hPal;
	GetPalette(hPal); 

	HDC hDCMem = GetMemDC();
	if (NULL == hDCMem)
		return E_FAIL;

	HBITMAP hGray = m_aGrayBitmaps.GetAt(nBitmapIndex);
	if (NULL == hGray)
	{
		CDib dibGray;
		dibGray.FromBitmap(hImage, hDC, 8);
		hGray = dibGray.MakeGray();
		m_aGrayBitmaps.SetAt(nBitmapIndex, hGray);
		if (NULL == hGray)
			return E_FAIL;
	}

	BOOL bResult;
	HBITMAP hBitmapOld = SelectBitmap(hDCMem, hGray);
	if (bScale)
	{
		bResult = ::StretchBlt(hDC, 
							   x, 
							   y,
							   nHeight,
							   nWidth,
							   hDCMem,
							   nRelativeIndex * iseV1.m_sizeImage.cx,
							   0,
							   iseV1.m_sizeImage.cx,
							   iseV1.m_sizeImage.cy,
							   dwRop);
	}
	else
	{
		bResult = ::BitBlt(hDC, 
						   x, 
						   y,
						   iseV1.m_sizeImage.cx,
						   iseV1.m_sizeImage.cy,
						   hDCMem,
						   nRelativeIndex * iseV1.m_sizeImage.cx,
						   0,
						   dwRop);
	}
	SelectBitmap(hDCMem, hBitmapOld);
	if (!bResult)
		return E_FAIL;
	return S_OK;
}

//
// BitBltMask
//

HRESULT ImageSizeEntry::BitBltMask (HDC hDC, BOOL bMask, int nIndex, int x, int y, DWORD dwRop, OLE_COLOR ocMaskColor, BOOL bScale, int nHeight, int nWidth)
{
	if (-1 == nIndex)
		return E_FAIL;

	HBITMAP hMask = NULL;
	int nRelativeIndex;
	if (!bMask)
	{
		if (NULL == m_hMask)
		{
			m_hMask = ::CreateBitmap(iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, 1, 1, 0);
			if (NULL == m_hMask)
				return E_FAIL;
		}

		CreateMaskBitmap(hDC, 
						 m_aImageBitmaps.GetAt(BitmapIndex(nIndex, eImage)), 
						 RelativeIndex(nIndex, eImage),
						 m_hMask,
						 ocMaskColor);
		hMask = m_hMask;
	}
	else
	{
		nRelativeIndex = RelativeIndex(nIndex, eMask);
		hMask = m_aMaskBitmaps.GetAt(BitmapIndex(nIndex, eMask));
	}
	assert(hMask);
	if (NULL == hMask)
		return E_FAIL;

	HPALETTE hPal;
	GetPalette(hPal); 

	HDC hDCMem = GetMemDC();
	if (NULL == hDCMem)
		return E_FAIL;

	BOOL bResult = FALSE;
	HBITMAP hBitmapOld = SelectBitmap(hDCMem, hMask);
	if (bScale)
	{
		if (0 != nHeight || 0 != nWidth)
		{
			bResult = ::StretchBlt(hDC, 
								   x, 
								   y,
								   nHeight,
								   nWidth,
								   hDCMem,
								   !bMask ? 0 : nRelativeIndex * iseV1.m_sizeImage.cx,
								   0,
								   iseV1.m_sizeImage.cx,
								   iseV1.m_sizeImage.cy,
								   dwRop);
		}
	}
	else
	{
		bResult = ::BitBlt(hDC, 
						   x, 
						   y,
						   iseV1.m_sizeImage.cx,
						   iseV1.m_sizeImage.cy,
						   hDCMem,
						   !bMask ? 0 : nRelativeIndex * iseV1.m_sizeImage.cx,
						   0,
						   dwRop);
	}
	SelectBitmap(hDCMem, hBitmapOld);
	if (!bResult)
		return E_FAIL;
	return S_OK;
}

//
// BitBltDisabled
//

HRESULT ImageSizeEntry::BitBltDisabled (HDC hDC, int nIndex, int nMaskIndex, int x, int y, DWORD dwRop, OLE_COLOR ocMaskColor, BOOL bScale, int nHeight, int nWidth)
{
	if (-1 == nIndex)
		return E_FAIL;

	int nBitmapIndex = BitmapIndex(nIndex, eImage);
	int nRelativeIndex = RelativeIndex(nIndex, eImage);
	HBITMAP hImage = m_aImageBitmaps.GetAt(nBitmapIndex);
	if (NULL == hImage)
		return E_FAIL;

	HRESULT hResult;
	HPALETTE hPal;
	GetPalette(hPal); 

	HDC hDCScreen = GetDC(NULL);
	if (NULL == hDCScreen)
		return E_FAIL;

	BOOL bResult = FALSE;
	COLORREF crBackOld = -1;
	COLORREF crForeOld = -1;
	HBITMAP hBitmapOld;
	HBITMAP hMask = NULL;
	HDC hDCColorBase;
	HDC hDCMask = ::CreateCompatibleDC(hDCScreen);
	if (NULL == hDCMask)
		goto CleanUp;

	if (bScale)
		hMask = ::CreateBitmap(nHeight, nWidth, 1, 1, 0);
	else
		hMask = ::CreateBitmap(iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, 1, 1, 0);

	if (NULL == hMask)
		goto CleanUp;

	hBitmapOld = SelectBitmap(hDCMask, hMask);

	bResult = ::PatBlt(hDCMask, 0, 0, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, BLACKNESS);
	if (!bResult)
		goto CleanUp;

	// Fix the white portion
	hDCColorBase = GetMemDC();
	if (NULL == hDCColorBase)
		goto CleanUp;

	::SetBkColor(hDCColorBase, RGB(255, 255, 255));
	::SetTextColor(hDCColorBase, RGB(0, 0, 0));
	hResult = BitBlt(hDCMask, nIndex, 0, 0, SRCPAINT);
	if (FAILED(hResult))
		goto CleanUp;

	::SetBkColor(hDCColorBase, RGB(192, 192, 192));
	::SetTextColor(hDCColorBase, RGB(0, 0, 0));
	hResult = BitBlt(hDCMask, nIndex, 0, 0, SRCPAINT);
	if (FAILED(hResult))
		goto CleanUp;

	::SetBkColor(hDCColorBase, RGB(255, 255, 255));
	::SetTextColor(hDCColorBase, RGB(0, 0, 0));
	if (-1 == nMaskIndex)
		hResult = BitBltMask(hDCMask, FALSE, nIndex, 0, 0, SRCPAINT);
	else
		hResult = BitBltMask(hDCMask, TRUE, nMaskIndex, 0, 0, SRCPAINT);
	if (FAILED(hResult))
		goto CleanUp;

	// HighLight
	crBackOld = SetBkColor(hDC, RGB(255,255,255));
	crForeOld = SetTextColor(hDC, RGB(0,0,0));
	if (bScale)
		bResult = ::StretchBlt(hDC, x+1, y+1, nHeight, nWidth, hDCMask, 0, 0, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, SRCAND);
	else
		bResult = ::BitBlt(hDC, x+1, y+1, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, hDCMask, 0, 0, SRCAND);
	if (!bResult)
		goto CleanUp;

	SetBkColor(hDC, RGB(0,0,0));
	SetTextColor(hDC, GetSysColor(COLOR_BTNHILIGHT));
	if (bScale)
		bResult = ::StretchBlt(hDC, x+1, y+1, nHeight, nWidth, hDCMask, 0, 0, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, SRCPAINT);
	else
		bResult = ::BitBlt(hDC, x+1, y+1, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, hDCMask, 0, 0, SRCPAINT);
	if (!bResult)
		goto CleanUp;

	// Shadow
	SetBkColor(hDC, RGB(255,255,255));
	SetTextColor(hDC, RGB(0,0,0));
	if (bScale)
		bResult = ::StretchBlt(hDC, x, y, nHeight, nWidth, hDCMask, 0, 0, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, SRCAND);
	else
		bResult = ::BitBlt(hDC, x, y, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, hDCMask, 0, 0, SRCAND);
	if (!bResult)
		goto CleanUp;

	SetBkColor(hDC, RGB(0,0,0));
	SetTextColor(hDC, GetSysColor(COLOR_BTNSHADOW));
	if (bScale)
		bResult = ::StretchBlt(hDC, x, y, nHeight, nWidth, hDCMask, 0, 0, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, SRCPAINT);
	else
		bResult = ::BitBlt(hDC, x, y, iseV1.m_sizeImage.cx, iseV1.m_sizeImage.cy, hDCMask, 0, 0, SRCPAINT);
	if (!bResult)
		goto CleanUp;

CleanUp:

	if (-1 != crBackOld)
		SetBkColor(hDC, crBackOld);
	if (-1 != crForeOld)
		SetTextColor(hDC, crForeOld);

	if (hMask)
	{
		SelectBitmap(hDCMask, hBitmapOld);
		DeleteBitmap(hMask);
	}

	if (hDCMask)
		DeleteDC(hDCMask);

	if (hDCScreen)
		ReleaseDC(NULL, hDCScreen);
	return S_OK;
}

//
// RemoveImage
//

HRESULT ImageSizeEntry::RemoveImage(int nIndex)
{
	if (-1 == nIndex)
		return E_FAIL;

	int nSize = m_aImageFreeList.GetSize();
	for (int nFreeIndex = 0; nFreeIndex < nSize; nFreeIndex++)
	{
		if (nIndex == m_aImageFreeList.GetAt(nFreeIndex))
			return E_FAIL;
	}
	HRESULT hResult = m_aImageFreeList.Add(nIndex);
	return hResult;
}

//
// RemoveMask
//

HRESULT ImageSizeEntry::RemoveMask(int nIndex)
{
	if (-1 == nIndex)
		return E_FAIL;

	int nSize = m_aMaskFreeList.GetSize();
	for (int nFreeIndex = 0; nFreeIndex < nSize; nFreeIndex++)
	{
		if (nIndex == m_aMaskFreeList.GetAt(nFreeIndex))
			return E_FAIL;
	}
	HRESULT hResult = m_aMaskFreeList.Add(nIndex);
	return hResult;
}

//
// GetSize
//

HRESULT ImageSizeEntry::GetSize(SIZE& sizeImage)
{
	sizeImage = iseV1.m_sizeImage;
	return S_OK;
}

//
// Exchange
//

HRESULT ImageSizeEntry::Exchange(CBar* pBar, IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		long nStreamSize;
		short nSize;
		short nSize2;
		ImageEntry* pImageEntry;
		HBITMAP     hBitmap;
		HRESULT     hResult;
		HPALETTE    hPal;
		long        nKey;
		int		    nIndex;
		int         nTemp;
		CDib		theBitmap;

		HDC hDC = GetDC(NULL);
		if (NULL == hDC)
			return E_FAIL;

		GetPalette(hPal);

		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving
			//

			nStreamSize = sizeof(nSize) + sizeof(iseV1);
			hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nSize = sizeof(iseV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&iseV1, nSize, 0);
			if (FAILED(hResult))
				return hResult;
			
			nSize = m_aImageBitmaps.GetSize();
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				hBitmap = m_aImageBitmaps.GetAt(nIndex);
				assert(hBitmap);
				if (hBitmap)
				{
					theBitmap.FromBitmap(hBitmap, hDC);
					theBitmap.Write(pStream);
				}
			}
			
			nSize = m_aImageFreeList.GetSize();
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;
			
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				nTemp = m_aImageFreeList.GetAt(nIndex);
				hResult = pStream->Write(&nTemp, sizeof(int), 0);
				if (FAILED(hResult))
					return hResult;
			}

			nSize = m_aMaskBitmaps.GetSize();
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;
			
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				hBitmap = m_aMaskBitmaps.GetAt(nIndex);
				assert(hBitmap);
				if (hBitmap)
				{
					theBitmap.FromBitmap(hBitmap, hDC);
					theBitmap.Write(pStream);
				}
			}
			
			nSize = m_aMaskFreeList.GetSize();
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;
			
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				nTemp = m_aMaskFreeList.GetAt(nIndex);
				hResult = pStream->Write(&nTemp, sizeof(int), 0);
				if (FAILED(hResult))
					return hResult;
			}

			nSize = m_mapImages.GetCount();
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;
			
			FPOSITION posMap = m_mapImages.GetStartPosition();
			while (posMap)
			{
				m_mapImages.GetNextAssoc(posMap, (LPVOID&)nKey, (LPVOID&)pImageEntry);
				if (pImageEntry)
				{
					hResult = pImageEntry->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}
			}
		}
		else
		{
			//
			// Loading
			//

			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= sizeof(nSize);
			if (nStreamSize <= 0)
				goto FinishedReading;

			nSize2 = sizeof(iseV1);
			hResult = pStream->Read(&iseV1, nSize < nSize2 ? nSize : nSize2, 0);
			if (FAILED(hResult))
				return hResult;

			if (nSize2 < nSize)
			{
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}
			nStreamSize -= nSize;
			if (nStreamSize <= 0)
				goto FinishedReading;

			nOffset.HighPart = 0;
			nOffset.LowPart = nStreamSize;
			hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
			if (FAILED(hResult))
				return hResult;

FinishedReading:
			m_aGrayBitmaps.RemoveAll();
			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				theBitmap.Read(pStream);
				if (8 == GetGlobals().m_nBitDepth)
					pBar->m_pColorQuantizer->ProcessImage(theBitmap);

				hResult = m_aImageBitmaps.Add(theBitmap.CreateBitmap(hDC));
				m_aGrayBitmaps.Add(NULL);
			}
			
			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				hResult = pStream->Read(&nTemp, sizeof(nTemp), 0);
				if (FAILED(hResult))
					return hResult;
				
				hResult = m_aImageFreeList.Add(nTemp);
			}

			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				theBitmap.Read(pStream);
				hResult = m_aMaskBitmaps.Add(theBitmap.CreateMonochromeBitmap(hDC));
			}
			
			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				hResult = pStream->Read(&nTemp, sizeof(nTemp), 0);
				if (FAILED(hResult))
					return hResult;
				
				hResult = m_aMaskFreeList.Add(nTemp);
			}

			DWORD dwImageId;
			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			for (nIndex = 0; nIndex < nSize; nIndex++)
			{
				pImageEntry = new ImageEntry(this);
				if (NULL == pImageEntry)
					return E_OUTOFMEMORY;
				
				hResult = pImageEntry->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
				
				pImageEntry->GetImageId(dwImageId);
				m_mapImages.SetAt((LPVOID)dwImageId, (LPVOID)pImageEntry);
				m_theImageMgr.SetImageEntry(dwImageId, pImageEntry);
			}
		}
		
		ReleaseDC(NULL, hDC);
		return hResult;
	}
	CATCH
	{
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
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
// CompactBitmap
//

BOOL ImageSizeEntry::CompactBitmap(FArray&   faBitmaps, 
								   FArray&   faFreeList, 
								   int&      nInitialBitmapSize, 
								   int       nGrowBy, 
								   SIZE	     sizeImage, 
								   ImageType eType)
{
	int nImage = 1;
	int nNumOfBitmaps = faBitmaps.GetSize();
	int nNumOfAvailableSlots = faFreeList.GetSize();
	if (0 == nNumOfBitmaps && 0 == nNumOfAvailableSlots)
	{
		nInitialBitmapSize = nGrowBy;
		return TRUE;
	}

	int nMaxNumOfImages = nInitialBitmapSize;
	for (nImage = 1; nImage < nNumOfBitmaps; nImage++)
		nMaxNumOfImages += nGrowBy;

	int nNumOfImages = nMaxNumOfImages - nNumOfAvailableSlots;
	if (nNumOfImages <= 0)
	{
		switch (eType)
		{
		case eMask:
		case eImage:
			nInitialBitmapSize = nGrowBy;
			for (nImage = 0; nImage < nNumOfBitmaps; nImage++)
				DeleteBitmap((HBITMAP)faBitmaps.GetAt(nImage));
			faBitmaps.RemoveAll();
			faFreeList.RemoveAll();
			break;
		}
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
		hBitmapNew = CreateBitmap(nNumOfImages * sizeImage.cx, iseV1.m_sizeImage.cy, 1, 1, 0);
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
	FPOSITION posMap = m_mapImages.GetStartPosition();
	while (posMap)
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
					 nRelativeIndex * iseV1.m_sizeImage.cx, 
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

	BOOL bResult;
	for (nImage = 0; nImage < nNumOfBitmaps; nImage++)
	{
		bResult = DeleteBitmap((HBITMAP)faBitmaps.GetAt(nImage));
		assert(bResult);
	}

	faBitmaps.RemoveAll();
	faFreeList.RemoveAll();
	HRESULT hResult = faBitmaps.Add(hBitmapNew);
	return SUCCEEDED(hResult);
}

//
// Compact
//

HRESULT ImageSizeEntry::Compact()
{
	BOOL bResult = CompactBitmap(m_aImageBitmaps, 
								 m_aImageFreeList, 
								 iseV1.m_nImageInitial, 
								 iseV1.m_nImageGrowBy, 
								 iseV1.m_sizeImage, 
								 eImage);
	if (!bResult)
		return E_FAIL;

	bResult = CompactBitmap(m_aMaskBitmaps, 
						    m_aMaskFreeList, 
						    iseV1.m_nMaskInitial, 
						    iseV1.m_nMaskGrowBy, 
						    iseV1.m_sizeImage, 
						    eMask);
	if (!bResult)
		return E_FAIL;

//	if (m_hBitmapGray)
//	{
//		int nResult = DeleteBitmap(m_hBitmapGray);
//		assert(0 != nResult);
//		m_hBitmapGray = NULL;
//	}
	return S_OK;
}

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

//
// StWritePalette
//

HRESULT StWritePalette(IStream* pStream, HPALETTE hPal)
{
	WORD nColors = 0;
	if (hPal)
	{
		GetObject(hPal, sizeof(WORD), &nColors);
		if (0 != nColors)
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
		}
	}
	pStream->Write(&nColors, sizeof(WORD), 0);
	return NOERROR;
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
    HBITMAP m_hbmMask = ::CreateBitmap(nWidth, nHeight, 1, 1, 0);
    
	// Select the mono bitmap into its DC.
    HBITMAP hbmMaskOld = SelectBitmap(hDCMask, m_hbmMask);
    
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
	return m_hbmMask;
}
