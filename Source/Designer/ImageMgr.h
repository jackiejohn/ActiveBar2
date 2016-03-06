#ifndef __CImageMgr_H__
#define __CImageMgr_H__
#include "..\Interfaces.h"
#include "Map.h"

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

struct DumpContext;

//
// Conversion Structures
//

//
// ImageInfo
//

struct ImageInfo
{
	ImageInfo();
	~ImageInfo();

	int      nImageIndex; 
	BOOL     bImageValid[4];
	BOOL     bMaskValid[4];
	COLORREF crMaskColor[4];
	DWORD    dwCookie;

	// Used only for special sized images
	HBITMAP bmp[4];
	HBITMAP bmpMask[4];
	HBITMAP bmpGray;
};

//
// BitmapMgr
//

struct BitmapMgr
{
	BitmapMgr();
	~BitmapMgr();

	unsigned short nCount;
	HBITMAP  hImage[4];
	HBITMAP  hMask[4];
	SIZE     sizeDefault;
	int      nImageIndex;
	int      nImageCount[4];
	int      nMaxCount[4];

	FArray   faImages;		

	void GetImageBitmap(ImageInfo* pImageInfo, int nType, HBITMAP& hImage);
	void GetMaskBitmap(ImageInfo* pImageInfo, int nType, HBITMAP& hImage);
};

//
// ImageEntry
//

class ImageSizeEntry;

class ImageEntry
{
public:
	ImageEntry(ImageSizeEntry* pImageSizeEntry, DWORD dwImageId = 0);
	ImageEntry();
	~ImageEntry();

	const ULONG& AddRef();
	ULONG Release();
	const ULONG& RefCnt() const;

	STDMETHOD(GetImageId)(DWORD& dwImageId);

	STDMETHOD(GetImageBitmap)(HBITMAP& hBitmapImage);
	STDMETHOD(SetImageBitmap)(HBITMAP hBitmapImage);

	STDMETHOD(GetMaskBitmap)(HBITMAP& hBitmapImage);
	STDMETHOD(SetMaskBitmap)(HBITMAP hBitmapImage);

	STDMETHOD(GetSize)(SIZE& sizeImage);

	STDMETHOD(BitBlt)(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	STDMETHOD(BitBltGray)(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	STDMETHOD(BitBltMask)(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	STDMETHOD(BitBltDisabled)(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);

	STDMETHOD(ReleaseImage)();
	STDMETHOD(ReleaseMask)();

	STDMETHOD(Exchange)(IStream* pStream, BOOL bSave, DWORD dwVersion);

	// Internal Use
	void ImageIndex(int nValue);
	int& ImageIndex();

	void MaskIndex(int nValue);
	int& MaskIndex();

	BOOL& MaskCreated();
	void MaskCreated(const BOOL& bMaskCreated);

	HRESULT GetMaskColor(OLE_COLOR* ocMaskColor);
	HRESULT SetMaskColor(OLE_COLOR ocMaskColor);

#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif

private:
	ImageSizeEntry* m_pImageSizeEntry;
	OLE_COLOR		m_ocMaskColor;
	DWORD			m_dwImageId;
	ULONG			m_nRefCnt;
	int				m_nImageIndex;
	int				m_nMaskIndex;
};

inline STDMETHODIMP ImageEntry::GetImageId(DWORD& dwImageId)
{
	dwImageId = m_dwImageId;
	return S_OK;
}

inline int& ImageEntry::ImageIndex()
{
	return m_nImageIndex;
}

inline void ImageEntry::ImageIndex(int nValue)
{
	m_nImageIndex = nValue;
}

inline int& ImageEntry::MaskIndex()
{
	return m_nMaskIndex;
}

inline void ImageEntry::MaskIndex(int nValue)
{
	m_nMaskIndex = nValue;
}

inline const ULONG& ImageEntry::AddRef()
{
	m_nRefCnt++;
	return m_nRefCnt;
}

inline const ULONG& ImageEntry::RefCnt() const
{
	return m_nRefCnt;
}

//
// ImageSizeEntry
//

class CImageMgr;

class ImageSizeEntry
{
public:
	enum ImageType
	{
		eImage,
		eMask
	};

	ImageSizeEntry(const SIZE& sizeImage, CImageMgr& theImageMgr);
	~ImageSizeEntry();

	STDMETHOD(GetSize)(SIZE& sizeImage);

	STDMETHOD(Add)(ImageEntry* pImageEntry);
	STDMETHOD(Remove)(ImageEntry* pImageEntry);
	
	STDMETHOD(GetImageBitmap)(const int& nIndex, const int& nMaskIndex, HBITMAP& hBitmapImage);
	STDMETHOD(SetImageBitmap)(int& nIndex, HBITMAP hBitmapImage);

	STDMETHOD(GetMaskBitmap)(const int& nIndex, BOOL bMask, HBITMAP& hBitmapImage, OLE_COLOR ocColor = -1);
	STDMETHOD(SetMaskBitmap)(int& nIndex, HBITMAP hBitmapImage);

	STDMETHOD(RemoveImage)(int nIndex);
	STDMETHOD(RemoveMask)(int nIndex);

	// Drawing
	STDMETHOD(BitBlt)(HDC hDC, int nIndex, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	STDMETHOD(BitBltGray)(HDC hDC, int nIndex, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	STDMETHOD(BitBltMask)(HDC hDC, BOOL bMask, int nIndex, int x, int y, DWORD dwRop, OLE_COLOR ocMaskColor = -1, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	STDMETHOD(BitBltDisabled)(HDC hDC, int nIndex, int nMaskIndex, int x, int y, DWORD dwRop, OLE_COLOR ocColor = -1, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);

	// Persistance 
	STDMETHOD(Exchange)(IStream* pStream, BOOL bSave, DWORD dwVersion);
	STDMETHOD(Compact)(); 

	BOOL CompactBitmap(FArray& faBitmaps, FArray& faFreeList, int& nInitialBitmapSize, int nGrowBy, SIZE sizeImage, ImageType eType);
	void GetColorDepth(short& nDepth);
	void GetPalette(HPALETTE& hPalette);

	void GrowBy(int nGrowBy, ImageType eType);
	void InitialSize(int nSize, ImageType eType);
	
	int BitmapIndex(int nIndex, ImageType eType);
	int RelativeIndex(int nIndex, ImageType eType);

	HBITMAP& GetMask();

#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif

private:
	void CreateMaskBitmap(HDC hDC, HBITMAP hBitmap, int nRelativeIndex, HBITMAP& hBitmapMask, COLORREF crColor = -1);

	SIZE	 m_sizeImage;
	
	FArray   m_faImageBitmaps;
	FArray   m_faImageFreeList;

	FArray   m_faMaskBitmaps;
	FArray   m_faMaskFreeList;
	
	FMap     m_mapImages;

	int      m_nImageInitial;
	int		 m_nImageGrowBy;
	int      m_nMaskInitial;
	int		 m_nMaskGrowBy;

	HBITMAP m_hBitmapGray;
	HBITMAP m_hMask;

	CImageMgr& m_theImageMgr;

	ImageEntry** m_ppArrayImageEntry;
	int m_nImageCount;
	friend class ImageEntry;
};

inline int ImageSizeEntry::BitmapIndex(int nIndex, ImageType nType)
{
	if (eImage == nType)
	{
		if (nIndex < m_nImageInitial)
			return 0;
		else 
			return (((nIndex - m_nImageInitial) / m_nImageGrowBy) + 1);
	}
	else
	{
		if (nIndex < m_nMaskInitial)
			return 0;
		else 
			return (((nIndex - m_nMaskInitial) / m_nMaskGrowBy) + 1);
	}
}

inline int ImageSizeEntry::RelativeIndex(int nIndex, ImageType nType)
{
	if (eImage == nType)
	{
		if (nIndex < m_nImageInitial)
			return nIndex;
		else 
			return ((nIndex - m_nImageInitial) % m_nImageGrowBy);
	}
	else
	{
		if (nIndex < m_nMaskInitial)
			return nIndex;
		else 
			return ((nIndex - m_nMaskInitial) % m_nMaskGrowBy);
	}
}

inline void ImageSizeEntry::GrowBy(int nGrowBy, ImageType nType)
{
	if (eImage == nType)
	{
		if (m_faImageBitmaps.GetSize() > 0)
			return;
		m_nImageGrowBy = nGrowBy;
	}
	else
	{
		if (m_faMaskBitmaps.GetSize() > 0)
			return;
		m_nMaskGrowBy = nGrowBy;
	}
}

inline void ImageSizeEntry::InitialSize(int nSize, ImageType nType)
{
	if (eImage == nType)
	{
		if (m_faImageBitmaps.GetSize() > 0)
			return;
		m_nImageInitial = nSize;
	}
	else
	{
		if (m_faMaskBitmaps.GetSize() > 0)
			return;
		m_nMaskInitial = nSize;
	}
}

inline HBITMAP& ImageSizeEntry::GetMask()
{
	return m_hMask;
}

//
// CImageMgr
//

class CImageMgr
{
public:
	CImageMgr();
	ULONG m_refCount;
	int m_objectIndex;
	~CImageMgr();

	enum ImageConstants
	{
		eDefaultImageSize = 16,
		eDefaultImageInitialElements = 175,
		eDefaultImageGrowBy = 25,
		eDefaultMaskInitialElements = 175,
		eDefaultMaskGrowBy = 25
	};

	// IImageMgr members

	STDMETHOD(get_BitDepth)(short *retval);
	STDMETHOD(put_BitDepth)(short val);
	STDMETHOD(get_Palette)(OLE_HANDLE *retval);
	STDMETHOD(put_Palette)(OLE_HANDLE val);
	STDMETHOD(CreateImage)(long*nImageId);
	STDMETHOD(CreateImage2)(long*nImageId,  int nCx,  int nCy);
	STDMETHOD(AddRefImage)( long nImageId);
	STDMETHOD(RefCntImage)( long nImageId, long*nRefCnt);
	STDMETHOD(ReleaseImage)( long nImageId);
	STDMETHOD(Size)( long nImageId,  long*nCx,  long*nCy);
	STDMETHOD(put_ImageBitmap)( long nImageId,  OLE_HANDLE hBitmap);
	STDMETHOD(get_ImageBitmap)( long nImageId,  OLE_HANDLE*hBitmap);
	STDMETHOD(put_MaskBitmap)( long nImageId,  OLE_HANDLE hBitmap);
	STDMETHOD(get_MaskBitmap)( long nImageId,  OLE_HANDLE*hBitmap);
	STDMETHOD(BitBltEx)(OLE_HANDLE hDC,  long nImageId,  long nX,  long nY, long nRop, ImageStyles isStyle);
	STDMETHOD(ScaleBlt)( OLE_HANDLE hDC,  long nImageId,  long nX,  long nY,  long nWidth,  long nHeight,  long nRop, ImageStyles isStyle);
	STDMETHOD(Compact)();
	STDMETHOD(ImageInitialSize)( short nCx,  short nCy,  short nSize);
	STDMETHOD(ImageGrowBy)( short nCx,  short nCy,  short nGrowBy);
	STDMETHOD(MaskInitialSize)( short nCx,  short nCy,  short nSize);
	STDMETHOD(MaskGrowBy)( short nCx,  short nCy,  short nGrowBy);
	STDMETHOD(put_MaskColor)(long nImageId, OLE_COLOR MaskColor);
	STDMETHOD(get_MaskColor)( long nImageId,  OLE_COLOR *MaskColor);
	// DEFS for IImageMgr
	
	
	// Events 
	//{END INTERFACEDEFS}
	ImageSizeEntry* GetImageSizeEntry(const SIZE& sizeImage);
	void SetImageEntry(DWORD dwImageId, ImageEntry* pImageEntry);
	void CleanUp();

	STDMETHOD(Exchange)(IStream* pStream, BOOL bSave, DWORD dwVersion);
	STDMETHOD(Convert)(BitmapMgr& theBitmapMgr, FArray& faTools);
	STDMETHOD(GetConvertInfo)(IStream* pStream, BOOL bSave, BitmapMgr& theBitmapMgr);

	ImageSizeEntry** m_ppArrayImageSizeEntry;
	int				 m_nCount;
	HPALETTE m_hPal;
	USHORT   m_nVersion;
	DWORD    m_dwMasterImageId;
	FMap     m_mapImages;
	FMap     m_mapImageSizes;
	FMap     m_mapConvertIds;
	int      m_nDepth;
};

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,int objectIndex);

inline void CImageMgr::SetImageEntry(DWORD dwImageId, ImageEntry* pImageEntry)
{
	m_mapImages.SetAt((LPVOID)dwImageId, (LPVOID)pImageEntry);
}

//
// Global Utility Functions
//

HRESULT StReadHBITMAP(IStream* pStream, HBITMAP* phBitmap, HBITMAP* phGrayBitmap, BOOL bMonoChrome, HPALETTE hPal);
HRESULT StWriteHBITMAP(IStream* pStream, HBITMAP hBitmap, HPALETTE hPal, int nBitsPerPixel);

HRESULT StReadPalette(IStream* pStream, HPALETTE& hPal);
HRESULT StWritePalette(IStream* pStream, HPALETTE hPal);

#endif
