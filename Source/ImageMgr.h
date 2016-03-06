#ifndef __CImageMgr_H__
#define __CImageMgr_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "interfaces.h"

class CColorQuantizer;

//
//  Image Manager
//
//	Motivation behind it's current design
//
//	The ImageManager should be able to handle different size images in an efficient manner.
//	Should create mask bitmaps on the fly for images that do not have one defined.
//	Images are referenced counted, this allows them to be shared between tools. 
//
//	Classes
//	ImageMgr
//		Is the interface that clients use to create, destroy, and retrieve images.  It manages the ImageSizeEntry 
// and the ImageEntry Classes.
//
//	ImageSizeEntry
//		Maintains the storage for the images, provides functions to get, set, and display the images.  The storage 
// of the images are maintained in an array of bitmaps, the size of the bitmaps are determined by the image size and by the Initial and GrowBy sizes. 
//
//	ImageEntry
//		Represents a single image.  Maintains the Reference Count, Image Id, MaskIndex, ImageIndex, Mask Color for 
// the image.  
//

class ImageSizeEntry;
struct ImageEntryV1;

//
// ImageEntry
//

class ImageEntry
{
public:
	ImageEntry(ImageSizeEntry* pImageSizeEntry, DWORD dwImageId = 0, VARIANT_BOOL vbDesignerCreated = VARIANT_FALSE);
	~ImageEntry();

	const ULONG& AddRef();
	ULONG Release();
	const ULONG& RefCnt() const;

	HRESULT GetImageId(DWORD& dwImageId);

	HRESULT GetImageBitmap(HBITMAP& hBitmapImage, VARIANT_BOOL vbUseMask);
	HRESULT SetImageBitmap(HBITMAP hBitmapImage, VARIANT_BOOL vbDesignerUpdated);

	HRESULT GetMaskBitmap(HBITMAP& hBitmapImage);
	HRESULT SetMaskBitmap(HBITMAP hBitmapImage, VARIANT_BOOL vbDesignerUpdated);

	HRESULT GetSize(SIZE& sizeImage);

	HRESULT BitBlt(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	HRESULT BitBltGray(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	HRESULT BitBltMask(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	HRESULT BitBltDisabled(HDC hDC, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);

	HRESULT ReleaseImage();
	HRESULT ReleaseMask();

	DWORD& ImageId();
	ULONG& RefCount();

#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif

private:
	friend class ImageSizeEntry;
	friend class CImageMgr;

	HRESULT Exchange(IStream* pStream, VARIANT_BOOL vbSave);

	void ImageIndex(int nValue);
	int& ImageIndex();

	void MaskIndex(int nValue);
	int& MaskIndex();

	void GetMaskColor (OLE_COLOR* pocMaskColor);
	void SetMaskColor (OLE_COLOR ocMaskColor);

	ImageSizeEntry* m_pImageSizeEntry;

#pragma pack(push)
#pragma pack (1)
	struct ImageEntryV1
	{
		ImageEntryV1();

		VARIANT_BOOL    m_vbDesignerCreated:1;
		VARIANT_BOOL    m_vbDesignerUpdated:1;
		OLE_COLOR       m_ocMaskColor;
		DWORD			m_dwImageId;
		ULONG			m_nRefCnt;
		int				m_nImageIndex;
		int				m_nMaskIndex;
	} iepV1;
#pragma pack(pop)
};

inline DWORD& ImageEntry::ImageId()
{
	return iepV1.m_dwImageId;
}

inline ULONG& ImageEntry::RefCount()
{
	return iepV1.m_nRefCnt;
}

inline HRESULT ImageEntry::GetImageId(DWORD& dwImageId)
{
	dwImageId = iepV1.m_dwImageId;
	return S_OK;
}

inline int& ImageEntry::ImageIndex()
{
	return iepV1.m_nImageIndex;
}

inline void ImageEntry::ImageIndex(int nValue)
{
	iepV1.m_nImageIndex = nValue;
}

inline int& ImageEntry::MaskIndex()
{
	return iepV1.m_nMaskIndex;
}

inline void ImageEntry::MaskIndex(int nValue)
{
	iepV1.m_nMaskIndex = nValue;
}

inline const ULONG& ImageEntry::AddRef()
{
	iepV1.m_nRefCnt++;
	return iepV1.m_nRefCnt;
}

inline const ULONG& ImageEntry::RefCnt() const
{
	return iepV1.m_nRefCnt;
}


//
// GetMaskColor
//

inline void ImageEntry::GetMaskColor (OLE_COLOR* pocMaskColor)
{
	*pocMaskColor = iepV1.m_ocMaskColor;
}

//
// PutMaskColor
//

inline void ImageEntry::SetMaskColor (OLE_COLOR ocMaskColor)
{
	iepV1.m_ocMaskColor = ocMaskColor;
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

	HRESULT GetSize(SIZE& sizeImage);

	HRESULT Add(ImageEntry* pImageEntry);
	HRESULT Remove(ImageEntry* pImageEntry);
	
	HRESULT GetImageBitmap(const int& nIndex, const int& nMaskIndex, HBITMAP& hBitmapImage, VARIANT_BOOL vbUseMask);
	HRESULT SetImageBitmap(int& nIndex, HBITMAP hBitmapImage);

	HRESULT GetMaskBitmap(const int& nIndex, BOOL bMask, HBITMAP& hBitmapMask, OLE_COLOR ocColor = -1);
	HRESULT SetMaskBitmap(int& nIndex, HBITMAP hBitmapImage);

	HRESULT RemoveImage(int nIndex);
	HRESULT RemoveMask(int nIndex);

	// Drawing
	HRESULT BitBlt(HDC hDC, int nIndex, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	HRESULT BitBltGray(HDC hDC, int nIndex, int x, int y, DWORD dwRop, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	HRESULT BitBltMask(HDC hDC, BOOL bMask, int nIndex, int x, int y, DWORD dwRop, OLE_COLOR ocMaskColor = -1, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);
	HRESULT BitBltDisabled(HDC hDC, int nIndex, int nMaskIndex, int x, int y, DWORD dwRop, OLE_COLOR ocMaskColor = -1, BOOL bScale = FALSE, int nHeight = 0, int nWidth = 0);

	// Persistance 
	HRESULT Exchange(CBar* pBar, IStream* pStream, VARIANT_BOOL vbSave);

	BOOL CompactBitmap(FArray&   faBitmaps, 
					   FArray&   faFreeList, 
					   int&      nInitialBitmapSize, 
					   int       nGrowBy, 
					   SIZE	     sizeImage, 
					   ImageType eType);
	HRESULT Compact(); 

	void GetPalette(HPALETTE& hPalette);

	void GrowBy(int nGrowBy, ImageType eType);
	void InitialSize(int nSize, ImageType eType);
	
	int BitmapIndex(int nIndex, ImageType eType);
	int RelativeIndex(int nIndex, ImageType eType);

	CImageMgr& GetImageMgr() const;

	HBITMAP& GetMask();

#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif

private:
	BOOL CreateMaskBitmap(HDC hDC, HBITMAP hBitmap, int nRelativeIndex, HBITMAP& hBitmapMask, OLE_COLOR ocColor = -1);

	TypedArray<HBITMAP> m_aImageBitmaps;
	TypedArray<int> m_aImageFreeList;

	TypedArray<HBITMAP> m_aGrayBitmaps;

	TypedArray<HBITMAP> m_aMaskBitmaps;
	TypedArray<int> m_aMaskFreeList;
	
	FMap m_mapImages;

#pragma pack(push)
#pragma pack (1)
	struct ImageSizeEntryV1
	{
		ImageSizeEntryV1();

		SIZE m_sizeImage;
		int  m_nImageInitial;
		int	 m_nImageGrowBy;
		int  m_nMaskInitial;
		int	 m_nMaskGrowBy;
	} iseV1;
#pragma pack(pop)

	HBITMAP m_hMask;

	CImageMgr& m_theImageMgr;

	friend class ImageEntry;
};

inline CImageMgr& ImageSizeEntry::GetImageMgr() const
{
	return m_theImageMgr;
}

inline int ImageSizeEntry::BitmapIndex(int nIndex, ImageType nType)
{
	switch (nType)
	{
	case eImage:
		return ((nIndex < iseV1.m_nImageInitial) ? 0 : (((nIndex - iseV1.m_nImageInitial) / iseV1.m_nImageGrowBy) + 1));
	case eMask:
		return ((nIndex < iseV1.m_nMaskInitial) ? 0 : (((nIndex - iseV1.m_nMaskInitial) / iseV1.m_nMaskGrowBy) + 1));
	default:
		assert(FALSE);
	}
	return -1;
}

inline int ImageSizeEntry::RelativeIndex(int nIndex, ImageType nType)
{
	switch (nType)
	{
	case eImage:
		return ((nIndex < iseV1.m_nImageInitial) ? nIndex : ((nIndex - iseV1.m_nImageInitial) % iseV1.m_nImageGrowBy));
	case eMask:
		return ((nIndex < iseV1.m_nMaskInitial) ? nIndex : ((nIndex - iseV1.m_nMaskInitial) % iseV1.m_nMaskGrowBy));
	default:
		assert(FALSE);
	}
	return -1;
}

inline void ImageSizeEntry::GrowBy(int nGrowBy, ImageType nType)
{
	switch (nType)
	{
	case eImage:
		if (0 == m_aImageBitmaps.GetSize())
			iseV1.m_nImageGrowBy = nGrowBy;
		break;

	case eMask:
		if (0 == m_aMaskBitmaps.GetSize())
			iseV1.m_nMaskGrowBy = nGrowBy;
		break;
	default:
		assert(FALSE);
	}
}

inline void ImageSizeEntry::InitialSize(int nSize, ImageType nType)
{
	switch (nType)
	{
	case eImage:
		if (0 == m_aImageBitmaps.GetSize())
			iseV1.m_nImageInitial = nSize;
		break;
	case eMask:
		if (0 == m_aMaskBitmaps.GetSize())
			iseV1.m_nMaskInitial = nSize;
		break;
	default:
		assert(FALSE);
	}
}

inline HBITMAP& ImageSizeEntry::GetMask()
{
	return m_hMask;
}

//
// CImageMgr
//

class CImageMgr : public IUnknown
{
public:
	CImageMgr();
	ULONG m_refCount;
	int m_objectIndex;
	~CImageMgr();

#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif

	enum ImageConstants
	{
		eDefaultImageSize = 16,
		eDefaultImageInitialElements = 100,
		eDefaultImageGrowBy = 25,
		eDefaultMaskInitialElements = 100,
		eDefaultMaskGrowBy = 25
	};

	void SetImageEntry(DWORD dwImageId, ImageEntry* pImageEntry);
	BOOL FindImageId(DWORD nId);

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CImageMgr *CreateInstance(IUnknown *pUnkOuter);
	// IUnknown members
	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
	STDMETHOD_(ULONG, AddRef)(void);
	STDMETHOD_(ULONG, Release)(void);

	// IImageMgr members

	STDMETHOD(get_Palette)(OLE_HANDLE *retval);
	STDMETHOD(put_Palette)(OLE_HANDLE val);
	STDMETHOD(CreateImage)(long*nImageId, VARIANT_BOOL bDesignerCreated);
	STDMETHOD(CreateImageEx)(long*nImageId,  int nCx,  int nCy, VARIANT_BOOL bDesignerCreated);
	STDMETHOD(AddRefImage)( long nImageId);
	STDMETHOD(RefCntImage)( long nImageId, long*nRefCnt);
	STDMETHOD(ReleaseImage)( long*nImageId);
	STDMETHOD(Size)( long nImageId,  long*nCx,  long*nCy);
	STDMETHOD(PutImageBitmap)( long nImageId,  VARIANT_BOOL vbDesignerCreated,  OLE_HANDLE hBitmap);
	STDMETHOD(GetImageBitmap)( long nImageId,  VARIANT_BOOL vbUseMask,  OLE_HANDLE*hBitmap );
	STDMETHOD(PutMaskBitmap)( long nImageId,  VARIANT_BOOL vbDesignerCreated,  OLE_HANDLE hBitmap);
	STDMETHOD(GetMaskBitmap)( long nImageId,  OLE_HANDLE*hBitmap);
	STDMETHOD(BitBltEx)(OLE_HANDLE hDC,  long nImageId,  long nX,  long nY, long nRop, ImageStyles isStyle);
	STDMETHOD(ScaleBlt)( OLE_HANDLE hDC,  long nImageId,  long nX,  long nY,  long nWidth,  long nHeight,  long nRop, ImageStyles isStyle);
	STDMETHOD(Compact)();
	STDMETHOD(BitmapInitialSize)( short nCx,  short nCy,  short nSize);
	STDMETHOD(BitmapGrowBy)( short nCx,  short nCy,  short nGrowBy);
	STDMETHOD(MaskBitmapInitialSize)( short nCx,  short nCy,  short nSize);
	STDMETHOD(MaskBitmapGrowBy)( short nCx,  short nCy,  short nGrowBy);
	STDMETHOD(get_MaskColor)( long nImageId,  OLE_COLOR *MaskColor);
	STDMETHOD(put_MaskColor)( long nImageId,  OLE_COLOR MaskColor);
	// DEFS for IImageMgr
	
	
	// Events 
	//{END INTERFACEDEFS}
	ImageSizeEntry* GetImageSizeEntry(const SIZE& sizeImage);
	HRESULT Exchange (CBar* pBar, IStream*  pStream, VARIANT_BOOL vbSave);
	HRESULT ExchangeConfig(CBar* pBar, IStream* pStream, VARIANT_BOOL vbSave);

	void CleanUp();
	HRESULT BitBltExInline(OLE_HANDLE hDC,  ImageEntry* pImageEntry,  long nX,  long nY, long nRop, ImageStyles isStyle);


#pragma pack(push)
#pragma pack (1)
	struct ImageManagerV1
	{
		ImageManagerV1();

		DWORD m_dwMasterImageId;
	} imV1;
#pragma pack(pop)
	HPALETTE m_hPal;

	FMap     m_mapImages;
	FMap     m_mapImageSizes;
	FMap     m_mapConvertIds;

	long m_nImageLoad;
};

inline void CImageMgr::SetImageEntry(DWORD dwImageId, ImageEntry* pImageEntry)
{
	m_mapImages.SetAt((LPVOID)dwImageId, (LPVOID)pImageEntry);
}

inline BOOL CImageMgr::FindImageId(DWORD nId)
{
	ImageEntry* pImageEntry;
	return m_mapImages.Lookup((LPVOID)nId, (LPVOID&)pImageEntry);
}

inline HRESULT CImageMgr::BitBltExInline(OLE_HANDLE hDC,  ImageEntry* pImageEntry,  long nX,  long nY, long nRop, ImageStyles isStyle)
{
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
// Global Utility Functions
//

HRESULT StReadPalette(IStream* pStream, HPALETTE& hPal);
HRESULT StWritePalette(IStream* pStream, HPALETTE hPal);
HBITMAP CreateMaskBitmap(HDC hDC, HBITMAP hBitmap, int nWidth, int nHeight);

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif
