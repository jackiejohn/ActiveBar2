#ifndef __SUPPORT_H__
#define __SUPPORT_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "enumx.h"
#include "ocidl.h"

#define STR_BSTR   0
#define STR_OLESTR 1
#define BSTRFROMANSI(x)    (BSTR)MakeWideStrFromAnsi((LPSTR)(x), STR_BSTR)
#define OLESTRFROMANSI(x)  (LPOLESTR)MakeWideStrFromAnsi((LPSTR)(x), STR_OLESTR)
#define BSTRFROMRESID(x)   (BSTR)MakeWideStrFromResourceId(x, STR_BSTR)
#define OLESTRFROMRESID(x) (LPOLESTR)MakeWideStrFromResourceId(x, STR_OLESTR)
#define COPYOLESTR(x)      (LPOLESTR)MakeWideStrFromWide(x, STR_OLESTR)
#define COPYBSTR(x)        (BSTR)MakeWideStrFromWide(x, STR_BSTR)

extern LPWSTR MakeWideStrFromAnsi(LPSTR, BYTE bType);
extern LPWSTR MakeWideStrFromResourceId(WORD, BYTE bType);
extern LPWSTR MakeWideStrFromWide(LPWSTR, BYTE bType);

extern void WINAPI CopyOLEVERB(void *pvDest,const void *pvSrc,DWORD cbCopy);
extern void WINAPI CopyAndAddRefObject(void *, const void *, DWORD);


extern void PixelToHiMetric(const SIZEL *pixsizel,SIZEL *hisizel);
extern void HiMetricToPixel(const SIZEL *hisizel,SIZEL *pixsizel);

extern HDC _CreateOleDC(DVTARGETDEVICE *ptd);

class TempBuffer
{
public:
	TempBuffer(ULONG cb)
    {
		m_fAlloc = (cb > 120);
		if (m_fAlloc)
			m_pbBuf = new char[cb];
		else
			m_pbBuf = &m_szBufT;
	}
	~TempBuffer()
	{
		if (m_pbBuf && m_fAlloc) 
			delete m_pbBuf; 
	}
	void *GetBuffer(void)
	{
		return m_pbBuf; 
	}
private:
	void *m_pbBuf;
	char  m_szBufT[120];  // temp buffer for small cases.
	int   m_fAlloc;
};

// this is pretty inefficient but is convenient
#define MAKE_ANSIPTR_FROMWIDE(ptrname, pwszUnicode) \
    WCHAR *__pwsz##ptrname = pwszUnicode?pwszUnicode:L""; \
    long __l##ptrname = (lstrlenW(__pwsz##ptrname) + 1) * sizeof(WCHAR); \
    TempBuffer __TempBuffer##ptrname(__l##ptrname); \
    WideCharToMultiByte(CP_ACP, 0, __pwsz##ptrname, -1, (LPSTR)__TempBuffer##ptrname.GetBuffer(), __l##ptrname, NULL, NULL); \
    LPSTR ptrname = (LPSTR)__TempBuffer##ptrname.GetBuffer()

#define MAKE_WIDEPTR_FROMANSI(ptrname, ansistr) \
    long __l##ptrname = (lstrlen(ansistr) + 1) * sizeof(WCHAR); \
    TempBuffer __TempBuffer##ptrname(__l##ptrname); \
    MultiByteToWideChar(CP_ACP, 0, ansistr, -1, (LPWSTR)__TempBuffer##ptrname.GetBuffer(), __l##ptrname); \
    LPWSTR ptrname = (LPWSTR)__TempBuffer##ptrname.GetBuffer()

#ifndef _UNICODE 
#define MAKE_TCHARPTR_FROMWIDE(ptrname, pwszUnicode) \
    WCHAR *__pwsz##ptrname = pwszUnicode?pwszUnicode:L""; \
    long __l##ptrname = (lstrlenW(__pwsz##ptrname) + 1) * sizeof(WCHAR); \
    TempBuffer __TempBuffer##ptrname(__l##ptrname); \
    WideCharToMultiByte(CP_ACP, 0, __pwsz##ptrname, -1, (LPSTR)__TempBuffer##ptrname.GetBuffer(), __l##ptrname, NULL, NULL); \
    LPTSTR ptrname = (LPTSTR)__TempBuffer##ptrname.GetBuffer()
#else 
#define MAKE_TCHARPTR_FROMWIDE(ptrname, pwszUnicode) \
	LPTSTR ptrname = (LPTSTR)pwszUnicode;
#endif

#ifndef _UNICODE 
#define MAKE_WIDEPTR_FROMTCHAR(ptrname, ansistr) \
    long __l##ptrname = (lstrlen(ansistr) + 1) * sizeof(WCHAR); \
    TempBuffer __TempBuffer##ptrname(__l##ptrname); \
    MultiByteToWideChar(CP_ACP, 0, ansistr, -1, (LPWSTR)__TempBuffer##ptrname.GetBuffer(), __l##ptrname); \
    LPWSTR ptrname = (LPWSTR)__TempBuffer##ptrname.GetBuffer()
#else
	LPTSTR ptrname = (LPTSTR)ansistr;
#endif

extern const WCHAR wszCtlSaveStream []; //def used by IPersistStorage


typedef LRESULT (CALLBACK *CALLBACKFUNC)(HWND hwnd,UINT msg,WPARAM  wParam,LPARAM  lParam);
HWND CreateReflectWindow(BOOL fVisible,HWND hwndParent,int x,int y,SIZEL *pSize,CALLBACKFUNC pReflectWindowProc);
HWND GetParkingWindow();
extern WNDPROC g_ParkingWindowProc;
extern HWND g_hwndParking;

class CPictureHolder;

extern void PersistBin(IStream *pStream,void *ptr,int size,BOOL save);
extern void PersistBSTR(IStream *pStream,BSTR *bstr,BOOL save);
extern void PersistPict(IStream *pStream,CPictureHolder *pict,BOOL save);
extern void PersistVariant(IStream *pStream,VARIANT *v,BOOL save);

extern HRESULT PersistBagI2(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,short *ptr,BOOL save);
extern HRESULT PersistBagI4(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,long *ptr,BOOL save);
extern HRESULT PersistBagR4(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,float *ptr,BOOL save);
extern HRESULT PersistBagR8(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,double *ptr,BOOL save);
extern HRESULT PersistBagBSTR(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,BSTR *ptr,BOOL save);
extern HRESULT PersistBagBOOL(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,VARIANT_BOOL *ptr,BOOL save);
extern HRESULT PersistBagCurrency(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,CY *ptr,BOOL save);
extern HRESULT PersistBagDate(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,DATE *ptr,BOOL save);
extern HRESULT PersistBagPict(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,CPictureHolder *pict,BOOL save);
extern HRESULT PersistBagVariant(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,VARIANT *ptr,BOOL save);

extern int     s_iXppli;            // Pixels per logical inch along width
extern int     s_iYppli;            // Pixels per logical inch along height
extern void CacheScreenMetrics(void);
#define HIMETRIC_PER_INCH   2540
#define TWIPS_PER_INCH   1440
#define MAP_PIX_TO_TWIPS(x,ppli)   ( (TWIPS_PER_INCH*(x) + ((ppli)>>1)) / (ppli) )
#define MAP_PIX_TO_LOGHIM(x,ppli)   ( (HIMETRIC_PER_INCH*(x) + ((ppli)>>1)) / (ppli) )
#define MAP_LOGHIM_TO_PIX(x,ppli)   ( ((ppli)*(x) + HIMETRIC_PER_INCH/2) / HIMETRIC_PER_INCH )
#define MAP_TWIPS_TO_PIX(x,ppli)   ( ((ppli)*(x)) / TWIPS_PER_INCH )

extern void PixelToHiMetric(const SIZEL *pixsizel,SIZEL *hisizel);
extern void HiMetricToPixel(const SIZEL *hisizel,SIZEL *pixsizel);
extern void PixelToTwips(const SIZEL *pixsizel,SIZEL *twipssizel);
extern void TwipsToPixel(const SIZEL *twipssizel,SIZEL *pixsizel);

#endif
