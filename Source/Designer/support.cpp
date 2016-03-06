//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <stddef.h>
#include "ipserver.h"
#include "Resource.h"
#include "Globals.h"
#include "..\EventLog.h"
#include "support.h"
#include "debug.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern HINSTANCE g_hInstance;

const WCHAR wszCtlSaveStream [] = L"CONTROLSAVESTREAM";
short _SpecialKeyState()
{
    // don't appear to be able to reduce number of calls to GetKeyState
    //
    BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
    BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
    BOOL bAlt   = (GetKeyState(VK_MENU) < 0);

    return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}

LPWSTR MakeWideStrFromAnsi(LPSTR psz,BYTE  bType)
{
    LPWSTR pwsz;
    int i;

    // arg checking.
    //
    if (!psz)
        return NULL;

    // compute the length of the required BSTR
    //
    i =  MultiByteToWideChar(CP_ACP, 0, psz, -1, NULL, 0);
    if (i <= 0) return NULL;

    // allocate the widestr
    //
    switch (bType) {
      case STR_BSTR:
        // -1 since it'll add it's own space for a NULL terminator
        //
        pwsz = (LPWSTR) SysAllocStringLen(NULL, i - 1);
        break;
      case STR_OLESTR:
        pwsz = (LPWSTR) CoTaskMemAlloc(i * sizeof(WCHAR));
        break;
      default:
		TRACE(1, "Bogus String Type in MakeWideStrFromAnsi.\n");
		return NULL;
    }

    if (!pwsz) return NULL;
    MultiByteToWideChar(CP_ACP, 0, psz, -1, pwsz, i);
    pwsz[i - 1] = 0;
    return pwsz;
}

//=--------------------------------------------------------------------------=
// MakeWideStrFromResId
//=--------------------------------------------------------------------------=
// given a resource ID, load it, and allocate a wide string for it.
//
// Parameters:
//    WORD            - [in] resource id.
//    BYTE            - [in] type of string desired.
//
// Output:
//    LPWSTR          - needs to be cast to desired string type.
//
// Notes:
//
LPWSTR MakeWideStrFromResourceId(WORD    wId,BYTE    bType)
{
    int i;
	TCHAR szTmp[512];
#ifdef UNICODE
	LPWSTR pwszTmp;
#endif

	i = LoadString(g_hInstance, wId, szTmp, 512*sizeof(TCHAR));
    // load the string from the resources.
    if (!i) 
		return NULL;

#ifdef UNICODE
	pwszTmp = (LPWSTR)CoTaskMemAlloc((i * sizeof(WCHAR)) + sizeof(WCHAR));
	memcpy(pwszTmp,szTmp,(i * sizeof(WCHAR)) + sizeof(WCHAR));
	return pwszTmp;
#else
    return MakeWideStrFromAnsi(szTmp, bType);
#endif
}

//=--------------------------------------------------------------------------=
// MakeWideStrFromWide
//=--------------------------------------------------------------------------=
// given a wide string, make a new wide string with it of the given type.
//
// Parameters:
//    LPWSTR            - [in]  current wide str.
//    BYTE              - [in]  desired type of string.
//
// Output:
//    LPWSTR
//
// Notes:
//
LPWSTR MakeWideStrFromWide
(
    LPWSTR pwsz,
    BYTE   bType
)
{
    LPWSTR pwszTmp;
    int i;

    if (!pwsz) return NULL;

    // just copy the string, depending on what type they want.
    //
    switch (bType) {
      case STR_OLESTR:
        i = lstrlenW(pwsz);
        pwszTmp = (LPWSTR)CoTaskMemAlloc((i * sizeof(WCHAR)) + sizeof(WCHAR));
        if (!pwszTmp) return NULL;
        memcpy(pwszTmp, pwsz, (sizeof(WCHAR) * i) + sizeof(WCHAR));
        break;

      case STR_BSTR:
        pwszTmp = (LPWSTR)SysAllocString(pwsz);
        break;
    }

    return pwszTmp;
}

//////////////// IOLEOBJECT

void WINAPI CopyOLEVERB(void *pvDest,const void *pvSrc,DWORD cbCopy)
{
	memcpy(pvDest,pvSrc,cbCopy);
	((OLEVERB *)pvDest)->lpszVerbName=MakeWideStrFromResourceId((UINT)(((OLEVERB *)pvSrc)->lpszVerbName), STR_OLESTR);
}


int     s_iXppli;            // Pixels per logical inch along width
int     s_iYppli;            // Pixels per logical inch along height
static  BYTE    s_fGotScreenMetrics; // Are above valid?

void CacheScreenMetrics(void)
{
    HDC hDCScreen;

    // we have to critical section this in case two threads are converting
    // things at the same time
    //
#ifdef DEF_CRITSECTION
    EnterCriticalSection(&g_CriticalSection);
#endif
    if (s_fGotScreenMetrics)
        goto Done;

    // we want the metrics for the screen
    //
    hDCScreen = GetDC(NULL);

    ASSERT(hDCScreen, "couldn't get a DC for the screen.");
    s_iXppli = GetDeviceCaps(hDCScreen, LOGPIXELSX);
    s_iYppli = GetDeviceCaps(hDCScreen, LOGPIXELSY);

    ReleaseDC(NULL, hDCScreen);
    s_fGotScreenMetrics = TRUE;

    // we're done with our critical seciton.  clean it up
    //
  Done:;
#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif
}

void PixelToHiMetric(const SIZEL *pixsizel,SIZEL *hisizel)
{
    CacheScreenMetrics();
    hisizel->cx = MAP_PIX_TO_LOGHIM(pixsizel->cx, s_iXppli);
    hisizel->cy = MAP_PIX_TO_LOGHIM(pixsizel->cy, s_iYppli);

}

void HiMetricToPixel(const SIZEL *hisizel,SIZEL *pixsizel)
{
    CacheScreenMetrics();
    pixsizel->cx = MAP_LOGHIM_TO_PIX(hisizel->cx, s_iXppli);
    pixsizel->cy = MAP_LOGHIM_TO_PIX(hisizel->cy, s_iYppli);
}

void PixelToTwips(const SIZEL *pixsizel,SIZEL *twipssizel)
{
    CacheScreenMetrics();
    twipssizel->cx = MAP_PIX_TO_TWIPS(pixsizel->cx, s_iXppli);
    twipssizel->cy = MAP_PIX_TO_TWIPS(pixsizel->cy, s_iYppli);

}

void TwipsToPixel(const SIZEL *twipssizel, SIZEL *pixsizel)
{
    CacheScreenMetrics();
    pixsizel->cx = MAP_TWIPS_TO_PIX(twipssizel->cx, s_iXppli);
    pixsizel->cy = MAP_TWIPS_TO_PIX(twipssizel->cy, s_iYppli);

}

//=--------------------------------------------------------------------------=
// _CreateOleDC
//=--------------------------------------------------------------------------=
// creates an HDC given a DVTARGETDEVICE structure.
HDC _CreateOleDC(DVTARGETDEVICE *ptd)
{
    LPDEVMODEW   pDevModeW;
    HDC          hdc;
	LPOLESTR     lpwszDriverName;
    LPOLESTR     lpwszDeviceName;
    LPOLESTR     lpwszPortName;

    // return screen DC for NULL target device
    if (!ptd)
        return CreateDC(_T("DISPLAY"), NULL, NULL, NULL);

    if (ptd->tdExtDevmodeOffset == 0)
        pDevModeW = NULL;
    else
        pDevModeW = (LPDEVMODEW)((LPSTR)ptd + ptd->tdExtDevmodeOffset);

    lpwszDriverName = (LPOLESTR)((BYTE*)ptd + ptd->tdDriverNameOffset);
    lpwszDeviceName = (LPOLESTR)((BYTE*)ptd + ptd->tdDeviceNameOffset);
    lpwszPortName   = (LPOLESTR)((BYTE*)ptd + ptd->tdPortNameOffset);

#ifdef UNICODE
	hdc = CreateDC(lpwszDriverName, lpwszDeviceName, lpwszPortName, pDevModeW);
#else

    DEVMODEA     DevModeA, *pDevModeA;
   
    MAKE_ANSIPTR_FROMWIDE(pszDriverName, lpwszDriverName);
    MAKE_ANSIPTR_FROMWIDE(pszDeviceName, lpwszDeviceName);
    MAKE_ANSIPTR_FROMWIDE(pszPortName,   lpwszPortName);

    if (pDevModeW) 
	{
        WideCharToMultiByte(CP_ACP, 0, pDevModeW->dmDeviceName, -1, (LPSTR)DevModeA.dmDeviceName, CCHDEVICENAME, NULL, NULL);
		memcpy(&DevModeA.dmSpecVersion, &pDevModeW->dmSpecVersion,
		offsetof(DEVMODEA, dmFormName) - offsetof(DEVMODEA, dmSpecVersion));
        WideCharToMultiByte(CP_ACP, 0, pDevModeW->dmFormName, -1, (LPSTR)DevModeA.dmFormName, CCHFORMNAME, NULL, NULL);
		memcpy(&DevModeA.dmLogPixels, &pDevModeW->dmLogPixels, sizeof(DEVMODEA) - offsetof(DEVMODEA, dmLogPixels));
        if (pDevModeW->dmDriverExtra) 
		{
            pDevModeA = (DEVMODEA *)HeapAlloc(g_hHeap, 0, sizeof(DEVMODEA) + pDevModeW->dmDriverExtra);
            if (!pDevModeA) 
				return NULL;
            memcpy(pDevModeA, &DevModeA, sizeof(DEVMODEA));
            memcpy(pDevModeA + 1, pDevModeW + 1, pDevModeW->dmDriverExtra);
        }
		else
            pDevModeA = &DevModeA;

		DevModeA.dmSize = sizeof(DEVMODEA);
    } 
	else
        pDevModeA = NULL;

	hdc = CreateDC(pszDriverName, pszDeviceName, pszPortName, pDevModeA);
    if (pDevModeA != &DevModeA) 
		HeapFree(g_hHeap, 0, pDevModeA);
#endif

    return hdc;
}


LPTSTR g_szReflectClassName= _T("XBuilderReflector");
BYTE g_fRegisteredReflect=FALSE;

HWND CreateReflectWindow(BOOL fVisible,HWND hwndParent,int x,int y,SIZEL *pSize,CALLBACKFUNC pReflectWindowProc)
{
    WNDCLASS wndclass;
    // first thing to do is register the window class.  crit sect this
    // so we don't have to move it into the control
    //
#ifdef DEF_CRITSECTION
    EnterCriticalSection(&g_CriticalSection);
#endif
    if (!g_fRegisteredReflect) 
	{

        memset(&wndclass, 0, sizeof(wndclass));
        wndclass.lpfnWndProc = pReflectWindowProc;
        wndclass.hInstance   = g_hInstance;
        wndclass.lpszClassName = g_szReflectClassName;
        if (!RegisterClass(&wndclass)) 
		{
            TRACE(1,"Couldn't Register Parking Window Class!");
#ifdef DEF_CRITSECTION
            LeaveCriticalSection(&g_CriticalSection);
#endif
            return NULL;
        }
        g_fRegisteredReflect = TRUE;
    }
#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif

    // go and create the window.
    return CreateWindowEx(0, g_szReflectClassName, NULL,
                          WS_CHILD | WS_CLIPSIBLINGS |((fVisible) ? WS_VISIBLE : 0),
                          x, y, pSize->cx, pSize->cy,
                          hwndParent,
                          NULL, g_hInstance, NULL);
}


LPTSTR szparkingclass=_T("XBuilder_Parking");
WNDPROC g_ParkingWindowProc = NULL;
HWND g_hwndParking=NULL;

HWND GetParkingWindow(void)
{
    WNDCLASS wndclass;

    // crit sect this creation for apartment threading support.
    //
//    EnterCriticalSection(&g_CriticalSection);
    if (g_hwndParking)
        goto CleanUp;

    ZeroMemory(&wndclass, sizeof(wndclass));
    wndclass.lpfnWndProc = (g_ParkingWindowProc) ? g_ParkingWindowProc : DefWindowProc;
    wndclass.hInstance   = g_hInstance;
    wndclass.lpszClassName = szparkingclass;

    if (!RegisterClass(&wndclass)) {
        TRACE(1, "Couldn't Register Parking Window Class!");
        goto CleanUp;
    }

    g_hwndParking = CreateWindow(szparkingclass, NULL, WS_POPUP, 0, 0, 0, 0, NULL, NULL, g_hInstance, NULL);
    if (g_hwndParking != NULL)
       ++g_cLocks;

    ASSERT(g_hwndParking, "Couldn't Create Global parking window!!");


  CleanUp:
//    LeaveCriticalSection(&g_CriticalSection);
    return g_hwndParking;
}

void WINAPI CopyAndAddRefObject(void *pDest,const void *pSource,DWORD dwSize)
{
    *((IUnknown **)pDest) = *((IUnknown **)pSource);
    if ((*((IUnknown **)pDest))!=NULL)
		(*((IUnknown **)pDest))->AddRef();
    return;
}

///////////////// PERSISTANCE

void PersistBin(IStream *pStream,void *ptr,int size,BOOL save)
{
	if (save)
		pStream->Write(ptr,size,NULL);
	else
		pStream->Read(ptr,size,NULL);
}

void PersistBSTR(IStream *pStream,BSTR *pbstr,BOOL save)
{
	short size;
	BSTR bstr;
	if (save)
	{
		bstr=*pbstr;
		if (bstr==NULL || *bstr==0)
			size=0;
		else
			size=(wcslen(bstr)+1)*sizeof(WCHAR);
		pStream->Write(&size,sizeof(short),NULL);
		if (size)
			pStream->Write(bstr,size,0);
	}
	else
	{
		pStream->Read(&size,sizeof(short),NULL);
		SysFreeString(*pbstr);
		if (size!=0)
		{
			*pbstr=SysAllocStringByteLen(NULL,size);
			pStream->Read(*pbstr,size,NULL);
		}
		else
			*pbstr=NULL;
	}
}

void PersistPict(IStream *pStream,CPictureHolder *pict,BOOL save)
{
}


void PersistVariant(IStream *pStream,VARIANT *v,BOOL save)
{
}


HRESULT PersistBagI2(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,short *ptr,BOOL save)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_I2;
	if (save)
	{
		v.iVal=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		hr=pPropBag->Read(propName,&v,pErrorLog);
		if (SUCCEEDED(hr))
			*ptr=v.iVal;
	}
	return hr;
}

HRESULT PersistBagI4(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,long *ptr,BOOL save)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_I4;
	if (save)
	{
		v.lVal=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		hr=pPropBag->Read(propName,&v,pErrorLog);
		if (SUCCEEDED(hr))
			*ptr=v.lVal;
	}
	return hr;
}

HRESULT PersistBagR4(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,float *ptr,BOOL save)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_R4;
	if (save)
	{
		v.fltVal=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		hr=pPropBag->Read(propName,&v,pErrorLog);
		if (SUCCEEDED(hr))
			*ptr=v.fltVal;
	}
	return hr;
}

HRESULT PersistBagR8(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,double *ptr,BOOL save)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_R8;
	if (save)
	{
		v.dblVal=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		hr=pPropBag->Read(propName,&v,pErrorLog);
		if (SUCCEEDED(hr))
			*ptr=v.dblVal;
	}
	return hr;
}

HRESULT PersistBagBSTR(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,BSTR *ptr,BOOL save)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_BSTR;
	if (save)
	{
		v.bstrVal=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		v.bstrVal=NULL;
		hr=pPropBag->Read(propName,&v,pErrorLog);
		SysFreeString(*ptr);
		if (SUCCEEDED(hr))
			*ptr=v.bstrVal;
		else
			*ptr=NULL;
	}
	return hr;
}


HRESULT PersistBagBOOL(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,VARIANT_BOOL *ptr,BOOL save)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_BOOL;
	if (save)
	{
		v.boolVal=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		hr=pPropBag->Read(propName,&v,pErrorLog);
		if (SUCCEEDED(hr))
			*ptr=v.boolVal;
	}
	return hr;
}

HRESULT PersistBagCurrency(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,CY *ptr,BOOL save)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_CY;
	if (save)
	{
		v.cyVal=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		hr=pPropBag->Read(propName,&v,pErrorLog);
		if (SUCCEEDED(hr))
			*ptr=v.cyVal;
	}
	return hr;
}

HRESULT PersistBagDate(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,DATE *ptr,BOOL save)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_DATE;
	if (save)
	{
		v.date=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		hr=pPropBag->Read(propName,&v,pErrorLog);
		if (SUCCEEDED(hr))
			*ptr=v.date;
	}
	return hr;
}

HRESULT PersistBagPict(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,CPictureHolder *pict,BOOL save)
{
	return E_FAIL;
}


HRESULT PersistBagVariant(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,VARIANT *ptr,BOOL save)
{
	HRESULT hr;
	if (save)
		hr=pPropBag->Write(propName,ptr);
	else
		hr=pPropBag->Read(propName,ptr,pErrorLog);
	return hr;
}

