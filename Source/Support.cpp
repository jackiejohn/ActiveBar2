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
#include "support.h"
#include "debug.h"
#include "globals.h"
#include "EventLog.h"
#include "Utility.h"
#include "resource.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif 
const WCHAR wszCtlSaveStream [] = L"CONTROLSAVESTREAM";

extern HINSTANCE g_hInstance;

void BSTRAppend(BSTR *str,WCHAR *str2)
{
	if (0 == str2)
		return;

	if (0 == (*str))
	{
		*str=SysAllocString(str2);
		return;
	}
	int len1=wcslen(*str);
	int len2=wcslen(str2);
	BSTR newStr=SysAllocStringByteLen(NULL,(len1+len2+1)*sizeof(WCHAR));
	memcpy(newStr,*str,len1*sizeof(WCHAR));
	memcpy(newStr+len1,str2,(len2+1)*sizeof(WCHAR));
	SysFreeString(*str);
	*str=newStr;;
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

LPTSTR g_szReflectClassName= _T("XBuilderReflector - Customize");
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
		BOOL bClassAlreadyRegistered = GetClassInfo(g_hInstance, g_szReflectClassName, &wndclass);
		if (!bClassAlreadyRegistered)
		{
			memset(&wndclass, 0, sizeof(wndclass));
			wndclass.lpfnWndProc = pReflectWindowProc;
			wndclass.hInstance   = g_hInstance;
			wndclass.lpszClassName = g_szReflectClassName;
			if (!RegisterClass(&wndclass)) 
			{
				TRACE(1, "Couldn't Register Parking Window Class!");
	#ifdef DEF_CRITSECTION
				LeaveCriticalSection(&g_CriticalSection);
	#endif
				return NULL;
			}
		}
	    g_fRegisteredReflect = TRUE;
    }
#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif

    // go and create the window.
    return CreateWindowEx(0, 
						  g_szReflectClassName, 
						  0,
                          WS_CHILD | WS_CLIPSIBLINGS |((fVisible) ? WS_VISIBLE : 0),
                          x, 
						  y, 
						  pSize->cx, 
						  pSize->cy,
                          hwndParent,
                          0, 
						  g_hInstance, 
						  0);
}


LPTSTR szparkingclass=_T("XBuilder_Parking - Customize");
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

//
// PERSISTANCE
//

void PersistBin(IStream *pStream,void *ptr,int size,VARIANT_BOOL vbSave)
{
	if (VARIANT_TRUE == vbSave)
		pStream->Write(ptr,size,NULL);
	else
		pStream->Read(ptr,size,NULL);
}

void PersistBSTR(IStream *pStream,BSTR *pbstr,VARIANT_BOOL vbSave)
{
	short size;
	BSTR bstr;
	if (VARIANT_TRUE == vbSave)
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



void PersistVariant(IStream *pStream,VARIANT *v,VARIANT_BOOL vbSave)
{
}


HRESULT PersistBagI2(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,short *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_I2;
	if (VARIANT_TRUE == vbSave)
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

HRESULT PersistBagI4(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,long *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_I4;
	if (VARIANT_TRUE == vbSave)
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

HRESULT PersistBagR4(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,float *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_R4;
	if (VARIANT_TRUE == vbSave)
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

HRESULT PersistBagR8(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,double *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_R8;
	if (VARIANT_TRUE == vbSave)
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

HRESULT PersistBagBSTR(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,BSTR *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_BSTR;
	if (VARIANT_TRUE == vbSave)
	{
		v.bstrVal=*ptr;
		hr=pPropBag->Write(propName,&v);
	}
	else
	{
		v.bstrVal=NULL;
		hr=pPropBag->Read(propName,&v,pErrorLog);
		if (SUCCEEDED(hr))
			*ptr=v.bstrVal;
	}
	return hr;
}


HRESULT PersistBagBOOL(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,VARIANT_BOOL *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_BOOL;
	if (VARIANT_TRUE == vbSave)
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

HRESULT PersistBagCurrency(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,CY *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_CY;
	if (VARIANT_TRUE == vbSave)
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

HRESULT PersistBagDate(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,DATE *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	v.vt=VT_DATE;
	if (VARIANT_TRUE == vbSave)
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

HRESULT PersistBagPict(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,CPictureHolder *pict,VARIANT_BOOL vbSave)
{
	return E_FAIL;
}


HRESULT PersistBagVariant(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,VARIANT *ptr,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	if (VARIANT_TRUE == vbSave)
		hr=pPropBag->Write(propName,ptr);
	else
		hr=pPropBag->Read(propName,ptr,pErrorLog);
	return hr;
}

int __cdecl Fatoi(const char *nptr)
{
        int c;              /* current char */
        long total;         /* current total */
        int sign;           /* if '-', then negative, otherwise positive */

        /* skip whitespace */
        while ( ((int)(unsigned char)*nptr)==32 )
            ++nptr;

        c = (int)(unsigned char)*nptr++;
        sign = c;           /* save sign indication */
        if (c == '-' || c == '+')
            c = (int)(unsigned char)*nptr++;    /* skip sign */

        total = 0;

        while (c>='0' && c<='9') {
            total = 10 * total + (c - '0');     /* accumulate digit */
            c = (int)(unsigned char)*nptr++;    /* get next char */
        }

        if (sign == '-')
            return -total;
        else
            return total;   /* return result, negated if necessary */
}

//
// Error Stuff
//
 
ErrorDialog::ErrorDialog(TCHAR *val) 
	: FDialog(IDD_ERROR) 
{
	errorStr=val;
};

BOOL ErrorDialog::DialogProc(UINT message,WPARAM wParam,LPARAM lParam)
{
	switch(message)
	{
	case WM_INITDIALOG:
		CenterDialog(GetParent(m_hWnd));
		SendMessage(GetDlgItem(m_hWnd,IDC_EDIT1),WM_SETTEXT,0,(LPARAM)errorStr);
		break;
	}
	return FDialog::DialogProc(message,wParam,lParam);
}

void ErrorObject::ThrowError(HRESULT hrExcep, BSTR bstrDescription, DWORD dwHelpContextID)
{
	//
    // First get the Create Error Info Object.
    //

	ICreateErrorInfo* pCreateErrorInfo;
	HRESULT hResult = CreateErrorInfo(&pCreateErrorInfo);
    if (FAILED(hResult)) 
		return;

	// set up some default information on it.
    pCreateErrorInfo->SetGUID((REFIID)INTERFACEOFOBJECT(0));
    
#ifdef _UNICODE
	pCreateErrorInfo->SetHelpFile((LPWSTR)HELPFILEOFOBJECT(0));
#else
	pCreateErrorInfo->SetHelpFile(m_bstrHelpFile);
#endif
    pCreateErrorInfo->SetHelpContext(dwHelpContextID);
    pCreateErrorInfo->SetDescription(bstrDescription);

    // load in the source
#ifdef _UNICODE
	pCreateErrorInfo->SetSource((LPWSTR)NAMEOFOBJECT(0));
#else
	WCHAR wszTmp[256];
    MultiByteToWideChar(CP_ACP, 0, NAMEOFOBJECT(0), -1, wszTmp, 256);
    pCreateErrorInfo->SetSource(wszTmp);
#endif

    // now set the Error info up with the system
    IErrorInfo* pErrorInfo;
    hResult = pCreateErrorInfo->QueryInterface(IID_IErrorInfo, (void**)&pErrorInfo);
    CLEANUP_ON_FAILURE(hResult);

    SetErrorInfo(0, pErrorInfo);
    pErrorInfo->Release();

CleanUp:
    pCreateErrorInfo->Release();
    return;
}

int ErrorObject::LookUpErrorIndex(UINT nErrorId)
{
	for (int nIndex = 1; nIndex < m_nErrorTableSize; nIndex++)
	{
		if (m_pErrorTable[nIndex].nErrorId == nErrorId)
			return nIndex;
	}
	return 0;
}

void ErrorObject::SendError(UINT nErrorId, BSTR bstrExtendedInfo)
{
	BSTR bstrDescription = NULL;
	int nIndex = LookUpErrorIndex(nErrorId);
	MAKE_WIDEPTR_FROMTCHAR(wDesc, LoadStringRes(m_pErrorTable[nIndex].nDescResId));
	bstrDescription=SysAllocString(wDesc);
	
	if (bstrExtendedInfo && *bstrExtendedInfo)
	{
		BSTRAppend(&bstrDescription,L"\r\n [ Extended Info :");
		BSTRAppend(&bstrDescription, bstrExtendedInfo);
		BSTRAppend(&bstrDescription,L"]");
	}
	if (m_nAsyncError)
	{
		IReturnString* pRetString = CReturnString::CreateInstance(0);
		if (pRetString)
		{
			pRetString->put_Value(bstrDescription);
			IReturnBool* pRetBool = CRetBool::CreateInstance(0);
			if (pRetBool)
			{
				pRetBool->put_Value(VARIANT_FALSE);
				AsyncError((short)nErrorId,
						   (IReturnString*)pRetString,
						   nErrorId,
						   (LPWSTR)NAMEOFOBJECT(0),
						   (LPWSTR)HELPFILEOFOBJECT(0),
						   m_pErrorTable[nIndex].nHelpContextId,
						   (IReturnBool*)pRetBool);
				pRetBool->Release();
			}
			pRetString->Release();
		}
	}
	else
		ThrowError(E_FAIL, bstrDescription, m_pErrorTable[nIndex].nHelpContextId);

	SysFreeString(bstrDescription);
}

