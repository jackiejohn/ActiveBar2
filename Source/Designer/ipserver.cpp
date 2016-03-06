//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <olectl.h>
#include <crtdbg.h>
#include "..\EventLog.h"
#include "ipserver.h"
#include "fregkey.h"
#include "support.h"
#include "GdiUtil.h"
#include "IconEdit.h"
#include "resource.h"
#include "DesignerInterfaces.h"

//{BEGIN AUTOINCLUDES}
#include "catidreg.h"
//{END AUTOINCLUDES}


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


//{BEGIN_FN()
extern PROPERTYPAGEINFO DesignerPageObject;
extern IUnknown *CreateFN_CDesignerPage(IUnknown *);
extern AUTOMATIONOBJECTDESC DesignerObject;
extern IUnknown *CreateFN_CDesigner(IUnknown *);
//{END_FN()

BEGIN_OBJECTTABLE()
DEFINE_OBJECTTABLEENTRY(TYPE_PROPPAGE,&DesignerPageObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&DesignerObject,FALSE),
END_OBJECTTABLE()
static _se_translator_function s_seFunction = NULL;

//
// Globals static variable
//

static Globals* s_pGlobal = NULL;

//
// GetGlobals
//

Globals& GetGlobals()
{
	return *s_pGlobal;
}

#ifdef _DEBUG
int AllocHook(int nAllocType, void* pUserData, size_t size, int nBlockType, long nRequestNumber, const unsigned char* szFilename, int nLineNumber)
{
	switch (nAllocType)
	{
	case _HOOK_ALLOC:
		switch (nRequestNumber)
		{
		case 86:
		case 85:
		case 53:
		case 52:
			if (szFilename)
			{
				TRACE2(1, "_HOOK_ALLOC Size: %i Request Number: %i\n", size, nRequestNumber);
				TRACE2(1, "FileName: %s LineNumber: %i\n", szFilename, nLineNumber);
			}
			else
			{
				TRACE2(1, "_HOOK_ALLOC Size: %i Request Number: %i\n", size, nRequestNumber);
			}
			break;
		}
		break;
	}
	return TRUE;
}
#endif

void InitializeLibrary()
{
/*
#ifdef _DEBUG
	_CrtSetReportMode( _CRT_WARN, _CRTDBG_MODE_DEBUG);
	_CrtSetReportMode( _CRT_ERROR, _CRTDBG_MODE_DEBUG);

	int iFlags = _CrtSetDbgFlag(_CRTDBG_REPORT_FLAG);
	iFlags |= _CRTDBG_CHECK_ALWAYS_DF | _CRTDBG_DELAY_FREE_MEM_DF | _CRTDBG_LEAK_CHECK_DF;

	_CrtSetDbgFlag(iFlags);
	_CrtSetAllocHook(AllocHook);
#endif
*/
	s_seFunction = _set_se_translator( trans_func );
	s_pGlobal = new Globals;
	assert(s_pGlobal);
}

void UninitializeLibrary()
{
	_set_se_translator(s_seFunction);
	delete s_pGlobal;
	UIUtilities::CleanupUIUtilities();
/*
#ifdef _DEBUG
		TRACE(1, "---------------- UNINITIALIZE LIBRARY MEMORY CHECK -----------\n");
		_CrtCheckMemory();
		_CrtDumpMemoryLeaks();
		TRACE(1, "---------------- UNINITIALIZE LIBRARY WINDOW MAP CHECK -----------\n");
#endif
*/
}

#ifdef LICENSING_SUPPORT
const GUID IID_IClassFactory2={0xB196B28F,0xBAB4,0x101A,{0xB6,0x9C,0x00,0xAA,0x00,0x34,0x1D,0x07}};
#endif



LONG  g_cLocks=0; // ref count for LockServer
HANDLE g_hHeap=NULL;
HINSTANCE g_hInstance;
LCID g_lcidLocale=LOCALE_SYSTEM_DEFAULT;
CRITICAL_SECTION g_CriticalSection;

class HeapMng
{
public:
	HeapMng() 
	{
		if (!g_hHeap)
			g_hHeap=HeapCreate(0,0,0);
	};
	~HeapMng() 
	{
		if (g_hHeap)
		{
			HeapDestroy(g_hHeap);
			g_hHeap=NULL;
		}
	};
} heapMng;


BOOL g_fSysWinNT;
BOOL g_fSysWin95;
BOOL g_fSysWin95Shell;

//--------------------------------------------------------------------------
// DllMain
//--------------------------------------------------------------------------
extern "C"
BOOL WINAPI DllMain(HANDLE hInstance,DWORD dwReason,void *pvReserved)
{
    switch (dwReason) 
	{
      case DLL_PROCESS_ATTACH:
        {
		// Get OS/Version information
        DWORD dwVer = GetVersion();
        DWORD dwWinVer;

        //  swap the two lowest bytes of dwVer so that the major and minor version
        //  numbers are in a usable order.
        //  for dwWinVer: high byte = major version, low byte = minor version
        //     OS               Sys_WinVersion  (as of 5/2/95)
        //     =-------------=  =-------------=
        //     Win95            0x035F   (3.95)
        //     WinNT ProgMan    0x0333   (3.51)
        //     WinNT Win95 UI   0x0400   (4.00)
        //
        dwWinVer = (UINT)(((dwVer & 0xFF) << 8) | ((dwVer >> 8) & 0xFF));
        g_fSysWinNT = FALSE;
        g_fSysWin95 = FALSE;
        g_fSysWin95Shell = FALSE;

        if (dwVer < 0x80000000) {
            g_fSysWinNT = TRUE;
            g_fSysWin95Shell = (dwWinVer >= 0x0334);
        } else  {
            g_fSysWin95 = TRUE;
            g_fSysWin95Shell = TRUE;
        }
        // initialize a critical section for apartment threading support
        InitializeCriticalSection(&g_CriticalSection);
        // create an initial heap for everybody to use.
        // currently, we're going to let the system make things thread-safe,
        // which will make them a little slower, but hopefully not enough
        // to notice
        //
        g_hInstance = (HINSTANCE)hInstance;

        // give the user a chance to initialize whatever
        //
        InitializeLibrary();
        return TRUE;
        }

      // do  a little cleaning up!
      //
      case DLL_PROCESS_DETACH:
        // clean up our critical seciton
        DeleteCriticalSection(&g_CriticalSection);


        // give the user a chance to do some cleaning up
        UninitializeLibrary();
        return TRUE;
    }

    return TRUE;
}


//-------------------------------------------------------
// REGISTRATION HELPERS

#define GUID_STR_LEN 40

BOOL RegisterUnknownObject(UNKNOWNOBJECTDESC *objDesc)
{
	FRegKey classKey,objectKey,subKey,inprocKey;
	if (!classKey.Open(HKEY_CLASSES_ROOT,_T("CLSID")))
		return FALSE;
    TCHAR guidStr[GUID_STR_LEN];
#ifdef UNICODE
	if (!StringFromGUID2(*(objDesc->rclsid), guidStr,GUID_STR_LEN))
		return FALSE;
#else
	if (!StringFromGuidA(*(objDesc->rclsid), guidStr)) 
		return FALSE;
#endif
	objectKey.Create(classKey,guidStr);
	objectKey.SetValue(NULL,objDesc->pszObjectName);
	TCHAR dllPath[MAX_PATH];
	GetModuleFileName(g_hInstance,dllPath,MAX_PATH);
	inprocKey.Create(objectKey,_T("InprocServer32"));
	inprocKey.SetValue(NULL,dllPath);
	inprocKey.SetValue(_T("ThreadingModel"),_T("Apartment"));
	return TRUE;
}

BOOL UnregisterUnknownObject(UNKNOWNOBJECTDESC *objDesc)
{
	TCHAR guidStr[GUID_STR_LEN];
#ifdef UNICODE
	if (!StringFromGUID2(*(objDesc->rclsid), guidStr,GUID_STR_LEN))
		return FALSE;
#else
	if (!StringFromGuidA(*(objDesc->rclsid), guidStr)) 
		return FALSE;
#endif
	FRegKey clsidKey;
	clsidKey.Open(HKEY_CLASSES_ROOT,_T("CLSID"));
	clsidKey.DeleteSubKey(guidStr);
    return TRUE;
}

BOOL RegisterAutomationObject(AUTOMATIONOBJECTDESC *objDesc)
{
	if (RegisterUnknownObject(&(objDesc->unknownObj))==FALSE)
		return FALSE;
	FRegKey objKey,objVerKey;
	TCHAR fullName[128];
	TCHAR shortName[128];
	TCHAR versionStr[6];
	TCHAR guidStr[GUID_STR_LEN];

#ifdef UNICODE
	if (!StringFromGUID2(*(objDesc->unknownObj.rclsid), guidStr,GUID_STR_LEN))
		return FALSE;
#else
	if (!StringFromGuidA(*(objDesc->unknownObj.rclsid), guidStr)) 
		return FALSE;
#endif

	lstrcpy(shortName,LIBRARY_NAME);
	lstrcat(shortName,_T("."));
	lstrcat(shortName,objDesc->unknownObj.pszShortName);

    objKey.Create(HKEY_CLASSES_ROOT,shortName);
	// HKEY_CLASSES_ROOT\<LibraryName>.<ObjectName> = <ObjectName> Object
	objKey.SetValue(NULL,objDesc->unknownObj.pszObjectName);
	// HKEY_CLASSES_ROOT\<LibraryName>.<ObjectName>\CLSID = <CLSID>
	objKey.SetKeyValue(_T("CLSID"),guidStr);
    
	wsprintf(versionStr,_T("%d"),objDesc->versionMajor);
	lstrcpy(fullName,shortName);
	lstrcat(fullName,_T("."));
	lstrcat(fullName,versionStr);
	// HKEY_CLASSES_ROOT\<LibraryName>.<ObjectName>\CurVer = <ObjectName>.Object.<VersionNumber>
	objKey.SetKeyValue(_T("CurVer"),fullName);


	objVerKey.Create(HKEY_CLASSES_ROOT,fullName);
	// HKEY_CLASSES_ROOT\<LibraryName>.<ObjectName>.<VersionNumber> = <ObjectName> Object
	objVerKey.SetValue(NULL,objDesc->unknownObj.pszObjectName);
	// HKEY_CLASSES_ROOT\<LibraryName>.<ObjectName>.<VersionNumber>\CLSID = <CLSID>
	objVerKey.SetKeyValue(_T("CLSID"),guidStr);
    
	FRegKey classesKey,clsidKey;
	if (!classesKey.Open(HKEY_CLASSES_ROOT,_T("CLSID")))
		return FALSE;
	clsidKey.Create(classesKey,guidStr);
    // HKEY_CLASSES_ROOT\CLSID\<CLSID>\ProgID = <LibraryName>.<ObjectName>.<VersionNumber>
	clsidKey.SetKeyValue(_T("ProgID"),fullName);
    // HKEY_CLASSES_ROOT\CLSID\<CLSID>\VersionIndependentProgID = <LibraryName>.<ObjectName>
	clsidKey.SetKeyValue(_T("VersionIndependentProgID"),shortName);
    // HKEY_CLASSES_ROOT\CLSID\<CLSID>\TypeLib = <LibidOfTypeLibrary>
#ifdef UNICODE
	if (!StringFromGUID2(LIBID_PROJECT, guidStr,GUID_STR_LEN))
		return FALSE;
#else
	if (!StringFromGuidA(LIBID_PROJECT, guidStr)) 
		return FALSE;
#endif
	clsidKey.SetKeyValue(_T("TypeLib"),guidStr);
	return TRUE;
}

BOOL RegisterControlObject(CONTROLOBJECTDESC *objectDesc)
{
	if (!RegisterAutomationObject(&(objectDesc->automationDesc)))
		return FALSE;
	// create Control entry
	FRegKey classKey,objectKey,subKey,miscStatKey;
	if (!classKey.Open(HKEY_CLASSES_ROOT,_T("CLSID")))
		return FALSE;
    TCHAR guidStr[GUID_STR_LEN];
	TCHAR szTmp[MAX_PATH];

#ifdef UNICODE
	if (!StringFromGUID2(*(objectDesc->automationDesc.unknownObj.rclsid), guidStr,GUID_STR_LEN))
		return FALSE;
#else
	if (!StringFromGuidA(*(objectDesc->automationDesc.unknownObj.rclsid), guidStr)) 
		return FALSE;
#endif
	
	objectKey.Create(classKey,guidStr);
	subKey.Create(objectKey,_T("Control"));	
	// MiscStatus
	miscStatKey.Create(objectKey,_T("MiscStatus"));
	miscStatKey.SetValue(NULL,_T("0"));
	subKey.Create(miscStatKey,_T("1"));
	wsprintf(szTmp, _T("%d"), objectDesc->dwOleMiscFlags);
	subKey.SetValue(NULL,szTmp);
	// Toolbox bitmap
	GetModuleFileName(g_hInstance, szTmp, MAX_PATH);
    wsprintf(guidStr, _T(", %d"), objectDesc->wToolboxId);
    lstrcat(szTmp, guidStr);
	objectKey.SetKeyValue(_T("ToolboxBitmap32"),szTmp);
	
	// Version info
	wsprintf(szTmp,_T("%d.%d"),objectDesc->automationDesc.versionMajor,objectDesc->automationDesc.versionMinor);
	objectKey.SetKeyValue(_T("Version"),szTmp);

	return TRUE;
}

BOOL UnregisterAutomationObject(AUTOMATIONOBJECTDESC *objDesc)
{
		FRegKey key;
	TCHAR temp[128];
	TCHAR versionStr[10];
	lstrcpy(temp,LIBRARY_NAME);
	lstrcat(temp,_T("."));
	lstrcat(temp,objDesc->unknownObj.pszShortName);
    //   HKEY_CLASSES_ROOT\<LibraryName>.<ObjectName> [\] *
	RecurseSubKeyDelete(HKEY_CLASSES_ROOT,temp);
	//   HKEY_CLASSES_ROOT\<LibraryName>.<ObjectName>.<VersionNumber> [\] *
	wsprintf(versionStr,_T("%d"),objDesc->versionMajor);
	lstrcat(temp,_T("."));
	lstrcat(temp,versionStr);
	RecurseSubKeyDelete(HKEY_CLASSES_ROOT,temp);
	return UnregisterUnknownObject(&(objDesc->unknownObj));
}



//=--------------------------------------------------------------------------=
// DllRegisterServer
//=--------------------------------------------------------------------------=
// registers the Automation server
//
// Output:
//    HRESULT
//
// Notes:
//
STDAPI DllRegisterServer(void)
{
//{BEGIN OBJECTREGISTER}
if (!RegisterUnknownObject((UNKNOWNOBJECTDESC *)&DesignerPageObject))
	return E_FAIL;
if (!RegisterAutomationObject(&DesignerObject))
	return E_FAIL;
//{END OBJECTREGISTER}
	ITypeLib *pTypeLib;
    HRESULT hr;
    WCHAR	wszTmp[MAX_PATH];
	// Register type library

	
#ifdef UNICODE
	GetModuleFileName(g_hInstance, wszTmp, MAX_PATH);
#else
	DWORD   dwPathLen;
	char    szTmp[MAX_PATH];
	dwPathLen = GetModuleFileName(g_hInstance, szTmp, MAX_PATH);
	MultiByteToWideChar(CP_ACP, 0, szTmp, -1, (LPWSTR)wszTmp, MAX_PATH*sizeof(WCHAR));
#endif

    hr = LoadTypeLib(wszTmp, &pTypeLib);
    RETURN_ON_FAILURE(hr);
    hr = RegisterTypeLib(pTypeLib, wszTmp, NULL);
    pTypeLib->Release();
    RETURN_ON_FAILURE(hr);
	return S_OK;
}



//=--------------------------------------------------------------------------=
// DllUnregisterServer
//=--------------------------------------------------------------------------=
// unregister's the Automation server

STDAPI DllUnregisterServer(void)
{
//{BEGIN OBJECTUNREGISTER}
	UnregisterUnknownObject((UNKNOWNOBJECTDESC *)&DesignerPageObject);
	UnregisterAutomationObject(&DesignerObject);
//{END OBJECTUNREGISTER}
	
//{BEGIN UNREGISTER TYPELIB}
	FRegKey typeLibKey;
	typeLibKey.Open(HKEY_CLASSES_ROOT,_T("TypeLib"));
	OLECHAR wKeyName[39];
	StringFromGUID2(LIBID_PROJECT,wKeyName,39);
	#ifndef UNICODE
	TCHAR keyName[39];
	WideCharToMultiByte(CP_ACP,0,wKeyName,39,keyName,39,NULL,NULL);
	typeLibKey.DeleteSubKey(keyName);
	#else
	typeLibKey.DeleteSubKey(wKeyName);
	#endif
//{END UNREGISTER TYPELIB}	
	return S_OK;
}


//=--------------------------------------------------------------------------=
// DllCanUnloadNow
//=--------------------------------------------------------------------------=
// we are being asked whether or not it's okay to unload the DLL.  just check
// the lock counts on remaining objects ...
//
// Output:
//    HRESULT        - S_OK, can unload now, S_FALSE, can't.
//
// Notes:
//
STDAPI DllCanUnloadNow(void)
{
    // if there are any objects lying around, then we can't unload.  The
    // controlling CUnknownObject class that people should be inheriting from
    // takes care of this
    return (g_cLocks) ? S_FALSE : S_OK;
}


//=--------------------------------------------------------------------------=
// DllGetClassObject
//=--------------------------------------------------------------------------=
int IndexOfOleObject(REFCLSID rclsid)
{
    int x = 0;
	UNKNOWNOBJECTDESC *descPtr;
	while (x<(sizeof(g_objectTable)/sizeof(g_objectTable[0]))) 
	{
		descPtr=(UNKNOWNOBJECTDESC *)(g_objectTable[x].objectDefPtr);
		if (descPtr->rclsid!=NULL && (*(descPtr->rclsid))==rclsid)
			return x;
        x++;
    }
    return -1;
}

ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef)
{
	HRESULT hr;
	AUTOMATIONOBJECTDESC *descPtr=(AUTOMATIONOBJECTDESC *)objectDef;
    if (descPtr->pTypeInfo==NULL)
	{ //Try to load typelibrary
		ITypeLib *pTypeLib;
        ITypeInfo *pTypeInfoTmp;
        hr = LoadRegTypeLib(LIBID_PROJECT, descPtr->versionMajor, descPtr->versionMinor,
                            lcid, &pTypeLib);
		if (FAILED(hr))
		{
			// might try to load tlb. insert code as needed
			return NULL;
		}

		hr = pTypeLib->GetTypeInfoOfGuid(*(descPtr->riid), &pTypeInfoTmp);
        descPtr->pTypeInfo=pTypeInfoTmp;
		pTypeLib->Release();
	}

	if (descPtr->pTypeInfo==NULL)
		return NULL;
	descPtr->pTypeInfo->AddRef();
	return descPtr->pTypeInfo;

}

STDAPI DllGetClassObject(REFCLSID rclsid,REFIID riid,void **ppvObjOut)
{
    HRESULT hr;
    void   *pv;
    int     iIndex;
    // arg checking
    if (!ppvObjOut)
        return E_INVALIDARG;
    // first of all, make sure they're asking for something we work with.
    iIndex = IndexOfOleObject(rclsid);
    if (iIndex == -1)
        return CLASS_E_CLASSNOTAVAILABLE;
    // create the blank object.
    pv = (void *)new CObjectFactory(iIndex);
    if (!pv)
        return E_OUTOFMEMORY;
    // QI for whatever the user has asked for.
    hr = ((IUnknown *)pv)->QueryInterface(riid, ppvObjOut);
    ((IUnknown *)pv)->Release();
    return hr;
}

/////////////// CLASS FACTORY STUFF
CObjectFactory::CObjectFactory(int iIndex)
{
	index=iIndex;
	m_refCount=1;
}

CObjectFactory::~CObjectFactory()
{
}

STDMETHODIMP CObjectFactory::QueryInterface(REFIID riid, void **ppvObj)
{
	if (riid==IID_IUnknown || riid==IID_IClassFactory 
#ifdef LICENSING_SUPPORT
		||
		(riid==IID_IClassFactory2 && g_objectTable[index].licensed)
#endif
		)
	{
		AddRef();
		*ppvObj=this;
		return NOERROR;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CObjectFactory::AddRef()
{
	return ++m_refCount;
}

STDMETHODIMP_(ULONG) CObjectFactory::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}

// IClassFactory methods
HRESULT CreateOleObjectFromIndex(IUnknown *pUnkOuter,int iIndex,void **ppvObjOut,REFIID riid)
{
    IUnknown *pUnk = NULL;
    HRESULT   hr;

    // go and create the object
    pUnk = CREATEFNOFOBJECT(iIndex)(pUnkOuter);

    // sanity check and make sure the object actually got allocated.
    RETURN_ON_NULLALLOC(pUnk);

    // make sure we support aggregation here properly -- if they gave us
    // a controlling unknown, then they -must- ask for IUnknown, and we'll
    // give them the private unknown the object gave us.
    if (pUnkOuter) 
	{
        if (!(riid==IID_IUnknown)) 
		{
            pUnk->Release();
            return E_INVALIDARG;
        }
        *ppvObjOut = (void *)pUnk;
        hr = S_OK;
    } 
	else 
	{
        // QI for whatever the user wants.
        //
        hr = pUnk->QueryInterface(riid, ppvObjOut);
        pUnk->Release();
        RETURN_ON_FAILURE(hr);
    }
    return hr;
}

STDMETHODIMP CObjectFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObjOut)
{
    if (!ppvObjOut)
        return E_INVALIDARG;
#ifdef LICENSING_SUPPORT
    EnterCriticalSection(&g_CriticalSection);
    if (!g_fCheckedForLicense) {
        g_fMachineHasLicense = CheckForLicense();
        g_fCheckedForLicense = TRUE;
    }
    LeaveCriticalSection(&g_CriticalSection);
    // check to see if they have the appropriate license to create this stuff
    if (!g_fMachineHasLicense)
        return CLASS_E_NOTLICENSED;
#endif
    return CreateOleObjectFromIndex(pUnkOuter, index, ppvObjOut, riid);
}

STDMETHODIMP CObjectFactory::LockServer(BOOL fLock)
{
	if (fLock)
		++g_cLocks;
	else
		--g_cLocks;
	return NOERROR;
}

#ifdef LICENSING_SUPPORT
// IClassFactory2 methods
STDMETHODIMP CObjectFactory::GetLicInfo(LICINFO *pLicInfo)
{
	return E_FAIL;
}

STDMETHODIMP CObjectFactory::RequestLicKey(DWORD dwReserved, BSTR *pbstrKey)
{
	return E_FAIL;
}

STDMETHODIMP CObjectFactory::CreateInstanceLic(IUnknown *pUnkOuter, IUnknown *pUnkReserved, REFIID riid, BSTR bstrKey, void **ppvObjOut)
{
	return E_FAIL;
}

#endif





/// MINIMIZE CRT

/*
#ifndef _DEBUG

int __cdecl _purecall()
{
	DebugBreak();
	return 0;
}

extern "C" const int _fltused = 0;

void* __cdecl malloc(size_t n)
{
	if (g_hHeap== NULL)
	{
		g_hHeap= HeapCreate(0, 0, 0);
		if (g_hHeap== NULL)
			return NULL;
	}


#ifdef _MALLOC_ZEROINIT
	void* p = HeapAlloc(g_hHeap, 0, n);
	if (p != NULL)
		memset(p, 0, n);
	return p;
#else
	return HeapAlloc(g_hHeap, 0, n);
#endif
}

void* __cdecl calloc(size_t n, size_t s)
{
#ifdef _MALLOC_ZEROINIT
	return malloc(n * s);
#else
	void* p = malloc(n * s);
	if (p != NULL)
		memset(p, 0, n * s);
	return p;
#endif
}

void* __cdecl realloc(void* p, size_t n)
{
	return (p == NULL) ? malloc(n) : HeapReAlloc(g_hHeap, 0, p, n);
}

void __cdecl free(void* p)
{
	if (p != NULL)
		HeapFree(g_hHeap, 0, p);
}

void* __cdecl operator new(size_t n)
{
	return malloc(n);
}

void __cdecl operator delete(void* p)
{
	free(p);
}

#endif  //_DEBUG

*/
