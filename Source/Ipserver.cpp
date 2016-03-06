//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//
//
//
//	RegisterCLSIDInCategory(*ActiveBar2Object.automationDesc.unknownObj.rclsid,CATID_SafeForScripting);
//	RegisterCLSIDInCategory(*ActiveBar2Object.automationDesc.unknownObj.rclsid,CATID_SafeForInitializing);
//	RegisterCLSIDInCategory(*ActiveBar2Object.automationDesc.unknownObj.rclsid,CATID_Control);
//
//

#include "precomp.h"
#include <olectl.h>
#include <assert.h>
#include <crtdbg.h>
#include "fregkey.h"
#include "support.h"
#include "resource.h"
#include <ObjSafe.h>
#include "ipserver.h"
#include <time.h>

//{BEGIN AUTOINCLUDES}
#include "catidreg.h"
//{END AUTOINCLUDES}

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//{BEGIN_FN()
extern CONTROLOBJECTDESC ActiveBar2Object;
extern IUnknown *CreateFN_CBar(IUnknown *);
extern AUTOMATIONOBJECTDESC BandsObject;
extern IUnknown *CreateFN_CBands(IUnknown *);
extern AUTOMATIONOBJECTDESC BandObject;
extern IUnknown *CreateFN_CBand(IUnknown *);
extern AUTOMATIONOBJECTDESC ToolObject;
extern IUnknown *CreateFN_CTool(IUnknown *);
extern AUTOMATIONOBJECTDESC ToolsObject;
extern IUnknown *CreateFN_CTools(IUnknown *);
extern AUTOMATIONOBJECTDESC ReturnStringObject;
extern IUnknown *CreateFN_CReturnString(IUnknown *);
extern AUTOMATIONOBJECTDESC ChildBandsObject;
extern IUnknown *CreateFN_CChildBands(IUnknown *);
extern AUTOMATIONOBJECTDESC ComboListObject;
extern IUnknown *CreateFN_CBList(IUnknown *);
extern AUTOMATIONOBJECTDESC ReturnBoolObject;
extern IUnknown *CreateFN_CRetBool(IUnknown *);
extern CONTROLOBJECTDESC CustomizeListboxObject;
extern IUnknown *CreateFN_CCustomizeListbox(IUnknown *);
extern UNKNOWNOBJECTDESC ImageMgrObject;
extern IUnknown *CreateFN_CImageMgr(IUnknown *);
extern AUTOMATIONOBJECTDESC ShortCutObject;
extern IUnknown *CreateFN_ShortCutStore(IUnknown *);
//{END_FN()

BEGIN_OBJECTTABLE()
DEFINE_OBJECTTABLEENTRY(TYPE_CONTROL,&ActiveBar2Object,TRUE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&BandsObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&BandObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&ToolObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&ToolsObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&ReturnStringObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&ChildBandsObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&ComboListObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&ReturnBoolObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_CONTROL,&CustomizeListboxObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_UNKNOWN,&ImageMgrObject,FALSE),
DEFINE_OBJECTTABLEENTRY(TYPE_AUTOMATION,&ShortCutObject,FALSE),
END_OBJECTTABLE()
void InitGlobalHeap()
{
	g_hHeap=HeapCreate(0,0,0);
}

void CleanupGlobalHeap()
{
	if (g_hHeap)
	{
		HeapDestroy(g_hHeap);
		g_hHeap=NULL;
	}
}
/*
#ifdef LICENSING_SUPPORT
const GUID IID_IClassFactory2={0xB196B28F,0xBAB4,0x101A,{0xB6,0x9C,0x00,0xAA,0x00,0x34,0x1D,0x07}};
#endif
*/
//BOOL g_bDoCheckExpire;

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
BOOL g_fSysWin98Shell;
extern void InitializeLibrary();
extern void UninitializeLibrary();
void InitHeap()
{
	if (!g_hHeap)
		g_hHeap = HeapCreate(0,0,0);
}

void CleanupHeap()
{
	if (g_hHeap)
	{
		HeapDestroy(g_hHeap);
		g_hHeap = NULL;
	}
}

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
			OSVERSIONINFO osvi;
			osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
			g_fSysWinNT = FALSE;
			g_fSysWin95 = FALSE;
			g_fSysWin95Shell = FALSE;
			g_fSysWin98Shell = FALSE;

			BOOL bResult = GetVersionEx(&osvi);
			assert(bResult);
			if (bResult)
			{
				switch (osvi.dwPlatformId)
				{
				case VER_PLATFORM_WIN32_WINDOWS:
					g_fSysWin95 = TRUE;
					if (osvi.dwMajorVersion > 3)
					{
						g_fSysWin95Shell = TRUE;
						g_fSysWin98Shell = (osvi.dwMajorVersion > 4 || ((osvi.dwMajorVersion == 4) && (osvi.dwMinorVersion > 0)));
					}
					break;

				case VER_PLATFORM_WIN32_NT:
					g_fSysWinNT = TRUE;
					if (osvi.dwMajorVersion > 3)
					{
						g_fSysWin95Shell = TRUE;
						if (osvi.dwMajorVersion > 4)
							g_fSysWin98Shell = TRUE;
					}
					break;
				}
			}
			else
			{
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
				//     WinNT Win98 UI   0x0500   (5.00)
				//
				dwWinVer = (UINT)(((dwVer & 0xFF) << 8) | ((dwVer >> 8) & 0xFF));
				g_fSysWinNT = FALSE;
				g_fSysWin95 = FALSE;
				g_fSysWin95Shell = FALSE;

				if (dwVer < 0x80000000) 
				{
					g_fSysWinNT = TRUE;
					g_fSysWin95Shell = (dwWinVer >= 0x0334);
					g_fSysWin98Shell = (dwWinVer >= 0x0500);
				} 
				else  
				{
					g_fSysWin95 = TRUE;
					g_fSysWin95Shell = TRUE;
				}
			}

			// initialize a critical section for apartment threading support
			InitializeCriticalSection(&g_CriticalSection);
			// create an initial heap for everybody to use.
			// currently, we're going to let the system make things thread-safe,
			// which will make them a little slower, but hopefully not enough
			// to notice
			//
			g_hInstance = (HINSTANCE)hInstance;

			//
			// give the user a chance to initialize whatever
			//
//			g_bDoCheckExpire=TRUE;
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

#ifdef AR20
static HRESULT DllDontRegisterServer(void)
#else
STDAPI DllRegisterServer(void)
#endif

{
	RegisterCLSIDInCategory(*ActiveBar2Object.automationDesc.unknownObj.rclsid,CATID_SafeForScripting);
	RegisterCLSIDInCategory(*ActiveBar2Object.automationDesc.unknownObj.rclsid,CATID_SafeForInitializing);
	RegisterCLSIDInCategory(*ActiveBar2Object.automationDesc.unknownObj.rclsid,CATID_Control);
//{BEGIN OBJECTREGISTER}
if (!RegisterControlObject(&ActiveBar2Object))
	return E_FAIL;
if (!RegisterAutomationObject(&BandsObject))
	return E_FAIL;
if (!RegisterAutomationObject(&BandObject))
	return E_FAIL;
if (!RegisterAutomationObject(&ToolObject))
	return E_FAIL;
if (!RegisterAutomationObject(&ToolsObject))
	return E_FAIL;
if (!RegisterAutomationObject(&ReturnStringObject))
	return E_FAIL;
if (!RegisterAutomationObject(&ChildBandsObject))
	return E_FAIL;
if (!RegisterAutomationObject(&ComboListObject))
	return E_FAIL;
if (!RegisterAutomationObject(&ReturnBoolObject))
	return E_FAIL;
if (!RegisterControlObject(&CustomizeListboxObject))
	return E_FAIL;
if (!RegisterUnknownObject(&ImageMgrObject))
	return E_FAIL;
if (!RegisterAutomationObject(&ShortCutObject))
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

#ifdef AR20
static HRESULT DllDontUnregisterServer(void)
#else
STDAPI DllUnregisterServer(void)
#endif
{
//{BEGIN OBJECTUNREGISTER}
	UnregisterControlObject(ActiveBar2Object);
	UnregisterAutomationObject(&BandsObject);
	UnregisterAutomationObject(&BandObject);
	UnregisterAutomationObject(&ToolObject);
	UnregisterAutomationObject(&ToolsObject);
	UnregisterAutomationObject(&ReturnStringObject);
	UnregisterAutomationObject(&ChildBandsObject);
	UnregisterAutomationObject(&ComboListObject);
	UnregisterAutomationObject(&ReturnBoolObject);
	UnregisterControlObject(CustomizeListboxObject);
	UnregisterUnknownObject(&ImageMgrObject);
	UnregisterAutomationObject(&ShortCutObject);
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

void UnLoadTypeInfo()
{
	for (int cnt = 0; cnt < (sizeof(g_objectTable)/sizeof(g_objectTable[0])); ++cnt)
	{
		if (g_objectTable[cnt].type == TYPE_AUTOMATION || g_objectTable[cnt].type == TYPE_CONTROL)
		{
			AUTOMATIONOBJECTDESC *descPtr=(AUTOMATIONOBJECTDESC *)(g_objectTable[cnt].objectDefPtr);
			if (descPtr->pTypeInfo)
			{
				descPtr->pTypeInfo->Release();
				descPtr->pTypeInfo=NULL;
			}
		}
	}
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

/*	if (g_bDoCheckExpire)
	{
		#define ExpireYear 1999
		#define ExpireMonth 12
		time_t curTime;
		time(&curTime);
		struct tm *lt;
		lt=localtime(&curTime);
		
		if ((lt->tm_year+1900)>ExpireYear || ((lt->tm_year+1900)==ExpireYear && (lt->tm_mon)>=ExpireMonth))
		{
			::MessageBox(::GetFocus(),"This beta copy has expired","ActiveBar 2.0",MB_ICONSTOP);
			return E_FAIL;
		}
		g_bDoCheckExpire=FALSE;
	}
*/	

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
		(riid==IID_IClassFactory2)
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

BOOL g_fMachineHasLicense=FALSE;
BOOL g_fCheckedForLicense=FALSE;

WCHAR runtimeLicenseKey[]=L"&$4525$^$$%  ( & RLK AB2.0 Copyright Data Dynamics";
char runtimeLicenseKey2[]=":)giyukolmmnytooiupztuffee";
	// {188E10CB-AE75-4EAE-9B9D-EA57CF3D57E5}
DEFINE_GUID(IID_ActiveBarValue, 0x188e10cb, 0xae75, 0x4eae, 0x9b, 0x9d, 0xea, 0x57, 0xcf, 0x3d, 0x57, 0xe5);
					   
BOOL IsLicenseValid()
{
	if (g_fCheckedForLicense)
		return g_fMachineHasLicense;
	else
	{
		g_fMachineHasLicense=FALSE;
		g_fCheckedForLicense=TRUE;
		// Check for design time license on machine
		FRegKey licensesKey;
		if (licensesKey.Open(HKEY_CLASSES_ROOT,_T("Licenses")))
		{
			char buffer[255];
			long cbBuffer=255;
			if (RegQueryValue(licensesKey,"E0EF4837-4BE5-4CC9-AA9A-554AA8D4816A",buffer,&cbBuffer)==ERROR_SUCCESS)
			{
				WCHAR szValue[40];
				if (StringFromGUID2(IID_ActiveBarValue, (WCHAR*)szValue, 40) > 0)
				{
					MAKE_WIDEPTR_FROMTCHAR(wValue, buffer);
					if (0 == wcscmp(wValue, szValue))
					{
						g_fMachineHasLicense=TRUE;
						return TRUE;
					}
				}
			}
		}
	}
	return FALSE;
}

BOOL VerifyLicenseKey(BSTR licKey)
{
	if (wcscmp(licKey,runtimeLicenseKey)==0)
	{
		g_fMachineHasLicense=TRUE;
		return TRUE;
	}
	else
		return FALSE;
}

BOOL GetLicenseKey(DWORD dwReserved, BSTR *bstr)
{
	if (IsLicenseValid())
	{
		*bstr=SysAllocString(runtimeLicenseKey);
		return TRUE;
	}
	return FALSE;
}


STDMETHODIMP CObjectFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObjOut)
{
    if (!ppvObjOut)
        return E_INVALIDARG;

	if (index==0)
	{
	#ifdef LICENSING_SUPPORT
	//    EnterCriticalSection(&g_CriticalSection);
	    if (!g_fCheckedForLicense && !g_fMachineHasLicense)  
		{
			g_fMachineHasLicense = IsLicenseValid();
			g_fCheckedForLicense = TRUE;
		}
	//    LeaveCriticalSection(&g_CriticalSection);
	    // check to see if they have the appropriate license to create this stuff
	    //if (!g_fMachineHasLicense)
	    //    return CLASS_E_NOTLICENSED;
	#endif
	}

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
	BSTR bstr = NULL;
	pLicInfo->fLicVerified = IsLicenseValid();
	pLicInfo->fRuntimeKeyAvail = GetLicenseKey(0, &bstr);
	if (bstr != NULL)
		SysFreeString(bstr);
	return S_OK;
}

STDMETHODIMP CObjectFactory::RequestLicKey(DWORD dwReserved, BSTR *pbstrKey)
{
	*pbstrKey = NULL;
	if (IsLicenseValid())
	{
		if (GetLicenseKey(dwReserved, pbstrKey))
			return S_OK;
		else
			return E_FAIL;
	}
	else
		return CLASS_E_NOTLICENSED;
}

STDMETHODIMP CObjectFactory::CreateInstanceLic(IUnknown *pUnkOuter, IUnknown *pUnkReserved, REFIID riid, BSTR bstrKey, void **ppvObjOut)
{
	*ppvObjOut= NULL;
	if (((bstrKey != NULL) && !VerifyLicenseKey(bstrKey)) ||
		((bstrKey == NULL) && !IsLicenseValid()))
	{
		;//return CLASS_E_NOTLICENSED;   // These are commented out since we'll run with licnotvaliddialog
	}
	else
	{
		g_fMachineHasLicense=TRUE;
	}
	// outer objects must ask for IUnknown only
	assert(pUnkOuter == NULL || riid == IID_IUnknown);

	return CreateOleObjectFromIndex(pUnkOuter, index, ppvObjOut, riid);
}

#endif




/*
/// MINIMIZE CRT


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