#ifndef __IPSERVER_H__
#define __IPSERVER_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "debug.h"

//{PREPROCESSOR DEFS}
#define LICENSING_SUPPORT
//{PREPROCESSOR DEFS}
//{LIBRARY INFO}
#define LIBID_PROJECT LIBID_ActiveBar2Library
#define LIBRARY_NAME "ActiveBar2Library"
//{LIBRARY INFO}
































































































































































































































































































































































































































































































//=--------------------------------------------------------------------------=
// Useful macros
//=--------------------------------------------------------------------------=
//
// handy error macros, randing from cleaning up, to returning to clearing
// rich error information as well.
//
#define RETURN_ON_FAILURE(hr) if (FAILED(hr)) return hr
#define RETURN_ON_NULLALLOC(ptr) if (!(ptr)) return E_OUTOFMEMORY
#define CLEANUP_ON_FAILURE(hr) if (FAILED(hr)) goto CleanUp
#define RELEASE_OBJECT(ptr)    if (ptr) { IUnknown *pUnk = (ptr); (ptr) = NULL; pUnk->Release(); }
#define QUICK_RELEASE(ptr)     if (ptr) (ptr)->Release();

#define DO_GUIDS_MATCH(riid1, riid2) ((riid1.Data1 == riid2.Data1) && (riid1 == riid2))

#define EVENTREFIIDOFCONTROL(x) (((CONTROLOBJECTDESC *)(objectDef))->piidEvents)
#define PROPPAGECOUNT(x) (((CONTROLOBJECTDESC *)(objectDef))->cPropPages)
#define PROPPAGES(x) (((CONTROLOBJECTDESC *)(objectDef))->rgPropPageGuids)
#define CLSIDOFOBJECT(index) (*(((UNKNOWNOBJECTDESC *)(objectDef))->rclsid))
#define INTERFACEOFOBJECT(index) (*(((AUTOMATIONOBJECTDESC *)(objectDef))->riid))
#define HELPFILEOFOBJECT(index) (((AUTOMATIONOBJECTDESC *)(objectDef))->pszHelpFile)

extern HANDLE g_hHeap;
extern HINSTANCE g_hInstance;
extern LONG g_cLocks;
extern LCID g_lcidLocale;

typedef struct tagUNKNOWNOBJECTDESC
{
    const CLSID *rclsid;                    // CLSID of your object.      ONLY USE IF YOU'RE CoCreatable!
	LPCTSTR	pszShortName;
    LPCTSTR       pszObjectName;             // Name of your object.       ONLY USE IF YOU'RE CoCreatable!
    IUnknown    *(*pfnCreate)(IUnknown *);  // pointer to creation fn.    ONLY USE IF YOU'RE CoCreatable!
} UNKNOWNOBJECTDESC;

typedef struct tagAUTOMATIONOBJECTDESC
{
	UNKNOWNOBJECTDESC unknownObj;
	unsigned short versionMajor;                       // Version number of Object.  ONLY USE IF YOU'RE CoCreatable!
	unsigned short versionMinor;
    const IID   *riid;                           // object's type
    LPCTSTR       pszHelpFile;                    // the helpfile for this automation object.
    ITypeInfo   *pTypeInfo;                      // typeinfo for this object
    UINT         cTypeInfo;                      // number of refs to the type info
} AUTOMATIONOBJECTDESC;

typedef struct {
    AUTOMATIONOBJECTDESC automationDesc;
    const IID      *piidEvents;                    // IID of primary event interface
    DWORD           dwOleMiscFlags;                // control flags
    DWORD           dwActivationPolicy;            // IPointerInactive support
    VARIANT_BOOL    fOpaque;                       // is your control 100% opaque?
    VARIANT_BOOL    fWindowless;                   // do we do windowless if we can?
    WORD            wToolboxId;                    // resource ID of Toolbox Bitmap
    LPCTSTR          szWndClass;                    // name of window control class
    WORD            cPropPages;                    // number of property pages
    const GUID    **rgPropPageGuids;               // array of the property page GUIDs
    WORD            cCustomVerbs;                  // number of custom verbs
    const OLEVERB *rgCustomVerbs;                 // description of custom verbs
    WNDPROC         pfnSubClass;                   // for subclassed controls.
} CONTROLOBJECTDESC;


#define NAMEOFOBJECT(x) (((UNKNOWNOBJECTDESC *)objectDef)->pszObjectName)

//=--------------------------------------------------------------------------=
// structure that control writers will use to define property pages.
typedef struct tagPROPERTYPAGEINFO {
    UNKNOWNOBJECTDESC unknowninfo;
    WORD    wDlgResourceId;
    WORD    wTitleId;
    WORD    wDocStringId;
    LPCTSTR  szHelpFile;
    DWORD   dwHelpContextId;
} PROPERTYPAGEINFO;

#define DEFINE_PROPPAGE(clsid, objname,longName,fn,wDlgResourceId,wTitleId,wDocStringId,szHelpFile,dwHelpContextId) \
    PROPERTYPAGEINFO objname##Object = { { clsid,_T( #objname ),longName, fn }, wDlgResourceId,wTitleId,wDocStringId,szHelpFile,dwHelpContextId}

#define TEMPLATENAMEOFPROPPAGE(index)    MAKEINTRESOURCE(((PROPERTYPAGEINFO *)(m_pDesc))->wDlgResourceId)
#define TITLEIDOFPROPPAGE(index)         (((PROPERTYPAGEINFO *)(m_pDesc))->wTitleId)
#define DOCSTRINGIDOFPROPPAGE(index)     (((PROPERTYPAGEINFO *)(m_pDesc))->wDocStringId)
#define HELPCONTEXTOFPROPPAGE(index)     (((PROPERTYPAGEINFO *)(m_pDesc))->dwHelpContextId)
#define HELPFILEOFPROPPAGE(index)        (((PROPERTYPAGEINFO *)(m_pDesc))->szHelpFile)
#define SUBCLASSWNDPROCOFCONTROL(index)  (((CONTROLOBJECTDESC *)(m_pDesc))->pfnSubClass)
#define ACTIVATIONPOLICYOFCONTROL(index) ((CONTROLOBJECTDESC *)(m_pDesc->pInfo))->dwActivationPolicy
//=--------------------------------------------------------------------------=



#define CTLIVERB_PROPERTIES     1 // defined for hosts who dont handle the IOLEVERB_PROPERTIES verb correctly

#define CREATEFNOFOBJECT(index)   (((UNKNOWNOBJECTDESC *)(g_objectTable[(index)]).objectDefPtr)->pfnCreate)
#define DEFINE_UNKNOWNOBJECT(clsid, objname,longName,fn) \
UNKNOWNOBJECTDESC objname##Object = { clsid,_T( #objname ),longName, fn }
#define DEFINE_AUTOMATIONOBJECT(clsid, objname,longName,fn,versionMajor,versionMinor,iid,helpfile) \
AUTOMATIONOBJECTDESC objname##Object = { {clsid,_T( #objname ), longName, fn} ,versionMajor,versionMinor,iid,helpfile,NULL,0}
#define DEFINE_CONTROLOBJECT(clsid, objname,longName,fn,versionMajor,versionMinor,iid,helpfile,\
	piidEvents,dwOleMiscFlags,dwActivationPolicy,fOpaque,fWindowless,\
	wToolboxId,cPropPages,rgPropPageGuids,cCustomVerbs,rgCustomVerbs)\
CONTROLOBJECTDESC objname##Object =\
	{{ {clsid,_T( #objname ), longName, fn} ,versionMajor,versionMinor,iid,helpfile,NULL,0}, \
	piidEvents,dwOleMiscFlags,dwActivationPolicy,fOpaque,fWindowless,\
    wToolboxId,NULL/*WndClass*/,cPropPages,rgPropPageGuids,\
	cCustomVerbs,rgCustomVerbs,NULL};
	
typedef struct tagOBJECTTABLEENTRY
{
	int type;
	void *objectDefPtr;
	BOOL isLicensed;
} OBJECTTABLEENTRY;


#define TYPE_UNKNOWN 0   // object table entry type
#define TYPE_AUTOMATION 1
#define TYPE_CONTROL 2
#define TYPE_PROPPAGE 3

extern OBJECTTABLEENTRY g_objectTable[];

#define BEGIN_OBJECTTABLE() \
OBJECTTABLEENTRY g_objectTable[]={
#define END_OBJECTTABLE()\
};

#define DEFINE_OBJECTTABLEENTRY(type,objPtr,islic)\
{type,objPtr,islic}

//=--------------------------------------------------------------------------=
// this structure is like the OLEVERB structure, except that it has a resource ID
// instead of a string for the verb's name.  better support for localization.
//
typedef struct tagVERBINFO {
    LONG    lVerb;                // verb id
    ULONG   idVerbName;           // resource ID of verb name
    DWORD   fuFlags;              // verb flags
    DWORD   grfAttribs;           // Specifies some combination of the verb attributes in the OLEVERBATTRIB enumeration.
} VERBINFO;


#define UnregisterControlObject(x) UnregisterAutomationObject(&(x.automationDesc))


#ifdef LICENSING_SUPPORT
class CObjectFactory : public IClassFactory2
#else
class CObjectFactory : public IClassFactory
#endif
{
public:
    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void);
    // IClassFactory methods
    STDMETHOD(CreateInstance)(IUnknown *pUnkOuter, REFIID riid, void **ppbObjOut);
    STDMETHOD(LockServer)(BOOL fLock);
#ifdef LICENSING_SUPPORT
    // IClassFactory2 methods
    STDMETHOD(GetLicInfo)(LICINFO *pLicInfo);
    STDMETHOD(RequestLicKey)(DWORD dwReserved, BSTR *pbstrKey);
    STDMETHOD(CreateInstanceLic)(IUnknown *pUnkOuter, IUnknown *pUnkReserved, REFIID riid, BSTR bstrKey, void **ppvObjOut);
#endif

	CObjectFactory(int iIndex);
    ~CObjectFactory();

private:
	int index;
	ULONG m_refCount;
    
};

#define DECLARE_STANDARD_UNKNOWN() \
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut) { \
        return ExternalQueryInterface(riid, ppvObjOut); \
    } \
    STDMETHOD_(ULONG, AddRef)(void) { \
        return ExternalAddRef(); \
    } \
    STDMETHOD_(ULONG, Release)(void) { \
        return ExternalRelease(); \
    } \

extern BOOL g_fSysWinNT;
extern BOOL g_fSysWin95;
extern BOOL g_fSysWin95Shell;
extern BOOL g_fSysWin98Shell;
extern CRITICAL_SECTION g_CriticalSection;

void UnLoadTypeInfo();

#endif