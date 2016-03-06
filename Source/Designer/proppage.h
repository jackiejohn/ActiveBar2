#ifndef __PROPPAGE_H__


//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include <olectl.h>
#include "ipserver.h"

//=--------------------------------------------------------------------------=
// messages that we'll send to property pages to instruct them to accomplish
// tasks.

#define PPM_NEWOBJECTS    (WM_USER + 100)
#define PPM_APPLY         (WM_USER + 101)
#define PPM_EDITPROPERTY  (WM_USER + 102)
#define PPM_FREEOBJECTS   (WM_USER + 103)


class CPropertyPage : public IPropertyPage2 
{
public:
	PROPERTYPAGEINFO *m_pDesc;
    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void);

    // IPropertyPage methods
    STDMETHOD(SetPageSite)(LPPROPERTYPAGESITE pPageSite);
    STDMETHOD(Activate)(HWND hwndParent, LPCRECT lprc, BOOL bModal);
    STDMETHOD(Deactivate)(void);
    STDMETHOD(GetPageInfo)(LPPROPPAGEINFO pPageInfo);
    STDMETHOD(SetObjects)(ULONG cObjects, LPUNKNOWN FAR* ppunk);
    STDMETHOD(Show)(UINT nCmdShow);
    STDMETHOD(Move)(LPCRECT prect);
    STDMETHOD(IsPageDirty)(void);
    STDMETHOD(Apply)(void);
    STDMETHOD(Help)(LPCOLESTR lpszHelpDir);
    STDMETHOD(TranslateAccelerator)(LPMSG lpMsg);

    // IPropertyPage2 methods
    STDMETHOD(EditProperty)(THIS_ DISPID dispid);

    // constructor destructor
    CPropertyPage();
    virtual ~CPropertyPage();

    HINSTANCE GetResourceHandle(void);            // returns current resource handle.

	virtual HRESULT OnGetPageInfo(PROPPAGEINFO *pPropPageInfo) {return NOERROR;};
	virtual HRESULT OnActivate(){return NOERROR;};
	virtual HRESULT OnSetObjects() {return NOERROR;};
protected:
	ULONG m_refCount;
    IPropertyPageSite *m_pPropertyPageSite;       // pointer to our ppage site.
    void     MakeDirty();                         // makes the property page dirty.
	inline BOOL IsDirty() {return m_fDirty;};
    HWND     m_hwnd;                              // our hwnd.

    // the following two methods allow a property page implementer to get at all the
    // objects that we need to set here.
    IUnknown *FirstControl(DWORD *dwCookie);
    IUnknown *NextControl(DWORD *dwCookie);
	
	
protected:
    IUnknown **m_ppUnkObjects;                    // objects that we're working with.
    UINT     m_cObjects;                          // how many objects we're holding on to

    unsigned m_fActivated:1;
    unsigned m_fDirty:1;

    void     m_ReleaseAllObjects(void);           // clears out all objects we've got.
    HRESULT  m_EnsureLoaded(void);                // forces the load of the page.

    // default dialog proc for a page.
    static BOOL CALLBACK PropPageDlgProc(HWND, UINT, WPARAM, LPARAM);
    // all page implementers MUST implement the following function.
    virtual BOOL DialogProc(HWND, UINT, WPARAM, LPARAM) PURE;

	
};

#define __PROPPAGE_H__
#endif 