#ifndef __CDesigner_H__
#define __CDesigner_H__
#include "fontholder.h"
#include "DesignerInterfaces.h"

class CDlgDesigner;

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class CDesigner : public IDesigner
{
public:
	CDesigner(IUnknown *pUnkOuter);
	~CDesigner();

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CDesigner *CreateInstance(IUnknown *pUnkOuter);
	// IUnknown members
	DECLARE_STANDARD_UNKNOWN();

	// IDispatch members

	STDMETHOD(GetTypeInfoCount)( UINT *pctinfo);
	STDMETHOD(GetTypeInfo)( UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
	STDMETHOD(GetIDsOfNames)( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid);
	STDMETHOD(Invoke)( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr);
	// DEFS for IDispatch
	
	

	// IDesigner members

	STDMETHOD(OpenDesigner)( OLE_HANDLE hWndParent,  LPDISPATCH pActiveBar,  VARIANT_BOOL vbStandALone);
	STDMETHOD(CloseDesigner)();
	STDMETHOD(SetFocus)();
	STDMETHOD(UIDeactivateCloseDesigner)();
	// DEFS for IDesigner
	
	
	// Events 
	//{END INTERFACEDEFS}
	//{AGGREGATION SUPPORT}
	virtual HRESULT InternalQueryInterface(REFIID riid, void **ppvObjOut);
	inline IUnknown *PrivateUnknown (void) {
	    return &m_UnkPrivate;
	}

	IActiveBar2* m_pActiveBar;
	CDlgDesigner* m_pDlgDesigner; 
	HANDLE		 m_hDesignThread;
	DWORD		 m_dwThreadId;
	DWORD		 m_dwCurrentThreadId;
	HWND		 m_hWndMainParent;
	BOOL         m_bStandALone;
protected:
	HRESULT ExternalQueryInterface(REFIID riid, void **ppvObjOut) {
		return m_pUnkOuter->QueryInterface(riid, ppvObjOut);
	}
	ULONG ExternalAddRef(void) {
		return m_pUnkOuter->AddRef();
	}
	ULONG ExternalRelease(void) {
		return m_pUnkOuter->Release();
	}
	IUnknown *m_pUnkOuter;
private:
	class CPrivateUnknownObject : public IUnknown {
	public:
		STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut);
		STDMETHOD_(ULONG, AddRef)(void);
		STDMETHOD_(ULONG, Release)(void);
		CPrivateUnknownObject() : m_cRef(1) {}
	private:
		CDesigner *m_pMainUnknown();
		ULONG m_cRef;
	} m_UnkPrivate;

	friend class CPrivateUnknownObject;
	//{AGGREGATION SUPPORT}
};

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif
