#ifndef __CReturnString_H__
#define __CReturnString_H__
#include "interfaces.h"

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class CReturnString : public IReturnString
{
public:
	CReturnString();
	ULONG m_refCount;
	int m_objectIndex;
	~CReturnString();

	BSTR bstrVal;

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CReturnString *CreateInstance(IUnknown *pUnkOuter);
	// IUnknown members
	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
	STDMETHOD_(ULONG, AddRef)(void);
	STDMETHOD_(ULONG, Release)(void);

	// IDispatch members

	STDMETHOD(GetTypeInfoCount)( UINT *pctinfo);
	STDMETHOD(GetTypeInfo)( UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
	STDMETHOD(GetIDsOfNames)( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid);
	STDMETHOD(Invoke)( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr);
	// DEFS for IDispatch
	
	

	// IReturnString members

	STDMETHOD(get_Value)( BSTR *retval);
	STDMETHOD(put_Value)( BSTR retval);
	// DEFS for IReturnString
	
	
	// Events 
	//{END INTERFACEDEFS}
	BSTR m_bstrHelpFile;
};

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void* objectDef);

#endif
