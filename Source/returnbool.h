#ifndef __CRetBool_H__
#define __CRetBool_H__
#include "interfaces.h"

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class CRetBool : public IReturnBool
{
public:
	CRetBool();
	ULONG m_refCount;
	int m_objectIndex;
	~CRetBool();

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CRetBool *CreateInstance(IUnknown *pUnkOuter);
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
	
	

	// IReturnBool members

	STDMETHOD(get_Value)( VARIANT_BOOL *retval);
	STDMETHOD(put_Value)( VARIANT_BOOL retval);
	// DEFS for IReturnBool
	
	
	// Events 
	//{END INTERFACEDEFS}
	VARIANT_BOOL boolVal;
};


#endif
