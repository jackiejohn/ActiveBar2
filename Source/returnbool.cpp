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
#include "ipserver.h"
#include <stddef.h>       // for offsetof()
#include "support.h"
#include "resource.h"
#include "debug.h"
#include "returnbool.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern ITypeInfo *GetObjectTypeInfoEx(LCID lcid,REFIID iid);


//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CRetBool(IUnknown *pUnkOuter)
{
	return (IUnknown *)(IReturnBool *)new CRetBool();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_ReturnBool,ReturnBool,_T("ReturnBool"),CreateFN_CRetBool,2,0,&IID_IReturnBool,_T(""));
void *CRetBool::objectDef=&ReturnBoolObject;
CRetBool *CRetBool::CreateInstance(IUnknown *pUnkOuter)
{
	return new CRetBool();
}
//{OBJECT CREATEFN}
CRetBool::CRetBool()
	: m_refCount(1)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
}

CRetBool::~CRetBool()
{
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
// {END CLEANUP}
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CRetBool::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IUnknown))
	{
		AddRef();
		*ppvObjOut=this;
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IDispatch))
	{
		*ppvObjOut=(void *)(IDispatch *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IReturnBool))
	{
		*ppvObjOut=(void *)(IReturnBool *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
STDMETHODIMP_(ULONG) CRetBool::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CRetBool::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
// IDispatch members

STDMETHODIMP CRetBool::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CRetBool::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfoEx(lcid,IID_IReturnBool);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CRetBool::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfoEx(lcid,IID_IReturnBool);
	if (pTypeInfo==NULL)
		return E_FAIL;
	hr=pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgdispid);
	pTypeInfo->Release();
	return hr;
}
STDMETHODIMP CRetBool::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
{
#ifdef _DEBUG
	if (dispidMember<0)
		TRACE1(1, "IDispatch::Invoke -%X\n",-dispidMember)
	else
		TRACE1(1, "IDispatch::Invoke %X\n",dispidMember)
#endif
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfoEx(lcid,IID_IReturnBool);
	if (pTypeInfo==NULL)
		return E_FAIL;
	#ifdef ERRORINFO_SUPPORT
	SetErrorInfo(0L,NULL); // should be called if ISupportErrorInfo is used
	#endif
	hr = pTypeInfo->Invoke((IDispatch *)this, dispidMember, wFlags,
		pdispparams, pvarResult,
        pexcepinfo, puArgErr);
    pTypeInfo->Release();
	return hr;
}

// ReturnBool members

STDMETHODIMP CRetBool::get_Value(VARIANT_BOOL *retval)
{
	*retval=boolVal;
	return NOERROR;
}
STDMETHODIMP CRetBool::put_Value(VARIANT_BOOL val)
{
	boolVal=val;
	return NOERROR;
}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
