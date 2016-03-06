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
#include "returnstring.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CReturnString(IUnknown *pUnkOuter)
{
	return (IUnknown *)(IReturnString *)new CReturnString();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_ReturnString,ReturnString,_T("ReturnString"),CreateFN_CReturnString,2,0,&IID_IReturnString,_T(""));
void *CReturnString::objectDef=&ReturnStringObject;
CReturnString *CReturnString::CreateInstance(IUnknown *pUnkOuter)
{
	return new CReturnString();
}
//{OBJECT CREATEFN}
CReturnString::CReturnString()
	: m_refCount(1)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
	bstrVal=NULL;
}

CReturnString::~CReturnString()
{
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
// {END CLEANUP}
	SysFreeString(bstrVal);
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CReturnString::QueryInterface(REFIID riid, void **ppvObjOut)
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
	if (DO_GUIDS_MATCH(riid,IID_IReturnString))
	{
		*ppvObjOut=(void *)(IReturnString *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
STDMETHODIMP_(ULONG) CReturnString::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CReturnString::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
STDMETHODIMP CReturnString::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CReturnString::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CReturnString::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;
	hr=pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgdispid);
	pTypeInfo->Release();
	return hr;
}
STDMETHODIMP CReturnString::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
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
	pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
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
STDMETHODIMP CReturnString::get_Value(BSTR *retval)
{
	*retval=SysAllocString(bstrVal);
	return NOERROR;
}
STDMETHODIMP CReturnString::put_Value(BSTR val)
{
	SysFreeString(bstrVal);
	bstrVal=SysAllocString(val);
	return NOERROR;
}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
