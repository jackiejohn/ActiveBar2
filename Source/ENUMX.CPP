//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "enumx.h"

extern HANDLE g_hHeap;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CEnumX::CEnumX(REFCLSID rclsid,int cElements,int cbElementSize,void *rgElements,
    void (WINAPI *pfnCopyElement)(void *, const void *, DWORD))
	:
	m_iid(rclsid),
	m_cElements(cElements),
	m_cbElementSize(cbElementSize),
	m_iCurrent(0),
	m_rgElements(rgElements),
	m_pfnCopyElement(pfnCopyElement)
{
	m_pBaseObj=NULL;
	m_refCount=1;
}

CEnumX::~CEnumX()
{
    if (m_pBaseObj)
        m_pBaseObj->Release();
    else 
	    if (m_rgElements) HeapFree(g_hHeap, 0, m_rgElements);
}

STDMETHODIMP CEnumX::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (riid==IID_IUnknown)
	{
		*ppvObjOut=(IUnknown *)this;
		AddRef();
		return NOERROR;
	}
	if (riid==m_iid)
	{
		*ppvObjOut=(IEnumX *)this;
		AddRef();
		return NOERROR;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CEnumX::AddRef(void) {return ++m_refCount;}

STDMETHODIMP_(ULONG) CEnumX::Release(void) 
{
	if (0!=--m_refCount) return m_refCount;
	delete this;
	return 0;
}

STDMETHODIMP CEnumX::Next(unsigned long cElm,void *rgDest,unsigned long *pcElmOut)
{
    unsigned long cElementsFetched = 0;
    void         *pElementDest = rgDest;
    const void   *pElementSrc = (const BYTE *)m_rgElements + (m_cbElementSize * m_iCurrent);

//	if (m_cElements==0 && m_iid==IID_IEnumVARIANT)
//		(*((VARIANT **)rgDest))->pdispVal=NULL;

    while (cElementsFetched < cElm) 
	{
        // last element ?
        if (m_iCurrent >= m_cElements)
            break;
        // copy the element
        m_pfnCopyElement(pElementDest, pElementSrc, m_cbElementSize);
        // advance counters
        pElementDest = (LPBYTE)pElementDest + m_cbElementSize;
        pElementSrc  = (const BYTE *)pElementSrc + m_cbElementSize;
        m_iCurrent++;
        cElementsFetched++;
    }
    if (pcElmOut)
        *pcElmOut = cElementsFetched;

	if (cElementsFetched < cElm)
		return S_FALSE;
	
	return S_OK;

    //return (cElementsFetched < cElm)? S_FALSE : S_OK;
}

STDMETHODIMP CEnumX::Skip(unsigned long cSkip)
{
    // end ??
    if (m_iCurrent + (int)cSkip > m_cElements) 
	{
        m_iCurrent = m_cElements;
        return S_FALSE;
    }
    m_iCurrent += cSkip;
    return S_OK;
}

STDMETHODIMP CEnumX::Reset(void)
{
    m_iCurrent = 0;
    return S_OK;
}

STDMETHODIMP CEnumX::Clone(IEnumX **ppEnumClone)
{
    CEnumX *pNewEnum;

    pNewEnum = new CEnumX(m_iid, m_cElements, m_cbElementSize, m_rgElements, m_pfnCopyElement);
    if (!pNewEnum)
		return E_OUTOFMEMORY;
    pNewEnum->m_pBaseObj=this;
	AddRef();
    *ppEnumClone = (IEnumX *)pNewEnum;
    return S_OK;
}

