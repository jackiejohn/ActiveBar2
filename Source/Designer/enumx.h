#ifndef __XENUM_H__
#define __XENUM_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class IEnumX: public IUnknown 
{
public:
	STDMETHOD(Next)(ULONG celt, LPVOID rgelt, ULONG *pceltFetched) PURE;
	STDMETHOD(Skip)(ULONG celt) PURE;
	STDMETHOD(Reset)(void) PURE;
	STDMETHOD(Clone)(IEnumX **ppenum) PURE;
};

// This is a generic enumerator implementation. It relies on a copyelement function
// passed to it during construction. 
// !!!rgElements should be alloced using HeapAlloc!!!
// 
// The pfnCopyElement func. Looks like :
// void WINAPI MyCopyFunc(void *pvDest,const void *pvSrc,DWORD cbCopy)
//{
//    MYSTRUCT *pDest = (VERBINFO *) pvDest;
//    const MYSTRUCT * pSrc = (const MYSTRUCT *) pvSrc;
//    pSrc->Clone(pDest);
//}
// 

class CEnumX : public IEnumX
{
public:
	// IUnknown members
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void);
	// IEnum* members
	STDMETHOD(Next)(ULONG celt, LPVOID rgelt, ULONG *pceltFetched);
	STDMETHOD(Skip)(ULONG celt);
	STDMETHOD(Reset)(void);
	STDMETHOD(Clone)(IEnumX **ppenum);

	CEnumX(REFIID riid, int cElement, int cbElement, void *rgElements,
                 void (WINAPI * pfnCopyElement)(void *, const void *, DWORD));
	~CEnumX();
	
protected:
	ULONG m_refCount;
	IID m_iid;			// type of enumerator
	int m_cElements;    // number of elements
	int m_cbElementSize;// element size
    int m_iCurrent;     // curpos
    VOID * m_rgElements;// element array
    CEnumX *m_pBaseObj; // base object pointer if this is a clone
    void  (WINAPI * m_pfnCopyElement)(void *, const void *, DWORD); // copy func
};


#endif
