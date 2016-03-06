//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Stream.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

FileStream::FileStream()
	: m_hFile(NULL)
{
	m_bCloseHandleOnRelease=FALSE;
	m_nRefCount=1;
}

FileStream::~FileStream()
{
	if (m_bCloseHandleOnRelease && NULL != m_hFile)
		CloseHandle(m_hFile);
}

STDMETHODIMP FileStream::QueryInterface(REFIID riid, void **ppvObj)
{
	if (riid==IID_IUnknown || riid==IID_IStream)
	{
		*ppvObj=this;
		AddRef();
		return NOERROR;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) FileStream::AddRef(void)
{
	return ++m_nRefCount;
}

STDMETHODIMP_(ULONG) FileStream::Release(void)
{
	if (0!=--m_nRefCount)
		return m_nRefCount;
	delete this;
	return 0;
}

STDMETHODIMP FileStream::Read(void __RPC_FAR *pv,ULONG cb,ULONG __RPC_FAR *pcbRead)
{
	DWORD readRes;
	BOOL val=ReadFile(m_hFile,pv,cb,&readRes,NULL);
	if (pcbRead)
		*pcbRead=readRes;
	if (val==TRUE)
	{
		if (cb==readRes)
			return S_OK;
		return S_FALSE;
	}
	return E_FAIL;
}
        
STDMETHODIMP FileStream::Write(const void __RPC_FAR *pv,ULONG cb,ULONG __RPC_FAR *pcbWritten)
{
	DWORD writeRes;
	BOOL val=WriteFile(m_hFile,pv,cb,&writeRes,NULL);
	if (pcbWritten)
		*pcbWritten=writeRes;
	if (val==TRUE)
		return NOERROR;
	return E_FAIL;
}

STDMETHODIMP FileStream::Seek(LARGE_INTEGER dlibMove,DWORD dwOrigin,ULARGE_INTEGER __RPC_FAR *plibNewPosition)
{
	DWORD pos=SetFilePointer((HANDLE)m_hFile, dlibMove.LowPart, &(dlibMove.HighPart), dwOrigin); 
	if (plibNewPosition)
	{
		plibNewPosition->HighPart=0;
		plibNewPosition->LowPart=pos;
	}
	return NOERROR;
}

STDMETHODIMP FileStream::SetSize(ULARGE_INTEGER libNewSize)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP FileStream::CopyTo(IStream __RPC_FAR *pstm,ULARGE_INTEGER cb,ULARGE_INTEGER __RPC_FAR *pcbRead,ULARGE_INTEGER __RPC_FAR *pcbWritten)
{
	return E_NOTIMPL;
}
STDMETHODIMP FileStream::Commit(DWORD grfCommitFlags)
{
	return NOERROR;
}
        
STDMETHODIMP  FileStream::Revert( void)
{
	return E_NOTIMPL;
}

STDMETHODIMP  FileStream::LockRegion(ULARGE_INTEGER libOffset,ULARGE_INTEGER cb,DWORD dwLockType)
{
	return E_NOTIMPL;
}

STDMETHODIMP  FileStream::UnlockRegion(ULARGE_INTEGER libOffset,ULARGE_INTEGER cb,DWORD dwLockType)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP  FileStream::Stat(STATSTG __RPC_FAR *pstatstg,DWORD grfStatFlag)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP  FileStream::Clone(IStream __RPC_FAR *__RPC_FAR *ppstm)
{
	return E_NOTIMPL;
}
