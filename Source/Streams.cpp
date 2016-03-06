//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Streams.h"

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
	if (m_bCloseHandleOnRelease && m_hFile)
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

//
// MemStream
//

MemStream::MemStream(IStream *pStream,ULONG dataSize)
{
	m_nRefCount=1;
	m_pStream=pStream;
	m_pStream->AddRef();
	m_nDataSize=dataSize;
}

MemStream::~MemStream()
{
	m_pStream->Release();
}

STDMETHODIMP MemStream::QueryInterface(THIS_ REFIID riid, LPVOID FAR* ppvObj)
{
	if (IsEqualIID(riid, IID_IUnknown) ||
		IsEqualIID(riid, IID_IPersist) ||
		IsEqualIID(riid, IID_IPersistStream))
	{
		AddRef();
		*ppvObj = this;
		return S_OK;
	}

	*ppvObj = NULL;
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) MemStream::AddRef(THIS)
{
	return ++m_nRefCount;
}

STDMETHODIMP_(ULONG) MemStream::Release(THIS)
{
	if (0!=--m_nRefCount)
		return m_nRefCount;
	delete this;
	return 0;
}

static const CLSID _clsidBlobProperty =
{ 0xf6f07540, 0x42ec, 0x11ce, { 0x81, 0x35, 0x0, 0xaa, 0x0, 0x4b, 0xb8, 0x51 } };

STDMETHODIMP MemStream::GetClassID( /* [out] */ CLSID __RPC_FAR *pClassID)
{
	*pClassID = _clsidBlobProperty;
	return S_OK;
}

STDMETHODIMP MemStream::IsDirty( void)
{
	return S_OK;
}

STDMETHODIMP MemStream::Load( /* [unique][in] */ IStream __RPC_FAR *pStm)
{
	ULARGE_INTEGER temp,temp2;
	ULARGE_INTEGER bigSize;
	pStm->Read(&m_nDataSize,sizeof(ULONG),NULL);
	bigSize.LowPart=m_nDataSize;
	bigSize.HighPart=0;
	return pStm->CopyTo(m_pStream,bigSize,&temp,&temp2);
	
}

STDMETHODIMP MemStream::Save( /* [unique][in] */ IStream __RPC_FAR *pStm,/* [in] */ BOOL fClearDirty)
{
	ULARGE_INTEGER temp,temp2;
	LARGE_INTEGER pos;
	ULARGE_INTEGER bigSize;

	pos.LowPart=0;pos.HighPart=0;
	m_pStream->Seek(pos,STREAM_SEEK_SET,&temp);
	pStm->Write(&m_nDataSize,sizeof(ULONG),NULL);
	bigSize.LowPart=m_nDataSize;
	bigSize.HighPart=0;
	return m_pStream->CopyTo(pStm,bigSize,&temp,&temp2);
}

STDMETHODIMP MemStream::GetSizeMax( /* [out] */ ULARGE_INTEGER __RPC_FAR *pcbSize)
{
	if (pcbSize)
	{
		pcbSize->LowPart=m_nDataSize;
		pcbSize->HighPart=0;
		return S_OK;
	}
	return S_FALSE;
}

CResourceStream::CResourceStream()
	: m_hMem(NULL),
	  m_nRefCount(1)
{
}

CResourceStream::~CResourceStream()
{
	Cleanup();
}

void CResourceStream::Cleanup()
{
	GlobalUnlock(m_hMem);
	if (m_hMem && m_bAutoFree)
	{
		GlobalFree(m_hMem);
		m_hMem = NULL;
	}
}

void CResourceStream::SetData(HGLOBAL hGlobal, UINT size, BOOL bAutoFree)
{
	m_hMem = hGlobal;
	m_pData = (LPBYTE)GlobalLock(m_hMem);
	m_bAutoFree = bAutoFree;
	m_nDataSize = size;
	m_nDataPos = 0;
}

HRESULT CResourceStream::GetIStream(HMODULE   hModule, 
								    LPCTSTR   szName,
								    LPCTSTR   szResourceType,
								    IStream** ppOut)
{
	HRSRC hResource = FindResource(hModule, szName, szResourceType);
	if (NULL == hResource)
		return E_FAIL;

	CResourceStream* pStream = new CResourceStream;
	if (NULL == pStream)
		return E_OUTOFMEMORY;

	pStream->m_hMem = LoadResource(hModule, hResource);
	pStream->m_pData = (LPBYTE)LockResource(pStream->m_hMem);
	pStream->m_bAutoFree = FALSE;
	pStream->m_nDataSize = SizeofResource(hModule, hResource);
	pStream->m_nDataPos = 0;
	*ppOut = pStream;
	return NOERROR;
}

STDMETHODIMP CResourceStream::QueryInterface(THIS_ REFIID riid, LPVOID FAR* ppvObj)
{
	if (IsEqualIID(riid, IID_IUnknown) ||
		IsEqualIID(riid, IID_IStream))
	{
		AddRef();
		*ppvObj = this;
		return S_OK;
	}

	*ppvObj = NULL;
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CResourceStream::AddRef(THIS)
{
	return ++m_nRefCount;
}

STDMETHODIMP_(ULONG) CResourceStream::Release(THIS)
{
	if (0 != --m_nRefCount)
		return m_nRefCount;
	delete this;
	return 0;
}

STDMETHODIMP CResourceStream::Read(void* pv, ULONG cb, ULONG* pcbRead)
{
	ULONG nReadSize;
	if (cb < (m_nDataSize-m_nDataPos))
		nReadSize = cb;
	else
		nReadSize = m_nDataSize-m_nDataPos;

	if (pcbRead)
		*pcbRead = nReadSize;

	if (0 == nReadSize)
		return S_FALSE;
	
	memcpy(pv, (m_pData + m_nDataPos), nReadSize);
	m_nDataPos += nReadSize;
	return NOERROR;
}
        
STDMETHODIMP CResourceStream::Write(const void* pv, ULONG cb, ULONG* pcbWritten)
{
	return E_NOTIMPL;
}

STDMETHODIMP CResourceStream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)
{
	if (0 != dlibMove.HighPart)
		return E_FAIL;

	switch(dwOrigin)
	{
	case STREAM_SEEK_SET:
		if (dlibMove.LowPart > m_nDataSize)
			return STG_E_INVALIDFUNCTION;
		m_nDataPos=dlibMove.LowPart;
		break;

	case STREAM_SEEK_CUR:
		if ((dlibMove.LowPart + m_nDataPos) > m_nDataSize || (dlibMove.LowPart + m_nDataPos) < 0)
			return STG_E_INVALIDFUNCTION;
		m_nDataPos += dlibMove.LowPart;
		break;

	case STREAM_SEEK_END:
		if ((dlibMove.LowPart + m_nDataSize) > m_nDataSize || (dlibMove.LowPart + m_nDataSize) < 0)
			return STG_E_INVALIDFUNCTION;
		m_nDataPos = m_nDataSize+dlibMove.LowPart;
		break;
	}
	if (plibNewPosition)
	{
		plibNewPosition->HighPart = 0;
		plibNewPosition->LowPart = m_nDataPos;
	}
	return NOERROR;
}

STDMETHODIMP CResourceStream::SetSize(ULARGE_INTEGER libNewSize)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP CResourceStream::CopyTo(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten)
{
	return E_NOTIMPL;
}
STDMETHODIMP CResourceStream::Commit(DWORD grfCommitFlags)
{
	return NOERROR;
}
        
STDMETHODIMP  CResourceStream::Revert()
{
	return E_NOTIMPL;
}

STDMETHODIMP  CResourceStream::LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
	return E_NOTIMPL;
}

STDMETHODIMP  CResourceStream::UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
	return E_NOTIMPL;
}

STDMETHODIMP  CResourceStream::Stat(STATSTG* pstatstg, DWORD grfStatFlag)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP  CResourceStream::Clone(IStream** ppStm)
{
	return E_NOTIMPL;
}

#pragma intrinsic(memcpy)

CHeapStream::CHeapStream()
{
	m_refCount=1;
	m_pBuffer=NULL;
	m_dataSize=0;
	m_bufferSize=0;
	m_curPos=0;
}

CHeapStream::~CHeapStream()
{
	if (m_pBuffer)
		free(m_pBuffer);
}

static const int GROWBLOCKSIZE=256;

BOOL CHeapStream::Grow(int sizeReq)
{
	int newSize=sizeReq>(m_bufferSize+GROWBLOCKSIZE) ?
		sizeReq+GROWBLOCKSIZE : m_bufferSize+GROWBLOCKSIZE;
	LPBYTE pNewBuffer;
	if (m_pBuffer)
		pNewBuffer=(LPBYTE)realloc(m_pBuffer,newSize);
	else
		pNewBuffer=(LPBYTE)malloc(newSize);
	if (!pNewBuffer)
		return FALSE; // out of memory
	m_pBuffer=pNewBuffer;
	m_bufferSize=newSize;
	return TRUE;
}


STDMETHODIMP CHeapStream::QueryInterface(REFIID riid, void **ppvObj)
{
	if (riid==IID_IUnknown || riid==IID_IStream)
	{
		*ppvObj=this;
		AddRef();
		return NOERROR;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CHeapStream::AddRef(void)
{
	return ++m_refCount;
}

STDMETHODIMP_(ULONG) CHeapStream::Release(void)
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}

STDMETHODIMP CHeapStream::Read(void __RPC_FAR *pv,ULONG cb,ULONG __RPC_FAR *pcbRead)
{
	ULONG readAmount;
	readAmount=m_dataSize-m_curPos;
	if (readAmount==0)
	{
		if (pcbRead)
			*pcbRead=0;
		return NOERROR;
	}
	if (cb<readAmount)
		readAmount=cb;

	if (cb!=0)
	{
		memcpy(pv,m_pBuffer+m_curPos,readAmount);
	}
	if (pcbRead)
		*pcbRead=readAmount;
	m_curPos+=readAmount;
	return NOERROR;
}
        
STDMETHODIMP CHeapStream::Write(const void __RPC_FAR *pv,ULONG cb,ULONG __RPC_FAR *pcbWritten)
{
	if ((cb+(ULONG)m_curPos)>(ULONG)m_bufferSize)
	{
		if (!Grow(cb+m_curPos))
			return E_OUTOFMEMORY;
	}


	memcpy(m_pBuffer+m_curPos,pv,cb);

	if (pcbWritten)
		*pcbWritten=cb;

	if (((ULONG)cb+(ULONG)m_curPos)>(ULONG)m_dataSize) 
		m_dataSize=cb+m_curPos;
	m_curPos+=cb;
	return NOERROR;
}

STDMETHODIMP CHeapStream::Seek(LARGE_INTEGER dlibMove,DWORD dwOrigin,ULARGE_INTEGER __RPC_FAR *plibNewPosition)
{
	switch(dwOrigin)
	{
	case STREAM_SEEK_SET:
		if (dlibMove.LowPart>(ULONG)m_dataSize) // beyond buffer ?
		{
			if (plibNewPosition)
			{
				plibNewPosition->LowPart=(ULONG)m_curPos;
				plibNewPosition->HighPart=0;
			}
			return E_FAIL;
		}
		m_curPos=dlibMove.LowPart;
		if (plibNewPosition)
		{
			plibNewPosition->LowPart=(ULONG)m_curPos;
			plibNewPosition->HighPart=0;
		}
		return NOERROR;
	case STREAM_SEEK_CUR:
		{
			LONG newPos=m_curPos;
			newPos+=dlibMove.LowPart;
			if (newPos<0)
				return E_FAIL;
			if (newPos>m_dataSize)
			{
				Grow(newPos);
				m_dataSize=newPos;
			}
			m_curPos=newPos;
			if (plibNewPosition)
			{
				plibNewPosition->LowPart=m_curPos;
				plibNewPosition->HighPart=0;
			}
		}
		return NOERROR;
	case STREAM_SEEK_END:
		{
			LONG newPos=m_dataSize;
			newPos+=dlibMove.LowPart;
			if (newPos<0)
				return E_FAIL;
			if (newPos>m_dataSize)
				return E_FAIL;
			m_curPos=newPos;
			if (plibNewPosition)
			{
				plibNewPosition->LowPart=m_curPos;
				plibNewPosition->HighPart=0;
			}
		}
		return NOERROR;
	default:
		return STG_E_INVALIDFUNCTION;
	}
	return E_FAIL;
}

STDMETHODIMP CHeapStream::SetSize(ULARGE_INTEGER libNewSize)
{
	if (libNewSize.LowPart>(ULONG)m_dataSize)
	{
		Grow(libNewSize.LowPart);
		m_dataSize=libNewSize.LowPart;
	}
	else
	{
		m_dataSize=libNewSize.LowPart;
	}
	return NOERROR;
}
        
STDMETHODIMP CHeapStream::CopyTo(IStream __RPC_FAR *pstm,ULARGE_INTEGER cb,ULARGE_INTEGER __RPC_FAR *pcbRead,ULARGE_INTEGER __RPC_FAR *pcbWritten)
{
	HRESULT hr;
	int copyAmount=m_dataSize-m_curPos;
	if (cb.LowPart<(ULONG)copyAmount)
		copyAmount=cb.LowPart;
	if (copyAmount==0)
		return NOERROR;
	ULONG written;
	hr=pstm->Write(m_pBuffer+m_curPos,copyAmount,&written);
	if (FAILED(hr))
		return hr;
	if (pcbWritten)
	{
		pcbWritten->LowPart=written;
		pcbWritten->HighPart=0;
	}
	m_curPos+=copyAmount;
	if (pcbRead)
	{
		pcbRead->LowPart=(ULONG)copyAmount;
		pcbRead->HighPart=0;
	}
	return NOERROR;
}

STDMETHODIMP CHeapStream::Commit(DWORD grfCommitFlags)
{
	return NOERROR;
}
        
STDMETHODIMP  CHeapStream::Revert( void)
{
	return E_NOTIMPL;
}

        
STDMETHODIMP  CHeapStream::LockRegion(ULARGE_INTEGER libOffset,ULARGE_INTEGER cb,DWORD dwLockType)
{
	return E_NOTIMPL;
}

STDMETHODIMP  CHeapStream::UnlockRegion(ULARGE_INTEGER libOffset,ULARGE_INTEGER cb,DWORD dwLockType)
{
	return E_NOTIMPL;
}

        
STDMETHODIMP  CHeapStream::Stat(STATSTG __RPC_FAR *pstatstg,DWORD grfStatFlag)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP  CHeapStream::Clone(IStream __RPC_FAR *__RPC_FAR *ppstm)
{
	CHeapStream *pNewStream;
	pNewStream=new CHeapStream();
	if (!pNewStream)
		return E_OUTOFMEMORY;
	if (m_dataSize!=0)
	{
		if (pNewStream->Grow(m_dataSize)==FALSE)
		{
			delete pNewStream;
			return E_OUTOFMEMORY;
		}
		memcpy(pNewStream->m_pBuffer,m_pBuffer,m_dataSize);
		pNewStream->m_dataSize=m_dataSize;
		pNewStream->m_curPos=m_curPos;
	}
	*ppstm=(IStream *)pNewStream;
	return NOERROR;
}

void CHeapStream::TrimLeft(int size)
{
	assert(m_dataSize>=size);
	if (m_dataSize!=size)
		memcpy(m_pBuffer,m_pBuffer+size,m_dataSize-size);
	m_dataSize-=size;
	if (m_curPos>=size)
		m_curPos-=size;
}

//
// BufferedFileStream
//

BufferedFileStream::BufferedFileStream()
	: m_hFile(NULL),
	  m_pBuffer(NULL)
{
	m_bWriting = FALSE;
	m_nBufferSize = 2097152;
//	m_nBufferSize = 512;
	m_nBufferSizeInitial = 0;
	m_bCloseHandleOnRelease=FALSE;
	m_nCurrentBufferPosition = 0;
	m_nRefCount=1;
	m_bInitialRead = TRUE;
}

BufferedFileStream::~BufferedFileStream()
{
	if (m_bWriting)
		Flush();
	if (m_bCloseHandleOnRelease && m_hFile)
		CloseHandle(m_hFile);
	delete [] m_pBuffer;
}

void BufferedFileStream::SetHandle(HANDLE hFile, BOOL bCloseHandleOnRelease)
{
	m_hFile = hFile;
	m_bCloseHandleOnRelease = bCloseHandleOnRelease;
	m_pBuffer = new BYTE[m_nBufferSize];
	while (NULL == m_pBuffer)
	{
		m_nBufferSize /= 2;
		if (m_nBufferSize <= 0)
			break;
		m_pBuffer = new BYTE[m_nBufferSize];
	}
	m_nBufferSizeInitial = m_nBufferSize;
	assert(m_pBuffer);
};

STDMETHODIMP BufferedFileStream::QueryInterface(REFIID riid, void **ppvObj)
{
	if (riid==IID_IUnknown || riid==IID_IStream)
	{
		*ppvObj=this;
		AddRef();
		return NOERROR;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) BufferedFileStream::AddRef(void)
{
	return ++m_nRefCount;
}

STDMETHODIMP_(ULONG) BufferedFileStream::Release(void)
{
	if (0!=--m_nRefCount)
		return m_nRefCount;
	delete this;
	return 0;
}

STDMETHODIMP BufferedFileStream::Read(void __RPC_FAR *pv, ULONG cb, ULONG __RPC_FAR *pcbRead)
{
	assert(!m_bWriting);
	BOOL bResult;
	DWORD dwReadResult = 1;
	if (m_bInitialRead)
	{
		m_bInitialRead = FALSE;
		// Read Stuff In
		bResult = ReadFile(m_hFile, m_pBuffer, m_nBufferSize, &dwReadResult, NULL);
		if (0 == dwReadResult)
			return E_FAIL;

		if (!bResult)
			return E_FAIL;

		if (m_nBufferSize != dwReadResult)
			m_nBufferSize = dwReadResult;

		m_nCurrentBufferPosition = 0;
	}
	DWORD cb2 = cb;
	DWORD dwRead = 0;
	DWORD dwLeft = 0;
	DWORD dwInputBufferPosition = 0;
	while (m_nBufferSize - m_nCurrentBufferPosition < cb2)
	{
		dwLeft = m_nBufferSize - m_nCurrentBufferPosition;
		memcpy((LPBYTE)pv + dwRead, m_pBuffer + m_nCurrentBufferPosition, dwLeft);
		cb2 -= dwLeft;
		dwRead += dwLeft;
		bResult = ReadFile(m_hFile, m_pBuffer, m_nBufferSize, &dwReadResult, NULL);
		if (!bResult)
			return E_FAIL;

		if (m_nBufferSize != dwReadResult)
			m_nBufferSize = dwReadResult;
		m_nCurrentBufferPosition = 0;
	}
	if (cb2 > 0)
	{
		memcpy((LPBYTE)pv + dwRead, m_pBuffer + m_nCurrentBufferPosition, cb2);
		m_nCurrentBufferPosition += cb2;
		dwRead += cb2;
	}
	if (pcbRead)
		*pcbRead = dwRead;
	return NOERROR;
}
        
STDMETHODIMP BufferedFileStream::Write(const void __RPC_FAR *pv,ULONG cb,ULONG __RPC_FAR *pcbWritten)
{
	m_bWriting = TRUE;
	LPBYTE pByte = (LPBYTE)pv;
	DWORD dwNumberOfBytesWritten;
	BOOL bResult;
	DWORD dwLeft = 0;
	DWORD dwInputBufferPosition = 0;
	DWORD cb2 = cb;
	while (m_nBufferSize < m_nCurrentBufferPosition + cb2)
	{
		dwLeft = m_nBufferSize - m_nCurrentBufferPosition;
		memcpy(m_pBuffer + m_nCurrentBufferPosition, (LPBYTE)(pByte + dwInputBufferPosition), dwLeft);
		bResult = WriteFile(m_hFile, m_pBuffer, m_nBufferSize, &dwNumberOfBytesWritten, NULL);
		if (!bResult)
			return E_FAIL;
		m_nCurrentBufferPosition = 0;
		dwInputBufferPosition += dwLeft;
		cb2 -= dwLeft;
	}
	if (cb2 > 0)
	{
		memcpy(m_pBuffer + m_nCurrentBufferPosition, (LPBYTE)(pByte + dwInputBufferPosition), cb2);
		m_nCurrentBufferPosition += cb2;
		dwInputBufferPosition += cb2;
	}
	if (pcbWritten)
		*pcbWritten = dwInputBufferPosition;
	return NOERROR;
}

BOOL BufferedFileStream::Flush()
{
	BOOL bResult;
	DWORD dwNumberOfBytesWritten;
	if (m_nCurrentBufferPosition > 0)
		bResult = WriteFile(m_hFile, m_pBuffer, m_nCurrentBufferPosition, &dwNumberOfBytesWritten, NULL);
	if (bResult)
		bResult = FlushFileBuffers(m_hFile);
	return bResult;
}

STDMETHODIMP BufferedFileStream::Seek(LARGE_INTEGER dlibMove,DWORD dwOrigin,ULARGE_INTEGER __RPC_FAR *plibNewPosition)
{
//FILE_BEGIN         0
//FILE_CURRENT       1
//FILE_END           2
//STREAM_SEEK_SET	 0
//STREAM_SEEK_CUR	 1
//STREAM_SEEK_END	 2

	DWORD pos=SetFilePointer((HANDLE)m_hFile, dlibMove.LowPart, &(dlibMove.HighPart), dwOrigin); 
	if (plibNewPosition)
	{
		plibNewPosition->HighPart=0;
		plibNewPosition->LowPart=pos;
	}

	if (STREAM_SEEK_CUR == dwOrigin && 0 == dlibMove.LowPart)
		return NOERROR;

	m_nCurrentBufferPosition = 0;
	m_nBufferSize = m_nBufferSizeInitial;
	m_bInitialRead = TRUE;
	return NOERROR;
}

STDMETHODIMP BufferedFileStream::SetSize(ULARGE_INTEGER libNewSize)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP BufferedFileStream::CopyTo(IStream __RPC_FAR *pstm,ULARGE_INTEGER cb,ULARGE_INTEGER __RPC_FAR *pcbRead,ULARGE_INTEGER __RPC_FAR *pcbWritten)
{
	return E_NOTIMPL;
}
STDMETHODIMP BufferedFileStream::Commit(DWORD grfCommitFlags)
{
	return NOERROR;
}
        
STDMETHODIMP  BufferedFileStream::Revert( void)
{
	return E_NOTIMPL;
}

STDMETHODIMP  BufferedFileStream::LockRegion(ULARGE_INTEGER libOffset,ULARGE_INTEGER cb,DWORD dwLockType)
{
	return E_NOTIMPL;
}

STDMETHODIMP  BufferedFileStream::UnlockRegion(ULARGE_INTEGER libOffset,ULARGE_INTEGER cb,DWORD dwLockType)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP  BufferedFileStream::Stat(STATSTG __RPC_FAR *pstatstg,DWORD grfStatFlag)
{
	return E_NOTIMPL;
}
        
STDMETHODIMP  BufferedFileStream::Clone(IStream __RPC_FAR *__RPC_FAR *ppstm)
{
	return E_NOTIMPL;
}
