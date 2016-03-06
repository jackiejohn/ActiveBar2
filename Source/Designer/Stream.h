#ifndef STREAM_INCLUDED
#define STREAM_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

//
// FileStream
//

class FileStream : public IStream
{
public:
	FileStream();
	~FileStream();
	ULONG m_nRefCount;
	void SetHandle(HANDLE hFile, BOOL bCloseHandleOnRelease);

	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void);
    

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE Read( 
            /* [length_is][size_is][out] */ void __RPC_FAR *pv,
            /* [in] */ ULONG cb,
            /* [out] */ ULONG __RPC_FAR *pcbRead) ;
        
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Write( 
            /* [size_is][in] */ const void __RPC_FAR *pv,
            /* [in] */ ULONG cb,
            /* [out] */ ULONG __RPC_FAR *pcbWritten) ;
        
	virtual /* [local] */ HRESULT STDMETHODCALLTYPE Seek( 
            /* [in] */ LARGE_INTEGER dlibMove,
            /* [in] */ DWORD dwOrigin,
            /* [out] */ ULARGE_INTEGER __RPC_FAR *plibNewPosition) ;
        
    virtual HRESULT STDMETHODCALLTYPE SetSize( 
            /* [in] */ ULARGE_INTEGER libNewSize) ;
        
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE CopyTo( 
            /* [unique][in] */ IStream __RPC_FAR *pstm,
            /* [in] */ ULARGE_INTEGER cb,
            /* [out] */ ULARGE_INTEGER __RPC_FAR *pcbRead,
            /* [out] */ ULARGE_INTEGER __RPC_FAR *pcbWritten) ;
        
    virtual HRESULT STDMETHODCALLTYPE Commit( 
            /* [in] */ DWORD grfCommitFlags) ;
        
    virtual HRESULT STDMETHODCALLTYPE Revert( void) ;
        
    virtual HRESULT STDMETHODCALLTYPE LockRegion( 
            /* [in] */ ULARGE_INTEGER libOffset,
            /* [in] */ ULARGE_INTEGER cb,
            /* [in] */ DWORD dwLockType) ;
        
    virtual HRESULT STDMETHODCALLTYPE UnlockRegion( 
            /* [in] */ ULARGE_INTEGER libOffset,
            /* [in] */ ULARGE_INTEGER cb,
            /* [in] */ DWORD dwLockType) ;
        
    virtual HRESULT STDMETHODCALLTYPE Stat( 
            /* [out] */ STATSTG __RPC_FAR *pstatstg,
            /* [in] */ DWORD grfStatFlag) ;
        
    virtual HRESULT STDMETHODCALLTYPE Clone( 
            /* [out] */ IStream __RPC_FAR *__RPC_FAR *ppstm) ;

protected:
	HANDLE m_hFile;
	BOOL   m_bCloseHandleOnRelease;
};

inline void FileStream::SetHandle(HANDLE hFile, BOOL bCloseHandleOnRelease)
{
	m_hFile = hFile;
	m_bCloseHandleOnRelease = bCloseHandleOnRelease;
};

#endif