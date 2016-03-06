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
// MemStream
//

class MemStream : public IPersistStream
{
public:
	enum
	{
		eBlockSize = 1024
	};

	MemStream(IStream* pStream, ULONG nDataSize);
	~MemStream();
		
	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj);
    STDMETHOD_(ULONG, AddRef)(THIS);
    STDMETHOD_(ULONG, Release)(THIS);

	HRESULT STDMETHODCALLTYPE GetClassID( /* [out] */ CLSID __RPC_FAR *pClassID);
	HRESULT STDMETHODCALLTYPE IsDirty( void);
	HRESULT STDMETHODCALLTYPE Load( /* [unique][in] */ IStream __RPC_FAR *pStm);
	HRESULT STDMETHODCALLTYPE Save( /* [unique][in] */ IStream __RPC_FAR *pStm,/* [in] */ BOOL fClearDirty);
	HRESULT STDMETHODCALLTYPE GetSizeMax( /* [out] */ ULARGE_INTEGER __RPC_FAR *pcbSize);

protected:
	IStream* m_pStream;
	ULONG    m_nRefCount;
	ULONG    m_nDataSize;
};

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

//
// CResourceStream
//

class CResourceStream : public IStream
{
public:
	CResourceStream();
	~CResourceStream();

	STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj);
    STDMETHOD_(ULONG, AddRef)(THIS);
    STDMETHOD_(ULONG, Release)(THIS);

	STDMETHOD(Read)(void* pv, ULONG cb, ULONG* pcbRead);
    STDMETHOD(Write)(const void* pv, ULONG cb, ULONG* pcbWritten);
	STDMETHOD(Seek)(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition);
    STDMETHOD(SetSize)(ULARGE_INTEGER libNewSize);
	STDMETHOD(CopyTo)(IStream* pStm,ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten);
    STDMETHOD(Commit)(DWORD grfCommitFlags);
    STDMETHOD(Revert)();
	STDMETHOD(LockRegion)(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
    STDMETHOD(UnlockRegion)(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
    STDMETHOD(Stat)(STATSTG* pStatStg, DWORD grfStatFlag);
    STDMETHOD(Clone)(IStream** ppStm);

	static HRESULT GetIStream(HMODULE   hModule, 
					          LPCTSTR   szName,
					          LPCTSTR   szResourceType,
					          IStream** ppOut);
protected:
	void Cleanup();
	void SetData(HGLOBAL hGlobal, UINT size, BOOL bAutoFree);
	HGLOBAL m_hMem;
	LPBYTE  m_pData;
	ULONG   m_nRefCount;
	ULONG   m_nDataSize;
	ULONG   m_nDataPos;
	BOOL    m_bAutoFree;
};

//
// CHeapStream
//

class CHeapStream : public IStream
{
public:
	CHeapStream();
	~CHeapStream();
	DWORD GetSize() {return m_dataSize;};
	DWORD GetSizeLeft() {return m_dataSize-m_curPos;};
	LPBYTE GetData() {return m_pBuffer;};
	void TrimLeft(int size);
	ULONG m_refCount;

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
	LPBYTE m_pBuffer;
	LONG m_dataSize;
	LONG m_bufferSize;
	LONG m_curPos;

	BOOL Grow(int sizeReq);

	friend class CHexReader;
};


// 
// BufferedFileStream
//

class BufferedFileStream : public IStream
{
public:
	BufferedFileStream();
	~BufferedFileStream();
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

	BOOL Flush();

protected:
	HANDLE m_hFile;
	LPBYTE m_pBuffer;
	DWORD  m_nCurrentBufferPosition;
	DWORD  m_nBufferSize;
	DWORD  m_nBufferSizeInitial;
	BOOL   m_bCloseHandleOnRelease;
	BOOL   m_bInitialRead;
	BOOL   m_bWriting;
};

#endif