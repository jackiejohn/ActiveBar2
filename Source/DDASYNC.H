//---------------------------------------------------------------------
//  Copyright (c) 1995-1996, Data Dynamics. All rights reserved.
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//---------------------------------------------------------------------
#ifndef __DDASYNC_H__
#define __DDASYNC_H__


//#define COM_NO_WINDOWS_H
#include "urlmon.h"
#include "FWnd.h"
class CHeapStream;

typedef HRESULT (_stdcall *PFN_CREATEURLMONIKER)(LPMONIKER pMkCtx, LPCWSTR szURL, LPMONIKER FAR * ppmk);
typedef HRESULT (_stdcall *PFN_REGISTERBINDSTATUSCALLBACK)(LPBC pBC, IBindStatusCallback *pBSCb,                     
                                IBindStatusCallback**  ppBSCBPrev, DWORD dwReserved);
typedef HRESULT (_stdcall *PFN_REVOKEBINDSTATUSCALLBACK)(LPBC pBC, IBindStatusCallback *pBSCb);

class UrlMonikerLib
{
public:
	UrlMonikerLib();
	~UrlMonikerLib();

	HRESULT CreateURLMoniker(LPMONIKER pMkCtx, LPCWSTR szURL, LPMONIKER FAR * ppmk);
	HRESULT RegisterBindStatusCallback(LPBC pBC, IBindStatusCallback *pBSCb,                     
                                IBindStatusCallback**  ppBSCBPrev, DWORD dwReserved);
	HRESULT RevokeBindStatusCallback(LPBC pBC, IBindStatusCallback *pBSCb);
protected:
	BOOL loadtry; // is set to TRUE after initial attempt to load has been made
	HINSTANCE urlmonlib;

	PFN_CREATEURLMONIKER m_pCreateURLMoniker;
	PFN_REGISTERBINDSTATUSCALLBACK m_pRegisterBindStatusCallback;
	PFN_REVOKEBINDSTATUSCALLBACK m_pRevokeBindStatusCallback;
	BOOL CheckLoaded();
};

#define ASYNC_MSG_DATAREADY 1
#define ASYNC_MSG_DATAPUSH 2
#define ASYNC_MSG_PROGRESS 3
#define ASYNC_MSG_ERROR 4

#define ASYNC_STATE_READY 0
#define ASYNC_STATE_DOWNLOADING 1
#define ASYNC_STATE_LOADED 2
#define ASYNC_STATE_ERROR 3

class AsyncDownloadItem;
class CDownloadSink;

typedef void (*ASYNCDOWNLOADCALLBACK)(void *cookie,AsyncDownloadItem *item,DWORD dwReason);

class AsyncDownloadItem
{
public:
	AsyncDownloadItem();
	~AsyncDownloadItem();

	void Init(void *cookie,IOleClientSite *pClientSite,ASYNCDOWNLOADCALLBACK pFuncCallBack);
	IOleClientSite *m_pClientSite;
	HRESULT StartDownload( LPOLESTR url);
	HRESULT CancelDownload();
	void Shutdown();

	BOOL m_bDownloadComplete;
	ULONG m_progress,m_progressMax;
	HRESULT m_error;
	DWORD m_readAmount;
	CHeapStream *m_pStream;

	int m_readyState;
protected:
	void OnDataReady(CHeapStream *pStream,DWORD readAmount);
	void OnDataPush(CHeapStream *pStream,DWORD readAmount);
	void OnProgress(ULONG ulProgress,ULONG ulProgressMax,ULONG ulStatusCode,LPCWSTR pwzStatusText);
	void OnDownloadError(HRESULT hr);
	HRESULT GetBindHost();
	HRESULT GetAMoniker(LPOLESTR url,IMoniker **ppmkr);
    IBindHost *m_host;
    ASYNCDOWNLOADCALLBACK m_pCallback;
	CDownloadSink * sink;
	IBindCtx * m_pBindCtx ;
	BOOL m_registration;


	void *m_cookie;
	friend class CDownloadSink;
};

class DDDownWin : public FWnd
{
public:
	DDDownWin(CDownloadSink *pOwner) :FWnd()
		{m_pOwner=pOwner;};
	BOOL CreateWin();
	virtual LRESULT WindowProc(UINT message,WPARAM wParam,LPARAM lParam);
protected:
	CDownloadSink *m_pOwner;
};


class CDownloadSink : public IBindStatusCallback
{
public:
    CDownloadSink(IUnknown *punkOuter,AsyncDownloadItem *);
    ~CDownloadSink();

    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut);
    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();

    STDMETHOD(OnStartBinding)(
            /* [in] */ DWORD grfBSCOption,
            /* [in] */ IBinding *pib);
    STDMETHOD(GetPriority)(
            /* [out] */ LONG *pnPriority);
    STDMETHOD(OnLowResource)(
            /* [in] */ DWORD reserved);
    STDMETHOD(OnProgress)(
            /* [in] */ ULONG ulProgress,
            /* [in] */ ULONG ulProgressMax,
            /* [in] */ ULONG ulStatusCode,
            /* [in] */ LPCWSTR szStatusText);
    STDMETHOD(OnStopBinding)(
            /* [in] */ HRESULT hresult,
            /* [in] */ LPCWSTR szError);

    STDMETHOD(GetBindInfo)(
            /* [out] */ DWORD *grfBINDINFOF,
            /* [unique][out][in] */ BINDINFO *pbindinfo);

    STDMETHOD(OnDataAvailable)(
            /* [in] */ DWORD grfBSCF,
            /* [in] */ DWORD dwSize,
            /* [in] */ FORMATETC *pformatetc,
            /* [in] */ STGMEDIUM *pstgmed);

    STDMETHOD(OnObjectAvailable)(
            /* [in] */ REFIID riid,
            /* [iid_is][in] */ IUnknown *punk);

    CDownloadSink * Next()                  { return(m_next); }
    void            Next(CDownloadSink *n)  { m_next = n; }

    DISPID          DispId()  { return(m_propId); }
    IBinding *      Binding() { return(m_binding); }

	DWORD readAmount;
	CHeapStream *		dataStream; // complete data stream based on hmem
		
	AsyncDownloadItem*   m_control;
    DISPID          m_propId;
        
	void Pull();
private:
	CDownloadSink * m_next;
	IBinding *      m_binding;
	DWORD           m_ref;
	IStream *		m_stream;
	BOOL cache_exit;
	
	CRITICAL_SECTION m_downloadCrit;
	BOOL m_bFinal;
	BOOL m_bCancelled;
	DDDownWin *m_pTempWin;
	DWORD m_readAmount;
	DWORD m_dwSize;
	DWORD m_curBufferSize;
	friend AsyncDownloadItem;
	friend DDDownWin;
};

#endif
