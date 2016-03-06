//---------------------------------------------------------------------
//  Copyright (c) 1995-1996, Data Dynamics. All rights reserved.
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//---------------------------------------------------------------------

#include "precomp.h"
#include "ddasync.h"
#include "streams.h"
#include "debug.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

UrlMonikerLib g_urllib;

CDownloadSink::CDownloadSink(IUnknown *punkOuter,AsyncDownloadItem *control)
{
	m_control = control;
    m_next    = 0;
    m_binding = 0;
    m_ref     = 0;
    m_stream  = 0;

	dataStream=NULL;
	readAmount=0;
	cache_exit=FALSE;

	InitializeCriticalSection(&m_downloadCrit);

	m_bFinal=FALSE;
	m_bCancelled=FALSE;

	m_pTempWin=NULL;
	m_readAmount=0;

	m_curBufferSize=16384;
}

CDownloadSink::~CDownloadSink()
{
	TRACE(1,"Download sink object delete\n");
	if( m_binding )
		m_binding->Release();
	if (dataStream!=NULL)
		dataStream->Release();
	if (m_stream)
	{
		m_stream->Release();
		m_stream=NULL;
	}

	DeleteCriticalSection(&m_downloadCrit);

	if (m_pTempWin!=NULL && m_pTempWin->IsWindow())
	{
		DestroyWindow(m_pTempWin->hWnd());
		delete m_pTempWin;
	}
}

STDMETHODIMP CDownloadSink::QueryInterface(const GUID &iid,void **ppv )
{
	if( IsEqualGUID(iid,IID_IUnknown) || 
		IsEqualGUID(iid,IID_IBindStatusCallback))
    {
		*ppv = this;
        AddRef();
        return NOERROR;
	}
	return ResultFromScode(E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CDownloadSink::AddRef()
{
	return(++m_ref);
}

STDMETHODIMP_(ULONG) CDownloadSink::Release()
{
	TRACE(1,"Download sink release\n");
	if(!--m_ref)
	{
		TRACE(1,"Download sink destroy\n");
		delete this;
        return(0);
	}
    return m_ref;
}

STDMETHODIMP CDownloadSink::GetBindInfo( DWORD *grfBINDF, BINDINFO *pbindInfo)
{
    *grfBINDF = BINDF_ASYNCHRONOUS | BINDF_ASYNCSTORAGE /*| BINDF_PULLDATA*/ | BINDF_GETNEWESTVERSION;
    return NOERROR;
}


STDMETHODIMP CDownloadSink::OnStartBinding(DWORD /*grfBSCOption*/,IBinding *pib)
{
	// BUGBUG: should check to see options are what we think they are
    m_binding = pib;
    pib->AddRef();
    return(NOERROR);
}

STDMETHODIMP CDownloadSink::GetPriority( LONG *pnPriority)
{
	return ResultFromScode(E_NOTIMPL);
}

STDMETHODIMP CDownloadSink::OnProgress(
     ULONG ulProgress,
     ULONG ulProgressMax,
     ULONG ulStatusCode,
     LPCWSTR pwzStatusText
)
{ 
	if (ulProgressMax!=0 && ulProgress==ulProgressMax)
		cache_exit=TRUE;
	m_control->OnProgress(ulProgress,
				ulProgressMax,ulStatusCode,pwzStatusText);
	return NOERROR;
}

STDMETHODIMP CDownloadSink::OnDataAvailable(DWORD grfBSCF,DWORD newsize,FORMATETC *pFmtetc,STGMEDIUM *pstgmed)
{
	HRESULT hr;
	
	m_dwSize=newsize;

    if( !m_stream )
	{
		m_stream = pstgmed->pstm;
		m_stream->AddRef();
	}
	
	if (dataStream==NULL)
	{
		dataStream=new CHeapStream();
		hr= (dataStream==NULL) ? E_OUTOFMEMORY: NOERROR;
		if (FAILED(hr))
			return hr;

		// Start DDDownloadWindow
		if (!m_pTempWin)
		{
			m_pTempWin=new DDDownWin(this);
			if (!m_pTempWin)
				return E_OUTOFMEMORY;
			m_pTempWin->CreateWin();
		}
	}

	// read data from m_stream into dataStream
	EnterCriticalSection(&m_downloadCrit);
	m_bFinal=(grfBSCF & BSCF_LASTDATANOTIFICATION) || cache_exit;
	LeaveCriticalSection(&m_downloadCrit);
	
	return NOERROR;
}

STDMETHODIMP CDownloadSink::OnObjectAvailable(REFIID riid,IUnknown *punk)
{
	return ResultFromScode(E_NOTIMPL);
}


STDMETHODIMP CDownloadSink::OnLowResource( DWORD reserved)
{
	// BUGBUG: really should have this kind of harsh policy on this ...
	m_binding->Abort();
	return(S_OK);
}

STDMETHODIMP CDownloadSink::OnStopBinding(HRESULT hrError, LPCWSTR szError)
{
	if (m_binding)
		m_binding->Release();
	m_binding = 0;
	if (FAILED(hrError) && hrError!=E_ABORT)
		m_control->OnDownloadError(hrError);

	return(NOERROR);
}


AsyncDownloadItem::AsyncDownloadItem()
{
	m_host=NULL;
	m_pClientSite=NULL;
	m_readyState=ASYNC_STATE_READY;
	sink=NULL;
	m_pBindCtx=NULL;
}

void AsyncDownloadItem::Init(void *cookie,IOleClientSite *pClientSite,ASYNCDOWNLOADCALLBACK pFuncCallBack)
{
	if (m_pClientSite)
		m_pClientSite->Release();
	m_pClientSite=pClientSite;
	m_pClientSite->AddRef();
	m_pCallback=pFuncCallBack;
	m_cookie=cookie;
	
}

void AsyncDownloadItem::Shutdown()
{
	CancelDownload();
	if (m_pClientSite)
		m_pClientSite->Release();
	m_pClientSite=NULL;
}


AsyncDownloadItem::~AsyncDownloadItem()
{
	Shutdown();
	if( m_host )
		m_host->Release();
}

HRESULT AsyncDownloadItem::GetBindHost()
{
	if( m_host ) // already loaded ?
		return(NOERROR);
    // Try service provider first...
	IServiceProvider * serviceProvider = 0;
    HRESULT hr = m_pClientSite->QueryInterface(IID_IServiceProvider,(void**)&serviceProvider);

    if( SUCCEEDED(hr) )
    {
		hr = serviceProvider->QueryService(SID_IBindHost,IID_IBindHost,(void**)&m_host);
		if( FAILED(hr) )
		{
	        // Some containers put IBindHost directly on the client site
			hr = serviceProvider->QueryInterface(IID_IBindHost,(void**)&m_host);
		}

        serviceProvider->Release();
    }
	
	if( FAILED(hr) )
    {
        // Some containers put IBindHost directly on the client site
        hr = m_pClientSite->QueryInterface(IID_IBindHost,(void**)&m_host);
		if (FAILED(hr))
		{
			IOleContainer *   container = 0;
			hr=m_pClientSite->GetContainer(&container);
			if (!FAILED(hr))
			{
				hr=container->QueryInterface(IID_IBindHost,(void **)&m_host);
				container->Release();
			}
		}
    }
    
	return(hr);
}




HRESULT AsyncDownloadItem::GetAMoniker( LPOLESTR url, IMoniker ** ppmkr )
{
	HRESULT hr = GetBindHost();
/*	if (FAILED(hr))
	{
		m_host= new _AfxBindHost();
		m_host->AddRef();
		hr=NOERROR;
	}*/
	if( SUCCEEDED(hr) )
	{
		hr = m_host->CreateMoniker(url,NULL, ppmkr,0);
	}

    if( FAILED(hr) )
    {
		// FUTURE: This really should be a call to MkParseDisplayNameEx!!!
		hr = g_urllib.CreateURLMoniker(0,url,ppmkr);
		// hr = ::MkParseDisplayNameEx(0, url, 0, ppmkr);
    }
	return( hr );
}

HRESULT AsyncDownloadItem::StartDownload( LPOLESTR url)
{
	IMoniker * pmkr=NULL;
	m_bDownloadComplete=FALSE;

	HRESULT hr = GetAMoniker( url, &pmkr );

	assert(m_pBindCtx==NULL);

	hr = ::CreateBindCtx(0,&m_pBindCtx);    

	sink=NULL;

	if( SUCCEEDED(hr) )
	{
		sink = new CDownloadSink(0,this);
        if( sink )
		{
			sink->AddRef();
		}
	}

	if( SUCCEEDED(hr) && !sink )
		hr = E_OUTOFMEMORY;

	m_registration=FALSE;
	if( SUCCEEDED(hr) )
	{
		hr = g_urllib.RegisterBindStatusCallback(m_pBindCtx, sink,0, 0) ;
		m_registration=TRUE;
	}

	IStream * strm = 0;

	if( SUCCEEDED(hr) )
		hr = pmkr->BindToStorage( m_pBindCtx, 0, IID_IStream, (void**)&strm );

	if( strm )
		strm->Release();

	if (FAILED(hr))
	{
		if (m_registration==TRUE)
			g_urllib.RevokeBindStatusCallback(m_pBindCtx,sink);
		pmkr->Release();
		if( m_pBindCtx )
		{
			m_pBindCtx->Release();
			m_pBindCtx=NULL;
		}
	}

	
	if (hr==NOERROR || hr==S_ASYNCHRONOUS || SUCCEEDED(hr))
		m_readyState=ASYNC_STATE_DOWNLOADING;
	return(hr);
}

HRESULT AsyncDownloadItem::CancelDownload()
{
	if (m_readyState==ASYNC_STATE_DOWNLOADING)
	{
		if (sink!=NULL && sink->m_binding!=NULL)
		{
			sink->m_binding->Abort();
			if (NULL!=sink && NULL!=sink->m_binding)
			{
				sink->m_binding->Release();
				sink->m_binding=NULL;
			}
		}

		if (sink)
		{
			if (m_registration==TRUE && m_pBindCtx!=NULL)
				g_urllib.RevokeBindStatusCallback(m_pBindCtx,sink);

			if (m_pBindCtx)
			{
				m_pBindCtx->Release();
				m_pBindCtx=NULL;
			}
		}
	}
	if (sink!=NULL)
	{
		sink->Release();
		sink=NULL;
	}
	m_readyState=ASYNC_STATE_READY;
	return NOERROR;
}

void AsyncDownloadItem::OnDataPush(CHeapStream *pStream,DWORD readAmount)
{
	m_pStream=pStream;
	m_readAmount=readAmount;
	(*m_pCallback)(m_cookie,this,ASYNC_MSG_DATAPUSH);
}

void AsyncDownloadItem::OnDataReady(CHeapStream *pStream,DWORD readAmount)
{
	m_pStream=pStream;
	m_readAmount=readAmount;
	(*m_pCallback)(m_cookie,this,ASYNC_MSG_DATAREADY);
	m_readyState=ASYNC_STATE_LOADED;

	if (m_registration==TRUE && m_pBindCtx!=NULL)
		g_urllib.RevokeBindStatusCallback(m_pBindCtx,sink);

	if (m_pBindCtx)
	{
		m_pBindCtx->Release();
		m_pBindCtx=NULL;
	}
}

void AsyncDownloadItem::OnProgress(ULONG ulProgress,ULONG ulProgressMax,ULONG ulStatusCode,LPCWSTR pwzStatusText)
{
	m_progress=ulProgress;
	m_progressMax=ulProgressMax;
	(*m_pCallback)(m_cookie,this,ASYNC_MSG_PROGRESS);
}

void AsyncDownloadItem::OnDownloadError(HRESULT hr)
{
	m_error=hr;
	(*m_pCallback)(m_cookie,this,ASYNC_MSG_ERROR);
}
	
//-------------------URL MONIKER DYNAMIC CALL ------------------------------------

UrlMonikerLib::UrlMonikerLib()
{
	urlmonlib=NULL;
	loadtry=FALSE;
}


UrlMonikerLib::~UrlMonikerLib()
{
	try
	{
		if (urlmonlib!=NULL)
			FreeLibrary(urlmonlib);
	}
	catch(...)
	{
	}
}


BOOL UrlMonikerLib::CheckLoaded()
{
	if (loadtry==FALSE)
	{
		loadtry=TRUE;
		urlmonlib=LoadLibrary("URLMON.DLL");
		if (urlmonlib==NULL)
			return FALSE;
		m_pCreateURLMoniker=(PFN_CREATEURLMONIKER)GetProcAddress(urlmonlib,"CreateURLMoniker");
		m_pRegisterBindStatusCallback=(PFN_REGISTERBINDSTATUSCALLBACK)GetProcAddress(urlmonlib,"RegisterBindStatusCallback");
		m_pRevokeBindStatusCallback=(PFN_REVOKEBINDSTATUSCALLBACK)GetProcAddress(urlmonlib,"RevokeBindStatusCallback");
	}
	if (urlmonlib==NULL)
		return FALSE;
	return TRUE;
}

HRESULT UrlMonikerLib::CreateURLMoniker(LPMONIKER pMkCtx, LPCWSTR szURL, LPMONIKER FAR * ppmk)
{
	if (CheckLoaded()==FALSE)
		return E_FAIL;
	if (m_pCreateURLMoniker==NULL)
		return E_FAIL;
	return m_pCreateURLMoniker(pMkCtx, szURL, ppmk);
}

HRESULT UrlMonikerLib::RegisterBindStatusCallback(LPBC pBC, IBindStatusCallback *pBSCb,
                                IBindStatusCallback**  ppBSCBPrev, DWORD dwReserved)
{
	if (CheckLoaded()==FALSE)
		return E_FAIL;
	if (m_pRegisterBindStatusCallback==NULL)
		return E_FAIL;
	return m_pRegisterBindStatusCallback(pBC,pBSCb,ppBSCBPrev, dwReserved);
}

HRESULT UrlMonikerLib::RevokeBindStatusCallback(LPBC pBC, IBindStatusCallback *pBSCb)
{
	if (CheckLoaded()==FALSE)
		return E_FAIL;
	if (m_pRevokeBindStatusCallback==NULL)
		return E_FAIL;
	return m_pRevokeBindStatusCallback(pBC,pBSCb);
}

void CDownloadSink::Pull()
{
	HRESULT hr;
	DWORD bytesRead,bytesWritten;

	if (m_dwSize!=0 && m_readAmount<m_dwSize)
	{
		char *buffer;

		DWORD buffersize=m_dwSize;
		if (buffersize>m_curBufferSize)
			buffersize=m_curBufferSize;
		
		if (m_curBufferSize<(128*1024))
			m_curBufferSize*=2;

		buffer=(char *)malloc(buffersize);

		if (buffer==NULL)
			return;
		
		if ((buffersize+m_readAmount)>m_dwSize)
			buffersize=m_dwSize-m_readAmount;

		hr=m_stream->Read(buffer,buffersize,&bytesRead);
		if (bytesRead!=0)
		{
			m_readAmount+=bytesRead;
			if (bytesRead!=0)
			{
				LARGE_INTEGER dMove; dMove.LowPart=0;dMove.HighPart=0;
				ULARGE_INTEGER backupPos;
				dataStream->Seek(dMove,STREAM_SEEK_CUR,&backupPos);
				dataStream->Seek(dMove,STREAM_SEEK_END,NULL);
				dataStream->Write(buffer,bytesRead,&bytesWritten);
				dMove.LowPart=backupPos.LowPart;
				dMove.HighPart=backupPos.HighPart;
				dataStream->Seek(dMove,STREAM_SEEK_SET,&backupPos);
			}
			
			m_control->OnDataPush(dataStream,readAmount);
		}

		free(buffer);
	}

    if (m_bFinal && m_control->m_bDownloadComplete==FALSE && m_readAmount>=m_dwSize)
	{
		LARGE_INTEGER moveAmount;
//		ULARGE_INTEGER newPosition;
		moveAmount.LowPart=0;
		moveAmount.HighPart=0;
		m_control->m_bDownloadComplete=TRUE;
		//dataStream->Seek(moveAmount,STREAM_SEEK_SET,&newPosition);
		m_control->OnDataReady(dataStream,readAmount);
		dataStream->Release();
		dataStream=NULL;
		if (m_binding)
			m_binding->Abort();
	}
}



BOOL DDDownWin::CreateWin()
{
	RegisterWindow(DD_WNDCLASS("DownPump"),0,NULL,NULL);
	if (CreateEx(0,"",
		WS_POPUP,
		0,0,0,0,GetDesktopWindow(),NULL,NULL)==0)
		return FALSE;
	
	SetTimer(m_hWnd,10,200,NULL);

	return TRUE;
}

LRESULT DDDownWin::WindowProc(UINT message,WPARAM wParam,LPARAM lParam)
{
	switch(message)
	{
	case WM_TIMER:
		{
			try
			{
				m_pOwner->Pull();
				if ((m_pOwner->m_bFinal && m_pOwner->m_readAmount>=m_pOwner->m_dwSize) || m_pOwner->m_bCancelled)
					KillTimer(m_hWnd,10);
			}
			catch(...)
			{
			}
		}
		break;
	case WM_NCDESTROY:
		break;
	}
	return FWnd::WindowProc(message,wParam,lParam);
}

