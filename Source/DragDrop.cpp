#include "precomp.h"
#include "Interfaces.h"
#include "Debug.h"
#include "Globals.h"
#include "Resource.h"
#include "Resource.h"
#include "Designer.h"
#include "Dock.h"
#include "DragDrop.h"

extern HINSTANCE g_hInstance;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// SetDefFormatEtc
//

void SetDefFormatEtc(FORMATETC& fe, CLIPFORMAT& cf, DWORD dwMed)
{
    fe.cfFormat = cf;
    fe.dwAspect = DVASPECT_CONTENT;
    fe.ptd = 0;
    fe.tymed = dwMed;
	fe.lindex = -1;
}

//
// CToolDropSource
//

//
// QueryInterface
//

HRESULT __stdcall CToolDropSource::QueryInterface(REFIID riid, void** ppvObject)
{
	if (riid == IID_IDropSource || riid == IID_IUnknown)
	{
		*ppvObject = static_cast<IDropSource*>(this);
		AddRef();
		return S_OK;
	}
	*ppvObject = 0;
	return E_NOINTERFACE;
}
        
//
// AddRef
//

ULONG __stdcall CToolDropSource::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

//
// Release
//

ULONG __stdcall CToolDropSource::Release()
{
	if (InterlockedDecrement(&m_cRef) == 0)
	{
		delete this;
		return 0;
	}
	return m_cRef;
}

//
// QueryContinueDrag
//

HRESULT __stdcall CToolDropSource::QueryContinueDrag(BOOL  bEscapePressed,
													 DWORD grfKeyState)
{
	if (bEscapePressed)
		return DRAGDROP_S_CANCEL;

	if (!(grfKeyState & MK_LBUTTON))
		return DRAGDROP_S_DROP;

	return S_OK;
}
        
//
// GiveFeedback
//

HRESULT __stdcall CToolDropSource::GiveFeedback(DWORD dwEffect)
{
	switch (dwEffect)
	{
	case DROPEFFECT_COPY:
		SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOLCOPY)));
		return S_OK;
	case DROPEFFECT_MOVE:
		SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOLMOVE)));
		return S_OK;
	case DROPEFFECT_NONE:
		SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOLNOT)));
		return S_OK;
	}
	return DRAGDROP_S_USEDEFAULTCURSORS;
}


//
// CBandDropTarget
//

CBandDropTarget::~CBandDropTarget()
{
}

HRESULT __stdcall CBandDropTarget::QueryInterface(REFIID riid, void** ppvObject)
{
	if (riid == IID_IDropTarget || riid == IID_IUnknown)
	{
		*ppvObject = static_cast<IDropTarget*>(this);
		AddRef();
		return S_OK;
	}
	*ppvObject = 0;
	return E_NOINTERFACE;
}
        
ULONG __stdcall CBandDropTarget::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

ULONG __stdcall CBandDropTarget::Release()
{
	if (InterlockedDecrement(&m_cRef) == 0)
	{
		delete this;
		return 0;
	}
	return m_cRef;
}

//
// DragEnter
//

HRESULT __stdcall CBandDropTarget::DragEnter(IDataObject* pDataObject, 
											 DWORD		  grfKeyState,
											 POINTL		  pt,
											 DWORD*		  pdwEffect)
{
	DWORD dwEffect = *pdwEffect;
	*pdwEffect = DROPEFFECT_NONE;

	if (m_pDataObject)
	{
		m_pDataObject->Release();
		m_pDataObject = NULL;
	}
	m_pDataObject = pDataObject;
	if (NULL == m_pDataObject)
		return E_FAIL;

	m_pDataObject->AddRef();

	POINT pt2;
	pt2.x = pt.x;
	pt2.y = pt.y;
	if (m_pBandDragDrop->DragEnter(pt2))
		*pdwEffect = DROPEFFECT_COPY|DROPEFFECT_MOVE;
	return NOERROR;
}

//
// DragOver
//

HRESULT __stdcall CBandDropTarget::DragOver(DWORD  grfKeyState,
											POINTL pt,
											DWORD* pdwEffect)
{
	DWORD dwEffect = *pdwEffect;
	*pdwEffect = DROPEFFECT_NONE;
	UINT nFlags = 0; 
	if (NULL != m_pDataObject)
	{
		POINT pt2;
		pt2.x = pt.x;
		pt2.y = pt.y;
		if (m_pBandDragDrop->DragOver(pt2))
			*pdwEffect = DROPEFFECT_COPY|DROPEFFECT_MOVE;
	}
	return NOERROR;
}
        
//
// DragLeave
//

HRESULT __stdcall CBandDropTarget::DragLeave()
{
	if (NULL != m_pDataObject)
	{
		m_pDataObject->Release();
		m_pDataObject = NULL;
	}
	m_pBandDragDrop->DragLeave();
	return NOERROR;
}
        
//
// Drop
//

HRESULT __stdcall CBandDropTarget::Drop(IDataObject* pDataObject,
									    DWORD		 grfKeyState,
										POINTL		 pt,
										DWORD*		 pdwEffect)
{
	POINT pt2;
	pt2.x = pt.x;
	pt2.y = pt.y;
	*pdwEffect = DROPEFFECT_MOVE;
	if (m_pBandDragDrop->Drop(pt2, pDataObject))
	{
		if (!(0x8000 & GetKeyState(VK_CONTROL)))
			*pdwEffect = DROPEFFECT_COPY;
	}
	m_pDataObject->Release();
	m_pDataObject = NULL;
	return NOERROR;
}

//
// BandDragDrop
//

BandDragDrop::~BandDragDrop()
{
	if (m_pTool)
		m_pTool->Release();
}

//
// DragEnter
//

BOOL BandDragDrop::DragEnter(const POINT& pt)
{
	m_pBar->m_diDropInfo.nDropIndex = -1;

	//
	// If the current tool has subbands. We cannot drop it into subbands so we have 
	// to close down any open popups
	//

	if (NULL == m_pBar->m_pCustomizeDragLock && 
		m_pBar->m_diCustSelection.pTool &&
		m_pBar->m_diCustSelection.pTool == (CTool*)m_pTool && 
		m_pBar->m_diCustSelection.pTool->HasSubBand())
	{
		m_pBar->m_diCustSelection.pBand->SetPopupIndex(-1);
		m_pBar->m_pCustomizeDragLock = m_pBar->m_diCustSelection.pTool;
	}

	POINT pt2 = pt;
	m_pWnd->ScreenToClient(pt2);
	CBand* pBand = GetBand(pt2);
	if (pBand)
	{
		OffsetPoint(pBand, pt2);
		pBand->CalcDropLoc(pt2, &m_pBar->m_diDropInfo);
		if (-1 != m_pBar->m_diDropInfo.nDropIndex)
		{
			pBand->SetPopupIndex(m_pBar->m_diDropInfo.nDropIndex + (1 == m_pBar->m_diDropInfo.nDropDir));
			m_pBar->DrawDropSign(m_pBar->m_diDropInfo);
		}
		return TRUE;
	}
	return FALSE;
}

//
// DragOver
//

BOOL BandDragDrop::DragOver(const POINT& pt)
{
	POINT pt2 = pt;
	m_pWnd->ScreenToClient(pt2);
	CBand* pBand = GetBand(pt2);
	if (pBand)
	{
		try
		{
			CBar::DropInfo diNew;
			OffsetPoint(pBand, pt2);
			pBand->CalcDropLoc(pt2, &diNew);
			if (0 != memcmp(&diNew, &m_pBar->m_diDropInfo, sizeof(CBar::DropInfo)))
			{
				if (-1 != m_pBar->m_diDropInfo.nDropIndex)
				{
					m_pBar->m_diDropInfo.nDropIndex = -1;
					m_pBar->DrawDropSign(m_pBar->m_diDropInfo);

				}
				memcpy(&m_pBar->m_diDropInfo, &diNew, sizeof(CBar::DropInfo));
				if (-1 != m_pBar->m_diDropInfo.nDropIndex)
				{
					m_pBar->m_diDropInfo.pBand->SetPopupIndex(m_pBar->m_diDropInfo.nDropIndex + (1 == m_pBar->m_diDropInfo.nDropDir));
					m_pBar->DrawDropSign(m_pBar->m_diDropInfo);
				}
			}
		}
		catch (SEException& e)
		{
			assert(FALSE);
			e.ReportException(__FILE__, __LINE__);
		}
		return TRUE;
	}
	return FALSE;
}

//
// DragLeave
//

BOOL BandDragDrop::DragLeave()
{
	if (-1 != m_pBar->m_diDropInfo.nDropIndex)
	{
		m_pBar->m_diDropInfo.nDropIndex = -1;
		m_pBar->DrawDropSign(m_pBar->m_diDropInfo);
	}
	return TRUE;
}

//
// Drop
//

BOOL BandDragDrop::Drop(const POINT& pt, IDataObject* pDataObject)
{
	POINT pt2 = pt;
	m_pWnd->ScreenToClient(pt2);
	if (-1 != m_pBar->m_diDropInfo.nDropIndex)
	{
		m_pBar->m_diDropInfo.nDropIndex = -1;
		m_pBar->DrawDropSign(m_pBar->m_diDropInfo);
	}
	CBand* pBand = GetBand(pt2);
	if (pBand)
	{
		if (CBand::eDropRight == m_pBar->m_diDropInfo.nDropDir || 
			CBand::eDropBottom == m_pBar->m_diDropInfo.nDropDir)
		{
			m_pBar->m_diDropInfo.nToolIndex++;
		}

		STGMEDIUM stm;
		FORMATETC fe[2];
		SetDefFormatEtc(fe[0], GetGlobals().m_nIDClipToolIdFormat, TYMED_ISTREAM);
		SetDefFormatEtc(fe[1], GetGlobals().m_nIDClipBandToolIdFormat, TYMED_ISTREAM);
		for (int nIndex = 0; nIndex < 2 && NULL == m_pTool; nIndex++)
		{
			if (SUCCEEDED(pDataObject->GetData(&fe[nIndex], &stm)))
			{
				LARGE_INTEGER lnMoveAmount; 
				lnMoveAmount.LowPart = lnMoveAmount.HighPart = 0;
				ULARGE_INTEGER lnCurPos;

				HRESULT hResult = stm.pstm->Seek(lnMoveAmount, STREAM_SEEK_SET, &lnCurPos);
				if (FAILED(hResult))
					return hResult;

				switch (nIndex)
				{
				case 0:
					hResult = m_pBar->ExchangeToolById((LPDISPATCH)stm.pstm, VARIANT_FALSE, (ITool**)&m_pTool);
					break;

				case 1:
					{
						BSTR bstrBand = NULL;
						BSTR bstrPage = NULL;
						hResult = m_pBar->ExchangeToolByBandPageToolId((LPDISPATCH)stm.pstm, bstrBand, bstrPage, VARIANT_FALSE, (ITool**)&m_pTool);
					}
					break;
				}
			}
		}
		if (NULL == m_pTool)
			return FALSE;

		HRESULT hResult;
		ITools* pTools = NULL;
		pBand->GetActiveTools((CTools**)&pTools);
		if (NULL == pTools)
			goto Cleanup;

		short nToolCount;
		hResult = pTools->Count(&nToolCount);
		if (FAILED(hResult))
			goto Cleanup;
			
		ITool* pCmpTool;
		VARIANT vi;
		vi.vt = VT_I2;
		for (vi.iVal = 0; vi.iVal < nToolCount; vi.iVal++)
		{
			hResult = pTools->Item(&vi, (Tool**)&pCmpTool);
			if (SUCCEEDED(hResult) && NULL != pCmpTool)
			{
				if (pCmpTool == m_pTool)
				{
					//
					// Source equals Destination
					//

					pCmpTool->Release();
					// Remove selected item first (only 1)
					pTools->Remove(&vi);

					// reduce drop index if prev item deleted
					if (m_pBar->m_diDropInfo.nToolIndex > vi.iVal) 
						m_pBar->m_diDropInfo.nToolIndex--; 
					break;
				}
				pCmpTool->Release();
			}
		}
		pTools->Release();
		if (SUCCEEDED(hResult))
		{
			hResult = pBand->DSGInsertTool(m_pBar->m_diDropInfo.nToolIndex, m_pTool);
			if (SUCCEEDED(hResult))
			{
				hResult = m_pBar->RecalcLayout();
				m_pBar->SetModified();
				if (m_pBar->m_pDesigner)
					m_pBar->m_pDesigner->SetDirty();
				m_pTool->Release();
				m_pTool = NULL;
				return TRUE;
			}
		}
Cleanup:
		if (pTools)
			pTools->Release();
		m_pTool->Release();
		m_pTool = NULL;
	}
	return FALSE;
}
