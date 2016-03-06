//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Interfaces.h"
#include "Debug.h"
#include "Globals.h"
#include "Resource.h"
#include "ChildBands.h"
#include "Dock.h"
#include "Bands.h"
#include "Designer\DragDrop.h"
#include "PopupWin.h"
#include "DropSource.h"

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

CBandDropTarget::CBandDropTarget(long nDestination)
	: m_pBandDragDrop(NULL),
	  m_pDataObject(NULL),
	  m_pTool(NULL)
{
	m_cRef = 1;
	SetDefFormatEtc(m_fe[0], GetGlobals().m_nIDClipToolIdFormat, TYMED_ISTREAM);
	SetDefFormatEtc(m_fe[1], GetGlobals().m_nIDClipBandToolIdFormat, TYMED_ISTREAM);
	SetDefFormatEtc(m_fe[2], GetGlobals().m_nIDClipBandChildBandToolIdFormat, TYMED_ISTREAM);
	SetDefFormatEtc(m_fe[3], GetGlobals().m_nIDClipBandFormat, TYMED_ISTREAM);
	SetDefFormatEtc(m_fe[4], GetGlobals().m_nIDClipCategoryFormat, TYMED_ISTREAM);
	SetDefFormatEtc(m_fe[5], GetGlobals().m_nIDClipToolFormat, TYMED_ISTREAM);
}

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
	m_nFormatType = -1;
	*pdwEffect = DROPEFFECT_NONE;
	m_nIndex = -1;

	if (m_pDataObject)
	{
		m_pDataObject->Release();
		m_pDataObject = NULL;
	}

	m_pDataObject = pDataObject;
	if (NULL == m_pDataObject)
		return E_FAIL;

	m_pDataObject->AddRef();

	if (m_pTool)
	{
		//
		// This is the first tool used to calc the drop position
		// 

		m_pTool->Release();
		m_pTool = NULL;
	}

	CBar* pDestBar = (CBar*)m_pDestBar;
	pDestBar->m_diDropInfo.nDropIndex = -1;

	ULARGE_INTEGER lnCurPos;
	LARGE_INTEGER  lnMoveAmount; 
	STGMEDIUM	   stm;
	memset(&stm, NULL, sizeof(stm));
	HRESULT		   hResult;
	int			   nCount;

	for (int nIndex = 0; nIndex < eNumOfTypes; nIndex++)
	{
		if (FAILED(m_pDataObject->GetData(&m_fe[nIndex], &stm)))
			continue;

		lnMoveAmount.LowPart = lnMoveAmount.HighPart = 0;

		hResult = stm.pstm->Seek(lnMoveAmount, STREAM_SEEK_SET, &lnCurPos);
		if (FAILED(hResult))
		{
			stm.pstm->Release();
			return hResult;
		}

		hResult = stm.pstm->Read(&m_nSrcId, sizeof(m_nSrcId), NULL);
		if (FAILED(hResult))
		{
			stm.pstm->Release();
			return hResult;
		}

		hResult = stm.pstm->Read(&m_pSrcBar, sizeof(m_pSrcBar), NULL);
		if (FAILED(hResult))
		{
			stm.pstm->Release();
			return hResult;
		}

		m_nFormatType = m_fe[nIndex].cfFormat;
		m_nIndex = nIndex;
		break;
	}

	POINT pt2 = {pt.x, pt.y};
	CBand* pBand = m_pBandDragDrop->GetBand(pt2);

	if (-1 != m_nIndex)
	{
		switch (m_nIndex)
		{
		case CToolDataObject::eTool:
			{
				hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
				if (FAILED(hResult))
				{
					stm.pstm->Release();
					return hResult;
				}

				if (nCount > 0)
				{
					m_pTool = CTool::CreateInstance(NULL);
					if (m_pTool)
					{
						((CTool*)m_pTool)->SetBar((CBar*)m_pDestBar);
						
						hResult = m_pTool->DragDropExchange(stm.pstm, VARIANT_FALSE);
						if (FAILED(hResult))
						{
							stm.pstm->Release();
							m_pTool->Release();
							m_pTool = NULL;
							return hResult;
						}

						if (CToolDataObject::eDesignerMainToolDragDropId == m_nSrcId || CToolDataObject::eDesignerDragDropId == m_nSrcId || CToolDataObject::eLibraryDragDropId == m_nSrcId)
							*pdwEffect = DROPEFFECT_COPY;
						else if (MK_CONTROL & grfKeyState)
							*pdwEffect = DROPEFFECT_COPY;
						else
							*pdwEffect = DROPEFFECT_MOVE;
					}
				}
			}
			break;

		case CToolDataObject::eToolId:
			{
				if (m_pSrcBar != m_pDestBar)
				{
					stm.pstm->Release();
					return NOERROR;
				}

				hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
				if (FAILED(hResult))
				{
					stm.pstm->Release();
					return hResult;
				}

				if (nCount > 0)
				{
					m_pTool = CTool::CreateInstance(NULL);
					if (m_pTool)
					{
						hResult = ((CBar*)m_pDestBar)->ExchangeToolByIdentity((LPDISPATCH)stm.pstm, VARIANT_FALSE, (IDispatch**)&m_pTool);
						if (FAILED(hResult))
						{
							stm.pstm->Release();
							m_pTool->Release();
							m_pTool = NULL;
							return hResult;
						}
						if (CToolDataObject::eDesignerMainToolDragDropId == m_nSrcId || CToolDataObject::eDesignerDragDropId == m_nSrcId || CToolDataObject::eLibraryDragDropId == m_nSrcId)
							*pdwEffect = DROPEFFECT_COPY;
						else if (MK_CONTROL & grfKeyState)
							*pdwEffect = DROPEFFECT_COPY;
						else
							*pdwEffect = DROPEFFECT_MOVE;
					}
				}
			}
			break;

		case CToolDataObject::eBandToolId:
		case CToolDataObject::eBandChildBandToolId:
			{
				if (m_pSrcBar != m_pDestBar)
				{
					stm.pstm->Release();
					return NOERROR;
				}

				hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
				if (FAILED(hResult))
				{
					stm.pstm->Release();
					return hResult;
				}

				if (nCount > 0)
				{
					BSTR bstrBand = NULL;
					BSTR bstrChildBand = NULL;
					hResult = ((CBar*)m_pDestBar)->ExchangeToolByBandChildBandToolIdentity((LPDISPATCH)stm.pstm, 
																							 bstrBand, 
																							 bstrChildBand, 
																							 VARIANT_FALSE, 
																							 (IDispatch**)&m_pTool);
					SysFreeString(bstrBand);
					SysFreeString(bstrChildBand);

					if (FAILED(hResult))
					{
						stm.pstm->Release();
						return hResult;
					}
					
					if (CToolDataObject::eDesignerMainToolDragDropId == m_nSrcId || CToolDataObject::eDesignerDragDropId == m_nSrcId || CToolDataObject::eLibraryDragDropId == m_nSrcId)
						*pdwEffect = DROPEFFECT_COPY;
					else if (MK_CONTROL & grfKeyState)
						*pdwEffect = DROPEFFECT_COPY;
					else
						*pdwEffect = DROPEFFECT_MOVE;
				}
			}
			break;

		case CToolDataObject::eBand:
			{
				if (NULL == pBand)
					*pdwEffect = DROPEFFECT_COPY;
			}
			break;
		
		case CToolDataObject::eCategory:
			{
			}
			break;
		}
		stm.pstm->Release();
	}

	if (pBand)
	{
		m_pBandDragDrop->OffsetPoint(pBand, pt2);
		
		pBand->CalcDropLoc(pt2, &pDestBar->m_diDropInfo);
		
		if (-1 != pDestBar->m_diDropInfo.nDropIndex)
		{
			pDestBar->DrawDropSign(pDestBar->m_diDropInfo);
			pBand->SetPopupIndex(pDestBar->m_diDropInfo.nDropIndex + (1 == pDestBar->m_diDropInfo.nDropDir));
		}
	}

	//
	// If the current tool has subbands. We cannot drop it into subbands so we have 
	// to close down any open popups
	//

	if (NULL == pDestBar->m_pCustomizeDragLock && 
		pDestBar->m_diCustSelection.pTool &&
		pDestBar->m_diCustSelection.pTool == m_pTool && 
		pDestBar->m_diCustSelection.pTool->HasSubBand())
	{
		pDestBar->m_diCustSelection.pBand->SetPopupIndex(-1);
		pDestBar->m_pCustomizeDragLock = pDestBar->m_diCustSelection.pTool;
	}

	return NOERROR;
}

//
// DragOver
//

HRESULT __stdcall CBandDropTarget::DragOver(DWORD  grfKeyState,
											POINTL pt,
											DWORD* pdwEffect)
{
	*pdwEffect = DROPEFFECT_NONE;
	UINT  nFlags = 0; 
	CBar* pDestBar = (CBar*)m_pDestBar;

	if (NULL == pDestBar->m_pCustomizeDragLock && 
		pDestBar->m_diCustSelection.pTool &&
		pDestBar->m_diCustSelection.pTool == m_pTool && 
		pDestBar->m_diCustSelection.pTool->HasSubBand())
	{
		pDestBar->m_diCustSelection.pBand->SetPopupIndex(-1);
		pDestBar->m_pCustomizeDragLock = pDestBar->m_diCustSelection.pTool;
	}

	if ((CToolDataObject::eToolId == m_nIndex || CToolDataObject::eBandToolId == m_nIndex  || CToolDataObject::eBandChildBandToolId == m_nIndex) && m_pSrcBar != m_pDestBar)
		return NOERROR;

	CBar::DropInfo diNew;
	POINT pt2 = {pt.x, pt.y};
	CBand* pBand = m_pBandDragDrop->GetBand(pt2);
	if (pBand)
	{
		m_pBandDragDrop->OffsetPoint(pBand, pt2);
		pBand->CalcDropLoc(pt2, &diNew);

		switch (m_nIndex)
		{
		case CToolDataObject::eTool:
		case CToolDataObject::eToolId:
		case CToolDataObject::eBandToolId:
		case CToolDataObject::eBandChildBandToolId:
			{
				if (0 != memcmp(&diNew, &pDestBar->m_diDropInfo, sizeof(CBar::DropInfo)))
				{
					if (-1 != pDestBar->m_diDropInfo.nDropIndex)
					{
						pDestBar->m_diDropInfo.nDropIndex = -1;
						pDestBar->DrawDropSign(pDestBar->m_diDropInfo);
					}

					memcpy(&pDestBar->m_diDropInfo, &diNew, sizeof(CBar::DropInfo));
					
					if (-1 != pDestBar->m_diDropInfo.nDropIndex && pDestBar->m_diDropInfo.pBand)
					{
						pDestBar->m_diDropInfo.pBand->SetPopupIndex(pDestBar->m_diDropInfo.nDropIndex + (1 == pDestBar->m_diDropInfo.nDropDir));
						pDestBar->DrawDropSign(pDestBar->m_diDropInfo);
					}
				}
				int nToolCount = pBand->m_pTools->GetToolCount();
				if (nToolCount > 0 && pDestBar->m_diDropInfo.pBand)
				{
					pDestBar->m_diDropInfo.pBand->SetPopupIndex(pDestBar->m_diDropInfo.nDropIndex + (1 == pDestBar->m_diDropInfo.nDropDir));
					pDestBar->DrawDropSign(pDestBar->m_diDropInfo);
				}

				if (NULL == pDestBar->m_diDropInfo.pBand)
					*pdwEffect = DROPEFFECT_NONE;
				else if (CToolDataObject::eDesignerMainToolDragDropId == m_nSrcId || CToolDataObject::eDesignerDragDropId == m_nSrcId || CToolDataObject::eLibraryDragDropId == m_nSrcId)
					*pdwEffect = DROPEFFECT_COPY;
				else if (MK_CONTROL & grfKeyState)
					*pdwEffect = DROPEFFECT_COPY;
				else
					*pdwEffect = DROPEFFECT_MOVE;
			}
			break;
		}
	}
	else
	{
		switch (m_nIndex)
		{
		case CToolDataObject::eBand:
			if (-1 != pDestBar->m_diDropInfo.nDropIndex)
			{
				pDestBar->m_diDropInfo.nDropIndex = -1;
				pDestBar->DrawDropSign(pDestBar->m_diDropInfo);
				memcpy(&pDestBar->m_diDropInfo, &diNew, sizeof(CBar::DropInfo));
			}
			*pdwEffect = DROPEFFECT_COPY;
			break;

		case CToolDataObject::eCategory:
			break;
		}
	}
	return NOERROR;
}
        
//
// DragLeave
//

HRESULT __stdcall CBandDropTarget::DragLeave()
{
	CBar* pDestBar = (CBar*)m_pDestBar;
	if (-1 != pDestBar->m_diDropInfo.nDropIndex)
		pDestBar->DrawDropSign(pDestBar->m_diDropInfo);
	if (-1 != pDestBar->m_diDropInfo.nDropIndex)
	{
		pDestBar->m_diDropInfo.nDropIndex = -1;
		pDestBar->DrawDropSign(pDestBar->m_diDropInfo);
	}
	if (m_pDataObject)
	{
		m_pDataObject->Release();
		m_pDataObject = NULL;
	}
	if (m_pTool)
	{
		m_pTool->Release();
		m_pTool = NULL;
	}
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
	*pdwEffect = DROPEFFECT_NONE;
	CBar* pDestBar = (CBar*)m_pDestBar;
	
	if ((CToolDataObject::eToolId == m_nIndex || CToolDataObject::eBandToolId == m_nIndex  || CToolDataObject::eBandChildBandToolId == m_nIndex) && m_pSrcBar != m_pDestBar)
		return NOERROR;

	if (-1 != pDestBar->m_diDropInfo.nDropIndex)
	{
		pDestBar->m_diDropInfo.nDropIndex = -1;
		pDestBar->DrawDropSign(pDestBar->m_diDropInfo);
	}

	ULARGE_INTEGER lnCurPos;
	LARGE_INTEGER lnMoveAmount; 
	STGMEDIUM stm;
	memset(&stm, NULL, sizeof(stm));
	HRESULT hResult;
	CTool* pTool;
	int nCount;

	CBand* pBand = m_pBandDragDrop->GetBand(pt2);
	if (pBand)
	{
		if (ddBFCustomize & pBand->bpV1.m_dwFlags)
		{
			if (ddCBSNone != pBand->bpV1.m_cbsChildStyle)
				pBand = pBand->m_pChildBands->GetCurrentChildBand();

			if (CBand::eDropRight == pDestBar->m_diDropInfo.nDropDir || CBand::eDropBottom == pDestBar->m_diDropInfo.nDropDir)
				pDestBar->m_diDropInfo.nToolIndex++;

			if (SUCCEEDED(m_pDataObject->GetData(&m_fe[m_nIndex], &stm)))
			{
				lnMoveAmount.LowPart = lnMoveAmount.HighPart = 0;

				hResult = stm.pstm->Seek(lnMoveAmount, STREAM_SEEK_SET, &lnCurPos);
				if (FAILED(hResult))
					return hResult;

				hResult = stm.pstm->Read(&m_nSrcId, sizeof(m_nSrcId), NULL);
				if (FAILED(hResult))
					return hResult;

				hResult = stm.pstm->Read(&m_pSrcBar, sizeof(m_pSrcBar), NULL);
				if (FAILED(hResult))
					return hResult;

				switch (m_nIndex)
				{
				case CToolDataObject::eTool:
					{
						hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
						if (FAILED(hResult))
							return hResult;

						for (int nTool = 0; nTool < nCount; nTool++)
						{
							pTool = CTool::CreateInstance(NULL);
							assert(pTool);
							if (pTool)
							{
								pTool->SetBar(pDestBar);
								hResult = pTool->DragDropExchange(stm.pstm, VARIANT_FALSE);
								if (FAILED(hResult))
								{
									pTool->Release();
									continue;
								}
								hResult = pBand->DSGInsertTool(pDestBar->m_diDropInfo.nToolIndex + nTool, pTool);
								pTool->Release();
							}
						}
					}
					break;

				case CToolDataObject::eToolId:
					{
						hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
						if (FAILED(hResult))
							return hResult;

						for (int nTool = 0; nTool < nCount; nTool++)
						{
							pTool = CTool::CreateInstance(NULL);
							assert(pTool);
							if (pTool)
							{
								hResult = ((CBar*)m_pDestBar)->ExchangeToolByIdentity((LPDISPATCH)stm.pstm, 
																					  VARIANT_FALSE, 
																					  (IDispatch**)&pTool);
								if (FAILED(hResult))
								{
									pTool->Release();
									continue;
								}

								hResult = pBand->DSGInsertTool(pDestBar->m_diDropInfo.nToolIndex + nTool, pTool);
								pTool->Release();
							}
						}
					}
					break;

				case CToolDataObject::eBandToolId:
				case CToolDataObject::eBandChildBandToolId:
					{
						hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
						if (FAILED(hResult))
							return hResult;

						BSTR bstrBand = NULL;
						BSTR bstrChildBand = NULL;
						for (int nTool = 0; nTool < nCount; nTool++)
						{
							hResult = ((CBar*)m_pDestBar)->ExchangeToolByBandChildBandToolIdentity((LPDISPATCH)stm.pstm, 
																							       bstrBand, 
																							       bstrChildBand, 
																							       VARIANT_FALSE, 
																							       (IDispatch**)&pTool);
							if (FAILED(hResult))
							{
								pTool->Release();
								continue;
							}

							hResult = pBand->DSGInsertTool(pDestBar->m_diDropInfo.nToolIndex + nTool, pTool);
							pTool->Release();
							SysFreeString(bstrBand);
							SysFreeString(bstrChildBand);
						}
					}
					break;
				}
				stm.pstm->Release();
			}

			if (SUCCEEDED(hResult))
			{
				pDestBar->SetModified();
				
				if (pDestBar->m_pDesigner)
					pDestBar->m_pDesigner->SetDirty();
				
				if (CToolDataObject::eDesignerMainToolDragDropId == m_nSrcId || CToolDataObject::eDesignerDragDropId == m_nSrcId || CToolDataObject::eLibraryDragDropId == m_nSrcId)
					*pdwEffect = DROPEFFECT_COPY;
				else if (MK_CONTROL & grfKeyState)
					*pdwEffect = DROPEFFECT_COPY;
				else
					*pdwEffect = DROPEFFECT_MOVE;
				
				hResult = pDestBar->RecalcLayout();

				pBand->Refresh();
			}
		}
	}
	else
	{
		if (SUCCEEDED(m_pDataObject->GetData(&m_fe[m_nIndex], &stm)))
		{
			lnMoveAmount.LowPart = lnMoveAmount.HighPart = 0;

			hResult = stm.pstm->Seek(lnMoveAmount, STREAM_SEEK_SET, &lnCurPos);
			if (FAILED(hResult))
				return hResult;

			hResult = stm.pstm->Read(&m_nSrcId, sizeof(m_nSrcId), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = stm.pstm->Read(&m_pSrcBar, sizeof(m_pSrcBar), NULL);
			if (FAILED(hResult))
				return hResult;

			switch (m_nIndex)
			{
			case CToolDataObject::eBand:
				{
					CBand* pBand = CBand::CreateInstance(NULL);
					assert(pBand);
					if (pBand)
					{
						pBand->SetOwner((CBar*)m_pDestBar);
						hResult = pBand->DragDropExchange(stm.pstm, VARIANT_FALSE);
						if (SUCCEEDED(hResult))
						{
							((CBar*)m_pDestBar)->m_pBands->InsertBand(-1, pBand);
							*pdwEffect = DROPEFFECT_COPY;
							hResult = pDestBar->RecalcLayout();
						}
						pBand->Release();
					}
				}
				break;
			
			case CToolDataObject::eCategory:
				{
					CTools* pTools;
					hResult = pDestBar->get_Tools((Tools **)&pTools);
					if (SUCCEEDED(hResult))
					{
						BSTR bstrCategory = NULL;
						hResult = StReadBSTR(stm.pstm, bstrCategory);
						if (FAILED(hResult))
						{
							pTools->Release();
							stm.pstm->Release();
							return hResult;
						}

						SysFreeString(bstrCategory);

						hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
						if (FAILED(hResult))
						{
							pTools->Release();
							stm.pstm->Release();
							return hResult;
						}

						for (int nTool = 0; nTool < nCount; nTool++)
						{
							hResult = pTools->CreateTool((ITool**)&pTool);
							if (FAILED(hResult))
								continue;

							hResult = pTool->DragDropExchange(stm.pstm, VARIANT_FALSE);
							if (FAILED(hResult))
							{
								pTool->Release();
								continue;
							}

							hResult = pTools->Insert(-1, (Tool*)pTool);
							if (SUCCEEDED(hResult))
								*pdwEffect = DROPEFFECT_COPY;
							pTool->Release();
						}
						pTools->Release();
					}
				}
				break;
			}
			if (stm.pstm)
				stm.pstm->Release();
		}
	}
	if (pBand->m_pBar->m_pDesigner && DROPEFFECT_NONE != *pdwEffect && ddBTPopup == pBand->bpV1.m_btBands)
	{
		if (pBand->m_pBar->m_pPopupRoot)
			pBand->m_pBar->m_pPopupRoot->SetPopupIndex(-1);
	}
	return NOERROR;
}

//
// BandDragDrop
//

BandDragDrop::BandDragDrop()
	: m_pWnd(NULL)
{
}

BandDragDrop::~BandDragDrop()
{
}
