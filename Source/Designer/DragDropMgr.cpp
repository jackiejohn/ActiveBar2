//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "PreComp.h"
#include "Utility.h"
#ifdef DESIGNER
#include "Resource.h"
#include "Debug.h"
#else
#include "..\Resource.h"
#endif
#include "..\EventLog.h"
#include "DragDropMgr.h"

extern HINSTANCE g_hInstance;

//
// CDragDropMgr
//

static DWORD KeyState()
{
	DWORD dwKeyState = 0;
	if ((GetKeyState(VK_SHIFT) < 0))
		dwKeyState |= MK_SHIFT;
	if ((GetKeyState(VK_CONTROL) < 0))
		dwKeyState |= MK_CONTROL;
	if ((GetKeyState(VK_MENU) < 0))
		dwKeyState |= MK_ALT;
	return dwKeyState;
}

//
// RegisterDragDrop
//

HRESULT CDragDropMgr::RegisterDragDrop(OLE_HANDLE hWnd, LPUNKNOWN pDropTarget)
{
	try
	{
		if (!IsWindow((HWND)hWnd))
			return DRAGDROP_E_INVALIDHWND;

		LPUNKNOWN pTempDropTarget;
		if (m_mapTargets.Lookup(hWnd, pTempDropTarget))
			return DRAGDROP_E_ALREADYREGISTERED;

		m_mapTargets.SetAt(hWnd, pDropTarget);
		return S_OK;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// RevokeDragDrop
//

HRESULT CDragDropMgr::RevokeDragDrop(OLE_HANDLE hWnd)
{
	try
	{
		if (!IsWindow((HWND)hWnd))
			return DRAGDROP_E_INVALIDHWND;

		if (!m_mapTargets.RemoveKey(hWnd))
			return DRAGDROP_E_NOTREGISTERED;
		return S_OK;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// DoDragDrop
//

HRESULT CDragDropMgr::DoDragDrop(LPUNKNOWN pDataObject, 
								 LPUNKNOWN pDropSource, 
								 DWORD	   dwOKEffect, 
								 DWORD*	   pdwEffect)
{
	enum STATES
	{
		eNotActive,
		eDragOver,
		eDragEnter,
		eDragLeave,
		eDragDrop,
	};

	IDropTarget* pCurrentDropTarget = NULL;
	HCURSOR		 hCursorPrev = GetCursor();
	HRESULT		 hResult = DROPEFFECT_NONE;
	STATES		 eState;
	POINTL		 ptl;
	BOOL		 bProcessing = TRUE;
	HWND		 hWndCurrentDropTarget = NULL;
	HWND	     hWnd = GetForegroundWindow();
	HWND	     hWndTemp;
	MSG			 msg;

	SetCapture(hWnd);
	while (bProcessing)
	{
		hWndTemp = GetCapture();
		if (hWnd != hWndTemp)
			break;

		GetMessage(&msg, NULL, 0, 0);

		switch (msg.message)
		{
		case WM_CANCELMODE:
		case WM_RBUTTONUP:
			bProcessing = FALSE;
			break;

		case WM_KEYUP:
		case WM_KEYDOWN:
			switch (msg.wParam)
			{
			case VK_CONTROL:
				if (eDragOver == eState)
				{
					try
					{

						POINTL ptl;
						ptl.x = msg.pt.x;
						ptl.y = msg.pt.y;
						if (SUCCEEDED(pCurrentDropTarget->DragOver(KeyState(), ptl, pdwEffect)))
						{
							if (*pdwEffect & DROPEFFECT_COPY && !(dwOKEffect & DROPEFFECT_MOVE) && dwOKEffect & DROPEFFECT_COPY)
								((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_COPY);
							else if (*pdwEffect & DROPEFFECT_COPY && dwOKEffect & DROPEFFECT_MOVE && dwOKEffect & DROPEFFECT_COPY)
								((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_COPY);
							else if (*pdwEffect & DROPEFFECT_MOVE && (dwOKEffect & DROPEFFECT_MOVE))
								((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_MOVE);
							else
								((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_NONE);
						}
						else
						{
							HCURSOR hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOLNOT));
							assert(hCursor);
							SetCursor(hCursor);
						}
					}
					CATCH
					{
						bProcessing = FALSE;
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
				}
				break;

			case VK_ESCAPE:
				eState = eDragLeave;
				if (pCurrentDropTarget)
				{
					try
					{
						pCurrentDropTarget->DragLeave();
						pCurrentDropTarget->Release();
						pCurrentDropTarget = NULL;
						hWndCurrentDropTarget = NULL;
					}
					CATCH
					{
						pCurrentDropTarget = NULL;
						hWndCurrentDropTarget = NULL;
						bProcessing = FALSE;
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
				}
				bProcessing = FALSE;
				hResult = DRAGDROP_S_CANCEL;
				break;
			}
			break;

		case WM_LBUTTONUP:
			{
				HWND hWndDropTarget = WindowFromPoint(msg.pt);
				if (hWndDropTarget != hWndCurrentDropTarget)
				{
					if (hWndCurrentDropTarget)
					{
						try
						{
							eState = eDragLeave;
							if (pCurrentDropTarget)
							{
								pCurrentDropTarget->DragLeave();
								pCurrentDropTarget->Release();
								pCurrentDropTarget = NULL;
							}
							hWndCurrentDropTarget = NULL;
						}
						CATCH
						{
							pCurrentDropTarget = NULL;
							hWndCurrentDropTarget = NULL;
							bProcessing = FALSE;
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
					}
					else
					{
						if (m_mapTargets.Lookup((OLE_HANDLE)hWndDropTarget, (LPUNKNOWN&)pCurrentDropTarget))
							hResult = E_UNEXPECTED;
						else
							hResult = DRAGDROP_S_DROP;
					}
					*pdwEffect = DROPEFFECT_NONE;
				}
				else
				{
					try
					{
						POINTL ptl;
						ptl.x = msg.pt.x;
						ptl.y = msg.pt.y;
						eState = eDragDrop;											  
						hResult = pCurrentDropTarget->Drop((IDataObject*)pDataObject, KeyState(), ptl, pdwEffect);
						if (SUCCEEDED(hResult))
							hResult = DRAGDROP_S_DROP;
						else
							hResult = DRAGDROP_S_CANCEL;
			
						pCurrentDropTarget->Release();
						pCurrentDropTarget = NULL;
					}
					CATCH
					{
						pCurrentDropTarget = NULL;
						hWndCurrentDropTarget = NULL;
						bProcessing = FALSE;
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
				}
				bProcessing = FALSE;
			}
			break;

		case WM_MOUSEMOVE:
			{
				HWND hWndDropTarget = WindowFromPoint(msg.pt);
				if (hWndDropTarget != hWndCurrentDropTarget)
				{
					if (hWndCurrentDropTarget)
					{
						try
						{
							eState = eDragLeave;
							if (pCurrentDropTarget)
							{
								pCurrentDropTarget->DragLeave();
								pCurrentDropTarget->Release();
								pCurrentDropTarget = NULL;
							}
							hWndCurrentDropTarget = NULL;
						}
						CATCH
						{
							pCurrentDropTarget = NULL;
							hWndCurrentDropTarget = NULL;
							bProcessing = FALSE;
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
					}
					if (m_mapTargets.Lookup((OLE_HANDLE)hWndDropTarget, (LPUNKNOWN&)pCurrentDropTarget))
					{
						try
						{
							eState = eDragEnter;
							pCurrentDropTarget->AddRef();
							hWndCurrentDropTarget = hWndDropTarget;
							ptl.x = msg.pt.x;
							ptl.y = msg.pt.y;
							if (SUCCEEDED(pCurrentDropTarget->DragEnter((IDataObject*)pDataObject, KeyState(), ptl, pdwEffect)))
							{
								if (*pdwEffect & DROPEFFECT_COPY && !(dwOKEffect & DROPEFFECT_MOVE) && dwOKEffect & DROPEFFECT_COPY)
									((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_COPY);
								else if (*pdwEffect & DROPEFFECT_COPY && dwOKEffect & DROPEFFECT_MOVE && dwOKEffect & DROPEFFECT_COPY)
									((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_COPY);
								else if (*pdwEffect & DROPEFFECT_MOVE && (dwOKEffect & DROPEFFECT_MOVE))
									((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_MOVE);
								else
									((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_NONE);
							}
						}
						CATCH
						{
							bProcessing = FALSE;
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
					}
					else
					{
						HCURSOR hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOLNOT));
						assert(hCursor);
						SetCursor(hCursor);
					}
				}
				else
				{
					try
					{
						POINTL ptl;
						ptl.x = msg.pt.x;
						ptl.y = msg.pt.y;
						eState = eDragOver;
						if (SUCCEEDED(pCurrentDropTarget->DragOver(KeyState(), ptl, pdwEffect)))
						{
							if (*pdwEffect & DROPEFFECT_COPY && !(dwOKEffect & DROPEFFECT_MOVE) && dwOKEffect & DROPEFFECT_COPY)
								((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_COPY);
							else if (*pdwEffect & DROPEFFECT_COPY && dwOKEffect & DROPEFFECT_MOVE && dwOKEffect & DROPEFFECT_COPY)
								((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_COPY);
							else if (*pdwEffect & DROPEFFECT_MOVE && (dwOKEffect & DROPEFFECT_MOVE))
								((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_MOVE);
							else
								((IDropSource*)pDropSource)->GiveFeedback(DROPEFFECT_NONE);
						}
						else
						{
							HCURSOR hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOLNOT));
							assert(hCursor);
							SetCursor(hCursor);
						}
					}
					CATCH
					{
						bProcessing = FALSE;
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
				}
			}
			break;

		default:
			DispatchMessage(&msg);
			break;
		}
	}
	eState = eNotActive;
	ReleaseCapture();

	if (hCursorPrev)
		SetCursor(hCursorPrev);

	return hResult;
}

