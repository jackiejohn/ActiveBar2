//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Support.h"
#include "IpServer.h"
#include "Resource.h"
#include "Bar.h"
#include "Tool.h"
#include "Dock.h"
#include "CBList.h"
#include "Globals.h"
//#include <Basetsd.h>
#include <RichEdit.h>
#include "TpPopup.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// ActiveEdit
//

ActiveEdit::~ActiveEdit()
{
	if (NULL != m_pTool)
		m_pTool->SetEdit(NULL);
}

//
// Create
//

BOOL ActiveEdit::Create (HWND hWndParent, const CRect& rc, BOOL bXPLook)
{
	FWnd::CreateEx(bXPLook ? 0 : WS_EX_CLIENTEDGE,
				   _T(""),
				   m_dwStyle,
				   rc.left,
				   rc.top,
				   rc.Width(),
				   rc.Height(),
				   hWndParent,
				   (HMENU)CTool::eEdit,
				   0);
	if (NULL == m_hWnd)
		return FALSE;
	SubClassAttach(m_hWnd);

	return TRUE;
}

//
// Exchange
//

void ActiveEdit::Exchange(BOOL bFireEvent)
{
	static BOOL bInExchange = FALSE;
	if (bInExchange)
		return;
	
	bInExchange = TRUE;
	HRESULT hResult;
	int nLen = SendMessage(WM_GETTEXTLENGTH);
	if (nLen > 0)
	{
		TCHAR* szBuffer = new TCHAR[++nLen];
		if (szBuffer)
		{
			LRESULT lResult = SendMessage(WM_GETTEXT, nLen, (LPARAM)szBuffer);
			if (0 != lResult)
			{
				BSTR bstrText;
				hResult = m_pTool->get_Text(&bstrText);
				MAKE_TCHARPTR_FROMWIDE(szCmpText, bstrText);
				SysFreeString(bstrText);
				if (0 == lstrcmp(szCmpText, szBuffer))
				{
					delete [] szBuffer;
					bInExchange = FALSE;
					return;
				}

				MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
				hResult = m_pTool->put_Text(wBuffer);
				if (SUCCEEDED(hResult) && bFireEvent)
				{
					m_pTool->AddRef();                     
					::PostMessage(m_pTool->m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARTEXTCHANGE, 0, (LPARAM)m_pTool);
				}
			}
			delete [] szBuffer;
		}
	}
	else
	{
		hResult = m_pTool->put_Text(L"");
		if (SUCCEEDED(hResult) && bFireEvent)
		{
			BSTR bstrText;
			hResult = m_pTool->get_Text(&bstrText);
			int nLength = wcslen(bstrText);
			SysFreeString(bstrText);
			if (bFireEvent && nLength != nLen)
			{
				m_pTool->AddRef();
				::PostMessage(m_pTool->m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARTEXTCHANGE, 0, (LPARAM)m_pTool);
			}
		}
	}
	bInExchange = FALSE;
}

//
// WindowProc
//

LRESULT ActiveEdit::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_COMMAND:
		{
			switch (HIWORD(wParam))
			{
			case EN_CHANGE:
				Exchange();
				break;
			}
		}
		break;

	case WM_KEYDOWN:
		try
		{
			m_pTool->m_pBar->FireToolKeyDown((Tool*)m_pTool, (short*)&wParam, OCXShiftState());
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		break;

	case WM_CHAR:
		try
		{
			m_pTool->m_pBar->FireToolKeyPress((Tool*)m_pTool, (long*)&wParam);
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		break;

	case WM_KEYUP:
		try
		{
			m_pTool->m_pBar->FireToolKeyUp((Tool*)m_pTool, (short*)&wParam, OCXShiftState());
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		break;

	case WM_MOUSEACTIVATE:
		return MA_NOACTIVATE;

	case WM_KILLFOCUS:
		ShowWindow(SW_HIDE);
		break;

	case WM_DESTROY:
		{
			UnsubClass();
			delete this;
		}
		return 0;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// ActiveCombobox
//

ActiveCombobox::~ActiveCombobox()
{
	m_pTool->SetCombo(NULL);
}

//
// Create
//

BOOL ActiveCombobox::Create (HWND hWndParent, const CRect& rc, DWORD dwStyle)
{
	int nComboHeight = 120;
	if (!FWnd::Create(_T(""), 
					  dwStyle,
					  rc.left, rc.top, rc.Width(), nComboHeight,
					  hWndParent,
					  (HMENU)CTool::eCombobox,
					  0))
	{
		return FALSE;
	}

	SubClassAttach(m_hWnd);
	m_dwStyle = dwStyle;
	return NULL == m_hWnd ? FALSE : TRUE;
}

//
// WindowProc
//

LRESULT ActiveCombobox::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_MOUSEACTIVATE:
		return MA_NOACTIVATE;

	case WM_DESTROY:
		{
			long nStart, nLength;
			SendMessage(CB_GETEDITSEL, (LPARAM)&nStart, (LPARAM)&nLength);
			m_pTool->tpV1.m_nSelStart = (short)nStart; 
			m_pTool->tpV1.m_nSelLength = (short)nLength;
			m_pTool->m_bPressed = FALSE;
			m_pTool->m_bDropDownPressed = FALSE;
			UnsubClass();
			LRESULT lResult = FWnd::WindowProc(nMsg, wParam, lParam);
			m_pTool->m_pBand->Refresh();
			m_pTool->m_pBar->m_pTabbedTool = NULL;
			RemoveFromMap();
			delete this;
			return lResult;
		}
		break;

	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(m_hWnd, &ps);
			if (hDC)
			{
				CRect rcClient;
				GetClientRect(rcClient);
				rcClient.Inflate(0, -1);
				ComboStyles nStyle;
				m_pTool->get_CBStyle(&nStyle);
				if (VARIANT_FALSE == m_pTool->m_pBar->bpV1.m_vbXPLook)
				{
					m_pTool->DrawCombo(hDC, 
									   rcClient, 
									   m_pTool->m_bDropDownPressed, 
									   0 != (nStyle == ddCBSReadOnly || nStyle == ddCBSSortedReadOnly));
				}
				else
				{
					m_pTool->DrawXPCombo(hDC, 
										 rcClient, 
										 m_pTool->m_bDropDownPressed, 
										 0 != (nStyle == ddCBSReadOnly || nStyle == ddCBSSortedReadOnly));
				}
				EndPaint(m_hWnd, &ps);
			}
			return 0;
		}
		break;
	
	case WM_COMMAND:
		{
			switch (HIWORD(wParam))
			{
			case CBN_CLOSEUP:
				{
					// Save the text
					LRESULT lResult = SendMessage(CB_GETCURSEL);
					CTool* pToolTemp = m_pTool;
					pToolTemp->GetComboList()->lpV1.m_nListIndex = (short)lResult;
					if (CB_ERR != lResult)
					{
						int nLen = SendMessage(CB_GETLBTEXTLEN, lResult, 0);
						if (0 != nLen)
						{
							TCHAR* szBuffer = new TCHAR[nLen+1];
							if (szBuffer)
							{
								try
								{
									lResult = SendMessage(CB_GETLBTEXT, lResult, (LPARAM)szBuffer);
									if (CB_ERR != lResult)
									{
										MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
										HRESULT hResult = pToolTemp->put_Text(wBuffer);
										assert(SUCCEEDED(hResult));
										if (SUCCEEDED(hResult))
										{
											pToolTemp->AddRef();
											::PostMessage(m_pTool->m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARTEXTCHANGE, 0, (LPARAM)m_pTool);
										}
									}
								}
								CATCH
								{
									assert(FALSE);
									REPORTEXCEPTION(__FILE__, __LINE__)
								}
								delete [] szBuffer;
							}
						}
					}

					try
					{
						// Fire the combo close event
						pToolTemp->m_pBar->FireToolComboClose((Tool*)pToolTemp);
					}
					catch (...)
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}

					// Reset the tool
					pToolTemp->m_bPressed = FALSE;
					pToolTemp->m_bDropDownPressed = FALSE;
					if (IsWindow())
						ShowWindow(SW_HIDE);
				}
				break;

			case CBN_KILLFOCUS:
				SendMessage(WM_CANCELMODE);
				ShowWindow(SW_HIDE);
				break;

			case CBN_DROPDOWN:
				{
					m_pTool->m_bPressed = TRUE;
					m_pTool->m_bDropDownPressed = TRUE;
					CBList* pList = m_pTool->GetComboList();
					if (NULL == pList)
						break;

					int nPrevCount = pList->Count();

					m_pTool->m_pBar->FireComboDrop((Tool*)m_pTool);

					int nPrevLines = pList->lpV1.m_nLines;
					
					if (nPrevCount != pList->Count())
					{
						SendMessage(CB_RESETCONTENT, 0, 0);
						pList->m_hWndActive = m_hWnd;
						BSTR bstrItem;
						int nCount = pList->Count();
						for (int nItem = 0; nItem < nCount; nItem++)
						{
							bstrItem = pList->GetName(nItem);
							if (bstrItem)
							{
								MAKE_TCHARPTR_FROMWIDE(szItem, bstrItem);
								SendMessage(CB_ADDSTRING, 0, (LPARAM)szItem);
							}
						}
						switch (pList->lpV1.m_nStyle)
						{
						case ddCBSReadOnly:
						case ddCBSSortedReadOnly:
							{
								MAKE_TCHARPTR_FROMWIDE(szText, m_pTool->m_bstrText);
								int nIndex = SendMessage(CB_FINDSTRING, (WPARAM)-1, (LPARAM)szText);
								if (CB_ERR != nIndex)
									SendMessage(CB_SETCURSEL, nIndex, 0);
							}
							break;

						default:
							SendMessage(CB_SETCURSEL, pList->lpV1.m_nListIndex);
							break;
						}
					}
					if (nPrevLines == pList->lpV1.m_nLines)
						break;

					CRect rcCB;
					GetWindowRect(rcCB);
					int nItemHeight = SendMessage(CB_GETITEMHEIGHT, 0, 0);
					int nComboHeight = pList->lpV1.m_nLines * nItemHeight + rcCB.Height() + 2;
					::MoveWindow(m_hWnd,
								 0,
								 0,
								 rcCB.Width(),
								 nComboHeight,
								 FALSE);
				}
				break;

			case CBN_SELCHANGE:
				{
					try
					{
						int nIndex = SendMessage(CB_GETCURSEL);
						m_pTool->GetComboList()->lpV1.m_nListIndex = nIndex;
						assert (CB_ERR != nIndex);
						if (CB_ERR != nIndex)
						{
							int nLen = SendMessage(CB_GETLBTEXTLEN, nIndex, 0);
							assert(0 != nLen);
							if (0 != nLen)
							{
								TCHAR* szBuffer = new TCHAR[nLen + 1];
								assert(szBuffer);
								if (szBuffer)
								{
									HRESULT hResult;
									LRESULT lResult = SendMessage(CB_GETLBTEXT, nIndex, (LPARAM)szBuffer);
									assert(CB_ERR != lResult);
									if (CB_ERR != lResult)
									{
										MAKE_WIDEPTR_FROMTCHAR(wBuffer,szBuffer);
										hResult = m_pTool->put_Text(wBuffer);
										assert(SUCCEEDED(hResult));
									}
									UpdateWindow();
									delete szBuffer;
								}
							}
						}
						m_pTool->AddRef();
						::PostMessage(m_pTool->m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCOMBOSELCHANGE, 0, (LPARAM)m_pTool);
					}
					catch (...)
					{
						assert(FALSE);
					}
				}
				break;

			case CBN_EDITCHANGE:
				{
					static bool m_bInExchange = FALSE;

					if (m_bInExchange)
						break;

					m_bInExchange = TRUE;

					// Transfer text
					HRESULT hResult;
					int nLen = SendMessage(WM_GETTEXTLENGTH);
					if (nLen > 0)
					{
						TCHAR* szBuffer = new TCHAR[nLen+1];
						if (NULL == szBuffer)
						{
							m_bInExchange = FALSE;
							break;
						}
						if (SendMessage(WM_GETTEXT, nLen+1, (LPARAM)szBuffer) > 0)
						{
							MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
							hResult = m_pTool->put_Text(wBuffer);
							if (SUCCEEDED(hResult))
							{
								m_pTool->AddRef();                     
								::PostMessage(m_pTool->m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARTEXTCHANGE, 0, (LPARAM)m_pTool);
							}
						}
						else
						{
							hResult = m_pTool->put_Text(L"");
							if (SUCCEEDED(hResult))
							{
								m_pTool->AddRef();                     
								::PostMessage(m_pTool->m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARTEXTCHANGE, 0, (LPARAM)m_pTool);
							}
						}
						delete [] szBuffer;
					}
					else
					{
						hResult = m_pTool->put_Text(L"");
						if (SUCCEEDED(hResult))
						{
							m_pTool->AddRef();                     
							::PostMessage(m_pTool->m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARTEXTCHANGE, 0, (LPARAM)m_pTool);
						}
					}

					m_bInExchange = FALSE;
				}
				break;
			}
		}
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}
