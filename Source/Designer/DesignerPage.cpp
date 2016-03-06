//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "ipserver.h"
#include "Resource.h"
#include "Dialogs.h"
#include "Browser.h"
#include "Stream.h"
#include "support.h"
#include "debug.h"
#include "proppage.h"
#include "Utility.h"
#include "DesignerInterfaces.h"
#include "..\hlp\ActiveBarUG.hh"
#include "..\EventLog.h"
#include "..\Interfaces.h"
#include "..\StaticLink.h"
#include "DesignerPage.h"

IUnknown *CreateFN_CDesignerPage(IUnknown *)
{
	return (IUnknown *)new CDesignerPage();
}

//{DEF
DEFINE_PROPPAGE(&CLSID_DesignerPage,DesignerPage,_T("DesignerPage"),CreateFN_CDesignerPage,IDD_DESIGNER,IDS_DESIGNER,IDS_DESIGNERDOC,_T("Activebar.hlp"),0);

CDesignerPage::CDesignerPage()
	: CPropertyPage(),
	  m_pActiveBar(NULL),
	  m_pSplitter(NULL),
	  m_pDesignerTree(NULL),
	  m_pCategoryMgr(NULL)
{
	m_pDesc=&DesignerPageObject;
	
	// Initialize local variables
	m_strMsgTitle.LoadString(IDS_DESIGNERTITLE);
	
	VariantInit(&m_vLayoutData);

	m_pSplitter = new CSplitter();
	assert(m_pSplitter);

	m_pBrowser = new CBrowser(this);
	assert(m_pBrowser);
	
	m_pDesignerTree = new CDesignerTree;
	assert(m_pDesignerTree);
	
	m_pCategoryMgr = new CategoryMgr;
	assert(m_pCategoryMgr);
}

CDesignerPage::~CDesignerPage()
{
	Cleanup();
}

//
// Cleanup
//

void CDesignerPage::Cleanup()
{
	try
	{
		delete m_ptbDesigner;
		delete m_pBrowser;
		delete m_pCategoryMgr;
		delete m_pDesignerTree;
		delete m_pSplitter;
		VariantClear(&m_vLayoutData);
		if (m_pActiveBar)
			m_pActiveBar->Release();
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
}

//
// IsButtonDisabled
//

BOOL CDesignerPage::IsButtonDisabled(UINT nId)
{
	DWORD dwStyle = GetWindowLong(GetDlgItem(GetParent(GetParent(m_hwnd)), nId), GWL_STYLE);
	return (WS_DISABLED & dwStyle) != 0;

}

//
// SaveDisableButtons
//

void CDesignerPage::SaveDisableButtons()
{ 
	m_bOk = IsButtonDisabled(eOk);
	if (!m_bOk)
		EnableWindow(GetDlgItem(GetParent(GetParent(m_hwnd)), eOk), FALSE);

	m_bCancel = IsButtonDisabled(eCancel);
	if (!m_bCancel)
		EnableWindow(GetDlgItem(GetParent(GetParent(m_hwnd)), eCancel), FALSE);

	m_bApply = IsButtonDisabled(eApply);
	if (!m_bApply)
		EnableWindow(GetDlgItem(GetParent(GetParent(m_hwnd)), eApply), FALSE);

	m_bHelp = IsButtonDisabled(eHelp);
	if (!m_bHelp)
		EnableWindow(GetDlgItem(GetParent(GetParent(m_hwnd)), eHelp), FALSE);
}

//
// RestoreButtons
//

void CDesignerPage::RestoreButtons()
{
	if (!m_bOk)
		EnableWindow(GetDlgItem(GetParent(GetParent(m_hwnd)), eOk), TRUE);

	if (!m_bCancel)
		EnableWindow(GetDlgItem(GetParent(GetParent(m_hwnd)), eCancel), TRUE);

	if (!m_bApply)
		EnableWindow(GetDlgItem(GetParent(GetParent(m_hwnd)), eApply), TRUE);

	if (!m_bHelp)
		EnableWindow(GetDlgItem(GetParent(GetParent(m_hwnd)), eHelp), TRUE);
}

//
// SetDirty
//

STDMETHODIMP CDesignerPage::SetDirty()
{
	try
	{
		MakeDirty();
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
	return NOERROR;
}

//
// CreateTool
//

STDMETHODIMP CDesignerPage::CreateTool(IDispatch* pBandIn, long nToolType)
{
	try
	{
		IBand* pBand = (IBand*)pBandIn;
		if (NULL == pBand)
			return NOERROR;

		int nToolId = FindLastToolId(m_pActiveBar) + 1;
		
		DDString strTool;
		strTool.Format(IDS_TOOLNAME, nToolId);
		
		BSTR bstrTool = strTool.AllocSysString();
		if (NULL == bstrTool)
			return E_FAIL;

		ITools* pTools;
		HRESULT hResult = pBand->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
			return hResult;

		ITool* pTool;
		hResult = pTools->Add(nToolId, bstrTool, (Tool**)&pTool);
		pTools->Release();
		SysFreeString(bstrTool);
		if (FAILED(hResult))
			return hResult;

		if (ddTTSeparator != nToolType)
		{
			hResult = pTool->put_Caption(bstrTool);
			if (FAILED(hResult))
				return hResult;
		}

		hResult = pTool->put_ControlType((ToolTypes)nToolType);
		if (FAILED(hResult))
			return hResult;

		hResult = m_pActiveBar->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
			return hResult;

		hResult = pTools->Insert(-1, (Tool*)pTool);
		pTools->Release();
		if (FAILED(hResult))
		{
			pTool->Release();
			return hResult;
		}
		pTool->Release();
		m_pActiveBar->RecalcLayout();
		SetDirty();
		if (m_pDesignerTree)
		{
			BSTR bstrBandName;
			hResult = pBand->get_Name(&bstrBandName);
			if (FAILED(hResult))
				return hResult;
			
			IChildBands* pChildBands;
			hResult = pBand->get_ChildBands((ChildBands**)&pChildBands);
			if (FAILED(hResult))
			{
				SysFreeString(bstrBandName);
				return hResult;
			}

			IBand* pChildBand;
			hResult = pChildBands->get_CurrentChildBand((Band**)&pChildBand);
			if (FAILED(hResult))
			{
				SysFreeString(bstrBandName);
				return hResult;
			}
			
			BSTR bstrPageName;
			hResult = pChildBand->get_Name(&bstrPageName);
			pChildBand->Release();
			if (FAILED(hResult))
			{
				SysFreeString(bstrBandName);
				return hResult;
			}

			m_pDesignerTree->ToolSelected(bstrBandName, bstrPageName, nToolId);
			SysFreeString(bstrBandName);
			SysFreeString(bstrPageName);
		}
		return NOERROR;
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
		return E_FAIL;
	}	
}

STDMETHODIMP CDesignerPage::Close()
{
	PostMessage(GetParent(GetParent(m_hwnd)), WM_CLOSE, 0, 0);
	return NOERROR;
}

STDMETHODIMP CDesignerPage::TranslateAccelerator(LPMSG pmsg)
{
	try
	{
		switch (pmsg->message)
		{
		case WM_KEYUP:
		case WM_KEYDOWN:
			switch (pmsg->wParam)
			{
			case VK_ESCAPE:
				try
				{
					if (m_pDesignerTree)
					{
						HWND hWndEdit = TreeView_GetEditControl(m_pDesignerTree->hWnd());
						if (hWndEdit && pmsg->hwnd == hWndEdit)
						{
							::SendMessage(hWndEdit, pmsg->message, pmsg->wParam, pmsg->lParam);
							return S_OK;
						}
					}
				}
				catch (SEException& e)
				{
					assert(FALSE);
					e.ReportException(__FILE__, __LINE__);
				}
				break;

			case VK_UP:
			case VK_DOWN:
			case VK_LEFT:
			case VK_RIGHT:
				try
				{
					if (m_pBrowser && pmsg->hwnd == m_pBrowser->Edit().hWnd())
					{
						m_pBrowser->Edit().SendMessage(pmsg->message, pmsg->wParam, pmsg->lParam);
						return S_OK;
					}
				}
				catch (SEException& e)
				{
					assert(FALSE);
					e.ReportException(__FILE__, __LINE__);
				}
				break;

			case VK_RETURN:
				try
				{
					if (m_pBrowser && pmsg->hwnd == m_pBrowser->Edit().hWnd())
					{
						m_pBrowser->Edit().SendMessage(pmsg->message, pmsg->wParam, pmsg->lParam);
						return S_OK;
					}
					else if (m_pDesignerTree)
					{
						HWND hWndEdit = TreeView_GetEditControl(m_pDesignerTree->hWnd());
						if (hWndEdit && pmsg->hwnd == hWndEdit)
						{
							::SendMessage(hWndEdit, pmsg->message, pmsg->wParam, pmsg->lParam);
							return S_OK;
						}
					}
				}
				catch (SEException& e)
				{
					assert(FALSE);
					e.ReportException(__FILE__, __LINE__);
				}
				break;

			case VK_F4:
				try
				{
					if (m_pBrowser && pmsg->hwnd == m_pBrowser->hWnd())
					{
						m_pBrowser->SendMessage(pmsg->message, pmsg->wParam, pmsg->lParam);
						return S_OK;
					}
				}
				catch (SEException& e)
				{
					assert(FALSE);
					e.ReportException(__FILE__, __LINE__);
				}
				break;
			}
		}
		if (IsDialogMessage(m_hwnd, pmsg))
			return S_OK;
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
	return S_FALSE;
}

BOOL CDesignerPage::DialogProc(HWND hWnd,UINT msg,WPARAM wParam,LPARAM lParam)
{
	try
	{
		switch(msg)
		{
		case WM_INITDIALOG:
			{
				try
				{
					//
					// Setup the splitter
					//

					HWND hWndChild;
					if (m_pSplitter)
					{
						m_pSplitter->hWnd(m_hwnd);

						CRect rcTree;
						hWndChild = GetDlgItem(m_hwnd, IDC_TREE_DESIGNER);
						if (hWndChild)
							::GetWindowRect(hWndChild, &rcTree);				
						ScreenToClient(m_hwnd, rcTree);
						
						CRect rcProperty;
						hWndChild = GetDlgItem(m_hwnd, IDC_LST_PROPERTIES);
						if (hWndChild)
							::GetWindowRect(hWndChild, &rcProperty);
						ScreenToClient(m_hwnd, rcProperty);

						m_pSplitter->SetLeftLimit(rcTree.left + 4);
						m_pSplitter->SetRightLimit(rcProperty.right - 8);
						CRect rc;
						if (GetGlobals().m_thePreferences.LoadWinRect(PREFID_DESIGNERTRACKER2, rc))
						{
							//
							// I have to size the windows based on the tracker's position
							//
							m_pSplitter->Tracker(rc);
							POINT pt = {rc.top, rc.left};
							OnLButtonUp(0, pt);
							InvalidateRect(m_hwnd, NULL, FALSE);
						}
						else
							m_pSplitter->Tracker(rcTree.right, rcProperty.left, rcTree.top, rcTree.bottom);
					}

					//
					// Create the designer object
					//

					m_pDesignerTree->SubClassAttach(GetDlgItem(m_hwnd, IDC_TREE_DESIGNER));
					m_pBrowser->SubClassAttach(GetDlgItem(m_hwnd, IDC_LST_PROPERTIES));

					//
					// We stuck a static control just to kill it and replace it with our
					// Designer ToolBar
					//
					
					CRect rc;
					hWndChild = GetDlgItem(m_hwnd, IDC_TOOLBAR);
					if (hWndChild)
					{
						GetWindowRect(hWndChild, &rc);
						ScreenToClient(m_hwnd, rc);
						DestroyWindow(hWndChild);
					}

					//
					// Create the Designer ToolBar
					//

					CreateDesignerToolbar(rc.left, rc.top, rc.Width(), rc.Height());
				}
				catch (SEException& e)
				{
					assert(FALSE);
					e.ReportException(__FILE__, __LINE__);
				}
			}
			break;

		case WM_SYSCOLORCHANGE:
		case WM_SETTINGCHANGE:
			m_pDesignerTree->SettingsChanged();
			break;

		case WM_HELP:
			{
				LPHELPINFO pHI = (LPHELPINFO)lParam; 
				try
				{
					UINT dwData = HX_ActiveBar_Designer;
					switch (pHI->iCtrlId)
					{
					case IDC_TREE_DESIGNER:
						dwData = HX_Tools_and_Bands_Treeview;
						break;

					case IDC_LST_PROPERTIES:
						dwData = HX_Properties_List;
						break;

					case ID_TOOLBAR:
						dwData = HX_Designer_Toolbar;
						break;
					}
					WinHelp(hWnd, HELP_CONTEXT, dwData);
					return TRUE;
				}
				catch (SEException& e)
				{
					assert(FALSE);
					e.ReportException(__FILE__, __LINE__);
				}
				break;
			}
			break;

		case PPM_NEWOBJECTS:
			try
			{
				if (m_cObjects > 0)
				{
					//
					// Check for the ActiveBar Interface
					//

					IActiveBar2* pActiveBar = NULL;
					HRESULT hResult = m_ppUnkObjects[0]->QueryInterface(IID_IActiveBar2, (void**)&pActiveBar);
					if (FAILED(hResult))
					{
						MessageBox(IDS_ERR_QUERYACTIVEBARINTERFACE);
						break;
					}

					if (pActiveBar == m_pActiveBar)
					{
						pActiveBar->Release();
						return TRUE;
					}

					if (m_pActiveBar)
						m_pActiveBar->Release();
					m_pActiveBar = pActiveBar;

					BSTR bstrVersion;
					hResult = m_pActiveBar->get_Version(&bstrVersion);
					if (FAILED(hResult))
					{
						MessageBox(IDS_ERR_VERSIONNUMBER);
						return FALSE;
					}

					MAKE_TCHARPTR_FROMWIDE(szVersionAB, bstrVersion);

					CVersionInfo viFileVersion(g_hInstance, _T("designer.dll"));
					LPTSTR szVersion = (LPTSTR)viFileVersion.GetValue(_T("FileVersion"));
					TCHAR *ptr = szVersion;
					while (*ptr)
					{
						if (*ptr==',') 
							*ptr='.';
						else if (*ptr==32)
							lstrcpy(ptr,ptr+1);
						else
							++ptr;
					}
					if (0 != _tcscmp(szVersionAB, szVersion))
					{
						MessageBox(IDS_ERR_VERSIONNUMBER);
						return FALSE;
					}

					//
					// Store the layout information in an array of byte
					//

					VariantClear(&m_vLayoutData);

					hResult = m_pActiveBar->Save(NULL, NULL, ddSOByteArray, &m_vLayoutData); 
					if (FAILED(hResult))
						MessageBox(IDS_ERR_FAILEDTOSETLAYOUT);

					//
					// Initialize the Designer with ActiveBar
					//

					IBarPrivate* pPrivate;
					hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
					if (SUCCEEDED(hResult))
					{
						pPrivate->DesignerInitialize((IDragDropManager*)GetGlobals().GetDragDropMgr(), this);
						pPrivate->Release();
					}

					//
					// Initialize the Category Manager and the Tree
					//

					m_pCategoryMgr->ActiveBar(m_pActiveBar);
					TreeView_DeleteAllItems(GetDlgItem(m_hwnd, IDC_TREE_DESIGNER));
					m_pDesignerTree->Init(m_pActiveBar, m_pBrowser, m_pCategoryMgr);
				}
			}
			catch (SEException& e)
			{
				assert(FALSE);
				e.ReportException(__FILE__, __LINE__);
				if (!e.Continue())
					throw;
			}
			return TRUE;

		case PPM_APPLY:
			if (m_pActiveBar)
			{
				IBarPrivate* pPrivate;
				HRESULT hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
				if (SUCCEEDED(hResult))
				{
					pPrivate->SetDesignerModified();
					pPrivate->Release();
				}
			}
			return TRUE;

		case WM_DESTROY:
			try
			{

				//
				// These windows are subclassed
				//

				UIUtilities::CleanupUIUtilities();
				TreeView_DeleteAllItems(GetDlgItem(m_hwnd, IDC_TREE_DESIGNER));
				::SendMessage(GetDlgItem(hWnd, IDC_LST_PROPERTIES), LB_RESETCONTENT, 0, 0);
				m_pDesignerTree->UnsubClass();
				m_pBrowser->UnsubClass();
				WinHelp(m_hwnd, NULL, HELP_QUIT, NULL);
				delete m_pCategoryMgr;
				m_pCategoryMgr = NULL;
				if (m_pSplitter)
					GetGlobals().m_thePreferences.SaveWinRect(PREFID_DESIGNERTRACKER2, m_pSplitter->Tracker());
			}
			catch (SEException& e)
			{
				assert(FALSE);
				e.ReportException(__FILE__, __LINE__);
			}
			break;

		case WM_NCDESTROY:
			if (m_pActiveBar)
			{
				HRESULT hResult;
				if (m_fDirty && VT_EMPTY != m_vLayoutData.vt)
				{
					hResult = m_pActiveBar->Load(NULL, &m_vLayoutData, ddSOByteArray);
					if (FAILED(hResult))
						MessageBox(IDS_FAILEDTORESTORELAYOUT);
					else
						m_pActiveBar->RecalcLayout();
				}
				IBarPrivate* pPrivate;
				hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
				if (SUCCEEDED(hResult))
				{
					pPrivate->DesignerShutdown();
					pPrivate->Release();
				}
				m_pActiveBar->Release();
				m_pActiveBar = NULL;
			}
			break;

		case WM_NOTIFY:
			try
			{
				switch (wParam)
				{
				case IDC_TREE_DESIGNER:
					{
						HWND hWndChild = GetDlgItem(m_hwnd, IDC_TREE_DESIGNER);
						if (IsWindow(hWndChild))
						{
							LRESULT lResult = SendMessage(hWndChild, msg, wParam, lParam);
							SetWindowLong(m_hwnd, DWL_MSGRESULT, lResult);
							return lResult;
						}
					}
					break;

				case ID_SETPAGEDIRTY:
					SetDirty();
					break;
				}
			}
			catch (SEException& e)
			{
				assert(FALSE);
				e.ReportException(__FILE__, __LINE__);
			}
			break;

		case WM_VKEYTOITEM:
			try
			{
				if (IsWindow((HWND)lParam) && m_pBrowser && m_pBrowser->hWnd() == (HWND)lParam)
					return m_pBrowser->SendMessage(WM_VKEYTOITEM, wParam, lParam);
			}
			catch (SEException& e)
			{
				assert(FALSE);
				e.ReportException(__FILE__, __LINE__);
			}
			break;

		case WM_COMPAREITEM:
		case WM_DELETEITEM:
		case WM_DRAWITEM:
			if (IDC_LST_PROPERTIES == wParam)
			{
				try
				{
					if (m_pBrowser && IsWindow(m_pBrowser->hWnd()) && m_pBrowser->hWnd())
						return m_pBrowser->SendMessage(msg, wParam, lParam);
				}
				catch (SEException& e)
				{
					assert(FALSE);
					e.ReportException(__FILE__, __LINE__);
				}
			}
			break;

		case WM_COMMAND:
			OnCommand(wParam, lParam);
			break;

		case WM_LBUTTONDOWN:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				OnLButtonDown(wParam, pt);
			}
			break;

		case WM_MOUSEMOVE:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				OnMouseMove(wParam, pt);
			}
			break;

		case WM_LBUTTONUP:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				OnLButtonUp(wParam, pt);
			}
			break;
		}
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
	return FALSE;
}

void CDesignerPage::OnLButtonDown(UINT nFlags, POINT pt)
{
	try
	{
		if (m_pSplitter)
		{
			if (m_pSplitter->SplitterHitTest(pt))
				m_pSplitter->StartTracking();
		}
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
}

void CDesignerPage::OnMouseMove(UINT nFlags, POINT pt)
{
	try
	{
		if (m_pSplitter)
			m_pSplitter->DoTracking(pt);
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
}

void CDesignerPage::OnLButtonUp(UINT nFlags, POINT pt)
{
	try
	{
		if (m_pSplitter)
		{
			m_pSplitter->StopTracking(pt);
			
			HWND hWndChild = GetDlgItem(m_hwnd, IDC_TREE_DESIGNER);
			if (IsWindow(hWndChild))
			{
				CRect rcTree;
				GetWindowRect(hWndChild, &rcTree);
				ScreenToClient(m_hwnd, rcTree);
				rcTree.right = m_pSplitter->TrackerLeft();
				MoveWindow(hWndChild,
						   rcTree.left,
						   rcTree.top,
						   rcTree.Width(),
						   rcTree.Height(),
						   TRUE);
			}

			CRect rcProperty;
			hWndChild = GetDlgItem(m_hwnd, IDC_LST_PROPERTIES);
			if (IsWindow(hWndChild))
			{
				GetWindowRect(hWndChild, &rcProperty);
				ScreenToClient(m_hwnd, rcProperty);
				rcProperty.left = m_pSplitter->TrackerRight();
				MoveWindow(hWndChild,
						   rcProperty.left,
						   rcProperty.top,
						   rcProperty.Width(),
						   rcProperty.Height(),
						   TRUE);
			}
			CRect rcButton = rcProperty;
			hWndChild = GetDlgItem(m_hwnd, IDC_ST_DESC);
			if (IsWindow(hWndChild))
			{
				rcButton.top = rcButton.bottom + 5;
				rcButton.bottom = rcButton.top + 35;
				MoveWindow(hWndChild,
						   rcButton.left,
						   rcButton.top,
						   rcButton.Width(),
						   rcButton.Height(),
						   TRUE);
				rcButton.bottom += 30;
			}
			InvalidateRect(m_hwnd, NULL, TRUE);
		}
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
}

HRESULT CDesignerPage::CreateDesignerToolbar (long nX,  long nY,  long nWidth,  long nHeight)
{
	try
	{
		CRect rc(nX, nX+nWidth, nY, nY+nHeight);
		m_ptbDesigner = new DDToolBar;
		assert(m_ptbDesigner);
		if (NULL == m_ptbDesigner)
			return E_OUTOFMEMORY;

		m_ptbDesigner->LoadToolBar(MAKEINTRESOURCE(IDR_TBDESIGNER));
		m_ptbDesigner->SetStyle(DDToolBar::eRaised|DDToolBar::eIE);
		m_ptbDesigner->Create(m_hwnd, rc, ID_TOOLBAR);
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
		if (!e.Continue())
			throw;
	}
	return S_OK;
}

HRESULT CDesignerPage::OnCommand(long wParam, long lParam)
{
	try
	{
		switch(LOWORD(wParam))
		{
		case ID_FILEOPEN:
			EnableWindow(GetParent(GetParent(m_hwnd)), FALSE);
			if (m_pDesignerTree)
				m_pDesignerTree->OnOpenLayout();
			EnableWindow(GetParent(GetParent(m_hwnd)), TRUE);
			break;

		case ID_FILENEW:
			EnableWindow(GetParent(GetParent(m_hwnd)), FALSE);
			if (m_pDesignerTree)
				m_pDesignerTree->OnNewLayout();
			EnableWindow(GetParent(GetParent(m_hwnd)), TRUE);
			break;

		case ID_FILESAVE:
			EnableWindow(GetParent(GetParent(m_hwnd)), FALSE);
			if (m_pDesignerTree)
				m_pDesignerTree->OnSaveLayout();
			EnableWindow(GetParent(GetParent(m_hwnd)), TRUE);
			break;


		case ID_FILE_SAVEAS:
			EnableWindow(GetParent(GetParent(m_hwnd)), FALSE);
			if (m_pDesignerTree)
				m_pDesignerTree->OnSaveAsLayout();
			EnableWindow(GetParent(GetParent(m_hwnd)), TRUE);
			break;

		case ID_OPTIONS:
			EnableWindow(GetParent(GetParent(m_hwnd)), FALSE);
			OnOptions();
			EnableWindow(GetParent(GetParent(m_hwnd)), TRUE);
			break;

		case ID_REPORT:
			EnableWindow(GetParent(GetParent(m_hwnd)), FALSE);
			OnReport();
			EnableWindow(GetParent(GetParent(m_hwnd)), TRUE);
			break;

		case ID_HEADER:
			EnableWindow(GetParent(GetParent(m_hwnd)), FALSE);
			OnHeader();
			EnableWindow(GetParent(GetParent(m_hwnd)), TRUE);
			break;

		case ID_LIBRARY:
			{
				SaveDisableButtons();
				m_ptbDesigner->CheckButton(ID_LIBRARY, TRUE);
				m_ptbDesigner->EnableButton(ID_LIBRARY, FALSE);
				OnLibrary();
				m_ptbDesigner->CheckButton(ID_LIBRARY, FALSE);
				m_ptbDesigner->EnableButton(ID_LIBRARY, TRUE);
				RestoreButtons();
			}
			break;

		case ID_MENUGRAB:
			EnableWindow(GetParent(GetParent(m_hwnd)), FALSE);
			OnMenuGrab();
			EnableWindow(GetParent(GetParent(m_hwnd)), TRUE);
			break;

		case ID_SETPAGEDIRTY:
			MakeDirty();
			break;

		case IDC_LST_PROPERTIES:
			SendMessage((HWND)lParam, WM_COMMAND, wParam, lParam);
			break;

		case ID_GEN_SELECT:
			OnGenSelect();
			break;

		default:
			break;
		}
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
		if (!e.Continue())
			throw;
	}
	return S_OK;
}

//
// OnOptions
//

void CDesignerPage::OnOptions()
{
	try
	{
		CDlgOptions theOptions;
		theOptions.DoModal();
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
		if (!e.Continue())
			throw;
	}
}

//
// OnGenSelect
//

void CDesignerPage::OnGenSelect()
{
	//
	// Generate the select string
	//

	GenSelect theGenSelect;
	theGenSelect.m_strSelect = _T("Select Case tool.Name\n");
	
	VisitBarTools(m_pActiveBar, -1, &theGenSelect, GenSelect::DoIt);
	
	DDString strTemp;
	strTemp.Format(_T("%sEnd Select\n"), theGenSelect.m_strSelect);
	
	//
	// Put it on the clipbroad
	//

	if (!OpenClipboard(m_hwnd))
	{
		MessageBox (IDS_ERR_COULDNOTOPENCLIPBROAD, MB_ICONSTOP);
		return;
	}

	if (!EmptyClipboard())
	{
		MessageBox (IDS_ERR_COULDNOTEMPTYCLIPBROAD, MB_ICONSTOP);
		return;
	}

	HGLOBAL hTest = GlobalAlloc(GMEM_MOVEABLE|GMEM_DDESHARE, strTemp.GetLength());
	if (NULL == hTest)
	{
		MessageBox (IDS_ERR_COULDNOTALLOCATEMEMORY, MB_ICONSTOP);
		return;
	}

	LPVOID pText = GlobalLock(hTest);
	if (NULL == pText)
	{
		GlobalFree(hTest);
		MessageBox (IDS_ERR_COULDNOTLOCKMEMORY, MB_ICONSTOP);
		return;
	}

	memcpy(pText, (LPTSTR)strTemp, strTemp.GetLength());
	
	GlobalUnlock(hTest);

	HANDLE hBM = SetClipboardData(CF_TEXT, hTest);
	if (NULL == hBM)
	{
		MessageBox (IDS_ERR_COULDNOTSETCLIPBROAD, MB_ICONSTOP);
		GlobalFree(hTest);
		CloseClipboard();
		return;
	}

	if (!CloseClipboard())
	{
		MessageBox (IDS_ERR_COULDNOTCLOSECLIPBROAD, MB_ICONSTOP);
		return;
	}
}

//
// OnReport
//

void CDesignerPage::OnReport()
{
	try
	{
		TCHAR szBuffer[256];
		_tcscpy(szBuffer, _T("Band Report (*.txt)|*.txt|"));
		CFileDialog dlgSave(IDS_REPORT,
							FALSE,
							_T("*.txt"),
							_T("*.txt"),
							OFN_HIDEREADONLY,
							szBuffer,
							m_hwnd);

		if (IDOK != dlgSave.DoModal())
			return;

		HANDLE hFile = CreateFile(dlgSave.GetFileName(), 
								  GENERIC_WRITE, 
								  0, 
								  0, 
								  CREATE_ALWAYS,
								  FILE_ATTRIBUTE_NORMAL,
								  0);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			return;
		}

		Report theReport(hFile);
		
		_stprintf(theReport.m_szLine, _T("Active Bar Project Report\r\n"));
		WriteFile(theReport.m_hFile, theReport.m_szLine, _tcslen(theReport.m_szLine), (DWORD*)&theReport.m_dwFilePosition, 0);

		if (m_pCategoryMgr)
		{
			m_pCategoryMgr->Init();
			m_pCategoryMgr->DoReport(theReport);
		}

		_stprintf(theReport.m_szLine, _T("\r\nBands\r\n"));
		WriteFile(theReport.m_hFile, theReport.m_szLine, _tcslen(theReport.m_szLine), (DWORD*)&theReport.m_dwFilePosition, 0);
		VisitBandTools(m_pActiveBar, -1, &theReport, Report::DoIt);
		CloseHandle(hFile);
		ShellExecute(0, _T("open"), dlgSave.GetFileName(), 0, 0, SW_SHOWNORMAL);
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
		if (!e.Continue())
			throw;
	}
}

//
// OnHeader
//

void CDesignerPage::OnHeader()
{
	try
	{
		TCHAR szBuffer[256];
		_tcscpy(szBuffer, _T("Tool Header File (*.h)|*.h|"));
		CFileDialog dlgSave(IDS_HEADERFILE,
							FALSE,
							_T("*.h"),
							_T("*.h"),
							OFN_HIDEREADONLY,
							szBuffer,
							m_hwnd);

		if (IDOK != dlgSave.DoModal())
			return;

		HANDLE hFile = CreateFile(dlgSave.GetFileName(), 
								  GENERIC_WRITE, 
								  0, 
								  0, 
								  CREATE_ALWAYS,
								  FILE_ATTRIBUTE_NORMAL,
								  0);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			return;
		}

		Header theHeader(hFile);
		VisitBarTools(m_pActiveBar, -1, &theHeader, Header::DoIt);
		CloseHandle(hFile);
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
		if (!e.Continue())
			throw;
	}
}

//
// OnLibrary
//

void CDesignerPage::OnLibrary()
{
	try
	{
		CDlgLibrary theLibrary;
		theLibrary.DoModal();
	}
	catch (SEException& e)
	{
		assert(FALSE);
		e.ReportException(__FILE__, __LINE__);
	}
}

//
// OnMenuGrab
//

void CDesignerPage::OnMenuGrab()
{
	HWND hWndMenu = GrabMenu::FindMenuWindow(m_hwnd);
	if (IsWindow(hWndMenu))
	{
		HMENU hMenu = GetMenu(hWndMenu);
		if (hMenu)
		{
			GrabMenu theMenu(hMenu, m_pActiveBar);
			theMenu.DoIt();
		}
		else
			MessageBeep(0xFFFFFFFF);
	}
	else if (!GrabMenu::m_bCancelled)
		MessageBeep(0xFFFFFFFF);
}