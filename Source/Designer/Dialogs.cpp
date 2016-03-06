//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Debug.h"
#include "Support.h"
#include "Globals.h"
#include "..\Dib.h"
#include "..\EventLog.h"
#include "..\StaticLink.h"
#include "..\Errors.h"
#include "..\hlp\ActiveBarUG.hh"
#include "Trees.h"
#include "Designer.h"
#include "Browser.h"
#include "IconEdit.h"
#include "Resource.h"
#include "Dialogs.h"
#include <ShlObj.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;

static void DebugPrintBitmap(HBITMAP hBitmap, int nOffset = 0)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return;

	HDC hDCSrc = CreateCompatibleDC(hDC);
	HBITMAP hBitmapOld = SelectBitmap(hDCSrc, hBitmap);
	BITMAP bmInfo;
	GetObject(hBitmap, sizeof(bmInfo), &bmInfo);
	BitBlt(hDC, 
		   0, 
		   nOffset, 
		   bmInfo.bmWidth, 
		   bmInfo.bmHeight, 
		   hDCSrc, 
		   0, 
		   0, 
		   SRCCOPY);
	SelectBitmap(hDCSrc, hBitmapOld);
	DeleteDC(hDCSrc);
	ReleaseDC(NULL, hDC);
}

#endif

extern HINSTANCE g_hInstance;

#define WM_UPDATEWINDOW WM_USER + 100

//
// CDlgOptions
//

CDlgOptions::CDlgOptions()
	: FDialog(IDD_DLG_OPTIONS)
{
}

BOOL GetLoggingValue(DWORD& dwData)
{
	HKEY  hKey = NULL;
	long lResult = RegOpenKeyEx(HKEY_CURRENT_USER, 
							    _T("Software\\Data Dynamics\\ActiveBar\\2.0"),
								0,
							    KEY_ALL_ACCESS, 
							    &hKey);
	if (ERROR_SUCCESS != lResult) 
		return FALSE;

	DWORD dwType = REG_DWORD;
	DWORD dwSize = sizeof(DWORD);
	lResult = RegQueryValueEx(hKey, 
							_T("Log Events"), 
							0,
							&dwType,
							(LPBYTE)&dwData, 
							&dwSize);

	RegCloseKey(hKey);
	if (ERROR_SUCCESS != lResult) 
		return FALSE;
	return TRUE;
}

BOOL SetLoggingValue(DWORD dwData)
{
	HKEY  hKey = NULL;
	long lResult = RegOpenKeyEx(HKEY_CURRENT_USER, 
							    _T("Software\\Data Dynamics\\ActiveBar\\2.0"),
								0,
							    KEY_ALL_ACCESS, 
							    &hKey);
	if (ERROR_SUCCESS != lResult) 
		return FALSE;

	lResult = RegSetValueEx(hKey, 
							_T("Log Events"), 
							0,
							REG_DWORD,
							(LPBYTE)&dwData, 
							sizeof(DWORD));

	RegCloseKey(hKey);
	if (ERROR_SUCCESS != lResult) 
	{
		LPTSTR sz = GetLastErrorMsg();
		return FALSE;
	}
	return TRUE;
}

//
// DialogProc
//

BOOL CDlgOptions::DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_INITDIALOG:
			{
				CenterDialog(GetParent(m_hWnd));
				
				if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_DELETEBANDCONFIRM))
					::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_BAND), BM_SETCHECK, BST_CHECKED, 0);
				
				if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_DELETECHILDBANDCONFIRM))
					::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_PAGE), BM_SETCHECK, BST_CHECKED, 0);
				
				if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_DELETETOOLCONFIRM))
					::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_TOOL), BM_SETCHECK, BST_CHECKED, 0);
				
				if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_DELETECATEGORYCONFIRM))
					::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_CATEGORY), BM_SETCHECK, BST_CHECKED, 0);
				
				if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_AUTOCREATEMASKCONFIRM))
					::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_AUTOCREATEMASK), BM_SETCHECK, BST_CHECKED, 0);

				if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_SORTBANDS))
					::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_SORTBANDS), BM_SETCHECK, BST_CHECKED, 0);

				TCHAR szTemp[MAX_PATH];
				if (GetGlobals().m_thePreferences.GetLibraryPath(szTemp))
					SetWindowText(GetDlgItem(m_hWnd, IDC_ED_LIBPATH), szTemp);

				DWORD dwData;
				if (GetLoggingValue(dwData) && 1 == dwData)
					CheckDlgButton(m_hWnd, IDC_CHK_EVENTLOGGING, BST_CHECKED);

				SetWindowPos(m_hWnd, HWND_TOP, 0,0,0,0, SWP_NOMOVE|SWP_NOSIZE);
			}
			break;

		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				switch (LOWORD(wParam))
				{
				case IDC_BN_GETPATH:
					{
						DDString strBuffer;
						strBuffer.LoadString(IDS_LAYOUTTYPE);
						DDString strFileType;
						strFileType.LoadString(IDS_LAYOUTFILETYPE);
						CFileDialog dlgFile(IDS_CHOOSELIBRARY,
											TRUE,
											strFileType,
											strFileType,
											OFN_OVERWRITEPROMPT|OFN_PATHMUSTEXIST|OFN_HIDEREADONLY,
											strBuffer,
											m_hWnd);

						if (0 == dlgFile.DoModal())
							break;
						SetWindowText(GetDlgItem(m_hWnd, IDC_ED_LIBPATH), dlgFile.GetFileName());
					}
					break;
				}
			}
			break;
		}
		return FDialog::DialogProc(nMsg, wParam, lParam);
	}
	catch (...)
	{
		assert(FALSE);
		return 0;
	}
}

void CDlgOptions::OnOK()
{
	if (BST_CHECKED == ::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_BAND), BM_GETCHECK, 0, 0))
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_DELETEBANDCONFIRM, 0);
	else
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_DELETEBANDCONFIRM);

	if (BST_CHECKED == ::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_PAGE), BM_GETCHECK, 0, 0))
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_DELETECHILDBANDCONFIRM, 0);
	else
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_DELETECHILDBANDCONFIRM);

	if (BST_CHECKED == ::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_TOOL), BM_GETCHECK, 0, 0))
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_DELETETOOLCONFIRM, 0);
	else
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_DELETETOOLCONFIRM);

	if (BST_CHECKED == ::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_CATEGORY), BM_GETCHECK, 0, 0))
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_DELETECATEGORYCONFIRM, 0);
	else
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_DELETECATEGORYCONFIRM);

	if (BST_CHECKED == ::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_AUTOCREATEMASK), BM_GETCHECK, 0, 0))
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_AUTOCREATEMASKCONFIRM, 0);
	else
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_AUTOCREATEMASKCONFIRM);

	if (BST_CHECKED == ::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_AUTOCREATEMASK), BM_GETCHECK, 0, 0))
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_AUTOCREATEMASKCONFIRM, 0);
	else
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_AUTOCREATEMASKCONFIRM);

	if (BST_CHECKED == ::SendMessage(GetDlgItem(m_hWnd, IDC_CHK_SORTBANDS), BM_GETCHECK, 0, 0))
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_SORTBANDS, 0);
	else
		GetGlobals().m_thePreferences.SetConfirmSuppress(IDS_SORTBANDS);

	if (BST_CHECKED == IsDlgButtonChecked(m_hWnd, IDC_CHK_EVENTLOGGING)) 
		SetLoggingValue(1);
	else
		SetLoggingValue(0);
	
	TCHAR szTemp[MAX_PATH];
	GetWindowText(GetDlgItem(m_hWnd, IDC_ED_LIBPATH), szTemp, MAX_PATH);
	if (szTemp)
		GetGlobals().m_thePreferences.SetLibraryPath(szTemp);

	FDialog::OnOK();
}

//
// CDlgConfirm 
//

BOOL CDlgConfirm::DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_INITDIALOG:
		SetWindowPos(m_hWnd, HWND_TOP, 0,0,0,0, SWP_NOMOVE|SWP_NOSIZE);
		m_bDontShow = FALSE;
		CenterDialog(GetParent(m_hWnd));
		DDString strText;
		strText.LoadString(m_nStringId);
		SendMessage(GetDlgItem(m_hWnd, IDC_MSG), WM_SETTEXT, 0, (LPARAM)(LPCTSTR)strText);
		break;
	}
	return FDialog::DialogProc(nMsg, wParam, lParam);
}

UINT CDlgConfirm::Show(HWND hWndParent)
{
	UINT nResult = IDOK;
	if (GetGlobals().m_thePreferences.GetConfirmSuppress(m_nStringId))
		return nResult;

	nResult = DoModal(hWndParent);
	if (m_bDontShow)
		GetGlobals().m_thePreferences.SetConfirmSuppress(m_nStringId);
	return nResult;
}

void CDlgConfirm::OnOK()
{
	if (BST_CHECKED == SendMessage(GetDlgItem(m_hWnd, IDC_DONTSHOW), BM_GETCHECK, 0, 0))
		m_bDontShow = TRUE;
	FDialog::OnOK();
}

void CDlgConfirm::OnCancel()
{
	if (BST_CHECKED == SendMessage(GetDlgItem(m_hWnd, IDC_DONTSHOW), BM_GETCHECK, 0, 0))
		m_bDontShow = TRUE;
	FDialog::OnCancel();
}

//
// CIconEditor
//

CIconEditor::CIconEditor(ITool* pTool) 
	: FDialog(IDD_ICONEDITOR),
	  m_pTool(pTool)
{
	m_pIconEditor = new CIconEdit;
	assert(m_pIconEditor);
	m_strMsgTitle.LoadString(IDS_ICONEDITOR);
	m_bPreTranslate = TRUE;
};

CIconEditor::~CIconEditor()
{
	delete m_pIconEditor;
}

//
// DialogProc
//

BOOL CIconEditor::DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch(nMsg)
	{
	case WM_INITDIALOG:
		{
			TCHAR szName[128];
			if (GetWindowText(m_hWnd, szName, 128) > 0)
			{
				BSTR bstrName;
				HRESULT hResult = m_pTool->get_Name(&bstrName);
				if (SUCCEEDED(hResult))
				{
					MAKE_TCHARPTR_FROMWIDE(szToolName, bstrName);
					wsprintf(szName, _T("%s - %s"), szName, szToolName);
					SetWindowText(m_hWnd, szName);
					SysFreeString(bstrName);
				}
			}
			SetWindowPos(m_hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE|SWP_NOMOVE);

			BOOL bResult = m_pIconEditor->CreateWin(m_hWnd, GetParent(m_hWnd));
			assert(bResult);

			m_pIconEditor->SetFont((HFONT)SendMessage(m_hWnd, WM_GETFONT, 0, 0));
			SetBitmaps();
			
			CRect rcDefault;
			CRect rc;
			GetWindowRect(m_hWnd, &rcDefault);
			GetGlobals().m_thePreferences.LoadWinRect(PREFID_ICONEDITOR, rc, rcDefault);
			MoveWindow(m_hWnd, rc.left, rc.top, rc.Width(), rc.Height(), TRUE);
			CRect rcWin;
			GetClientRect(m_hWnd, &rcWin);
			OnSize(rcWin.Width(), rcWin.Height());
		}
		break;

	case WM_SIZE:
		OnSize(LOWORD(lParam), HIWORD(lParam));
		return 0;

	case WM_NCDESTROY:
		GetGlobals().m_thePreferences.SaveWinRect(PREFID_ICONEDITOR, m_hWnd);
		break;
	}
	return FDialog::DialogProc(nMsg, wParam, lParam);
}

//
// PreTranslateMessage
//

BOOL CIconEditor::PreTranslateMessage(MSG* msg)
{
	switch (msg->message)
	{
	case WM_KEYDOWN:
		{
			switch (msg->wParam)
			{
			case 'Z':
			case 'z':
				if (0x8000 & GetKeyState(VK_CONTROL))
				{
					m_pIconEditor->OnUndo();
					return TRUE;
				}
				break;

			case 'X':
			case 'x':
				if (0x8000 & GetKeyState(VK_CONTROL))
				{
					m_pIconEditor->Cut();
					return TRUE;
				}
				break;

			case 'V':
			case 'v':
				if (0x8000 & GetKeyState(VK_CONTROL))
				{
					m_pIconEditor->Paste();
					return TRUE;
				}
				break;

			case 'C':
			case 'c':
				if (0x8000 & GetKeyState(VK_CONTROL))
				{
					m_pIconEditor->Copy();
					return TRUE;
				}
				break;

			case 'D':
			case 'd':
				if (0x8000 & GetKeyState(VK_CONTROL))
				{
					m_pIconEditor->ClearPicture();
					return TRUE;
				}
				break;
			}
		}
		break;
	}
	return FALSE;
}

//
// OnSize
//

void CIconEditor::OnSize(int nWidth, int nHeight)
{
	HWND hWndOk = GetDlgItem(m_hWnd, IDOK);
	CRect rcButton;
	::GetClientRect(hWndOk, &rcButton);

	CRect rcDefault(0, 
				    nWidth - rcButton.Width() - 8, 
					0, 
					nHeight);
	m_pIconEditor->MoveWindow(rcDefault, TRUE);

	rcButton.Offset(nWidth - rcButton.Width() - 8, 4);
	
	MoveWindow(hWndOk, 
			   rcButton.left, 
			   rcButton.top, 
			   rcButton.Width(), 
			   rcButton.Height(), 
			   TRUE);

	rcButton.Offset(0, rcButton.Height() + 4);

	MoveWindow(GetDlgItem(m_hWnd, IDCANCEL), 
			   rcButton.left, 
			   rcButton.top, 
			   rcButton.Width(), 
			   rcButton.Height(), 
			   TRUE);
}

//
// OnOK
//

void CIconEditor::OnOK()
{
	FDialog::OnOK();
	BOOL bResult;
	HBITMAP hBitmap;
	HBITMAP hBitmapMask;
	HRESULT hResult;
	for (int nImage = 0; nImage < Painter::eNumOfImages; nImage++)
	{
		if (m_pIconEditor->IsDirty(nImage))
		{
			m_pIconEditor->SetImageIndex(nImage);
			m_pIconEditor->GetBitmaps(hBitmap, hBitmapMask);
			hResult = m_pTool->put_Bitmap((ImageTypes)nImage, (OLE_HANDLE)hBitmap);
			if (FAILED(hResult))
			{
				if (CUSTOM_CTL_SCODE(IDERR_IMAGENOTSET) == hResult)
					continue;

				DDString strError;
				strError.Format(IDS_SETIMAGE, nImage);
				MessageBox((LPTSTR)strError);
			}
			hResult = m_pTool->put_MaskBitmap((ImageTypes)nImage, (OLE_HANDLE)hBitmapMask);
			if (hBitmap)
			{
				bResult = DeleteBitmap(hBitmap);
				assert(bResult);
			}
			if (hBitmapMask)
			{
				bResult = DeleteBitmap(hBitmapMask);
				assert(bResult);
			}
		}
	}
}

//
// SetBitmaps
//

void CIconEditor::SetBitmaps()
{
	BOOL bResult;
	COLORREF crMask;
	HBITMAP hBitmap;
	HBITMAP hBitmapMask;

	HRESULT hResult = m_pTool->GetMaskColor(&crMask);
	for (int nImage = 0; nImage < Painter::eNumOfImages; nImage++)
	{
		hResult = m_pTool->get_Bitmap((ImageTypes)nImage, (OLE_HANDLE*)&hBitmap);
		if (FAILED(hResult))
			continue;

		hResult = m_pTool->get_MaskBitmap((ImageTypes)nImage, (OLE_HANDLE*)&hBitmapMask);
		if (FAILED(hResult))
		{
			bResult = DeleteBitmap(hBitmap);
			assert(bResult);
			continue;
		}
		m_pIconEditor->SetImageIndex(nImage);
		m_pIconEditor->SetBitmaps(hBitmap, hBitmapMask, crMask); 
		bResult = DeleteBitmap(hBitmap);
		assert(bResult);
		bResult = DeleteBitmap(hBitmapMask);
		assert(bResult);
	}
	m_pIconEditor->SetImageIndex(0);
}

//
// CDlgLibrary
//

CDlgLibrary::CDlgLibrary() 
	: FDialog(IDD_LIBRARY),
	  m_pActiveBar(NULL)
{
	m_pToolBar = new DDToolBar;
	assert(m_pToolBar);
	
	m_pBrowser = new CBrowser;
	assert(m_pBrowser);
	
	m_pSplitter = new CSplitter;
	assert(m_pSplitter);

	m_pDesignerTree = new CDesignerTree(CDesignerTree::eLibrary);
	assert(m_pDesignerTree);

	m_pCategoryMgr = new CategoryMgr;
	assert(m_pCategoryMgr);

	m_bPreTranslate = TRUE;
	m_bDirty = FALSE;
}

CDlgLibrary::~CDlgLibrary()
{
	delete m_pToolBar;
	delete m_pBrowser;
	delete m_pSplitter;
	delete m_pDesignerTree;
	delete m_pCategoryMgr;
	if (m_pActiveBar)
		m_pActiveBar->Release();
}

//
// DialogProc
//

BOOL CDlgLibrary::DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_INITDIALOG:
			{
				if (m_pSplitter)
					m_pSplitter->hWnd(m_hWnd);

				HWND hWndToolBar = GetDlgItem(m_hWnd, IDC_TOOLBAR);
				if (IsWindow(hWndToolBar) && m_pToolBar)
				{
					CRect rc;
					GetClientRect(hWndToolBar, &rc);
					ClientToScreen(hWndToolBar, rc);
					ScreenToClient(m_hWnd, rc);
					DestroyWindow(hWndToolBar);
					m_pToolBar->LoadToolBar(MAKEINTRESOURCE(IDR_TBLIBRARY));
					m_pToolBar->SetStyle(DDToolBar::eRaised|DDToolBar::eIE);
					hWndToolBar = m_pToolBar->Create((HWND)m_hWnd, rc, ID_TOOLBAR);
					assert(hWndToolBar);
				}

				HRESULT hResult = CoCreateInstance(CLSID_ActiveBar2, NULL, CLSCTX_INPROC_SERVER, IID_IActiveBar2, (void**)&m_pActiveBar);
				if (FAILED(hResult))
				{
					MessageBox(IDS_ERR_FAILEDTOCREATEDESIGNER);
					return FALSE;
				}

				if (GetGlobals().m_thePreferences.GetLibraryPath(m_pDesignerTree->CurFileName()))
				{
					VARIANT vFileName;
					vFileName.vt = VT_BSTR;
					MAKE_WIDEPTR_FROMTCHAR(wFile, m_pDesignerTree->CurFileName());
					vFileName.bstrVal = SysAllocString(wFile);
					if (vFileName.bstrVal)
					{
						hResult = m_pActiveBar->Load(L"", &vFileName, ddSOFile);
						if (FAILED(hResult))
							MessageBox(IDS_ERR_FAILEDTOLOADDEFAULTLIBRARY);
						VariantClear(&vFileName);
					}
					else
						MessageBox(IDS_ERR_FAILEDTOLOADDEFAULTLIBRARY);
				}
				
				IBarPrivate* pPrivate;
				hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
				if (SUCCEEDED(hResult))
				{
					pPrivate->put_Library(VARIANT_TRUE);
					pPrivate->Release();
				}

				if (m_pCategoryMgr)
				{
					m_pCategoryMgr->ActiveBar(m_pActiveBar);
					HWND hWndProperty = GetDlgItem(m_hWnd, IDC_LST_PROPERTIES);
					if (IsWindow(hWndProperty) && m_pBrowser)
					{
						m_pBrowser->SubClassAttach(hWndProperty);
						HWND hWndTree = GetDlgItem(m_hWnd, IDC_TREE_DESIGNER);
						if (IsWindow(hWndTree) && m_pDesignerTree)
						{
							m_pDesignerTree->SubClassAttach(hWndTree);
							m_pDesignerTree->Init(m_pActiveBar, m_pBrowser, m_pCategoryMgr);
						}
					}
				}
				CRect rcDefault;
				CRect rc;
				GetWindowRect(m_hWnd, &rcDefault);
				GetGlobals().m_thePreferences.LoadWinRect(PREFID_LIBRARY, rc, rcDefault);
				SetWindowPos(m_hWnd, NULL, rc.left, rc.top, rc.Width(), rc.Height(), SWP_SHOWWINDOW);
				OnSize();
			}
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
					WinHelp(m_hWnd, HELP_CONTEXT, dwData);
					return TRUE;
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;
			}
			break;

		case WM_SYSCOLORCHANGE:
		case WM_SETTINGCHANGE:
			m_pDesignerTree->SettingsChanged();
			break;

		case WM_SIZE:
			OnSize();
			break;

		case WM_DESTROY:
			if (m_pDesignerTree)
				TreeView_DeleteAllItems(m_pDesignerTree->hWnd());
			if (m_pBrowser)
				m_pBrowser->SendMessage(LB_RESETCONTENT);
			WinHelp(m_hWnd, NULL, HELP_QUIT, NULL);
			break;

		case WM_NCDESTROY:
			if (m_pDesignerTree)
				m_pDesignerTree->UnsubClass();
			if (m_pBrowser)
				m_pBrowser->UnsubClass();
			GetGlobals().m_thePreferences.SaveWinRect(PREFID_LIBRARY, m_hWnd);
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

		case WM_SYSCOMMAND:
			if (SC_CLOSE == wParam && GetGlobals().m_thePreferences.GetLibraryPath(m_pDesignerTree->CurFileName()))
			{
				VARIANT vFileName;
				vFileName.vt = VT_EMPTY;
				MAKE_WIDEPTR_FROMTCHAR(wFile, m_pDesignerTree->CurFileName());
				BSTR bstrFile = SysAllocString(wFile);
				if (bstrFile)
				{
					HRESULT hResult = m_pActiveBar->Save(L"", bstrFile, ddSOFile, &vFileName);
					VariantClear(&vFileName);
				}
			}
			break;

		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				switch (LOWORD(wParam))
				{
				case ID_FILEOPEN:
					if (m_pDesignerTree)
						m_pDesignerTree->OnOpenLayout();
					break;

				case ID_FILENEW:
					if (m_pDesignerTree)
						m_pDesignerTree->OnNewLayout();
					break;

				case ID_FILESAVE:
					if (m_pDesignerTree)
						m_pDesignerTree->OnSaveLayout();
					break;

				case ID_FILE_SAVEAS:
					if (m_pDesignerTree)
						m_pDesignerTree->OnSaveAsLayout();
					break;

				case ID_HELP_CONTENTS:
					WinHelp(m_hWnd, HELP_CONTEXT, HX_ActiveBar_Designer);
					break;

				case ID_HELP_ABOUT:
					if (m_pActiveBar)
						m_pActiveBar->About();
					break;

				case ID_CANCEL:
					EndDialog(m_hWnd, ID_CANCEL);
					break;

				case ID_FILEEXIT:
					if (GetGlobals().m_thePreferences.GetLibraryPath(m_pDesignerTree->CurFileName()))
					{
						VARIANT vFileName;
						vFileName.vt = VT_EMPTY;
						MAKE_WIDEPTR_FROMTCHAR(wFile, m_pDesignerTree->CurFileName());
						BSTR bstrFile = SysAllocString(wFile);
						if (bstrFile)
						{
							HRESULT hResult = m_pActiveBar->Save(L"", bstrFile, ddSOFile, &vFileName);
							VariantClear(&vFileName);
						}
					}
					EndDialog(m_hWnd, IDOK);
					break;
				}
				break;

			case EN_CHANGE:
			case EN_UPDATE:
			case EN_SETFOCUS:
			case EN_KILLFOCUS:
				return 0;

			case LBN_DBLCLK:
			case LBN_SELCANCEL:
			case LBN_SELCHANGE:
				switch (LOWORD(wParam))
				{
				case IDC_LST_PROPERTIES:
					::SendMessage((HWND)lParam, WM_COMMAND, wParam, lParam);
					break;
				}
				break;
			}
			break;

		case WM_NOTIFY:
			switch (wParam)
			{
			case IDC_TREE_DESIGNER:
				{
					HWND hWndChild = GetDlgItem(m_hWnd, IDC_TREE_DESIGNER);
					if (IsWindow(hWndChild))
					{
						LRESULT lResult = SendMessage(hWndChild, nMsg, wParam, lParam);
						SetWindowLong(m_hWnd, DWL_MSGRESULT, lResult);
						return lResult;
					}
				}
				break;

			case ID_SETPAGEDIRTY:
				m_bDirty = TRUE;
				break;
			}
			break;

		case WM_CHARTOITEM:
		case WM_VKEYTOITEM:
			{
				LRESULT lResult = SendMessage((HWND)lParam, WM_VKEYTOITEM, wParam, lParam);
				SetWindowLong(m_hWnd, DWL_MSGRESULT, lResult);
				return lResult;
			}
			break;

		case WM_COMPAREITEM:
		case WM_DELETEITEM:
		case WM_DRAWITEM:
			if (IDC_LST_PROPERTIES == wParam)
			{
				HWND hWndChild = GetDlgItem(m_hWnd, IDC_LST_PROPERTIES);
				if (IsWindow(hWndChild))
					return SendMessage(hWndChild, nMsg, wParam, lParam);
			}
			break;
		}
		return FDialog::DialogProc(nMsg, wParam, lParam);
	}
	catch (...)
	{
		assert(FALSE);
		return 0;
	}
}


//
// PreTranslateMessage
//

BOOL CDlgLibrary::PreTranslateMessage(MSG* pMsg)
{
	switch (pMsg->message)
	{
	case WM_KEYUP:
	case WM_KEYDOWN:
		{
			switch (pMsg->wParam)
			{
			case VK_RETURN:
			case VK_ESCAPE:
				try
				{
					if (m_pDesignerTree)
					{
						HWND hWndEdit = TreeView_GetEditControl(m_pDesignerTree->hWnd());
						if (hWndEdit && pMsg->hwnd == hWndEdit)
						{
							::SendMessage(hWndEdit, pMsg->message, pMsg->wParam, pMsg->lParam);
							return TRUE;
						}
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;
			}
		}
		break;
	}
	return FALSE;
}

//
// OnSize
//

void CDlgLibrary::OnSize()
{
	CRect rcClient;
	GetClientRect(m_hWnd, &rcClient);

	CRect rcTree = rcClient;
	rcTree.Inflate(-5, -5);
	rcTree.top += 35;
	rcTree.right = rcClient.Width() / 2 - 2; 
	m_pDesignerTree->MoveWindow(rcTree);

	CRect rcBrowser = rcClient;
	rcBrowser.Inflate(-5, -5);
	rcBrowser.top += 35;
	rcBrowser.left = rcClient.Width() / 2 + 2; 
	m_pBrowser->MoveWindow(rcBrowser);

	CRect rcToolBar = rcClient;
	rcToolBar.Inflate(-5, -5);
	rcToolBar.bottom = rcToolBar.top + 30;
	m_pToolBar->MoveWindow(rcToolBar, TRUE);

	if (m_pSplitter)
	{
		m_pSplitter->SetLeftLimit(rcTree.left + 4);
		m_pSplitter->SetRightLimit(rcBrowser.right - 8);
		m_pSplitter->Tracker(rcTree.right, rcBrowser.left, rcTree.top, rcTree.bottom);
	}
	m_pBrowser->InvalidateRect(NULL, TRUE);
}

//
// OnLButtonDown
//

void CDlgLibrary::OnLButtonDown(UINT nFlags, POINT pt)
{
	if (m_pSplitter->SplitterHitTest(pt))
		m_pSplitter->StartTracking();
}

//
// OnMouseMove
//

void CDlgLibrary::OnMouseMove(UINT nFlags, POINT pt)
{
	m_pSplitter->DoTracking(pt);
}

//
// OnLButtonUp
//

void CDlgLibrary::OnLButtonUp(UINT nFlags, POINT pt)
{
	m_pSplitter->StopTracking(pt);
	
	CRect rcTree;
	GetWindowRect(GetDlgItem(m_hWnd, IDC_TREE_DESIGNER), &rcTree);
	ScreenToClient(m_hWnd, rcTree);
	rcTree.right = m_pSplitter->TrackerLeft();
	MoveWindow(GetDlgItem(m_hWnd, IDC_TREE_DESIGNER),
		       rcTree.left,
			   rcTree.top,
			   rcTree.Width(),
			   rcTree.Height(),
			   TRUE);

	CRect rcProperty;
	GetWindowRect(GetDlgItem(m_hWnd, IDC_LST_PROPERTIES), &rcProperty);
	ScreenToClient(m_hWnd, rcProperty);
	rcProperty.left = m_pSplitter->TrackerRight();
	MoveWindow(GetDlgItem(m_hWnd, IDC_LST_PROPERTIES),
		       rcProperty.left,
			   rcProperty.top,
			   rcProperty.Width(),
			   rcProperty.Height(),
			   TRUE);
	
	InvalidateRect(GetDlgItem(m_hWnd, IDC_LST_PROPERTIES), NULL, TRUE);
}

//
// CDlgDesigner
//

CDlgDesigner::CDlgDesigner() 
	: FDialog(IDD_DESIGNER1),
	  m_pActiveBar(NULL)
{
	m_pDesignerTree = new CDesignerTree;
	assert(m_pDesignerTree);

	m_pCategoryMgr = new CategoryMgr;
	assert(m_pCategoryMgr);

	m_pSplitter = new CSplitter;
	assert(m_pSplitter);

	m_pToolBar = new DDToolBar;
	assert(m_pToolBar);
	
	m_pBrowser = new CBrowser;
	assert(m_pBrowser);
	
	m_bPreTranslate = TRUE;
	m_bFileDirty = FALSE;
	m_bDirty = FALSE;

	m_vLayoutData.vt = VT_EMPTY;
}

CDlgDesigner::~CDlgDesigner()
{
	delete m_pDesignerTree;
	delete m_pCategoryMgr;
	delete m_pSplitter;
	delete m_pToolBar;
	delete m_pBrowser;
	if (m_pActiveBar)
		m_pActiveBar->Release();
}

void CDlgDesigner::Apply()
{
	if (m_pActiveBar)
	{
		IBarPrivate* pPrivate;
		HRESULT hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
		if (SUCCEEDED(hResult))
		{
			m_bDirty = FALSE;
			pPrivate->SetDesignerModified();
			pPrivate->Release();
		}
	}
}

//
// Init
//

void CDlgDesigner::Init(IActiveBar2* pActiveBar, BOOL bStandALone)
{
	m_bStandALone = bStandALone;
	if (m_pActiveBar)
	{
		m_pActiveBar->Release();
		m_pActiveBar = NULL;
	}
	m_pActiveBar = pActiveBar;
	m_pActiveBar->AddRef();
	m_strMsgTitle.LoadString(IDS_DESIGNERTITLE);
}

//
// DialogProc
//

extern HWND g_hDlgCurrent;

BOOL CDlgDesigner::DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_INITDIALOG:
			{
				if (NULL == m_pActiveBar)
					break;

				g_hDlgCurrent = m_hWnd;

				if (m_pSplitter)
					m_pSplitter->hWnd(m_hWnd);

				HWND hWndToolBar = GetDlgItem(m_hWnd, IDC_TOOLBAR);
				if (IsWindow(hWndToolBar) && m_pToolBar)
				{
					CRect rc;
					GetClientRect(hWndToolBar, &rc);
					ClientToScreen(hWndToolBar, rc);
					ScreenToClient(m_hWnd, rc);
					DestroyWindow(hWndToolBar);
					m_pToolBar->LoadToolBar(MAKEINTRESOURCE(IDR_TBDESIGNER));
					m_pToolBar->SetStyle(DDToolBar::eRaised|DDToolBar::eIE);
					m_pToolBar->Create((HWND)m_hWnd, rc, ID_TOOLBAR);
				}

				BSTR bstrVersion;
				HRESULT hResult = m_pActiveBar->get_Version(&bstrVersion);
				if (FAILED(hResult))
				{
					MessageBox(IDS_ERR_VERSIONNUMBER);
					OnSize();
					return TRUE;
				}

				MAKE_TCHARPTR_FROMWIDE(szVersionAB, bstrVersion);
				SysFreeString(bstrVersion);

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
					PostQuitMessage(0);
					return TRUE;
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

				if (m_pCategoryMgr)
				{
					m_pCategoryMgr->ActiveBar(m_pActiveBar);
					HWND hWndProperty = GetDlgItem(m_hWnd, IDC_LST_PROPERTIES);
					if (IsWindow(hWndProperty) && m_pBrowser)
					{
						m_pBrowser->SubClassAttach(hWndProperty);
						m_pBrowser->Init();
						HWND hWndTree = GetDlgItem(m_hWnd, IDC_TREE_DESIGNER);
						if (IsWindow(hWndTree) && m_pDesignerTree)
						{
							m_pDesignerTree->SubClassAttach(hWndTree);
							m_pDesignerTree->Init(m_pActiveBar, m_pBrowser, m_pCategoryMgr);
						}
					}
				}
				CRect rcDefault;
				CRect rc;
				GetWindowRect(m_hWnd, &rcDefault);
				GetGlobals().m_thePreferences.LoadWinRect(PREFID_DESIGNER, rc, rcDefault);
				SetWindowPos(m_hWnd, NULL, rc.left, rc.top, rc.Width(), rc.Height(), SWP_SHOWWINDOW);
				OnSize();
				if (m_pSplitter)
				{
					//
					// This is not complete
					//

					if (GetGlobals().m_thePreferences.LoadWinRect(PREFID_DESIGNERTRACKER, rc))
					{
						//
						// I have to size the windows based on the tracker's position
						//
						m_pSplitter->Tracker(rc);
						POINT pt = {rc.top, rc.left};
						OnLButtonUp(0, pt);
						InvalidateRect(m_hWnd, NULL, FALSE);
					}
				}
				HICON hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_MAINWIN));
				if (hIcon)
				{
					SendMessage(m_hWnd, WM_SETICON, ICON_BIG, (LPARAM)hIcon);
					DestroyIcon(hIcon);
				}
			}
			break;

		case WM_UPDATEWINDOW:
			InvalidateRect(m_hWnd, NULL, TRUE);
			m_pBrowser->InvalidateRect(NULL, TRUE);
			InvalidateRect(GetDlgItem(m_hWnd, IDC_ST_DESC), NULL, TRUE);
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

					case IDC_STATICTREE:
						dwData = HX_Designer_Toolbar;
						break;
					}
					WinHelp(m_hWnd, HELP_CONTEXT, dwData);
					return TRUE;
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;
			}
			break;
			
		case WM_SYSCOLORCHANGE:
		case WM_SETTINGCHANGE:
			m_pDesignerTree->SettingsChanged();
			break;

		case WM_SIZE:
			OnSize();
			return 0;

		case WM_DESTROY:
			try
			{
				if (m_pBrowser && IsWindow(m_pBrowser->hWnd()))
				{
					m_pBrowser->SendMessage(LB_RESETCONTENT);
					m_pBrowser->ReleaseObjects();
				}

				if (m_pDesignerTree && IsWindow(m_pDesignerTree->hWnd()))
					TreeView_DeleteAllItems(m_pDesignerTree->hWnd());

				UIUtilities::CleanupUIUtilities();

				if (IsWindow(m_hWnd))
					WinHelp(m_hWnd, NULL, HELP_QUIT, NULL);

				DWORD dwStyle = ::GetWindowLong(m_hWnd, GWL_STYLE);
				if (!(dwStyle & WS_MINIMIZE))
				{
					GetGlobals().m_thePreferences.SaveWinRect(PREFID_DESIGNER, m_hWnd);
					if (m_pSplitter)
						GetGlobals().m_thePreferences.SaveWinRect(PREFID_DESIGNERTRACKER, m_pSplitter->Tracker());
				}

				//
				// We need to make sure all extra reference counts are removed at this point.
				//

				if (m_pDesignerTree)
					m_pDesignerTree->UnsubClass();

				if (m_pBrowser)
					m_pBrowser->UnsubClass();

				delete m_pCategoryMgr;
				m_pCategoryMgr = NULL;

				if (m_bStandALone)
				{
					if (m_bFileDirty || m_bDirty)
					{
						int nResult = MessageBox(IDS_COMMITCHANGES, MB_YESNO);
						if (IDYES == nResult)
						{
							if (m_pDesignerTree)
								m_pDesignerTree->OnSaveLayout();
						}
					}
				}
				else
				{
					IBarPrivate* pPrivate;
					HRESULT hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
					if (SUCCEEDED(hResult))
					{
						if (m_bDirty && VT_EMPTY != m_vLayoutData.vt)
						{
							hResult = m_pActiveBar->Load(NULL, &m_vLayoutData, ddSOByteArray);
							if (FAILED(hResult))
								MessageBox(IDS_FAILEDTORESTORELAYOUT);
							else
								m_pActiveBar->RecalcLayout();
						}
						pPrivate->Release();
					}
				}
				g_hDlgCurrent = NULL;
				PostQuitMessage(0);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__);
			}
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

		case WM_SYSCOMMAND:
			if (SC_CLOSE == wParam && m_bDirty)
			{
				if (m_pActiveBar)
				{
					int nResult = MessageBox(IDS_COMMITCHANGES, MB_YESNO);
					if (IDYES == nResult)
					{
						IBarPrivate* pPrivate;
						HRESULT hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
						if (SUCCEEDED(hResult))
						{
							m_bDirty = FALSE;
							pPrivate->SetDesignerModified();
							pPrivate->Release();
							m_pActiveBar->RecalcLayout();
						}
					}
				}
			}
			break;

		case WM_COMMAND:
			try
			{
				switch (HIWORD(wParam))
				{
				case EN_CHANGE:
				case EN_UPDATE:
				case EN_SETFOCUS:
				case EN_KILLFOCUS:
					return 0;

				case LBN_DBLCLK:
				case LBN_SELCANCEL:
				case LBN_SELCHANGE:
					switch (LOWORD(wParam))
					{
					case IDC_LST_PROPERTIES:
						::SendMessage((HWND)lParam, WM_COMMAND, wParam, lParam);
						break;
					}
					break;

				case BN_CLICKED:
					switch (LOWORD(wParam))
					{
					case IDOK:
					case IDAPPLY:
					
						if (m_bDirty)
							m_bFileDirty = TRUE;
						Apply();
						if (IDOK == LOWORD(wParam))
						{
							DestroyWindow(m_hWnd);
							return 0;
						}
						else
							EnableWindow(GetDlgItem(m_hWnd, IDAPPLY), FALSE);
						break;

					case IDCANCEL:
						DestroyWindow(m_hWnd);
						return 0;

					case IDHELP:
						WinHelp(m_hWnd, HELP_CONTEXT, HX_ActiveBar_Designer);
						break;

					case ID_FILENEW:
						if (m_pDesignerTree)
						{
							m_pDesignerTree->OnNewLayout();
							if (m_bStandALone)
							{
								m_bFileDirty = FALSE;
								m_bDirty = FALSE;
							}
						}
						break;

					case ID_FILE_SAVEAS:
						if (m_pDesignerTree)
						{
							m_pDesignerTree->OnSaveAsLayout();
							if (m_bStandALone)
							{
								m_bFileDirty = FALSE;
								m_bDirty = FALSE;
							}
						}
						break;

					case ID_FILEOPEN:
						if (m_pDesignerTree)
						{
							m_pDesignerTree->OnOpenLayout();
							if (m_bStandALone)
							{
								m_bFileDirty = FALSE;
								m_bDirty = FALSE;
							}
						}
						break;

					case ID_FILESAVE:
						{
							if (m_pDesignerTree)
								m_pDesignerTree->OnSaveLayout();
							if (m_bStandALone)
							{
								m_bFileDirty = FALSE;
								m_bDirty = FALSE;
							}
						}
						break;

					case ID_FILEEXIT:
						DestroyWindow(m_hWnd);
						break;

					case ID_OPTIONS:
						OnOptions();
						break;

					case ID_REPORT:
						OnReport();
						break;

					case ID_HEADER:
						OnHeader();
						break;

					case ID_LIBRARY:
						OnLibrary();
						break;

					case ID_MENUGRAB:
						OnMenuGrab();
						break;

					case ID_HELP_CONTENTS:
						WinHelp(m_hWnd, HELP_CONTEXT, HX_ActiveBar_Designer);
						break;

					case ID_HELP_ABOUT:
						if (m_pActiveBar)
							m_pActiveBar->About();
						break;

					case ID_GEN_SELECT:
						OnGenSelect();
						break;

					case ID_VERIFYIMAGEREFERENCECOUNT:
						OnVerifyImageReferenceCount();
						break;
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__);
			}
			break;

		case WM_NOTIFY:
			switch (wParam)
			{
			case IDC_TREE_DESIGNER:
				{
					try
					{
						LPNMHDR pnmh = (LPNMHDR) lParam; 
						HWND hWndChild = GetDlgItem(m_hWnd, IDC_TREE_DESIGNER);
						if (IsWindow(hWndChild) && pnmh->hwndFrom == hWndChild)
						{
							LRESULT lResult = SendMessage(hWndChild, nMsg, wParam, lParam);
							SetWindowLong(m_hWnd, DWL_MSGRESULT, lResult);
							return lResult;
						}
					}
					CATCH 
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
						return 0;
					}
				}
				break;

			case ID_SETPAGEDIRTY:
				SetDirty();
				break;
			}
			break;

		case WM_CHARTOITEM:
		case WM_VKEYTOITEM:
			{
				LRESULT lResult = SendMessage((HWND)lParam, WM_VKEYTOITEM, wParam, lParam);
				SetWindowLong(m_hWnd, DWL_MSGRESULT, lResult);
				return lResult;
			}
			break;

		case WM_COMPAREITEM:
		case WM_DELETEITEM:
		case WM_DRAWITEM:
			if (IDC_LST_PROPERTIES == wParam && IsWindow(m_hWnd))
			{
				HWND hWndChild = GetDlgItem(m_hWnd, IDC_LST_PROPERTIES);
				if (IsWindow(hWndChild))
					return SendMessage(hWndChild, nMsg, wParam, lParam);
			}
			break;
		}
		return FDialog::DialogProc(nMsg, wParam, lParam);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
}

//
// PreTranslateMessage
//

BOOL CDlgDesigner::PreTranslateMessage(MSG* pMsg)
{
	switch (pMsg->message)
	{
	case WM_KEYUP:
	case WM_KEYDOWN:
		{
			switch (pMsg->wParam)
			{
			case VK_ESCAPE:
			case VK_RETURN:
				try
				{
					if (m_pDesignerTree)
					{
						HWND hWndEdit = TreeView_GetEditControl(m_pDesignerTree->hWnd());
						if (IsWindow(hWndEdit) && pMsg->hwnd == hWndEdit)
						{
							::SendMessage(hWndEdit, pMsg->message, pMsg->wParam, pMsg->lParam);
							return TRUE;
						}
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;
			}
		}
	}
	return FALSE;
}

//
// SetDirty
//

STDMETHODIMP CDlgDesigner::SetDirty()
{
	try
	{
		m_bDirty = TRUE;
		EnableWindow(GetDlgItem(m_hWnd, IDAPPLY), TRUE);
		::InvalidateRect(GetDlgItem(m_hWnd, IDAPPLY), NULL, FALSE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return NOERROR;
}

//
// CreateTool
//

STDMETHODIMP CDlgDesigner::CreateTool(IDispatch* pBandIn, long nToolType)
{
	try
	{
		IBand* pBand = (IBand*)pBandIn;

		if (NULL == pBand)
			return NOERROR;

		int nToolId = FindLastToolId(m_pActiveBar) + 1;
		
		DDString strTool;
		strTool.Format(IDS_TOOLNAME, nToolId);
		
		ITools* pTools;
		HRESULT hResult = pBand->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
		{
			pTools->Release();
			return hResult;
		}

		BSTR bstrTool = strTool.AllocSysString();
		if (NULL == bstrTool)
		{
			pTools->Release();
			return E_FAIL;
		}

		ITool* pTool;
		hResult = pTools->Add(nToolId, bstrTool, (Tool**)&pTool);
		pTools->Release();
		if (FAILED(hResult))
		{
			SysFreeString(bstrTool);
			return hResult;
		}

		if (ddTTSeparator != nToolType)
		{
			hResult = pTool->put_Caption(bstrTool);
			SysFreeString(bstrTool);
			if (FAILED(hResult))
			{
				pTool->Release();
				return hResult;
			}
		}
		else
			SysFreeString(bstrTool);

		hResult = pTool->put_ControlType((ToolTypes)nToolType);
		if (FAILED(hResult))
		{
			pTool->Release();
			return hResult;
		}

		hResult = m_pActiveBar->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
		{
			pTool->Release();
			return hResult;
		}

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
			pChildBands->Release();
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
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}	
}

//
// OnSize
//

void CDlgDesigner::OnSize()
{
	CRect rcClient;
	GetClientRect(m_hWnd, &rcClient);

	HDWP hDWP = ::BeginDeferWindowPos(8); 
	if (NULL == hDWP)
		return;

	CRect rcTemp = rcClient;
	rcTemp.Inflate(-5, -5);
	rcTemp.bottom = rcTemp.top + 30;
	hDWP = DeferWindowPos(hDWP, m_pToolBar->hWnd(), NULL, rcTemp.left, rcTemp.top, rcTemp.Width(), rcTemp.Height(), SWP_SHOWWINDOW | SWP_NOACTIVATE);

	CRect rcTree = rcClient;
	rcTree.Inflate(-5, -40);
	rcTree.bottom -= 10;
	rcTree.right = rcClient.Width() / 2 - 2; 
	hDWP = DeferWindowPos(hDWP, m_pDesignerTree->hWnd(), NULL, rcTree.left, rcTree.top, rcTree.Width(), rcTree.Height(), SWP_SHOWWINDOW);

	CRect rcBrowser = rcClient;
	rcBrowser.Inflate(-5, -40);
	rcBrowser.bottom -= 50;
	rcBrowser.left = rcClient.Width() / 2 + 2; 
	hDWP = DeferWindowPos(hDWP, m_pBrowser->hWnd(), NULL, rcBrowser.left, rcBrowser.top, rcBrowser.Width(), rcBrowser.Height(), SWP_SHOWWINDOW);

	CRect rcButton = rcBrowser;
	HWND hWndTemp = GetDlgItem(m_hWnd, IDC_ST_DESC);
	if (hWndTemp)
	{
		rcButton.top = rcButton.bottom + 5;
		rcButton.bottom = rcButton.top + 35;
		hDWP = DeferWindowPos(hDWP, hWndTemp, NULL, rcButton.left, rcButton.top, rcButton.Width(), rcButton.Height(), SWP_SHOWWINDOW);
		rcBrowser.bottom += 30;
	}
	
	if (m_pSplitter)
	{
		m_pSplitter->SetLeftLimit(rcTree.left + 4);
		m_pSplitter->SetRightLimit(rcBrowser.right - 8);
		m_pSplitter->Tracker(rcTree.right, rcBrowser.left, rcTree.top, rcTree.bottom);
	}

	hWndTemp = GetDlgItem(m_hWnd, IDHELP);
	if (hWndTemp)
	{
		GetClientRect(hWndTemp, &rcTemp);
		rcButton.top = rcBrowser.bottom + 20;
		rcButton.left = rcBrowser.right - rcTemp.Width();
		rcButton.bottom = rcButton.top + rcTemp.Height();
		rcButton.right = rcBrowser.right;
		hDWP = DeferWindowPos(hDWP, hWndTemp, NULL, rcButton.left, rcButton.top, rcButton.Width(), rcButton.Height(), SWP_SHOWWINDOW);
	}

	hWndTemp = GetDlgItem(m_hWnd, IDAPPLY);
	if (hWndTemp)
	{
		GetClientRect(hWndTemp, &rcTemp);
		rcButton.Offset(-(rcTemp.Width() + 5), 0);
		hDWP = DeferWindowPos(hDWP, hWndTemp, NULL, rcButton.left, rcButton.top, rcButton.Width(), rcButton.Height(), SWP_SHOWWINDOW);
	}

	hWndTemp = GetDlgItem(m_hWnd, IDCANCEL);
	if (hWndTemp)
	{
		GetClientRect(hWndTemp, &rcTemp);
		rcButton.Offset(-(rcTemp.Width() + 5), 0);
		hDWP = DeferWindowPos(hDWP, hWndTemp, NULL, rcButton.left, rcButton.top, rcButton.Width(), rcButton.Height(), SWP_SHOWWINDOW);
	}
	
	hWndTemp = GetDlgItem(m_hWnd, IDOK);
	if (hWndTemp)
	{
		GetClientRect(hWndTemp, &rcTemp);
		rcButton.Offset(-(rcTemp.Width() + 5), 0);
		hDWP = DeferWindowPos(hDWP, hWndTemp, NULL, rcButton.left, rcButton.top, rcButton.Width(), rcButton.Height(), SWP_SHOWWINDOW);
	}	
	BOOL bResult = EndDeferWindowPos(hDWP);
	assert(bResult);
	PostMessage(m_hWnd, WM_UPDATEWINDOW, 0, 0);
}

//
// OnLButtonDown
//

void CDlgDesigner::OnLButtonDown(UINT nFlags, POINT pt)
{
	if (m_pSplitter && m_pSplitter->SplitterHitTest(pt))
		m_pSplitter->StartTracking();
}

//
// OnMouseMove
//

void CDlgDesigner::OnMouseMove(UINT nFlags, POINT pt)
{
	if (m_pSplitter)
		m_pSplitter->DoTracking(pt);
}

//
// OnLButtonUp
//

void CDlgDesigner::OnLButtonUp(UINT nFlags, POINT pt)
{
	if (NULL == m_pSplitter)
		return;

	m_pSplitter->StopTracking(pt);
	
	CRect rcTree;
	GetWindowRect(GetDlgItem(m_hWnd, IDC_TREE_DESIGNER), &rcTree);
	ScreenToClient(m_hWnd, rcTree);
	rcTree.right = m_pSplitter->TrackerLeft();
	MoveWindow(GetDlgItem(m_hWnd, IDC_TREE_DESIGNER),
		       rcTree.left,
			   rcTree.top,
			   rcTree.Width(),
			   rcTree.Height(),
			   TRUE);

	CRect rcProperty;
	GetWindowRect(GetDlgItem(m_hWnd, IDC_LST_PROPERTIES), &rcProperty);
	ScreenToClient(m_hWnd, rcProperty);
	rcProperty.left = m_pSplitter->TrackerRight();
	MoveWindow(GetDlgItem(m_hWnd, IDC_LST_PROPERTIES),
		       rcProperty.left,
			   rcProperty.top,
			   rcProperty.Width(),
			   rcProperty.Height(),
			   TRUE);
	
	InvalidateRect(GetDlgItem(m_hWnd, IDC_LST_PROPERTIES), NULL, TRUE);

	GetWindowRect(GetDlgItem(m_hWnd, IDC_ST_DESC), &rcProperty);
	ScreenToClient(m_hWnd, rcProperty);
	rcProperty.left = m_pSplitter->TrackerRight();

	MoveWindow(GetDlgItem(m_hWnd, IDC_ST_DESC),
		       rcProperty.left,
			   rcProperty.top,
			   rcProperty.Width(),
			   rcProperty.Height(),
			   TRUE);

	InvalidateRect(GetDlgItem(m_hWnd, IDC_ST_DESC), NULL, TRUE);
}

//
// OnOptions
//

void CDlgDesigner::OnOptions()
{
	try
	{
		CDlgOptions theOptions;
		theOptions.DoModal();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnReport
//

void CDlgDesigner::OnReport()
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
							m_hWnd);

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
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnHeader
//

void CDlgDesigner::OnHeader()
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
							m_hWnd);

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
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnGenSelect
//

void CDlgDesigner::OnGenSelect()
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

	if (!OpenClipboard(m_hWnd))
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
// OnVerifyImageReferenceCount
//

void CDlgDesigner::OnVerifyImageReferenceCount()
{
	if (m_pActiveBar)
	{
		IBarPrivate* pPrivate;
		HRESULT hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
		if (SUCCEEDED(hResult))
		{
			pPrivate->VerifyAndCorrectImages();
			SetDirty();
			pPrivate->Release();
		}
	}
}

//
// OnLibrary
//

void CDlgDesigner::OnLibrary()
{
	try
	{
		static BOOL bShown = FALSE;
		if (bShown)
			return;
		bShown = TRUE;
		CDlgLibrary theLibrary;
		theLibrary.DoModal();
		bShown = FALSE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnMenuGrab
//

void CDlgDesigner::OnMenuGrab()
{
	HWND hWndMenu = GrabMenu::FindMenuWindow(GetParent(m_hWnd));
	if (IsWindow(hWndMenu))
	{
		HMENU hMenu = GetMenu(hWndMenu);
		if (hMenu)
		{
			GrabMenu theMenu(hMenu, m_pActiveBar);
			theMenu.DoIt();
			SetDirty();
		}
		else
			MessageBeep(0xFFFFFFFF);
	}
	else if (!GrabMenu::m_bCancelled)
		MessageBeep(0xFFFFFFFF);
}

//
// CColor
//

CColor::CColor()
{
	m_nRed = 0;
	m_nGreen = 0;
	m_nBlue = 0;
	RGBToHLS();
}

CColor::CColor(COLORREF crColor)
{
	m_nRed = GetRValue(crColor);
	m_nGreen = GetGValue(crColor);
	m_nBlue = GetBValue(crColor);
	RGBToHLS();
}

void CColor::RGBToHLS()
{
	int cMax = max(max(m_nRed, m_nGreen), m_nBlue);
	int cMin = min(min(m_nRed, m_nGreen), m_nBlue);

	m_nLum = (((cMax + cMin) * eLSMAX) + eRGBMAX ) / (2 * eRGBMAX);

	if (cMax == cMin) 
	{   
		//
		// Achromatic case 
		//
		// r = g = b
		//

		m_nSat = 0;
		m_nHue = eUNDEFINED;
	}
	else 
	{     
		//
		// Chromatic Case
		//
		// Saturation
		//

		if (m_nLum <= (eLSMAX / 2))
			m_nSat = (((cMax - cMin) * eLSMAX) + ((cMax + cMin) / 2)) / (cMax + cMin);
		else
			m_nSat = (((cMax - cMin) * eLSMAX) + ((2 * eRGBMAX - cMax - cMin) / 2)) / (2 * eRGBMAX - cMax - cMin);

		//
		// Hue
		//

		short nRDelta = (((cMax - m_nRed) * (eHUEMAX/6)) + ((cMax - cMin) / 2)) / (cMax - cMin);
		short nGDelta = (((cMax - m_nGreen) * (eHUEMAX/6)) + ((cMax - cMin) / 2)) / (cMax - cMin);
		short nBDelta = (((cMax - m_nBlue) * (eHUEMAX/6)) + ((cMax - cMin) / 2)) / (cMax - cMin);

		if (m_nRed == cMax)
			m_nHue = nBDelta - nGDelta;
		else if (m_nGreen == cMax)
			m_nHue = (eHUEMAX / 3) + nRDelta - nBDelta;
		else 
			// m_nBlue == cMax
			m_nHue = ((2 * eHUEMAX) / 3) + nGDelta - nRDelta;

		if (m_nHue < 0)
			m_nHue += eHUEMAX;

		if (m_nHue > eHUEMAX)
			m_nHue -= eHUEMAX;
	}
}

//
// Utility routine for HLStoRGB
//

inline static short HueToRGB(short n1, short n2, short nHue)
{ 
	//
	// range check: note values passed add/subtract thirds of range
	if (nHue < 0)
		nHue += CColor::eHUEMAX;

	if (nHue > CColor::eHUEMAX)
		nHue -= CColor::eHUEMAX;

    //
	// return r, g, or b value from this tridrant
	//

	if (nHue < (CColor::eHUEMAX / 6))
		return (n1 + (((n2 - n1) * nHue + (CColor::eHUEMAX / 12)) / (CColor::eHUEMAX / 6)));

	if (nHue < (CColor::eHUEMAX / 2))
		return n2;

	if (nHue < ((CColor::eHUEMAX * 2) / 3))
		return (n1 + (((n2 - n1) * (((CColor::eHUEMAX * 2) / 3) - nHue) + (CColor::eHUEMAX / 12)) / (CColor::eHUEMAX / 6))); 
	else
		return n1;
} 

void CColor::HLSToRGB()
{
	if (0 == m_nSat) 
	{            
		//
		// Achromatic case
		//

		m_nRed = m_nGreen = m_nBlue = (m_nLum * eRGBMAX) / eLSMAX;
		if (eUNDEFINED != m_nHue) 
		{
			// ERROR
			m_nRed = m_nGreen = m_nBlue = 0;
		}
	}
	else  
	{
		// chromatic case
		// set up magic numbers

		short nMagic2;
		if (m_nLum <= (eLSMAX / 2))
			nMagic2 = (m_nLum * (eLSMAX + m_nSat) + (eLSMAX / 2)) / eLSMAX;
		else
			nMagic2 = m_nLum + m_nSat - ((m_nLum * m_nSat) + (eLSMAX / 2)) / eLSMAX;

		short nMagic1 = 2 * m_nLum - nMagic2;

		// Get RGB, change units from HLSMAX to RGBMAX
		
		m_nRed =   (HueToRGB(nMagic1, nMagic2, m_nHue + (eHUEMAX / 3)) * eRGBMAX + (eHUEMAX / 2)) / eHUEMAX; 
		m_nGreen = (HueToRGB(nMagic1, nMagic2, m_nHue) * eRGBMAX + (eHUEMAX / 2)) / eHUEMAX;
		m_nBlue =  (HueToRGB(nMagic1, nMagic2, m_nHue - (eHUEMAX / 3)) * eRGBMAX + (eHUEMAX / 2)) / eHUEMAX; 
	}
}

//
// CColorDisplay
//

CColorDisplay::CColorDisplay()
{
	BOOL bResult = FWnd::RegisterWindow(DD_WNDCLASS("DDColor"),
										CS_SAVEBITS | CS_DBLCLKS,
										NULL,
										LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)),
										(HICON)NULL);
	assert(bResult);
}

// WindowProc
//

LRESULT CColorDisplay::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(m_hWnd, &ps);
			if (hDC)
			{
				CRect rcBound;
				GetClientRect(rcBound);
				
				UIUtilities::FillSolidRect(hDC, rcBound, m_crColor);
				
				EndPaint(m_hWnd, &ps);
				return 0;
			}
		}
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// CColorSpectrum
//

CColorSpectrum::CColorSpectrum()
{
	m_bCreated = FALSE;
	long nBaseUnit = GetDialogBaseUnits();
	int nLow = LOWORD(nBaseUnit);
	m_nWidth = (103 * nLow) / 4;
	int nHigh = HIWORD(nBaseUnit);
	m_nHeight = (96 * nHigh) / 8;
	Initialize();
}

CColorSpectrum::~CColorSpectrum()
{
}

//
// Initialize
//

void CColorSpectrum::Initialize()
{
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		HDC hDCOff = m_ffDraw.RequestDC(hDC, m_nWidth, m_nHeight);
		CRect rcClient;
		rcClient.right = m_nWidth;
		rcClient.bottom = m_nHeight;
		Draw(hDCOff, rcClient);
		m_bCreated = TRUE;
		CDib theDib;
		theDib.FromBitmap(m_ffDraw.GetBitmap(), hDC);
		HANDLE hFile = CreateFile("c:\\Spec.dib", 
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
		theDib.Write(hFile);
		CloseHandle(hFile);
		ReleaseDC(NULL, hDC);
	}
}

LRESULT CColorSpectrum::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(m_hWnd, &ps);
			if (hDC)
			{
				CRect rcBound;
				GetClientRect(rcBound);
				
				HDC hDCOff = m_ffDraw.GetDC();
				if (hDCOff != hDC)
					m_ffDraw.Paint(hDC, 0, 0, rcBound.Width(), rcBound.Height());
				
				EndPaint(m_hWnd, &ps);
				return 0;
			}
		}
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// Draw
//

void CColorSpectrum::Draw(HDC hDC, CRect& rcBound)
{
	int nBitDepth = GetDeviceCaps(hDC, BITSPIXEL);
	CColor cValue;
	COLORREF crColor;
	int nWidth = rcBound.Width();
	int nHeight = rcBound.Height();
	int nHue;
	int nTemp;
	if (nBitDepth > 8)
	{
		for (int nSat = nHeight; nSat > -1; nSat--)
		{
			for (nHue = 0; nHue < nWidth; nHue++)
			{
				cValue.SetHLS(nHue * CColor::eHUEMAX / nWidth, 120, nSat * CColor::eLSMAX / nHeight);
				nTemp = nHeight - nSat;
				crColor = cValue.Color();
				SetPixel(hDC, nHue, nTemp, crColor);
			}
		}
	}
	else
	{
		CRect rcColor;
		int nSatDelta = nHeight / 4;
		int nHueDelta = nWidth / 4;
		HBRUSH hBrush;
		for (int nSat = nHeight; nSat > -1; nSat -= 4)
		{
			for (nHue = 0; nHue < nWidth; nHue += 4)
			{
				cValue.SetHLS(nHue * CColor::eHUEMAX / nWidth, 120, nSat * CColor::eLSMAX / nHeight);
				nTemp = nHeight - nSat;
				crColor = cValue.Color();
				hBrush = CreateSolidBrush(crColor);
				rcColor.left = nHue;
				rcColor.right = rcColor.left + nHueDelta;
				rcColor.top = nTemp;
				rcColor.bottom = rcColor.top + nSatDelta;
				FillRect(hDC, &rcColor, hBrush);
				DeleteBrush(hBrush);
			}
		}
	}
}

//
// CLuminosity
//

CLuminosity::CLuminosity()
{
}

CLuminosity::~CLuminosity()
{
}

LRESULT CLuminosity::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hDC = BeginPaint(m_hWnd, &ps);
			if (hDC)
			{
				CRect rcBound;
				GetClientRect(rcBound);
				
				HDC hDCOff = m_ffDraw.RequestDC(hDC, rcBound.Width(), rcBound.Height());
				if (NULL == hDCOff)
					hDCOff = hDC;

				Draw(hDCOff, rcBound);
				
				if (hDCOff != hDC)
					m_ffDraw.Paint(hDC, 0, 0);
				
				EndPaint(m_hWnd, &ps);
				return 0;
			}
		}
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// Draw
//

void CLuminosity::Draw(HDC hDC, CRect& rcBound)
{
	int nHeight = rcBound.Height();
	CRect rc = rcBound;
	rc.bottom = rc.top + 3;
	CColor cValue(m_crColor);
	short nHue, nLum, nSat;
	cValue.GetHLS(nHue, nLum, nSat);
	for (nLum = nHeight; nLum > -1; nLum -= 3)
	{
		cValue.SetHLS(nHue, (nLum * 240) / nHeight, nSat);
		UIUtilities::FillSolidRect(hDC, rc, cValue.Color());
		rc.Offset(0, 3);
	}
}

//
// CDefineColor
//

CDefineColor::CDefineColor()
	: FDialog(IDD_DLG_DEFINECOLOR)
{
	m_bLock = FALSE;
}

CDefineColor::~CDefineColor()
{
}

void CDefineColor::GetRGB()
{
	short nRed, nGreen, nBlue;
	TCHAR szBuffer[MAX_PATH];
	GetWindowText(GetDlgItem(m_hWnd, IDC_ED_RED), szBuffer, MAX_PATH);
	nRed = atoi(szBuffer);
	GetWindowText(GetDlgItem(m_hWnd, IDC_ED_GREEN), szBuffer, MAX_PATH);
	nGreen = atoi(szBuffer);
	GetWindowText(GetDlgItem(m_hWnd, IDC_ED_BLUE), szBuffer, MAX_PATH);
	nBlue = atoi(szBuffer);
	m_cColor.SetRGB(nRed, nGreen, nBlue);
}

void CDefineColor::SetRGB()
{
	short nRed, nGreen, nBlue;
	m_cColor.GetRGB(nRed, nGreen, nBlue);
	DDString strText;
	strText.Format(_T("%i"), nRed);
	SetWindowText(GetDlgItem(m_hWnd, IDC_ED_RED), strText);
	strText.Format(_T("%i"), nGreen);
	SetWindowText(GetDlgItem(m_hWnd, IDC_ED_GREEN), strText);
	strText.Format(_T("%i"), nBlue);
	SetWindowText(GetDlgItem(m_hWnd, IDC_ED_BLUE), strText);
}

void CDefineColor::GetHLS()
{
	short nHue, nLum, nSat;
	TCHAR szBuffer[MAX_PATH];
	nHue = atoi(szBuffer);
	GetWindowText(GetDlgItem(m_hWnd, IDC_ED_HUE), szBuffer, MAX_PATH);
	nLum = atoi(szBuffer);
	GetWindowText(GetDlgItem(m_hWnd, IDC_ED_LUM), szBuffer, MAX_PATH);
	nSat = atoi(szBuffer);
	GetWindowText(GetDlgItem(m_hWnd, IDC_ED_SAT), szBuffer, MAX_PATH);
	m_cColor.SetHLS(nHue, nLum, nSat);
}

void CDefineColor::SetHLS()
{
	short nHue, nLum, nSat;
	m_cColor.GetHLS(nHue, nLum, nSat);
	DDString strText;
	strText.Format(_T("%i"), nHue);
	SetWindowText(GetDlgItem(m_hWnd, IDC_ED_HUE), strText);
	strText.Format(_T("%i"), nLum);
	SetWindowText(GetDlgItem(m_hWnd, IDC_ED_LUM), strText);
	strText.Format(_T("%i"), nSat);
	SetWindowText(GetDlgItem(m_hWnd, IDC_ED_SAT), strText);
}

void CDefineColor::Color(COLORREF& crColor)
{
  m_wndColorDisplay.Color(crColor);
//  m_wndColorSpectrum.Color(crColor);
  m_cColor.Color(crColor);
}

BOOL CDefineColor::DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_INITDIALOG:
			{
				HWND hWndColor = GetDlgItem(m_hWnd, IDC_STATIC_COLOR);
				CRect rcColor;
				GetWindowRect(hWndColor, &rcColor);
				DestroyWindow(hWndColor);

				m_bLock = TRUE;
				
				SetRGB();
				SetHLS();
				
				HWND hWndLuminosity = GetDlgItem(m_hWnd, IDC_LUMINOSITY);
				m_wndLum.SubClassAttach(hWndLuminosity);
				m_wndLum.Color(m_cColor.Color());
				
				HWND hWndSpectrum = GetDlgItem(m_hWnd, IDC_COLORSPECTRUM);
				HDC hDC = GetDC(NULL);
				int nBitDepth = GetDeviceCaps(hDC, BITSPIXEL);
				if (nBitDepth <= 8)
				{
					HBITMAP hBitmap = (HBITMAP)LoadImage(g_hInstance, 
													     MAKEINTRESOURCE(IDB_SPECLOW), 
													     IMAGE_BITMAP,
													     0,
													     0,
													     LR_DEFAULTCOLOR|LR_DEFAULTSIZE);
					SendMessage(hWndSpectrum, STM_SETIMAGE, IMAGE_BITMAP, (LRESULT)hBitmap);
				}
				ReleaseDC(NULL, hDC);
				
				ScreenToClient(m_hWnd, rcColor);
				m_wndColorDisplay.CreateEx(WS_EX_CLIENTEDGE, _T(""), WS_CHILD|WS_VISIBLE, rcColor.left, rcColor.top, rcColor.Width(), rcColor.Height(), m_hWnd);
				m_wndColorDisplay.Color(m_cColor.Color());
				
				m_bLock = FALSE;
				
				m_hCursorLuminosity = (HICON)LoadImage(g_hInstance, 
													   MAKEINTRESOURCE(IDC_LUMINOSITY), 
													   IMAGE_CURSOR,
													   4,
													   8,
													   LR_DEFAULTCOLOR);
				assert(m_hCursorLuminosity);
				CRect rcClient;
				GetWindowRect(hWndLuminosity, &rcClient);
				m_ptLuminosity.x = rcClient.left;
				m_ptLuminosity.y = rcClient.top;
				ScreenToClient(m_hWnd, &m_ptLuminosity);
				m_ptLuminosity.x += rcClient.Width();
				short nHue, nLum, nSat;
				m_cColor.GetHLS(nHue, nLum, nSat);
				rcClient.Inflate(-2, -2);
				m_ptLuminosity.y += (rcClient.Height() - ((rcClient.Height() * nLum) / 240));

				
				::SendMessage(GetDlgItem(m_hWnd, IDC_ED_RED), EM_SETLIMITTEXT, 3, 0);
				::SendMessage(GetDlgItem(m_hWnd, IDC_ED_GREEN), EM_SETLIMITTEXT, 3, 0);
				::SendMessage(GetDlgItem(m_hWnd, IDC_ED_BLUE), EM_SETLIMITTEXT, 3, 0);

				::SendMessage(GetDlgItem(m_hWnd, IDC_SPIN_RED), UDM_SETRANGE, 0, MAKELONG(CColor::eRGBMAX, 0));
				::SendMessage(GetDlgItem(m_hWnd, IDC_SPIN_GREEN), UDM_SETRANGE, 0, MAKELONG(CColor::eRGBMAX, 0));
				::SendMessage(GetDlgItem(m_hWnd, IDC_SPIN_BLUE), UDM_SETRANGE, 0, MAKELONG(CColor::eRGBMAX, 0));

				::SendMessage(GetDlgItem(m_hWnd, IDC_ED_HUE), EM_SETLIMITTEXT, 3, 0);
				::SendMessage(GetDlgItem(m_hWnd, IDC_ED_SAT), EM_SETLIMITTEXT, 3, 0);
				::SendMessage(GetDlgItem(m_hWnd, IDC_ED_LUM), EM_SETLIMITTEXT, 3, 0);

				::SendMessage(GetDlgItem(m_hWnd, IDC_SPIN_HUE), UDM_SETRANGE, 0, MAKELONG(CColor::eHUEMAX, 0));
				::SendMessage(GetDlgItem(m_hWnd, IDC_SPIN_SAT), UDM_SETRANGE, 0, MAKELONG(CColor::eLSMAX, 0));
				::SendMessage(GetDlgItem(m_hWnd, IDC_SPIN_LUM), UDM_SETRANGE, 0, MAKELONG(CColor::eLSMAX, 0));
				SetWindowPos(m_hWnd, HWND_TOP, 0,0,0,0, SWP_NOMOVE|SWP_NOSIZE);
			}
			break;

		case WM_LBUTTONDOWN:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				OnLButtonDown(wParam, pt);
			}
			break;

		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcBound;
					GetClientRect(m_hWnd, &rcBound);
					LRESULT lResult = FDialog::DialogProc(nMsg, wParam, lParam);
					Draw(hDC, rcBound);
					EndPaint(m_hWnd, &ps);
					return lResult;
				}
			}
			break;

		case WM_DESTROY:
			{
//				m_wndColorDisplay.UnsubClass();
//				m_wndColorSpectrum.UnsubClass();
				m_wndLum.UnsubClass();
				DestroyCursor(m_hCursorLuminosity);
			}
			break;

		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				switch (LOWORD(wParam))
				{
				case ID_ADDCOLOR:
					OnAddColor();
					break;

				case ID_CLOSE:
					OnClose();
					break;

				case ID_HELP2:
					WinHelp(m_hWnd, HELP_CONTEXT, 0);
					break;
				}
				break;

			case EN_UPDATE:
				{
					HWND hWnd = (HWND)lParam;
					TCHAR szBuffer[MAX_PATH];
					GetWindowText(hWnd, szBuffer, MAX_PATH);
					int nNumber = atoi(szBuffer);
					switch (LOWORD(wParam))
					{
					case IDC_ED_RED:
						if (nNumber > CColor::eRGBMAX)
							SetWindowText(hWnd, _T("255"));
						break;

					case IDC_ED_GREEN:
						if (nNumber > CColor::eRGBMAX)
							SetWindowText(hWnd, _T("255"));
						break;
					
					case IDC_ED_BLUE:
						if (nNumber > CColor::eRGBMAX)
							SetWindowText(hWnd, _T("255"));
						break;

					case IDC_ED_HUE:
						if (nNumber > CColor::eHUEMAX)
							SetWindowText(hWnd, _T("240"));
						break;
					
					case IDC_ED_LUM:
						if (nNumber > CColor::eLSMAX)
							SetWindowText(hWnd, _T("240"));
						break;

					case IDC_ED_SAT:
						if (nNumber > CColor::eLSMAX)
							SetWindowText(hWnd, _T("240"));
						break;
					}
				}
				break;

			case EN_CHANGE:
				switch (LOWORD(wParam))
				{
				case IDC_ED_RED:
				case IDC_ED_GREEN:
				case IDC_ED_BLUE:
					{
						if (!m_bLock)
						{
							m_bLock = TRUE;
							GetRGB();
							SetHLS();
							m_wndColorDisplay.Color(m_cColor.Color());
							m_wndLum.Color(m_cColor.Color());
							m_bLock = FALSE;
						}
					}
					break;

				case IDC_ED_HUE:
				case IDC_ED_LUM:
				case IDC_ED_SAT:
					{
						if (!m_bLock)
						{
							m_bLock = TRUE;
							GetHLS();
							SetRGB();
							m_wndColorDisplay.Color(m_cColor.Color());
							m_wndLum.Color(m_cColor.Color());
							m_bLock = FALSE;
						}
					}
					break;
				}
				break;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return FDialog::DialogProc(nMsg, wParam, lParam);
}

//
// OnLButtonDown
//

void CDefineColor::OnLButtonDown(WPARAM wParam, const POINT& pt)
{
	POINT ptScreen = pt;
	ClientToScreen(m_hWnd, &ptScreen);

	//
	// CColorSpectrum
	//

	short nHue;
	short nSat;
	short nLum;
	m_cColor.GetHLS(nHue, nLum, nSat);

	BOOL bColorChanged = FALSE;
	CRect rcChild;
	GetWindowRect(GetDlgItem(m_hWnd, IDC_COLORSPECTRUM), &rcChild);
	if (PtInRect(&rcChild, ptScreen))
	{
		nHue = (ptScreen.x  - rcChild.left) * CColor::eHUEMAX / rcChild.Width();
		nSat = 240 - ((ptScreen.y - rcChild.top) * 240 / rcChild.Height());
		bColorChanged = TRUE;
	}

	//
	// CLuminosity
	//

	GetWindowRect(GetDlgItem(m_hWnd, IDC_LUMINOSITY), &rcChild);
	if (PtInRect(&rcChild, ptScreen) && !bColorChanged)
	{
		CRect rcFill(m_ptLuminosity.x, 
					 m_ptLuminosity.x + 4, 
					 m_ptLuminosity.y, 
					 m_ptLuminosity.y + 8);
		HDC hDC = GetDC(m_hWnd);
		if (hDC)
		{
			bColorChanged = TRUE;

			UIUtilities::FillSolidRect(hDC, rcFill, GetSysColor(COLOR_3DFACE));
			
			nLum = 240 - ((ptScreen.y - rcChild.top) * 240 / rcChild.Height());
			
			m_ptLuminosity.x = rcChild.left;
			m_ptLuminosity.y = ptScreen.y;
			ScreenToClient(m_hWnd, &m_ptLuminosity);
			m_ptLuminosity.x += rcChild.Width();
			m_ptLuminosity.y -= 3;

			BOOL bResult = DrawIconEx(hDC, 
									  m_ptLuminosity.x, 
									  m_ptLuminosity.y, 
									  m_hCursorLuminosity, 
									  4, 
									  8, 
									  0, 
									  NULL, 
									  DI_NORMAL);
			ReleaseDC(m_hWnd, hDC);
		}
	}

	if (bColorChanged)
	{
		m_bLock = TRUE;
		m_cColor.SetHLS(nHue, nLum, nSat);
		SetRGB();
		SetHLS();
		m_wndColorDisplay.Color(m_cColor.Color());
		m_wndLum.Color(m_cColor.Color());
		m_bLock = FALSE;
	}
}

//
// Draw
//

void CDefineColor::Draw(HDC hDC, CRect& rcBound)
{
	BOOL bResult = DrawIconEx(hDC, 
							  m_ptLuminosity.x, 
							  m_ptLuminosity.y, 
							  m_hCursorLuminosity, 
							  4, 
							  8, 
							  0, 
							  NULL, 
							  DI_NORMAL);
}

//
// OnAddColor
//

void CDefineColor::OnAddColor()
{
	BOOL bResult = EndDialog(m_hWnd, ID_ADDCOLOR);
	assert(bResult);
}

//
// OnClose
//

void CDefineColor::OnClose()
{
	BOOL bResult = EndDialog(m_hWnd, ID_CLOSE);
	assert(bResult);
}

