//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <assert.h>
#include "map.h"
#include "debug.h"
#include "Support.h"
#include "GDIUtil.h"
#include "Globals.h"
#include "..\EventLog.h"
#include "FDialog.h"

//
// Extended dialog templates (new in Win95)
//

#pragma pack(push, 1)

typedef struct
{
	WORD dlgVer;
	WORD signature;
	DWORD helpID;
	DWORD exStyle;
	DWORD style;
	WORD cDlgItems;
	short x;
	short y;
	short cx;
	short cy;
} DLGTEMPLATEEX;

typedef struct
{
	DWORD helpID;
	DWORD exStyle;
	DWORD style;
	short x;
	short y;
	short cx;
	short cy;
	DWORD id;
} DLGITEMTEMPLATEEX;

#pragma pack(pop)

extern HINSTANCE g_hInstance;

static FDialog* s_pCreateTarget;
static FDialog* s_pTranslateTarget=NULL;
static HHOOK s_hHookOldMsgFilter;
static FMap	s_mapDialog;

//
// FDialog
//

FDialog::FDialog(UINT nId)
	: m_hTemplate(NULL)
{
	m_bPreTranslate=FALSE;
	m_nDialogId=nId;
#ifdef _DDINC_MODELESS_CODE
	m_bModeless = FALSE;
#endif
}

FDialog::~FDialog()
{
	if (m_hTemplate)
		GlobalFree(m_hTemplate);
}
	
inline static BOOL IsDialogEx(const DLGTEMPLATE* pTemplate)
{
	return ((DLGTEMPLATEEX*)pTemplate)->signature == 0xFFFF;
}

inline static BOOL HasFont(const DLGTEMPLATE* pTemplate)
{
	return (DS_SETFONT &
		(IsDialogEx(pTemplate) ? ((DLGTEMPLATEEX*)pTemplate)->style :
		pTemplate->style));
}

inline static int FontAttrSize(BOOL bDialogEx)
{
	return sizeof(WORD) * (bDialogEx ? 3 : 1);
}

static BYTE* GetFontSizeField(const DLGTEMPLATE* pTemplate)
{
	BOOL bDialogEx = IsDialogEx(pTemplate);
	WORD* pw;

	if (bDialogEx)
		pw = (WORD*)((DLGTEMPLATEEX*)pTemplate + 1);
	else
		pw = (WORD*)(pTemplate + 1);

	if (*pw == (WORD)-1)        // Skip menu name string or ordinal
		pw += 2; // WORDs
	else
		while(*pw++);

	if (*pw == (WORD)-1)        // Skip class name string or ordinal
		pw += 2; // WORDs
	else
		while(*pw++);

	while (*pw++);          // Skip caption string

	return (BYTE*)pw;
}

BOOL FDialog::SetTemplate(const DLGTEMPLATE* pTemplate, UINT cb)
{
	m_dwTemplateSize = cb;
	if ((m_hTemplate = GlobalAlloc(GPTR, m_dwTemplateSize + LF_FACESIZE * 2)) == NULL)
		return FALSE;
	DLGTEMPLATE* pNew = (DLGTEMPLATE*)GlobalLock(m_hTemplate);
	memcpy((BYTE*)pNew, pTemplate, (size_t)m_dwTemplateSize);
	m_bSystemFont = (::HasFont(pNew) == 0);
	ExecuteDlgInit((void*)pTemplate);
	GlobalUnlock(m_hTemplate);
	return TRUE;
}

//
// ExecuteDlgInit
//

BOOL FDialog::ExecuteDlgInit()
{
	// find resource handle
	LPVOID lpResource = NULL;
	HGLOBAL hResource = NULL;
	HRSRC hDlgInit = ::FindResource(g_hInstance, MAKEINTRESOURCE(m_nDialogId), MAKEINTRESOURCE(240));
	if (hDlgInit != NULL)
	{
		// load it
		hResource = LoadResource(g_hInstance, hDlgInit);
		if (hResource == NULL)
			return FALSE;
		// lock it
		lpResource = LockResource(hResource);
		assert(lpResource != NULL);
	}

	// execute it
	BOOL bResult = ExecuteDlgInit(lpResource);

	// cleanup
	if (lpResource != NULL && hResource != NULL)
	{
		UnlockResource(hResource);
		FreeResource(hResource);
	}
	return bResult;
}

//
// ExecuteDlgInit
//

BOOL FDialog::ExecuteDlgInit(LPVOID pResource)
{
	BOOL bSuccess = TRUE;
	if (pResource != NULL)
	{
		UNALIGNED WORD* lpnRes = (WORD*)pResource;
		while (bSuccess && *lpnRes != 0)
		{
			WORD nIDC = *lpnRes++;
			WORD nMsg = *lpnRes++;
			DWORD dwLen = *((UNALIGNED DWORD*&)lpnRes)++;
			// In Win32 the WM_ messages have changed.  They have
			// to be translated from the 32-bit values to 16-bit
			// values here.

			#define WIN16_LB_ADDSTRING  0x0401
			#define WIN16_CB_ADDSTRING  0x0403

			if (nMsg == WIN16_LB_ADDSTRING)
				nMsg = LB_ADDSTRING;
			else if (nMsg == WIN16_CB_ADDSTRING)
				nMsg = CB_ADDSTRING;

			// check for invalid/unknown message types
			assert(nMsg == LB_ADDSTRING || nMsg == CB_ADDSTRING);

#ifdef _DEBUG
			// For AddStrings, the count must exactly delimit the
			// string, including the NULL termination.  This check
			// will not catch all mal-formed ADDSTRINGs, but will
			// catch some.
			if (nMsg == LB_ADDSTRING || nMsg == CB_ADDSTRING)
				assert(*((LPBYTE)lpnRes + (UINT)dwLen - 1) == 0);
#endif

			if (nMsg == LB_ADDSTRING || nMsg == CB_ADDSTRING)
			{
				// List/Combobox returns -1 for error
				if (::SendDlgItemMessageA(m_hWnd, nIDC, nMsg, 0, (LONG)lpnRes) == -1)
					bSuccess = FALSE;
			}

			// skip past data
			lpnRes = (WORD*)((LPBYTE)lpnRes + (UINT)dwLen);
		}
	}
	return bSuccess;
}

BOOL FDialog::IsSysFont()
{
	DLGTEMPLATE* pTemplate = (DLGTEMPLATE*)GlobalLock(m_hTemplate);
	BOOL hasFont = ::HasFont(pTemplate);
	GlobalUnlock(m_hTemplate);
	if (!hasFont)
		return FALSE;

	BOOL retval=FALSE;

	BYTE* pb = GetFontSizeField(pTemplate);
	pb += FontAttrSize(IsDialogEx(pTemplate));

	#if defined(_UNICODE) || (defined(OLE2ANSI) && !defined(_MAC))
		if (lstrcmp(pb,"MS Sans Serif")==0 || lstrcmp(pb,"Helv")==0)
			retval=TRUE;
	#else
		// Convert Unicode font name to MBCS
		if (wcscmp((WCHAR *)pb,L"MS Sans Serif")==0 || wcscmp((WCHAR *)pb,L"Helv")==0)
			retval=TRUE;
	#endif
	return retval;
}

BOOL FDialog::SetFont(LPCTSTR lpFaceName, WORD nFontSize)
{
	if (m_dwTemplateSize == 0)
		return FALSE;

	DLGTEMPLATE* pTemplate = (DLGTEMPLATE*)GlobalLock(m_hTemplate);

	BOOL bDialogEx = IsDialogEx(pTemplate);
	BOOL bHasFont = ::HasFont(pTemplate);
	int cbFontAttr = FontAttrSize(bDialogEx);

	if (bDialogEx)
		((DLGTEMPLATEEX*)pTemplate)->style |= DS_SETFONT;
	else
		pTemplate->style |= DS_SETFONT;

#ifdef _UNICODE
	int cbNew = cbFontAttr + ((lstrlen(lpFaceName) + 1) * sizeof(TCHAR));
	BYTE* pbNew = (BYTE*)lpFaceName;
#else
	WCHAR wszFaceName [LF_FACESIZE];
	int cbNew = cbFontAttr + 2 * MultiByteToWideChar(CP_ACP, 0, lpFaceName, -1, wszFaceName, LF_FACESIZE);
	BYTE* pbNew = (BYTE*)wszFaceName;
#endif

	BYTE* pb = GetFontSizeField(pTemplate);
	int cbOld = bHasFont ? cbFontAttr + 2 * (wcslen((WCHAR*)(pb + cbFontAttr)) + 1) : 0;

	BYTE* pOldControls = (BYTE*)(((DWORD)pb + cbOld + 3) & ~3);
	BYTE* pNewControls = (BYTE*)(((DWORD)pb + cbNew + 3) & ~3);

	BYTE nCtrl = bDialogEx ? (BYTE)((DLGTEMPLATEEX*)pTemplate)->cDlgItems :
		(BYTE)pTemplate->cdit;

	if (cbNew != cbOld && nCtrl > 0)
		memmove(pNewControls, pOldControls, (size_t)(m_dwTemplateSize - (pOldControls - (BYTE*)pTemplate)));

	*(WORD*)pb = nFontSize;
	memmove(pb + cbFontAttr, pbNew, cbNew - cbFontAttr);

	m_dwTemplateSize += (pNewControls - pOldControls);

	GlobalUnlock(m_hTemplate);
	m_bSystemFont = FALSE;
	return TRUE;
}


BOOL CALLBACK FDialog::FDialogProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	FDialog* pDlg;
	if (s_mapDialog.Lookup(hWnd, (void*&)pDlg) == NULL)
	{ 
		if (message == WM_INITDIALOG)
		{
			// Creating a new window so enter into window map
			pDlg = s_pCreateTarget;
			pDlg->m_hWnd = hWnd;
			s_mapDialog.SetAt(hWnd, (void*&)pDlg);
			pDlg->ExecuteDlgInit();
		}
#ifdef _DEBUG
	else 
	{
		char str[80];
		wsprintf(str,"Dialog windowMap error %X, using DefDlgProc\n",message);
		TRACE(1, str);
		return -1;
	}
#else
	else
		return -1;
#endif		
	}
	return pDlg->DialogProc(message,wParam,lParam);
}

int FDialog::DoModal(HWND hWndParent)
{
	int retval = 0;
	try
	{
		s_pCreateTarget = this;
		if (m_bPreTranslate)
			SetPreTranslateHook();
		try
		{
			retval = DialogBox(g_hInstance, MAKEINTRESOURCE(m_nDialogId), hWndParent, (DLGPROC)FDialogProc);
#ifdef _DEBUG
			if (-1 == retval)
				::MessageBox(NULL, GetLastErrorMsg(), _T("Error"), MB_OK);
#endif
		}
		catch (...)
		{
			assert(FALSE);
			return -1;
		}
		if (m_bPreTranslate)
			ClearPreTranslateHook();
	}
	catch (...)
	{
		assert(FALSE);
		return -1;
	}
	return retval;
}

#ifdef _DDINC_MODELESS_CODE
void FDialog::DoModeless(HWND hWndParent)
{
	m_bModeless = TRUE;
	s_pCreateTarget=this;
	CreateDialog(g_hInstance,MAKEINTRESOURCE(m_nDialogId),hWndParent,(DLGPROC)FDialogProc);
	if (m_bPreTranslate)
		SetPreTranslateHook();
}
#endif

BOOL FDialog::DialogProc(UINT message,WPARAM wParam,LPARAM lParam)
{
	try
	{
		switch(message)
		{
		case WM_INITDIALOG:
			return TRUE;
		case WM_NCDESTROY:
			s_mapDialog.RemoveKey(m_hWnd);
#ifdef _DDINC_MODELESS_CODE
			if (m_bModeless)
			{
				if (m_bPreTranslate)
					ClearPreTranslateHook();
				delete this;
			}
#endif
			return 0;

		case WM_COMMAND:
			switch(LOWORD(wParam))
			{
			case IDOK:
				OnOK();
				return TRUE;
			case IDCANCEL:
				OnCancel();
				return TRUE;
			}
		default:
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return FALSE;
}

void FDialog::OnOK()
{
	EndDialog(m_hWnd,IDOK);
}

void FDialog::OnCancel()
{
	EndDialog(m_hWnd,IDCANCEL);
}

void FDialog::CenterDialog(HWND hWndRelativeWindow)
{
	CRect rcBase, rcMy;
	int nScrCX=GetSystemMetrics(SM_CXSCREEN);
	int nScrCY=GetSystemMetrics(SM_CYSCREEN);
	
	if (NULL == hWndRelativeWindow)
	{
		rcBase.right = nScrCX;
		rcBase.bottom = nScrCY;
	}
	else
		GetWindowRect(hWndRelativeWindow, &rcBase);

	GetWindowRect(m_hWnd, &rcMy);
	
	int nLeft = rcBase.left + (rcBase.Width()-rcMy.Width())/2;
	int nTop = rcBase.top + (rcBase.Height()-rcMy.Height())/2;
	int nRight = nLeft + rcMy.Width();
	int nBottom = nTop + rcMy.Height();
	
	if (nRight > nScrCX)
		nLeft = nScrCX - rcMy.Width();
	
	if (nBottom >= nScrCY)
		nTop = nScrCY - rcMy.Height();

	if (nLeft < 0) 
		nLeft = 0;
	
	if (nTop < 0)
		nTop = 0;
	
	MoveWindow(m_hWnd, nLeft, nTop, rcMy.Width(), rcMy.Height(), FALSE);
}

LRESULT CALLBACK FDialogMsgFilterHook(int code, WPARAM wParam, LPARAM lParam)
{
	if ((code < 0 && code != MSGF_DDEMGR) || s_pTranslateTarget == NULL)
	{
		return ::CallNextHookEx(s_hHookOldMsgFilter,code, wParam, lParam);
	}
	return (LRESULT)s_pTranslateTarget->PreTranslateMessage((LPMSG)lParam);
}

void FDialog::SetPreTranslateHook()
{
	s_pTranslateTarget=this;
	s_hHookOldMsgFilter=::SetWindowsHookEx(WH_MSGFILTER,(HOOKPROC)FDialogMsgFilterHook, NULL, ::GetCurrentThreadId());
}

void FDialog::ClearPreTranslateHook()
{
	::UnhookWindowsHookEx(s_hHookOldMsgFilter);
	s_pTranslateTarget = NULL;
}

//
// CFileDialog
//

CFileDialog::CFileDialog(UINT     nTitleId,
						 BOOL     bOpenFileDialog,
						 LPCTSTR  szDefExt,
						 LPCTSTR  szFileName, 
						 DWORD    dwFlags, 
						 LPTSTR   szFilter, 
						 HWND     hWndParent,
						 UINT     nResourceId,
						 LPOFNHOOKPROC pFn)
{
	m_bOpenFileDialog = bOpenFileDialog;

	memset(&m_ofn, 0, sizeof(OPENFILENAME));
	
	m_ofn.lStructSize = sizeof(OPENFILENAME);
	m_ofn.hwndOwner = hWndParent;
	m_szInitFileName[0] = NULL;
	m_ofn.lpstrFile = m_szInitFileName;
	m_ofn.nMaxFile = MAX_PATH;
	m_ofn.lpstrFileTitle = m_szFileTitle;
	m_ofn.nMaxFileTitle = MAX_PATH;
	m_ofn.Flags = dwFlags;
	m_ofn.Flags |= OFN_SHOWHELP;
	m_ofn.lpstrDefExt = szDefExt;

	if (nResourceId)
	{
		m_ofn.hInstance = g_hInstance;
		m_ofn.lpTemplateName = MAKEINTRESOURCE(nResourceId);
		m_ofn.Flags |= OFN_ENABLETEMPLATE | OFN_EXPLORER;
	}
	
	if (nTitleId)
	{
		m_strTitle.LoadString(nTitleId);
		m_ofn.lpstrTitle = m_strTitle;
	}
	
	if (pFn)
	{
		m_ofn.lpfnHook = pFn;
		m_ofn.Flags |= OFN_ENABLEHOOK;
	}

	//
	// Translate filter into commdlg format (lots of \0)
	//
	
	if (szFilter)
	{
		LPTSTR pChar = szFilter;
		
		int nLen = _tcslen(szFilter);

		// convert '|' to '\0'
		while ((pChar = _tcschr(pChar, '|')))
			*pChar++ = '\0';

		m_ofn.lpstrFilter = szFilter;
		*(szFilter+nLen) = NULL;
	}
}

CFileDialog::~CFileDialog()
{
}

UINT CFileDialog::DoModal()
{
	HWND hWndFocus = GetFocus();
	
	UINT nResult;
	if (m_bOpenFileDialog)
		nResult = GetOpenFileName(&m_ofn);
	else
		nResult = GetSaveFileName(&m_ofn);

	if (0 == nResult)
	{
		nResult = CommDlgExtendedError();
	}

	if (::IsWindow(hWndFocus))
		::SetFocus(hWndFocus);
	return nResult;
}

//
// CFontDialog
//

CFontDialog::CFontDialog(HFONT hFont)
{
	GetObject(hFont, sizeof(LOGFONT), &m_lf);
	memset(&m_theChooseFont, 0, sizeof(m_theChooseFont));
	m_theChooseFont.lStructSize = sizeof(m_theChooseFont);
    m_theChooseFont.hwndOwner;
	m_theChooseFont.hDC;     
	m_theChooseFont.lpLogFont = &m_lf; 
    m_theChooseFont.iPointSize = 0;
	m_theChooseFont.Flags = CF_SCREENFONTS|CF_INITTOLOGFONTSTRUCT|CF_EFFECTS;
    m_theChooseFont.rgbColors = 0;
	m_theChooseFont.lCustData = 0; 
    m_theChooseFont.lpfnHook = NULL;
	m_theChooseFont.lpTemplateName = NULL; 
    m_theChooseFont.hInstance = NULL;
	m_theChooseFont.lpszStyle = NULL; 
    m_theChooseFont.nFontType = 0;
    m_theChooseFont.nSizeMin = 0;
	m_theChooseFont.nSizeMax = 0;
}

HFONT CFontDialog::GetFont()
{
	return CreateFontIndirect(&m_lf);
}

//
// CColorDialog
//

CColorDialog::CColorDialog(COLORREF crColor)
{

	memset(&m_theChooseColor, 0, sizeof(m_theChooseColor));
	m_theChooseColor.lStructSize = sizeof(m_theChooseColor);
	m_theChooseColor.Flags = CC_ANYCOLOR|CC_RGBINIT|CC_FULLOPEN;
	m_theChooseColor.lpCustColors = m_crUserDefined;
	m_theChooseColor.rgbResult = crColor;
}

COLORREF& CColorDialog::GetColor()
{
	return m_theChooseColor.rgbResult;
}
