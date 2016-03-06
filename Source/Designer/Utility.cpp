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
#include "Resource.h"
#include "..\Interfaces.h"
#include "..\EventLog.h"
#include "Utility.h"

extern HINSTANCE g_hInstance;

//
// DDString
//

DDString::DDString(TCHAR* szString)
{
	m_szBuffer = new TCHAR[_tcslen(szString)+1];
	assert(m_szBuffer);
	if (m_szBuffer)
		_tcscpy(m_szBuffer, szString);
}
	
//
// ~DDString
//

DDString::~DDString()
{
	delete [] m_szBuffer;
}

//
// AllocSysString
//

BSTR DDString::AllocSysString() const
{ 
	MAKE_WIDEPTR_FROMTCHAR(wString, m_szBuffer);
	return SysAllocString(wString);
}

//
// LoadStringm_hWndActive
//

BOOL DDString::LoadString (UINT nResourceId)
{
	delete [] m_szBuffer;
	static TCHAR szBuffer[512];
	::LoadString(g_hInstance, nResourceId, szBuffer, 512);
	m_szBuffer = new TCHAR[_tcslen(szBuffer)+1];
	if (NULL == m_szBuffer)
		return FALSE;
	_tcscpy(m_szBuffer, szBuffer);
	return TRUE;
}

#define FORCE_ANSI      0x10000
#define FORCE_UNICODE   0x20000

void DDString::FormatV(LPCTSTR lpszFormat, va_list argList)
{
	va_list argListSave = argList;

	// make a guess at the maximum length of the resulting string
	int nMaxLen = 0;
	for (LPCTSTR lpsz = lpszFormat; *lpsz != '\0'; lpsz = _tcsinc(lpsz))
	{
		// handle '%' character, but watch out for '%%'
		if (*lpsz != '%' || *(lpsz = _tcsinc(lpsz)) == '%')
		{
			nMaxLen += _tclen(lpsz);
			continue;
		}

		int nItemLen = 0;

		// handle '%' character with format
		int nWidth = 0;
		for (; *lpsz != '\0'; lpsz = _tcsinc(lpsz))
		{
			// check for valid flags
			if (*lpsz == '#')
				nMaxLen += 2;   // for '0x'
			else if (*lpsz == '*')
				nWidth = va_arg(argList, int);
			else if (*lpsz == '-' || *lpsz == '+' || *lpsz == '0' ||
				*lpsz == ' ')
				;
			else // hit non-flag character
				break;
		}
		// get width and skip it
		if (nWidth == 0)
		{
			// width indicated by
			nWidth = _ttoi(lpsz);
			for (; *lpsz != '\0' && _istdigit(*lpsz); lpsz = _tcsinc(lpsz))
				;
		}
		assert(nWidth >= 0);

		int nPrecision = 0;
		if (*lpsz == '.')
		{
			// skip past '.' separator (width.precision)
			lpsz = _tcsinc(lpsz);

			// get precision and skip it
			if (*lpsz == '*')
			{
				nPrecision = va_arg(argList, int);
				lpsz = _tcsinc(lpsz);
			}
			else
			{
				nPrecision = _ttoi(lpsz);
				for (; *lpsz != '\0' && _istdigit(*lpsz); lpsz = _tcsinc(lpsz))
					;
			}
			assert(nPrecision >= 0);
		}

		// should be on type modifier or specifier
		int nModifier = 0;
		switch (*lpsz)
		{
		// modifiers that affect size
		case 'h':
			nModifier = FORCE_ANSI;
			lpsz = _tcsinc(lpsz);
			break;
		case 'l':
			nModifier = FORCE_UNICODE;
			lpsz = _tcsinc(lpsz);
			break;

		// modifiers that do not affect size
		case 'F':
		case 'N':
		case 'L':
			lpsz = _tcsinc(lpsz);
			break;
		}

		// now should be on specifier
		switch (*lpsz | nModifier)
		{
		// single characters
		case 'c':
		case 'C':
			nItemLen = 2;
			va_arg(argList, TCHAR);
			break;
		case 'c'|FORCE_ANSI:
		case 'C'|FORCE_ANSI:
			nItemLen = 2;
			va_arg(argList, char);
			break;
		case 'c'|FORCE_UNICODE:
		case 'C'|FORCE_UNICODE:
			nItemLen = 2;
			va_arg(argList, WCHAR);
			break;

		// strings
		case 's':
		{
			LPCTSTR pstrNextArg = va_arg(argList, LPCTSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6;  // "(null)"
			else
			{
			   nItemLen = lstrlen(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
			break;
		}

		case 'S':
		{
#ifndef _UNICODE
			LPWSTR pstrNextArg = va_arg(argList, LPWSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6;  // "(null)"
			else
			{
			   nItemLen = wcslen(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
#else
			LPCSTR pstrNextArg = va_arg(argList, LPCSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6; // "(null)"
			else
			{
			   nItemLen = lstrlenA(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
#endif
			break;
		}

		case 's'|FORCE_ANSI:
		case 'S'|FORCE_ANSI:
		{
			LPCSTR pstrNextArg = va_arg(argList, LPCSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6; // "(null)"
			else
			{
			   nItemLen = lstrlenA(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
			break;
		}

		case 's'|FORCE_UNICODE:
		case 'S'|FORCE_UNICODE:
		{
			LPWSTR pstrNextArg = va_arg(argList, LPWSTR);
			if (pstrNextArg == NULL)
			   nItemLen = 6; // "(null)"
			else
			{
			   nItemLen = wcslen(pstrNextArg);
			   nItemLen = max(1, nItemLen);
			}
			break;
		}
		}

		// adjust nItemLen for strings
		if (nItemLen != 0)
		{
			nItemLen = max(nItemLen, nWidth);
			if (nPrecision != 0)
				nItemLen = min(nItemLen, nPrecision);
		}
		else
		{
			switch (*lpsz)
			{
			// integers
			case 'd':
			case 'i':
			case 'u':
			case 'x':
			case 'X':
			case 'o':
				va_arg(argList, int);
				nItemLen = 32;
				nItemLen = max(nItemLen, nWidth+nPrecision);
				break;

			case 'e':
			case 'f':
			case 'g':
			case 'G':
				va_arg(argList, double);
				nItemLen = 128;
				nItemLen = max(nItemLen, nWidth+nPrecision);
				break;

			case 'p':
				va_arg(argList, void*);
				nItemLen = 32;
				nItemLen = max(nItemLen, nWidth+nPrecision);
				break;

			// no output
			case 'n':
				va_arg(argList, int*);
				break;

			default:
				assert(FALSE);  // unknown formatting option
			}
		}

		// adjust nMaxLen for output nItemLen
		nMaxLen += nItemLen;
	}

	delete [] m_szBuffer;
	m_szBuffer = new TCHAR[nMaxLen+1];
	assert(m_szBuffer);
	if (m_szBuffer)
		_vstprintf(m_szBuffer, lpszFormat, argListSave);
	va_end(argListSave);
}

void DDString::Format(LPCTSTR szFormat, ...)
{
	va_list argList;
	va_start(argList, szFormat);
	FormatV(szFormat, argList);
	va_end(argList);
}

void DDString::Format(UINT nFormatId, ...)
{
	DDString strFormat;
	strFormat.LoadString(nFormatId);

	va_list argList;
	va_start(argList, nFormatId);
	FormatV(strFormat, argList);
	va_end(argList);
}

//
// OleLoadPictureHelper
//
// Open and load file into HGLOBAL, convert it to IStream and call OleLoadPicture
//

IPicture* OleLoadPictureHelper(LPCTSTR szFileName)
{
	HANDLE    hFile = NULL;
	IPicture* pPict = NULL;
	HGLOBAL   hMem = NULL;
	LPBYTE    pData = NULL;
	try 
	{
		hFile = CreateFile(szFileName, 
						   GENERIC_READ,
						   FILE_SHARE_READ,
						   NULL,
						   OPEN_EXISTING,
						   FILE_ATTRIBUTE_NORMAL,
						   NULL);
		if (INVALID_HANDLE_VALUE == hFile)
			return NULL;
		
		DWORD dwFileSizeHigh;
		DWORD dwFileSize = GetFileSize(hFile,&dwFileSizeHigh);
		hMem = GlobalAlloc(GMEM_MOVEABLE, dwFileSize);
		if (NULL == hMem)
			throw NULL;

		pData = (LPBYTE)GlobalLock(hMem);
		if (NULL == pData)
			throw NULL;

		if (ReadFile(hFile, pData, dwFileSize, &dwFileSizeHigh, NULL) == FALSE)
			throw NULL;

		GlobalUnlock(hMem);
		IStream* pStream;
		HRESULT hResult = CreateStreamOnHGlobal(hMem, TRUE, &pStream);
		if (FAILED(hResult))
			throw NULL;
		
		// now we have the stream
		hResult = OleLoadPicture(pStream, 0, FALSE, IID_IPicture, (LPVOID*)&pPict);
		pStream->Release();

		CloseHandle(hFile);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)

		if (pData)
			GlobalUnlock(hMem);
		
		if (hMem)
			GlobalFree(hMem);

		CloseHandle(hFile);
		return NULL;
	}
	return pPict;
}

//
// RemoveAmperstand
//

void RemoveAmperstand(TCHAR* szString)
{
	char* pChar = szString;
	while (*pChar)
	{
		if (*pChar == '&')
		{
			TCHAR* pChar2 = pChar;
			TCHAR* pChar3 = pChar+1;
			while (*pChar3)
				*pChar2++ = *pChar3++;
			*pChar2 = NULL;
		}
		pChar++;
	}
}

void CleanCaption(BSTR bstrName, WCHAR* wCaption)
{
	if (bstrName)
	{
		if (wcslen(bstrName) < 255)
			wcscpy(wCaption, bstrName);
		else
		{
			wcsncpy(wCaption, bstrName, 254);
			wCaption[255] = NULL;
		}

		WCHAR* wChar = wCaption;
		while (*wChar)
		{
			if (L'\t' == *wChar)
			{
				*wChar = NULL;
				break;
			}
			if (L'&' == *wChar)
			{
				WCHAR* pChar = wChar;
				while (*pChar)
					*pChar++ = *(pChar+1);
				break;
			}
			wChar++;
		}
	}
	else
		wCaption[0] = NULL;
}

//
// VisitBarTools
//

BOOL VisitBarTools(IActiveBar2* pActiveBar, long nToolId, void* pData, PFNMODIFYTOOL pToolFunction)
{
	try
	{
		ITools* pTools;
		HRESULT hResult = pActiveBar->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
			return FALSE;

		short nCount;
		hResult = pTools->Count(&nCount);
		if (FAILED(hResult))
		{
			pTools->Release();
			return FALSE;
		}

		VARIANT vTool;
		vTool.vt = VT_I4;

		ITool* pTool;
		long nCurToolId;
		for (vTool.lVal = 0; vTool.lVal < nCount; vTool.lVal++)
		{
			hResult = pTools->Item(&vTool, (Tool**)&pTool);
			if (FAILED(hResult))
				continue;

			hResult = pTool->get_ID(&nCurToolId);
			if (FAILED(hResult))
			{
				pTool->Release();
				continue;
			}

			if (-1 == nToolId || nCurToolId == nToolId)
				(*pToolFunction)(NULL, NULL, pTool, pData);
			pTool->Release();
		}
		pTools->Release();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

//
// VisitBandTools
//
// Look through all bands and tools 
//

BOOL VisitBandTools(IActiveBar2* pActiveBar, long nToolId, void* pData, PFNMODIFYTOOL pToolFunction)
{
	try
	{
		IChildBands* pChildBands;
		IBands* pBands;
		IBand*  pBand;
		IBand*  pChildBand;
		ITools* pTools;
		ITool*  pTool;
		VARIANT vTool, vChildBand, vBand;
		vBand.vt = vChildBand.vt = vTool.vt = VT_I4;
		short nToolCount, nChildBandCount, nBandCount;
		long nCurToolId;

		HRESULT hResult = pActiveBar->get_Bands((Bands**)&pBands);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBands->Count(&nBandCount);
		if (FAILED(hResult))
		{
			pBands->Release();
			return FALSE;
		}

		for (vBand.lVal = 0; vBand.lVal < nBandCount; vBand.lVal++)
		{
			hResult = pBands->Item(&vBand, (Band**)&pBand);
			if (FAILED(hResult))
				continue;

			if (-1 == nToolId)
				(*pToolFunction)(pBand, NULL, NULL, pData);

			hResult = pBand->get_Tools((Tools**)&pTools);
			if (SUCCEEDED(hResult))
			{
				hResult = pTools->Count(&nToolCount);
				if (FAILED(hResult))
				{
					pTools->Release();
					pBand->Release();
					continue;
				}

				for (vTool.lVal = 0; vTool.lVal < nToolCount; vTool.lVal++)
				{
					hResult = pTools->Item(&vTool, (Tool**)&pTool);
					if (FAILED(hResult))
						continue;

					hResult = pTool->get_ID(&nCurToolId);
					if (FAILED(hResult))
					{
						pTool->Release();
						continue;
					}

					if (-1 == nToolId || nCurToolId == nToolId)
						(*pToolFunction)(NULL, NULL, pTool, pData);

					pTool->Release();
				}
				pTools->Release();
			}

			hResult = pBand->get_ChildBands((ChildBands**)&pChildBands);
			if (SUCCEEDED(hResult))
			{
				hResult = pChildBands->Count(&nChildBandCount);
				if (FAILED(hResult))
				{
					pChildBands->Release();
					pBands->Release();
					continue;
				}

				for (vChildBand.lVal = 0; vChildBand.lVal < nChildBandCount; vChildBand.lVal++)
				{
					hResult = pChildBands->Item(&vChildBand, (Band**)&pChildBand);
					if (FAILED(hResult))
						continue;

					if (-1 == nToolId)
						(*pToolFunction)(NULL, pChildBand, NULL, pData);

					hResult = pChildBand->get_Tools((Tools**)&pTools);
					if (FAILED(hResult))
					{
						pChildBand->Release();
						continue;
					}

					hResult = pTools->Count(&nToolCount);
					if (FAILED(hResult))
					{
						pTools->Release();
						pChildBand->Release();
						continue;
					}

					for (vTool.lVal = 0; vTool.lVal < nToolCount; vTool.lVal++)
					{
						hResult = pTools->Item(&vTool, (Tool**)&pTool);
						if (FAILED(hResult))
							continue;

						hResult = pTool->get_ID(&nCurToolId);
						if (FAILED(hResult))
						{
							pTool->Release();
							continue;
						}

						if (-1 == nToolId || nCurToolId == nToolId)
							(*pToolFunction)(NULL, pChildBand, pTool, pData);

						pTool->Release();
					}
					pTools->Release();
					pChildBand->Release();
				}
				pChildBands->Release();
			}
			pBand->Release();
		}
		pBands->Release();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

//
// DoIt
//

void Report::DoIt(IBand* pBand, IBand* pChildBand, ITool* pTool, void* pData)
{
	HRESULT hResult;
	Report* pReport = (Report*)pData;

	try
	{
		if (pBand)
		{
			BSTR bstrName;
			hResult = pBand->get_Name(&bstrName);
			if (SUCCEEDED(hResult))
			{
				MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
				SysFreeString(bstrName);

				_stprintf(pReport->m_szLine, _T("\r\n\tBand: %-40.40s\r\n\r\n"), szName);
				WriteFile(pReport->m_hFile, pReport->m_szLine, _tcslen(pReport->m_szLine), (DWORD*)&pReport->m_dwFilePosition, 0);

				_stprintf(pReport->m_szLine, _T("\t\t%-40.40s %s\t\t%s\r\n"), _T("Tool Name"), _T("Tool Id"), _T("Shortcut"));
				WriteFile(pReport->m_hFile, pReport->m_szLine, _tcslen(pReport->m_szLine), (DWORD*)&pReport->m_dwFilePosition, 0);
			}
		}
		else if (pChildBand && NULL == pTool)
		{
			BSTR bstrName;
			hResult = pChildBand->get_Caption(&bstrName);
			SysFreeString(bstrName);
			if (SUCCEEDED(hResult))
			{
				MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
				_stprintf(pReport->m_szLine, _T("\r\n\t\tChild Band: %-40.40s\r\n"), szName);
				WriteFile(pReport->m_hFile, pReport->m_szLine, _tcslen(pReport->m_szLine), (DWORD*)&pReport->m_dwFilePosition, 0);
				_stprintf(pReport->m_szLine, _T("\t\t\t%-40.40s %s\t\t\r\n"), _T("Tool Name"), _T("Tool Id"), _T("Shortcut"));
				WriteFile(pReport->m_hFile, pReport->m_szLine, _tcslen(pReport->m_szLine), (DWORD*)&pReport->m_dwFilePosition, 0);
			}
		}
		else if (pTool)
		{
			long nToolId;
			BSTR bstrName;
			hResult = pTool->get_Name(&bstrName);
			if (SUCCEEDED(hResult))
			{
				MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
				SysFreeString(bstrName);
				VARIANT vShortCut;
				vShortCut.vt = VT_EMPTY;
				hResult = pTool->get_ShortCuts(&vShortCut);
				if (SUCCEEDED(hResult))
				{
					long nLBound;
					HRESULT hResult = SafeArrayGetLBound(vShortCut.parray, 1,  &nLBound);
					if (FAILED(hResult))
						return;
					
					long nUBound;
					hResult = SafeArrayGetUBound(vShortCut.parray, 1,  &nUBound);
					if (FAILED(hResult))
						return;

					IShortCut** ppShortCut;
					hResult = SafeArrayAccessData(vShortCut.parray, (void**)&ppShortCut);
					if (FAILED(hResult))
						return;

					BSTR bstrDesc = NULL;
					int nLength = 1;
					int nCount = nUBound - nLBound + 1;
					if (nLBound < nCount)
						hResult = ppShortCut[nLBound]->get_Value(&bstrDesc);
					SafeArrayUnaccessData(vShortCut.parray);

					pTool->get_ID(&nToolId);
					if (pChildBand)
						WriteFile(pReport->m_hFile, _T("\t"), _tcslen(_T("\t")), (DWORD*)&pReport->m_dwFilePosition, 0);

					if (SUCCEEDED(hResult))
					{
						MAKE_TCHARPTR_FROMWIDE(szShortCut, bstrDesc);
						_stprintf(pReport->m_szLine, _T("\t\t%-40.40s %li\t\t\t%s\r\n"), szName, nToolId, szShortCut);
						SysFreeString(bstrDesc);
					}
					else
						_stprintf(pReport->m_szLine, _T("\t\t%-40.40s %li\r\n"), szName, nToolId);
					WriteFile(pReport->m_hFile, pReport->m_szLine, _tcslen(pReport->m_szLine), (DWORD*)&pReport->m_dwFilePosition, 0);
					VariantClear(&vShortCut);
				}
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// DoIt
//

void Header::DoIt(IBand* pBand, IBand* pChildBand, ITool* pTool, void* pData)
{
	try
	{
		Header* pHeader = (Header*)pData;
		
		if (pTool)
		{
			HRESULT hResult;
			long nToolId;
			BSTR bstrName;
			hResult = pTool->get_Name(&bstrName);
			if (SUCCEEDED(hResult))
			{
				MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
				SysFreeString(bstrName);
				hResult = pTool->get_ID(&nToolId);
				if (SUCCEEDED(hResult))
				{
					_tcsupr(szName);
					TCHAR* pChar = szName;
					while (*pChar)
					{
						if (' ' == *pChar)
							*pChar = '_';
						pChar++;
					}
					_stprintf(pHeader->m_szLine, _T("#define IDD_%-40.40s %li\r\n"), szName, nToolId);
					WriteFile(pHeader->m_hFile, pHeader->m_szLine, _tcslen(pHeader->m_szLine), (DWORD*)&pHeader->m_dwFilePosition, 0);
				}
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// GenSelect
//

void GenSelect::DoIt(IBand* pBand, IBand* pChildBand, ITool* pTool, void* pData)
{
	GenSelect* pGenSelect = (GenSelect*)pData;
	if (pTool)
	{
		ToolTypes ttType;
		HRESULT hResult = pTool->get_ControlType(&ttType);
		if (ddTTSeparator == ttType)
			return;
		
		BSTR bstrName;
		hResult = pTool->get_SubBand(&bstrName);
		if (!(NULL == bstrName || NULL == *bstrName))
		{
			SysFreeString(bstrName);
			return;
		}
		SysFreeString(bstrName);

		hResult = pTool->get_Name(&bstrName);
		if (SUCCEEDED(hResult))
		{
			MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
			SysFreeString(bstrName);
			DDString strTemp;
			strTemp.Format(_T("%s\tcase \"%s\"\n"), pGenSelect->m_strSelect, szName);
			pGenSelect->m_strSelect = strTemp;
		}
	}
}

//
// CleanName
//

static BOOL CleanName(LPTSTR szName, LPTSTR& szShortCut)
{
	if (szName)
	{
		TCHAR* ch = szName;
		while (*ch)
		{
			if ('\t' == *ch)
			{
				szShortCut = ch+1;
				*ch = NULL;
				break;
			}
			while ('&' == *ch || '.' == *ch || ' ' == *ch)
			{
				TCHAR* chTmp = ch;
				while (*chTmp)
					*chTmp++ = *(chTmp + 1);
			}
			ch++;
		}
	}
	return TRUE;
}

static BOOL CleanCaption(LPTSTR szCaption)
{
	if (szCaption)
	{
		TCHAR* ch = szCaption;
		while (*ch)
		{
			if ('\t' == *ch)
			{
				*ch = NULL;
				break;
			}
			ch++;
		}
	}
	return TRUE;
}

BOOL GrabMenu::m_bCancelled = FALSE;

//
// GrabMenu
//

GrabMenu::GrabMenu(HMENU hMenu, IActiveBar2* pActiveBar)
{
	m_pActiveBar = pActiveBar;
	m_hMenu = hMenu;
}

//
// CreateSubMenu
//

BOOL GrabMenu::CreateSubMenu(IBands* pBands, ITool* pTool, BSTR bstrName, BSTR bstrCaption, HMENU hMenu, int nItem)
{
	MENUITEMINFO menuItem;
	DDString strToolCaption;
	DDString strToolName;
	HRESULT hResult;
	LPTSTR  szShortCut = NULL;
	ITools* pBarTools = NULL;
	ITools* pTools = NULL;
	IBand* pBand = NULL;
	ITool* pSubTool = NULL;
	TCHAR szMenu[MAX_PATH];
	TCHAR szCaption[MAX_PATH];
	HMENU hSubMenu = NULL;
	BOOL bResult;
	BSTR bstrSubName = NULL;
	BSTR bstrSubCaption = NULL;
	long nToolId;
	int nSubItemCount;
	int nSubItem;
	int nResult;
	
	hSubMenu = GetSubMenu(hMenu, nItem);
	if (NULL == hSubMenu)
		return FALSE;

	nSubItemCount = GetMenuItemCount(hSubMenu);
	if (nSubItemCount < 1)
		return FALSE;

	hResult = pBands->Add(bstrName, (Band**)&pBand);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	hResult = pBand->put_Type(ddBTPopup);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	hResult = pBand->put_Caption(bstrCaption);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	hResult = pTool->put_SubBand(bstrName);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	hResult = pBand->get_Tools((Tools**)&pTools);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	for (nSubItem = 0; nSubItem < nSubItemCount; nSubItem++)
	{
		menuItem.cbSize = sizeof(menuItem);
		menuItem.fMask = MIIM_STATE | MIIM_TYPE;
		menuItem.dwTypeData = szMenu;
		menuItem.cch = MAX_PATH;
		try
		{
			nResult = GetMenuItemInfo(hSubMenu, nSubItem, TRUE, &menuItem);
			if (0 == nResult)
				continue;

			if (MFT_SEPARATOR & menuItem.fType)
			{
				strToolCaption.Format(_T("%s"), _T("Separator"));
				strToolName.Format(_T("mi%s"), _T("Separator"));
			}
			else
			{
				_tcsncpy(szCaption, szMenu, MAX_PATH-1);
				if (!CleanCaption(szCaption))
					continue;

				strToolCaption.Format(_T("%s"), szCaption);

				if (!CleanName(szMenu, szShortCut))
					continue;

				strToolName.Format(_T("mi%s"), szMenu);
			}

			//
			// Create a Tool here
			//

			bstrSubName = strToolName.AllocSysString();
			
			nToolId = FindLastToolId(m_pActiveBar) + 1;
			hResult = pTools->Add(nToolId, bstrSubName, (Tool**)&pSubTool);
			if (FAILED(hResult))
				continue;

			if (szShortCut)
			{
				VARIANT vShortCut;
				vShortCut.vt = VT_ARRAY|VT_DISPATCH;
				vShortCut.parray = NULL;

				SAFEARRAYBOUND rgsabound[1];    
				rgsabound[0].lLbound = 0;
				rgsabound[0].cElements = 1;
				vShortCut.parray = SafeArrayCreate(VT_DISPATCH, 1, rgsabound);
				if (vShortCut.parray)
				{
					IShortCut** ppShortCut;

					// Get a pointer to the elements of the array.
					HRESULT hResult = SafeArrayAccessData(vShortCut.parray, (void**)&ppShortCut);
					if (SUCCEEDED(hResult))
					{
						HRESULT hResult = CoCreateInstance(CLSID_ShortCut, 
														   NULL, 
														   CLSCTX_INPROC_SERVER, 
														   IID_IShortCut, 
														   (LPVOID*)ppShortCut);
						if (SUCCEEDED(hResult))
						{
							MAKE_WIDEPTR_FROMTCHAR(wShortCut, szShortCut);
							BSTR bstr = SysAllocString(wShortCut);
							if (bstr)
							{
								hResult = (*ppShortCut)->put_Value(bstr);
								SysFreeString(bstr);
							}
						}
						SafeArrayUnaccessData(vShortCut.parray);

						hResult = pSubTool->put_ShortCuts(vShortCut);
						
						(*ppShortCut)->Release();
					}
				}
			}

			if (MFT_SEPARATOR & menuItem.fType)
				hResult = pSubTool->put_ControlType(ddTTSeparator);
			else
			{
				bstrSubCaption = strToolCaption.AllocSysString();
				hResult = pSubTool->put_Caption(bstrSubCaption);
				if (FAILED(hResult))
					continue;

				if (MFS_CHECKED & menuItem.fState)
					hResult = pSubTool->put_Checked(VARIANT_TRUE);
				else if (MFS_DEFAULT & menuItem.fState)
					hResult = pSubTool->put_Default(VARIANT_TRUE);
				else if (MFS_UNCHECKED & menuItem.fState)
					hResult = pSubTool->put_Checked(VARIANT_FALSE);
			}

			hResult = pSubTool->put_Category(bstrName);

			hResult = m_pActiveBar->get_Tools((Tools**)&pBarTools);
			if (SUCCEEDED(hResult))
			{
				hResult = pBarTools->Insert(-1, (Tool*)pSubTool);
				pBarTools->Release();
			}
			CreateSubMenu(pBands, pSubTool, bstrSubName, bstrSubCaption, hSubMenu, nSubItem);

			pSubTool->Release();
			pSubTool = NULL;
			SysFreeString(bstrSubName);
			bstrSubName = NULL;
			SysFreeString(bstrSubCaption);
			bstrSubCaption = NULL;
			szShortCut = NULL;
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}
Cleanup:

	if (pBand)
		pBand->Release();

	if (pTools)
		pTools->Release();
	
	return TRUE;
}

//
// DoIt
//

BOOL GrabMenu::DoIt()
{
	int nItemCount = GetMenuItemCount(m_hMenu);
	if (nItemCount < 1)
		return FALSE;
	
	//
	// Create a band for the menu bar
	//
	
	IBands* pBands = NULL;
	HRESULT hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
	if (FAILED(hResult))
		return FALSE;

	DDString strToolName;
	DDString strToolCaption;
	LPTSTR  szShortCut = NULL;
	ITools* pBarTools = NULL;
	ITools* pTools = NULL;
	IBand* pBand = NULL;
	ITool* pTool = NULL;
	TCHAR szMenu[MAX_PATH];
	TCHAR szCaption[MAX_PATH];
	BOOL bResult = TRUE;
	BSTR bstrName;
	BSTR bstrCaption;
	long nResult;
	long nToolId;
	int nItem;
	hResult = pBands->Add(L"Menu Bar", (Band**)&pBand);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	hResult = pBand->put_Type(ddBTMenuBar);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	hResult = pBand->get_Tools((Tools**)&pTools);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	for (nItem = 0; nItem < nItemCount; nItem++)
	{
		try
		{
			nResult = GetMenuString(m_hMenu, nItem, szMenu, MAX_PATH, MF_BYPOSITION);
			if (0 == nResult)
				continue;
			
			_tcsncpy(szCaption, szMenu, MAX_PATH-1);
			if (!CleanCaption(szCaption))
				continue;
			
			strToolCaption.Format(_T("%s"), szMenu);

			if (!CleanName(szMenu, szShortCut))
				continue;

			strToolName.Format(_T("mi%s"), szMenu);

			//
			// Create a Tool here
			//

			bstrName = strToolName.AllocSysString();

			nToolId = FindLastToolId(m_pActiveBar) + 1;
			hResult = pTools->Add(nToolId, bstrName, (Tool**)&pTool);
			if (FAILED(hResult))
				continue;

			bstrCaption = strToolCaption.AllocSysString();
			hResult = pTool->put_Caption(bstrCaption);
			if (FAILED(hResult))
				continue;

			hResult = m_pActiveBar->get_Tools((Tools**)&pBarTools);
			if (SUCCEEDED(hResult))
			{
				hResult = pBarTools->Insert(-1, (Tool*)pTool);
				pBarTools->Release();
			}

			CreateSubMenu(pBands, pTool, bstrName, bstrCaption, m_hMenu, nItem);

			SysFreeString(bstrName);
			SysFreeString(bstrCaption);
			pTool->Release();
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}

Cleanup:
	if (pBands)
		pBands->Release();

	if (pBand)
		pBand->Release();

	if (pTools)
		pTools->Release();

	return bResult;
}

//
// SnapShotWndProc
//

LRESULT CALLBACK SnapShotWndProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	return DefWindowProc(hWnd, nMsg, wParam, lParam);
}

HWND GrabMenu::FindMenuWindow(HWND hWndParent)
{
	try
	{
		HWND hWndMenu = NULL;
		HCURSOR hCursorPrev;
		POINT   ptCurrent;
		m_bCancelled = FALSE;


		WNDCLASS wndClass;
		wndClass.style = CS_HREDRAW|CS_VREDRAW;
		wndClass.lpfnWndProc = SnapShotWndProc;
		wndClass.cbClsExtra = 0;
		wndClass.cbWndExtra = 0;
		wndClass.hInstance = g_hInstance;
		wndClass.hIcon = NULL;
		wndClass.hCursor = NULL;
		wndClass.hbrBackground = NULL;
		wndClass.lpszMenuName=NULL;
		wndClass.lpszClassName = _T("DDMenuGrab");
		RegisterClass(&wndClass);

		::ShowWindow(hWndParent, SW_SHOWMINIMIZED);

		HWND hWnd = CreateWindowEx(WS_EX_TRANSPARENT | WS_EX_TOPMOST,
								   wndClass.lpszClassName,
								   _T(""),
								   WS_POPUP,
								   0,
								   0,
								   GetSystemMetrics(SM_CXSCREEN),
								   GetSystemMetrics(SM_CYSCREEN),
								   NULL,
								   NULL,
								   g_hInstance,
								   NULL);

		::ShowWindow(hWnd, SW_SHOWNORMAL);

		hCursorPrev = ::SetCursor(NULL);
		SetCapture(hWnd);
		HCURSOR hCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDC_MENUGRAB));
		assert(hCursor);
		if (hCursor)
			SetCursor(hCursor);

		MSG msg;
		while (GetCapture() == hWnd)
		{
			GetMessage(&msg, NULL, 0, 0);

			switch (msg.message)
			{
			case WM_CANCELMODE:
				m_bCancelled = TRUE;
				goto ExitLoop;
			
			case WM_KEYUP:
				switch (msg.wParam)
				{
				case VK_ESCAPE:
					m_bCancelled = TRUE;
					goto ExitLoop;

				case VK_RETURN:
					GetCursorPos(&ptCurrent);
					goto ExitLoop;
				}
				break;

			case WM_LBUTTONDOWN:
				GetCursorPos(&ptCurrent);
				goto ExitLoop;

			default:	
				DispatchMessage(&msg);
				break;
			}
		}

ExitLoop:
		if (GetCapture() == hWnd)
			ReleaseCapture();

		DestroyWindow(hWnd);
		SetCursor(hCursorPrev);
		HWND hWndTemp = WindowFromPoint(ptCurrent);
		::ShowWindow(hWndParent, SW_RESTORE);
		return hWndTemp;
	}
	CATCH
	{
		assert(FALSE);
		::ShowWindow(hWndParent, SW_RESTORE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return NULL;
}

//
// FindLastToolId
//

long FindLastToolId(IActiveBar2* pActiveBar)
{
	long nMaxToolId = -1;
	try
	{
		ITools* pTools;

		HRESULT hResult = pActiveBar->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
			return -2;

		ITool* pTool;
		long nToolId;
		short nCount;
		VARIANT vIndex;
		vIndex.vt = VT_I2;
		pTools->Count(&nCount);
		for (vIndex.iVal = 0; vIndex.iVal < nCount; vIndex.iVal++)
		{
			hResult = pTools->Item(&vIndex, (Tool**)&pTool);
			if (FAILED(hResult))
				continue;

			pTool->get_ID(&nToolId);
			if (nToolId > nMaxToolId && !(nToolId & eSpecialToolId))
				nMaxToolId = nToolId;
			pTool->Release();
		}
		pTools->Release();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return nMaxToolId;
}

//
// WinHelp
//

BOOL WinHelp(HWND hWndMain, UINT uCommand, DWORD dwData)
{
	DDString strHelpFile;
	strHelpFile.LoadString(IDS_ACTIVEBARHELP);
	TCHAR  szHelpFilePath[_MAX_PATH];
	TCHAR* szEnd = 0;
	
	DWORD dwResult = GetModuleFileName(g_hInstance, szHelpFilePath, _MAX_PATH);
	if (dwResult > 0)
		szEnd = _tcsrchr(szHelpFilePath, static_cast<TCHAR>('\\'));

	if (szEnd)
	{
		*(++szEnd) = NULL;
		_tcscat(szHelpFilePath, strHelpFile);
	}
	else
		_tcscpy(szHelpFilePath, strHelpFile);

	return WinHelp(hWndMain, szHelpFilePath, uCommand, dwData);
}
//
// StWriteBSTR
//

HRESULT StWriteBSTR(IStream* pStream, BSTR bstr)
{
	HRESULT hResult;
	short nSize = 0;
	if (bstr)
	{
		nSize = (wcslen(bstr)+1) * sizeof(WCHAR);
		hResult = pStream->Write(&nSize, sizeof(short), NULL);
		if (FAILED(hResult))
			return hResult;
		hResult = pStream->Write(bstr, nSize, NULL);
		if (FAILED(hResult))
			return hResult;
	}
	else
	{
		hResult = pStream->Write(&nSize, sizeof(short), NULL);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}

//
// StReadBSTR
//

HRESULT StReadBSTR(IStream* pStream, BSTR& bstr)
{
	SysFreeString(bstr);
	bstr = NULL;
	short nSize;
	HRESULT hResult = pStream->Read(&nSize, sizeof(short), NULL);
	if (FAILED(hResult))
		return hResult;

	if (nSize > 0)
	{
		bstr = SysAllocStringByteLen(NULL, nSize);
		if (NULL == bstr)
			return E_OUTOFMEMORY;

		pStream->Read(bstr, nSize, NULL);
		if (FAILED(hResult))
		{
			SysFreeString(bstr);
			bstr = NULL;
			return hResult;
		}
	}
	return NOERROR;
}

