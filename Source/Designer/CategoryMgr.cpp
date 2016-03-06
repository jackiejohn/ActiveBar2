//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "..\Interfaces.h"
#include "Globals.h"
#include "Resource.h"
#include "Support.h"
#include "DragDrop.h"
#include "Dialogs.h"
#include "Browser.h"
#include "StringUtil.h"
#include "CategoryMgr.h"

extern HINSTANCE g_hInstance;

//
// CatEntry
// 

CatEntry::CatEntry(BSTR bstrCategory, CreatedTypes eCreated)
	: m_bstrCategory(NULL)
{
	m_eCreated = eCreated;
	if (bstrCategory)
	{
		m_bstrCategory = SysAllocString(bstrCategory);
		assert(m_bstrCategory);
	}
}

CatEntry::~CatEntry()
{
	SysFreeString (m_bstrCategory);
	m_bstrCategory = NULL;
	
	int nCount = m_aTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		m_aTools.GetAt(nIndex)->Release();

	m_aTools.RemoveAll();
}

//
// DeleteTool
//

BOOL CatEntry::DeleteTool(ITool* pTool)
{
	int nCount = m_aTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (m_aTools.GetAt(nIndex) == pTool)
		{
			m_aTools.RemoveAt(nIndex);
			pTool->Release();
			return TRUE;
		}
	}
	return FALSE;
}

//
// FindTool
//

BOOL CatEntry::FindTool(ITool* pToolIn)
{
	int nCount = m_aTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (m_aTools.GetAt(nIndex) == pToolIn)
			return TRUE;
	}
	return FALSE;
}

//
// AddTool
//

BOOL CatEntry::AddTool(ITool* pTool)
{
	HRESULT hResult = pTool->put_Category(m_bstrCategory);
	if (FAILED(hResult))
		return FALSE;
	pTool->AddRef();
	m_aTools.Add(pTool);
	return TRUE;
}

//
// IterateCategoryEntry
//

void CatEntry::IterateCategoryEntry(PFNITERATETOOLS pfnIterate, void* pData)
{
	int nCount = m_aTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		(*pfnIterate)(m_aTools.GetAt(nIndex), pData);
}

//
// CategoryMgr
//

CategoryMgr::CategoryMgr()
	: m_pTools(NULL),
	  m_pActiveBar(NULL)
{
	m_strNone.LoadString(IDS_CAT_NONE);
	m_strNewTool.LoadString(IDS_NEW_TOOL);
}

CategoryMgr::~CategoryMgr()
{
	Cleanup();
	if (m_pActiveBar)
		m_pActiveBar->Release();

	if (m_pTools)
		m_pTools->Release();
}

//
// ToolCount
//

short CategoryMgr::ToolCount()
{
	short nToolCount = 0;
	if (m_pTools)
		m_pTools->Count(&nToolCount);
	return nToolCount;
}

//
// Cleanup
//

void CategoryMgr::Cleanup()
{
	int nCount = m_aCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		delete m_aCategories.GetAt(nIndex);
	m_aCategories.RemoveAll();
}

// 
// ClearABCreated
//

void CategoryMgr::ClearABCreated()
{
	CatEntry* pEntry;
	int nCount = m_aCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pEntry = m_aCategories.GetAt(nIndex);
		if (CatEntry::eActiveBar == pEntry->m_eCreated)
		{
			delete pEntry;
			m_aCategories.RemoveAt(nIndex);
			nCount--;
			nIndex--;
		}
	}
}

//
// IsToolInToolsCollection
//

BOOL CategoryMgr::IsToolInToolsCollection(ITool* pTool, int* pnIndex)
{
	if (pnIndex)
		*pnIndex = -1;
	
	BOOL bResult = FALSE;
	short nCount;
	HRESULT hResult = m_pTools->Count(&nCount);
	if (SUCCEEDED(hResult))
	{
		ITool* pCmpTool;
		VARIANT vIndex;
		vIndex.vt = VT_I2;
		for (vIndex.iVal = 0; vIndex.iVal < nCount; vIndex.iVal++)
		{
			m_pTools->Item(&vIndex, (Tool**)&pCmpTool);
			if (pCmpTool == pTool)
			{
				if (pnIndex)
					*pnIndex = vIndex.iVal;
				pCmpTool->Release();
				bResult = TRUE;
				break;
			}
			pCmpTool->Release();
		}
	}
	return bResult;
}

//
// CloneTool
//

ITool* CategoryMgr::CloneTool(int nInsertIndex, ITool* pTool, BOOL bClone, BSTR bstrCategory)
{
	// Cloning a Tool
	HRESULT hResult = m_pTools->InsertTool(nInsertIndex, pTool, !bClone);
	if (FAILED(hResult))
		return NULL;

	ITool* pClonedTool = NULL;
	VARIANT vIndex;
	vIndex.vt = VT_I2;
	vIndex.iVal = nInsertIndex;
	hResult = m_pTools->Item(&vIndex, (Tool**)&pClonedTool);
	if (FAILED(hResult) || NULL == pClonedTool)
		return NULL;

	// Getting it back out to change it's category
	hResult = pClonedTool->put_Category(bstrCategory);

	CatEntry* pCatEntry = FindEntry(bstrCategory);
	if (pCatEntry)
		pCatEntry->AddTool(pClonedTool);

	return pClonedTool;
}

//
// FindEntry
//

CatEntry* CategoryMgr::FindEntry(BSTR bstrCategory)
{
	// Stick it into our structure
	CatEntry* pCatEntry;
	int nCount = m_aCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pCatEntry = m_aCategories.GetAt(nIndex);
		if (NULL == pCatEntry)
			continue;

		if ((NULL == pCatEntry->m_bstrCategory && NULL == bstrCategory) ||
			(NULL == pCatEntry->m_bstrCategory && bstrCategory && NULL == *bstrCategory) ||
			(pCatEntry->m_bstrCategory && NULL == *pCatEntry->m_bstrCategory && NULL == bstrCategory) ||
			(pCatEntry->m_bstrCategory && bstrCategory && 0 == wcscmp(pCatEntry->m_bstrCategory, bstrCategory)))
		{
			return pCatEntry;
		}
	}
	return NULL;
}

//
// ChangeName
//

void CategoryMgr::ChangeName(CatEntry* pCatEntry, LPTSTR szCategoryName)
{
	SysFreeString(pCatEntry->m_bstrCategory);
	pCatEntry->m_bstrCategory = NULL;

	MAKE_WIDEPTR_FROMTCHAR(wCategoryName, szCategoryName);
	pCatEntry->m_bstrCategory = SysAllocString(wCategoryName);

	ITool* pTool;
	int nCount = pCatEntry->m_aTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pTool = (ITool*)pCatEntry->m_aTools.GetAt(nIndex);
		pTool->put_Category(pCatEntry->m_bstrCategory);
	}

	int nNotifyCount = m_aRegistered.GetSize();
	for (nIndex = 0; nIndex < nNotifyCount; nIndex++)
		m_aRegistered.GetAt(nIndex)->CategoryChanged(pCatEntry, CategoryMgrNotify::eCategoryNameChanged);
}

//
// CreateEntry
//

CatEntry* CategoryMgr::CreateEntry(BSTR bstrCategory, CatEntry::CreatedTypes eCreated)
{
	CatEntry* pCatEntry = FindEntry(bstrCategory);
	if (NULL == pCatEntry)
	{
		pCatEntry = new CatEntry(bstrCategory, eCreated);
		if (NULL == pCatEntry)
			return NULL;

		m_aCategories.Add(pCatEntry);

		int nNotifyCount = m_aRegistered.GetSize();
		for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
			m_aRegistered.GetAt(nIndex)->CategoryChanged(pCatEntry, CategoryMgrNotify::eNewCategory);
	}
	return pCatEntry;
}

//
// DeleteEntry
//

BOOL CategoryMgr::DeleteEntry(BSTR bstrCategory)
{
	CatEntry* pCatEntry;
	int nCount = m_aCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pCatEntry = (CatEntry*)m_aCategories.GetAt(nIndex);

		if (NULL == pCatEntry)
			continue;

		if ((NULL == pCatEntry->m_bstrCategory && NULL == bstrCategory) ||
			(NULL == pCatEntry->m_bstrCategory && bstrCategory && NULL == *bstrCategory) ||
			(pCatEntry->m_bstrCategory && NULL == *pCatEntry->m_bstrCategory && NULL == bstrCategory) ||
			(pCatEntry->m_bstrCategory && bstrCategory && 0 == wcscmp(pCatEntry->m_bstrCategory, bstrCategory)))
		{

			if (NULL == pCatEntry->m_bstrCategory || NULL == *pCatEntry->m_bstrCategory)
				return FALSE;
			
			m_aCategories.RemoveAt(nIndex);

			int nCount = pCatEntry->m_aTools.GetSize();
			for (nIndex = 0; nIndex < nCount; nIndex++)
				m_pTools->DeleteTool((ITool*)pCatEntry->m_aTools.GetAt(nIndex));

			int nNotifyCount = m_aRegistered.GetSize();
			for (nIndex = 0; nIndex < nNotifyCount; nIndex++)
			{
				m_aRegistered.GetAt(nIndex)->CategoryChanged(pCatEntry, 
															  CategoryMgrNotify::eCategoryDeleted);
			}
			
			delete pCatEntry;
			return TRUE;
		}
	}
	return FALSE;
}

//
// AddTool
//

BOOL CategoryMgr::AddTool(ITool* pTool)
{
	BSTR bstrCategory;
	HRESULT hResult = pTool->get_Category(&bstrCategory);
	if (FAILED(hResult))
		return FALSE;

	CatEntry* pCatEntry = FindEntry(bstrCategory);
	if (NULL == pCatEntry)
		pCatEntry = new CatEntry(bstrCategory);

	SysFreeString(bstrCategory);
	if (NULL == pCatEntry)
		return FALSE;
	pCatEntry->AddTool(pTool);
	return TRUE;
}

//
// CreateTool
//

ITool* CategoryMgr::CreateTool(BSTR bstrCategory, BOOL bDisplayDialog, int ttTool)
{
	CatEntry* pCatEntry = FindEntry(bstrCategory);
	if (NULL == pCatEntry)
		return NULL;

	MAKE_WIDEPTR_FROMTCHAR(wCaption, m_strNewTool);

	// Now check greatest NewTool%d you can find.in m_pTools and increment by 1
	long    nMaxToolId = 0;
	long    nLargestId = 0;
	long    nToolId;
	BSTR    bstrCaption;
	size_t  sCaptionLen = wcslen(wCaption);
	ITool*  pCmpTool;
	short   nElem;
	VARIANT vIndex;
	HRESULT hResult;
	vIndex.vt = VT_I2;
	hResult = m_pTools->Count(&nElem);
	for (vIndex.iVal = 0; vIndex.iVal < nElem; vIndex.iVal++)
	{
		hResult = m_pTools->Item(&vIndex, (Tool**)&pCmpTool);
		if (FAILED(hResult))
			continue;

		hResult = pCmpTool->get_ID(&nToolId);
		if (FAILED(hResult))
		{
			SysFreeString(bstrCaption);
			bstrCaption = NULL;
			continue;
		}

		if (nToolId > nMaxToolId && !(nToolId & eSpecialToolId))
			nMaxToolId = nToolId;

		hResult = pCmpTool->get_Caption(&bstrCaption);
		if (FAILED(hResult))
			continue;

		if (NULL == bstrCaption)
		{
			SysFreeString(bstrCaption);
			bstrCaption = NULL;
			continue;
		}

		if (wcslen(bstrCaption) >= sCaptionLen)
		{
			if (0 == memcmp(bstrCaption, wCaption, sizeof(WCHAR)*sCaptionLen))
			{
				// yes match. So check value
				int nNewVal = _wtoi(bstrCaption+sCaptionLen);
				if (nNewVal > nLargestId)
					nLargestId = nNewVal;
			}
		}
		SysFreeString(bstrCaption);
		bstrCaption = NULL;
		pCmpTool->Release();
	}
	DDString strNewTool;
	strNewTool.Format(_T("%s%d"), m_strNewTool, ++nMaxToolId);
	MAKE_WIDEPTR_FROMTCHAR(wBuffer, strNewTool);

	BSTR bstrName = SysAllocString(wBuffer);
	if (NULL == bstrName)
		return NULL;

	ITool* pTool;
	hResult = m_pTools->Add(nMaxToolId, bstrName, (Tool**)&pTool);
	if (FAILED(hResult))
	{
		SysFreeString(bstrName);
		bstrName = NULL;
		return NULL;
	}
	hResult = pTool->put_Caption(bstrName);
	SysFreeString(bstrName);
	bstrName = NULL;

	hResult = pTool->put_Category(bstrCategory);

	if (ddTTButton != ttTool)
		pTool->put_ControlType((ToolTypes)ttTool);

	pCatEntry->AddTool(pTool);

	int nNotifyCount = m_aRegistered.GetSize();
	for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
		m_aRegistered.GetAt(nIndex)->ToolChanged(bstrCategory, 
												  pTool, 
												  CategoryMgrNotify::eNewTool);

	pTool->Release();
	return pTool;
}

//
// DeleteTool
//

BOOL CategoryMgr::DeleteTool(BSTR bstrCategory, ITool* pTool)
{
	BOOL bResult = FALSE;
	if (SUCCEEDED(m_pTools->DeleteTool(pTool)))
	{
		bResult = TRUE;
		CatEntry* pCatEntry = FindEntry(bstrCategory);
		if (pCatEntry)
			pCatEntry->DeleteTool(pTool);

		int nNotifyCount = m_aRegistered.GetSize();
		for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
			m_aRegistered.GetAt(nIndex)->ToolChanged(bstrCategory, pTool, CategoryMgrNotify::eToolDeleted);
	}
	return bResult;
}

//
// Init
//

void CategoryMgr::Init()
{
	short nCount;
	HRESULT hResult = m_pTools->Count(&nCount);
	if (FAILED(hResult))
		return;

	//
	// Always going to have Category None
	//

	CatEntry* pCatEntry = new CatEntry(NULL);
	assert (pCatEntry);
	if (pCatEntry)
		m_aCategories.Add(pCatEntry);

	VARIANT vTool;
	ITool*  pTool;
	long    nToolId;
	BSTR    bstrCatName;

	vTool.vt = VT_I2;
	for (vTool.iVal = 0; vTool.iVal < nCount; vTool.iVal++)
	{
		hResult = m_pTools->Item(&vTool,(Tool**)&pTool);
		if (FAILED(hResult))
			continue;

		hResult = pTool->get_ID(&nToolId);
		if (FAILED(hResult))
		{
			pTool->Release();
			continue;
		}

		hResult = pTool->get_Category(&bstrCatName);
		if (FAILED(hResult))
		{
			pTool->Release();
			continue;
		}
		
		CatEntry* pCatEntry = FindEntry(bstrCatName);
		if (NULL == pCatEntry)
		{
			pCatEntry = new CatEntry(bstrCatName);
			assert (pCatEntry);
			if (pCatEntry)
				m_aCategories.Add(pCatEntry);
		}

		if (pCatEntry && !pCatEntry->FindTool(pTool))
			pCatEntry->AddTool(pTool);

		pTool->Release();

		SysFreeString(bstrCatName);
		bstrCatName = NULL;
	}
	
	int nNotifyCount = m_aRegistered.GetSize();
	for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
		m_aRegistered.GetAt(nIndex)->ManagerChanged(CategoryMgrNotify::eMgrInit);
}

//
// IterateCategoryImpl
//

static void IterateCategoryReport (ITool* pTool, void* pData)
{
	Report* pReport = (Report*)pData;

	long nToolId;
	BSTR bstrName;
	HRESULT hResult = pTool->get_Name(&bstrName);
	if (SUCCEEDED(hResult))
	{
		MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
		pTool->get_ID(&nToolId);
		_stprintf(pReport->m_szLine, _T("\t\t%-40.40s %li\r\n"), szName, nToolId);
		WriteFile(pReport->m_hFile, pReport->m_szLine, _tcslen(pReport->m_szLine), (DWORD*)&pReport->m_dwFilePosition, 0);
		SysFreeString(bstrName);
	}
}

//
// IterateCategoriesReport
//

static void IterateCategoriesReport (CatEntry* pEntry, void* pData)
{
	Report* pReport = (Report*)pData;
	
	MAKE_TCHARPTR_FROMWIDE(szCategory, pEntry->Category());
	_stprintf(pReport->m_szLine, _T("\tCategory: %s\r\n\r\n"), szCategory);
	WriteFile(pReport->m_hFile, pReport->m_szLine, _tcslen(pReport->m_szLine), (DWORD*)&pReport->m_dwFilePosition, 0);

	_stprintf(pReport->m_szLine, _T("\t\t%-40.40s %s\r\n"), _T("Tool Name"), _T("Tool Id"));
	WriteFile(pReport->m_hFile, pReport->m_szLine, _tcslen(pReport->m_szLine), (DWORD*)&pReport->m_dwFilePosition, 0);

	pEntry->IterateCategoryEntry(IterateCategoryReport, pData);

	WriteFile(pReport->m_hFile, _T("\r\n"), _tcslen(_T("\r\n")), (DWORD*)&pReport->m_dwFilePosition, 0);
}

//
// DoReport
//

BOOL CategoryMgr::DoReport(const Report& theReport)
{
	_stprintf(const_cast<TCHAR*>(theReport.m_szLine), _T("\r\nCategories\r\n"));
	WriteFile(theReport.m_hFile, theReport.m_szLine, _tcslen(theReport.m_szLine), (DWORD*)&theReport.m_dwFilePosition, 0);

	IterateCategories(IterateCategoriesReport, (void*)&theReport);
	return TRUE;
}

//
// FillCategoryListbox
//

void CategoryMgr::FillCategoryListbox(HWND hWndListbox, BOOL bNone)
{
	SendMessage(hWndListbox, WM_SETREDRAW, FALSE, 0);	
	SendMessage(hWndListbox, LB_RESETCONTENT, 0, 0);

	CatEntry* pCatEntry;
	int nCount = m_aCategories.GetSize();
	int nIndex2;
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pCatEntry = m_aCategories.GetAt(nIndex);
		if (NULL == pCatEntry || NULL == pCatEntry->m_bstrCategory || NULL == *pCatEntry->m_bstrCategory)
			continue;

		MAKE_TCHARPTR_FROMWIDE(szItem, pCatEntry->m_bstrCategory);
		nIndex2 = SendMessage(hWndListbox, LB_ADDSTRING, 0, (LPARAM)szItem);
		SendMessage(hWndListbox, LB_SETITEMDATA, (WPARAM)nIndex2, (LPARAM)pCatEntry);
	}
	if (nCount > 0)
		SendMessage(hWndListbox, LB_SETCURSEL, 0, 0);

	SendMessage(hWndListbox, WM_SETREDRAW, TRUE, 0);	
}

//
// FillCategoryCombo
//

void CategoryMgr::FillCategoryCombo(HWND hWndCombo)
{
	SendMessage(hWndCombo, WM_SETREDRAW, FALSE, 0);	
	SendMessage(hWndCombo, CB_RESETCONTENT, 0, 0);

	CatEntry* pCatEntry;
	int nCount = m_aCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pCatEntry = m_aCategories.GetAt(nIndex);
		if (NULL == pCatEntry || NULL == pCatEntry->m_bstrCategory || NULL == *pCatEntry->m_bstrCategory)
			continue;

		MAKE_TCHARPTR_FROMWIDE(szItem, pCatEntry->m_bstrCategory);
		SendMessage(hWndCombo, CB_ADDSTRING, 0, (LPARAM)szItem);
	}
	if (nCount > 0)
		SendMessage(hWndCombo, CB_SETCURSEL, 0, 0);
	SendMessage(hWndCombo, WM_SETREDRAW, TRUE, 0);	
}

//
// FillToolListbox
//

void CategoryMgr::FillToolListbox(HWND hWndListbox, BSTR bstrCategory)
{
	SendMessage(hWndListbox, WM_SETREDRAW, FALSE, 0);	
	SendMessage(hWndListbox, LB_RESETCONTENT, 0, 0);

	CatEntry* pCatEntry = FindEntry(bstrCategory);
	if (pCatEntry)
	{
		int nCount = pCatEntry->m_aTools.GetSize();
		for (int nTool = 0; nTool < nCount; nTool++)
			SendMessage(hWndListbox, LB_ADDSTRING, 0, (LPARAM)pCatEntry->m_aTools.GetAt(nTool));

		if (nCount > 0)
			SendMessage(hWndListbox, LB_SETCURSEL, 0, 0);
	}
	SendMessage(hWndListbox, WM_SETREDRAW, TRUE, 0);	
}

//
// ActiveBar
//

void CategoryMgr::ActiveBar(IActiveBar2* pActiveBar)
{
	if (m_pTools)
		m_pTools->Release();

	if (m_pActiveBar)
		m_pActiveBar->Release();
	
	m_pActiveBar = pActiveBar;
	m_pActiveBar->AddRef();

	m_pActiveBar->get_Tools((Tools**)&m_pTools);
}

//
// Register
//

void CategoryMgr::Register(CategoryMgrNotify* pCategoryMgrNotify)
{
	int nCount = m_aRegistered.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (pCategoryMgrNotify == m_aRegistered.GetAt(nIndex))
			return;
	}
	m_aRegistered.Add(pCategoryMgrNotify);
}

//
// Unregister
//

void CategoryMgr::Unregister(CategoryMgrNotify* pCategoryMgrNotify)
{
	int nCount = m_aRegistered.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (pCategoryMgrNotify == m_aRegistered.GetAt(nIndex))
		{
			m_aRegistered.RemoveAt(nIndex);
			return;
		}
	}
}

//
// IterateCategories
//

void CategoryMgr::IterateCategories(PFNITERATECATEGORYS pfnIterate, void* pData)
{
	int nCount = m_aCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		(*pfnIterate)(m_aCategories.GetAt(nIndex), pData);
}
