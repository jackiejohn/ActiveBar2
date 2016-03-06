//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "interfaces.h"
#include "Globals.h"
#include "Resource.h"
#include "Support.h"
#include "CategoryMgr.h"

extern HINSTANCE g_hInstance;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// CatEntry
// 

CatEntry::CatEntry(BSTR bstrCategory)
	: m_bstrCategory(NULL)
{
	if (bstrCategory)
	{
		m_bstrCategory = SysAllocString(bstrCategory);
		assert(m_bstrCategory);
	}
}

CatEntry::~CatEntry()
{
	SysFreeString (m_bstrCategory);

	int nCount = m_faTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		m_faTools.GetAt(nIndex)->Release();

	m_faTools.RemoveAll();
}

BOOL CatEntry::DeleteTool(ITool* pTool)
{
	int nCount = m_faTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (m_faTools.GetAt(nIndex) == pTool)
		{
			m_faTools.RemoveAt(nIndex);
			pTool->Release();
			return TRUE;
		}
	}
	return FALSE;
}

void CatEntry::AddTool(ITool* pTool)
{
	pTool->AddRef();
	HRESULT hResult = m_faTools.Add(pTool);
	if (FAILED(hResult))
	{
		pTool->Release();
		assert(FALSE);
	}
}

//
// IterateCategoryEntry
//

void CatEntry::IterateCategoryEntry(PFNITERATETOOLS pfnIterate, void* pData)
{
	int nCount = m_faTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		(*pfnIterate)(m_faTools.GetAt(nIndex), pData);
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
	if (m_pTools)
		m_pTools->Release();
	if (m_pActiveBar)
		m_pActiveBar->Release();
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
	int nCount = m_faCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		delete m_faCategories.GetAt(nIndex);
	m_faCategories.RemoveAll();
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
	m_pTools->InsertTool(nInsertIndex, pTool, !bClone);

	ITool* pClonedTool = NULL;
	VARIANT vIndex;
	vIndex.vt = VT_I2;
	vIndex.iVal = nInsertIndex;

	HRESULT hResult = m_pTools->Item(&vIndex, (Tool**)&pClonedTool);
	if (FAILED(hResult) || NULL == pClonedTool)
		return NULL;

	// Getting it back out to change it's category
	pClonedTool->put_Category(bstrCategory);

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
	int nCount = m_faCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pCatEntry = m_faCategories.GetAt(nIndex);
		if (NULL == pCatEntry)
			continue;

		if (pCatEntry->m_bstrCategory && NULL == bstrCategory)
			continue;

		if (NULL == pCatEntry->m_bstrCategory && bstrCategory)
			continue;

		if (pCatEntry->m_bstrCategory == bstrCategory)
			return pCatEntry;

		if (*pCatEntry->m_bstrCategory == (WCHAR)bstrCategory)
			return pCatEntry;

		if (0 == wcscmp(pCatEntry->m_bstrCategory, bstrCategory))
			return pCatEntry;
	}
	return NULL;
}

//
// ChangeName
//

void CategoryMgr::ChangeName(CatEntry* pCatEntry, LPTSTR szCategoryName)
{
	int nIndex = 0;
	SysFreeString(pCatEntry->m_bstrCategory);
	MAKE_WIDEPTR_FROMTCHAR(wCategoryName, szCategoryName);
	pCatEntry->m_bstrCategory = SysAllocString(wCategoryName);

	ITool* pTool;
	int nCount = pCatEntry->m_faTools.GetSize();
	for (nIndex = 0; nIndex < nCount; nIndex++)
	{
		pTool = (ITool*)pCatEntry->m_faTools.GetAt(nIndex);
		pTool->put_Category(pCatEntry->m_bstrCategory);
	}

	int nNotifyCount = m_faRegistered.GetSize();
	for (nIndex = 0; nIndex < nNotifyCount; nIndex++)
		m_faRegistered.GetAt(nIndex)->CategoryChanged(pCatEntry, CategoryMgrNotify::eCategoryNameChanged);
}

//
// CreateEntry
//

CatEntry* CategoryMgr::CreateEntry(BSTR bstrCategory)
{
	CatEntry* pCatEntry = FindEntry(bstrCategory);
	if (NULL == pCatEntry)
	{
		pCatEntry = new CatEntry(bstrCategory);
		if (NULL == pCatEntry)
			return NULL;

		HRESULT hResult = m_faCategories.Add(pCatEntry);
		if (FAILED(hResult))
		{
			delete pCatEntry;
			assert(FALSE);
			return NULL;
		}

		int nNotifyCount = m_faRegistered.GetSize();
		for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
			m_faRegistered.GetAt(nIndex)->CategoryChanged(pCatEntry, CategoryMgrNotify::eNewCategory);
	}
	return pCatEntry;
}

//
// DeleteEntry
//

BOOL CategoryMgr::DeleteEntry(BSTR bstrCategory)
{
	CatEntry* pCatEntry;
	int nCount = m_faCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pCatEntry = m_faCategories.GetAt(nIndex);
		if (pCatEntry && 0 == wcscmp(pCatEntry->m_bstrCategory, bstrCategory))
		{
			m_faCategories.RemoveAt(nIndex);

			int nCount = pCatEntry->m_faTools.GetSize();
			for (nIndex = 0; nIndex < nCount; nIndex++)
				m_pTools->DeleteTool(pCatEntry->m_faTools.GetAt(nIndex));

			int nNotifyCount = m_faRegistered.GetSize();
			for (nIndex = 0; nIndex < nNotifyCount; nIndex++)
				m_faRegistered.GetAt(nIndex)->CategoryChanged(pCatEntry, CategoryMgrNotify::eCategoryDeleted);
			
			delete pCatEntry;
			return TRUE;
		}
	}
	return FALSE;
}

//
// CreateTool
//

ITool* CategoryMgr::CreateTool(BSTR bstrCategory, IStream* pStream)
{
	ITool* pTool;
	m_pTools->CreateTool(&pTool);
	if (NULL == pTool)
		return FALSE;

	m_pTools->InsertTool(-1, pTool, VARIANT_FALSE);

	CatEntry* pCatEntry = FindEntry(bstrCategory);
	if (NULL != pCatEntry)
		pCatEntry->AddTool(pTool);

	pTool->DragDropExchange(pStream, VARIANT_FALSE);
	pTool->put_Category(bstrCategory);

	int nNotifyCount = m_faRegistered.GetSize();
	for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
		m_faRegistered.GetAt(nIndex)->ToolChanged(bstrCategory, pTool, CategoryMgrNotify::eNewTool);

	pTool->Release();
	return pTool;
}

//
// CreateTool
//

ITool* CategoryMgr::CreateTool(BSTR bstrCategory, BOOL bDisplayDialog)
{
	ITool* pTool;
	m_pTools->CreateTool(&pTool);
	if (NULL == pTool)
		return FALSE;

	m_pTools->InsertTool(-1, pTool, VARIANT_FALSE);

	CatEntry* pCatEntry = FindEntry(bstrCategory);
	if (NULL != pCatEntry)
		pCatEntry->AddTool(pTool);

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
	m_pTools->Count(&nElem);
	for (vIndex.iVal = 0; vIndex.iVal < nElem; vIndex.iVal++)
	{
		hResult = m_pTools->Item(&vIndex, (Tool**)&pCmpTool);
		if (!FAILED(hResult))
		{
			pCmpTool->get_Caption(&bstrCaption);
			pCmpTool->get_ID(&nToolId);
			if (nToolId > nMaxToolId)
				nMaxToolId = nToolId;

			if (bstrCaption)
			{
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
			}
			pCmpTool->Release();
		}
	}
	DDString strNewTool;
	strNewTool.Format(_T("%s%d"), m_strNewTool, ++nMaxToolId);
	MAKE_WIDEPTR_FROMTCHAR(wBuffer, strNewTool);

	BSTR bstrName = SysAllocString(wBuffer);
	pTool->put_Name(bstrName);
	pTool->put_ID(nMaxToolId);
	pTool->put_Category(bstrCategory);

	int nNotifyCount = m_faRegistered.GetSize();
	for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
		m_faRegistered.GetAt(nIndex)->ToolChanged(bstrCategory, pTool, CategoryMgrNotify::eNewTool);

	pTool->Release();
	return pTool;
}

//
// DeleteTool
//

BOOL CategoryMgr::DeleteTool(BSTR bstrCategory, ITool* pTool)
{
	BOOL bResult = FALSE;
	if (S_OK == m_pTools->DeleteTool(pTool))
	{
		bResult = TRUE;
		CatEntry* pCatEntry = FindEntry(bstrCategory);
		if (pCatEntry)
			pCatEntry->DeleteTool(pTool);

		int nNotifyCount = m_faRegistered.GetSize();
		for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
			m_faRegistered.GetAt(nIndex)->ToolChanged(bstrCategory, pTool, CategoryMgrNotify::eToolDeleted);
	}
	return bResult;
}

//
// Init
//

void CategoryMgr::Init()
{
	Cleanup();

	short nCount;
	HRESULT hResult = m_pTools->Count(&nCount);
	if (FAILED(hResult))
		return;

	VARIANT vTool;
	ITool*  pTool;
	BSTR    bstrCatName;

	vTool.vt = VT_I2;
	for (vTool.iVal = 0; vTool.iVal < nCount; vTool.iVal++)
	{
		hResult = m_pTools->Item(&vTool,(Tool**)&pTool);
		if (FAILED(hResult))
			continue;

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
			{
				hResult = m_faCategories.Add(pCatEntry);
				if (FAILED(hResult))
				{
					assert(FALSE);
					pTool->Release();
					SysFreeString(bstrCatName);
					delete pCatEntry;
					continue;
				}
			}
		}

		pCatEntry->AddTool(pTool);
		pTool->Release();
		SysFreeString(bstrCatName);
	}
	
	int nNotifyCount = m_faRegistered.GetSize();
	for (int nIndex = 0; nIndex < nNotifyCount; nIndex++)
		m_faRegistered.GetAt(nIndex)->ManagerChanged(CategoryMgrNotify::eMgrInit);
}

//
// FillCategoryListbox
//

void CategoryMgr::FillCategoryListbox(HWND hWndListbox, BOOL bNone)
{
	SendMessage(hWndListbox, WM_SETREDRAW, FALSE, 0);	
	SendMessage(hWndListbox, LB_RESETCONTENT, 0, 0);

	CatEntry* pCatEntry;
	int nCount = m_faCategories.GetSize();
	int nIndex2;
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pCatEntry = m_faCategories.GetAt(nIndex);
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
	int nCount = m_faCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pCatEntry = m_faCategories.GetAt(nIndex);
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
		int nCount = pCatEntry->m_faTools.GetSize();
		for (int nTool = 0; nTool < nCount; nTool++)
			SendMessage(hWndListbox, LB_ADDSTRING, 0, (LPARAM)pCatEntry->m_faTools.GetAt(nTool));

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
	{
		m_pTools->Release();
		m_pTools = NULL;
	}

	if (m_pActiveBar)
	{
		m_pActiveBar->Release();
		m_pActiveBar = NULL;
	}
	
	m_pActiveBar = pActiveBar;
	if (m_pActiveBar)
	{
		m_pActiveBar->AddRef();
		HRESULT hResult = m_pActiveBar->get_Tools((Tools**)&m_pTools);
		assert(SUCCEEDED(hResult));
	}
}

//
// Register
//

void CategoryMgr::Register(CategoryMgrNotify* pCategoryMgrNotify)
{
	int nCount = m_faRegistered.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (pCategoryMgrNotify == m_faRegistered.GetAt(nIndex))
			return;
	}
	HRESULT hResult = m_faRegistered.Add(pCategoryMgrNotify);
	assert(SUCCEEDED(hResult));
}

//
// Unregister
//

void CategoryMgr::Unregister(CategoryMgrNotify* pCategoryMgrNotify)
{
	int nCount = m_faRegistered.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (pCategoryMgrNotify == m_faRegistered.GetAt(nIndex))
		{
			m_faRegistered.RemoveAt(nIndex);
			return;
		}
	}
}

//
// IterateCategories
//

void CategoryMgr::IterateCategories(PFNITERATECATEGORYS pfnIterate, void* pData)
{
	int nCount = m_faCategories.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		(*pfnIterate)(m_faCategories.GetAt(nIndex), pData);
}
