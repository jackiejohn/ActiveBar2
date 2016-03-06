#ifndef CATEGORYMGR_INCLUDED
#define CATEGORYMGR_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Resource.h"

class CBrowser;
struct CategoryMgr;
interface ITool;
interface ITools;
interface IActiveBar2;

//
// CatEntry
//

struct CatEntry
{
	CatEntry(BSTR bstrCategory = NULL);
	~CatEntry();

	void AddTool(ITool* pTool);
	BOOL DeleteTool(ITool* pTool);
	BSTR Category();
	TypedArray<ITool*>& GetTools();

	typedef void (*PFNITERATETOOLS) (ITool* pTool, void* pData);
	void IterateCategoryEntry(PFNITERATETOOLS pfnIterate, void* pData);
private:
	TypedArray<ITool*> m_faTools;
	BSTR			  m_bstrCategory;
	friend struct CategoryMgr;
};

inline BSTR CatEntry::Category()
{
	return m_bstrCategory;
}

inline TypedArray<ITool*>& CatEntry::GetTools()
{
	return m_faTools;
}

//
// CategoryMgrNotify
//

struct CategoryMgrNotify
{
	enum ToolChange
	{
		eNewTool,
		eToolDeleted,
		eToolNameChanged,
	};

	enum CategoryChange
	{
		eNewCategory,
		eCategoryDeleted,
		eCategoryNameChanged
	};

	enum ManagerChange
	{
		eMgrInit,
	};

	virtual void CategoryChanged(CatEntry* pCatEntry, CategoryChange ccChange) = 0;
	virtual void ToolChanged(BSTR bstrCategory, ITool* pTool, ToolChange tcChange) = 0;
	virtual void ManagerChanged(ManagerChange mcChange) = 0;
};

//
// CategoryMgr
//

struct CategoryMgr
{
	CategoryMgr();
	~CategoryMgr();

	enum Constants
	{
		eImageCategories,
		eImageTools
	};

	enum SpecialToolIds
	{
		eToolIdSysCommand = 0x80004000,
		eToolIdMDIButtons = 0x80004001
	};

	void Init();

	typedef void (*PFNITERATECATEGORYS) (CatEntry* pCatEntry, void* pData);
	void IterateCategories(PFNITERATECATEGORYS pfnIterate, void* pData);

	BOOL IsToolInToolsCollection(ITool* pTool, int* pnIndex = 0);
	ITool* CloneTool(int nInsertIndex, ITool* pTool, BOOL bClone, BSTR bstrCategory);
	void ActiveBar(IActiveBar2* pActiveBar);

	ITools* GetTools();
	short ToolCount();

	ITool* CreateTool(BSTR bstrCategory, IStream* pStream);
	ITool* CreateTool(BSTR bstrCategory, BOOL bDisplayDialog = FALSE);
	BOOL DeleteTool(BSTR bstrCategory, ITool* pTool);

	CatEntry* CreateEntry(BSTR bstrCategory);
	CatEntry* FindEntry(BSTR bstrCategory);
	BOOL DeleteEntry(BSTR bstrCategory);
	void ChangeName(CatEntry* pCatEntry, LPTSTR szCategoryName);
	
	void FillCategoryCombo(HWND hWndCombo);
	void FillCategoryListbox(HWND hWndListbox, BOOL bNone = TRUE);
	void FillToolListbox(HWND hWndListbox, BSTR bstrCategory);

	void Register(CategoryMgrNotify* pCategoryMgrNotify);
	void Unregister(CategoryMgrNotify* pCategoryMgrNotify);

private:
	void Cleanup();

	TypedArray<CategoryMgrNotify*> m_faRegistered;
	TypedArray<CatEntry*>		   m_faCategories;
	IActiveBar2*				   m_pActiveBar;
	DDString					   m_strNone;
	DDString					   m_strNewTool;
	ITools*						   m_pTools;
};

inline ITools* CategoryMgr::GetTools()
{
	return m_pTools;
}

#endif