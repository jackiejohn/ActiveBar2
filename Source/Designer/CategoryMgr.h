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

class CCategoryTarget;
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
	enum CreatedTypes
	{
		eDesigner,
		eActiveBar
	};

	CatEntry(BSTR bstrCategory = NULL, CreatedTypes eCreated = eActiveBar);
	~CatEntry();

	BOOL AddTool(ITool* pTool);
	BSTR Category();
	BOOL DeleteTool(ITool* pTool);
	BOOL FindTool(ITool* pToolIn);
	TypedArray<ITool*>& GetTools();

	typedef void (*PFNITERATETOOLS) (ITool* pTool, void* pData);
	void IterateCategoryEntry(PFNITERATETOOLS pfnIterate, void* pData);
private:
	TypedArray<ITool*> m_aTools;
	CreatedTypes       m_eCreated;
	BSTR			   m_bstrCategory;
	friend struct CategoryMgr;
};

inline BSTR CatEntry::Category()
{
	return m_bstrCategory;
}

inline TypedArray<ITool*>& CatEntry::GetTools()
{
	return m_aTools;
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
		eImageCategory,
		eImageTools
	};

	void Init();

	BOOL DoReport(const Report& theReport);

	typedef void (*PFNITERATECATEGORYS) (CatEntry* pCatEntry, void* pData);
	void IterateCategories(PFNITERATECATEGORYS pfnIterate, void* pData);

	void ActiveBar(IActiveBar2* pActiveBar);
	IActiveBar2* ActiveBar();

	ITool* CloneTool(int nInsertIndex, ITool* pTool, BOOL bClone, BSTR bstrCategory);

	BOOL IsToolInToolsCollection(ITool* pTool, int* pnIndex = 0);

	ITools* GetTools();
	short ToolCount();

	int CategoryCount();

	BOOL AddTool(ITool* pTool);
	ITool* CreateTool(BSTR bstrCategory, BOOL bDisplayDialog = FALSE, int ttTool = 0);
	BOOL DeleteTool(BSTR bstrCategory, ITool* pTool);

	CatEntry* CreateEntry(BSTR bstrCategory, CatEntry::CreatedTypes eCreated = CatEntry::eActiveBar);

	CatEntry* FindEntry(BSTR bstrCategory);

	BOOL DeleteEntry(BSTR bstrCategory);

	void ChangeName(CatEntry* pCatEntry, LPTSTR szCategoryName);
	
	void FillCategoryCombo(HWND hWndCombo);
	void FillCategoryListbox(HWND hWndListbox, BOOL bNone = TRUE);
	void FillToolListbox(HWND hWndListbox, BSTR bstrCategory);

	void Register(CategoryMgrNotify* pCategoryMgrNotify);
	void Unregister(CategoryMgrNotify* pCategoryMgrNotify);

	void ClearABCreated();

private:
	void Cleanup();

	TypedArray<CategoryMgrNotify*> m_aRegistered;
	TypedArray<CatEntry*> m_aCategories;
	IActiveBar2* m_pActiveBar;
	DDString m_strNone;
	DDString m_strNewTool;
	ITools* m_pTools;
};

inline int CategoryMgr::CategoryCount()
{
	return m_aCategories.GetSize();
}

inline ITools* CategoryMgr::GetTools()
{
	return m_pTools;
}

inline IActiveBar2* CategoryMgr::ActiveBar()
{
	return m_pActiveBar;
}

#endif