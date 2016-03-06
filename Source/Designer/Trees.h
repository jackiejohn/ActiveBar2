#ifndef DESIGNERTREES
#define DESIGNERTREES

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Resource.h"
#include "CategoryMgr.h"
#include "..\Interfaces.h"
#include "..\PrivateInterfaces.h"

class CToolDataObject;

//
// CTreeCtrl
//

class CTreeCtrl : public FWnd
{
public:
	CTreeCtrl();
	virtual ~CTreeCtrl();

	enum 
	{
		eTextLength = 50,
		eEditLabelTimer = 101,
		eEditLabelTimerInterval = 1000
	};

	IActiveBar2* ActiveBar();
	
	void Expand(HTREEITEM hItem, BOOL bExpand = TRUE);
	int FindIndex(HTREEITEM hParent, HTREEITEM hChild);
	int ChildCount(HTREEITEM hParent);
	
	HTREEITEM& GetCurrentItem();
	TV_ITEM&   GetCurrentTreeViewItem();
	BOOL       GetParent(HTREEITEM hItem, TV_ITEM& tvParent, DWORD dwMask);
	BOOL       GetSelected(UINT nMask);
	
	BOOL       HitTestAtCursor(POINT& pt, UINT nMask);
	BOOL       HitTest(POINT& pt, UINT nMask);
	
	BOOL SelectItems(HTREEITEM hFromItem, HTREEITEM hToItem);
	UINT GetItemState(HTREEITEM hItem, UINT nState) const;
	BOOL SetItemState(HTREEITEM hItem, UINT nState, UINT nStateMask);
	LPARAM GetItemData(HTREEITEM hItem);
	void ClearSelection();
	void SelectMultiple(HTREEITEM hClickedItem, UINT nFlags);

	UINT GetSelectedCount() const;
	HTREEITEM GetFirstSelectedItem();
	HTREEITEM GetNextSelectedItem(HTREEITEM hItem);
	HTREEITEM GetPrevSelectedItem(HTREEITEM hItem);
	HTREEITEM HitTest(POINT pt, UINT* pFlags) const;
	HTREEITEM InsertNode (HTREEITEM hParent, LPTSTR szText, int nImage, int nChild = -2, LPARAM lParam = -2, HTREEITEM hAfter = TVI_LAST);

	void SelectDropTarget(HTREEITEM hItem);
	void Refresh(HTREEITEM hParent);

protected:

	BOOL LoadImages(UINT nResourceId);
	BOOL PopupMenu (UINT nId, int nSubMenuPos, POINT& pt);
	void DeleteChildren(HTREEITEM hParent);

	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual void OnCommand(WPARAM wParam, LPARAM lParam) {};
	virtual LRESULT OnNotify(NMHDR* pNMHdr) {return TRUE;};

	LRESULT OnLButtonDown(WPARAM wParam, POINT& pt);
	LRESULT OnLButtonUp(WPARAM wParam, POINT& pt);
	LRESULT OnMouseMove(UINT nFlags, POINT& pt);
	LRESULT OnRButtonDown(WPARAM wParam, POINT& pt);
	void OnTimer(WPARAM wParam);

	TypedArray<HTREEITEM> m_aSelectedItems;
	CFlickerFree m_ff;
	IActiveBar2* m_pActiveBar;
	HIMAGELIST   m_ilImages;
	HTREEITEM    m_hCurrentItem;
	HTREEITEM    m_hFirstSelectedItem;
	TV_ITEM      m_tvCurrentItem;
	POINT	     m_ptClick;
	TCHAR        m_szCurrentText[eTextLength];
	BOOL		 m_bMultSelectActive;
};

inline void CTreeCtrl::Expand(HTREEITEM hItem, BOOL bExpand)
{
	UINT nFlag = TVE_COLLAPSE;
	if (bExpand)
		nFlag = TVE_EXPAND;
	TreeView_Expand(m_hWnd, hItem, nFlag);
}

inline BOOL CTreeCtrl::SetItemState(HTREEITEM hItem, UINT nState, UINT nStateMask)
{ 
	assert(::IsWindow(m_hWnd)); 
	TVITEM tvItem;
	tvItem.hItem = hItem;
	tvItem.mask = TVIF_STATE;
	tvItem.state = nState;
	tvItem.stateMask = nStateMask;
	return TreeView_SetItem(m_hWnd, &tvItem); 
}

inline LPARAM CTreeCtrl::GetItemData(HTREEITEM hItem)
{
	TVITEM tvItem;
	tvItem.hItem = hItem;
	tvItem.mask = TVIF_PARAM;
	if (!TreeView_GetItem(m_hWnd, &tvItem))
		return -1;
	return tvItem.lParam;
}

inline UINT CTreeCtrl::GetItemState(HTREEITEM hItem, UINT nStateMask) const
{ 
	assert(::IsWindow(m_hWnd));
	TVITEM item;
	item.hItem = hItem;
	item.mask = TVIF_STATE;
	item.stateMask = nStateMask;
	item.state = 0;
	::SendMessage(m_hWnd, TVM_GETITEM, 0, (LPARAM)&item);
	return item.state;
}

inline IActiveBar2* CTreeCtrl::ActiveBar()
{
	return m_pActiveBar;
}

inline HTREEITEM& CTreeCtrl::GetCurrentItem()
{
	return m_hCurrentItem;
}

inline TV_ITEM& CTreeCtrl::GetCurrentTreeViewItem()
{
	return m_tvCurrentItem;
}

class CDesignerTarget;
class CBrowser;

//
// ObjectUpdateNotify
//

struct IObjectChanged
{
	virtual void ObjectChanged(int nType = -1) = 0;
};

//
// CDesignerTree
//

class CDesignerTree : public CTreeCtrl, public IObjectChanged, public IDesignerNotifications, public CategoryMgrNotify 
{
public:
	enum LayoutTypes
	{
		ActiveBar2 = 1,
		ActiveBar1 = 2
	};

	enum ImageTypes
	{
		eRoot,
		eBands,
		eBand,
		eMenu,
		eChildMenu,
		ePopup,
		eStatusBar,
		eChildBands,
		eChildBand,
		eButton,
		eDropDown,
		eTextbox,
		eCombobox,
		eLabel,
		eSeparator,
		eWindowList,
		eCustomTool,
		eForm,
		eCategories,
		eCategory
	};

	enum MenuTypes
	{
		eMenuRoot,
		eMenuBands,
		eMenuBandNone,
		eMenuBandChildBands,
		eMenuChildBands,
		eMenuChildBand,
		eMenuTool,
		eMenuCategories,
		eMenuCategory
	};

	enum DesignerTreeType
	{
		eBrowser, 
		eLibrary
	};

	CDesignerTree(DesignerTreeType eTreeType = eBrowser);
	~CDesignerTree();

	BOOL Init(IActiveBar2* pActivebar, CBrowser* pBrowser = NULL, CategoryMgr* pCategoryMgr = NULL);
	void OnNewLayout();
	void OnOpenLayout();
	void OnSaveLayout();
	void OnSaveAsLayout(BOOL bSaveFileName = FALSE);

	STDMETHOD (BandChanged)(LPDISPATCH pBand, LPDISPATCH pPage, BandChangeTypes nType);
	STDMETHOD (ToolSelected)(BSTR bstrBand, BSTR bstrPage, long nToolId);
	virtual void ObjectChanged(int nType = -1);

	LPTSTR CurFileName();
	CategoryMgr* GetCategoryMgr();

	void SettingsChanged();

	void Update();
	CToolDataObject* m_pDataObject;
private:
	void SetRoot();
	void DeleteTool(TVITEM tvParent, TVITEM tvTool, ITool* pTool);
	BOOL GetListOfSelectedTools(TypedArray<ITool*>& aTools);

	HTREEITEM FindCategoryNode(BSTR bstrCategory);

	// CategoryMgrNotify
	virtual void CategoryChanged(CatEntry* pCatEntry, CategoryMgrNotify::CategoryChange ccChange);
	virtual void ToolChanged(BSTR bstrCategory, ITool* pTool, CategoryMgrNotify::ToolChange tcChange);
	virtual void ManagerChanged(CategoryMgrNotify::ManagerChange mcChange);

	HTREEITEM FindBand(LPCTSTR szBandName);
	HTREEITEM FindChildBand(HTREEITEM hBand, LPCTSTR szChildBand);
	HTREEITEM FindTool(HTREEITEM hBand, long nToolIdIn);
	
	void FillCategories(HTREEITEM hRoot);
	void FillCategory(HTREEITEM hRoot, CatEntry* pCatEntry);
	void FillBands(HTREEITEM hRoot);
	void FillChildBands(IChildBands* pChildBands, HTREEITEM hChildBands);
	void FillTools(ITools* pTools, HTREEITEM hChildBand);
	
	ImageTypes GetBandImageType(BandTypes btTypes);
	ImageTypes GetToolImageType(ToolTypes ctTool);

	LRESULT OnNotify(LPNMHDR pNMHdr);
	void OnBeginDrag(NM_TREEVIEW*  pTreeView);
	void OnDeleteItem(NM_TREEVIEW* pTreeView);
	void OnDoubleClick();
	void OnExpanded(NM_TREEVIEW* pTreeView);
	void OnExpanding(NM_TREEVIEW* pTreeView);
	void OnGetDisplayInfo(TV_DISPINFO* pDispInfo);
	void OnRightClick();
	void OnClick();
	void OnSelChanged(NM_TREEVIEW* pNMHdr);

	LRESULT OnKeyDown(LPNMTVKEYDOWN pKeyDown);
	LRESULT OnBeginEditLabel(TV_DISPINFO* pDispInfo);
	LRESULT OnEndEditLabel(TV_DISPINFO* pDispInfo);
	
	void OnSetDisplayInfo(TV_DISPINFO* pDispInfo);

	void OnCommand(WPARAM wParam, LPARAM lParam);
	void OnCreateBand(UINT nType);
	void OnCreateChildBand();
	void OnCreateTool(ToolTypes ttTool);
	void OnDeleteBand();
	void OnDeleteChildBand();
	void OnDeleteTool();
	void OnApplyAll();
	void OnCreateCategory();
	void OnDeleteCategory();
	void OnLoadAB10(LPCTSTR szFileName);
	void OnLoadAB20(LPCTSTR szFileName);

	DesignerTreeType m_eTreeType;
	CDesignerTarget* m_pTreeCtrlTarget;
	CategoryMgr*     m_pCategoryMgr;
	CBrowser*	     m_pBrowser;
	DDString		 m_strNone;
	TCHAR		     m_szCurFileName[_MAX_PATH];
};

inline CategoryMgr* CDesignerTree::GetCategoryMgr()
{
	return m_pCategoryMgr;
}

//
// Update
//

inline void CDesignerTree::Update()
{
	::SendMessage(::GetParent(m_hWnd), WM_NOTIFY, MAKELONG(ID_SETPAGEDIRTY, 0), (LPARAM)m_hWnd);
}

inline LPTSTR CDesignerTree::CurFileName()
{
	return &m_szCurFileName[0];
}

//
// CDesignerTarget
//

class CDesignerTarget : public IDropTarget
{
public:
	enum
	{
		eNumOfFormats = 6
	};

	CDesignerTarget(long nDestinationId, CDesignerTree& theTreeCtrl);

    virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject);
        
    virtual ULONG __stdcall AddRef();
        
    virtual ULONG __stdcall Release();

    virtual HRESULT __stdcall DragEnter(IDataObject* pDataObj, 
										DWORD        grfKeyState,
										POINTL       pt,
										DWORD*       pdwEffect);
        
    virtual HRESULT __stdcall DragOver(DWORD  grfKeyState,
									   POINTL pt,
									   DWORD* pdwEffect);
        
    virtual HRESULT __stdcall DragLeave();
        
    virtual HRESULT __stdcall Drop(IDataObject* pDataObj,
								   DWORD		grfKeyState,
								   POINTL		pt,
								   DWORD*		pdwEffect);
private:
	void FeedBack(int nImage, DWORD* pdwEffect, DWORD grfKeyState, WORD nToolFormat);

	CDesignerTree& m_theTreeCtrl;
	IDataObject*   m_pDataObject;
	IActiveBar2* m_pSrcBar;
	FORMATETC m_fe[6];
	DWORD m_dwStartDelay;
	DWORD m_dwCurrentDelay;
	CRect m_rcClient;
	CRect m_rcScroll;
	BOOL  m_bAutoScrolling;
	WORD  m_nToolFormat;
	long  m_nDestId;
	long  m_nSrcId;
	long  m_cRef;
	int   m_nIndex;
};

//
// CTreeCtrlSource
//

class CTreeCtrlSource : public IDropSource
{
public:
	CTreeCtrlSource()
	{
		m_cRef = 1;
	}

    virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject);
        
    virtual ULONG __stdcall AddRef();
        
    virtual ULONG __stdcall Release();

	virtual HRESULT __stdcall QueryContinueDrag(BOOL  bEscapePressed,
												DWORD grfKeyState);
        
    virtual HRESULT __stdcall GiveFeedback(DWORD dwEffect);
private:
	long m_cRef;
};

#endif