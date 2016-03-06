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
#include "Dialogs.h"
#include "Stream.h"
#include "Globals.h"
#include "Browser.h"
#include "Debug.h"
#include "DISPIDS.h"
#include "DISPIDS2.h"
#include "DesignerPage.h"
#include "DragDrop.h"
#include "DragDropMgr.h"
#include "..\EventLog.h"
#include "StringUtil.h"
#include "CategoryMgr.h"
#include "AB10Format.h"
#include "Trees.h"

extern HINSTANCE g_hInstance;
#define WM_CLEARSELECTION

//
// SetDefFormatEtc
//

static inline void SetDefFormatEtc(FORMATETC& fe, CLIPFORMAT& cf, DWORD dwMed)
{
    fe.cfFormat = cf;
    fe.dwAspect = DVASPECT_CONTENT;
    fe.ptd = 0;
    fe.tymed = dwMed;
	fe.lindex = -1;
}

//
// CTreeCtrl
//

CTreeCtrl::CTreeCtrl()
	: m_ilImages(NULL),
	  m_pActiveBar(NULL),
	  m_hCurrentItem(NULL),
	  m_hFirstSelectedItem(NULL)
{
	m_tvCurrentItem.pszText = m_szCurrentText;
	m_szClassName = WC_TREEVIEW;
	m_bMultSelectActive = FALSE;
}

CTreeCtrl::~CTreeCtrl()
{
	if (m_ilImages)
		ImageList_Destroy(m_ilImages);
}


void CTreeCtrl::Refresh(HTREEITEM hParent)
{
	try
	{
		SendMessage(WM_SETREDRAW, FALSE); 
		TV_ITEM tvItem;
		tvItem.hItem = hParent;
		tvItem.mask = TVIF_STATE|TVIF_HANDLE; 
		BOOL bResult = TreeView_GetItem(m_hWnd, &tvItem);
		if (bResult && TVIS_EXPANDED & tvItem.state)
		{
			tvItem.mask = TVIF_CHILDREN; 
			tvItem.cChildren = 1;
			bResult = TreeView_SetItem(m_hWnd, (LRESULT)&tvItem); 
			bResult = TreeView_Expand(m_hWnd, tvItem.hItem, TVE_COLLAPSE|TVE_COLLAPSERESET);
			bResult = TreeView_Expand(m_hWnd, tvItem.hItem, TVE_EXPAND);
			tvItem.cChildren = I_CHILDRENCALLBACK;
			bResult = TreeView_SetItem(m_hWnd, (LRESULT)&tvItem); 
		}
		SendMessage(WM_SETREDRAW, TRUE);
		UpdateWindow();
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// SelectDropTarget
//

void CTreeCtrl::SelectDropTarget(HTREEITEM hItem)
{
	static HTREEITEM hLastItem = NULL;

	if (hLastItem || NULL == hItem && hLastItem)
	{
		SetItemState(hLastItem, 0, TVIS_DROPHILITED);
		hLastItem = NULL;
	}

	BOOL bSelect = (GetItemState(hItem, TVIS_DROPHILITED) & TVIS_DROPHILITED);
	if (!bSelect)
	{
		hLastItem = hItem;
		SetItemState(hLastItem, TVIS_DROPHILITED, TVIS_DROPHILITED);
	}
}

//
// DeleteChildren
//

void CTreeCtrl::DeleteChildren(HTREEITEM hParent)
{
	TreeView_Expand(m_hWnd, hParent, TVE_COLLAPSE|TVE_COLLAPSERESET);
}

//
// LoadImages
//

BOOL CTreeCtrl::LoadImages(UINT nResourceId)
{
	if (m_ilImages)
		return TRUE;

	m_ilImages = ImageList_LoadImage(g_hInstance, 
									 MAKEINTRESOURCE(nResourceId), 
									 16, 
									 0, 
									 CLR_DEFAULT,
									 IMAGE_BITMAP,
									 LR_LOADTRANSPARENT|LR_CREATEDIBSECTION);
	if (NULL == m_ilImages)
		return FALSE;
	ImageList_SetBkColor(m_ilImages, CLR_NONE);
    TreeView_SetImageList(m_hWnd, m_ilImages, TVSIL_NORMAL); 
	return TRUE;
}

//
// WindowProc
//

LRESULT CTreeCtrl::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_LBUTTONDOWN:
		{
			POINT pt;
			pt.x = LOWORD(lParam);
			pt.y = HIWORD(lParam);
			LRESULT lResult = OnLButtonDown(wParam, pt);
			if (0 == lResult)
				return 0;
		}
		break;

	case WM_RBUTTONDOWN:
		{
			POINT pt;
			pt.x = LOWORD(lParam);
			pt.y = HIWORD(lParam);
			UINT nHitFlags = 0;
			HTREEITEM hClickedItem = HitTest(pt, &nHitFlags);
			if (TVHT_ONITEM & nHitFlags)
				TreeView_SelectItem(m_hWnd, hClickedItem);
		}
		break;

	case WM_KEYDOWN:
		{
			switch (wParam)
			{
			case VK_UP:
			case VK_DOWN:
				{
					// Find which item is currently selected
					HTREEITEM hSelectedItem = TreeView_GetSelection(m_hWnd);

					HTREEITEM hNextItem;
					if (VK_UP == wParam)
						hNextItem = TreeView_GetPrevVisible(m_hWnd, hSelectedItem);
					else
						hNextItem = TreeView_GetNextVisible(m_hWnd, hSelectedItem);

					if (!(GetKeyState(VK_SHIFT) & 0x8000))
					{
						// User pressed arrow key without holding 'Shift':
						// Clear multiple selection and let base class do normal 
						// selection work!
						ClearSelection();
						if (hNextItem)
							break;
					}

					LRESULT lResult = FWnd::WindowProc(nMsg, wParam, lParam);

					// Select from first selected item to the clicked item
					if (!m_hFirstSelectedItem)
					{
						m_hFirstSelectedItem = hSelectedItem;
						SetItemState(hSelectedItem, TVIS_DROPHILITED, TVIS_DROPHILITED);
					}
					SelectItems(m_hFirstSelectedItem, hNextItem);
					TreeView_SelectItem(m_hWnd, hNextItem);
					return lResult;
				}
				break;

			case VK_CONTROL:
				{
					if (!(HIWORD(lParam) & KF_REPEAT))
					{
						HTREEITEM hClickedItem = TreeView_GetSelection(m_hWnd); 
						if (hClickedItem)
							SelectMultiple(hClickedItem, MK_CONTROL);
					}
				}
				break;
			}
		}
		break;

	case WM_NOTIFY:
		return OnNotify(reinterpret_cast<LPNMHDR>(lParam));

	case WM_COMMAND:
		OnCommand(wParam, lParam);
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// OnLButtonDown
//

LRESULT CTreeCtrl::OnLButtonDown(WPARAM wParam, POINT& pt)
{
	try
	{
		UINT nHitFlags = 0;
		HTREEITEM hClickedItem = HitTest(pt, &nHitFlags);
		if (TVHT_ONITEM & nHitFlags && ((MK_CONTROL|MK_SHIFT) & wParam))
		{
			SelectMultiple(hClickedItem, wParam);
			return 0;
		}
		if (m_bMultSelectActive)
		{
			BOOL bFound = FALSE;
			TVITEM tvItem;
			tvItem.mask = TVIF_IMAGE|TVIF_PARAM; 
			tvItem.hItem = GetFirstSelectedItem();
			while (tvItem.hItem)
			{
				if (TreeView_GetItem(m_hWnd, &tvItem) && tvItem.hItem == hClickedItem)
				{
					bFound = TRUE;
					break;
				}
				tvItem.hItem = GetNextSelectedItem(tvItem.hItem);
			}
			if (bFound)
			{
				DWORD nDragStartTick = GetTickCount();
				DWORD nDragDelay = GetProfileInt(_T("Windows"), _T("DragDelay"), DD_DEFDRAGDELAY);
				MSG msg;
				BOOL m_bProcessing = TRUE;
				HWND hWndPrev = SetCapture(m_hWnd);
				while (m_hWnd == GetCapture() && m_bProcessing)
				{
					GetMessage(&msg, NULL, 0, 0);

					switch (msg.message)
					{
					case WM_CANCELMODE:
						m_bProcessing = FALSE;
						break;

					case WM_LBUTTONDOWN:
						break;

					case WM_LBUTTONUP:
						ClearSelection();
						TreeView_SelectItem(m_hWnd, hClickedItem);
						ReleaseCapture();
						return 0;

					case WM_MOUSEMOVE:
						if (abs(msg.pt.x - pt.x) > GetSystemMetrics(SM_CXDRAG) || abs(msg.pt.y - pt.y) > GetSystemMetrics(SM_CYDRAG) || (GetTickCount()-nDragStartTick) < nDragDelay)
						{
							HWND hWndParent = ::GetParent(m_hWnd);
							if (IsWindow(hWndParent))
							{
								NM_TREEVIEW tv;
								tv.hdr.hwndFrom = m_hWnd;
								tv.hdr.idFrom = GetWindowLong(m_hWnd, GWL_ID);
								tv.hdr.code = TVN_BEGINDRAG;
								tv.itemNew.hItem = hClickedItem;
								tv.itemNew.state = GetItemState(hClickedItem, 0xffffffff);
								tv.itemNew.lParam = GetItemData(hClickedItem);
								tv.ptDrag.x = pt.x;
								tv.ptDrag.y = pt.y;
								::PostMessage(hWndParent, WM_NOTIFY, tv.hdr.idFrom, (LPARAM)&tv);
								m_bProcessing = FALSE;
							}
						}
						break;

					default:
						TranslateMessage(&msg);
						DispatchMessage(&msg);
						break;
					}
				}
			}
			else
				ClearSelection();
		}
		else
			ClearSelection();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	} 
	return -1;
}

//
// OnLButtonUp
//

LRESULT CTreeCtrl::OnLButtonUp(WPARAM wParam, POINT& pt)
{
	return -1;
}

//
// PopupMenu
//

BOOL CTreeCtrl::PopupMenu(UINT nId, int nSubMenuPos, POINT& pt)
{
	HMENU hMenu = LoadMenu(g_hInstance, MAKEINTRESOURCE(nId)); 
	if (NULL == hMenu)
		return FALSE;
	HMENU hMenuSub = GetSubMenu(hMenu, nSubMenuPos);
	if (NULL == hMenuSub)
	{
		DestroyMenu(hMenu);
		return FALSE;
	}
	ClientToScreen(pt);
	BOOL bResult = TrackPopupMenu(hMenuSub, 
								  TPM_LEFTALIGN, 
								  pt.x, 
								  pt.y, 
								  0, 
								  m_hWnd, 
								  0);
	DestroyMenu(hMenu);
	return bResult;
}

//
// ChildCount
//

int CTreeCtrl::ChildCount(HTREEITEM hParent)
{
	int nCount = 0;
	HTREEITEM hTreeItem = TreeView_GetChild(m_hWnd, hParent);
	while (hTreeItem)
	{
		nCount++;
		hTreeItem = TreeView_GetNextSibling(m_hWnd, hTreeItem);
	}
	return nCount;
}

//
// FindIndex
//

int CTreeCtrl::FindIndex(HTREEITEM hParent, HTREEITEM hChild)
{
	int nIndex = -1;
	HTREEITEM hTreeItem = TreeView_GetChild(m_hWnd, hParent);
	while (hTreeItem)
	{
		nIndex++;
		if (hTreeItem == hChild)
			break;
		hTreeItem = TreeView_GetNextSibling(m_hWnd, hTreeItem);
	}
	return nIndex;
}

//
// InsertNode
//

HTREEITEM CTreeCtrl::InsertNode (HTREEITEM hParent, 
								 LPTSTR    szText, 
								 int       nImage, 
								 int       nChild, 
								 LPARAM    lParam, 
								 HTREEITEM hAfter)
{
	TV_INSERTSTRUCT theInsertItem;
    theInsertItem.hParent = hParent;     
	theInsertItem.hInsertAfter = hAfter;
    theInsertItem.item.mask = TVIF_IMAGE|TVIF_SELECTEDIMAGE|TVIF_TEXT;
	
	if (-2 != nChild)
		theInsertItem.item.mask |= TVIF_CHILDREN;

	if (-2 != lParam)
		theInsertItem.item.mask |= TVIF_PARAM;

	theInsertItem.item.hItem = 0;     
	theInsertItem.item.state = 0; 
    theInsertItem.item.stateMask = 0; 
    theInsertItem.item.pszText = szText;    
	if (LPSTR_TEXTCALLBACK != szText)
		theInsertItem.item.cchTextMax = lstrlen(szText) + 1; 
	else
		theInsertItem.item.cchTextMax = -1; 

    theInsertItem.item.iImage = nImage;   
	theInsertItem.item.iSelectedImage = nImage;   
	theInsertItem.item.cChildren = nChild; 
    theInsertItem.item.lParam = lParam; 	

	HTREEITEM hTreeItem = TreeView_InsertItem(m_hWnd, &theInsertItem);
	if (NULL == hTreeItem)
		return NULL;

	if (TVI_ROOT != hParent)
	{
		CRect rcItem;
		if (TreeView_GetItemRect(m_hWnd, hParent, &rcItem, FALSE))
			InvalidateRect(&rcItem, FALSE);
	}
	return hTreeItem;
}

//
// HitTest
//

BOOL CTreeCtrl::HitTestAtCursor(POINT& pt, UINT nMask)
{
	m_hCurrentItem = NULL;
	TV_HITTESTINFO hitInfo;
	hitInfo.flags = TVHT_ONITEM;
	GetCursorPos(&hitInfo.pt);
	ScreenToClient(hitInfo.pt);
	m_hCurrentItem = TreeView_HitTest(m_hWnd, &hitInfo);
	if (!(TVHT_ONITEM & hitInfo.flags))
		return FALSE;

	m_hCurrentItem = hitInfo.hItem;
	m_tvCurrentItem.mask = nMask;
	m_tvCurrentItem.hItem = m_hCurrentItem;
	pt = hitInfo.pt;
	return TreeView_GetItem(m_hWnd, &m_tvCurrentItem);
}

//
// HitTest
//

BOOL CTreeCtrl::HitTest(POINT& pt, UINT nMask)
{
	TV_HITTESTINFO hitInfo;
	hitInfo.pt = pt;
	ScreenToClient(hitInfo.pt);
	hitInfo.flags = TVHT_ONITEM;
	m_hCurrentItem = TreeView_HitTest(m_hWnd, &hitInfo);
	if (NULL == m_hCurrentItem && !(TVHT_ONITEM & hitInfo.flags))
		return FALSE;

	m_tvCurrentItem.mask = nMask;
	m_tvCurrentItem.hItem = m_hCurrentItem;
	return TreeView_GetItem(m_hWnd, &m_tvCurrentItem);
}

//
// GetSelected
//

BOOL CTreeCtrl::GetSelected(UINT nMask)
{
	m_hCurrentItem = TreeView_GetSelection(m_hWnd); 
	if (NULL == m_hCurrentItem)
		return FALSE;

	m_tvCurrentItem.mask = nMask;
	m_tvCurrentItem.hItem = m_hCurrentItem;
	return TreeView_GetItem(m_hWnd, &m_tvCurrentItem);
}

//
// GetParent
//

BOOL CTreeCtrl::GetParent(HTREEITEM hItem, TV_ITEM& tvParent, DWORD dwMask)
{
	HTREEITEM hParent = TreeView_GetParent(m_hWnd, hItem);
	if (NULL == hParent)
		return FALSE;
	tvParent.mask = dwMask;
	tvParent.hItem = hParent;
	return TreeView_GetItem(m_hWnd, &tvParent);
}

//
// Select visible items between specified 'from' and 'to' item (including these!)
// If the 'to' item is above the 'from' item, it traverses the tree in reverse 
// direction. Selection on other items is cleared!
//

BOOL CTreeCtrl::SelectItems(HTREEITEM hFromItem, HTREEITEM hToItem)
{
	// Determine direction of selection 
	// (see what item comes first in the tree)
	HTREEITEM hItem = TreeView_GetRoot(m_hWnd);

	while (hItem && hItem != hFromItem && hItem != hToItem)
		hItem = TreeView_GetNextVisible(m_hWnd, hItem);

	if (!hItem)
		return FALSE;	// Items not visible in tree

	// Really select the 'to' item (which will deselect 
	// the previously selected item)

	BOOL bReverse = hItem == hToItem;

	hItem = TreeView_GetRoot(m_hWnd);
	BOOL bSelect = FALSE;

	while (hItem)
	{
		if (hItem == (bReverse ? hToItem : hFromItem))
			bSelect = TRUE;

		if (bSelect)
		{
			//
			// Selecting
			//

			if (!(GetItemState(hItem, TVIS_DROPHILITED) & TVIS_DROPHILITED))
				SetItemState(hItem, TVIS_DROPHILITED, TVIS_DROPHILITED);
		}
		else
		{
			//
			// Clearing
			//

			if (GetItemState(hItem, TVIS_DROPHILITED) & TVIS_DROPHILITED)
				SetItemState(hItem, 0, TVIS_DROPHILITED);
		}

		if (hItem == (bReverse ? hFromItem : hToItem))
			bSelect = FALSE;

		hItem = TreeView_GetNextVisible(m_hWnd, hItem);
	}
	return TRUE;
}

//
// Clear selected state on all visible items
//

void CTreeCtrl::ClearSelection()
{
	m_bMultSelectActive = FALSE;
	m_hFirstSelectedItem = NULL;
	for (HTREEITEM hItem = TreeView_GetRoot(m_hWnd); hItem; hItem = TreeView_GetNextVisible(m_hWnd, hItem))
	{
		if (GetItemState(hItem, TVIS_DROPHILITED) & TVIS_DROPHILITED)
			SetItemState(hItem, 0, TVIS_DROPHILITED);
	}
}

//
// SelectMultiple
//

void CTreeCtrl::SelectMultiple( HTREEITEM hClickedItem, UINT nFlags )
{
	HTREEITEM hSelectedItem = TreeView_GetSelection(m_hWnd);
	// Action depends on whether the user holds down the Shift or Ctrl key
	if (nFlags & MK_SHIFT)
	{
		m_bMultSelectActive = TRUE;
		// Select from first selected item to the clicked item
		if (!m_hFirstSelectedItem)
			m_hFirstSelectedItem = hSelectedItem;
		SelectItems(m_hFirstSelectedItem, hClickedItem);
		TreeView_SelectItem(m_hWnd, hClickedItem);
	}
	else if (nFlags & MK_CONTROL)
	{
		m_bMultSelectActive = TRUE;
		// Is the clicked item already selected ?
		DWORD bIsClickedItemSelected = GetItemState(hClickedItem, TVIS_DROPHILITED) & TVIS_DROPHILITED;
		BOOL bIsSelectedItemSelected = FALSE;

		if (hSelectedItem)
			bIsSelectedItemSelected = GetItemState(hSelectedItem, TVIS_DROPHILITED) & TVIS_DROPHILITED;

		// We want the newly selected item to toggle its selected state,
		// so unselect now if it was already selected before
		if (bIsClickedItemSelected)
			SetItemState(hClickedItem, 0, TVIS_DROPHILITED|TVIS_SELECTED);
		else
			SetItemState(hClickedItem, TVIS_SELECTED|TVIS_DROPHILITED, TVIS_DROPHILITED|TVIS_SELECTED);

		// Store as first selected item (if not already stored)
		if (NULL ==  m_hFirstSelectedItem)
			m_hFirstSelectedItem = hClickedItem;
	}
}

//
// Get number of selected items
//

UINT CTreeCtrl::GetSelectedCount() const
{
	// Only visible items should be selected!
	UINT uCount = 0;
	for (HTREEITEM hItem = TreeView_GetRoot(m_hWnd); hItem; hItem = TreeView_GetNextVisible(m_hWnd, hItem))
		if (GetItemState(hItem, TVIS_DROPHILITED) & TVIS_DROPHILITED)
			uCount++;
	return uCount;
}

//
// Helpers to list out selected items. (Use similar to GetFirstVisibleItem(), 
// GetNextVisibleItem() and GetPrevVisibleItem()!)
//

HTREEITEM CTreeCtrl::GetFirstSelectedItem()
{
	for (HTREEITEM hItem = TreeView_GetRoot(m_hWnd); hItem; hItem = TreeView_GetNextVisible(m_hWnd, hItem))
		if (GetItemState(hItem, TVIS_DROPHILITED) & TVIS_DROPHILITED)
			return hItem;
	return NULL;
}

//
// GetNextSelectedItem
//

HTREEITEM CTreeCtrl::GetNextSelectedItem(HTREEITEM hItem)
{
	for (hItem = TreeView_GetNextVisible(m_hWnd, hItem); hItem; hItem = TreeView_GetNextVisible(m_hWnd, hItem))
		if (GetItemState(hItem, TVIS_DROPHILITED) & TVIS_DROPHILITED)
			return hItem;
	return NULL;
}

//
// GetPrevSelectedItem
//

HTREEITEM CTreeCtrl::GetPrevSelectedItem(HTREEITEM hItem)
{
	for (hItem = TreeView_GetPrevVisible(m_hWnd, hItem); hItem; hItem = TreeView_GetPrevVisible(m_hWnd, hItem))
		if (GetItemState(hItem, TVIS_DROPHILITED) & TVIS_DROPHILITED)
			return hItem;
	return NULL;
}

//
// HitTest
//

HTREEITEM CTreeCtrl::HitTest(POINT pt, UINT* pFlags) const
{
	assert(::IsWindow(m_hWnd));
	TVHITTESTINFO hti;
	hti.pt = pt;
	HTREEITEM h = (HTREEITEM)::SendMessage(m_hWnd, TVM_HITTEST, 0, (LPARAM)&hti);
	if (pFlags)
		*pFlags = hti.flags;
	return h;
}

//
// CDesignerTree
//

CDesignerTree::CDesignerTree(DesignerTreeType eTreeType)
	: m_pDataObject(NULL),
	  m_pTreeCtrlTarget(NULL)
{
	m_eTreeType = eTreeType;
	m_szMsgTitle.LoadString(IDS_ACTIVEBAR);
	m_strNone.LoadString(IDS_NONE);
	*m_szCurFileName = NULL;
}

CDesignerTree::~CDesignerTree()
{
	if (IsWindow(m_hWnd))
		TreeView_DeleteAllItems(m_hWnd);

	if (m_pActiveBar)
	{
		IBarPrivate* pPrivate;
		HRESULT hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
		if (SUCCEEDED(hResult))
		{
			pPrivate->RevokeBandChange();
			pPrivate->Release();
		}
	}
	
	UnsubClass();

	GetGlobals().GetDragDropMgr()->RevokeDragDrop((OLE_HANDLE)m_hWnd);
	if (m_pTreeCtrlTarget)
		m_pTreeCtrlTarget->Release();
}

void CDesignerTree::SettingsChanged()
{
	if (m_ilImages)
	{
		ImageList_Destroy(m_ilImages);
		m_ilImages = NULL;
	}
	BOOL bResult = LoadImages(IDR_DESIGNER);
	assert(bResult);
}

//
// FindBand
//

HTREEITEM CDesignerTree::FindBand(LPCTSTR szBandName)
{
	try
	{
		HTREEITEM hTreeItem = TreeView_GetRoot(m_hWnd);
		if (NULL == hTreeItem)
			return NULL;

		BOOL bResult = TreeView_Expand(m_hWnd, hTreeItem, TVE_EXPAND);

		hTreeItem = TreeView_GetChild(m_hWnd, hTreeItem);
		if (NULL == hTreeItem)
			return NULL;


		hTreeItem = TreeView_GetNextSibling(m_hWnd, hTreeItem);
		if (NULL == hTreeItem)
			return NULL;

		TV_ITEM tvItem;
		tvItem.mask = TVIF_STATE; 
		tvItem.hItem = hTreeItem;
		if (TreeView_GetItem(m_hWnd, &tvItem) && !(TVIS_EXPANDED & tvItem.state))
		{
			bResult = TreeView_Expand(m_hWnd, hTreeItem, TVE_EXPAND);
			if (bResult)
			{
				if (TreeView_GetItem(m_hWnd, &tvItem) && !(TVIS_EXPANDED & tvItem.state))
					return NULL;
			}
			else
				return NULL;
		}

		hTreeItem = TreeView_GetChild(m_hWnd, hTreeItem);
		if (NULL == hTreeItem)
			return NULL;

		// Find Band in the Tree
		const int nSize = 100;
		TCHAR szBuffer[nSize];
		tvItem.mask = TVIF_TEXT; 
		tvItem.pszText = szBuffer;
		tvItem.cchTextMax = nSize;
		while (hTreeItem)
		{
			tvItem.hItem = hTreeItem;
			if (TreeView_GetItem(m_hWnd, &tvItem) && 0 == _tcsicmp(tvItem.pszText, szBandName))
				return hTreeItem;
			hTreeItem = TreeView_GetNextSibling(m_hWnd, hTreeItem);
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return NULL;
}

//
// FindChildBand
//

HTREEITEM CDesignerTree::FindChildBand(HTREEITEM hChildBands, LPCTSTR szChildBand)
{
	try
	{
		HTREEITEM hChildBand = TreeView_GetChild(m_hWnd, hChildBands);
		if (NULL == hChildBand)
			return NULL;

		// Find ChildBand in the Tree
		const int nSize = 100;
		TCHAR szBuffer[nSize];
		TV_ITEM tvItem;
		tvItem.mask = TVIF_TEXT; 
		tvItem.pszText = szBuffer;
		tvItem.cchTextMax = nSize;
		while (hChildBand)
		{
			tvItem.hItem = hChildBand;
			if (TreeView_GetItem(m_hWnd, &tvItem) && 0 == _tcsicmp(tvItem.pszText, szChildBand))
				return hChildBand;
			hChildBand = TreeView_GetNextSibling(m_hWnd, hChildBand);
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return NULL;
}

//
// FindTool
//

HTREEITEM CDesignerTree::FindTool(HTREEITEM hTool, long nToolIdIn)
{
	try
	{
		HRESULT hResult;
		long nToolId;
		ITool* pTool;
		TV_ITEM tvTool;
		tvTool.mask = TVIF_PARAM; 
		while (hTool)
		{
			tvTool.hItem = hTool;
			if (TreeView_GetItem(m_hWnd, &tvTool))
			{
				pTool = reinterpret_cast<ITool*>(tvTool.lParam);
				if (pTool)
				{
					hResult = pTool->get_ID(&nToolId);
					if (FAILED(hResult))
						continue;
					if (nToolId == nToolIdIn)
						return hTool;
				}
			}
			hTool = TreeView_GetNextSibling(m_hWnd, hTool);
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return NULL;
}

//
// BandChanged
//

STDMETHODIMP CDesignerTree::BandChanged(LPDISPATCH pBandIn, LPDISPATCH pChildBandIn, BandChangeTypes nType)
{
	try
	{
		IBand* pBand = (IBand*)pBandIn;
		IBand* pChildBand = (IBand*)pChildBandIn;
		switch (nType)
		{
		case ddBandModified:
			{
				//
				// Band Modified
				//

				BSTR bstrName;
				HRESULT hResult = pBand->get_Name(&bstrName);
				if (FAILED(hResult))
					return E_FAIL;

				MAKE_TCHARPTR_FROMWIDE(szBandName, bstrName);
				SysFreeString(bstrName);
				bstrName = NULL;

				// Find Band in the Tree
				HTREEITEM hBand = FindBand(szBandName);
				if (NULL == hBand)
					return NOERROR;

				BandTypes btType;
				hResult = pBand->get_Type(&btType);
				if (FAILED(hResult))
					return E_FAIL;

				int iImage = (int)GetBandImageType(btType);
				TVITEM tvItem;
				tvItem.hItem = hBand;
				tvItem.mask = TVIF_IMAGE|TVIF_SELECTEDIMAGE; 
				if (TreeView_GetItem(m_hWnd, &tvItem) && iImage != tvItem.iImage)
				{
					tvItem.iImage = tvItem.iSelectedImage = iImage;
					TreeView_SetItem(m_hWnd, &tvItem); 
				}

				if (ddBTNormal == btType)
				{
					ChildBandStyles cbStyle;
					hResult = pBand->get_ChildBandStyle(&cbStyle);
					if (FAILED(hResult))
						return E_FAIL;

					if (ddCBSNone != cbStyle)
					{
						HTREEITEM hChildBands = TreeView_GetChild(m_hWnd, hBand);
						if (NULL == hChildBands)
							return NOERROR;

						hResult = pBand->get_ChildBands((ChildBands**)&pChildBand);
						if (FAILED(hResult))
							return E_FAIL;

						hResult = pChildBand->get_Name(&bstrName);
						pChildBand->Release();
						if (FAILED(hResult))
							return E_FAIL;

						MAKE_TCHARPTR_FROMWIDE(szChildBand, bstrName);
						SysFreeString(bstrName);
						bstrName = NULL;

						HTREEITEM hChildBand = FindChildBand(hChildBands, szChildBand);
						if (NULL == hChildBand)
						{
							Refresh(hChildBand);
							return E_FAIL;
						}
					}
					else
						Refresh(hBand);
				}
				else
					Refresh(hBand);
			}
			break;

		case ddBandDeleted:
			{
				//
				// Band Deleted
				//

				BSTR bstrName;
				HRESULT hResult = pBand->get_Name(&bstrName);
				if (FAILED(hResult))
					return E_FAIL;

				MAKE_TCHARPTR_FROMWIDE(szBandName, bstrName);
				SysFreeString(bstrName);
				bstrName = NULL;

				HTREEITEM hBand = FindBand(szBandName);
				if (hBand)
					TreeView_DeleteItem(m_hWnd, hBand);
			}
			break;

		case ddBandAdded:
			{
				//
				// Band Added
				//

				if (!IsWindow(m_hWnd))
					return E_FAIL;

				HTREEITEM hTree = TreeView_GetRoot(m_hWnd);
				if (NULL == hTree)
					return E_FAIL;

				hTree = TreeView_GetChild(m_hWnd, hTree); 
				if (NULL == hTree)
					return E_FAIL;

				hTree = TreeView_GetNextSibling(m_hWnd, hTree);
				if (NULL == hTree)
					return E_FAIL;

				BSTR bstrName;
				HRESULT hResult = pBand->get_Name(&bstrName);
				if (FAILED(hResult))
					return E_FAIL;

				MAKE_TCHARPTR_FROMWIDE(szBandName, bstrName);
				SysFreeString(bstrName);
				bstrName = NULL;

				BandTypes btType;
				hResult = pBand->get_Type(&btType);
				if (FAILED(hResult))
					return E_FAIL;

				TVITEM tvItem;
				tvItem.mask = TVIF_STATE;
				tvItem.hItem = hTree;
				BOOL bResult = TreeView_GetItem(m_hWnd, &tvItem);
				if (bResult && tvItem.state & TVIS_EXPANDED)
				{
					pBand->AddRef();
					HTREEITEM hBand = InsertNode(hTree, 
												 LPSTR_TEXTCALLBACK, 
												 GetBandImageType(btType),
												 I_CHILDRENCALLBACK, 
												 (long)pBand);
				}
				else
				{
					SetItemState(hTree, 0, TVIS_EXPANDEDONCE);
					bResult = TreeView_Expand(m_hWnd, hTree, TVE_TOGGLE);
				}
				CRect rcItem;
				bResult = TreeView_GetItemRect(m_hWnd, hTree, &rcItem, FALSE);
				if (bResult)
					InvalidateRect(&rcItem, FALSE);

			}
			break;
		}
		return S_OK;
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// Is called by ActiveBar to select a Band, ChildBand, or a Tool by pressing on the 
// object.
//

STDMETHODIMP CDesignerTree::ToolSelected(BSTR bstrBand, BSTR bstrChildBand, long nToolId)
{
	try
	{
		if (NULL == bstrBand)
			return S_OK;

		MAKE_TCHARPTR_FROMWIDE(szBand, bstrBand);
		HTREEITEM hBand = FindBand(szBand);
		if (NULL == hBand)
			return S_OK;

		TV_ITEM tvItem;
		tvItem.mask = TVIF_STATE; 
		tvItem.hItem = hBand;
		if (!TreeView_GetItem(m_hWnd, &tvItem))
			return S_OK;

		if (TVIS_EXPANDED & tvItem.state)
			Refresh(hBand);

		TreeView_Expand(m_hWnd, hBand, TVE_EXPAND);
		HTREEITEM hTool = NULL;
		HTREEITEM hChildBands = TreeView_GetChild(m_hWnd, hBand);
		if (NULL == hChildBands)
		{
			hTool = TreeView_GetChild(m_hWnd, hBand);
			if (hTool)
			{
				hTool = FindTool(hTool, nToolId);
				if (hTool)
					TreeView_Select(m_hWnd, hTool, TVGN_CARET); 
				else
					TreeView_Select(m_hWnd, hBand, TVGN_CARET); 
			}
			else
				TreeView_Select(m_hWnd, hBand, TVGN_CARET); 
			return S_OK;
		}

		tvItem.mask = TVIF_IMAGE|TVIF_STATE; 
		tvItem.hItem = hChildBands;
		if (!TreeView_GetItem(m_hWnd, &tvItem))
			return S_OK;

		if (eChildBands == tvItem.iImage)
		{
			//
			// A band with ChildBands
			//

			TreeView_Expand(m_hWnd, hChildBands, TVE_EXPAND);
			if (NULL == bstrChildBand)
			{
				TreeView_Select(m_hWnd, hBand, TVGN_CARET); 
				return S_OK;
			}

			MAKE_TCHARPTR_FROMWIDE(szChildBand, bstrChildBand);
			HTREEITEM hChildBand = FindChildBand(hChildBands, szChildBand);
			if (NULL == hChildBand)
				return S_OK;

			TreeView_Expand(m_hWnd, hChildBand, TVE_EXPAND);
			hTool = TreeView_GetChild(m_hWnd, hChildBand);
			if (hTool)
			{
				hTool = FindTool(hTool, nToolId);
				if (hTool)
					TreeView_Select(m_hWnd, hTool, TVGN_CARET); 
				else
					TreeView_Select(m_hWnd, hChildBand, TVGN_CARET); 
			}
			else
				TreeView_Select(m_hWnd, hChildBand, TVGN_CARET); 
		}
		else
		{
			//
			// A band without childbands
			//

			hTool = TreeView_GetChild(m_hWnd, hBand);
			if (hTool)
			{
				hTool = FindTool(hTool, nToolId);
				if (hTool)
					TreeView_Select(m_hWnd, hTool, TVGN_CARET); 
				else
					TreeView_Select(m_hWnd, hBand, TVGN_CARET); 
			}
			else
				TreeView_Select(m_hWnd, hBand, TVGN_CARET); 
		}
		return S_OK;
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// UpdateBarTool
//

void UpdateBarTool (IBand* pBand, IBand* pChildBand, ITool* pTool, void* pData)
{
	ITool* pModifiedTool = static_cast<ITool*>(pData);
	if (pModifiedTool == pTool)
		return;
	
	pModifiedTool->CopyTo(&pTool);
}

//
// ObjectChanged
//

void CDesignerTree::ObjectChanged(int nType)
{
	try
	{
		if (!GetSelected(TVIF_IMAGE|TVIF_PARAM))
			return;

		CProperty* pProperty = m_pBrowser->CurrentProperty();

		switch (m_tvCurrentItem.iImage)
		{
		case eBand:
		case eMenu:
		case ePopup:
		case eStatusBar:
		case eChildMenu:
			{
				IBand* pBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
				if (NULL == pBand)
				{
					MessageBox(IDS_ERR_FAILEDTOGETBANDOBJECT);
					return;
				}

				if (NULL == pProperty && -1 == nType)
					return;

				int nMemId;
				if (-1 == nType)
					nMemId = pProperty->MemId();
				else
					nMemId = nType;

				switch (nMemId)
				{
				case DISPID_BANDTYPE:
					{
						BandTypes btType;
						HRESULT hResult = pBand->get_Type(&btType);
						if (FAILED(hResult))
							return;

						int iImage = (int)GetBandImageType(btType);
						if (m_tvCurrentItem.iImage != iImage)
						{
							m_tvCurrentItem.mask |= TVIF_IMAGE|TVIF_SELECTEDIMAGE; 
							m_tvCurrentItem.iImage = m_tvCurrentItem.iSelectedImage = iImage;
							TreeView_SetItem(m_hWnd, &m_tvCurrentItem); 
						}
						if (m_pBrowser)
						{
							m_pBrowser->Clear();
							m_pBrowser->SetTypeInfo(pBand, this);
						}
					}
					break;

				case DISPID_CHILDBANDSTYLE:
					{
						BOOL bExpanded = FALSE;
						if (m_tvCurrentItem.state & TVIS_EXPANDED)
							bExpanded = TRUE;

						TreeView_Expand(m_hWnd, m_tvCurrentItem.hItem, TVE_COLLAPSE|TVE_COLLAPSERESET);

						if (bExpanded)
							TreeView_Expand(m_hWnd, m_tvCurrentItem.hItem, TVE_EXPAND);

						CRect rcItem;
						BOOL bResult = TreeView_GetItemRect(m_hWnd, m_hCurrentItem, &rcItem, FALSE);
						if (bResult)
							InvalidateRect(&rcItem, FALSE);
					}
					break;
				}
			}
			break;

		case eButton:
		case eDropDown:
		case eTextbox:
		case eCombobox:
		case eLabel:
		case eWindowList:
		case eSeparator:
		case eCustomTool:
		case eForm:
			{
				// Update caption
				ITool* pTool = reinterpret_cast<ITool*>(m_tvCurrentItem.lParam);
				if (NULL == pTool)
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOL);
					return;
				}

				if (pProperty)
				{
					switch (pProperty->MemId())
					{
					case DISPID_CONTROLTYPE:
						{
							ToolTypes ctTool;
							HRESULT hResult = pTool->get_ControlType(&ctTool);
							if (FAILED(hResult))
								return;

							int iImage = (int)GetToolImageType(ctTool);
							if (m_tvCurrentItem.iImage != iImage)
							{
								m_tvCurrentItem.mask |= TVIF_IMAGE|TVIF_SELECTEDIMAGE; 
								m_tvCurrentItem.iImage = m_tvCurrentItem.iSelectedImage = iImage;
								TreeView_SetItem(m_hWnd, &m_tvCurrentItem); 
							}
							if (m_pBrowser)
							{
								m_pBrowser->Clear();
								m_pBrowser->SetTypeInfo(pTool, this);
							}
						}
						break;
					}
				}
			}
			break;
		}
		Update();
		InvalidateRect(NULL, TRUE);
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}

}

//
// IterateCategoriesImpl
//

static void IterateCategoriesImpl (CatEntry* pEntry, void* pData)
{
	HTREEITEM hCurrentItem = ((CDesignerTree*)pData)->GetCurrentItem();
	if (NULL == pEntry->Category() || NULL == *pEntry->Category())
	{
		DDString strTemp;
		strTemp.LoadString(IDS_CAT_NONE);
		((CDesignerTree*)pData)->InsertNode (hCurrentItem, 
											 strTemp, 
											 CDesignerTree::eCategory, 
											 I_CHILDRENCALLBACK, 
											 (long)pEntry);
	}
	else
	{
		MAKE_TCHARPTR_FROMWIDE(szCategory, pEntry->Category());
		((CDesignerTree*)pData)->InsertNode (hCurrentItem, 
											 szCategory, 
											 CDesignerTree::eCategory, 
											 I_CHILDRENCALLBACK, 
											 (long)pEntry);
	}
}

//
// FillCategories
//

int CALLBACK CompareCategories(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	CatEntry* pCatEntry1 = (CatEntry*)lParam1;
	CatEntry* pCatEntry2 = (CatEntry*)lParam2;
	if (NULL == pCatEntry1->Category())
		return -1;
	if (NULL == pCatEntry2->Category())
		return 1;
	return wcscmp(pCatEntry1->Category(), pCatEntry2->Category());
}

void CDesignerTree::FillCategories(HTREEITEM hRoot)
{
	m_hCurrentItem = hRoot;
	m_pCategoryMgr->ClearABCreated();
	m_pCategoryMgr->Init();
	m_pCategoryMgr->IterateCategories(IterateCategoriesImpl, (void*)this);
	TVSORTCB theSort;
	theSort.hParent = m_hCurrentItem;
	theSort.lParam = NULL;
	theSort.lpfnCompare = CompareCategories;
	TreeView_SortChildrenCB(m_hWnd, &theSort, NULL);
}

//
// FillCategory
//

void CDesignerTree::FillCategory(HTREEITEM hRoot, CatEntry* pCatEntry)
{
	int nCount = pCatEntry->GetTools().GetSize();
	if (nCount < 1)
		return;

	HRESULT hResult;
	ITool* pTool;
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pTool = pCatEntry->GetTools().GetAt(nIndex);
		if (NULL == pTool)
			continue;

		ToolTypes ctTool;
		hResult = pTool->get_ControlType(&ctTool);
		if (FAILED(hResult))
			continue;

		pTool->AddRef();
		InsertNode (hRoot, LPSTR_TEXTCALLBACK, GetToolImageType(ctTool), 0, (long)pTool);
	}
}

//
// GetBandImageType
//

CDesignerTree::ImageTypes CDesignerTree::GetBandImageType(BandTypes btType)
{
	static ImageTypes itBands[] = {eBand, eMenu, ePopup, eStatusBar, eChildMenu};
	return itBands[btType];
}

CDesignerTree::ImageTypes CDesignerTree::GetToolImageType(ToolTypes ctTool)
{
	static ImageTypes itTools[] = {eButton, eDropDown, eCombobox, eTextbox, eLabel, eSeparator, eCustomTool, eForm, eWindowList};
	return itTools[ctTool];
}

//
// FillBands
//

void CDesignerTree::FillBands(HTREEITEM hRoot)
{
	short nBandCount;
	IBands* pBands;
	HRESULT hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
	if (FAILED(hResult))
	{
		MessageBox(IDS_ERR_FAILEDTOGETBANDSOBJECT);
		return;
	}

	hResult = pBands->Count(&nBandCount);
	if (FAILED(hResult))
	{
		pBands->Release();
		return;
	}

	HTREEITEM hBand;
	IBand* pBand;
	VARIANT vBand;
	vBand.vt = VT_I4;

	try
	{
		for (vBand.lVal = 0; vBand.lVal < nBandCount; vBand.lVal++)
		{
			hResult = pBands->Item(&vBand, (Band**)&pBand);
			if (FAILED(hResult))
				continue;

			BandTypes btType;
			hResult = pBand->get_Type(&btType);
			if (FAILED(hResult))
			{
				pBand->Release();
				continue;
			}

			hBand = InsertNode (hRoot, 
								LPSTR_TEXTCALLBACK, 
								GetBandImageType(btType),
								I_CHILDRENCALLBACK, 
								(long)pBand);
			if (NULL == hBand)
			{
				pBand->Release();
				continue;
			}
		}
		if (!GetGlobals().m_thePreferences.GetConfirmSuppress(IDS_SORTBANDS))
			TreeView_SortChildren(m_hWnd, hRoot, 0);
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	pBands->Release();
}

//
// FillChildBands
//

void CDesignerTree::FillChildBands(IChildBands* pChildBands, HTREEITEM hChildBands)
{
	short nBandCount;
	HRESULT hResult = pChildBands->Count(&nBandCount);
	if (FAILED(hResult))
		return;

	HTREEITEM hChildBand;
	IBand* pChildBand;
	VARIANT vChildBand;
	vChildBand.vt = VT_I4;
	try
	{
		for (vChildBand.lVal = 0; vChildBand.lVal < nBandCount; vChildBand.lVal++)
		{
			hResult = pChildBands->Item(&vChildBand, (Band**)&pChildBand);
			if (FAILED(hResult))
				continue;

			hChildBand = InsertNode (hChildBands, 
									 LPSTR_TEXTCALLBACK, 
									 eChildBand, 
									 I_CHILDRENCALLBACK, 
									 (long)pChildBand);
			if (NULL == hChildBand)
			{
				pChildBand->Release();
				continue;
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
// FillTools
//

void CDesignerTree::FillTools(ITools* pTools, HTREEITEM hBand)
{
	HTREEITEM hTool;
	VARIANT vTool;
	ITool* pTool;
	short nToolCount;
	long nToolId;

	vTool.vt = VT_I4;

	HRESULT hResult = pTools->Count(&nToolCount);
	if (FAILED(hResult))
		return;

	try
	{
		for (vTool.lVal = 0; vTool.lVal < nToolCount; vTool.lVal++)
		{
			hResult = pTools->Item(&vTool, (Tool**)&pTool);
			if (FAILED(hResult))
				continue;

			assert(pTool);
			hResult = pTool->get_ID(&nToolId);
			if (FAILED(hResult))
			{
				pTool->Release();
				continue;
			}

			ToolTypes ttTool;
			hResult = pTool->get_ControlType(&ttTool);
			if (FAILED(hResult))
			{
				pTool->Release();
				continue;
			}

			hTool = InsertNode (hBand, 
								LPSTR_TEXTCALLBACK,
								GetToolImageType(ttTool), 
								-2, 
								(long)pTool);
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		CONTINUE
	}
}

//
// Init
//

BOOL CDesignerTree::Init(IActiveBar2* pActiveBar, CBrowser* pBrowser, CategoryMgr* pCategoryMgr)
{
	try
	{
		m_pActiveBar = pActiveBar;
		IBarPrivate* pPrivate;
		HRESULT hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
		if (SUCCEEDED(hResult))
		{
			pPrivate->RegisterBandChange(this);
			pPrivate->Release();
		}

		if (pBrowser)
		{
			m_pBrowser = pBrowser;
			m_pBrowser->Clear();
		}

		BOOL bResult = LoadImages(IDR_DESIGNER);
		if (!bResult)
			return FALSE;

		if (pCategoryMgr)
		{
			m_pCategoryMgr = pCategoryMgr;
			m_pCategoryMgr->Register(this);
		}

		if (NULL == m_pTreeCtrlTarget)
		{
			m_pTreeCtrlTarget = new CDesignerTarget(eBrowser == m_eTreeType ? CToolDataObject::eDesignerDragDropId : CToolDataObject::eLibraryDragDropId, *this);
			assert(m_pTreeCtrlTarget);
			if (NULL == m_pTreeCtrlTarget)
				return FALSE;
		
			GetGlobals().GetDragDropMgr()->RegisterDragDrop((OLE_HANDLE)m_hWnd, (LPUNKNOWN)m_pTreeCtrlTarget);
		}
		SetRoot();
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
// SetRoot
//

void CDesignerTree::SetRoot()
{
	m_pActiveBar->AddRef();
	HTREEITEM hTree = InsertNode (TVI_ROOT, 
								  m_szMsgTitle, 
								  eRoot, 
								  I_CHILDRENCALLBACK, 
								  (long)m_pActiveBar);
	if (NULL == hTree)
		return;

	BOOL bResult = TreeView_Expand(m_hWnd, hTree, TVE_EXPAND);
	if (bResult)
		bResult = TreeView_Select(m_hWnd, hTree, TVGN_CARET); 
}

//
// OnCommand
//

void CDesignerTree::OnCommand(WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (LOWORD(wParam))
		{
		case ID_INSERTBAND_NORMAL:
			OnCreateBand(ddBTNormal);
			break;

		case ID_INSERTBAND_MENU:
			OnCreateBand(ddBTMenuBar);
			break;

		case ID_INSERTBAND_CHILDMENU:
			OnCreateBand(ddBTChildMenuBar);
			break;

		case ID_INSERTBAND_POPUP:
			OnCreateBand(ddBTPopup);
			break;
		
		case ID_INSERTBAND_STATUSBAR:
			OnCreateBand(ddBTStatusBar);
			break;

		case ID_DELETEBAND:
			OnDeleteBand();
			break;
		
		case ID_CREATECHILDBAND:
			OnCreateChildBand();
			break;

		case ID_DELETECHILDBAND:
			OnDeleteChildBand();
			break;

		case ID_INSERTTOOL_BUTTON:
			OnCreateTool(ddTTButton);
			break;

		case ID_INSERTTOOL_DROPDOWNBUTTON:
			OnCreateTool(ddTTButtonDropDown);
			break;

		case ID_INSERTTOOL_TEXTBOX:
			OnCreateTool(ddTTEdit);
			break;

		case ID_INSERTTOOL_COMBOBOX:
			OnCreateTool(ddTTCombobox);
			break;

		case ID_INSERTTOOL_LABEL:
			OnCreateTool(ddTTLabel);
			break;

		case ID_INSERTTOOL_SEPARATOR:
			OnCreateTool(ddTTSeparator);
			break;

		case ID_INSERTTOOL_WINDOWLIST:
			OnCreateTool(ddTTWindowList);
			break;

		case ID_INSERTTOOL_ACTIVEXCONTROL:
			OnCreateTool(ddTTControl);
			break;

		case ID_INSERTTOOL_ACTIVEXFORM:
			OnCreateTool(ddTTForm);
			break;

		case ID_DELETETOOL:
			OnDeleteTool();
			break;

		case ID_EDITBITMAP:
			try
			{
				ITool* pTool = reinterpret_cast<ITool*>(m_tvCurrentItem.lParam);
				if (NULL == pTool)
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOL);
					return;
				}
				EnableWindow(::GetParent(m_hWnd), FALSE);
				CIconEditor dlgIconEditor(pTool);
				if (IDOK == dlgIconEditor.DoModal())
				{
					ObjectChanged(m_tvCurrentItem.iImage);
					if (eBrowser == m_eTreeType)
						Update();
				}
				EnableWindow(::GetParent(m_hWnd), TRUE);
			}
			CATCH 
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
			break;

		case ID_APPLYALLTOOL:
			OnApplyAll();
			break;
		
		case ID_FILENEW:
			OnNewLayout();
			break;

		case ID_FILEOPEN:
			OnOpenLayout();
			break;

		case ID_FILESAVE:
			OnSaveLayout();
			break;

		case ID_FILE_SAVEAS:
			OnSaveAsLayout();
			break;

		case ID_CATEGORY_NEW:
			OnCreateCategory();
			break;

		case ID_CATEGORY_DELETE:
			OnDeleteCategory();
			break;
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnCreateBand
//

void CDesignerTree::OnCreateBand(UINT nType)
{
	try
	{
		IBands* pBands;
		HRESULT hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOGETBANDSOBJECT);
			return;
		}

		short nBandCount;
		hResult = pBands->Count(&nBandCount);
		if (FAILED(hResult))
		{
			pBands->Release();
			MessageBox(IDS_ERR_FAILEDTOGETBANDCOUNT);
			return;
		}

		DDString strBandName;
		VARIANT  vBand;
		IBand*   pBand;
		BOOL	 bFindingName = TRUE;
		BSTR	 bstrName;

		vBand.vt = VT_I4;
		while (bFindingName)
		{
			strBandName.Format(IDS_BANDNAME, ++nBandCount);

			for (vBand.lVal = 0; vBand.lVal < nBandCount; vBand.lVal++)
			{
				hResult = pBands->Item(&vBand, (Band**)&pBand);
				if (FAILED(hResult))
					continue;

				hResult = pBand->get_Name(&bstrName);
				pBand->Release();
				if (FAILED(hResult))
					continue;

				MAKE_TCHARPTR_FROMWIDE(szBandName, bstrName);

				//
				// Free the name
				//

				SysFreeString(bstrName);
				bstrName = NULL;

				if (0 == _tcscmp(szBandName, strBandName))
					break;
			}
			if (vBand.lVal == nBandCount)
				bFindingName = FALSE;
		}

		//
		// Creating and adding the new band to ActiveBar
		//

		BSTR bstrBand = strBandName.AllocSysString();
		if (NULL == bstrBand)
			return;

		hResult = pBands->Add(bstrBand, reinterpret_cast<Band**>(&pBand));
		pBands->Release();
		SysFreeString(bstrBand);
		bstrBand = NULL;
		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOCREATENEWBAND);
			return;
		}

		if (-1 != nType)
		{
			hResult = pBand->put_Type((BandTypes)nType);
			if (FAILED(hResult))
			{
				DDString strType;
				strType.LoadString(IDS_BANDTYPES + nType);
				DDString strMsg;
				strMsg.Format(IDS_ERR_FAILEDTOSETBANDTYPE, strType);
				MessageBox(strMsg);
			}
		}
		pBand->Release();
		if (eBrowser == m_eTreeType)
			Update();
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnCreateChildBand
//

void CDesignerTree::OnCreateChildBand()
{
	try
	{
		IChildBands* pChildBands = reinterpret_cast<IChildBands*>(m_tvCurrentItem.lParam);
		if (NULL == pChildBands)
		{
			MessageBox(IDS_ERR_FAILEDTOGETCHILDBANDSOBJECT);
			return;
		}

		short nCount;
		HRESULT hResult = pChildBands->Count(&nCount);
		if (FAILED(hResult))
		{
			pChildBands->Release();
			MessageBox(IDS_ERR_FAILEDTOGETCHILDBANDCOUNT);
			return;
		}

		DDString strChildBand;
		strChildBand.Format(IDS_CHILDBANDNAME, ++nCount);
		BSTR bstrChildBand = strChildBand.AllocSysString();
		if (NULL == bstrChildBand)
			return;

		IBand* pChildBand;
		hResult = pChildBands->Add(bstrChildBand, (Band**)&pChildBand);
		SysFreeString(bstrChildBand);
		bstrChildBand = NULL;
		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOCREATEANEWCHILDBAND);
			return;
		}

		if (m_tvCurrentItem.state & TVIS_EXPANDED)
		{
			HTREEITEM hChildBand = InsertNode (m_hCurrentItem, 
											   LPSTR_TEXTCALLBACK, 
											   eChildBand, 
											   I_CHILDRENCALLBACK,
											   (long)pChildBand);
			if (NULL == hChildBand)
			{
				MessageBox(IDS_ERR_FAILEDTOCREATEANEWCHILDBAND);
				return;
			}
		}
		else
			pChildBand->Release();

		if (eBrowser == m_eTreeType)
			Update();

		CRect rcItem;
		BOOL bResult = TreeView_GetItemRect(m_hWnd, m_hCurrentItem, &rcItem, FALSE);
		if (bResult)
			InvalidateRect(&rcItem, FALSE);
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION (__FILE__, __LINE__)
	}
}

//
// OnCreateTool
//

void CDesignerTree::OnCreateTool(ToolTypes ttTool)
{
	try
	{
		ITools* pTools = NULL;
		ITool* pTool;
		DDString strTool;
		int nToolId = FindLastToolId(m_pActiveBar) + 1;
		strTool.Format(IDS_TOOLNAME, nToolId);

		switch (m_tvCurrentItem.iImage)
		{
		case eBand:
		case eMenu:
		case eChildMenu:
		case ePopup:
		case eStatusBar:
		case eChildBand:
			{
				IBand* pBand = NULL;
				pBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
				if (NULL == pBand)
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOL);
					return;
				}
				
				HRESULT hResult = pBand->get_Tools((Tools**)&pTools);
				if (FAILED(hResult))
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOLS);
					return;
				}

				//
				// Adding the Tool to the Band's Tools collection
				//

				BSTR bstrTool = strTool.AllocSysString();
				if (NULL == bstrTool)
				{
					pTools->Release();
					return;
				}

				hResult = pTools->Add(nToolId, bstrTool, (Tool**)&pTool);
				pTools->Release();
				if (FAILED(hResult))
				{
					SysFreeString(bstrTool);
					MessageBox(IDS_ERR_FAILEDTOCREATEANEWTOOL);
					return;
				}


				if (ddTTSeparator != ttTool && ddTTWindowList != ttTool)
				{
					hResult = pTool->put_Caption(bstrTool);
					SysFreeString(bstrTool);
					bstrTool = NULL;
					if (FAILED(hResult))
					{
						pTool->Release();
						MessageBox(IDS_ERR_FAILEDTOSETTHETOOLSCAPTION);
						return;
					}
				}
				else
				{
					SysFreeString(bstrTool);
					bstrTool = NULL;
				}

				if (ddTTButton != ttTool)
					pTool->put_ControlType((ToolTypes)ttTool);

				//
				// Adding the Tool to the Main Tools collection
				//

				hResult = m_pActiveBar->get_Tools((Tools**)&pTools);
				if (FAILED(hResult))
				{
					pTool->Release();
					MessageBox(IDS_ERR_FAILEDTOGETTOOLS);
					return;
				}

				//
				// Inserting a clone to the tools to the Main Tools Collection
				//

				hResult = pTools->InsertTool(-1, pTool, VARIANT_TRUE);
				pTools->Release();
				if (FAILED(hResult))
				{
					pTool->Release();
					MessageBox(IDS_ERR_FAILEDTOCREATEANEWTOOL);
					return;
				}
				m_pCategoryMgr->AddTool(pTool);

				BOOL bResult;
				if (m_tvCurrentItem.state & TVIS_EXPANDED)
				{
					HTREEITEM hTool = InsertNode (m_hCurrentItem, 
												  LPSTR_TEXTCALLBACK, 
												  (int)GetToolImageType(ttTool), 
												  -2,
												  (long)pTool);
					if (NULL == hTool)
						MessageBox(IDS_ERR_FAILEDTOCREATEANEWTOOL);
				}
				else
				{
					pTool->Release();
					SetItemState(m_hCurrentItem, 0, TVIS_EXPANDEDONCE);
					bResult = TreeView_Expand(m_hWnd, m_hCurrentItem, TVE_TOGGLE);
				}
				CRect rcItem;
				bResult = TreeView_GetItemRect(m_hWnd, m_hCurrentItem, &rcItem, FALSE);
				if (bResult)
					InvalidateRect(&rcItem, FALSE);
			}
			break;

		default:
			{
				CatEntry* pCatEntry = reinterpret_cast<CatEntry*>(m_tvCurrentItem.lParam);
				if (NULL == pCatEntry)
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOLGATEGORY);
					return;
				}

				ITool* pTool = m_pCategoryMgr->CreateTool(pCatEntry->Category(), TRUE, ttTool);
				if (NULL == pTool)
				{
					MessageBox(IDS_ERR_FAILEDTOCREATEANEWTOOL);
					return;
				}

				TV_ITEM tvItem1;
				tvItem1.mask = TVIF_STATE; 
				tvItem1.hItem = m_hCurrentItem;
				if (TreeView_GetItem(m_hWnd, &tvItem1) && !(TVIS_EXPANDED & tvItem1.state))
					TreeView_Expand(m_hWnd, m_hCurrentItem, TVE_EXPAND);
			}
			break;
		}

		if (eBrowser == m_eTreeType)
			Update();
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnDeleteBand
//

void CDesignerTree::OnDeleteBand()
{
	try
	{
		CDlgConfirm dlgConfirm(IDS_DELETEBANDCONFIRM);
		UINT nResult = dlgConfirm.Show(m_hWnd);
		if (IDCANCEL == nResult)
			return;

		IBand* pBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
		if (NULL == pBand)
		{
			MessageBox(IDS_ERR_FAILEDTOGETBANDOBJECT);
			return;
		}
		
		IBands* pBands;
		HRESULT hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOGETBANDSOBJECT);
			return;
		}

		short nBandCount;
		hResult = pBands->Count(&nBandCount);
		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOGETBANDCOUNT);
			pBands->Release();
			return;
		}

		IBand* pBandLoop;
		VARIANT vBand;
		vBand.vt = VT_I4;
		for (vBand.lVal = 0; vBand.lVal < nBandCount; vBand.lVal++)
		{
			hResult = pBands->Item(&vBand, (Band**)&pBandLoop);
			if (FAILED(hResult))
				continue;

			if (pBandLoop == pBand)
			{
				pBandLoop->Release();
				hResult = pBands->Remove(&vBand);
				if (FAILED(hResult))
				{
					if (eBrowser == m_eTreeType)
						Update();
				}
				break;
			}
			pBandLoop->Release();
		}
		pBands->Release();
		if (eBrowser == m_eTreeType)
			Update();
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnApplyAll
//

void CDesignerTree::OnApplyAll()
{
	try
	{
		ITool* pTool = reinterpret_cast<ITool*>(m_tvCurrentItem.lParam);
		if (NULL == pTool)
		{
			MessageBox(IDS_ERR_FAILEDTOGETTOOL);
			return;
		}
		long nToolId;
		if (SUCCEEDED(pTool->get_ID(&nToolId)))
		{
			VisitBandTools(m_pActiveBar, nToolId, pTool, UpdateBarTool);

			// Added this because of CR 1306, I think that took it out because of
			// another CR.

			VisitBarTools(m_pActiveBar, nToolId, pTool, UpdateBarTool);

			if (eBrowser == m_eTreeType)
				Update();
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnDeleteChildBand
//

void CDesignerTree::OnDeleteChildBand()
{
	try
	{
		CDlgConfirm dlgConfirm(IDS_DELETECHILDBANDCONFIRM);
		UINT nResult = dlgConfirm.Show(m_hWnd);
		if (IDCANCEL == nResult)
			return;

		IBand* pChildBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
		if (NULL == pChildBand)
		{
			MessageBox(IDS_ERR_FAILEDTOGETCHILDBANDOBJECT);
			return;
		}

		TV_ITEM tvParent;
		if (!GetParent(m_hCurrentItem, tvParent, TVIF_IMAGE|TVIF_HANDLE|TVIF_PARAM))
		{
			MessageBox(IDS_ERR_FAILEDTOGETCHILDBANDSOBJECT);
			return;
		}

		IChildBands* pChildBands = reinterpret_cast<IChildBands*>(tvParent.lParam);

		short nCount;
		HRESULT hResult = pChildBands->Count(&nCount);
		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOGETCHILDBANDCOUNT);
			return;
		}

		IBand* pBandLoop;
		VARIANT vItem;
		vItem.vt = VT_I4;
		for (vItem.lVal = 0; vItem.lVal < nCount; vItem.lVal++)
		{
			hResult = pChildBands->Item(&vItem, (Band**)&pBandLoop);
			if (FAILED(hResult))
				continue;

			if (pBandLoop == pChildBand)
			{
				pBandLoop->Release();
				pChildBands->Remove(&vItem);
				TreeView_DeleteItem(m_hWnd, m_hCurrentItem);
				if (eBrowser == m_eTreeType)
					Update();
				break;
			}
			pBandLoop->Release();
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// DeleteTool
//

void CDesignerTree::DeleteTool(TVITEM tvParent, TVITEM tvTool, ITool* pTool)
{
	switch (tvParent.iImage)
	{
	case eBand:
	case eMenu:
	case ePopup:
	case eStatusBar:
	case eChildBand:
	case eChildMenu:
		{
			ITools* pTools = NULL;
			IBand* pBand = reinterpret_cast<IBand*>(tvParent.lParam);
			if (NULL == pBand)
			{
				MessageBox(IDS_ERR_FAILEDTOGETTOOL);
				return;
			}

			HRESULT hResult = pBand->get_Tools((Tools**)&pTools);
			if (FAILED(hResult))
			{
				MessageBox(IDS_ERR_FAILEDTOGETTOOLS);
				return;
			}

			short nCount;
			hResult = pTools->Count(&nCount);
			if (FAILED(hResult))
			{
				MessageBox(IDS_ERR_FAILEDTOGETTOOLCOUNT);
				return;
			}

			ITool* pToolLoop;
			VARIANT vItem;
			vItem.vt = VT_I4;
			for (vItem.lVal = 0; vItem.lVal < nCount; vItem.lVal++)
			{
				hResult = pTools->Item(&vItem, (Tool**)&pToolLoop);
				if (FAILED(hResult))
					continue;

				if (pToolLoop == pTool)
				{
					TreeView_DeleteItem(m_hWnd, tvTool.hItem);
					pTool->Release();
					pTools->Remove(&vItem);
					if (eBrowser == m_eTreeType)
						Update();
					break;
				}
				pToolLoop->Release();
			}
			pTools->Release();
		}
		break;

	case eCategory: 
		{
			CatEntry* pCatEntry = reinterpret_cast<CatEntry*>(tvParent.lParam);
			if (NULL == pCatEntry)
			{
				MessageBox(IDS_ERR_FAILEDTODELETETOOL);
				return;
			}

			if (!m_pCategoryMgr->DeleteTool(pCatEntry->Category(), pTool))
			{
				MessageBox(IDS_ERR_FAILEDTODELETETOOL);
				return;
			}
		}
		break;

	default:
		assert(FALSE);
		break;
	}
}

//
// OnDeleteTool
//

void CDesignerTree::OnDeleteTool()
{
	try
	{
		CDlgConfirm dlgConfirm(IDS_DELETETOOLCONFIRM);
		UINT nResult = dlgConfirm.Show(m_hWnd);
		if (IDCANCEL == nResult)
			return;

		TV_ITEM tvParent;
		if (m_bMultSelectActive)
		{
			TVITEM tvTool;
			TVITEM tvDelete;
			tvTool.mask = TVIF_IMAGE|TVIF_PARAM|TVIF_HANDLE; 
			tvTool.hItem = GetFirstSelectedItem();
			while (tvTool.hItem)
			{
				tvDelete = tvTool;
				tvTool.hItem = GetNextSelectedItem(tvTool.hItem);
				if (TreeView_GetItem(m_hWnd, &tvDelete))
				{
					switch (tvDelete.iImage)
					{
					case eButton:
					case eDropDown:
					case eCombobox:
					case eTextbox:
					case eLabel:
					case eSeparator:
					case eWindowList:
					case eCustomTool:
					case eForm:
						{
							if (!GetParent(tvDelete.hItem, tvParent, TVIF_IMAGE|TVIF_HANDLE|TVIF_PARAM))
							{
								MessageBox(IDS_ERR_FAILEDTOGETCHILDBANDOBJECT);
								return;
							}
							DeleteTool(tvParent, tvDelete, reinterpret_cast<ITool*>(tvDelete.lParam));
						}
						break;
					}
				}
			}
		}
		else
		{
			ITool* pTool = reinterpret_cast<ITool*>(m_tvCurrentItem.lParam);
			if (NULL == pTool)
			{
				MessageBox(IDS_ERR_FAILEDTOGETTOOL);
				return;
			}

			if (!GetParent(m_hCurrentItem, tvParent, TVIF_IMAGE|TVIF_HANDLE|TVIF_PARAM))
			{
				MessageBox(IDS_ERR_FAILEDTOGETCHILDBANDOBJECT);
				return;
			}
			DeleteTool(tvParent, m_tvCurrentItem, pTool);
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	if (eBrowser == m_eTreeType)
		Update();
}


//
// OnSaveLayout
//

void CDesignerTree::OnSaveLayout()
{
	try
	{
		// Get file name

		if (NULL == m_szCurFileName || lstrlen(m_szCurFileName) < 1)
		{
			OnSaveAsLayout(TRUE);
			return;
		}

		MAKE_WIDEPTR_FROMTCHAR(wFileName, m_szCurFileName);
		BSTR bstrFileName = SysAllocString(wFileName);
		if (NULL == bstrFileName)
		{
			MessageBox(IDS_ERR_FAILEDTOSAVELAYOUT);
			return;
		}

		VARIANT vEmpty;
		vEmpty.vt = VT_EMPTY;
		HRESULT hResult = m_pActiveBar->Save(L"", bstrFileName, ddSOFile, &vEmpty);

		SysFreeString(bstrFileName);

		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOSAVELAYOUT);
			return;
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		CONTINUE
	}
}

//
// OnSaveAsLayout
//

void CDesignerTree::OnSaveAsLayout(BOOL bSaveFileName)
{
	try
	{
		// Get file name

		DDString strBuffer;
		strBuffer.LoadString(IDS_SAVELAYOUTTYPE);
		DDString strFileType;
		strFileType.LoadString(IDS_LAYOUTTYPE);
		CFileDialog dlgFile(IDS_SAVELAYOUT,
							FALSE,
							strFileType,
							strFileType,
							OFN_OVERWRITEPROMPT|OFN_HIDEREADONLY,
							strBuffer,
							m_hWnd);

		if (0 == dlgFile.DoModal())
		{
			// It was cancelled or there was an error
			return;
		}

		// Now save to file
		MAKE_WIDEPTR_FROMTCHAR(wFileName, dlgFile.GetFileName());
		BSTR bstrFileName = SysAllocString(wFileName);
		VARIANT vEmpty;
		vEmpty.vt = VT_EMPTY;
		HRESULT hResult = m_pActiveBar->Save(L"", bstrFileName, ddSOFile, &vEmpty);

		SysFreeString(bstrFileName);

		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOSAVELAYOUT);
			return;
		}
		if (bSaveFileName)
			_tcscpy(m_szCurFileName, dlgFile.GetFileName());
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		CONTINUE
	}
}

// 
// OnNewLayout
//

void CDesignerTree::OnNewLayout()
{
	IBands* pBands;
	HRESULT hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
	if (FAILED(hResult))
	{
		MessageBox(IDS_ERR_FAILEDTOGETBANDSOBJECT);
		return;
	}

	hResult = pBands->RemoveAll();
	if (FAILED(hResult))
	{
		MessageBox(IDS_ERR_FAILEDTOGETBANDSOBJECT);
		return;
	}

	ITools* pTools;
	hResult = m_pActiveBar->get_Tools((Tools**)&pTools);
	if (FAILED(hResult))
	{
		MessageBox(IDS_ERR_FAILEDTOGETTOOLS);
		return;
	}
	
	hResult = pTools->RemoveAll();
	if (FAILED(hResult))
	{
		MessageBox(IDS_ERR_FAILEDTOGETTOOLS);
		return;
	}
	
	hResult = m_pActiveBar->RecalcLayout();

	TreeView_DeleteAllItems(m_hWnd);
	if (m_pBrowser)
		m_pBrowser->SendMessage(LB_RESETCONTENT);
	SetRoot();
	Update();
}

//
// OnOpenLayout
//

void CDesignerTree::OnOpenLayout()
{
	try
	{
		// Get file name

		DDString strBuffer;
		strBuffer.LoadString(IDS_LAYOUTTYPE);
		DDString strFileType;
		strFileType.LoadString(IDS_LAYOUTFILETYPE);
		CFileDialog dlgFile(IDS_OPENLAYOUT,
							TRUE,
							strFileType,
							strFileType,
							OFN_OVERWRITEPROMPT|OFN_PATHMUSTEXIST|OFN_HIDEREADONLY,
							strBuffer,
							m_hWnd);

		dlgFile.SetFileName(m_szCurFileName);
		if (0 == dlgFile.DoModal())
		{
			// It was cancelled or there was an error
			return;
		}

		//
		// Get the File Extention
		//

		TCHAR* szEnd = _tcsrchr(dlgFile.GetFileName(), static_cast<TCHAR>('.'));
		szEnd++;

		HCURSOR hCursorPrev = GetCursor();
		HCURSOR hCursorWait = LoadCursor(NULL, MAKEINTRESOURCE(IDC_WAIT));
		if (hCursorWait)
			SetCursor(hCursorWait);

		TreeView_DeleteAllItems(m_hWnd);
		if (m_pBrowser)
			m_pBrowser->SendMessage(LB_RESETCONTENT);

		switch (dlgFile.GetFilterIndex())
		{
		case ActiveBar1:
			{
				if (0 == lstrcmpi(szEnd, _T("tb")))
					OnLoadAB10(dlgFile.GetFileName());
				else if (0 == lstrcmpi(szEnd, _T("tb2")))
					OnLoadAB20(dlgFile.GetFileName());
				else
					MessageBox(IDS_ERR_FAILEDTOLOADLAYOUT);
			}
			break;

		case ActiveBar2:
			{
				if (0 == lstrcmpi(szEnd, _T("tb2")))
					OnLoadAB20(dlgFile.GetFileName());
				else if (0 == lstrcmpi(szEnd, _T("tb")))
					OnLoadAB10(dlgFile.GetFileName());
				else
					MessageBox(IDS_ERR_FAILEDTOLOADLAYOUT);
			}
			break;

		default:
			assert(FALSE);
			break;
		}
		SetRoot();
		if (hCursorPrev)
			SetCursor(hCursorPrev);
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		CONTINUE
	}
}

//
// OnLoadAB20
//

void CDesignerTree::OnLoadAB20(LPCTSTR szFileName)
{
	try
	{
		lstrcpy(m_szCurFileName, szFileName);
		MAKE_WIDEPTR_FROMTCHAR(wFileName, m_szCurFileName);
		VARIANT vFileName;
		vFileName.vt = VT_BSTR;
		vFileName.bstrVal = SysAllocString(wFileName);

		HRESULT hResult = m_pActiveBar->Load(L"", &vFileName, ddSOFile);
		VariantClear(&vFileName);
		if (FAILED(hResult))
		{
			MessageBox(IDS_ERR_FAILEDTOLOADLAYOUT);
			return;
		}
		else
			Update();
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnLoadAB10
//

void CDesignerTree::OnLoadAB10(LPCTSTR szFileName)
{
	try
	{
		HANDLE hFile = CreateFile(szFileName, 
								  GENERIC_READ, 
								  FILE_SHARE_READ, 
								  NULL, 
								  OPEN_EXISTING, 
								  FILE_ATTRIBUTE_NORMAL, 
								  NULL);

		if (INVALID_HANDLE_VALUE == hFile)
		{
			LPVOID lpMsgBuf;
			FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
						  NULL,
						  GetLastError(),
						  MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
						  (LPTSTR) &lpMsgBuf,    
						  0,    
						  NULL );
			MessageBox((char*)lpMsgBuf, MB_OK|MB_ICONINFORMATION);
			// Free the buffer.
			LocalFree(lpMsgBuf);
			return;
		}

		FileStream* pFileStream = new FileStream;
		if (NULL == pFileStream)
		{
			CloseHandle((HANDLE)hFile);
			MessageBox(IDS_ERR_FAILEDTOLOADLAYOUT);
			return;
		}

		pFileStream->SetHandle(hFile, TRUE);
		AB10 theActiveBar10;
		if (theActiveBar10.ReadState(pFileStream))
		{
			if (!theActiveBar10.Convert(m_pActiveBar))
				MessageBox(IDS_ERR_FAILEDTOCONVERT);
			else
				Update();
		}
		pFileStream->Release();
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnCreateCategory
//

void CDesignerTree::OnCreateCategory()
{
	try
	{
		int nCount = m_pCategoryMgr->CategoryCount();
		DDString strTemp;
		strTemp.Format(IDS_CATEGORYCREATE, nCount);
		BSTR bstrTemp = strTemp.AllocSysString();
		if (NULL == bstrTemp)
		{
			MessageBox(IDS_ERR_FAILEDTOCREATECATEGORY);
			return;
		}
		CatEntry* pEntry = m_pCategoryMgr->FindEntry(bstrTemp);
		while (pEntry)
		{
			SysFreeString(bstrTemp);
			nCount++;
			strTemp.Format(IDS_CATEGORYCREATE, nCount);
			bstrTemp = strTemp.AllocSysString();
			if (NULL == bstrTemp)
			{
				MessageBox(IDS_ERR_FAILEDTOCREATECATEGORY);
				return;
			}
			pEntry = m_pCategoryMgr->FindEntry(bstrTemp);
		}

		pEntry = m_pCategoryMgr->CreateEntry(bstrTemp, CatEntry::eDesigner);
		SysFreeString(bstrTemp);
		if (NULL == pEntry)
		{
			MessageBox(IDS_ERR_FAILEDTOCREATECATEGORY);
			return;
		}
		Update();
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnDeleteCategory
//

void CDesignerTree::OnDeleteCategory()
{
	try
	{
		CDlgConfirm dlgConfirm(IDS_DELETECATEGORYCONFIRM);
		UINT nResult = dlgConfirm.Show(m_hWnd);
		if (IDCANCEL == nResult)
			return;

		if (NULL == m_hCurrentItem)
		{
			MessageBox(IDS_ERR_FAILEDTODELETECATEGORY);
			return;
		}

		TV_ITEM tvItem;
		tvItem.mask = TVIF_IMAGE|TVIF_PARAM; 
		tvItem.hItem = m_hCurrentItem;
		if (TreeView_GetItem(m_hWnd, &tvItem))
		{
			CatEntry* pCatEntry = reinterpret_cast<CatEntry*>(tvItem.lParam);
			if (NULL == pCatEntry)
			{
				MessageBox(IDS_ERR_FAILEDTODELETECATEGORY);
				return;
			}

			if (!m_pCategoryMgr->DeleteEntry(pCatEntry->Category()))
			{
				MessageBox(IDS_ERR_FAILEDTODELETECATEGORY);
				return;
			}
			Update();
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnNotify
//

LRESULT CDesignerTree::OnNotify(LPNMHDR pNMHdr)
{
	try
	{
		switch (pNMHdr->code)
		{
		case TVN_SELCHANGED:
			OnSelChanged((NM_TREEVIEW*)pNMHdr);
			break;

		case TVN_ITEMEXPANDING:
			OnExpanding((NM_TREEVIEW*)pNMHdr);
			break;
		
		case TVN_ITEMEXPANDED:
			OnExpanded((NM_TREEVIEW*)pNMHdr);
			break;
		
		case NM_DBLCLK:
			OnDoubleClick();
			break;

		case NM_RCLICK:
			OnRightClick();
			break;

		case TVN_KEYDOWN:
			{
				LRESULT lResult = OnKeyDown((LPNMTVKEYDOWN)pNMHdr);
				if (0 != lResult)
					return lResult;
			}
			break;

		case TVN_DELETEITEM:
			OnDeleteItem((NM_TREEVIEW*)pNMHdr);
			break;
		
		case TVN_BEGINDRAG:
			OnBeginDrag((NM_TREEVIEW*)pNMHdr);
			break;

		case TVN_GETDISPINFO:
			OnGetDisplayInfo((TV_DISPINFO*)pNMHdr);
			break;

		case TVN_BEGINLABELEDIT:
			return OnBeginEditLabel((TV_DISPINFO*)pNMHdr);

		case TVN_ENDLABELEDIT:
			return OnEndEditLabel((TV_DISPINFO*)pNMHdr);
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return FALSE;
}

//
// OnExpanded
//

void CDesignerTree::OnExpanded(NM_TREEVIEW* pTreeView)
{
	try
	{
		if (TVE_COLLAPSE & pTreeView->action)
			DeleteChildren(pTreeView->itemNew.hItem);
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnExpanding
//

void CDesignerTree::OnExpanding(NM_TREEVIEW* pTreeView)
{
	try
	{
		if (TVE_EXPAND & pTreeView->action)
		{
			TVITEM tvItem;
			tvItem.mask = TVIF_IMAGE|TVIF_PARAM;
			tvItem.hItem = pTreeView->itemNew.hItem;
			BOOL bResult = TreeView_GetItem(m_hWnd, &tvItem);
			if (bResult)
			{
				switch (tvItem.iImage)
				{
				case eRoot:
					{
						m_pActiveBar->AddRef();
						InsertNode (pTreeView->itemNew.hItem, 
								    _T("Categories"), 
									eCategories, 
									I_CHILDRENCALLBACK, 
									(long)m_pActiveBar);

						m_pActiveBar->AddRef();
						InsertNode (pTreeView->itemNew.hItem, 
								    _T("Bands"), 
									eBands, 
									I_CHILDRENCALLBACK, 
									(long)m_pActiveBar);
					}
					break;

				case eBands:
					FillBands(pTreeView->itemNew.hItem);
					break;

				case eCategories:
					FillCategories(pTreeView->itemNew.hItem);
					break;

				case eCategory:
					FillCategory(pTreeView->itemNew.hItem, reinterpret_cast<CatEntry*>(tvItem.lParam));
					break;

				case eBand:
				case eMenu:
				case eChildMenu:
				case ePopup:
				case eStatusBar:
					{
						try
						{
							IBand* pBand = reinterpret_cast<IBand*>(tvItem.lParam);
							if (NULL == pBand)
								return;

							HRESULT hResult;
							if (eBand == tvItem.iImage)
							{
								ChildBandStyles cbsStyle;
								hResult = pBand->get_ChildBandStyle(&cbsStyle);
								if (FAILED(hResult))
									return;

								if (ddCBSNone == cbsStyle)
								{
									ITools* pTools;
									hResult = pBand->get_Tools((Tools**)&pTools);
									if (FAILED(hResult))
										return;

									FillTools(pTools, pTreeView->itemNew.hItem);

									pTools->Release();
								}
								else
								{
									IChildBands* pChildBands;
									hResult = pBand->get_ChildBands((ChildBands**)&pChildBands);
									if (FAILED(hResult))
										return;

									DDString strChildBands;
									strChildBands.LoadString(IDS_CHILDBANDS);
									HTREEITEM hChildBands = InsertNode (pTreeView->itemNew.hItem, 
																		strChildBands, 
																		eChildBands, 
																		I_CHILDRENCALLBACK, 
																		(long)pChildBands);
								}
							}
							else
							{
								ITools* pTools;
								hResult = pBand->get_Tools((Tools**)&pTools);
								if (FAILED(hResult))
									return;

								FillTools(pTools, pTreeView->itemNew.hItem);

								pTools->Release();
							}
						}
						CATCH 
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
					}
					break;

				case eChildBands:
					{
						IChildBands* pChildBands = reinterpret_cast<IChildBands*>(tvItem.lParam);
						if (NULL == pChildBands)
							return;

						FillChildBands(pChildBands, pTreeView->itemNew.hItem);
					}
					break;

				case eChildBand:
					{
						IBand* pChildBand = reinterpret_cast<IBand*>(tvItem.lParam);
						if (NULL == pChildBand)
							return;

						ITools* pTools;
						HRESULT hResult = pChildBand->get_Tools((Tools**)&pTools);
						if (FAILED(hResult))
							return;

						FillTools(pTools, pTreeView->itemNew.hItem);

						pTools->Release();
					}
					break;
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
// OnGetDisplayInfo
//

void CDesignerTree::OnGetDisplayInfo(TV_DISPINFO* pDispInfo)
{
	try
	{
		if (TVIF_CHILDREN & pDispInfo->item.mask)
		{
			TV_ITEM tvItem;
			tvItem.mask = TVIF_IMAGE|TVIF_PARAM; 
			tvItem.hItem = pDispInfo->item.hItem;
			if (TreeView_GetItem(m_hWnd, &tvItem))
			{
				HRESULT hResult;
				short nCount = 0;
				switch (tvItem.iImage)
				{
				case eRoot:
					try
					{
						pDispInfo->item.cChildren = 1;
					}
					CATCH 
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
					break;

				case eBands:
					try
					{
						IBands* pBands;
						hResult = m_pActiveBar->get_Bands((Bands**)&pBands);
						if (FAILED(hResult))
						{
							MessageBox(IDS_ERR_FAILEDTOGETBANDSOBJECT);
							return;
						}

						hResult = pBands->Count(&nCount);
						pBands->Release();
						if (FAILED(hResult))
							return;

						pDispInfo->item.cChildren = 0;
						if (nCount > 0)
							pDispInfo->item.cChildren = 1;
					}
					CATCH 
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
					break;

				case eCategories:
					try
					{
						pDispInfo->item.cChildren = 1;
					}
					CATCH 
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
					break;

				case eCategory:
					try
					{
						pDispInfo->item.cChildren = 0;
						CatEntry* pCatEntry = reinterpret_cast<CatEntry*>(tvItem.lParam);
						if (pCatEntry && pCatEntry->GetTools().GetSize() > 0)
							pDispInfo->item.cChildren = 1;
					}
					CATCH 
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
					break;

				case eBand:
				case eMenu:
				case eChildMenu:
				case ePopup:
				case eStatusBar:
					try
					{
						IBand* pBand = reinterpret_cast<IBand*>(tvItem.lParam);
						if (NULL == pBand)
							return;

						if (eBand == tvItem.iImage)
						{
							ChildBandStyles cbsStyle;
							hResult = pBand->get_ChildBandStyle(&cbsStyle);
							if (FAILED(hResult))
								return;

							if (ddCBSNone == cbsStyle)
							{
								ITools* pTools;
								hResult = pBand->get_Tools((Tools**)&pTools);
								if (FAILED(hResult))
									return;

								hResult = pTools->Count(&nCount);
								pTools->Release();
								if (FAILED(hResult))
									return;

								pDispInfo->item.cChildren = 0;
								if (nCount > 0)
									pDispInfo->item.cChildren = 1;
							}
							else
								pDispInfo->item.cChildren = 1;
						}
						else
						{
							ITools* pTools;
							hResult = pBand->get_Tools((Tools**)&pTools);
							if (FAILED(hResult))
								return;

							hResult = pTools->Count(&nCount);
							pTools->Release();
							if (FAILED(hResult))
								return;

							pDispInfo->item.cChildren = 0;
							if (nCount > 0)
								pDispInfo->item.cChildren = 1;
						}
					}
					CATCH 
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
					break;

				case eChildBands:
					try
					{
						IChildBands* pChildBands = reinterpret_cast<IChildBands*>(tvItem.lParam);
						if (NULL == pChildBands)
							return;

						hResult = pChildBands->Count(&nCount);
						if (FAILED(hResult))
							return;

						pDispInfo->item.cChildren = 0;
						if (nCount > 0)
							pDispInfo->item.cChildren = 1;
					}
					CATCH 
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
					break;

				case eChildBand:
					try
					{
						IBand* pChildBand = reinterpret_cast<IBand*>(tvItem.lParam);
						if (NULL == pChildBand)
							return;

						ITools* pTools;
						hResult = pChildBand->get_Tools((Tools**)&pTools);
						if (FAILED(hResult))
							return;
						
						hResult = pTools->Count(&nCount);
						pTools->Release();
						if (FAILED(hResult))
							return;

						pDispInfo->item.cChildren = 0;
						if (nCount > 0)
							pDispInfo->item.cChildren = 1;
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
		if (TVIF_TEXT & pDispInfo->item.mask)
		{
			TV_ITEM tvItem;
			tvItem.mask = TVIF_IMAGE|TVIF_PARAM; 
			tvItem.hItem = pDispInfo->item.hItem;
			if (!TreeView_GetItem(m_hWnd, &tvItem))
				return; 

			HRESULT hResult;
			BSTR bstrName;
			switch (tvItem.iImage)
			{
			case eBand:
			case eMenu:
			case eChildMenu:
			case ePopup:
			case eStatusBar:
			case eChildBand:
				{
					IBand* pBand = reinterpret_cast<IBand*>(tvItem.lParam);
					hResult = pBand->get_Name(&bstrName);
					if (SUCCEEDED(hResult))
					{
						MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
						lstrcpy(pDispInfo->item.pszText, szName);
						SysFreeString(bstrName);
					}
				}
				break;

			case eCategory:
				try
				{
					CatEntry* pCatEntry = reinterpret_cast<CatEntry*>(tvItem.lParam);
					if (pCatEntry)
						pDispInfo->item.pszText = NULL;
				}
				CATCH 
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case eButton:
			case eDropDown:
			case eTextbox:
			case eCombobox:
			case eLabel:
			case eSeparator:
			case eWindowList:
			case eCustomTool:
			case eForm:
				try
				{
					ITool* pTool = reinterpret_cast<ITool*>(tvItem.lParam);
					if (pTool)
					{
						hResult = pTool->get_Name(&bstrName);
						if (SUCCEEDED(hResult))
						{
							MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
							lstrcpy(pDispInfo->item.pszText, szName);
							SysFreeString(bstrName);
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
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnSelChanged
//

void CDesignerTree::OnSelChanged(NM_TREEVIEW* pNMHdr)
{
	try 
	{
		if (m_pBrowser)
		{
			m_pBrowser->Clear();
		}

		m_tvCurrentItem.hItem = m_hCurrentItem = pNMHdr->itemNew.hItem;
		m_tvCurrentItem.mask = TVIF_IMAGE|TVIF_PARAM;

		if (TreeView_GetItem(m_hWnd, &m_tvCurrentItem))
		{
			switch (m_tvCurrentItem.iImage)
			{
			case eRoot:
				if (m_pBrowser)
					m_pBrowser->SetTypeInfo(m_pActiveBar, this);
				break;

			case eBand:
			case eMenu:
			case eChildMenu:
			case ePopup:
			case eStatusBar:
				{
					IBand* pBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
					if (NULL == pBand)
						return;

					if (m_pBrowser)
						m_pBrowser->SetTypeInfo(pBand, this);
				}
				break;

			case eChildBands:
				{
					IChildBands* pChildBands = reinterpret_cast<IChildBands*>(m_tvCurrentItem.lParam);
					if (NULL == pChildBands)
						return;

					if (m_pBrowser)
						m_pBrowser->SetTypeInfo(pChildBands, this);
				}
				break;

			case eChildBand:
				{
					IBand* pChildBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
					if (NULL == pChildBand)
						return;

					if (m_pBrowser)
						m_pBrowser->SetTypeInfo(pChildBand, this);
				}
				break;

			case eButton:
			case eDropDown:
			case eTextbox:
			case eCombobox:
			case eSeparator:
			case eLabel:
			case eWindowList:
			case eCustomTool:
			case eForm:
				{
					ITool* pTool = reinterpret_cast<ITool*>(m_tvCurrentItem.lParam);
					if (NULL == pTool)
						return;

					if (m_pBrowser)
						m_pBrowser->SetTypeInfo(pTool, this);
				}
				break;
			}
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		CONTINUE
	}
}

//
// OnDoubleClick
//

void CDesignerTree::OnDoubleClick()
{
	try
	{
		POINT pt;
		if (HitTestAtCursor(pt, TVIF_IMAGE|TVIF_PARAM))
		{
			switch (m_tvCurrentItem.iImage)
			{
			case eButton:
			case eDropDown:
			case eTextbox:
			case eCombobox:
			case eSeparator:
			case eLabel:
			case eWindowList:
			case eCustomTool:
			case eForm:
				{
					ITool* pTool = reinterpret_cast<ITool*>(m_tvCurrentItem.lParam);
					if (NULL == pTool)
					{
						MessageBox(IDS_ERR_FAILEDTOGETTOOL);
						return;
					}
					static BOOL bIconUp = FALSE;
					if (bIconUp)
					{
						MessageBeep(MB_ICONEXCLAMATION);
						return;
					}

					EnableWindow(::GetParent(m_hWnd), FALSE);
					bIconUp = TRUE;
					CIconEditor dlgIconEditor(pTool);
					if (IDOK == dlgIconEditor.DoModal(::GetParent(::GetParent(m_hWnd))))
					{
						ObjectChanged(m_tvCurrentItem.iImage);
						if (eBrowser == m_eTreeType)
							Update();
					}
					EnableWindow(::GetParent(m_hWnd), TRUE);
					bIconUp = FALSE;
				}
				break;
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
// OnRClick
//

void CDesignerTree::OnRightClick()
{
	try
	{
		POINT pt;
		if (HitTestAtCursor(pt, TVIF_IMAGE|TVIF_PARAM))
		{
			switch (m_tvCurrentItem.iImage)
			{
			case eRoot:
				PopupMenu (IDR_MNU_BAND, eMenuRoot, pt);
				break;

			case eBands:
				PopupMenu (IDR_MNU_BAND, eMenuBands, pt);
				break;

			case eBand:
			case eMenu:
			case eChildMenu:
			case ePopup:
			case eStatusBar:
				{
					IBand* pBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
					if (NULL == pBand)
						return;

					if (eBand == m_tvCurrentItem.iImage)
					{
						ChildBandStyles cbsStyle;
						HRESULT hResult = pBand->get_ChildBandStyle(&cbsStyle);
						if (FAILED(hResult))
							return;

						if (ddCBSNone == cbsStyle)
							PopupMenu (IDR_MNU_BAND, eMenuBandNone, pt);
						else
							PopupMenu (IDR_MNU_BAND, eMenuBandChildBands, pt);
					}
					else
						PopupMenu (IDR_MNU_BAND, eMenuBandNone, pt);
				}
				break;

			case eChildBands:
				PopupMenu (IDR_MNU_BAND, eMenuChildBands, pt);
				break;

			case eChildBand:
				PopupMenu (IDR_MNU_BAND, eMenuChildBand, pt);
				break;

			case eButton:
			case eDropDown:
			case eTextbox:
			case eCombobox:
			case eSeparator:
			case eLabel:
			case eWindowList:
			case eCustomTool:
			case eForm:
				PopupMenu (IDR_MNU_BAND, eMenuTool, pt);
				break;

			case eCategories:
				PopupMenu (IDR_MNU_BAND, eMenuCategories, pt);
				break;

			case eCategory:
				{
					CatEntry* pEntry = reinterpret_cast<CatEntry*>(m_tvCurrentItem.lParam);
					if (NULL == pEntry)
						return;

					PopupMenu (IDR_MNU_BAND, eMenuCategory, pt);
				}
				break;
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
// OnKeyDown
//

LRESULT CDesignerTree::OnKeyDown(LPNMTVKEYDOWN pKeyDown)
{
	try
	{
		switch (pKeyDown->wVKey)
		{
		case VK_UP:
		case VK_DOWN:
		case VK_LEFT:
		case VK_RIGHT:
			return -1;

		case VK_F2:
			TreeView_EditLabel(m_hWnd, m_hCurrentItem);
			return -1;

		case VK_F10:
			{
				CRect rc;
				if (!TreeView_GetItemRect(m_hWnd, m_hCurrentItem, &rc, TRUE))
					break;

				POINT pt;
				pt.x = rc.left;
				pt.y = rc.bottom;

				switch (m_tvCurrentItem.iImage)
				{
				case eRoot:
					PopupMenu (IDR_MNU_BAND, eMenuRoot, pt);
					return -1;

				case eBands:
					PopupMenu (IDR_MNU_BAND, eMenuBands, pt);
					return -1;

				case eCategories:
					PopupMenu (IDR_MNU_BAND, eMenuCategories, pt);
					return -1;

				case eCategory:
					PopupMenu (IDR_MNU_BAND, eMenuCategory, pt);
					return -1;

				case eBand:
				case eMenu:
				case ePopup:
				case eChildMenu:
				case eStatusBar:
					{
						IBand* pBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
						if (NULL == pBand)
							break;

						if (eBand == m_tvCurrentItem.iImage)
						{
							ChildBandStyles cbsStyle;
							HRESULT hResult = pBand->get_ChildBandStyle(&cbsStyle);
							if (FAILED(hResult))
								break;

							if (ddCBSNone == cbsStyle)
								PopupMenu (IDR_MNU_BAND, eMenuBandNone, pt);
							else
								PopupMenu (IDR_MNU_BAND, eMenuBandChildBands, pt);
						}
						else
							PopupMenu (IDR_MNU_BAND, eMenuBandNone, pt);
						return -1;
					}
					break;

				case eChildBands:
					PopupMenu (IDR_MNU_BAND, eMenuChildBands, pt);
					return -1;

				case eChildBand:
					PopupMenu (IDR_MNU_BAND, eMenuChildBand, pt);
					return -1;

				case eButton:
				case eDropDown:
				case eTextbox:
				case eCombobox:
				case eSeparator:
				case eLabel:
				case eWindowList:
				case eCustomTool:
				case eForm:
					PopupMenu (IDR_MNU_BAND, eMenuTool, pt);
					return -1;
				}
			}
			break;


		case VK_DELETE:
			if (!GetSelected(TVIF_IMAGE|TVIF_PARAM))
				break;

			switch (m_tvCurrentItem.iImage)
			{
			case eCategory:
				OnDeleteCategory();
				return -1;

			case eBand:
			case eMenu:
			case eChildMenu:
			case ePopup:
			case eStatusBar:
				OnDeleteBand();
				return -1;

			case eChildBand:
				OnDeleteChildBand();
				return -1;

			case eButton:
			case eDropDown:
			case eTextbox:
			case eCombobox:
			case eSeparator:
			case eLabel:
			case eWindowList:
			case eCustomTool:
			case eForm:
				OnDeleteTool();
				return -1;
			}
			break;

		case VK_INSERT:
			if (!GetSelected(TVIF_IMAGE|TVIF_PARAM))
				break;

			switch (m_tvCurrentItem.iImage)
			{
			case eBands:
				OnCreateBand(-1);
				return -1;

			case eChildBands:
				OnCreateChildBand();
				return -1;

			case eBand:
			case eMenu:
			case eChildMenu:
			case ePopup:
			case eStatusBar:
				{
					IBand* pBand = reinterpret_cast<IBand*>(m_tvCurrentItem.lParam);
					if (NULL == pBand)
					{
						MessageBeep(MB_ICONEXCLAMATION);
						break;
					}

					if (eBand == m_tvCurrentItem.iImage)
					{
						ChildBandStyles cbsStyle;
						HRESULT hResult = pBand->get_ChildBandStyle(&cbsStyle);
						if (FAILED(hResult))
						{
							MessageBeep(MB_ICONEXCLAMATION);
							break;
						}

						if (ddCBSNone != cbsStyle)
						{
							MessageBeep(MB_ICONEXCLAMATION);
							break;
						}
						OnCreateTool(ddTTButton);
					}
					else
						OnCreateTool(ddTTButton);
					return -1;
				}
				break;

			case eChildBand:
				OnCreateTool(ddTTButton);
				return -1;

			case eCategories:
				OnCreateCategory();
				return -1;
				break;

			case eCategory:
				OnCreateTool(ddTTButton);
				return -1;
				break;
			}
			break;

		default:
			if (isprint(pKeyDown->wVKey))
			{
				BYTE bKeyState[256];
				GetKeyboardState(bKeyState);
				BOOL bCap = 0x1 & bKeyState[VK_CAPITAL];
				BOOL bShift = 0x8 & bKeyState[VK_SHIFT];
				if (!bShift && !bCap || bCap && bShift)
					pKeyDown->wVKey = _totlower(pKeyDown->wVKey);

				m_pBrowser->PostMessage(GetGlobals().WM_SETEDITFOCUS, pKeyDown->wVKey);
				return -1;
			}
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return 0;
}

//
// OnDeleteItem
//

void CDesignerTree::OnDeleteItem(NM_TREEVIEW* pTreeView)
{
	try
	{
		TVITEM tvDeleteItem;
		tvDeleteItem.mask = TVIF_PARAM|TVIF_IMAGE;
		tvDeleteItem.hItem = pTreeView->itemOld.hItem;

		if (TreeView_GetItem(m_hWnd, &tvDeleteItem))
		{
			switch (tvDeleteItem.iImage)
			{
			case eCategory:
				break;

			default:
				if (pTreeView->itemOld.lParam)
					reinterpret_cast<LPUNKNOWN>(pTreeView->itemOld.lParam)->Release();
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
// Getting a list of the tools
//

BOOL CDesignerTree::GetListOfSelectedTools(TypedArray<ITool*>& aTools)
{
	TVITEM tvItem;
	tvItem.mask = TVIF_IMAGE|TVIF_PARAM; 
	tvItem.hItem = GetFirstSelectedItem();
	while (tvItem.hItem)
	{
		if (TreeView_GetItem(m_hWnd, &tvItem))
		{
			switch (tvItem.iImage)
			{
			case eButton:
			case eDropDown:
			case eCombobox:
			case eTextbox:
			case eLabel:
			case eWindowList:
			case eSeparator:
			case eCustomTool:
			case eForm:
				aTools.Add((ITool*)tvItem.lParam);
				break;
			}
			tvItem.hItem = GetNextSelectedItem(tvItem.hItem);
		}
	}
	return TRUE;
}

//
// OnBeginDrag
//

void CDesignerTree::OnBeginDrag(LPNMTREEVIEW pTreeView)
{
	CTreeCtrlSource* pSource = NULL;
	HRESULT hResult;
	
	TV_ITEM tvDragItem;

	ITools* pTools = NULL;
	ITool*  pTool = NULL; 

	IBarPrivate* pPrivate = NULL;

	BSTR bstrChildBand = NULL;
	BSTR bstrBand = NULL;

	DWORD dwEffect;
	
	tvDragItem.mask = TVIF_PARAM|TVIF_IMAGE;
	tvDragItem.hItem = pTreeView->itemNew.hItem;

	SetItemState(tvDragItem.hItem, TVIS_DROPHILITED, TVIS_DROPHILITED);

	hResult = m_pActiveBar->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
	if (FAILED(hResult))
		goto Cleanup;

	if (!TreeView_GetItem(m_hWnd, &tvDragItem))
	{
		assert(FALSE);
		return;
	}

	try
	{
		TypedArray<ITool*> aTools;
		switch (tvDragItem.iImage)
		{
		case eButton:
		case eDropDown:
		case eCombobox:
		case eTextbox:
		case eLabel:
		case eWindowList:
		case eSeparator:
		case eCustomTool:
		case eForm:
			{
				//
				// Finding the parent of a Tool Object
				//

				TV_ITEM tvParent;
				if (!GetParent(tvDragItem.hItem, tvParent, TVIF_HANDLE|TVIF_PARAM|TVIF_IMAGE))
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOL);
					goto Cleanup;
				}

				switch (tvParent.iImage)
				{
				case eBand:
				case eMenu:
				case ePopup:
				case eChildMenu:
				case eStatusBar:
					{
						//
						// Parent is a Band
						//

						IBand* pBand = reinterpret_cast<IBand*>(tvParent.lParam);
						if (NULL == pBand)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}
						
						//
						// Get the Band Name
						//

						hResult = pBand->get_Name(&bstrBand);
						if (FAILED(hResult))
							goto Cleanup;

						//
						// Get a list of tools
						//

						if (!GetListOfSelectedTools(aTools))
							goto Cleanup;

						m_pDataObject = new CToolDataObject(eBrowser == m_eTreeType ? CToolDataObject::eDesignerDragDropId : CToolDataObject::eLibraryDragDropId, 
															m_pActiveBar, 
															aTools, 
															bstrBand, 
															bstrChildBand);
						if (NULL == m_pDataObject)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						DWORD dwEffect;
						pSource = new CTreeCtrlSource;
						if (NULL == pSource)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						if (eBrowser == m_eTreeType)
							pPrivate->put_CustomizeDragLock(aTools.GetAt(0));

						hResult = GetGlobals().GetDragDropMgr()->DoDragDrop(m_pDataObject, 
																			pSource, 
																			DROPEFFECT_COPY | DROPEFFECT_MOVE,
																			&dwEffect);
						if (eBrowser == m_eTreeType)
							pPrivate->put_CustomizeDragLock(NULL);

						if (DRAGDROP_S_DROP != hResult)
							goto Cleanup;

						if (DROPEFFECT_MOVE & dwEffect)
						{
							hResult = pBand->get_Tools((Tools**)&pTools);
							if (SUCCEEDED(hResult))
							{
								int nCount = aTools.GetSize();
								for (int nTool = 0; nTool < nCount; nTool++)
									pTools->DeleteTool(aTools.GetAt(nTool));
							}
						}
						Refresh(tvParent.hItem);
					}
					break;

				case eChildBand:
					{
						IBand* pChildBand = reinterpret_cast<IBand*>(tvParent.lParam);
						if (NULL == pChildBand)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						hResult = pChildBand->get_Name(&bstrChildBand);
						if (FAILED(hResult))
							goto Cleanup;

						//
						// Get the Child Band Node
						//
						
						TV_ITEM tvChildBands;
						if (!GetParent(tvParent.hItem, tvChildBands, TVIF_HANDLE|TVIF_PARAM))
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						//
						// Get the Parent Band Node
						//

						TV_ITEM tvBand;
						if (!GetParent(tvChildBands.hItem, tvBand, TVIF_HANDLE|TVIF_PARAM))
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						//
						// Get the Parent Band Object
						//

						IBand* pBand = reinterpret_cast<IBand*>(tvBand.lParam);
						if (NULL == pBand)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						//
						// Get the Band's Name
						//

						hResult = pBand->get_Name(&bstrBand);
						if (FAILED(hResult))
							goto Cleanup;

						if (!GetListOfSelectedTools(aTools))
							goto Cleanup;

						m_pDataObject = new CToolDataObject(eBrowser == m_eTreeType ? CToolDataObject::eDesignerDragDropId : CToolDataObject::eLibraryDragDropId, 
															m_pActiveBar, 
															aTools, 
															bstrBand, 
															bstrChildBand);
						if (eBrowser == m_eTreeType)
							pPrivate->put_CustomizeDragLock(NULL);

						if (NULL == m_pDataObject)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						pSource = new CTreeCtrlSource;
						if (NULL == pSource)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}


						if (eBrowser == m_eTreeType)
							pPrivate->put_CustomizeDragLock(aTools.GetAt(0));

						hResult = GetGlobals().GetDragDropMgr()->DoDragDrop(m_pDataObject, 
																			pSource, 
																			DROPEFFECT_COPY | DROPEFFECT_MOVE,
																			&dwEffect);
						if (eBrowser == m_eTreeType)
							pPrivate->put_CustomizeDragLock(NULL);

						if (DRAGDROP_S_DROP != hResult)
							goto Cleanup;

						if (DROPEFFECT_MOVE & dwEffect)
						{
							hResult = pChildBand->get_Tools((Tools**)&pTools);
							if (SUCCEEDED(hResult))
							{
								int nCount = aTools.GetSize();
								for (int nTool = 0; nTool < nCount; nTool++)
									pTools->DeleteTool(aTools.GetAt(nTool));
							}
						}
						Refresh(tvParent.hItem);
					}
					break;

				case eCategory:
					{
						CatEntry* pCateEntry = reinterpret_cast<CatEntry*>(tvParent.lParam);
						assert(pCateEntry);
						if (pCateEntry)
						{
							if (!GetListOfSelectedTools(aTools))
								goto Cleanup;
						}

						m_pDataObject = new CToolDataObject(eBrowser == m_eTreeType ? CToolDataObject::eDesignerMainToolDragDropId : CToolDataObject::eLibraryDragDropId, 
															m_pActiveBar, 
															aTools, 
															bstrBand, 
															bstrChildBand);
						if (NULL == m_pDataObject)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						pSource = new CTreeCtrlSource;
						if (NULL == pSource)
						{
							MessageBox(IDS_ERR_FAILEDTOGETTOOL);
							goto Cleanup;
						}

						if (eBrowser == m_eTreeType)
							pPrivate->put_CustomizeDragLock(aTools.GetAt(0));

						hResult = GetGlobals().GetDragDropMgr()->DoDragDrop(m_pDataObject, 
																			pSource, 
																			DROPEFFECT_COPY | DROPEFFECT_MOVE,
																			&dwEffect);
						if (eBrowser == m_eTreeType)
							pPrivate->put_CustomizeDragLock(NULL);

						if (DRAGDROP_S_DROP != hResult)
							goto Cleanup;

						if (DROPEFFECT_MOVE & dwEffect)
						{
							int nCount = aTools.GetSize();
							for (int nTool = 0; nTool < nCount; nTool++)
							{
								pTool = aTools.GetAt(nTool);
								if (pTool)
									pCateEntry->DeleteTool(pTool);
							}
						}
						Refresh(tvParent.hItem);
					}
					break;

				default:
					assert(FALSE);
					break;
				}
			}
			break;


		case eBand:
		case eMenu:
		case ePopup:
		case eChildMenu:
		case eChildBand:
		case eStatusBar:
			{
				IBand* pBand = reinterpret_cast<IBand*>(tvDragItem.lParam);
				if (NULL == pBand)
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOL);
					goto Cleanup;
				}

				m_pDataObject = new CToolDataObject(eBrowser == m_eTreeType ? CToolDataObject::eDesignerDragDropId : CToolDataObject::eLibraryDragDropId, 
													m_pActiveBar, 
													pBand, 
													aTools);
				if (NULL == m_pDataObject)
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOL);
					goto Cleanup;
				}

				DWORD dwEffect;
				try
				{
					pSource = new CTreeCtrlSource;
					if (NULL == pSource)
					{
						MessageBox(IDS_ERR_FAILEDTOGETTOOL);
						goto Cleanup;
					}

					hResult = GetGlobals().GetDragDropMgr()->DoDragDrop(m_pDataObject, 
																		pSource, 
																		DROPEFFECT_COPY | DROPEFFECT_MOVE,
																		&dwEffect);
				}
				CATCH 
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
					goto Cleanup;
				}
			}
			break;

		case eCategory:
			{
				CatEntry* pCatEntry = reinterpret_cast<CatEntry*>(tvDragItem.lParam);
				if (NULL == pCatEntry)
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOL);
					goto Cleanup;
				}

				m_pDataObject = new CToolDataObject(eBrowser == m_eTreeType ? CToolDataObject::eDesignerDragDropId : CToolDataObject::eLibraryDragDropId, 
													m_pActiveBar, 
													pCatEntry->Category(), 
													pCatEntry->GetTools());
				if (NULL == m_pDataObject)
				{
					MessageBox(IDS_ERR_FAILEDTOGETTOOL);
					goto Cleanup;
				}

				DWORD dwEffect;
				try
				{
					pSource = new CTreeCtrlSource;
					if (NULL == pSource)
					{
						MessageBox(IDS_ERR_FAILEDTOGETTOOL);
						goto Cleanup;
					}

					hResult = GetGlobals().GetDragDropMgr()->DoDragDrop(m_pDataObject, 
																		pSource, 
																		DROPEFFECT_COPY | DROPEFFECT_MOVE,
																		&dwEffect);
				}
				CATCH 
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
					goto Cleanup;
				}
			}
			break;
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
Cleanup:
	try
	{
		if (bstrChildBand)
			SysFreeString(bstrChildBand);

		if (bstrBand)
			SysFreeString(bstrBand);

		if (pSource)
			pSource->Release();

		if (pTools)
			pTools->Release();

		if (m_pDataObject)
		{
			m_pDataObject->Release();
			m_pDataObject = NULL;
		}
		if (pPrivate)
			pPrivate->Release();
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnBeginEditLabel
//

LRESULT CDesignerTree::OnBeginEditLabel(TV_DISPINFO* pDispInfo)
{
	BOOL bResult = FALSE;
	pDispInfo->item.mask = TVIF_IMAGE; 
	TreeView_GetItem(m_hWnd, &pDispInfo->item);
	switch (pDispInfo->item.iImage)
	{
	case eRoot:
	case eBands:
	case eChildBands:
	case eCategories:
		bResult = TRUE;
		break;

	case eCategory: 
		if (0 == _tcscmp(m_strNone, pDispInfo->item.pszText))
			bResult = TRUE;
		break;
	}
	return bResult;
}

//
// OnEndEditLabel
//

LRESULT CDesignerTree::OnEndEditLabel(TV_DISPINFO* pDispInfo)
{
	BOOL bResult = TRUE;
	if (pDispInfo->item.pszText)
	{
		bResult = TRUE;
		switch (m_tvCurrentItem.iImage)
		{
		case eBand:
		case eMenu:
		case ePopup:
		case eStatusBar:
		case eChildBand:
			{
				DDString strTemp(pDispInfo->item.pszText);
				IBand* pBand = reinterpret_cast<IBand*>(pDispInfo->item.lParam);
				if (pBand)
				{
					HRESULT hResult = pBand->put_Name(strTemp.AllocSysString());
					if (SUCCEEDED(hResult) && m_pBrowser)
					{
						m_pBrowser->Clear();
						m_pBrowser->SetTypeInfo(pBand, this);
					}
				}
			}
			break;

		case eButton:
		case eDropDown:
		case eTextbox:
		case eCombobox:
		case eSeparator:
		case eLabel:
		case eWindowList:
		case eCustomTool:
		case eForm:
			{
				DDString strTemp(pDispInfo->item.pszText);
				ITool* pTool = reinterpret_cast<ITool*>(pDispInfo->item.lParam);
				if (pTool)
				{
					HRESULT hResult = pTool->put_Name(strTemp.AllocSysString());
					if (SUCCEEDED(hResult) && m_pBrowser)
					{
						m_pBrowser->Clear();
						m_pBrowser->SetTypeInfo(pTool, this);
					}
				}
			}
			break;

		case eCategory: 
			{
				// Category
				if (m_pCategoryMgr && pDispInfo)
				{
					m_pCategoryMgr->ChangeName(reinterpret_cast<CatEntry*>(pDispInfo->item.lParam), pDispInfo->item.pszText);
				}
			}
			break;
		}
	}
	return bResult;
}

//
// CategoryChanged
//

void CDesignerTree::CategoryChanged(CatEntry*						  pCatEntry, 
									CategoryMgrNotify::CategoryChange ccChange)
{
	switch (ccChange)
	{
	case CategoryMgrNotify::eCategoryDeleted:
		{
			HTREEITEM hTreeItem = TreeView_GetRoot(m_hWnd);
			if (NULL == hTreeItem)
				return;

			hTreeItem = TreeView_GetChild(m_hWnd, hTreeItem); 
			if (NULL == hTreeItem)
				return;

			hTreeItem = TreeView_GetChild(m_hWnd, hTreeItem);
			if (NULL == hTreeItem)
				return;

			TV_ITEM tvItem;
			CatEntry* pCatEntry2;
			while (hTreeItem)
			{
				tvItem.mask = TVIF_IMAGE|TVIF_PARAM; 
				tvItem.hItem = hTreeItem;
				TreeView_GetItem(m_hWnd, &tvItem); 
				pCatEntry2 = reinterpret_cast<CatEntry*>(tvItem.lParam);
				if (pCatEntry2 == pCatEntry)
				{
					TreeView_DeleteItem(m_hWnd, hTreeItem);
					break;
				}
				hTreeItem = TreeView_GetNextSibling(m_hWnd, hTreeItem);
			}
		}
		break;

	case CategoryMgrNotify::eNewCategory:
		{
			MAKE_TCHARPTR_FROMWIDE(szCategory, pCatEntry->Category());
			HTREEITEM hTreeItem = TreeView_GetRoot(m_hWnd);
			if (NULL == hTreeItem)
				return;

			HTREEITEM hChild = TreeView_GetChild(m_hWnd, hTreeItem); 
			if (hChild)
			{
				TV_ITEM tvItem1;
				tvItem1.mask = TVIF_STATE; 
				tvItem1.hItem = hChild;
				if (TreeView_GetItem(m_hWnd, &tvItem1))
				{
					if (TVIS_EXPANDED & tvItem1.state)
					{
						InsertNode (hChild, 
								    szCategory, 
								    eCategory, 
								    I_CHILDRENCALLBACK, 
								    (long)pCatEntry);
					}
					CRect rcItem;
					if (TreeView_GetItemRect(m_hWnd, hChild, &rcItem, FALSE))
						InvalidateRect(&rcItem, FALSE);
				}
			}
		}
		break;
	}
}

//
// ToolChanged
//

void CDesignerTree::ToolChanged(BSTR						  bstrCategory, 
								ITool*						  pTool, 
								CategoryMgrNotify::ToolChange tcChange)
{
	switch (tcChange)
	{
	case CategoryMgrNotify::eNewTool:
		{
			HTREEITEM hCatRoot = FindCategoryNode(bstrCategory);
			if (hCatRoot)
			{
				TV_ITEM tvItem1;
				tvItem1.mask = TVIF_STATE; 
				tvItem1.hItem = hCatRoot;
				if (TreeView_GetItem(m_hWnd, &tvItem1))
				{
					if (!(TVIS_EXPANDED & tvItem1.state))
						return;

					BSTR bstrName;
					HRESULT hResult = pTool->get_Name(&bstrName);
					if (SUCCEEDED(hResult))
					{
						MAKE_TCHARPTR_FROMWIDE(szTool, bstrName);
						SysFreeString(bstrName);
						bstrName = NULL;
						
						ToolTypes ttType;
						hResult = pTool->get_ControlType(&ttType);
						
						pTool->AddRef();
						InsertNode (hCatRoot, 
									LPSTR_TEXTCALLBACK, 
									GetToolImageType(ttType), 
									0, 
									(long)pTool);
					}
				}
			}
		}
		break;

	case CategoryMgrNotify::eToolDeleted:
		{
			HTREEITEM hCatRoot = FindCategoryNode(bstrCategory);
			if (hCatRoot)
			{
				ITool*  pTool2;
				TV_ITEM tvItem;
				tvItem.mask = TVIF_PARAM; 
				HTREEITEM hChild = TreeView_GetChild(m_hWnd, hCatRoot); 
				while (hChild)
				{
					tvItem.hItem = hChild;
					if (TreeView_GetItem(m_hWnd, &tvItem))
					{
						pTool2 = (ITool*)tvItem.lParam;
						if (pTool2 == pTool)
						{
							TreeView_DeleteItem(m_hWnd, hChild);
							break;
						}
					}
					hChild = TreeView_GetNextSibling(m_hWnd, hChild);
				}
			}
		}
		break;
	}
}

//
// ManagerChanged
//

void CDesignerTree::ManagerChanged(CategoryMgrNotify::ManagerChange mcChange)
{
}

//
// FindCategoryNode
//

HTREEITEM CDesignerTree::FindCategoryNode(BSTR bstrCategory)
{
	HTREEITEM hTreeItem = TreeView_GetRoot(m_hWnd);
	if (NULL == hTreeItem)
		return NULL;

	hTreeItem = TreeView_GetChild(m_hWnd, hTreeItem);
	if (NULL == hTreeItem)
		return NULL;

	hTreeItem = TreeView_GetChild(m_hWnd, hTreeItem);
	if (NULL == hTreeItem)
		return NULL;

	TV_ITEM tvItem;
	CatEntry* pCatEntry2;
	while (hTreeItem)
	{
		tvItem.mask = TVIF_IMAGE|TVIF_PARAM; 
		tvItem.hItem = hTreeItem;
		TreeView_GetItem(m_hWnd, &tvItem); 

		pCatEntry2 = reinterpret_cast<CatEntry*>(tvItem.lParam);

		if ((NULL == pCatEntry2->Category() && NULL == bstrCategory) ||
			(NULL == pCatEntry2->Category() && bstrCategory && NULL == *bstrCategory) ||
			(pCatEntry2->Category() && NULL == *pCatEntry2->Category() && NULL == bstrCategory) ||
			(pCatEntry2->Category() && bstrCategory && 0 == wcscmp(pCatEntry2->Category(), bstrCategory)))
		{
			return hTreeItem;
		}
		
		hTreeItem = TreeView_GetNextSibling(m_hWnd, hTreeItem);
	}
	return NULL;
}

//
// CDesignerTarget
//

CDesignerTarget::CDesignerTarget(long nDestId, CDesignerTree& theTreeCtrl)
	: m_theTreeCtrl(theTreeCtrl),
	  m_pDataObject(NULL),
	  m_pSrcBar(NULL)
{
	m_cRef = 1;
	m_nToolFormat = -1;
	m_nIndex = -1;
	m_nDestId = nDestId;
    SetDefFormatEtc(m_fe[0], GetGlobals().m_nIDClipToolIdFormat, TYMED_ISTREAM);
    SetDefFormatEtc(m_fe[1], GetGlobals().m_nIDClipBandToolIdFormat, TYMED_ISTREAM);
    SetDefFormatEtc(m_fe[2], GetGlobals().m_nIDClipBandChildBandToolIdFormat, TYMED_ISTREAM);
    SetDefFormatEtc(m_fe[3], GetGlobals().m_nIDClipBandFormat, TYMED_ISTREAM);
    SetDefFormatEtc(m_fe[4], GetGlobals().m_nIDClipCategoryFormat, TYMED_ISTREAM);
    SetDefFormatEtc(m_fe[5], GetGlobals().m_nIDClipToolFormat, TYMED_ISTREAM);
}

HRESULT __stdcall CDesignerTarget::QueryInterface(REFIID riid, void** ppvObject)
{
	if (riid == IID_IDropTarget || riid == IID_IUnknown)
	{
		*ppvObject = static_cast<IDropTarget*>(this);
		AddRef();
		return S_OK;
	}
	*ppvObject = 0;
	return E_NOINTERFACE;
}
        
ULONG __stdcall CDesignerTarget::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

ULONG __stdcall CDesignerTarget::Release()
{
	if (0 == InterlockedDecrement(&m_cRef))
	{
		delete this;
		return 0;
	}
	return m_cRef;
}

void CDesignerTarget::FeedBack(int nImage, DWORD* pdwEffect, DWORD grfKeyState, WORD nToolFormat)
{
	switch (m_theTreeCtrl.GetCurrentTreeViewItem().iImage)
	{
	case CDesignerTree::eBand:
	case CDesignerTree::eMenu:
	case CDesignerTree::eChildMenu:
	case CDesignerTree::ePopup:
	case CDesignerTree::eStatusBar:
	case CDesignerTree::eChildBand:
		{
			//
			// Band
			//

			if (m_pSrcBar == m_theTreeCtrl.ActiveBar())
			{
				// Same ActiveBar
				if (GetGlobals().m_nIDClipToolIdFormat == nToolFormat ||
					GetGlobals().m_nIDClipBandToolIdFormat == nToolFormat ||
					GetGlobals().m_nIDClipBandChildBandToolIdFormat == nToolFormat)
				{
					if (m_nSrcId != m_nDestId)
						*pdwEffect = DROPEFFECT_COPY;
					else if (MK_CONTROL & grfKeyState)
						*pdwEffect = DROPEFFECT_COPY;
					else
						*pdwEffect = DROPEFFECT_MOVE;
				}
				else if (GetGlobals().m_nIDClipToolFormat == nToolFormat)
				{
					if (m_nSrcId == m_nDestId)
						*pdwEffect = DROPEFFECT_MOVE;
					else if (MK_CONTROL & grfKeyState)
						*pdwEffect = DROPEFFECT_COPY;
					else
						*pdwEffect = DROPEFFECT_MOVE;
				}
			}
			else
			{
				// Different ActiveBars
				if (GetGlobals().m_nIDClipToolFormat == nToolFormat)
					*pdwEffect = DROPEFFECT_COPY;
			}
		}
		break;

	case CDesignerTree::eChildBands:
		{
			//
			// Bands
			//

			if (GetGlobals().m_nIDClipBandFormat == nToolFormat)
			{
				if ((m_nSrcId == CToolDataObject::eLibraryDragDropId && CToolDataObject::eLibraryDragDropId != m_nDestId) || 
					(CToolDataObject::eDesignerMainToolDragDropId == m_nSrcId && CToolDataObject::eDesignerMainToolDragDropId != m_nDestId))
					*pdwEffect = DROPEFFECT_COPY;
				else if (MK_CONTROL & grfKeyState)
					*pdwEffect = DROPEFFECT_COPY;
				else
					*pdwEffect = DROPEFFECT_MOVE;
			}
		}
		break;

	case CDesignerTree::eButton:
	case CDesignerTree::eDropDown:
	case CDesignerTree::eTextbox:
	case CDesignerTree::eCombobox:
	case CDesignerTree::eSeparator:
	case CDesignerTree::eLabel:
	case CDesignerTree::eWindowList:
	case CDesignerTree::eCustomTool:
	case CDesignerTree::eForm:
		{
			//
			// Tools
			//

			TV_ITEM tvCategory;
			tvCategory.iImage = -1;
			tvCategory.hItem = TreeView_GetParent(m_theTreeCtrl.hWnd(), m_theTreeCtrl.GetCurrentItem());
			if (NULL != tvCategory.hItem)
			{
				tvCategory.mask = TVIF_IMAGE;
				if (!TreeView_GetItem(m_theTreeCtrl.hWnd(), &tvCategory))
					tvCategory.iImage = -1;
			}
			if (CDesignerTree::eCategory != tvCategory.iImage)
			{
				if (m_pSrcBar == m_theTreeCtrl.ActiveBar())
				{
					// Same ActiveBar
					if (GetGlobals().m_nIDClipToolIdFormat == nToolFormat ||
						GetGlobals().m_nIDClipBandToolIdFormat == nToolFormat ||
						GetGlobals().m_nIDClipBandChildBandToolIdFormat == nToolFormat)
					{
						if (m_nSrcId != m_nDestId)
							*pdwEffect = DROPEFFECT_COPY;
						else if (MK_CONTROL & grfKeyState)
							*pdwEffect = DROPEFFECT_COPY;
						else
							*pdwEffect = DROPEFFECT_MOVE;
					}
					else if (GetGlobals().m_nIDClipToolFormat == nToolFormat)
					{
						if (m_nSrcId == m_nDestId)
							*pdwEffect = DROPEFFECT_MOVE;
						else if (MK_CONTROL & grfKeyState)
							*pdwEffect = DROPEFFECT_COPY;
						else
							*pdwEffect = DROPEFFECT_MOVE;
					}
				}
				else
				{
					// Different ActiveBars
					if (GetGlobals().m_nIDClipToolFormat == nToolFormat)
						*pdwEffect = DROPEFFECT_COPY;
				}
			}
		}
		break;

	case CDesignerTree::eCategory:
		{
			//
			// Tools
			//

			if (m_pSrcBar == m_theTreeCtrl.ActiveBar())
			{
				// Same ActiveBar
				if (GetGlobals().m_nIDClipBandToolIdFormat == nToolFormat ||
					GetGlobals().m_nIDClipBandChildBandToolIdFormat == nToolFormat)
				{
					if (m_nSrcId != m_nDestId)
						*pdwEffect = DROPEFFECT_COPY;
					else if (MK_CONTROL & grfKeyState)
						*pdwEffect = DROPEFFECT_COPY;
					else
						*pdwEffect = DROPEFFECT_MOVE;
				}
				else if (GetGlobals().m_nIDClipToolFormat == nToolFormat ||
					     GetGlobals().m_nIDClipToolIdFormat == nToolFormat)
				{
					if (m_nSrcId == m_nDestId)
						*pdwEffect = DROPEFFECT_MOVE;
					else if (MK_CONTROL & grfKeyState)
						*pdwEffect = DROPEFFECT_COPY;
					else
						*pdwEffect = DROPEFFECT_MOVE;
				}
			}
			else
			{
				// Different ActiveBars
				if (GetGlobals().m_nIDClipToolFormat == nToolFormat)
					*pdwEffect = DROPEFFECT_COPY;
			}
		}
		break;

	case CDesignerTree::eBands:
		if (GetGlobals().m_nIDClipBandFormat == nToolFormat)
			*pdwEffect = DROPEFFECT_COPY;
		break;

	case CDesignerTree::eCategories:
		if (GetGlobals().m_nIDClipCategoryFormat == nToolFormat)
			*pdwEffect = DROPEFFECT_COPY;
		break;
	}
}

//
// DragEnter
//

HRESULT __stdcall CDesignerTarget::DragEnter(IDataObject* pDataObject, 
											 DWORD		  grfKeyState,
											 POINTL		  pt,
											 DWORD*		  pdwEffect)
{
	*pdwEffect = DROPEFFECT_NONE;
	HRESULT hResult = NOERROR;
	m_nIndex = -1;
	m_nToolFormat = -1;
	m_dwStartDelay = 0;
	m_dwCurrentDelay = 0;
	m_theTreeCtrl.GetClientRect(m_rcClient);
	m_rcScroll = m_rcClient;
	m_rcScroll.Inflate(-GetGlobals().m_nScrollInset, -GetGlobals().m_nScrollInset);

	
	m_pDataObject = pDataObject;
	m_pDataObject->AddRef();

	//
	// Is it a valid data object
	//

	POINT pt2 = {pt.x, pt.y};
	if (m_theTreeCtrl.HitTest(pt2, TVIF_IMAGE|TVIF_PARAM))
	{
		m_theTreeCtrl.SelectDropTarget(m_theTreeCtrl.GetCurrentItem());
		int nImage = m_theTreeCtrl.GetCurrentTreeViewItem().iImage;

		ULARGE_INTEGER lnCurPos;
		LARGE_INTEGER lnMoveAmount; 
		STGMEDIUM stm;
		for (int nIndex = 0; nIndex < eNumOfFormats; nIndex++)
		{
			hResult = pDataObject->GetData(&m_fe[nIndex], &stm);
			if (SUCCEEDED(hResult))
			{
				lnMoveAmount.LowPart = lnMoveAmount.HighPart = 0;

				hResult = stm.pstm->Seek(lnMoveAmount, STREAM_SEEK_SET, &lnCurPos);
				if (SUCCEEDED(hResult))
				{
					hResult = stm.pstm->Read(&m_nSrcId, sizeof(m_nSrcId), NULL);
					if (SUCCEEDED(hResult))
					{
						hResult = stm.pstm->Read(&m_pSrcBar, sizeof(m_pSrcBar), NULL);
						if (SUCCEEDED(hResult))
						{
							FeedBack(nImage, pdwEffect, grfKeyState, m_fe[nIndex].cfFormat);
							if (DROPEFFECT_NONE != *pdwEffect)
							{
								m_nToolFormat = m_fe[nIndex].cfFormat;
								m_nIndex = nIndex;
								stm.pstm->Release();
								break;
							}
						}
					}
				}
				stm.pstm->Release();
			}
		}
	}
	return hResult;
}

//
// DragOver
//

HRESULT __stdcall CDesignerTarget::DragOver(DWORD  grfKeyState, POINTL pt, DWORD* pdwEffect)
{
	*pdwEffect = DROPEFFECT_NONE;
	POINT pt2 = {pt.x, pt.y};
	UINT nFlags = 0; 

	ULARGE_INTEGER lnCurPos;
	LARGE_INTEGER lnMoveAmount; 
	STGMEDIUM stm;
	HRESULT hResult = NOERROR;
	m_nIndex = -1;

	if (m_theTreeCtrl.HitTest(pt2, TVIF_IMAGE|TVIF_PARAM))
	{
		m_theTreeCtrl.SelectDropTarget(m_theTreeCtrl.GetCurrentItem());
		int nImage = m_theTreeCtrl.GetCurrentTreeViewItem().iImage;
		for (int nIndex = 0; nIndex < eNumOfFormats; nIndex++)
		{
			hResult = m_pDataObject->GetData(&m_fe[nIndex], &stm);
			if (SUCCEEDED(hResult))
			{
				lnMoveAmount.LowPart = lnMoveAmount.HighPart = 0;

				hResult = stm.pstm->Seek(lnMoveAmount, STREAM_SEEK_SET, &lnCurPos);
				if (SUCCEEDED(hResult))
				{
					hResult = stm.pstm->Read(&m_nSrcId, sizeof(m_nSrcId), NULL);
					if (SUCCEEDED(hResult))
					{
						hResult = stm.pstm->Read(&m_pSrcBar, sizeof(m_pSrcBar), NULL);
						if (SUCCEEDED(hResult))
						{
							FeedBack(nImage, pdwEffect, grfKeyState, m_fe[nIndex].cfFormat);
							if (DROPEFFECT_NONE != *pdwEffect)
							{
								m_nToolFormat = m_fe[nIndex].cfFormat;
								m_nIndex = nIndex;
								stm.pstm->Release();
								break;
							}
						}
					}
				}
				stm.pstm->Release();
			}
		}
	}

	//
	// Scrolling stuff
	//

	m_theTreeCtrl.ScreenToClient(pt2);
	if (PtInRect(&m_rcClient, pt2) && !PtInRect(&m_rcScroll, pt2))
	{
		if ((m_bAutoScrolling  &&
			(int)(m_dwCurrentDelay - m_dwStartDelay) < GetGlobals().m_nScrollInterval) ||
			(int)(m_dwCurrentDelay - m_dwStartDelay) >= GetGlobals().m_nScrollDelay)
		{
			CRect rc, rcClient;
			m_theTreeCtrl.GetClientRect(rcClient);
	
			// Handle horizontal scrolling
			rc = rcClient;
			rc.left = m_rcScroll.right;
			if (PtInRect (&rc, pt2))
				m_theTreeCtrl.SendMessage(WM_HSCROLL, MAKELONG(SB_LINERIGHT, 0));

			rc.left = 0;
			rc.right = m_rcScroll.left;
			if (PtInRect(&rc, pt2))
				m_theTreeCtrl.SendMessage(WM_HSCROLL, MAKELONG(SB_LINELEFT, 0));

			// Handle vertical scrolling
			rc = m_rcClient;
			rc.top = m_rcScroll.bottom;
			if (PtInRect(&rc, pt2))
				m_theTreeCtrl.SendMessage(WM_VSCROLL, MAKELONG(SB_LINEDOWN, 0));

			rc.top = 0;
			rc.bottom = m_rcScroll.top;
			if (PtInRect(&rc, pt2))
				m_theTreeCtrl.SendMessage(WM_VSCROLL, MAKELONG(SB_LINEUP, 0));

			*pdwEffect = DROPEFFECT_SCROLL;
			m_bAutoScrolling = TRUE;
			m_dwStartDelay = (int)GetTickCount();
		}
		m_dwCurrentDelay = (int)GetTickCount();
	}
	else
	{
		m_bAutoScrolling = FALSE;
		m_dwStartDelay = 0;
		m_dwCurrentDelay = 0;
	}
	return NOERROR;
}
        
//
// DragLeave
//

HRESULT __stdcall CDesignerTarget::DragLeave()
{
	m_theTreeCtrl.SelectDropTarget(NULL);
	if (m_pDataObject)
	{
		m_pDataObject->Release();
		m_pDataObject = NULL;
	}
	return NOERROR;
}
        
//
// Drop
//

HRESULT __stdcall CDesignerTarget::Drop(IDataObject* pDataObject,
									    DWORD		 grfKeyState,
										POINTL		 pt,
										DWORD*		 pdwEffect)
{
	HRESULT hResult = S_OK;
	HTREEITEM hRefreshItem = NULL;
	IBarPrivate* pPrivate = NULL;
	STGMEDIUM stm;
	memset(&stm, 0, sizeof(stm));
	ITools* pTools = NULL;
	IBand* pBand = NULL;
	ITool* pTool = NULL;
	BOOL bUpdate = FALSE;
	try
	{
		*pdwEffect = DROPEFFECT_NONE;
		if (NULL == m_pDataObject)
			return NOERROR;
		POINT pt2 = {pt.x, pt.y};

		m_theTreeCtrl.SelectDropTarget(NULL);
		if (!m_theTreeCtrl.HitTestAtCursor(pt2, TVIF_PARAM|TVIF_IMAGE))
			return NOERROR;
		
		int nNodeType = m_theTreeCtrl.GetCurrentTreeViewItem().iImage;
		int nCount;
		
		hResult = m_pDataObject->GetData(&m_fe[m_nIndex], &stm);
		if (FAILED(hResult))
			goto Cleanup;

		LARGE_INTEGER lnMoveAmount; 
		lnMoveAmount.LowPart = lnMoveAmount.HighPart = 0;
		ULARGE_INTEGER lnCurPos;

		hResult = stm.pstm->Seek(lnMoveAmount, STREAM_SEEK_SET, &lnCurPos);
		if (FAILED(hResult))
			goto Cleanup;

		hResult = stm.pstm->Read(&m_nSrcId, sizeof(m_nSrcId), NULL);
		if (FAILED(hResult))
			goto Cleanup;

		hResult = stm.pstm->Read(&m_pSrcBar, sizeof(m_pSrcBar), NULL);
		if (FAILED(hResult))
			goto Cleanup;

		hResult = m_theTreeCtrl.ActiveBar()->QueryInterface(IID_IBarPrivate, (void**)&pPrivate);
		if (FAILED(hResult))
			goto Cleanup;

		switch (m_nIndex)
		{
		case CToolDataObject::eTool:
			{

				hResult = m_theTreeCtrl.ActiveBar()->get_Tools((Tools **)&pTools);
				if (SUCCEEDED(hResult))
				{
					hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
					for (int nTool = 0; nTool < nCount; nTool++)
					{
						hResult = pTools->CreateTool(&pTool);
						if (FAILED(hResult))
						{
							assert(FALSE);
							continue;
						}

						hResult = pTool->DragDropExchange(stm.pstm, VARIANT_FALSE);
						if (FAILED(hResult))
						{
							assert(FALSE);
							pTool->Release();
							continue;
						}

						switch (nNodeType)
						{
						case CDesignerTree::eBand:
						case CDesignerTree::eMenu:
						case CDesignerTree::ePopup:
						case CDesignerTree::eStatusBar:
						case CDesignerTree::eChildMenu:
						case CDesignerTree::eChildBand:
							{
								//
								// Tool Dropped on a band
								//

								IBand* pBand = reinterpret_cast<IBand*>(m_theTreeCtrl.GetCurrentTreeViewItem().lParam);
								assert(pBand);
								if (pBand)
								{
									ITools* pBandTools;
									hResult = pBand->get_Tools((Tools**)&pBandTools);
									if (SUCCEEDED(hResult))
									{
										hResult = pBandTools->InsertTool(-1, pTool, VARIANT_FALSE);
										if (SUCCEEDED(hResult))
										{
											FeedBack(nNodeType, pdwEffect, grfKeyState, m_fe[m_nIndex].cfFormat);
											hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
											bUpdate = TRUE;
										}
										pBandTools->Release();
									}
								}
							}
							break;

					case CDesignerTree::eButton:
					case CDesignerTree::eDropDown:
					case CDesignerTree::eTextbox:
					case CDesignerTree::eCombobox:
					case CDesignerTree::eSeparator:
					case CDesignerTree::eLabel:
					case CDesignerTree::eWindowList:
					case CDesignerTree::eCustomTool:
					case CDesignerTree::eForm:
						{
							HTREEITEM hTool = m_theTreeCtrl.GetCurrentItem();
							if (NULL == hTool)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							TV_ITEM tvBand;
							tvBand.hItem = TreeView_GetParent(m_theTreeCtrl.hWnd(), hTool);
							if (NULL == tvBand.hItem)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							tvBand.mask = TVIF_PARAM|TVIF_IMAGE;
							if (!TreeView_GetItem(m_theTreeCtrl.hWnd(), &tvBand))
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							switch (tvBand.iImage)
							{
							case CDesignerTree::eBand:
							case CDesignerTree::eMenu:
							case CDesignerTree::ePopup:
							case CDesignerTree::eStatusBar:
							case CDesignerTree::eChildMenu:
							case CDesignerTree::eChildBand:
								break;

							default:
								hResult = E_FAIL;
								goto Cleanup;
							}

							pBand = reinterpret_cast<IBand*>(tvBand.lParam);
							if (NULL == pBand)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}
						
							int nIndex = m_theTreeCtrl.FindIndex(tvBand.hItem, hTool);

							ITools* pBandTools;
							hResult = pBand->get_Tools((Tools**)&pBandTools);
							if (FAILED(hResult))
								goto Cleanup;

							hResult = pBandTools->InsertTool(nIndex + nTool, pTool, VARIANT_TRUE);
							if (SUCCEEDED(hResult) && !bUpdate)
							{
								FeedBack(nNodeType, pdwEffect, grfKeyState, m_fe[m_nIndex].cfFormat);
								bUpdate = TRUE;
								hRefreshItem = tvBand.hItem;
							}
							pBandTools->Release();
						}
						break;

					case CDesignerTree::eCategory:
							{
								//
								// Tool Dropped on a Category
								//

								CatEntry* pCateEntry = reinterpret_cast<CatEntry*>(m_theTreeCtrl.GetCurrentTreeViewItem().lParam);
								if (pCateEntry)
								{
									// Same ActiveBar
									if (m_pSrcBar == m_theTreeCtrl.ActiveBar())
									{
										if (CToolDataObject::eDesignerMainToolDragDropId == m_nSrcId)
											*pdwEffect = DROPEFFECT_MOVE;
										else
											*pdwEffect = DROPEFFECT_COPY;
										pCateEntry->AddTool(pTool);
										bUpdate = TRUE;
										hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
									}
									else
									{
										// different activeBar
										hResult = pTools->InsertTool(-1, pTool, VARIANT_FALSE);
										if (SUCCEEDED(hResult))
										{
											pCateEntry->AddTool(pTool);
											*pdwEffect = DROPEFFECT_COPY;
											bUpdate = TRUE;
											hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
										}
									}
								}
							}
							break;
						}
						pTool->Release();
					}
					pTools->Release();
				}
			}
			break;

		case CToolDataObject::eToolId:
			{
				if (m_pSrcBar != m_theTreeCtrl.ActiveBar())
				{
					m_theTreeCtrl.MessageBox(_T("Different ActiveBar"));
					goto Cleanup;
				}

				hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
				for (int nTool = 0; nTool < nCount; nTool++)
				{
					hResult = pPrivate->ExchangeToolByIdentity((LPDISPATCH)stm.pstm, 
															   VARIANT_FALSE, 
															   (IDispatch**)&pTool);
					if (FAILED(hResult))
						continue;

					switch (nNodeType)
					{
					case CDesignerTree::eBand:
					case CDesignerTree::eMenu:
					case CDesignerTree::ePopup:
					case CDesignerTree::eStatusBar:
					case CDesignerTree::eChildMenu:
					case CDesignerTree::eChildBand:
						{
							//
							// Tool Dropped on a band
							//

							IBand* pBand = reinterpret_cast<IBand*>(m_theTreeCtrl.GetCurrentTreeViewItem().lParam);
							if (pBand)
							{
								ITools* pBandTools;
								hResult = pBand->get_Tools((Tools**)&pBandTools);
								if (SUCCEEDED(hResult))
								{
									ITool* pToolClone;
									hResult = pTool->Clone(&pToolClone);
									pTool->Release();
									pTool = pToolClone;

									hResult = pBandTools->InsertTool(-1, pTool, VARIANT_FALSE);
									if (SUCCEEDED(hResult) && !bUpdate)
									{
										FeedBack(nNodeType, pdwEffect, grfKeyState, m_fe[m_nIndex].cfFormat);
										bUpdate = TRUE;
										hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
									}
									pBandTools->Release();
								}
							}
						}
						break;

					case CDesignerTree::eButton:
					case CDesignerTree::eDropDown:
					case CDesignerTree::eTextbox:
					case CDesignerTree::eCombobox:
					case CDesignerTree::eSeparator:
					case CDesignerTree::eLabel:
					case CDesignerTree::eWindowList:
					case CDesignerTree::eCustomTool:
					case CDesignerTree::eForm:
						{
							HTREEITEM hTool = m_theTreeCtrl.GetCurrentItem();
							if (NULL == hTool)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							TV_ITEM tvBand;
							tvBand.hItem = TreeView_GetParent(m_theTreeCtrl.hWnd(), hTool);
							if (NULL == tvBand.hItem)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							tvBand.mask = TVIF_PARAM|TVIF_IMAGE;
							if (!TreeView_GetItem(m_theTreeCtrl.hWnd(), &tvBand))
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							switch (tvBand.iImage)
							{
							case CDesignerTree::eBand:
							case CDesignerTree::eMenu:
							case CDesignerTree::ePopup:
							case CDesignerTree::eStatusBar:
							case CDesignerTree::eChildMenu:
							case CDesignerTree::eChildBand:
								break;

							default:
								hResult = E_FAIL;
								goto Cleanup;
							}

							pBand = reinterpret_cast<IBand*>(tvBand.lParam);
							if (NULL == pBand)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}
						
							int nIndex = m_theTreeCtrl.FindIndex(tvBand.hItem, hTool);

							hResult = pBand->get_Tools((Tools**)&pTools);
							if (FAILED(hResult))
								goto Cleanup;

							ITool* pToolClone;
							hResult = pTool->Clone(&pToolClone);
							pTool->Release();
							pTool = pToolClone;

							hResult = pTools->InsertTool(nIndex + nTool, pTool, VARIANT_TRUE);
							if (SUCCEEDED(hResult) && !bUpdate)
							{
								FeedBack(nNodeType, pdwEffect, grfKeyState, m_fe[m_nIndex].cfFormat);
								bUpdate = TRUE;
								hRefreshItem = tvBand.hItem;
							}
							pTools->Release();
						}
						break;

					case CDesignerTree::eCategory:
						{
							//
							// Tool Dropped on a Category
							//

							CatEntry* pCateEntry = reinterpret_cast<CatEntry*>(m_theTreeCtrl.GetCurrentTreeViewItem().lParam);
							assert(pCateEntry);
							if (pCateEntry)
							{
								if (m_nSrcId == CToolDataObject::eLibraryDragDropId && CToolDataObject::eLibraryDragDropId != m_nDestId)
								{
									//
									// Tool came from the library and we are not the library
									//

									hResult = m_theTreeCtrl.ActiveBar()->get_Tools((Tools **)&pTools);
									if (SUCCEEDED(hResult))
									{
										hResult = pTools->InsertTool(-1, pTool, VARIANT_TRUE);
										if (SUCCEEDED(hResult))
										{
											pCateEntry->AddTool(pTool);
											if (!bUpdate)
											{
												*pdwEffect = DROPEFFECT_COPY;
												bUpdate = TRUE;
												hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
											}
										}
										pTools->Release();
									}
								}
								else if (pCateEntry->AddTool(pTool))
								{
									//
									// Tool came from ActiveBar
									//

									if (!bUpdate)
									{
										if (MK_CONTROL & grfKeyState)
											*pdwEffect = DROPEFFECT_COPY;
										else
											*pdwEffect = DROPEFFECT_MOVE;
										bUpdate = TRUE;
										hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
									}
								}
								else
								{
									assert(FALSE);
								}
							}
						}
						break;
					}
					pTool->Release();
				}
			}
			break;

		case CToolDataObject::eBandToolId:
		case CToolDataObject::eBandChildBandToolId:
			{
				if (m_pSrcBar != m_theTreeCtrl.ActiveBar())
				{
					m_theTreeCtrl.MessageBox(_T("Different ActiveBar"));
					goto Cleanup;
				}
				BSTR bstrBand = NULL;
				BSTR bstrChildBand = NULL;
				hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
				for (int nTool = 0; nTool < nCount; nTool++)
				{
					hResult = pPrivate->ExchangeToolByBandChildBandToolIdentity((LPDISPATCH)stm.pstm, 
																			    bstrBand, 
																			    bstrChildBand, 
																			    VARIANT_FALSE, 
																			    (IDispatch**)&pTool);
					if (FAILED(hResult))
						continue;

					SysFreeString(bstrBand);
					SysFreeString(bstrChildBand);

					if (m_pSrcBar != m_theTreeCtrl.ActiveBar())
					{
						m_theTreeCtrl.MessageBox(_T("Different ActiveBar"));

						hResult = E_FAIL;
						pTool->Release();
						goto Cleanup;
					}

					switch (nNodeType)
					{
					case CDesignerTree::eBand:
					case CDesignerTree::eMenu:
					case CDesignerTree::ePopup:
					case CDesignerTree::eStatusBar:
					case CDesignerTree::eChildMenu:
					case CDesignerTree::eChildBand:
						{
							//
							// Tool Dropped on a band
							//

							IBand* pBand = reinterpret_cast<IBand*>(m_theTreeCtrl.GetCurrentTreeViewItem().lParam);
							if (pBand)
							{
								//
								// Tool Dropped on a band
								//

								IBand* pBand = reinterpret_cast<IBand*>(m_theTreeCtrl.GetCurrentTreeViewItem().lParam);
								if (pBand)
								{
									hResult = pBand->get_Tools((Tools**)&pTools);
									if (SUCCEEDED(hResult))
									{
										hResult = pTools->InsertTool(-1, pTool, VARIANT_TRUE);
										if (SUCCEEDED(hResult) && !bUpdate)
										{
											FeedBack(nNodeType, pdwEffect, grfKeyState, m_fe[m_nIndex].cfFormat);
											bUpdate = TRUE;
											hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
										}
										pTools->Release();
									}
								}
							}
						}
						break;

					case CDesignerTree::eButton:
					case CDesignerTree::eDropDown:
					case CDesignerTree::eTextbox:
					case CDesignerTree::eCombobox:
					case CDesignerTree::eSeparator:
					case CDesignerTree::eLabel:
					case CDesignerTree::eWindowList:
					case CDesignerTree::eCustomTool:
					case CDesignerTree::eForm:
						{
							HTREEITEM hTool = m_theTreeCtrl.GetCurrentItem();
							if (NULL == hTool)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							TV_ITEM tvBand;
							tvBand.hItem = TreeView_GetParent(m_theTreeCtrl.hWnd(), hTool);
							if (NULL == tvBand.hItem)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							tvBand.mask = TVIF_PARAM|TVIF_IMAGE;
							if (!TreeView_GetItem(m_theTreeCtrl.hWnd(), &tvBand))
							{
								hResult = E_FAIL;
								goto Cleanup;
							}

							switch (tvBand.iImage)
							{
							case CDesignerTree::eBand:
							case CDesignerTree::eMenu:
							case CDesignerTree::ePopup:
							case CDesignerTree::eStatusBar:
							case CDesignerTree::eChildMenu:
							case CDesignerTree::eChildBand:
								break;

							default:
								hResult = E_FAIL;
								goto Cleanup;
							}

							pBand = reinterpret_cast<IBand*>(tvBand.lParam);
							if (NULL == pBand)
							{
								hResult = E_FAIL;
								goto Cleanup;
							}
						
							int nIndex = m_theTreeCtrl.FindIndex(tvBand.hItem, hTool);

							hResult = pBand->get_Tools((Tools**)&pTools);
							if (FAILED(hResult))
								goto Cleanup;

							hResult = pTools->InsertTool(nIndex + nTool, pTool, VARIANT_TRUE);
							if (SUCCEEDED(hResult) && !bUpdate)
							{
								FeedBack(nNodeType, pdwEffect, grfKeyState, m_fe[m_nIndex].cfFormat);
								hRefreshItem = tvBand.hItem;
								bUpdate = TRUE;
							}
							pTools->Release();
						}
						break;

					case CDesignerTree::eCategory:
						{
							//
							// Tool Dropped on a Category
							//

							CatEntry* pCateEntry = reinterpret_cast<CatEntry*>(m_theTreeCtrl.GetCurrentTreeViewItem().lParam);
							if (pCateEntry)
							{
								if (pCateEntry->AddTool(pTool))
								{
									if (!bUpdate)
									{
										FeedBack(nNodeType, pdwEffect, grfKeyState, m_fe[m_nIndex].cfFormat);
										hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
										bUpdate = TRUE;
									}
									ITools* pTools;
									ITool* pTool2;
									hResult = m_theTreeCtrl.ActiveBar()->get_Tools((Tools**)&pTools);
									if (SUCCEEDED(hResult))
									{
										BOOL bFound = FALSE;
										short nCount;
										hResult = pTools->Count(&nCount);
										if (SUCCEEDED(hResult))
										{
											VARIANT vIndex;
											vIndex.vt = VT_I2;
											for (vIndex.iVal = 0; vIndex.iVal < nCount; vIndex.iVal++)
											{
												hResult = pTools->Item(&vIndex, (Tool**)&pTool2);
												if (SUCCEEDED(hResult) && pTool == pTool2)
												{
													pTool2->Release();
													bFound = TRUE;
													break;
												}
												pTool2->Release();
											}
										}
										if (!bFound)
											pTools->InsertTool(-1, pTool, FALSE);
										pTools->Release();
									}
								}
							}
						}
						break;
					}
					pTool->Release();
				}
			}
			break;

		case CToolDataObject::eBand:
			{
				switch (nNodeType)
				{
				case CDesignerTree::eBands:
					{
						IBands* pBands;
						hResult = m_theTreeCtrl.ActiveBar()->get_Bands((Bands**)&pBands);
						if (SUCCEEDED(hResult))
						{
							hResult = pBands->Add(L"Temp", (Band**)&pBand);
							pBands->Release();
							if (FAILED(hResult))
								goto Cleanup;

							hResult = pBand->DragDropExchange(stm.pstm, VARIANT_FALSE);
							pBand->Release();
							if (FAILED(hResult))
								goto Cleanup;

							if (!bUpdate)
							{
								*pdwEffect = DROPEFFECT_COPY;
								bUpdate = TRUE;
								hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
							}
						}
					}
					break;

				case CDesignerTree::eChildBands:
					{
						IChildBands* pChildBands = reinterpret_cast<IChildBands*>(m_theTreeCtrl.GetCurrentTreeViewItem().lParam);
						if (pChildBands)
						{
							hResult = pChildBands->Add(L"Temp", (Band**)&pBand);
							if (FAILED(hResult))
								goto Cleanup;

							hResult = pBand->DragDropExchange(stm.pstm, VARIANT_FALSE);
							pBand->Release();
							if (FAILED(hResult))
								goto Cleanup;

							if (!bUpdate)
							{
								*pdwEffect = DROPEFFECT_COPY;
								bUpdate = TRUE;
								hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
							}
						}
					}
					break;
				}
			}
			break;
		
		case CToolDataObject::eCategory:
			{
				hResult = m_theTreeCtrl.ActiveBar()->get_Tools((Tools**)&pTools);
				if (SUCCEEDED(hResult))
				{
					BSTR bstrCategory = NULL;
					hResult = StReadBSTR(stm.pstm, bstrCategory);
					if (FAILED(hResult))
						goto Cleanup;

					int nCount;
					hResult = stm.pstm->Read(&nCount, sizeof(nCount), NULL);
					for (int nTool = 0; nTool < nCount; nTool++)
					{
						hResult = pTools->CreateTool(&pTool);
						if (FAILED(hResult))
							continue;

						hResult = pTool->DragDropExchange(stm.pstm, VARIANT_FALSE);
						if (FAILED(hResult))
						{
							pTool->Release();
							continue;
						}

						switch (nNodeType)
						{
						case CDesignerTree::eCategories:
							{
								CatEntry* pCatEntry = m_theTreeCtrl.GetCategoryMgr()->CreateEntry(bstrCategory); 
								if (pCatEntry)
								{
									if (pCatEntry->AddTool(pTool))
									{
										hResult = pTools->Insert(-1, (Tool*)pTool);
										if (SUCCEEDED(hResult) && !bUpdate)
										{
											*pdwEffect = DROPEFFECT_COPY;
											bUpdate = TRUE;
											hRefreshItem = m_theTreeCtrl.GetCurrentTreeViewItem().hItem;
										}
									}
								}
							}
							break;
						}
						pTool->Release();
					}
					SysFreeString(bstrCategory);
					pTools->Release();
				}
			}
			break;
		}
	}
	CATCH 
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}

Cleanup:
	if (pPrivate)
		pPrivate->Release();

	if (m_pDataObject)
	{
		m_pDataObject->Release();
		m_pDataObject = NULL;
	}

	if (stm.pstm)
		stm.pstm->Release();

	m_theTreeCtrl.SelectDropTarget(NULL);
	
	if (bUpdate)
	{
		m_theTreeCtrl.Refresh(hRefreshItem);
		m_theTreeCtrl.Update();
	}

	return hResult;
}

//
// CTreeCtrlSource
//

//
// QueryInterface
//

HRESULT __stdcall CTreeCtrlSource::QueryInterface(REFIID riid, void** ppvObject)
{
	if (riid == IID_IDropSource || riid == IID_IUnknown)
	{
		*ppvObject = static_cast<IDropSource*>(this);
		AddRef();
		return S_OK;
	}
	*ppvObject = 0;
	return E_NOINTERFACE;
}
        
//
// AddRef
//

ULONG __stdcall CTreeCtrlSource::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

//
// Release
//

ULONG __stdcall CTreeCtrlSource::Release()
{
	if (InterlockedDecrement(&m_cRef) == 0)
	{
		delete this;
		return 0;
	}
	return m_cRef;
}

//
// QueryContinueDrag
//

HRESULT __stdcall CTreeCtrlSource::QueryContinueDrag(BOOL  bEscapePressed, DWORD grfKeyState)
{
	if (bEscapePressed)
		return DRAGDROP_S_CANCEL;

	if (!(MK_LBUTTON & grfKeyState))
		return DRAGDROP_S_DROP;

	return S_OK;
}
        
//
// GiveFeedback
//

HRESULT __stdcall CTreeCtrlSource::GiveFeedback(DWORD dwEffect)
{
	switch (dwEffect)
	{
	case DROPEFFECT_COPY:
		SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOL)));
		return S_OK;

	case DROPEFFECT_MOVE:
		SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOLMOVE)));
		return S_OK;

	case DROPEFFECT_NONE:
		SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDCC_DRAGTOOLNOT)));
		return S_OK;
	}
	return DRAGDROP_S_USEDEFAULTCURSORS;
}

