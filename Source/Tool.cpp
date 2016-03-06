//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <MultiMon.h>
#include <OleCtl.h>
#include <RichEdit.h>
#include "IPServer.h"
#include <stddef.h>       // for offsetof()
#include "Debug.h"
#include "Errors.h"
#include "Globals.h"
#include "Dispids.h"
#include "Resource.h"
#include "CBList.h"
#include "Utility.h" 
#include "Flicker.h"
#include "Localizer.h"
#include "Dib.h"
#include "Bar.h"
#include "Dock.h"
#include "Band.h"
#include "PopupWin.h"
#include "MiniWin.h"
#include "ChildBands.h"
#include "ImageMgr.h"
#include "Customproxy.h"
#include "ShortCuts.h"
#include "TpPopup.h"
#ifdef _SCRIPTSTRING
#include "ScriptString.h"
#endif
#include "Bands.h"
#include "Tool.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;

static void DebugPrintBitmap(HBITMAP hBitmap, int nOffsetY = 0)
{
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		HDC hDCSrc = CreateCompatibleDC(hDC);
		HBITMAP hBitmapOld = SelectBitmap(hDCSrc, hBitmap);
		BITMAP bmInfo;
		GetObject(hBitmap, sizeof(bmInfo), &bmInfo);
		BitBlt(hDC, 
			   0, 
			   nOffsetY, 
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
}

#endif

//
// PaintDropDownBitmap
//

static void PaintDropDownBitmap(HDC hDC, int x, int y, HBITMAP hBitmap, COLORREF crMask, BOOL bDouble = FALSE)
{
	HDC hMemDC = GetMemDC();
	if (hMemDC)
	{
		HBITMAP hBitmapOld = SelectBitmap(hMemDC, hBitmap);

		SetTextColor(hDC, RGB(0, 0, 0));	
		SetBkColor(hDC, RGB(255, 255, 255));

		BOOL bResult = StretchBlt(hDC, x, y, (bDouble ? 10 : 5), (bDouble ? 6 : 3), hMemDC, 0, 0, 5, 3, SRCAND);
		
		SetBkColor(hDC, RGB(0, 0, 0));
		SetTextColor(hDC, crMask);

		bResult = StretchBlt(hDC, x, y, (bDouble ? 10 : 5), (bDouble ? 6 : 3), hMemDC, 0, 0, 5, 3, SRCPAINT);
		
		SelectBitmap(hMemDC, hBitmapOld);
	}
}

//
// CToolShortCut
//

CToolShortCut::~CToolShortCut()
{
	RemoveAll();
}

HRESULT CToolShortCut::RemoveAll()
{
	try
	{
		ShortCutStore* pShortCutStore;
		int nCount = m_aShortCutStores.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pShortCutStore = m_aShortCutStores.GetAt(nIndex);
			if (m_pBar)
				m_pBar->m_pShortCuts->Remove(pShortCutStore, m_pTool);
			pShortCutStore->Release();
		}
		m_aShortCutStores.RemoveAll();
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

HRESULT CToolShortCut::Add(ShortCutStore* pShortCutStore)
{
	try
	{
		ShortCutStore* pShortCutTool;
		HRESULT hResult = pShortCutStore->Clone((IShortCut**)&pShortCutTool);
		if (FAILED(hResult))
			return hResult;

		hResult = m_aShortCutStores.Add(pShortCutTool);
		if (FAILED(hResult))
			return hResult;

		if (m_pBar)
		{
			pShortCutTool->m_pBar = m_pBar;
			hResult = m_pBar->m_pShortCuts->Add(pShortCutStore, m_pTool);
			if (FAILED(hResult))
				return hResult;
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

HRESULT CToolShortCut::CopyTo(CToolShortCut& theToolShortCutIn)
{
	if (&theToolShortCutIn == this)
		return NOERROR;
	try
	{
		theToolShortCutIn.RemoveAll();
		ShortCutStore* pShortCutStore;
		int nCount = GetCount();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pShortCutStore = Get(nIndex);
			if (pShortCutStore)
			{
				pShortCutStore->m_pBar = m_pBar;
				theToolShortCutIn.Add(pShortCutStore);
			}
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

HRESULT CToolShortCut::Remove(ShortCutStore* pShortCutStore)
{
	try
	{
		ShortCutStore* pDeleteShortCutStore;
		int nCount = m_aShortCutStores.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pDeleteShortCutStore = m_aShortCutStores.GetAt(nIndex);
			if (*pShortCutStore == *pDeleteShortCutStore)
			{
				m_aShortCutStores.RemoveAt(nIndex);
				if (m_pBar)
					m_pBar->m_pShortCuts->Remove(pShortCutStore, m_pTool);
				pDeleteShortCutStore->Release();
				break;
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

HRESULT CToolShortCut::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		ShortCutStore* pShortCutStore;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			int nCount = m_aShortCutStores.GetSize();
			
			hResult = pStream->Write(&nCount,sizeof(nCount),NULL);
			if (FAILED(hResult))
				return hResult;

			for (int nIndex = 0; nIndex < nCount; nIndex++)
			{
				pShortCutStore = m_aShortCutStores.GetAt(nIndex);

				hResult = pShortCutStore->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
				RemoveAll();

				int nCount;
				hResult = pStream->Read(&nCount,sizeof(nCount),NULL);
				if (FAILED(hResult))
					return hResult;

				if (nCount > 1000)
					return E_FAIL;

				for (int nIndex = 0; nIndex < nCount; nIndex++)
				{
					pShortCutStore = ShortCutStore::CreateInstance(NULL);
					if (NULL == pShortCutStore)
						return E_OUTOFMEMORY;

					hResult = pShortCutStore->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;

					if (m_pBar)
					{
						pShortCutStore->m_pBar = m_pBar;

						hResult = m_pBar->m_pShortCuts->Add(pShortCutStore, m_pTool);
						if (FAILED(hResult))
							return hResult;
					}
					
					hResult = m_aShortCutStores.Add(pShortCutStore);
					if (FAILED(hResult))
						return hResult;
				}
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

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CTool(IUnknown *pUnkOuter)
{
	return (IUnknown *)(ITool *)new CTool();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_Tool,Tool,_T("Tool"),CreateFN_CTool,2,0,&IID_ITool,_T("Tool Object"));
void *CTool::objectDef=&ToolObject;
CTool *CTool::CreateInstance(IUnknown *pUnkOuter)
{
	return new CTool();
}
//{OBJECT CREATEFN}
CTool::CTool()
	: m_refCount(1)
	  ,m_pDispCustom(NULL),
	  m_bstrName(NULL),
	  m_bstrText(NULL),
	  m_bstrSubBand(NULL),
	  m_bstrCaption(NULL),
	  m_bstrCategory(NULL),
	  m_bstrToolTipText(NULL),
	  m_bstrDescription(NULL),
	  m_hWndActive(NULL),
	  m_hWndParent(NULL),
	  m_hWndPrevParent(NULL),
	  m_pBar(NULL),
	  m_pBand(NULL),
	  m_pTools(NULL),
	  m_pComboList(NULL),
	  m_pActiveEdit(NULL),
	  m_pActiveCombo(NULL),
	  m_szCleanCaptionNormal(NULL),
	  m_szCleanCaptionPopup(NULL),
	  m_pSubBandWindow(NULL)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
	VariantInit(&m_vTag);
	m_nCachedWidth = -1;
	m_nCachedHeight = -1;
	m_nComboNameOffset = -1;
	m_nShortCutOffset = 0;
	m_sizeImage.cx = CBar::eDefaultBarIconWidth;
	m_sizeImage.cy = CBar::eDefaultBarIconHeight;
	m_bDropDownPressed = FALSE;
	m_bPaintVisible = FALSE;
	m_bGrayImage = FALSE;
	m_bPressed = FALSE;
	m_bClone = FALSE;
	m_bFocus = FALSE;
	m_dwCustomFlag = 0;
	m_nCustomWidth = 0;
	m_bVerticalSeparator = FALSE;
	m_bShowTool = TRUE;
	m_bAdjusted = FALSE;
	m_bMFU = FALSE;
	m_bCreatedInternally = FALSE;
	m_bMoreTools = TRUE;
	m_dwStyleOld = 0;
	m_dwExStyleOld = 0;
	m_nImageLoad = -1;
	m_bMDICloseEnabled = TRUE;
	m_mdibActive = eNone;
}

CTool::~CTool()
{
	try
	{
		if (m_pBar && this == m_pBar->m_pActiveTool && tpV1.m_nToolId != CBar::eToolIdWindowList)
			m_pBar->m_pActiveTool = NULL;
	}
	catch (...)
	{
	}
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
// {END CLEANUP}
	SysFreeString(m_bstrName);
	SysFreeString(m_bstrToolTipText);
	SysFreeString(m_bstrCaption);
	SysFreeString(m_bstrDescription);
	SysFreeString(m_bstrCategory);
	SysFreeString(m_bstrText);
	SysFreeString(m_bstrSubBand);
	
	ReleaseImages();

	if (m_pDispCustom)
		put_Custom(NULL);

	if (m_pComboList)
		m_pComboList->Release();
	
	m_scTool.RemoveAll();

	VariantClear(&m_vTag);
	if (m_szCleanCaptionNormal)
		free(m_szCleanCaptionNormal);
	if (m_szCleanCaptionPopup)
		free(m_szCleanCaptionPopup);
}

#ifdef _DEBUG
void CTool::Dump(DumpContext& dc)
{
	MAKE_TCHARPTR_FROMWIDE(szName, m_bstrName);
	MAKE_TCHARPTR_FROMWIDE(szCaption, m_bstrCaption);
	_stprintf(dc.m_szBuffer,
			  _T("\tTool Name: %s Caption: %s\n"), 
			  szName, szCaption);
	dc.Write();
	_stprintf(dc.m_szBuffer,
			  _T("\tImage Ids Normal: %li, Pressed: %li, MouseOver: %li, Disabled: %li\n"), 
			  tpV1.m_nImageIds[0], tpV1.m_nImageIds[1], tpV1.m_nImageIds[2], tpV1.m_nImageIds[3]);
	dc.Write();
}
#endif

CTool::ToolPropV1::ToolPropV1()
{
	m_ttTools = ddTTButton;
	m_tsStyle = ddSStandard;
	m_cpTools = ddCPStandard;
	m_taTools = ddACenterCenter;
	m_lsLabelStyle = ddLSNormal;
	m_nWidth = -1;
	m_nHeight = -1;
	m_nToolId = 0;
	m_nUsage = 0;
	m_nHelpContextId = 0;
	m_vbVisible = VARIANT_TRUE;
	m_vbEnabled = VARIANT_TRUE;
	m_vbChecked = VARIANT_FALSE;
	m_vbDefault = VARIANT_FALSE;
	m_vbAutoRepeat = VARIANT_FALSE;
	m_vbDesignerCreated = VARIANT_FALSE;
	m_nAutoRepeatInterval = eAutoRepeatTime;
	m_mvMenuVisibility = ddMVAlwaysVisible;
	m_nImageWidth = -1;
	m_nImageHeight = -1;
	m_lsLabelBevel = ddLBFlat;
	m_nSelStart = 0;
	m_nSelLength = -1;
	m_asAutoSize = ddTASNone;

	for (int nIndex = 0; nIndex < eNumOfImageTypes; nIndex++)
		m_nImageIds[nIndex] = -1L;

	m_dwIdentity = 0;
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CTool::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IUnknown))
	{
		AddRef();
		*ppvObjOut=this;
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IDispatch))
	{
		*ppvObjOut=(void *)(IDispatch *)(ITool *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ITool))
	{
		*ppvObjOut=(void *)(ITool *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISupportErrorInfo))
	{
		*ppvObjOut=(void *)(ISupportErrorInfo *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ICustomHost))
	{
		*ppvObjOut=(void *)(ICustomHost *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPerPropertyBrowsing))
	{
		*ppvObjOut=(void *)(IPerPropertyBrowsing *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IDDPerPropertyBrowsing))
	{
		*ppvObjOut=(void *)(IDDPerPropertyBrowsing *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ICategorizeProperties))
	{
		*ppvObjOut=(void *)(ICategorizeProperties *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
STDMETHODIMP_(ULONG) CTool::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CTool::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	try
	{
		delete this;
	}
	catch(...)
	{
	}
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
STDMETHODIMP CTool::get_CBNewIndex(long *retval)
{
	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;

	*retval = pList->m_nNewIndex;
	return NOERROR;
}
STDMETHODIMP CTool::get_CBLines(short *retval)
{
	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;

	*retval = pList->lpV1.m_nLines;
	return NOERROR;
}
STDMETHODIMP CTool::put_CBLines(short val)
{
	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;

	pList->lpV1.m_nLines = val;
	return NOERROR;
}
STDMETHODIMP CTool::_CBAddItem(BSTR bstrItem, long Index)
{
	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;
	
	return pList->AddItem(bstrItem);
}
STDMETHODIMP CTool::get_SelStart(short *retval)
{
	if (!(ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
		return E_FAIL;

	if (m_hWndActive && IsWindow(m_hWndActive))
	{
		if (ddTTCombobox == tpV1.m_ttTools)
			SendMessage(m_hWndActive, CB_GETEDITSEL, (WPARAM)&tpV1.m_nSelStart, (LPARAM)&tpV1.m_nSelLength);
		else
			SendMessage(m_hWndActive, EM_GETSEL, (WPARAM)&tpV1.m_nSelStart, (LPARAM)&tpV1.m_nSelLength);
	}
	*retval = tpV1.m_nSelStart;
	return NOERROR;
}
STDMETHODIMP CTool::put_SelStart(short val)
{
	if (!(ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
		return E_FAIL;

	tpV1.m_nSelStart = val;
	if (m_hWndActive && IsWindow(m_hWndActive))
	{
		if (ddTTCombobox == tpV1.m_ttTools)
			SendMessage(m_hWndActive, CB_SETEDITSEL, tpV1.m_nSelStart, tpV1.m_nSelLength);
		else
			SendMessage(m_hWndActive, EM_SETSEL, tpV1.m_nSelStart, tpV1.m_nSelLength);
	}

	return NOERROR;
}
STDMETHODIMP CTool::get_SelLength(short *retval)
{
	if (!(ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
		return E_FAIL;

	if (NULL != m_hWndActive && IsWindow(m_hWndActive))
	{
		if (ddTTCombobox == tpV1.m_ttTools)
			SendMessage(m_hWndActive, CB_GETEDITSEL, (WPARAM)&tpV1.m_nSelStart, (LPARAM)&tpV1.m_nSelLength);
		else
			SendMessage(m_hWndActive, EM_GETSEL, (WPARAM)&tpV1.m_nSelStart, (LPARAM)&tpV1.m_nSelLength);
	}
	*retval = tpV1.m_nSelLength;

	return NOERROR;
}
STDMETHODIMP CTool::put_SelLength(short val)
{
	if (!(ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
		return E_FAIL;

	tpV1.m_nSelLength = val;
	if (NULL != m_hWndActive && IsWindow(m_hWndActive))
	{
		if (ddTTCombobox == tpV1.m_ttTools)
			SendMessage(m_hWndActive, CB_SETEDITSEL, tpV1.m_nSelStart, tpV1.m_nSelLength);
		else
			SendMessage(m_hWndActive, EM_SETSEL, tpV1.m_nSelStart, tpV1.m_nSelLength);
	}

	return NOERROR;
}
STDMETHODIMP CTool::get_AutoSize(ToolAutoSizeTypes *retval)
{
	*retval = tpV1.m_asAutoSize;
	return NOERROR;
}
STDMETHODIMP CTool::put_AutoSize(ToolAutoSizeTypes val)
{
	if (m_pTools && ddASNone != val)
	{
		if (m_pBar && m_pTools != m_pBar->m_pTools)
		{
			CTool* pTool;
			int nCount = m_pTools->GetToolCount();
			for (int nTool = 0; nTool < nCount; nTool++)
			{
				pTool = m_pTools->GetTool(nTool);
				if (pTool != this && ddASNone != pTool->tpV1.m_asAutoSize)
				{
					m_pBar->m_theErrorObject.SendError(IDERR_ONLYONETOOLCANHAVEAUTOSIZE, L"");
					return CUSTOM_CTL_SCODE(IDERR_ONLYONETOOLCANHAVEAUTOSIZE);
				}
			}
		}
	}
	tpV1.m_asAutoSize = val;
	return NOERROR;
}
STDMETHODIMP CTool::MapPropertyToCategory( DISPID dispid,  PROPCAT*ppropcat)
{
	return E_FAIL;
}
STDMETHODIMP CTool::GetCategoryName( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName)
{
	return E_FAIL;
}
STDMETHODIMP CTool::get_LabelBevel(LabelBevelStyles *retval)
{
	if (ddTTLabel != tpV1.m_ttTools)
		return E_FAIL;
	*retval = tpV1.m_lsLabelBevel;
	return NOERROR;
}
STDMETHODIMP CTool::put_LabelBevel(LabelBevelStyles val)
{
	if (ddTTLabel != tpV1.m_ttTools)
		return E_FAIL;

	if (val < ddLBFlat || val > ddLBInset)
		return E_INVALIDARG;

	tpV1.m_lsLabelBevel = val;
	return NOERROR;
}
STDMETHODIMP CTool::get_ImageWidth(short *retval)
{
	SIZE size;
	size.cx = tpV1.m_nImageWidth;
	if (size.cx > 0)
		PixelToTwips(&size, &size);
	*retval = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CTool::put_ImageWidth(short val)
{
	SIZE size;
	size.cx = val;
	if (val > 0)
		TwipsToPixel(&size, &size);
	tpV1.m_nImageWidth = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CTool::get_ImageHeight(short *retval)
{
	SIZE size;
	size.cx = tpV1.m_nImageHeight;
	if (size.cx > 0)
		PixelToTwips(&size, &size);
	*retval = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CTool::put_ImageHeight(short val)
{
	SIZE size;
	size.cx = val;
	if (val > 0)
		TwipsToPixel(&size, &size);
	tpV1.m_nImageHeight = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CTool::get_Left(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	if (NULL == m_pBand)
	{
		*retval = -1;
		return NOERROR;
	}
	int nIndex = m_pBand->GetIndexOfTool(this);
	CRect rcTool;
	if (!m_pBand->GetToolScreenRect(nIndex, rcTool))
	{
		m_pBar->m_theErrorObject.SendError(IDERR_WINDOWISNOTVALID, L"");
		return CUSTOM_CTL_SCODE(IDERR_WINDOWISNOTVALID);
	}
	SIZE size;
	size.cx = rcTool.left;
	PixelToTwips(&size, &size);
	*retval = size.cx;
	return NOERROR;
}
STDMETHODIMP CTool::get_Top(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	if (NULL == m_pBand)
	{
		*retval = -1;
		return NOERROR;
	}
	int nIndex = m_pBand->GetIndexOfTool(this);
	CRect rcTool;
	if (!m_pBand->GetToolScreenRect(nIndex, rcTool))
	{
		m_pBar->m_theErrorObject.SendError(IDERR_WINDOWISNOTVALID, L"");
		return CUSTOM_CTL_SCODE(IDERR_WINDOWISNOTVALID);
	}
	SIZE size;
	size.cx = rcTool.top;
	PixelToTwips(&size, &size);
	*retval = size.cx;
	return NOERROR;
}
STDMETHODIMP CTool::get_MenuVisibility(MenuVisibilityStyles *retval)
{
	*retval = tpV1.m_mvMenuVisibility;
	return NOERROR;
}
STDMETHODIMP CTool::put_MenuVisibility(MenuVisibilityStyles val)
{
	tpV1.m_mvMenuVisibility = val;
	return NOERROR;
}
STDMETHODIMP CTool::GetType(DISPID dispID, long* pnType)
{
	switch (dispID)
	{
	case DISPID_SUBBAND:
		*pnType = 1;
		return S_OK;

	case DISPID_SHORTCUTS:
		*pnType = 3;
		return S_OK;
	}
	return E_NOTIMPL;
}
STDMETHODIMP CTool::GetDisplayString( DISPID dispID, BSTR __RPC_FAR *pBstr)
{
	return E_NOTIMPL;
}
STDMETHODIMP CTool::MapPropertyToPage( DISPID dispID, CLSID __RPC_FAR *pClsid)
{
	return E_NOTIMPL;
}
STDMETHODIMP CTool::GetPredefinedStrings(DISPID dispID, CALPOLESTR __RPC_FAR *pCaStringsOut, CADWORD __RPC_FAR *pCaCookiesOut)
{
	switch (dispID)
	{
	case DISPID_SUBBAND:
		m_pBar->GetPopupBands(*pCaStringsOut);
		pCaCookiesOut->cElems = 0;
		pCaCookiesOut->pElems = NULL;
		return S_OK;
	}
	return E_NOTIMPL;
}
STDMETHODIMP CTool::GetPredefinedValue( DISPID dispID, DWORD dwCookie, VARIANT __RPC_FAR *pVarOut)
{
	return E_NOTIMPL;
}
STDMETHODIMP CTool::InterfaceSupportsErrorInfo( REFIID riid)
{
    return S_OK;
}

HBITMAP CTool::XPShiftUpShadow(HBITMAP bmSource, HBITMAP bmSourceMask, BOOL bShadow)
{
	COLORREF crTransparent = RGB(255, 0, 255);
	if (NULL == bmSource)
		return NULL;

	BOOL bResult;

	HDC hDC = GetDC(m_pBar->m_hWnd);

	BITMAP bm;                   // bitmap structure
    if (!GetObject(bmSource, sizeof(bm), (LPBYTE)&bm))
		return NULL;

	HBITMAP bmTemp = CreateCompatibleBitmap(hDC, bm.bmWidth + 2, bm.bmHeight + 2);
	HBITMAP bmReturn = CreateCompatibleBitmap(hDC, bm.bmWidth + 2, bm.bmHeight + 2);

	HDC hDCTemp = CreateCompatibleDC(hDC);
	HDC hDCSource = CreateCompatibleDC(hDC);
	HDC hDCSourceMask = CreateCompatibleDC(hDC);
	HDC hDCReturn = CreateCompatibleDC(hDC);

	HBITMAP bmTempOld = SelectBitmap(hDCTemp, bmTemp);
	HBITMAP bmSourceOld = SelectBitmap(hDCSource, bmSource);
	HBITMAP bmSourceMaskOld = SelectBitmap(hDCSourceMask, bmSourceMask);
	HBITMAP bmReturnOld = SelectBitmap(hDCReturn, bmReturn);

	HBRUSH aBrush = CreateSolidBrush(RGB(255, 0, 255));
	if (aBrush)
	{
		CRect rc;
		rc.left = rc.top = 0;
		rc.right = bm.bmWidth + 2;
		rc.bottom = bm.bmHeight + 2;

		FillRect(hDCTemp, &rc, aBrush);

		FillRect(hDCReturn, &rc, aBrush);

		bResult = DeleteBrush(aBrush);
		assert(bResult);
	}

	BitBlt(hDCTemp, 0, 0, bm.bmWidth, bm.bmHeight, hDCSource, 0, 0, SRCINVERT);
	BitBlt(hDCTemp, 0, 0, bm.bmWidth, bm.bmHeight, hDCSourceMask, 0, 0, SRCAND);
	BitBlt(hDCTemp, 0, 0, bm.bmWidth, bm.bmHeight, hDCSource, 0, 0, SRCINVERT);
	COLORREF crControlDark = GetSysColor(COLOR_3DSHADOW);
	COLORREF crPixel;
	if (bShadow)
	{
		// Add gray shadow.
		for (int nX = 0; nX < bm.bmWidth; nX++)
		{
			for (int nY = 0; nY < bm.bmHeight; nY++)
			{
				crPixel = GetPixel(hDCTemp, nX, nY);
				if (crPixel != crTransparent)
					SetPixel(hDCReturn, nX + 2, nY + 2, crControlDark);
			}
		}
	}

	for (int nX = 0; nX < bm.bmWidth; nX++)
	{
		for (int nY = 0; nY < bm.bmHeight; nY++)
		{
			crPixel = GetPixel(hDCTemp, nX, nY);
			if (crPixel != crTransparent)
				SetPixel(hDCReturn, nX, nY, crPixel);
		}
	}

	SelectBitmap(hDCTemp, bmTempOld);
	SelectBitmap(hDCSource, bmSourceOld);
	SelectBitmap(hDCSourceMask, bmSourceMaskOld);
	SelectBitmap(hDCReturn, bmReturnOld);
	
	DeleteDC(hDCTemp);
	DeleteDC(hDCSource);
	DeleteDC(hDCSourceMask);
	DeleteDC(hDCReturn);

	bResult = DeleteBitmap(bmTemp);
	assert(bResult);

	ReleaseDC(m_pBar->m_hWnd, hDC);

	return bmReturn;
}

HBITMAP CTool::XPDisableImage(HBITMAP bmSource, HBITMAP bmSourceMask, BOOL bShadow)
{
	COLORREF crTransparent = RGB(255,0,255);
	if (NULL == bmSource)
		return NULL;

	BOOL bResult;
	HDC hDC = GetDC(m_pBar->m_hWnd);

	BITMAP bm;                   // bitmap structure
    if (!GetObject(bmSource, sizeof(bm), (LPBYTE)&bm))
		return NULL;

	HBITMAP bmTemp = CreateCompatibleBitmap(hDC, bm.bmWidth, bm.bmHeight);
	HBITMAP bmReturn = CreateCompatibleBitmap(hDC, bm.bmWidth, bm.bmHeight);

	HDC hDCTemp = CreateCompatibleDC(hDC);
	HDC hDCSource = CreateCompatibleDC(hDC);
	HDC hDCSourceMask = CreateCompatibleDC(hDC);
	HDC hDCReturn = CreateCompatibleDC(hDC);

	HBITMAP bmTempOld = SelectBitmap(hDCTemp, bmTemp);
	HBITMAP bmSourceOld = SelectBitmap(hDCSource, bmSource);
	HBITMAP bmSourceMaskOld = SelectBitmap(hDCSourceMask, bmSourceMask);
	HBITMAP bmReturnOld = SelectBitmap(hDCReturn, bmReturn);

	HBRUSH aBrush = CreateSolidBrush(RGB(255, 0, 255));
	if (aBrush)
	{
		CRect rc;
		rc.left = rc.top = 0;
		rc.right = bm.bmWidth;
		rc.bottom = bm.bmHeight;
		FillRect(hDCTemp, &rc, aBrush);

		FillRect(hDCReturn, &rc, aBrush);

		bResult = DeleteBrush(aBrush);
		assert(bResult);
	}

	BitBlt(hDCTemp, 0, 0, bm.bmWidth, bm.bmHeight, hDCSource, 0, 0, SRCINVERT);
	BitBlt(hDCTemp, 0, 0, bm.bmWidth, bm.bmHeight, hDCSourceMask, 0, 0, SRCAND);
	BitBlt(hDCTemp, 0, 0, bm.bmWidth, bm.bmHeight, hDCSource, 0, 0, SRCINVERT);
	COLORREF crGray = GetSysColor(COLOR_3DSHADOW);
	COLORREF crPixel;

	for (int nX = 0; nX < bm.bmWidth; nX++)
	{
		for (int nY = 0; nY < bm.bmHeight; nY++)
		{
			crPixel = GetPixel(hDCTemp, nX, nY);
			if (crPixel == RGB(255, 255, 255))
				SetPixel(hDCReturn, nX, nY, crTransparent);
			else if (crPixel != crTransparent)
				SetPixel(hDCReturn, nX, nY, crGray);
		}
	}

	SelectBitmap(hDCTemp, bmTempOld);
	SelectBitmap(hDCSource, bmSourceOld);
	SelectBitmap(hDCSourceMask, bmSourceMaskOld);
	SelectBitmap(hDCReturn, bmReturnOld);
	
	DeleteDC(hDCTemp);
	DeleteDC(hDCSource);
	DeleteDC(hDCSourceMask);
	DeleteDC(hDCReturn);

	bResult = DeleteBitmap(bmTemp);
	assert(bResult);

	ReleaseDC(m_pBar->m_hWnd, hDC);

	return bmReturn;
}

//
// DrawPict
// 

STDMETHODIMP CTool::DrawPict(OLE_HANDLE   olehDC, 
							 int		  nX,
							 int		  nY,
							 int		  nWidth,
							 int		  nHeight, 
							 VARIANT_BOOL vbEnabled)
{
	try
	{
		if (NULL == m_pBar)
			return NOERROR;

		if (NULL == m_pBar->m_pImageMgr)
			return NOERROR;

		CImageMgr& theImageMgr = *m_pBar->m_pImageMgr; 

		//
		// Find the current image index to use
		//

		ImageTypes eImageIndex = ddITNormal;
		if (m_bPressed || VARIANT_TRUE == tpV1.m_vbChecked)
		{
			if (VARIANT_TRUE == vbEnabled || Customization())
				eImageIndex = ddITPressed;
			else
				eImageIndex = ddITDisabled;
		}
		else if (VARIANT_TRUE == vbEnabled || Customization())
		{
			if (m_pBar->m_pActiveTool == this || (m_pBand && m_pBand->m_pPopupToolSelected == this))
				eImageIndex = ddITHover;
			else if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook && m_bFocus && !m_bPressed)
				eImageIndex = ddITHover;
		}
		else
			eImageIndex = ddITDisabled;


		if (-1L == tpV1.m_nImageIds[eImageIndex])
		{
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				if (ddITHover == eImageIndex)
				{
					// Generate the Hover Image and put it into the Image List
					HBITMAP bmSource;
					HRESULT hResult = get_Bitmap(ddITNormal, (OLE_HANDLE*)&bmSource);
					if (SUCCEEDED(hResult))
					{
						HBITMAP bmSourceMask;
						HRESULT hResult = get_MaskBitmap(ddITNormal, (OLE_HANDLE*)&bmSourceMask);
						if (SUCCEEDED(hResult))
						{
							HBITMAP bmReturn = XPShiftUpShadow(bmSource, bmSourceMask, TRUE);
							if (bmReturn)
							{
								hResult = put_Bitmap(ddITHover, (OLE_HANDLE)bmReturn);
								hResult = m_pBar->m_pImageMgr->put_MaskColor(tpV1.m_nImageIds[ddITHover], RGB(255, 0, 255));
								DeleteBitmap(bmReturn);
							}
							DeleteBitmap(bmSourceMask);
						}
						DeleteBitmap(bmSource);
					}
				}
				else if (ddITDisabled == eImageIndex)
				{
					// Generate the Disabled Image and put it into the Image List
					HBITMAP bmSource;
					HRESULT hResult = get_Bitmap(ddITNormal, (OLE_HANDLE*)&bmSource);
					if (SUCCEEDED(hResult))
					{
						HBITMAP bmSourceMask;
						HRESULT hResult = get_MaskBitmap(ddITNormal, (OLE_HANDLE*)&bmSourceMask);
						if (SUCCEEDED(hResult))
						{
							HBITMAP bmReturn = XPDisableImage(bmSource, bmSourceMask, TRUE);
							if (bmReturn)
							{
								hResult = put_Bitmap(ddITDisabled, (OLE_HANDLE)bmReturn);
								hResult = m_pBar->m_pImageMgr->put_MaskColor(tpV1.m_nImageIds[ddITDisabled], RGB(255, 0, 255));
								DeleteBitmap(bmReturn);
							}
							DeleteBitmap(bmSourceMask);
						}
						DeleteBitmap(bmSource);
					}
				}
				else
					eImageIndex = ddITNormal;			
			}
			else
				eImageIndex = ddITNormal;
		}

		GetGlobals().EnterPaintIcon();

		if (VARIANT_TRUE == vbEnabled || ddITDisabled == eImageIndex || Customization())
		{
			HDC hDC = (HDC)olehDC;
			COLORREF crBackOld = SetBkColor(hDC, RGB(255, 255, 255));
			COLORREF crForeOld = SetTextColor(hDC, RGB(0, 0, 0));

			ImageStyles isImage; 
			if (m_pBar->m_pActiveTool == this || !m_bGrayImage)
				isImage = ddISNormal;
			else
				isImage = ddISGray;

			if (ddITHover == eImageIndex && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				nX -= 1;
				nY -= 1;
			}

			if (m_sizeImage.cx != nWidth || m_sizeImage.cy != nHeight)
			{
				theImageMgr.ScaleBlt(olehDC, tpV1.m_nImageIds[eImageIndex], nX, nY, nWidth, nHeight, SRCINVERT, isImage);
				theImageMgr.ScaleBlt(olehDC, tpV1.m_nImageIds[eImageIndex], nX, nY,  nWidth, nHeight, SRCAND, ddISMask);
				theImageMgr.ScaleBlt(olehDC, tpV1.m_nImageIds[eImageIndex], nX, nY,  nWidth, nHeight, SRCINVERT, isImage);
			}
			else
			{
				ImageEntry* pImageEntry;
				if (theImageMgr.m_mapImages.Lookup((LPVOID)tpV1.m_nImageIds[eImageIndex], (LPVOID&)pImageEntry))
				{
					theImageMgr.BitBltExInline(olehDC, pImageEntry, nX, nY, SRCINVERT, isImage);
					theImageMgr.BitBltExInline(olehDC, pImageEntry, nX, nY, SRCAND, ddISMask);
					theImageMgr.BitBltExInline(olehDC, pImageEntry, nX, nY, SRCINVERT, isImage);
				}
			}
			SetBkColor(hDC, crBackOld);
			SetTextColor(hDC, crForeOld);
		}
		else
		{
			//
			// Disabled
			//

			if (m_sizeImage.cx != nWidth || m_sizeImage.cy != nHeight)
				theImageMgr.ScaleBlt(olehDC, tpV1.m_nImageIds[eImageIndex], nX, nY, nWidth, nHeight, SRCAND, ddISDisabled);
			else
				theImageMgr.BitBltEx(olehDC, tpV1.m_nImageIds[eImageIndex], nX, nY,SRCAND, ddISDisabled);
		}
		GetGlobals().LeavePaintIcon();
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

// IDispatch members

STDMETHODIMP CTool::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CTool::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CTool::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;
	HRESULT hr=pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgdispid);
	pTypeInfo->Release();
	return hr;
}
STDMETHODIMP CTool::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
{
#ifdef _DEBUG
	if (dispidMember<0)
		TRACE1(1, "IDispatch::Invoke -%X\n",-dispidMember)
	else
		TRACE1(1, "IDispatch::Invoke %X\n",dispidMember)
#endif
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;
	#ifdef ERRORINFO_SUPPORT
	SetErrorInfo(0L,NULL); // should be called if ISupportErrorInfo is used
	#endif
	hr = pTypeInfo->Invoke((IDispatch *)(ITool *)this, dispidMember, wFlags,
		pdispparams, pvarResult,
        pexcepinfo, puArgErr);
    pTypeInfo->Release();
	return hr;
}

// ITool members

STDMETHODIMP CTool::get_AutoRepeatInterval(short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_nAutoRepeatInterval;
	return NOERROR;
}
STDMETHODIMP CTool::put_AutoRepeatInterval(short val)
{
	if (val < 0)
		return E_INVALIDARG;

	tpV1.m_nAutoRepeatInterval = val;
	return NOERROR;
}
STDMETHODIMP CTool::get_AutoRepeat(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_vbAutoRepeat;
	return NOERROR;
}
STDMETHODIMP CTool::put_AutoRepeat(VARIANT_BOOL val)
{
	tpV1.m_vbAutoRepeat = val;
	return NOERROR;
}
STDMETHODIMP CTool::get_List(long Index, BSTR*Item)
{
	if (NULL == Item)
		return E_INVALIDARG;

	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	if (!m_pBar->AmbientUserMode())
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;
	
	VARIANT vIndex;
	vIndex.vt = VT_I4;
	vIndex.lVal = Index;
	return pList->Item(&vIndex, Item);
}
STDMETHODIMP CTool::put_List( long Index, BSTR Item)
{
	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;
	
	return pList->InsertItem(Index, Item);
}
STDMETHODIMP CTool::CBRemoveItem(long Index)
{
	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;

	VARIANT vIndex;
	vIndex.vt = VT_I4;
	vIndex.lVal = Index;
	return pList->Remove(&vIndex);
}
STDMETHODIMP CTool::get_CBListCount(short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	if (!m_pBar->AmbientUserMode())
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;

	return 	pList->Count(retval);
}
STDMETHODIMP CTool::CBClear()
{
	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;
	
	return pList->Clear();
}
STDMETHODIMP CTool::CBAddItem(BSTR bstrItem, VARIANT Index)
{
	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;
	
	if (Index.vt == VT_ERROR)
		return pList->AddItem(bstrItem);
	else if (Index.vt == VT_I2)
		return pList->InsertItem(Index.iVal, bstrItem);
	else if (Index.vt == VT_I4)
		return pList->InsertItem(Index.lVal, bstrItem);
	return E_FAIL;
}
STDMETHODIMP CTool::get_LabelStyle(LabelStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddTTLabel != tpV1.m_ttTools)
		return E_FAIL;

	*retval = tpV1.m_lsLabelStyle;
	
	return NOERROR;
}
STDMETHODIMP CTool::put_LabelStyle(LabelStyles val)
{
	if (val < ddLSNormal || val > ddLSTime)
		return E_INVALIDARG;

	tpV1.m_lsLabelStyle = val;

	TCHAR szCaption[128];
	switch (tpV1.m_lsLabelStyle)
	{
	case ddLSInsert:
		{
			long nScanCode = MapVirtualKey(VK_INSERT, 0);
			nScanCode <<= 16;
			nScanCode |= 0x01000000;
			int nResult = GetKeyNameText(nScanCode, szCaption, 40); 
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szCaption);
			tpV1.m_tsStyle = ddSText;
			put_Caption(wBuffer);
		}
		break;
	
	case ddLSCapitalLock:
		{
			long nScanCode = MapVirtualKey(VK_CAPITAL, 0);
			nScanCode <<= 16;
			int nResult = GetKeyNameText(nScanCode, szCaption, 40); 
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szCaption);
			tpV1.m_tsStyle = ddSText;
			put_Caption(wBuffer);
		}
		break;
	
	case ddLSNumberLock:
		{
			long nScanCode = MapVirtualKey(VK_NUMLOCK, 0);
			nScanCode <<= 16;
			nScanCode |= 0x01000000;
			int nResult = GetKeyNameText(nScanCode, szCaption, 40); 
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szCaption);
			tpV1.m_tsStyle = ddSText;
			put_Caption(wBuffer);
		}
		break;

	case ddLSScrollLock:
		{
			long nScanCode = MapVirtualKey(VK_SCROLL, 0);
			nScanCode <<= 16;
			int nResult = GetKeyNameText(nScanCode, szCaption, 40); 
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szCaption);
			tpV1.m_tsStyle = ddSText;
			put_Caption(wBuffer);
		}
		break;

	case ddLSDate:
		{	
			TCHAR szDate[60];
			GetLocalTime(&m_theSystemTime);
			GetDateFormat(LOCALE_SYSTEM_DEFAULT, DATE_SHORTDATE, &m_theSystemTime, NULL, szDate, 60);
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szDate);
			tpV1.m_tsStyle = ddSText;
			put_Caption(wBuffer);
		}
		break;

	case ddLSTime:
		{	
			TCHAR szTime[60];
			GetLocalTime(&m_theSystemTime);
			GetTimeFormat(LOCALE_SYSTEM_DEFAULT, TIME_NOSECONDS, &m_theSystemTime, NULL, szTime, 60);
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szTime);
			tpV1.m_tsStyle = ddSText;
			put_Caption(wBuffer);
		}
		break;
	}

	return NOERROR;
}
STDMETHODIMP CTool::get_hWnd(OLE_HANDLE *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (NULL == m_hWndActive)
	{
		*retval = (OLE_HANDLE)m_hWndActive;
		return E_FAIL;
	}
	*retval = (OLE_HANDLE)m_hWndActive;
	return NOERROR;
}
STDMETHODIMP CTool::put_hWnd(OLE_HANDLE val)
{
	if (m_hWndActive)
	{
		SetParent(NULL);
		m_hWndPrevParent = m_hWndParent = NULL;
		m_dwStyleOld = 0;
		m_dwExStyleOld = 0;
	}
	m_hWndActive = reinterpret_cast<HWND>(val);
	if (m_hWndActive)
	{
		if (!IsWindow(m_hWndActive))
			return E_FAIL;

		DWORD dwStyle = GetWindowLong(m_hWndActive, GWL_STYLE);
		m_dwStyleOld = dwStyle;
		DWORD dwExStyle = GetWindowLong(m_hWndActive, GWL_EXSTYLE);
		m_dwExStyleOld = dwExStyle;
		if (ddTTForm == tpV1.m_ttTools)
		{
			dwStyle &= ~(WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX);
			SetWindowLong(m_hWndActive, GWL_STYLE, dwStyle | WS_CHILD);
			dwExStyle &= ~WS_EX_APPWINDOW;
			SetWindowLong(m_hWndActive, GWL_EXSTYLE, dwExStyle);
		}
		else if (ddTTControl == tpV1.m_ttTools)
			SetWindowLong(m_hWndActive, GWL_STYLE, dwStyle | WS_CHILD);

		if (m_pBand)
		{
			if (m_pBand->m_pBar)
				m_pBand->m_pBar->m_theFormsAndControls.Add(this);

			if (m_pBand->m_pDock)
				SetParent(m_pBand->m_pDock->hWnd());
			else if (m_pBand->m_pFloat)
				SetParent(m_pBand->m_pFloat->hWnd());
			else
				m_hWndPrevParent = ::GetParent(m_hWndActive);
		}
		else
			m_hWndPrevParent = ::GetParent(m_hWndActive);
	}
	else
	{
		if (m_pBand && m_pBand->m_pBar)
			m_pBand->m_pBar->m_theFormsAndControls.Remove(this);
		m_hWndActive = reinterpret_cast<HWND>(val);
	}
	return NOERROR;
}
STDMETHODIMP CTool::get_TagVariant(VARIANT *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	HRESULT hResult = VariantCopy(retval, &m_vTag);
	return hResult;
}
STDMETHODIMP CTool::put_TagVariant(VARIANT val)
{
	HRESULT hResult = VariantCopy(&m_vTag, &val);
	return hResult;
}
STDMETHODIMP CTool::get_Default(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_vbDefault;
	return NOERROR;
}
STDMETHODIMP CTool::put_Default(VARIANT_BOOL val)
{
	tpV1.m_vbDefault = val;
	return NOERROR;
}
CBList *CTool::GetComboList()
{
	if (NULL == m_pComboList)
	{
		m_pComboList = CBList::CreateInstance(NULL);
		assert(m_pComboList);
		if (m_pComboList)
			m_pComboList->SetTool(this);
	}
	return m_pComboList;
}

STDMETHODIMP CTool::get_CBWidth(short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;

	SIZE size;
	size.cx = pList->lpV1.m_nWidth;
	PixelToTwips(&size, &size);
	*retval = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CTool::put_CBWidth(short val)
{
	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;
	SIZE size;
	size.cx = val;
	TwipsToPixel(&size, &size);
	pList->lpV1.m_nWidth = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CTool::get_CBListIndex(short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;
	
	if (IsWindow(m_hWndActive))
		pList->lpV1.m_nListIndex = (short)SendMessage(m_hWndActive, CB_GETCURSEL, 0, 0);
	
	*retval = pList->lpV1.m_nListIndex;
	return NOERROR;
}
STDMETHODIMP CTool::put_CBListIndex(short val)
{
	CBList* pList = GetComboList();
	if (pList)
	{
		pList->lpV1.m_nListIndex = val;
		if (IsWindow(m_hWndActive))
			SendMessage(m_hWndActive, CB_SETCURSEL, (WPARAM)(int)pList->lpV1.m_nListIndex, 0);
		else
		{
			// do windowless update
			int nElem = pList->Count();
			if (val >= 0 && val < nElem)
			{
				BSTR bstr = pList->GetName(val);
				put_Text(bstr);
			}
			else if (-1 == val)
				put_Text(L"");
			if (m_pBar)
			{
				AddRef();
				::SendMessage(m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCOMBOSELCHANGE, 0, (LPARAM)this);
			}
		}
	}
	return NOERROR;
}

STDMETHODIMP CTool::get_CBStyle(ComboStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (pList)
		*retval = pList->lpV1.m_nStyle;
	else
		*retval = ddCBSNormal;
	return NOERROR;
}

STDMETHODIMP CTool::put_CBStyle(ComboStyles val)
{
	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;

	switch (val)
	{
    case ddCBSNormal:
	case ddCBSReadOnly:
		pList->lpV1.m_nStyle = val;
		pList->m_bNeedsSort = FALSE;
		break;

	case ddCBSSorted:
	case ddCBSSortedReadOnly:
		pList->lpV1.m_nStyle = val;
		pList->m_bNeedsSort = TRUE;
		break;

	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}

STDMETHODIMP CTool::put_CBList(ComboList * val)
{
	if (NULL == val)
		return E_INVALIDARG;

	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (NULL == pList)
		return E_FAIL;

	HRESULT hResult = ((IComboList*)val)->CopyTo((IComboList**)&pList);
	if (SUCCEEDED(hResult) && m_pBar)
		m_pBar->SetToolComboList(tpV1.m_nToolId, pList);
	return hResult;
}
STDMETHODIMP CTool::get_CBList(ComboList **retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	*retval = (ComboList*)GetComboList();
	if (NULL == *retval)
		return E_FAIL;

	((LPUNKNOWN)(*retval))->AddRef();
	return NOERROR;
}
STDMETHODIMP CTool::put_CBItemData( int Index,  long Data)
{
	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (pList)
	{
		VARIANT vtIndex;
		vtIndex.vt = VT_I4;
		vtIndex.lVal = Index;
		return pList->put_ItemData(vtIndex, Data);
	}
	return E_FAIL;
}
STDMETHODIMP CTool::get_CBItemData( int Index,  long *Data)
{
	if (NULL == Data)
		return E_INVALIDARG;

	if (ddTTCombobox != tpV1.m_ttTools)
		return E_FAIL;

	CBList* pList = GetComboList();
	if (pList)
	{
		VARIANT vtIndex;
		vtIndex.vt = VT_I4;
		vtIndex.lVal = Index;
		return pList->get_ItemData(vtIndex, Data);
	}
	return E_FAIL;
}
STDMETHODIMP CTool::get_Text(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval=SysAllocString(m_bstrText);
	return NOERROR;
}

STDMETHODIMP CTool::put_Text(BSTR val)
{
	try
	{
		SysFreeString(m_bstrText);
		m_bstrText = SysAllocString(val);
	}
	catch (...)
	{
		assert(FALSE);
	}

	try
	{
		if (m_pBar)
			m_pBar->ChangeToolText(tpV1.m_nToolId, val);
	}
	catch (...)
	{
		assert(FALSE);
	}

	if (IsWindow(m_hWndActive))
	{
		HWND hWndEdit;
		if (ddTTCombobox == tpV1.m_ttTools)
		{
			try
			{
				hWndEdit = GetWindow(m_hWndActive, GW_CHILD);
				if (IsWindow(hWndEdit))
				{
					DWORD dwStart, dwEnd;
					LRESULT lResult = SendMessage(hWndEdit, EM_GETSEL, (WPARAM)&dwStart, (LPARAM)&dwEnd);
					MAKE_TCHARPTR_FROMWIDE(szText, val);
					SendMessage(m_hWndActive, WM_SETTEXT, 0, (LPARAM)szText);
					SendMessage(hWndEdit, EM_SETSEL, dwStart, dwEnd);
				}
			}
			catch (...)
			{
				assert(FALSE);
			}
		}
		else
		{
			hWndEdit = m_hWndActive;
			if (IsWindow(hWndEdit))
			{
				DWORD dwStart, dwEnd;
				LRESULT lResult = SendMessage(hWndEdit, EM_GETSEL, (WPARAM)&dwStart, (LPARAM)&dwEnd);
				MAKE_TCHARPTR_FROMWIDE(szText, m_bstrText);
				SendMessage(m_hWndActive, WM_SETTEXT, 0, (LPARAM)szText);
				SendMessage(hWndEdit, EM_SETSEL, dwStart, dwEnd);
			}
		}
	}
	else if (ddTTCombobox == tpV1.m_ttTools && GetComboList())
	{
		if (ddCBSReadOnly == GetComboList()->lpV1.m_nStyle)
		{
			int nIndex = GetComboList()->GetPosOfItem(m_bstrText);
			if (-1 == nIndex)
			{
				m_pBar->m_theErrorObject.SendError(IDERR_COMBOBOXISREADONLY, val);
				return CUSTOM_CTL_SCODE(IDERR_COMBOBOXISREADONLY);
			}
		}
	}
	return NOERROR;
}

STDMETHODIMP CTool::get_ID(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval=tpV1.m_nToolId;
	return NOERROR;
}
STDMETHODIMP CTool::put_ID(long val)
{
	if (val & CBar::eSpecialToolId)
		return E_FAIL;

	tpV1.m_nToolId = val;
	return NOERROR;
}
STDMETHODIMP CTool::get_Name(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = SysAllocString(m_bstrName);
	return NOERROR;
}
STDMETHODIMP CTool::put_Name(BSTR val)
{
	SysFreeString(m_bstrName);
	m_bstrName = NULL;
	if (NULL == val)
		return NOERROR;
	m_bstrName=SysAllocString(val);
	if (NULL == m_bstrName)
		return E_OUTOFMEMORY;
	return NOERROR;
}
STDMETHODIMP CTool::get_HelpContextID(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval=tpV1.m_nHelpContextId;
	return NOERROR;
}
STDMETHODIMP CTool::put_HelpContextID(long val)
{
	tpV1.m_nHelpContextId=val;
	return NOERROR;
}
STDMETHODIMP CTool::get_TooltipText(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval=SysAllocString(m_bstrToolTipText);
	return NOERROR;
}
STDMETHODIMP CTool::put_TooltipText(BSTR val)
{
	SysFreeString(m_bstrToolTipText);
	m_bstrToolTipText=SysAllocString(val);
	return NOERROR;
}

STDMETHODIMP CTool::SetPicture(ImageTypes eType, LPPICTUREDISP pPicture, VARIANT vColor)
{
	try
	{
		if (NULL == m_pBar)
			return E_FAIL;

		if (eType < ddITNormal || eType > ddITDisabled)
			return DISP_E_BADINDEX;

		HRESULT hResult = NOERROR;
		if (NULL == pPicture)
		{
			if (-1 == tpV1.m_nImageIds[eType])
				return CUSTOM_CTL_SCODE(IDERR_IMAGENOTSET);

			hResult = m_pBar->m_pImageMgr->ReleaseImage(&tpV1.m_nImageIds[eType]);
			if (FAILED(hResult))
				return CUSTOM_CTL_SCODE(IDERR_IMAGENOTSET);

			return hResult;
		}
		IPicture* pPict = NULL;
		if (SUCCEEDED(pPicture->QueryInterface(IID_IPicture, (LPVOID*)&pPict)))
		{
			short nType;
			hResult = pPict->get_Type(&nType);
			if (SUCCEEDED(hResult))
			{
				switch (nType)
				{
				case PICTYPE_ICON:
					{
						HICON hIcon;
						hResult = pPict->get_Handle((OLE_HANDLE*)&hIcon);
						if (SUCCEEDED(hResult))
						{
							ICONINFO info;
							memset(&info, 0, sizeof(ICONINFO));
							info.fIcon = TRUE;
							if (GetIconInfo(hIcon, &info))
							{
								pPict->Release();
								hResult = put_Bitmap(eType, (OLE_HANDLE)info.hbmColor);
								BOOL bResult = DeleteBitmap(info.hbmColor);
								assert(bResult);
								if (SUCCEEDED(hResult))
								{
									if (VT_EMPTY == vColor.vt || VT_ERROR == vColor.vt )
									{
										hResult = put_MaskBitmap(eType, (OLE_HANDLE)info.hbmMask);
									}
									else if (VT_I4 == vColor.vt)
									{
										hResult = put_MaskBitmap(eType, (OLE_HANDLE)NULL);
										hResult = m_pBar->m_pImageMgr->put_MaskColor(tpV1.m_nImageIds[eType], vColor.lVal);
									}
								}
								bResult = DeleteBitmap(info.hbmMask);
								assert(bResult);
							}
							else
							{
								HBITMAP hbmNew = CreateBitmapFromPicture(pPict, FALSE);
								pPict->Release();
								if (hbmNew)
								{
									hResult = put_Bitmap(eType, (OLE_HANDLE)hbmNew);
									if (FAILED(hResult))
										return hResult;

									hResult = put_MaskBitmap(eType, (OLE_HANDLE)NULL);
									if (FAILED(hResult))
										return hResult;

									BOOL bResult = DeleteBitmap(hbmNew);
									assert(bResult);
								}
								else
								{
									DestroyIcon(hIcon);
									m_pBar->m_theErrorObject.SendError(IDERR_INVALIDPICTURE, L"SetPicture failed, no picture");
									return CUSTOM_CTL_SCODE(IDERR_INVALIDPICTURE);
								}
							}
//							DestroyIcon(hIcon);  was removed because of bug 1472  this could cause resource leak
						}
						else
						{
							pPict->Release();
							hResult = E_FAIL;
						}
					}
					break;

				default:
					{
						HBITMAP hbmNew = CreateBitmapFromPicture(pPict, FALSE);
						pPict->Release();
						if (hbmNew)
						{
							hResult = put_Bitmap(eType, (OLE_HANDLE)hbmNew);
							if (FAILED(hResult))
								return hResult;
						
							if (VT_I4 == vColor.vt)
							{
								hResult = m_pBar->m_pImageMgr->put_MaskColor(tpV1.m_nImageIds[eType], vColor.lVal);
								if (FAILED(hResult))
									return hResult;
							}
							
							hResult = put_MaskBitmap(eType, (OLE_HANDLE)NULL);
							if (FAILED(hResult))
								return hResult;

							BOOL bResult = DeleteBitmap(hbmNew);
							assert(bResult);
						}
						else 
							hResult = E_FAIL;
					}
					break;
				}
			}
		}
		if (eType == ddITNormal)
			CalcImageSize();

		if (m_pBand)
		{
			if (ddBTPopup == m_pBand->bpV1.m_btBands)
				m_pBand->Refresh();
			else
				m_pBand->ToolNotification(this, TNF_VIEWCHANGED);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

STDMETHODIMP CTool::SetPictureMask(ImageTypes eType,  LPPICTUREDISP pMaskPicture)
{
	try
	{
		if (NULL == m_pBar)
			return E_FAIL;

		if (eType < ddITNormal || eType > ddITDisabled)
			return DISP_E_BADINDEX;
		
		HRESULT hResult = NOERROR;
		if (NULL == pMaskPicture)
		{
			hResult = m_pBar->m_pImageMgr->ReleaseImage(&tpV1.m_nImageIds[eType]);
			return hResult;
		}
		IPicture* pPict;
		if (SUCCEEDED(pMaskPicture->QueryInterface(IID_IPicture,(LPVOID*)&pPict)))
		{
			HBITMAP hbmNew = CreateBitmapFromPicture(pPict, TRUE);
			if (hbmNew)
			{
				hResult = put_MaskBitmap(eType, (OLE_HANDLE)hbmNew);
				if (FAILED(hResult))
					return hResult;
				BOOL bResult = DeleteBitmap(hbmNew);
				assert(bResult);
				pPict->Release();
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

STDMETHODIMP CTool::GetPicture(ImageTypes nIndex, IPictureDisp **ppPicture)
{
	if (NULL == ppPicture)
		return E_INVALIDARG;

	*ppPicture = NULL;
	if (NULL == m_pBar)
		return E_FAIL;

	try
	{
		if (-1 == tpV1.m_nImageIds[nIndex])
			 return E_FAIL;

		if (NULL == m_pBar->m_pImageMgr)
		{
			m_pBar->m_theErrorObject.SendError(IDERR_INVALIDPICTURE, L"GetPicture failed, no picture");
			return CUSTOM_CTL_SCODE(IDERR_INVALIDPICTURE);
		}

		HBITMAP bmp;
		HRESULT hResult = m_pBar->m_pImageMgr->GetImageBitmap(tpV1.m_nImageIds[nIndex], VARIANT_TRUE, (OLE_HANDLE*)&bmp);
		if (FAILED(hResult))
		{
			m_pBar->m_theErrorObject.SendError(IDERR_INVALIDPICTURE, L"GetPicture failed, no picture");
			return CUSTOM_CTL_SCODE(IDERR_INVALIDPICTURE);
		}

		HBITMAP bmpMask;
		hResult = get_MaskBitmap(nIndex, (OLE_HANDLE*)&bmpMask);
		if (FAILED(hResult))
		{
			m_pBar->m_theErrorObject.SendError(IDERR_INVALIDPICTURE, L"GetPicture failed, no picture");
			return CUSTOM_CTL_SCODE(IDERR_INVALIDPICTURE);
		}

		// Create the icon
		ICONINFO icInfo;
		memset(&icInfo, 0, sizeof(ICONINFO));
		icInfo.fIcon = TRUE;
		icInfo.hbmMask = bmpMask;


		HDC hDC = GetDC(NULL);
		if (hDC)
		{
			HDC hDCSrc = CreateCompatibleDC(hDC);
			if (hDCSrc)
			{
				HBITMAP hBitmapOldSrc = SelectBitmap(hDCSrc, bmp);
				COLORREF crColor = ::GetPixel(hDCSrc, 0, 0);
				BITMAP bmInfo;
				GetObject(bmp, sizeof(bmInfo), &bmInfo);
				int y;
				for (int x = 0; x < bmInfo.bmWidth; x++)
					for (y = 0; y < bmInfo.bmHeight; y++)
						if (crColor == GetPixel(hDCSrc, x, y))
							SetPixel(hDCSrc, x, y, RGB(0,0,0));
				SelectBitmap(hDCSrc, hBitmapOldSrc);
				DeleteDC(hDCSrc);
			}
			ReleaseDC(NULL, hDC);
		}

		icInfo.fIcon = TRUE;
		icInfo.hbmColor = bmp;
		BOOL bResult;
		HICON hIcon = CreateIconIndirect(&icInfo);
		if (NULL == hIcon)
		{
			bResult = DeleteBitmap(bmp);
			assert(bResult);
			bResult = DeleteBitmap(bmpMask);
			assert(bResult);
			m_pBar->m_theErrorObject.SendError(IDERR_INVALIDPICTURE, L"GetPicture failed, no picture");
			return CUSTOM_CTL_SCODE(IDERR_INVALIDPICTURE);
		}
		
		PICTDESC pictDesc;
		pictDesc.cbSizeofstruct = sizeof(PICTDESC);
		// Create picture object from icon
		pictDesc.icon.hicon = hIcon;
		pictDesc.picType = PICTYPE_ICON;
		hResult = OleCreatePictureIndirect(&pictDesc, IID_IPictureDisp, TRUE, (LPVOID*)ppPicture);
		if (SUCCEEDED(hResult))
		{
			IPicture* pPicture;
			hResult = (*ppPicture)->QueryInterface(IID_IPicture, (LPVOID*)&pPicture);
			if (SUCCEEDED(hResult))
				pPicture->put_KeepOriginalFormat(VARIANT_TRUE);
			pPicture->Release();
		}
	    bResult = DeleteBitmap(bmp);
		assert(bResult);
		bResult = DeleteBitmap(bmpMask);
		assert(bResult);
		return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

STDMETHODIMP CTool::get_Bitmap(ImageTypes eType, OLE_HANDLE* pHandle)
{
	try
	{
		if (NULL == pHandle)
			return E_INVALIDARG;
		*pHandle = NULL;

		if (NULL == m_pBar)
			return E_FAIL;

		if (eType < ddITNormal || eType > ddITDisabled)
			return DISP_E_BADINDEX;

		HRESULT hResult = m_pBar->m_pImageMgr->GetImageBitmap(tpV1.m_nImageIds[eType], VARIANT_TRUE, pHandle);
		return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// put_Bitmap 
// 

STDMETHODIMP CTool::put_Bitmap(ImageTypes eType, OLE_HANDLE hBitmap) 
{
	try
	{
		if (NULL == m_pBar)
			return E_FAIL;

		if (eType < ddITNormal || eType > ddITDisabled)
			return DISP_E_BADINDEX;

		HRESULT hResult;
		VARIANT_BOOL vbDesignerCreated = m_pBar && m_pBar->m_pDesigner ? VARIANT_TRUE : VARIANT_FALSE;

		CImageMgr* pImageMgr = m_pBar->m_pImageMgr;

		if (NULL == hBitmap)
		{
			if (-1 == tpV1.m_nImageIds[eType])
				return CUSTOM_CTL_SCODE(IDERR_IMAGENOTSET);

			hResult = pImageMgr->ReleaseImage(&tpV1.m_nImageIds[eType]);
			if (FAILED(hResult))
				return CUSTOM_CTL_SCODE(IDERR_IMAGENOTSET);
		}
		else
		{
			BITMAP bmpInfo;
			GetObject((HBITMAP)hBitmap, sizeof(BITMAP), &bmpInfo);

			if (-1 == tpV1.m_nImageIds[eType])
			{
				hResult = pImageMgr->CreateImageEx(&tpV1.m_nImageIds[eType], bmpInfo.bmWidth, bmpInfo.bmHeight, vbDesignerCreated);
				if (FAILED(hResult))
					return hResult;
			}
			else
			{
				long nRefCnt;
				hResult = pImageMgr->RefCntImage(tpV1.m_nImageIds[eType], &nRefCnt);
				if (FAILED(hResult))
				{
					tpV1.m_nImageIds[eType] = -1;
					hResult = pImageMgr->CreateImageEx(&tpV1.m_nImageIds[eType], bmpInfo.bmWidth, bmpInfo.bmHeight, vbDesignerCreated);
					if (FAILED(hResult))
						return hResult;
				}
				else
				{
					SIZE sizeImage;
					hResult = pImageMgr->Size(tpV1.m_nImageIds[eType], &sizeImage.cx, &sizeImage.cy);
					if (FAILED(hResult))
						return hResult;

					if (nRefCnt > 1 || (sizeImage.cx != bmpInfo.bmWidth || sizeImage.cy != bmpInfo.bmHeight))
					{
						hResult = pImageMgr->ReleaseImage(&tpV1.m_nImageIds[eType]);
						if (FAILED(hResult))
							return hResult;

						hResult = pImageMgr->CreateImageEx(&tpV1.m_nImageIds[eType], bmpInfo.bmWidth, bmpInfo.bmHeight, vbDesignerCreated);
						if (FAILED(hResult))
							return hResult;
					}
				}
			}
			hResult = pImageMgr->PutImageBitmap(tpV1.m_nImageIds[eType], vbDesignerCreated, (OLE_HANDLE)hBitmap);
			if (FAILED(hResult))
				return hResult;
		}
		if (ddITNormal == eType)
			CalcImageSize();
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

STDMETHODIMP CTool::get_MaskBitmap(ImageTypes eType, OLE_HANDLE *hBitmap)
{
	try
	{
		if (NULL == hBitmap)
			return E_INVALIDARG;
		*hBitmap = NULL;

		if (NULL == m_pBar)
			return E_FAIL;

		if (eType < ddITNormal || eType > ddITDisabled)
			return DISP_E_BADINDEX;

		HRESULT hResult = m_pBar->m_pImageMgr->GetMaskBitmap(tpV1.m_nImageIds[eType], hBitmap);
		return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}
STDMETHODIMP CTool::put_MaskBitmap(ImageTypes eType, OLE_HANDLE hBitmap)
{
	try
	{
		if (NULL == m_pBar)
			return E_FAIL;

		if (eType < ddITNormal || eType > ddITDisabled)
			return DISP_E_BADINDEX;

		HRESULT hResult;
		if (-1 == tpV1.m_nImageIds[eType])
		{
			if (NULL == hBitmap)
				return E_FAIL;

			//
			// Setting the Mask and the Image Bitmap is not set
			//

			m_pBar->m_theErrorObject.SendError(IDERR_IMAGEBITMAPNOTSET,L"");
			return CUSTOM_CTL_SCODE(IDERR_IMAGEBITMAPNOTSET);
		}

		VARIANT_BOOL vbDesignerCreated = m_pBar && m_pBar->m_pDesigner ? VARIANT_TRUE : VARIANT_FALSE;

		CImageMgr* pImageMgr = m_pBar->m_pImageMgr;

		if (NULL == hBitmap)
		{
			hResult = pImageMgr->PutMaskBitmap(tpV1.m_nImageIds[eType], vbDesignerCreated, (OLE_HANDLE)hBitmap);
			if (FAILED(hResult))
				return hResult;
		}
		else
		{
			BITMAP bmpInfo;
			GetObject((HBITMAP)hBitmap, sizeof(BITMAP), &bmpInfo);

			SIZE sizeImage;
			hResult = pImageMgr->Size(tpV1.m_nImageIds[eType], &sizeImage.cx, &sizeImage.cy);
			if (FAILED(hResult))
				return hResult;

			if (sizeImage.cx != bmpInfo.bmWidth || sizeImage.cy != bmpInfo.bmHeight)
			{
				m_pBar->m_theErrorObject.SendError(IDERR_MASKINVALIDSIZE,L"");
				return CUSTOM_CTL_SCODE(IDERR_MASKINVALIDSIZE);
			}
			hResult = pImageMgr->PutMaskBitmap(tpV1.m_nImageIds[eType], vbDesignerCreated, (OLE_HANDLE)hBitmap);
			if (FAILED(hResult))
				return hResult;
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
STDMETHODIMP CTool::get_Visible(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_vbVisible;
	return NOERROR;
}

STDMETHODIMP CTool::put_Visible(VARIANT_BOOL val)
{
	tpV1.m_vbVisible = val;
	if (m_pBar)
		m_pBar->SetVisibleButton(tpV1.m_nToolId, tpV1.m_vbVisible);
	return NOERROR;
}

STDMETHODIMP CTool::get_Enabled(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_vbEnabled;
	return NOERROR;
}

STDMETHODIMP CTool::put_Enabled(VARIANT_BOOL val)
{
	tpV1.m_vbEnabled = val;
	if (m_pBar)
		m_pBar->EnableButton(tpV1.m_nToolId, tpV1.m_vbEnabled);
	return NOERROR;
}

STDMETHODIMP CTool::get_Checked(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_vbChecked;
	return NOERROR;
}

STDMETHODIMP CTool::put_Checked(VARIANT_BOOL val)
{
	tpV1.m_vbChecked = val;
	if (m_pBar)
		m_pBar->CheckButton(tpV1.m_nToolId, tpV1.m_vbChecked);
	return NOERROR;
}

STDMETHODIMP CTool::get_Caption(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = SysAllocString(m_bstrCaption);
	return NOERROR;
}
inline TCHAR* CleanCaptionHelper(BSTR bstrCaption,BOOL bPopup)
{
	WCHAR wCaption[256];

	if (bstrCaption)
	{
		if (wcslen(bstrCaption) < 255)
			wcscpy(wCaption, bstrCaption);
		else
		{
			wcsncpy(wCaption, bstrCaption, 254);
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
			if (bPopup && (L'\n' == *wChar || L'\r' == *wChar))
				*wChar = L' ';
			wChar++;
		}
	}
	else
		*wCaption = 0;

	TCHAR *retval;
	long len= (lstrlenW(wCaption) + 1) * sizeof(WCHAR);
	retval=(TCHAR *)malloc(len);
	if (retval==NULL)
		return NULL;

    WideCharToMultiByte(CP_ACP, 0, wCaption, -1, (LPSTR)retval, len, NULL, NULL);

    return retval;
}

STDMETHODIMP CTool::put_Caption(BSTR val)
{
	SysFreeString(m_bstrCaption);
	m_bstrCaption = NULL;
	m_bstrCaption = SysAllocString(val);
	
	if (m_szCleanCaptionNormal)
		free(m_szCleanCaptionNormal);
	if (m_szCleanCaptionPopup)
		free(m_szCleanCaptionPopup);

	m_szCleanCaptionNormal = CleanCaptionHelper(m_bstrCaption,FALSE);
	m_szCleanCaptionPopup = CleanCaptionHelper(m_bstrCaption,TRUE);
	if (m_pBand && ddBTMenuBar == m_pBand->bpV1.m_btBands)
	{
		if (m_pBar)
		{
			if (CBar::eMDIForm == m_pBar->m_eAppType || CBar::eSDIForm == m_pBar->m_eAppType)
				m_pBar->BuildAccelators(m_pBar->GetMenuBand());
			else if (m_pBar->m_pControlSite)
				m_pBar->m_pControlSite->OnControlInfoChanged();
		}
	}

	return NOERROR;
}
STDMETHODIMP CTool::get_Style(ToolStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_tsStyle;
	return NOERROR;
}
STDMETHODIMP CTool::put_Style(ToolStyles val)
{
    if (val < ddSStandard || val > ddSCheckedIconText)
		return E_INVALIDARG;
	tpV1.m_tsStyle = val;
	return NOERROR;
}
STDMETHODIMP CTool::get_CaptionPosition(CaptionPositionTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = tpV1.m_cpTools;
	return NOERROR;
}
STDMETHODIMP CTool::put_CaptionPosition(CaptionPositionTypes val)
{
	if (val < ddCPStandard || val > ddCPCenter)
		return E_INVALIDARG;
	tpV1.m_cpTools = val;
	return NOERROR;
}
STDMETHODIMP CTool::get_ShortCuts(VARIANT *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	try
	{
		HRESULT hResult = VariantClear(retval);
		if (FAILED(hResult))
			return hResult;

		int nCount = m_scTool.GetCount();

		SAFEARRAYBOUND rgsabound[1];    
		rgsabound[0].lLbound = 0;
		rgsabound[0].cElements = nCount;
		
		retval->parray = SafeArrayCreate(VT_DISPATCH, 1, rgsabound);
		if (NULL == retval->parray)
			return E_FAIL;

		retval->vt = VT_ARRAY|VT_DISPATCH;

		LPDISPATCH ppDispatch;
		hResult = SafeArrayAccessData(retval->parray, (void**)&ppDispatch);
		if (FAILED(hResult))
		{
			SafeArrayDestroy(retval->parray);
			return hResult;
		}

		ShortCutStore* pShortCut;
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pShortCut = m_scTool.Get(nIndex);
			assert(pShortCut);
			if (NULL == pShortCut)
				continue;

			hResult = pShortCut->Clone((IShortCut**)&ppDispatch[nIndex]);
			assert(SUCCEEDED(hResult));
		}

		SafeArrayUnaccessData(retval->parray);
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}

}
STDMETHODIMP CTool::put_ShortCuts(VARIANT val)
{
	try
	{
		if (0 == (val.vt & (VT_DISPATCH | VT_ARRAY)))
		{
			m_pBar->m_theErrorObject.SendError(IDERR_MUSTBEARRAYOFSTRINGS, NULL);
			return CUSTOM_CTL_SCODE(IDERR_MUSTBEARRAYOFSTRINGS);
		}

		long nLBound;
		HRESULT hResult = SafeArrayGetLBound(val.parray, 1,  &nLBound);
		if (FAILED(hResult))
			return hResult;
		
		long nUBound;
		hResult = SafeArrayGetUBound(val.parray, 1,  &nUBound);
		if (FAILED(hResult))
			return hResult;

		LPDISPATCH* ppDispatch;
		hResult = SafeArrayAccessData(val.parray, (void**)&ppDispatch);
		if (FAILED(hResult))
			return hResult;

		m_scTool.RemoveAll();

		int nCount = nUBound - nLBound + 1;
		for (int nIndex = nLBound; nIndex < nCount; nIndex++)
		{
			if (NULL == ppDispatch[nIndex])
				continue;

			hResult = m_scTool.Add((ShortCutStore*)ppDispatch[nIndex]);
			if (FAILED(hResult))
				continue;
		}
		SafeArrayUnaccessData(val.parray);
		if (m_pBar)
			m_pBar->SetToolShortCut(tpV1.m_nToolId, &m_scTool);
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}
STDMETHODIMP CTool::get_Description(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval=SysAllocString(m_bstrDescription);
	return NOERROR;
}
STDMETHODIMP CTool::put_Description(BSTR val)
{
	SysFreeString(m_bstrDescription);
	m_bstrDescription=SysAllocString(val);
	return NOERROR;
}
STDMETHODIMP CTool::get_ControlType(ToolTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_ttTools;
	return NOERROR;
}
STDMETHODIMP CTool::put_ControlType(ToolTypes val)
{
	switch (val)
	{
    case ddTTButton:
	case ddTTButtonDropDown:
	case ddTTCombobox:
	case ddTTEdit:
	case ddTTLabel:
	case ddTTSeparator:
	case ddTTControl:
	case ddTTForm:
	case ddTTWindowList:
		tpV1.m_ttTools = val;
		break;
	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}
STDMETHODIMP CTool::get_Custom(LPDISPATCH *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = NULL;
	if (!m_pBar->AmbientUserMode() || (m_pBar && VARIANT_TRUE == m_pBar->m_vbLibrary))
		return E_FAIL;

	*retval=m_pDispCustom;
	if (m_pDispCustom)
		m_pDispCustom->AddRef();
	return NOERROR;
}
STDMETHODIMP CTool::putref_Custom(IDispatch * *val)
{
	if (NULL == val)
		return E_INVALIDARG;

	return put_Custom(*val);
}
STDMETHODIMP CTool::put_Custom(LPDISPATCH val)
{
	try
	{
		if (m_pDispCustom == val)
			return NOERROR;

		if (m_pDispCustom)
		{
			//
			// If the PrevParent was set reset the parent of the control or form
			//

			if (IsWindow(m_hWndActive))
			{
				ShowWindow(m_hWndActive, SW_HIDE);
				if (IsWindow(m_hWndPrevParent))
				{
					::SetParent(m_hWndActive, m_hWndPrevParent);
					m_hWndPrevParent = NULL;
					m_hWndParent = NULL;

					if (0 != m_dwStyleOld)
						SetWindowLong(m_hWndActive, GWL_STYLE, m_dwStyleOld);

					if (0 != m_dwExStyleOld)
						SetWindowLong(m_hWndActive, GWL_EXSTYLE, m_dwExStyleOld);
				}
			}

			//
			// Release this control or form
			//

			m_pDispCustom->Release();
			m_pDispCustom = NULL;
			m_nCustomWidth = 0;
			m_dwStyleOld = 0;
			m_dwExStyleOld = 0;
		}
		IDispatch* pCustInterface;
		if (val) 
		{
			//
			// First check if interface supports ICustomTool
			//

			if (SUCCEEDED(val->QueryInterface(DIID_ICustomTool, (LPVOID*)&pCustInterface)))
			{
				//
				// ICustomTool Interface here
				//

				CCustomTool custTool(pCustInterface);
				custTool.SetHost((LPDISPATCH)(ITool*)this);
				if (SUCCEEDED(custTool.hr))
				{
					int w=16,h=16;
					custTool.GetSize(&w,&h);
					custTool.GetFlags((long*)&m_dwCustomFlag);
					if (FAILED(custTool.hr))
						m_dwCustomFlag = 0;
				}
				else
					m_dwCustomFlag = 0;

				pCustInterface->Release();
				if (0 == m_dwCustomFlag)
				{
					m_pDispCustom = NULL;
					return E_FAIL;
				}
				m_pDispCustom = val;
				if (m_pDispCustom)
					m_pDispCustom->AddRef();

				if (m_pBar)
					m_pBar->SetCustomInterface(tpV1.m_nToolId,val,m_dwCustomFlag);
			}
			else
			{
				//
				// Custom Controls and VB Forms
				//

				if (ddTTForm != tpV1.m_ttTools && ddTTControl != tpV1.m_ttTools)
				{
					MAKE_TCHARPTR_FROMWIDE(szBand, m_pBand->m_bstrName);
					MAKE_TCHARPTR_FROMWIDE(szTool, m_bstrName);
					DDString strError;
					strError.Format(_T("Band: %s, Tool: %s"), szBand, szTool);
					m_pBar->m_theErrorObject.SendError(IDERR_WRONGTOOLTYPE, strError.AllocSysString());
					return CUSTOM_CTL_SCODE(IDERR_WRONGTOOLTYPE);
				}

				//
				// Get the Window handle of the Control or Form
				//

				VARIANT vProperty;
				vProperty.vt = VT_EMPTY;
				HRESULT hResult = PropertyGet(val, L"hWnd", vProperty);
				if (FAILED(hResult)) 
				{
					MAKE_TCHARPTR_FROMWIDE(szBand, m_pBand->m_bstrName);
					MAKE_TCHARPTR_FROMWIDE(szTool, m_bstrName);
					DDString strError;
					strError.Format(_T("Band: %s, Tool: %s"), szBand, szTool);
					m_pBar->m_theErrorObject.SendError(IDERR_FAILEDTOGETWINDOWHANDLE, strError.AllocSysString());
					return CUSTOM_CTL_SCODE(IDERR_FAILEDTOGETWINDOWHANDLE);
				}

				m_hWndActive = (HWND)vProperty.lVal;
				if (!IsWindow(m_hWndActive))
				{
					MAKE_TCHARPTR_FROMWIDE(szBand, m_pBand->m_bstrName);
					MAKE_TCHARPTR_FROMWIDE(szTool, m_bstrName);
					DDString strError;
					strError.Format(_T("Band: %s, Tool: %s"), szBand, szTool);
					m_pBar->m_theErrorObject.SendError(IDERR_FAILEDTOGETWINDOWHANDLE, strError.AllocSysString());
					return CUSTOM_CTL_SCODE(IDERR_FAILEDTOGETWINDOWHANDLE);
				}

				DWORD dwStyle = GetWindowLong(m_hWndActive, GWL_STYLE);
				m_dwStyleOld = dwStyle;

				DWORD dwExStyle = GetWindowLong(m_hWndActive, GWL_EXSTYLE);
				m_dwExStyleOld = dwExStyle;

				if (ddTTForm == tpV1.m_ttTools)
				{
					CRect rcPrevClient;
					GetClientRect(m_hWndActive, &rcPrevClient);
					dwStyle &= ~WS_OVERLAPPEDWINDOW;
					SetWindowLong(m_hWndActive, GWL_STYLE, dwStyle);

					dwExStyle &= ~WS_EX_APPWINDOW;
					SetWindowLong(m_hWndActive, GWL_EXSTYLE, dwExStyle);

					SetWindowPos(m_hWndActive, NULL, 0, 0, rcPrevClient.Width(), rcPrevClient.Height(), SWP_NOZORDER|SWP_NOMOVE|SWP_NOACTIVATE);

					SendMessage(m_hWndActive, WM_SIZE, 0, 0);
				}
				if (ddTTControl == tpV1.m_ttTools)
				{
					VARIANT vProperty;
					vProperty.vt = VT_R4;
					vProperty.fltVal = (float)-100000.0;
					IOleObject* pObject;
					hResult = val->QueryInterface(IID_IOleObject, (void**)&pObject);
					if (SUCCEEDED(hResult))
					{
						// Regular Control
						IOleClientSite* pClientSite;
						hResult = pObject->GetClientSite(&pClientSite);
						pObject->Release();
						if (SUCCEEDED(hResult))
						{
							IOleControlSite* pControlSite;
							hResult = pClientSite->QueryInterface(IID_IOleControlSite, (void**)&pControlSite);
							pClientSite->Release();
							if (SUCCEEDED(hResult))
							{
								IDispatch* pDispatchEx;
								hResult = pControlSite->GetExtendedControl(&pDispatchEx);
								pControlSite->Release();
								if (SUCCEEDED(hResult))
								{
									hResult = PropertyPut(pDispatchEx, L"Left", vProperty);
									pDispatchEx->Release();
								}
							}
						}
					}
					else
						// VB Control
						hResult = PropertyPut(val, L"Left", vProperty);
				}

				if (m_pBar && m_pBar->m_pRelocate)
					m_pBar->m_pRelocate->FindControl(val);

				//
				// Add the child style to the control or form
				//

				SetWindowLong(m_hWndActive, GWL_STYLE, dwStyle | WS_CHILD);

				//
				// Set the parent of the control or form
				//

				if (m_pBand)
				{
					if (m_pBand->m_pDock)
						SetParent(m_pBand->m_pDock->hWnd());
					else if (m_pBand->m_pFloat)
						SetParent(m_pBand->m_pFloat->hWnd());
				}
				else
					m_hWndPrevParent = ::GetParent(m_hWndActive);

				//
				// AddRef to the control or form so we can keep it around
				//

				m_dwCustomFlag = 0;
				m_pDispCustom = val;
				m_pDispCustom->AddRef();

				//
				// Add it to the Forms and Controls collection so we can shutdown properly
				//

				if (m_pBand && m_pBand->m_pBar)
					m_pBand->m_pBar->m_theFormsAndControls.Add(this);
				CRect rc;
				GetWindowRect(m_hWndActive, &rc);
				m_nCustomWidth = rc.Width();
			}
		}
		else if (m_pBand && m_pBand->m_pBar)
		{
			//
			// We no longer need to hold on to this control or form so remove it from the collection
			//

			m_pBand->m_pBar->m_theFormsAndControls.Remove(this);
			m_hWndPrevParent = m_hWndParent = m_hWndActive = NULL;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

STDMETHODIMP CTool::get_SubBand(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval=SysAllocString(m_bstrSubBand);
	return NOERROR;
}
STDMETHODIMP CTool::put_SubBand(BSTR val)
{
	SysFreeString(m_bstrSubBand);
	m_bstrSubBand = NULL;
	m_bstrSubBand=SysAllocString(val);
	return NOERROR;
}
STDMETHODIMP CTool::get_Width(OLE_XSIZE_PIXELS *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (m_pBar && !m_pBar->AmbientUserMode() && -1 == tpV1.m_nWidth)
	{
		*retval = tpV1.m_nWidth;
		return NOERROR;
	}

	switch (tpV1.m_ttTools)
	{
	case ddTTControl:
	case ddTTForm:
		{
			if (m_pDispCustom)
			{
				VARIANT vProperty;
				vProperty.vt = VT_R4;
				HRESULT hResult = PropertyGet(m_pDispCustom, L"Width", vProperty);
				if (SUCCEEDED(hResult))
				{
					hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
					if (SUCCEEDED(hResult))
					{
						*retval = vProperty.lVal;
						return NOERROR;
					}
					else
						*retval = tpV1.m_nWidth;
				}
				else
					*retval = tpV1.m_nWidth;
			}
			else
				*retval = tpV1.m_nWidth;
		}
		break;

	default:
		if (-1 == tpV1.m_nWidth)
			*retval = m_rcTemp.Width();
		else
			*retval = tpV1.m_nWidth;
		break;
	}
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CTool::put_Width(OLE_XSIZE_PIXELS val)
{
	if (val < -1)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_INVALIDPROPERTYVALUE, NULL);
		return CUSTOM_CTL_SCODE(IDERR_INVALIDPROPERTYVALUE);
	}
	SIZE size;
	size.cx = val;
	if (val > 0)
		TwipsToPixel(&size, &size);
	switch (tpV1.m_ttTools)
	{
	case ddTTControl:
	case ddTTForm:
		{
			if (m_pDispCustom && m_pBar && m_pBar->AmbientUserMode())
			{
				VARIANT vProperty;
				vProperty.vt = VT_R4;
				vProperty.fltVal = (float)val;
				HRESULT hResult = PropertyPut(m_pDispCustom, L"Width", vProperty);
				tpV1.m_nWidth = size.cx;
			}
			else
				tpV1.m_nWidth = size.cx;
		}
		break;

	default:
		tpV1.m_nWidth = size.cx;
		break;
	}
	return NOERROR;
}
STDMETHODIMP CTool::get_Height(OLE_YSIZE_PIXELS *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (m_pBar && !m_pBar->AmbientUserMode() && -1 == tpV1.m_nHeight)
	{
		*retval = tpV1.m_nHeight;
		return NOERROR;
	}

	switch (tpV1.m_ttTools)
	{
	case ddTTControl:
	case ddTTForm:
		{
			if (m_pDispCustom)
			{
				VARIANT vProperty;
				vProperty.vt = VT_R4;
				HRESULT hResult = PropertyGet(m_pDispCustom, L"Height", vProperty);
				if (SUCCEEDED(hResult))
				{
					hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
					if (SUCCEEDED(hResult))
					{
						*retval = vProperty.lVal;
						return NOERROR;
					}
					else
						*retval = tpV1.m_nHeight;
				}
				else
					*retval = tpV1.m_nHeight;
			}
			else
				*retval = tpV1.m_nHeight;
		}
		break;

	default:
		if (-1 == tpV1.m_nHeight)
			*retval = m_rcTemp.Height() - 2 * CBand::eHToolPadding;
		else
			*retval = tpV1.m_nHeight;
		break;
	}
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CTool::put_Height(OLE_YSIZE_PIXELS val)
{
	if (val < -1)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_INVALIDPROPERTYVALUE, NULL);
		return CUSTOM_CTL_SCODE(IDERR_INVALIDPROPERTYVALUE);
	}
	SIZE size;
	size.cx = val;
	if (val > 0)
		TwipsToPixel(&size, &size);
	switch (tpV1.m_ttTools)
	{
	case ddTTControl:
	case ddTTForm:
		{
			if (m_pDispCustom && m_pBar && m_pBar->AmbientUserMode())
			{
				VARIANT vProperty;
				vProperty.vt = VT_R4;
				vProperty.fltVal = (float)val;
				HRESULT hResult = PropertyPut(m_pDispCustom, L"Height", vProperty);
				tpV1.m_nHeight=size.cx;
			}
			else
				tpV1.m_nHeight=size.cx;
		}
		break;

	default:
		tpV1.m_nHeight = size.cx;
		break;
	}
	return NOERROR;
}
STDMETHODIMP CTool::get_Alignment(ToolAlignmentTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = tpV1.m_taTools;
	return NOERROR;
}
STDMETHODIMP CTool::put_Alignment(ToolAlignmentTypes val)
{
	if (val < ddALeftTop || val > ddARightBottom)
		return E_INVALIDARG;
	tpV1.m_taTools = val;
	return NOERROR;
}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
void CTool::SetBar(CBar* pBar)
{
	m_pBar = pBar;
	m_scTool.Initialize(pBar, this);
}

void CTool::SetBand(CBand* pBand)
{
	m_pBand = pBand;
	if (m_pBand && m_pBand->m_pBar)
		SetBar(m_pBand->m_pBar);
}

static inline COLORREF ROP4(COLORREF crFore, COLORREF crBack)
{
	return ((((crBack) << 8) & 0xFF000000) | (crFore));
}

BOOL CTool::Customization()
{
	return (m_pBar->m_bCustomization && (NULL == m_pBand || ddCBSystem != m_pBand->bpV1.m_nCreatedBy));
}

//
// Draw
//

void CTool::Draw(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical, BOOL bVerticalPaint)
{
	try
	{
		ToolTypes ttTools = tpV1.m_ttTools;
		if (bVertical && (ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
			tpV1.m_ttTools = ddTTButton;

		switch (tpV1.m_ttTools)
		{
		case ddTTButton:
			try
			{
				DrawButton(hDC, rcPaint, btType, bVerticalPaint);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTButtonDropDown:
			try
			{
				DrawButtonDropDown(hDC, rcPaint, btType, bVerticalPaint);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTCombobox:
			try
			{
				DrawCombobox(hDC, rcPaint, btType, bVerticalPaint);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTEdit:
			try
			{
				DrawEdit(hDC, rcPaint, btType, bVerticalPaint);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;
		
		case ddTTLabel:
			try
			{
				DrawStatic(hDC, rcPaint, btType, bVerticalPaint);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTSeparator:
			try
			{
				DrawSeparator(hDC, rcPaint, btType, bVertical);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTMDIButtons:
			try
			{
				DrawMDIButtons(hDC, rcPaint, btType, bVerticalPaint);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTChildSysMenu:
			try
			{
				DrawChildSysMenu(hDC, rcPaint, btType, bVerticalPaint);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTMenuExpandTool:
			try
			{
				DrawMenuExpand(hDC, rcPaint);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTMoreTools:
			try
			{
				DrawMoreTools(hDC, rcPaint, btType, bVertical);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case ddTTAllTools:
			try
			{
				DrawAllTools(hDC, rcPaint, btType, bVertical);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;
		}

		if (ttTools != tpV1.m_ttTools)
			tpV1.m_ttTools = ttTools;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}


//
// CleanCaption
//

void CTool::CleanCaption(TCHAR* szCaption, BOOL bPopup)
{
	if (!bPopup)
	{
		if (m_szCleanCaptionNormal)
			lstrcpy(szCaption,m_szCleanCaptionNormal);
		else
			*szCaption=0;
	}
	else
	{
		if (m_szCleanCaptionPopup)
			lstrcpy(szCaption,m_szCleanCaptionPopup);
		else
			*szCaption=0;
	}
}

//
// CalcImageTextPos
//

void CTool::CalcImageTextPos(const CRect& rcPaint, 
							 BandTypes	  btType, 
							 BOOL		  bVertical, 
							 SIZE		  sizeTotal, 
							 SIZE		  sizeIcon, 
							 SIZE		  sizeText, 
							 BOOL		  bIcon, 
							 BOOL		  bText,
							 CRect&		  rcText,
							 CRect&		  rcImage)
{
	if (bVertical)
	{
		//swap
		int nTemp = sizeTotal.cy;
		sizeTotal.cy = sizeTotal.cx;
		sizeTotal.cx = nTemp;

		nTemp = sizeText.cy;
		sizeText.cy = sizeText.cx;
		sizeText.cx = nTemp;
	}

	int nHPadding = (NULL == m_pBand) ? CBand::eHToolPadding : m_pBand->bpV1.m_nToolsHPadding;
	int nVPadding = (NULL == m_pBand) ? CBand::eVToolPadding : m_pBand->bpV1.m_nToolsVPadding;

	CaptionPositionTypes cpText = tpV1.m_cpTools;
	if (ddCPStandard == cpText)
		cpText = ddCPRight;

	int nFontHeight = sizeText.cy;
	POINT ptIcon, ptText;
	if (bIcon & bText)
	{
		switch(cpText)
		{
		case ddCPLeft:
			ptIcon.x = nHPadding + sizeText.cx;
			ptIcon.y = (sizeTotal.cy - sizeIcon.cy) / 2;
			ptText.x = nHPadding;
			ptText.y = (sizeTotal.cy - nFontHeight) / 2;
			break;
	
		case ddCPStandard:
		case ddCPRight:
			ptIcon.x = nHPadding;
			ptIcon.y = (sizeTotal.cy - sizeIcon.cy) / 2;
			ptText.x = nHPadding + sizeIcon.cx;
			ptText.y = (sizeTotal.cy - nFontHeight) / 2;
			break;
		
		case ddCPAbove:
			ptText.x = (sizeTotal.cx - sizeText.cx) / 2;
			ptText.y = nVPadding;

			switch (abs(tpV1.m_taTools % 3))
			{
			case eHTALeft:
				ptIcon.x = nHPadding;
				break;

			case eHTACenter:
				ptIcon.x = (sizeTotal.cx - sizeIcon.cx) / 2;
				break;

			case eHTARight:
				ptIcon.x = sizeTotal.cx - sizeIcon.cx - 2 * nHPadding;
				break;
			}
			ptIcon.y = nVPadding + sizeText.cy;
			break;
		
		case ddCPCenter:
			ptText.x = (sizeTotal.cx - sizeText.cx) / 2;
			ptText.y = (sizeTotal.cy - sizeText.cy) / 2;
			ptIcon.x = (sizeTotal.cx - sizeIcon.cx) / 2;
			ptIcon.y = (sizeTotal.cy - sizeIcon.cy) / 2;
			break;

		case ddCPBelow:
			ptText.x = (sizeTotal.cx - sizeText.cx) / 2;
			ptText.y = sizeIcon.cy;

			switch (abs(tpV1.m_taTools % 3))
			{
			case eHTALeft: 
				ptIcon.x = nHPadding;
				break;

			case eHTACenter: 
				ptIcon.x = (sizeTotal.cx - sizeIcon.cx) / 2;
				break;

			case eHTARight: 
				ptIcon.x = sizeTotal.cx - sizeIcon.cx - 2 * nHPadding;
				break;
			}
			ptIcon.y = 0;
			break;
		}
	}
	else if (bText)
	{
		ptText.x = (sizeTotal.cx-sizeText.cx) / 2;
		ptText.y = (sizeTotal.cy-sizeText.cy) / 2;
	}
	else if (bIcon)
	{
		ptIcon.x = (sizeTotal.cx - sizeIcon.cx) / 2;
		ptIcon.y = (sizeTotal.cy - sizeIcon.cy) / 2;
	}

	//
	// Now we have sizeTotal, sizeIcon , sizeText , ptIcon and ptText
	// Let's calc the sizeTotal rectangles offset inside the rcPaint
	//

	int nXOffset = 0;
	switch(abs(tpV1.m_taTools % 3))
	{
	case eHTALeft:
		nXOffset = rcPaint.left;
		break;

	case eHTACenter:
		nXOffset = rcPaint.left + (rcPaint.Width() - sizeTotal.cx) / 2;
		break;

	case eHTARight:
		nXOffset = rcPaint.right - sizeTotal.cx;
		break;

	default:
		assert(FALSE);
		break;
	}

	int nYOffset = 0;
	switch (abs(tpV1.m_taTools / 3))
	{
	case eVTAAbove:
		nYOffset = rcPaint.top;
		break;

	case eVTACenter:
		nYOffset = rcPaint.top + (rcPaint.Height() - sizeTotal.cy) / 2;
		break;

	case eVTABelow:
		nYOffset = rcPaint.bottom - sizeTotal.cy;
		break;

	default:
		assert(FALSE);
		break;
	}

	POINT ptImageOffset;
	ptImageOffset.x = ptImageOffset.y = 0;
	
	rcImage.SetEmpty();
	if (bIcon)
	{
		if (0 != m_dwCustomFlag && m_dwCustomFlag >= 1)
		{
			rcImage.left = nXOffset + ptIcon.x + ptImageOffset.x;
			rcImage.top = nYOffset + ptIcon.y + ptImageOffset.y;
			rcImage.right = rcImage.left + sizeIcon.cx;
			rcImage.bottom = rcImage.top + sizeIcon.cy;
		}
		else if (-1 != tpV1.m_nImageIds[0])
		{
			rcImage.left = nXOffset + ptIcon.x + nHPadding + ptImageOffset.x;
			rcImage.top = nYOffset + ptIcon.y + nVPadding + ptImageOffset.y;
			rcImage.right = rcImage.left + sizeIcon.cx - 2 * nHPadding;
			rcImage.bottom = rcImage.top + sizeIcon.cy - 2 * nVPadding;
		}

		if (!rcImage.IsEmpty())
		{
			BOOL bResult = IntersectRect(&rcImage, &rcImage, &rcPaint);
//			assert(bResult);
		}
	}

	if (bText)
	{
		rcText.left = nXOffset + ptText.x;
		rcText.top = nYOffset + ptText.y;
		
		rcText.right = rcText.left + sizeText.cx;
		rcText.bottom = rcText.top + sizeText.cy;

		if (!rcText.IsEmpty())
		{
			BOOL bResult = IntersectRect(&rcText, &rcText, &rcPaint);
//			assert(bResult);
		}

		if (m_pBand && m_pBand->m_pParentBand && ddCBSSlidingTabs != m_pBand->m_pParentBand->bpV1.m_cbsChildStyle)
		{
			ptText.x += ptImageOffset.x;
			ptText.y += ptImageOffset.y;
		}
	}
	else
		rcText.SetEmpty();
}

//
// PopupBand
//

void CTool::PopupBand(HDC		   hDC, 
					  LPCTSTR      szCaption,
					  BandTypes    btType, 
					  const CRect& rcPaint,
					  const SIZE&  sizeTotal, 
					  const SIZE&  sizeChecked, 
					  const SIZE&  sizeIcon, 
					  const SIZE&  sizeText, 
					  BOOL		   bVertical, 
					  BOOL		   bChecked, 
					  BOOL		   bIcon, 
					  BOOL		   bText)
{
	try
	{
		int nIconHeight = max (sizeText.cy, sizeIcon.cy) + 6;
		if (nIconHeight > rcPaint.Height())
			nIconHeight = rcPaint.Height();

		BOOL bXPLook = VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook;

		if (-1L == tpV1.m_nImageIds[ddITNormal] && 0 == m_dwCustomFlag)
			bIcon = FALSE;

		BOOL bResult;
		if (m_bPressed && bXPLook)
		{
			HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
			if (aPen)
			{
				HPEN penOld = SelectPen(hDC, aPen);

				HBRUSH aBrush;
				if (VARIANT_FALSE == tpV1.m_vbEnabled)
					aBrush = CreateSolidBrush(m_pBar->m_crXPMenuBackground);
				else
					aBrush = CreateSolidBrush(m_pBar->m_crXPSelectedColor);

				if (aBrush)
				{
					HBRUSH brushOld = SelectBrush(hDC, aBrush);
				
					Rectangle(hDC, rcPaint.left, rcPaint.top, rcPaint.right, rcPaint.bottom);
					
					SelectBrush(hDC, brushOld);
					bResult = bResult = DeleteBrush(aBrush);
					assert(bResult);
				}
				SelectPen(hDC, penOld);
				bResult = DeletePen(aPen);
				assert(bResult);
			}
		}

		CRect rcImage;
		CRect rcText;
		if (bChecked)
		{
			if (VARIANT_TRUE == tpV1.m_vbChecked)
			{
				//
				// Draw Check
				//

				CRect rcImageBackground = rcPaint;
				rcImageBackground.right = rcImageBackground.left + nIconHeight;
				rcImageBackground.bottom = rcImageBackground.top + nIconHeight;
				if (bXPLook)
				{
					HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
					if (aPen)
					{
						HPEN penOld = SelectPen(hDC, aPen);

						HBRUSH aBrush = CreateSolidBrush(m_pBar->m_crXPCheckedColor);
						if (aBrush)
						{
							HBRUSH brushOld = SelectBrush(hDC, aBrush);
							
							rcImageBackground.Inflate(-1, -2);
							Rectangle(hDC, rcImageBackground.left, rcImageBackground.top, rcImageBackground.right, rcImageBackground.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(aBrush);
							assert(bResult);
						}
						
						SelectPen(hDC, penOld);
						bResult = DeletePen(aPen);
						assert(bResult);
					}
				}
				else
				{
					if (!m_bPressed)
						DrawDitherRect(hDC, rcImageBackground, m_pBar ? m_pBar->m_crHighLight : GetSysColor(COLOR_BTNHIGHLIGHT));
				}

				rcImage.left = rcPaint.left + ((nIconHeight - eMenuCheckWidth) / 2);
				rcImage.top = rcPaint.top + ((nIconHeight - eMenuCheckHeight) / 2);
				rcImage.right = rcImage.left + sizeIcon.cx;
				rcImage.bottom = rcImage.top + sizeIcon.cy;
				DrawChecked(hDC, rcImage);

				if (!bXPLook)
				{
					rcImage = rcPaint;
					rcImage.right = rcImage.left + nIconHeight;
					rcImage.bottom = rcImage.top + nIconHeight;
					DrawEdge(hDC, &rcImage, BDR_SUNKENOUTER, BF_RECT);
				}
			}

			if (bIcon)
			{
				// Draw Image
				rcImage.left = rcPaint.left + nIconHeight + ((nIconHeight - sizeIcon.cx) / 2);
				rcImage.top = rcPaint.top + ((nIconHeight - sizeIcon.cy) / 2);
				rcImage.right = rcImage.left + sizeIcon.cx;
				rcImage.bottom = rcImage.top + sizeIcon.cy;
				BOOL bPrevPressed = m_bPressed;
				m_bPressed = FALSE;
				DrawImage(hDC, rcImage);
				m_bPressed = bPrevPressed;
				
				if (m_bPressed && !bXPLook)
				{
					rcImage.Inflate(1, 1);
					DrawEdge(hDC, &rcImage, BDR_RAISEDINNER, BF_RECT);
				}
			}
		}
		else if ((VARIANT_TRUE == tpV1.m_vbChecked && !bChecked) || bIcon)
		{
			
			if (VARIANT_TRUE == tpV1.m_vbChecked && !bChecked)
			{
				// Drawing the background
				CRect rcImageBackground = rcPaint;
				rcImageBackground.right = rcImageBackground.left + nIconHeight;
				rcImageBackground.bottom = rcImageBackground.top + nIconHeight;
				if (bXPLook)
				{
					HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
					if (aPen)
					{
						HPEN penOld = SelectPen(hDC, aPen);

						HBRUSH aBrush = CreateSolidBrush(m_pBar->m_crXPCheckedColor);
						if (aBrush)
						{
							HBRUSH brushOld = SelectBrush(hDC, aBrush);
							rcImageBackground.Inflate(-1, -2);
							Rectangle(hDC, rcImageBackground.left, rcImageBackground.top, rcImageBackground.right, rcImageBackground.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(aBrush);
							assert(bResult);
						}
						
						SelectPen(hDC, penOld);
						bResult = DeletePen(aPen);
						assert(bResult);
					}
				}
				else if (!m_bPressed)
					DrawDitherRect(hDC, rcImageBackground, m_pBar ? m_pBar->m_crHighLight : GetSysColor(COLOR_BTNHIGHLIGHT));
			}
			
			if (VARIANT_TRUE == tpV1.m_vbChecked && !bChecked && -1L == tpV1.m_nImageIds[ddITPressed] && -1L == tpV1.m_nImageIds[ddITNormal] && 0 == m_dwCustomFlag)
			{
				// if the tool does not have an image draw the default check mark
				rcImage.left = rcPaint.left + ((nIconHeight - eMenuCheckWidth) / 2);
				rcImage.top = rcPaint.top + ((nIconHeight - eMenuCheckHeight) / 2);
				rcImage.right = rcImage.left + sizeIcon.cx;
				rcImage.bottom = rcImage.top + sizeIcon.cy;
				DrawChecked(hDC, rcImage);
			}
			else
			{
				// if the tool does have a image then draw it
				rcImage.left = rcPaint.left + ((nIconHeight - sizeIcon.cx) / 2);
				rcImage.top = rcPaint.top + ((nIconHeight - sizeIcon.cy) / 2);
				rcImage.right = rcImage.left + sizeIcon.cx;
				rcImage.bottom = rcImage.top + sizeIcon.cy;
				BOOL bPrevPressed = m_bPressed;
				m_bPressed = FALSE;
				DrawImage(hDC, rcImage);
				m_bPressed = bPrevPressed;
			}

			if (!bXPLook)
			{
				if (VARIANT_TRUE == tpV1.m_vbChecked)
				{
					rcImage = rcPaint;
					rcImage.right = rcImage.left + nIconHeight;
					rcImage.bottom = rcImage.top + nIconHeight;
					DrawEdge(hDC, &rcImage, BDR_SUNKENOUTER, BF_RECT);
				}
				else if (m_bPressed)
				{
					rcImage = rcPaint;
					rcImage.right = rcImage.left + nIconHeight;
					rcImage.bottom = rcImage.top + nIconHeight;
					DrawEdge(hDC, &rcImage, BDR_RAISEDINNER, BF_RECT);
				}
			}
		}

		if (bText)
		{
			rcText = rcPaint;
			DrawText(hDC, 
					 rcText, 
					 btType, 
					 bVertical, 
					 szCaption, 
					 m_bstrCaption,
					 (VARIANT_TRUE == tpV1.m_vbChecked && !bChecked) || bIcon, 
					 VARIANT_TRUE == tpV1.m_vbChecked && bChecked,
					 nIconHeight);

			if (HasSubBand())
				DrawSubBandIndicator(hDC, rcPaint);
			else
			{
				LPTSTR szShortCut = GetShortCutString();
				if (lstrlen(szShortCut) > 0)
				{
					HFONT hFontText = GetMyFont(btType, bVertical);
					if (VARIANT_TRUE == tpV1.m_vbDefault)
					{
						LOGFONT lf;
						GetObject(hFontText, sizeof(LOGFONT), &lf);
						lf.lfWeight = FW_BOLD;
						hFontText = CreateFontIndirect(&lf);
					}
					HFONT hFontTextOld = SelectFont(hDC, hFontText);
					CRect rc = rcPaint;
					CRect rcTmp;
#ifdef _SCRIPTSTRING
					if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
					{
						MAKE_WIDEPTR_FROMTCHAR(wShortCut, szShortCut);
						ScriptDrawText(hDC, 
									   wShortCut, 
									   wcslen(wShortCut),
									   &rcTmp, 
									   DT_RIGHT|DT_SINGLELINE|DT_CALCRECT);
						rc.right -= 6;
						rcTmp.Inflate(2, 0);
						rc.left = rc.right - rcTmp.Width();
						DrawText(hDC, rc, btType, bVertical, szShortCut, wShortCut, FALSE, FALSE, 0, TRUE);
					}
					else
#endif
					{
						::DrawText(hDC, szShortCut, lstrlen(szShortCut), &rcTmp, DT_RIGHT|DT_CALCRECT);
						rcTmp.Inflate(2, 0);
						rc.right -= 6;
						rc.left = rc.right - rcTmp.Width();
						DrawText(hDC, rc, btType, bVertical, szShortCut, NULL, FALSE, FALSE, 0, TRUE);
					}
					SelectFont(hDC, hFontTextOld);
					if (VARIANT_TRUE == tpV1.m_vbDefault)
					{
						BOOL bResult = DeleteFont(hFontText);
						assert(bResult);
					}
				}
			}
		}
	}
	catch (...)
	{
		assert(FALSE);
	}
}

//
// DrawButton
//

void CTool::DrawButton(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical)
{

	CRect rcImage, rcText;
	SIZE sizeTotal, sizeChecked, sizeIcon, sizeText;
	BOOL bChecked, bIcon, bText;
	TCHAR szCaption[512];
	BOOL bResult;
	try
	{
		CleanCaption(szCaption, ddBTPopup == btType);
		CalcSizeEx(hDC, btType, bVertical, sizeTotal, sizeChecked, sizeIcon, sizeText, bChecked, bIcon, bText);
	}
	catch (...)
	{
		assert(FALSE);
	}
	
	switch (btType)
	{
	case ddBTPopup:
		{
			PopupBand(hDC, 
					  szCaption, 
					  btType, 
					  rcPaint, 
					  sizeTotal, 
					  sizeChecked, 
					  sizeIcon, 
					  sizeText, 
					  bVertical, 
					  bChecked, 
					  bIcon, 
					  bText);
		}
		break;

	default:
		{
			try
			{
				CRect rcPaint2 = rcPaint;

				BOOL bFlipVert = FALSE;
				BOOL bFlipHorz = FALSE;

				BOOL bMenuBar = ddBTMenuBar == m_pBand->bpV1.m_btBands || ddBTChildMenuBar == m_pBand->bpV1.m_btBands;
				BOOL bXPLook = VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook;
				BOOL bHasSubBand = HasSubBand();
				BOOL bSubBandIndicator = (bHasSubBand && !bMenuBar && bText && !bIcon);
				if (bSubBandIndicator)
					rcPaint2.right -= eToolWithSubBandWidth;

				CalcImageTextPos(rcPaint2, 
								 btType, 
								 bVertical, 
								 sizeTotal, 
								 sizeIcon, 
								 sizeText, 
								 bIcon, 
								 bText,
								 rcText,
								 rcImage);
				
				if (m_pBar && bXPLook && (m_pBar->m_pActiveTool == this || m_bPressed || m_bFocus))
				{
					if (!((bMenuBar && NULL != m_pSubBandWindow) || (bHasSubBand && NULL != m_pSubBandWindow)))
					{
						// Drawing the selected XP Look
						HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
						if (aPen)
						{
							HPEN penOld = SelectPen(hDC, aPen);

							HBRUSH aBrush = CreateSolidBrush((m_bPressed && !bMenuBar) ? m_pBar->m_crXPPressed : m_pBar->m_crXPSelectedColor);
							if (aBrush)
							{
								HBRUSH brushOld = SelectBrush(hDC, aBrush);
								
								if (ddCBSSlidingTabs == m_pBand->m_pRootBand->bpV1.m_cbsChildStyle && (ddSIcon == tpV1.m_tsStyle || ddSIconText == tpV1.m_tsStyle || ddSCheckedIconText == tpV1.m_tsStyle))
									Rectangle(hDC, rcImage.left, rcImage.top, rcImage.right, rcImage.bottom);
								else
									Rectangle(hDC, rcPaint2.left, rcPaint2.top, rcPaint2.right, rcPaint2.bottom);

								SelectBrush(hDC, brushOld);
								bResult = DeleteBrush(aBrush);
								assert(bResult);
							}
							
							SelectPen(hDC, penOld);
							bResult = DeletePen(aPen);
							assert(bResult);
						}
					}
					else if (bHasSubBand && m_pSubBandWindow && m_bPressed || bMenuBar)
					{
						// Drawing the selected XP Look on items that have subands
						HPEN aPen;
						CRect rcTool;
						CRect rcPopup;
						BOOL bFlipVert = FALSE;
						BOOL bFlipHorz = FALSE;
						if (m_pSubBandWindow)
						{
							bFlipHorz = m_pSubBandWindow->FlipHorz();
							bFlipVert = m_pSubBandWindow->FlipVert();
							if (m_pSubBandWindow->IsWindow())
							{
								CRect rcBand;
								HWND hWndBand;
								m_pBand->GetBandRect(hWndBand, rcBand);
								m_pSubBandWindow->GetWindowRect(rcPopup);
								ScreenToClient(hWndBand, rcPopup);
								m_pBand->GetToolBandRect(m_pBand->m_pTools->GetToolIndex(this), rcTool, hWndBand);
							}
							aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
						}
						else
							aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);

						CRect rcButton;
						if (ddCBSSlidingTabs == m_pBand->m_pRootBand->bpV1.m_cbsChildStyle && (ddSIcon == tpV1.m_tsStyle || ddSIconText == tpV1.m_tsStyle || ddSCheckedIconText == tpV1.m_tsStyle))
							rcButton = rcImage;
						else
							rcButton = rcPaint2;

						if (aPen)
						{
							HPEN penOld = SelectPen(hDC, aPen);

							if (bVertical)
							{
								if (!bFlipHorz && !bFlipVert)
								{
									MoveToEx(hDC, rcButton.right, rcButton.top, NULL);
									LineTo(hDC, rcButton.left, rcButton.top);
									LineTo(hDC, rcButton.left, rcButton.bottom-1);
									LineTo(hDC, rcButton.right, rcButton.bottom-1);

									if (rcTool.bottom > rcPopup.bottom || rcTool.top > rcPopup.top)
									{
										rcPopup.Offset(-rcTool.left, -rcTool.top);
										rcTool.Offset(-rcTool.left, -rcTool.top);
										MoveToEx(hDC, rcTool.right-1, rcPopup.bottom-1, NULL);
										LineTo(hDC, rcTool.right-1, rcTool.bottom-1);
										MoveToEx(hDC, rcTool.right-1, rcTool.top , NULL);
										LineTo(hDC, rcTool.right-1, rcPopup.top);
									}
								}
								else if (bFlipHorz && !bFlipVert)
								{
									MoveToEx(hDC, rcButton.left, rcButton.top, NULL);
									LineTo(hDC, rcButton.right, rcButton.top);
									LineTo(hDC, rcButton.right, rcButton.bottom);
									MoveToEx(hDC, rcButton.left, rcButton.bottom, NULL);
									LineTo(hDC, rcButton.right, rcButton.bottom);
									
									if (rcTool.bottom > rcPopup.bottom || rcTool.top > rcPopup.top)
									{
										rcPopup.Offset(-rcTool.left, -rcTool.top);
										rcTool.Offset(-rcTool.left, -rcTool.top);
										MoveToEx(hDC, rcTool.left, rcPopup.bottom-1, NULL);
										LineTo(hDC, rcTool.left, rcTool.bottom);
										MoveToEx(hDC, rcTool.left, rcTool.top, NULL);
										LineTo(hDC, rcTool.left, rcPopup.top);
									}
								}
								else if (!bFlipHorz && bFlipVert)
								{
									MoveToEx(hDC, rcButton.right-1, rcButton.top, NULL);
									LineTo(hDC, rcButton.left, rcButton.top);
									LineTo(hDC, rcButton.left, rcButton.bottom-1);
									LineTo(hDC, rcButton.right-1, rcButton.bottom-1);
									if (rcTool.bottom > rcPopup.bottom || rcPopup.top > rcTool.top)
									{
										rcPopup.Offset(-rcTool.left, -rcTool.top);
										rcTool.Offset(-rcTool.left, -rcTool.top);
										MoveToEx(hDC, rcTool.right-1, rcPopup.bottom-1, NULL);
										LineTo(hDC, rcTool.right-1, rcTool.bottom);
										MoveToEx(hDC, rcTool.right-1, rcTool.top, NULL);
										LineTo(hDC, rcTool.right-1, rcPopup.top);
									}
								}
								else
								{
									MoveToEx(hDC, rcButton.left, rcButton.top, NULL);
									LineTo(hDC, rcButton.right, rcButton.top);
									LineTo(hDC, rcButton.right, rcButton.bottom);
									MoveToEx(hDC, rcButton.left, rcButton.bottom, NULL);
									LineTo(hDC, rcButton.right, rcButton.bottom);
									if (rcTool.bottom > rcPopup.bottom || rcPopup.top > rcTool.top)
									{
										rcPopup.Offset(-rcTool.left, -rcTool.top);
										rcTool.Offset(-rcTool.left, -rcTool.top);
										MoveToEx(hDC, rcTool.left, rcPopup.bottom-1, NULL);
										LineTo(hDC, rcTool.left, rcTool.bottom);
										MoveToEx(hDC, rcTool.left, rcTool.top, NULL);
										LineTo(hDC, rcTool.left, rcPopup.top + 1);
									}
								}
							}
							else
							{
								if ((!bFlipHorz && !bFlipVert) || (bFlipHorz && !bFlipVert))
								{
									rcButton.Offset(-1, 0);
									MoveToEx(hDC, rcButton.left, rcButton.bottom, NULL);
									LineTo(hDC, rcButton.left, rcButton.top);
									LineTo(hDC, rcButton.right, rcButton.top);
									LineTo(hDC, rcButton.right, rcButton.bottom + 2);
								}
								else
								{
									rcButton.Offset(-1, 0);
									MoveToEx(hDC, rcButton.left, rcButton.top, NULL);
									LineTo(hDC, rcButton.left, rcButton.bottom);
									LineTo(hDC, rcButton.right, rcButton.bottom);
									LineTo(hDC, rcButton.right, rcButton.top - 2);
								}
							}
							SelectPen(hDC, penOld);
							bResult = DeletePen(aPen);
							assert(bResult);
						}

						rcButton.Inflate(-1, -1);
						if (ddCBSSlidingTabs == m_pBand->m_pRootBand->bpV1.m_cbsChildStyle && (ddSIcon == tpV1.m_tsStyle || ddSIconText == tpV1.m_tsStyle || ddSCheckedIconText == tpV1.m_tsStyle))
							FillSolidRect(hDC, rcButton, (m_bPressed && !bMenuBar) ? m_pBar->m_crXPPressed : m_pBar->m_crXPSelectedColor);
						else
							FillSolidRect(hDC, rcButton, m_pBar->m_crXPBandBackground);
					}
				}

				if (VARIANT_TRUE == tpV1.m_vbChecked && m_pBar->m_pActiveTool != this)
				{
					// Drawing Checked Non Active Tools
					if (bXPLook)
					{
						HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
						if (aPen)
						{
							HPEN penOld = SelectPen(hDC, aPen);

							HBRUSH aBrush = CreateSolidBrush(m_pBar->m_crXPCheckedColor);
							if (aBrush)
							{
								HBRUSH brushOld = SelectBrush(hDC, aBrush);

								if (ddCBSSlidingTabs == m_pBand->m_pRootBand->bpV1.m_cbsChildStyle && (ddSIcon == tpV1.m_tsStyle || ddSIconText == tpV1.m_tsStyle || ddSCheckedIconText == tpV1.m_tsStyle))
									Rectangle(hDC, rcImage.left, rcImage.top, rcImage.right, rcImage.bottom);
								else
									Rectangle(hDC, rcPaint2.left, rcPaint2.top, rcPaint2.right, rcPaint2.bottom);

								SelectBrush(hDC, brushOld);
								bResult = DeleteBrush(aBrush);
								assert(bResult);
							}
							
							SelectPen(hDC, penOld);
							bResult = DeletePen(aPen);
							assert(bResult);
						}
					}
					else
					{
						if (m_pBand && m_pBand->m_pParentBand && ddCBSSlidingTabs == m_pBand->m_pParentBand->bpV1.m_cbsChildStyle)
							DrawDitherRect(hDC, rcImage, m_pBar ? m_pBar->m_crHighLight : GetSysColor(COLOR_BTNHIGHLIGHT));
						else
							DrawDitherRect(hDC, rcPaint2, m_pBar ? m_pBar->m_crHighLight : GetSysColor(COLOR_BTNHIGHLIGHT));
					}
				}

				if (bIcon)
				{
					if (m_bPressed && !bXPLook)
						rcImage.Offset(CBand::eBevelBorder, CBand::eBevelBorder);
					DrawImage(hDC, rcImage);
				}

				if (bText)
				{
					if (m_bPressed && !bXPLook)
						rcText.Offset(CBand::eBevelBorder, CBand::eBevelBorder);
					DrawText(hDC, rcText, btType, bVertical, szCaption, m_bstrCaption);
				}

				if (bSubBandIndicator)
				{ 
					CRect rcDropDownImage = rcPaint;
					rcDropDownImage.left = rcDropDownImage.right - eToolWithSubBandWidth;
					if (m_bPressed)
						rcDropDownImage.Offset(1, 1);
					DrawDropDownImage(hDC, rcDropDownImage);
				}

				if (bSubBandIndicator)
					DrawBorder(hDC, rcPaint, rcImage, btType, bVertical, m_bPressed);
				else
					DrawBorder(hDC, rcPaint2, rcImage, btType, bVertical, m_bPressed);
			}
			catch (...)
			{
				assert(FALSE);
			}
		}
		break;
	}
}

//
// DrawBorder
//

void CTool::DrawBorder(HDC hDC, const CRect& rcPaint, CRect rcImage, BandTypes btType, BOOL bVertical, BOOL bPressed)
{
	BOOL bDrawBevel = TRUE;
	if ((VARIANT_FALSE == tpV1.m_vbEnabled && ddBTChildMenuBar != btType && ddBTMenuBar != btType) || 
		(m_pBand && ddTSNone == m_pBand->bpV1.m_tsMouseTracking) ||
		VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
	{
		bDrawBevel = FALSE;
	}
	
	if (bDrawBevel)
	{
		CRect rc = rcPaint;

		if ((VARIANT_TRUE == tpV1.m_vbChecked || bPressed) && m_dwCustomFlag <= 1)
		{
			if (m_pBand && m_pBand->m_pParentBand && ddCBSSlidingTabs == m_pBand->m_pParentBand->bpV1.m_cbsChildStyle)
			{
				if (ddSText == tpV1.m_tsStyle)
				{
					rc.Inflate(-CBand::eBevelBorder, -CBand::eBevelBorder);
					DrawSlidingEdge(hDC, rc, TRUE);
					return;
				}
				rcImage.Inflate(CBand::eBevelBorder, CBand::eBevelBorder);
				DrawSlidingEdge(hDC, rcImage, TRUE);
			}
			else
			{
				rc.Inflate(CBand::eBevelBorder, CBand::eBevelBorder);
				m_pBar->DrawEdge(hDC, rc, BDR_SUNKENOUTER, BF_RECT);
			}
		}
		else if ((VARIANT_FALSE == tpV1.m_vbChecked && !bPressed && m_pBar->m_pActiveTool == this && !Customization()) || m_bFocus)
		{
			if (m_pBand && m_pBand->m_pParentBand && ddCBSSlidingTabs == m_pBand->m_pParentBand->bpV1.m_cbsChildStyle)
			{
				if (ddSText == tpV1.m_tsStyle)
				{
					rc.Inflate(-CBand::eBevelBorder, -CBand::eBevelBorder);
					DrawSlidingEdge(hDC, rc, FALSE);
					return;
				}
				rcImage.Inflate(CBand::eBevelBorder, CBand::eBevelBorder);
				DrawSlidingEdge(hDC, rcImage, FALSE);
			}
			else
			{
				rc.Inflate(CBand::eBevelBorder, CBand::eBevelBorder);
				m_pBar->DrawEdge(hDC, rc, BDR_RAISEDINNER, BF_RECT);
			}
		}
	}
}

//
// DrawButtonDropDown
//

void CTool::DrawButtonDropDown(HDC			hDC, 
							   const CRect& rcPaint, 
							   BandTypes    btType, 
							   BOOL		    bVertical)
{
	TCHAR szCaption[512];
	CleanCaption(szCaption, ddBTPopup == btType);
	BOOL bResult;
	
	CRect rcImage, rcText;
	SIZE sizeTotal, sizeChecked, sizeIcon, sizeText;
	BOOL bChecked, bIcon, bText;
	CalcSizeEx(hDC, btType, bVertical, sizeTotal, sizeChecked, sizeIcon, sizeText, bChecked, bIcon, bText);
	switch (btType)
	{
	case ddBTPopup:
		{
			PopupBand(hDC, 
					  szCaption, 
					  btType, 
					  rcPaint, 
					  sizeTotal, 
					  sizeChecked,
					  sizeIcon, 
					  sizeText, 
					  bVertical, 
					  bChecked,
					  bIcon, 
					  bText);
		}
		break;

	default:
		{
			CRect rcPaint2 = rcPaint;
			CRect rcPaintBorder = rcPaint;

			BOOL bMenuBar = ddBTMenuBar == m_pBand->bpV1.m_btBands || ddBTChildMenuBar == m_pBand->bpV1.m_btBands;
			BOOL bXPLook = VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook;
			BOOL bHasSubBand = HasSubBand();

			rcPaint2.right -= (eButtonDropDownWidth * (VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 2 : 1)) + 1;

			CalcImageTextPos(rcPaint2, 
							 btType, 
							 bVertical, 
							 sizeTotal,
							 sizeIcon, 
							 sizeText, 
							 bIcon, 
							 bText,
							 rcText,
							 rcImage);

			if (m_pBar && bXPLook && (m_pBar->m_pActiveTool == this || m_bPressed || m_bFocus || m_bDropDownPressed))
			{
TRACE(5, "No SubBand Are we here\n");
				if (!((bMenuBar && NULL != m_pSubBandWindow) || (bHasSubBand && NULL != m_pSubBandWindow)))
				{
					// Drawing the selected XP Look
					rcPaintBorder.Inflate(1, 1);
					HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
					if (aPen)
					{
						HPEN penOld = SelectPen(hDC, aPen);

						HBRUSH aBrush = CreateSolidBrush((m_bPressed && !bMenuBar) ? m_pBar->m_crXPPressed : m_pBar->m_crXPSelectedColor);
						if (aBrush)
						{
							HBRUSH brushOld = SelectBrush(hDC, aBrush);
							
							Rectangle(hDC, rcPaintBorder.left, rcPaintBorder.top, rcPaintBorder.right, rcPaintBorder.bottom);

							MoveToEx(hDC, rcPaintBorder.right - eButtonDropDownWidth, rcPaintBorder.top, NULL);
							LineTo(hDC, rcPaintBorder.right - eButtonDropDownWidth, rcPaintBorder.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(aBrush);
							assert(bResult);
						}
						
						SelectPen(hDC, penOld);
						bResult = DeletePen(aPen);
						assert(bResult);
					}
				}
				else if (bHasSubBand && m_pSubBandWindow && (m_bPressed || m_bDropDownPressed) || bMenuBar)
				{
TRACE(5, "bHasSubBand Are we here\n");
					// Drawing the selected XP Look on items that have subands
					HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
					if (aPen)
					{
						HPEN penOld = SelectPen(hDC, aPen);
						CRect rcTool;
						CRect rcPopup;
						BOOL bFlipVert = FALSE;
						BOOL bFlipHorz = FALSE;
						if (m_pSubBandWindow)
						{
							bFlipHorz = m_pSubBandWindow->FlipHorz();
							bFlipVert = m_pSubBandWindow->FlipVert();
							if (m_pSubBandWindow->IsWindow())
							{
								CRect rcBand;
								HWND hWndBand;
								m_pBand->GetBandRect(hWndBand, rcBand);
								m_pSubBandWindow->GetWindowRect(rcPopup);
								ScreenToClient(hWndBand, rcPopup);
								m_pBand->GetToolBandRect(m_pBand->m_pTools->GetToolIndex(this), rcTool, hWndBand);
							}
						}
						if (bVertical)
						{
							if (!bFlipHorz && !bFlipVert)
							{
								MoveToEx(hDC, rcPaintBorder.right, rcPaintBorder.top, NULL);
								LineTo(hDC, rcPaintBorder.left, rcPaintBorder.top);
								LineTo(hDC, rcPaintBorder.left, rcPaintBorder.bottom-1);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.bottom-1);

								if (rcTool.bottom > rcPopup.bottom || rcTool.top > rcPopup.top)
								{
									rcPopup.Offset(-rcTool.left, -rcTool.top);
									rcTool.Offset(-rcTool.left, -rcTool.top);
									MoveToEx(hDC, rcTool.right-1, rcPopup.bottom-1, NULL);
									LineTo(hDC, rcTool.right-1, rcTool.bottom-1);
									MoveToEx(hDC, rcTool.right-1, rcTool.top , NULL);
									LineTo(hDC, rcTool.right-1, rcPopup.top);
								}
							}
							else if (bFlipHorz && !bFlipVert)
							{
								MoveToEx(hDC, rcPaintBorder.left, rcPaintBorder.top, NULL);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.top);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.bottom);
								MoveToEx(hDC, rcPaintBorder.left, rcPaintBorder.bottom, NULL);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.bottom);
								
								if (rcTool.bottom > rcPopup.bottom || rcTool.top > rcPopup.top)
								{
									rcPopup.Offset(-rcTool.left, -rcTool.top);
									rcTool.Offset(-rcTool.left, -rcTool.top);
									MoveToEx(hDC, rcTool.left, rcPopup.bottom-1, NULL);
									LineTo(hDC, rcTool.left, rcTool.bottom);
									MoveToEx(hDC, rcTool.left, rcTool.top, NULL);
									LineTo(hDC, rcTool.left, rcPopup.top);
								}
							}
							else if (!bFlipHorz && bFlipVert)
							{
								MoveToEx(hDC, rcPaintBorder.right-1, rcPaintBorder.top, NULL);
								LineTo(hDC, rcPaintBorder.left, rcPaintBorder.top);
								LineTo(hDC, rcPaintBorder.left, rcPaintBorder.bottom-1);
								LineTo(hDC, rcPaintBorder.right-1, rcPaintBorder.bottom-1);
								if (rcTool.bottom > rcPopup.bottom || rcPopup.top > rcTool.top)
								{
									rcPopup.Offset(-rcTool.left, -rcTool.top);
									rcTool.Offset(-rcTool.left, -rcTool.top);
									MoveToEx(hDC, rcTool.right-1, rcPopup.bottom-1, NULL);
									LineTo(hDC, rcTool.right-1, rcTool.bottom);
									MoveToEx(hDC, rcTool.right-1, rcTool.top, NULL);
									LineTo(hDC, rcTool.right-1, rcPopup.top);
								}
							}
							else
							{
								MoveToEx(hDC, rcPaintBorder.left, rcPaintBorder.top, NULL);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.top);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.bottom);
								MoveToEx(hDC, rcPaintBorder.left, rcPaintBorder.bottom, NULL);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.bottom);
								if (rcTool.bottom > rcPopup.bottom || rcPopup.top > rcTool.top)
								{
									rcPopup.Offset(-rcTool.left, -rcTool.top);
									rcTool.Offset(-rcTool.left, -rcTool.top);
									MoveToEx(hDC, rcTool.left, rcPopup.bottom-1, NULL);
									LineTo(hDC, rcTool.left, rcTool.bottom);
									MoveToEx(hDC, rcTool.left, rcTool.top, NULL);
									LineTo(hDC, rcTool.left, rcPopup.top + 1);
								}
							}
						}
						else
						{
							if ((!bFlipHorz && !bFlipVert) || (bFlipHorz && !bFlipVert))
							{
								rcPaintBorder.Offset(-1, 0);
								MoveToEx(hDC, rcPaintBorder.left, rcPaintBorder.bottom, NULL);
								LineTo(hDC, rcPaintBorder.left, rcPaintBorder.top);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.top);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.bottom + 2);
							}
							else
							{
								rcPaintBorder.Offset(-1, 0);
								MoveToEx(hDC, rcPaintBorder.left, rcPaintBorder.top, NULL);
								LineTo(hDC, rcPaintBorder.left, rcPaintBorder.bottom);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.bottom);
								LineTo(hDC, rcPaintBorder.right, rcPaintBorder.top - 2);
							}
						}
						SelectPen(hDC, penOld);
						bResult = DeletePen(aPen);
						assert(bResult);
					}

					rcPaintBorder.Inflate(-1, -1);
					FillSolidRect(hDC, rcPaintBorder, m_pBar->m_crXPBandBackground);
				}
				else
					TRACE(5, "XP no drawing\n");
			}
			else
				TRACE(5, "No XP\n");


			if (bIcon)
				DrawImage(hDC, rcImage);

			if (bText)
				DrawText(hDC, rcText, btType, bVertical, szCaption, m_bstrCaption);

			BOOL bWasActive = m_pBar->m_pActiveTool == this;
			if (m_bDropDownPressed && !bWasActive)
				m_pBar->m_pActiveTool = this;

			DrawBorder(hDC, rcPaint2, rcImage, btType, bVertical, m_bPressed);

			if (!bWasActive)
				m_pBar->m_pActiveTool = NULL;

			CRect rcDropDownImage = rcPaint;
			rcDropDownImage.left = rcDropDownImage.right - (eButtonDropDownWidth * (VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 2 : 1)) + 1;
			
			DrawDropDownImage(hDC, rcDropDownImage, VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons);

			VARIANT_BOOL vbTemp = tpV1.m_vbChecked;
			tpV1.m_vbChecked = VARIANT_FALSE;
			DrawBorder(hDC, rcDropDownImage, rcImage, btType, bVertical, m_bDropDownPressed);
			tpV1.m_vbChecked = vbTemp;
		}
		break;
	}
}

//
// DrawCombobox
//

void CTool::DrawCombobox(HDC hDC, CRect rcPaint, BandTypes btType, BOOL bVertical)
{
	TCHAR szCaption[512];
	CleanCaption(szCaption, ddBTPopup == btType);
	BOOL bResult;
	
	CRect rcImage, rcText;
	SIZE sizeTotal, sizeChecked, sizeIcon, sizeText;
	BOOL bChecked, bIcon, bText;
	CalcSizeEx(hDC, btType, bVertical, sizeTotal, sizeChecked, sizeIcon, sizeText, bChecked, bIcon, bText);

	CRect rcCombo = rcPaint;
	switch (btType)
	{
	case ddBTPopup:
		{
			CRect rcText = rcPaint;
			rcText.right = rcText.left + m_nComboNameOffset;
			if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				if (m_bPressed)
				{
					HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
					if (aPen)
					{
						HPEN penOld = SelectPen(hDC, aPen);
						HBRUSH aBrush;
						
						if (VARIANT_FALSE == tpV1.m_vbEnabled)
							aBrush = CreateSolidBrush(m_pBar->m_crXPMenuBackground);
						else
							aBrush = CreateSolidBrush(m_pBar->m_crXPSelectedColor);

						if (aBrush)
						{
							HBRUSH brushOld = SelectBrush(hDC, aBrush);
							
							Rectangle(hDC, rcPaint.left, rcPaint.top, rcPaint.right, rcPaint.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(aBrush);
							assert(bResult);
						}
						SelectPen(hDC, penOld);
						bResult = DeletePen(aPen);
						assert(bResult);
					}
				}
				rcText.left += 20;
			}
			DrawText(hDC, rcText, btType, bVertical, szCaption, m_bstrCaption);
			rcCombo.left += m_nComboNameOffset;
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				DrawXPCombo(hDC, rcCombo, m_bPressed);
			else
				DrawCombo(hDC, rcCombo, m_bPressed);
		}
		break;

	default:
		switch(tpV1.m_tsStyle)
		{
		case ddSText:
		case ddSIconText:
		case ddSCheckedIconText:
			if (m_bstrCaption && *m_bstrCaption)
			{
				CRect rcText = rcPaint;
				rcText.right = rcText.left + m_nComboNameOffset;
				DrawText(hDC, rcText, btType, bVertical, szCaption, m_bstrCaption);
				rcCombo.left += m_nComboNameOffset;
			}
			break;
		}
		if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			DrawXPCombo(hDC, rcCombo, m_bPressed);
		else
			DrawCombo(hDC, rcCombo, m_bPressed);
		break;
	}
}

//
// DrawEdit
//

void CTool::DrawEdit(HDC hDC, CRect rcPaint, BandTypes btType, BOOL bVertical)
{
	TCHAR szCaption[512];
	CleanCaption(szCaption, ddBTPopup == btType);
	
	CRect rcImage;
	SIZE sizeTotal, sizeChecked, sizeIcon, sizeText;
	BOOL bChecked, bIcon, bText;
	CalcSizeEx(hDC, btType, bVertical, sizeTotal, sizeChecked, sizeIcon, sizeText, bChecked, bIcon, bText);

	CRect rcEdit = rcPaint;
	CRect rcText = rcPaint;
	switch (btType)
	{
	case ddBTPopup:
		{
			if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				if (m_bPressed)
				{
					HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
					HBRUSH aBrush = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
					if (VARIANT_FALSE == tpV1.m_vbEnabled)
						aBrush = CreateSolidBrush(m_pBar->m_crXPMenuBackground);
					else
						aBrush = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
					HBRUSH brushOld = SelectBrush(hDC, aBrush);
					HPEN penOld = SelectPen(hDC, aPen);
					Rectangle(hDC, rcPaint.left, rcPaint.top, rcPaint.right, rcPaint.bottom);
					SelectBrush(hDC, brushOld);
					SelectPen(hDC, penOld);
				}
				rcText.left += 20;
			}
			rcText.right = rcText.left + m_nComboNameOffset;
			DrawText(hDC, rcText, btType, bVertical, szCaption, m_bstrCaption);
			rcEdit.left += m_nComboNameOffset;
			DrawEdit(hDC, rcEdit);
		}
		break;

	default:
		switch(tpV1.m_tsStyle)
		{
		case ddSText:
		case ddSIconText:
		case ddSCheckedIconText:
			if (m_bstrCaption && *m_bstrCaption)
			{
				rcText.right = rcText.left + m_nComboNameOffset;
				DrawText(hDC, rcText, btType, bVertical, szCaption, m_bstrCaption);
				rcEdit.left += m_nComboNameOffset;
			}
			break;
		}
		DrawEdit(hDC, rcEdit);
		break;
	}
}

//
// DrawStatic
//

void CTool::DrawStatic(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical)
{
	TCHAR szCaption[256];
	switch (tpV1.m_lsLabelStyle)
	{
	case ddLSInsert:
		{
			if (GetKeyState(VK_INSERT) & 0x0001)
				tpV1.m_vbEnabled = VARIANT_TRUE;
			else
				tpV1.m_vbEnabled = VARIANT_FALSE;
		}
		break;
	
	case ddLSCapitalLock:
		{
			if (GetKeyState(VK_CAPITAL) & 0x0001)
				tpV1.m_vbEnabled = VARIANT_TRUE;
			else
				tpV1.m_vbEnabled = VARIANT_FALSE;
		}
		break;
	
	case ddLSNumberLock:
		{
			if (GetKeyState(VK_NUMLOCK) & 0x0001)
				tpV1.m_vbEnabled = VARIANT_TRUE;
			else
				tpV1.m_vbEnabled = VARIANT_FALSE;
		}
		break;

	case ddLSScrollLock:
		{
			if (GetKeyState(VK_SCROLL) & 0x0001)
				tpV1.m_vbEnabled = VARIANT_TRUE;
			else
				tpV1.m_vbEnabled = VARIANT_FALSE;
		}
		break;

	case ddLSDate:
		{	
			TCHAR szDate[60];
			GetDateFormat(LOCALE_SYSTEM_DEFAULT, DATE_SHORTDATE, &m_theSystemTime, NULL, szDate, 60);
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szDate);
			put_Caption(wBuffer);
		}
		break;

	case ddLSTime:
		{	
			TCHAR szTime[60];
			GetTimeFormat(LOCALE_SYSTEM_DEFAULT, TIME_NOSECONDS, &m_theSystemTime, NULL, szTime, 60);
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szTime);
			put_Caption(wBuffer);
		}
		break;
	}
	CleanCaption(szCaption, ddBTPopup == btType);
	
	SIZE sizeTotal, sizeChecked, sizeIcon, sizeText;
	BOOL bChecked, bIcon, bText;
	CalcSizeEx(hDC, btType, bVertical, sizeTotal, sizeChecked, sizeIcon, sizeText, bChecked, bIcon, bText);

	switch (btType)
	{
	case ddBTPopup:
		{
			PopupBand(hDC, 
					  szCaption, 
					  btType, 
					  rcPaint, 
					  sizeTotal, 
					  sizeChecked,
					  sizeIcon, 
					  sizeText, 
					  bVertical, 
					  bChecked,
					  bIcon, 
					  bText);
		}
		break;

	default:
		{
			CRect rcImage, rcText;
			CalcImageTextPos(rcPaint, 
							 btType, 
							 bVertical, 
							 sizeTotal, 
							 sizeIcon, 
							 sizeText, 
							 bIcon, 
							 bText,
							 rcText,
							 rcImage);
			if (bIcon && !rcImage.IsEmpty())
				DrawImage(hDC, rcImage);

			if (bText && !rcText.IsEmpty())
				DrawText(hDC, rcText, btType, bVertical, szCaption, m_bstrCaption);

			switch (tpV1.m_lsLabelBevel)
			{
			case ddLBInset:
				{
					CRect rcEdge = rcPaint;
					DrawEdge(hDC, const_cast<CRect*>(&rcEdge), BDR_SUNKENOUTER, BF_RECT);
				}
				break;

			case ddLBRaised:
				{
					CRect rcEdge = rcPaint;
					DrawEdge(hDC, const_cast<CRect*>(&rcEdge), BDR_RAISEDINNER, BF_RECT);
				}
				break;

			case ddLBFlat:
				{
					if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
					{
						CRect rcEdge = rcPaint;
						HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPBandBackground);
						if (penBorder)
						{
							HPEN penOld = SelectPen(hDC, penBorder);

							HBRUSH brushOld = SelectBrush(hDC, (HBRUSH)GetStockObject(HOLLOW_BRUSH));
							
							Rectangle(hDC, rcEdge.left, rcEdge.top, rcEdge.right, rcEdge.bottom);

							SelectBrush(hDC, brushOld);

							SelectPen(hDC, penOld);
							BOOL bResult = DeletePen(penBorder);
							assert(bResult);
						}
					}
				}
				break;
			}
		}
		break;
	}
}

//
// DrawSeparator
//

void CTool::DrawSeparator(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical)
{
	BOOL bResult;
	CRect rcSep = rcPaint;
	if (m_bVerticalSeparator)
	{
		rcSep.Inflate(0, -2);
		if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
		{
			rcSep.top++;
			HBRUSH hBrush = CreateSolidBrush(RGB(165, 166, 165));
			if (hBrush)
			{
	
				FillRect(hDC, &rcSep, hBrush);
				
				bResult = DeleteBrush(hBrush);
				assert(bResult);
			}
		}
		else
			m_pBar->DrawEdge(hDC, rcSep, BDR_SUNKENOUTER, BF_RECT);
	}
	else
	{
		switch (btType)
		{
		case ddBTPopup:
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				rcSep.left += m_pBand->m_sizeMaxIcon.cx;
			if (m_bstrCaption && *m_bstrCaption)
			{
				TCHAR szCaption[512];
				CleanCaption(szCaption, ddBTPopup == btType);
				
				CRect rcImage, rcText;
				SIZE sizeTotal, sizeChecked, sizeIcon, sizeText;
				BOOL bChecked, bIcon, bText;
				tpV1.m_ttTools = ddTTButton;
				CalcSizeEx(hDC, btType, bVertical, sizeTotal, sizeChecked, sizeIcon, sizeText, bChecked, bIcon, bText);
				CalcImageTextPos(rcPaint, 
								 btType, 
								 bVertical, 
								 sizeTotal, 
								 sizeIcon, 
								 sizeText, 
								 bIcon, 
								 bText,
								 rcText,
								 rcImage);
				tpV1.m_ttTools = ddTTSeparator;
				HFONT hFontText = GetMyFont(btType, FALSE);
				HFONT hFontOld = SelectFont(hDC, hFontText);
				int nModeOld = SetBkMode(hDC, TRANSPARENT);
				COLORREF crTextOld = SetTextColor(hDC, m_pBar->m_crForeground);

				int nTextLen = lstrlen(szCaption);
				::DrawText(hDC, szCaption, nTextLen, &rcText, DT_CENTER|DT_VCENTER);

				rcSep = rcPaint;
				rcSep.right = rcText.left - 2;
				rcSep.top +=((rcPaint.Height() - 2) / 2);
				rcSep.bottom = rcSep.top + 2;
				rcSep.left += 8;
				m_pBar->DrawEdge(hDC, rcSep, BDR_SUNKENOUTER, BF_RECT);

				rcSep = rcPaint;
				rcSep.left = rcText.right + 2;
				rcSep.top += ((rcPaint.Height() - 2) / 2);
				rcSep.bottom = rcSep.top + 2;
				rcSep.right -= 8;
				m_pBar->DrawEdge(hDC, rcSep, BDR_SUNKENOUTER, BF_RECT);

				SelectFont(hDC, hFontOld);
				SetBkMode(hDC, nModeOld);
				SetTextColor(hDC, crTextOld);
			}
			else
			{
				rcSep.Inflate(-8, -2);
				m_pBar->DrawEdge(hDC, rcSep, BDR_SUNKENOUTER, BF_RECT);
			}
			break;

		default:
			rcSep.Inflate(-2, 0);
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				rcSep.left++;
				HBRUSH hBrush = CreateSolidBrush(RGB(165, 166, 165));
				if (hBrush)
				{
					FillRect(hDC, &rcSep, hBrush);
					bResult = DeleteBrush(hBrush);
					assert(bResult);
				}
			}
			else
				m_pBar->DrawEdge(hDC, rcSep, BDR_SUNKENOUTER, BF_RECT);
			break;
		}
	}
}

//
// DrawMDIButtons
//

void CTool::DrawMDIButtons(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical)
{
	BOOL bResult;
	CRect rcButton = rcPaint;
	m_bMDICloseEnabled = TRUE;
	HWND hWndActive = (HWND)SendMessage(m_pBar->m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
	if (hWndActive)
	{
		HMENU hMenu = GetSystemMenu(hWndActive, FALSE);
		if (hMenu)
		{
			TCHAR szName[64];
			MENUITEMINFO theMenuInfo;
			memset(&theMenuInfo, 0, sizeof(MENUITEMINFO));
			theMenuInfo.cbSize = sizeof(MENUITEMINFO);
			theMenuInfo.fMask = MIIM_TYPE | MIIM_ID;
			theMenuInfo.fType = MFT_STRING;
			theMenuInfo.dwTypeData = szName;
			theMenuInfo.cch = 64;
			if (GetMenuItemInfo(hMenu, SC_CLOSE, FALSE, &theMenuInfo))
			{
				if (theMenuInfo.fState & MFS_DISABLED)
					m_bMDICloseEnabled = FALSE;
			}
		}
	}

	if (bVertical)
	{
		rcButton.Offset(-2, 0);
		rcButton.top = rcButton.bottom - m_pBar->m_sizeSMButton.cy;
		// CloseBox
		if (CBar::eCloseWindow & m_pBar->m_dwMdiButtons)
		{
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				if (eClose == m_mdibActive)
				{
					HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
					if (penBorder)
					{
						HPEN penOld = SelectPen(hDC, penBorder);

						HBRUSH brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
						if (brushBackground)
						{
							HBRUSH brushOld = SelectBrush(hDC, brushBackground);
							
							Rectangle(hDC, rcButton.left, rcButton.top, rcButton.right, rcButton.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(brushBackground);
							assert(bResult);
						}
						SelectPen(hDC, penOld);
						bResult = DeletePen(penBorder);
						assert(bResult);
					}
				}

				HDC hMemDC = GetMemDC();
				HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMDIButtons());

				BitBltMasked(hDC, 
							 rcButton.left + (rcButton.Width() - 8) / 2,
							 rcButton.top + (rcButton.Height() - 9) / 2,
							 8,
							 9,
							 hMemDC,
							 8 * 2,
							 0,
							 8,
							 9,
							 m_pBar->m_crForeground);

				SelectBitmap(hMemDC, hBitmapOld);
			}
			else
			{
				DrawFrameControl(hDC, 
								 &rcButton, 
								 DFC_CAPTION, 
								 m_bMDICloseEnabled ? DFCS_CAPTIONCLOSE : (DFCS_CAPTIONCLOSE | DFCS_INACTIVE));
			}
		}
		if ((CBar::eRestoreWindow|CBar::eMinimizeWindow) & m_pBar->m_dwMdiButtons)
		{
			rcButton.Offset(0, -(m_pBar->m_sizeSMButton.cy + 2));

			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				HDC hMemDC = GetMemDC();
				HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMDIButtons());

				if (eRestore == m_mdibActive)
				{
					HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
					if (penBorder)
					{
						HPEN penOld = SelectPen(hDC, penBorder);

						HBRUSH brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
						if (brushBackground)
						{
							HBRUSH brushOld = SelectBrush(hDC, brushBackground);
							
							Rectangle(hDC, rcButton.left, rcButton.top, rcButton.right, rcButton.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(brushBackground);
							assert(bResult);
						}
						SelectPen(hDC, penOld);
						bResult = DeletePen(penBorder);
						assert(bResult);
					}
				}

				BitBltMasked(hDC, 
							 rcButton.left + (rcButton.Width() - 8) / 2,
							 rcButton.top + (rcButton.Height() - 9) / 2,
							 8,
							 9,
							 hMemDC,
							 8,
							 0,
							 8,
							 9,
							 m_pBar->m_crForeground);

				rcButton.Offset(0, -m_pBar->m_sizeSMButton.cy);

				if (eMinimize == m_mdibActive)
				{
					HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
					if (penBorder)
					{
						HPEN penOld = SelectPen(hDC, penBorder);

						HBRUSH brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
						if (brushBackground)
						{
							HBRUSH brushOld = SelectBrush(hDC, brushBackground);
							
							Rectangle(hDC, rcButton.left, rcButton.top, rcButton.right, rcButton.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(brushBackground);
							assert(bResult);
						}
						SelectPen(hDC, penOld);
						bResult = DeletePen(penBorder);
						assert(bResult);
					}
				}
				BitBltMasked(hDC, 
							 rcButton.left + (rcButton.Width() - 8) / 2,
							 rcButton.top + (rcButton.Height() - 9) / 2,
							 8,
							 9,
							 hMemDC,
							 0,
							 0,
							 8,
							 9,
							 m_pBar->m_crForeground);
				SelectBitmap(hMemDC, hBitmapOld);
			}
			else
			{
				//	Restore
				DrawFrameControl(hDC,
								 &rcButton,
								 DFC_CAPTION,
								 (CBar::eRestoreWindow & m_pBar->m_dwMdiButtons) ? DFCS_CAPTIONRESTORE : (DFCS_CAPTIONRESTORE | DFCS_INACTIVE));

				rcButton.Offset(0, -m_pBar->m_sizeSMButton.cy);
				
				//	Minimize
				DrawFrameControl(hDC,
								 &rcButton,
								 DFC_CAPTION,
								 (CBar::eMinimizeWindow & m_pBar->m_dwMdiButtons) ? DFCS_CAPTIONMIN : (DFCS_CAPTIONMIN | DFCS_INACTIVE));
			}
		}
	}
	else
	{
		rcButton.top += 2;
		rcButton.left = rcButton.right - m_pBar->m_sizeSMButton.cx;
		rcButton.bottom = rcButton.top + m_pBar->m_sizeSMButton.cy;
		
		if (CBar::eCloseWindow & m_pBar->m_dwMdiButtons)
		{
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				if (eClose == m_mdibActive)
				{
					HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
					if (penBorder)
					{
						HPEN penOld = SelectPen(hDC, penBorder);

						HBRUSH brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
						if (brushBackground)
						{
							HBRUSH brushOld = SelectBrush(hDC, brushBackground);
							
							Rectangle(hDC, rcButton.left, rcButton.top, rcButton.right, rcButton.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(brushBackground);
							assert(bResult);
						}
						SelectPen(hDC, penOld);
						bResult = DeletePen(penBorder);
						assert(bResult);
					}
				}

				HDC hMemDC = GetMemDC();
				if (hMemDC)
				{
					HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMDIButtons());

					BitBltMasked(hDC, 
								 rcButton.left + (rcButton.Width() - 8) / 2,
								 rcButton.top + (rcButton.Height() - 9) / 2,
								 8,
								 9,
								 hMemDC,
								 8 * 2,
								 0,
								 8,
								 9,
								 m_pBar->m_crForeground);

					SelectBitmap(hMemDC, hBitmapOld);
				}
			}
			else
			{
				DrawFrameControl(hDC,
								 &rcButton,
								 DFC_CAPTION,
								 m_bMDICloseEnabled ? DFCS_CAPTIONCLOSE : (DFCS_CAPTIONCLOSE | DFCS_INACTIVE));
			}
		}
		if ((CBar::eRestoreWindow | CBar::eMinimizeWindow) & m_pBar->m_dwMdiButtons)
		{
			rcButton.Offset(-(m_pBar->m_sizeSMButton.cx + 2), 0);

			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				HDC hMemDC = GetMemDC();
				if (hMemDC)
				{
					HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMDIButtons());

					if (eRestore == m_mdibActive)
					{
						HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
						if (penBorder)
						{
							HPEN penOld = SelectPen(hDC, penBorder);

							HBRUSH brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
							if (brushBackground)
							{
								HBRUSH brushOld = SelectBrush(hDC, brushBackground);
								
								Rectangle(hDC, rcButton.left, rcButton.top, rcButton.right, rcButton.bottom);

								SelectBrush(hDC, brushOld);
								bResult = DeleteBrush(brushBackground);
								assert(bResult);
							}
							SelectPen(hDC, penOld);
							bResult = DeletePen(penBorder);
							assert(bResult);
						}
					}

					BitBltMasked(hDC, 
								 rcButton.left + (rcButton.Width() - 8) / 2,
								 rcButton.top + (rcButton.Height() - 9) / 2,
								 8,
								 9,
								 hMemDC,
								 8,
								 0,
								 8,
								 9,
								 m_pBar->m_crForeground);

					rcButton.Offset(-m_pBar->m_sizeSMButton.cx, 0);

					if (eMinimize == m_mdibActive)
					{
						HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
						if (penBorder)
						{
							HPEN penOld = SelectPen(hDC, penBorder);

							HBRUSH brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
							if (brushBackground)
							{
								HBRUSH brushOld = SelectBrush(hDC, brushBackground);
								
								Rectangle(hDC, rcButton.left, rcButton.top, rcButton.right, rcButton.bottom);

								SelectBrush(hDC, brushOld);
								bResult = DeleteBrush(brushBackground);
								assert(bResult);
							}
							SelectPen(hDC, penOld);
							bResult = DeletePen(penBorder);
							assert(bResult);
						}
					}

					BitBltMasked(hDC, 
								 rcButton.left + (rcButton.Width() - 8) / 2,
								 rcButton.top + (rcButton.Height() - 9) / 2,
								 8,
								 9,
								 hMemDC,
								 0,
								 0,
								 8,
								 9,
								 m_pBar->m_crForeground);

					SelectBitmap(hMemDC, hBitmapOld);
				}
			}
			else
			{
				DrawFrameControl(hDC,
								 &rcButton,
								 DFC_CAPTION,
								 (CBar::eRestoreWindow & m_pBar->m_dwMdiButtons) ? DFCS_CAPTIONRESTORE : (DFCS_CAPTIONRESTORE | DFCS_INACTIVE));

				rcButton.Offset(-m_pBar->m_sizeSMButton.cx, 0);
				
				DrawFrameControl(hDC,
								 &rcButton,
								 DFC_CAPTION,
								 (CBar::eMinimizeWindow & m_pBar->m_dwMdiButtons) ? DFCS_CAPTIONMIN: (DFCS_CAPTIONMIN | DFCS_INACTIVE));
			}
		}
	}
}

//
// DrawChildSysMenu
//

void CTool::DrawChildSysMenu(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical)
{
	if (CBar::eSysMenu & m_pBar->m_dwMdiButtons)
	{
		//
		// Get Child Window's Icon
		//

		HICON hIcon = NULL;
		HWND hWndActive = (HWND)SendMessage(m_pBar->m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
		if (hWndActive)
		{
			// Try to load the small icon
			hIcon = (HICON)SendMessage(hWndActive, WM_GETICON, (WPARAM)ICON_SMALL, 0);
			if (NULL == hIcon)
			{
				// Try to load big one
				hIcon = (HICON)SendMessage(hWndActive, WM_GETICON, (WPARAM)ICON_BIG, 0);
				if (NULL == hIcon)
				{
					hIcon = (HICON)GetClassLong(hWndActive, GCL_HICONSM);
					if (NULL == hIcon)
						// Last chance is the Window's icon
						hIcon = LoadIcon(NULL, MAKEINTRESOURCE(IDI_WINLOGO));
				}
			}
			CRect rcImage, rcText;
			SIZE sizeTotal, sizeChecked, sizeIcon, sizeText;
			BOOL bChecked, bIcon, bText;
			CalcSizeEx(hDC, btType, bVertical, sizeTotal, sizeChecked, sizeIcon, sizeText, bChecked, bIcon, bText);
			DrawIconEx(hDC, 
					   rcPaint.left, 
					   rcPaint.top,
					   hIcon,
					   sizeIcon.cx,
					   sizeIcon.cy,
					   0,
					   0,
					   DI_NORMAL);
		}
	}
}

//
// DrawMenuExpand
//

void CTool::DrawMenuExpand(HDC hDC, CRect rcPaint)
{
	BOOL bResult;
	if (m_bPressed && !Customization())
	{
		if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
		{
			HPEN penBack = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
			if (penBack)
			{
				HPEN penOld = SelectPen(hDC, penBack);
				
				HBRUSH hBrush = CreateSolidBrush(m_pBand->m_pBar->m_crXPSelectedColor);
				if (hBrush)
				{
					HBRUSH brushOld = SelectBrush(hDC, hBrush);

					RoundRect(hDC, rcPaint.left, rcPaint.top, rcPaint.right, rcPaint.bottom, 1, 1);
					
					SelectBrush(hDC, brushOld);
					bResult = DeleteBrush(hBrush);
					assert(bResult);
				}
				SelectPen(hDC, penOld);
				bResult = DeletePen(penBack);
				assert(bResult);
			}
			rcPaint.Inflate(-1, -1);
		}
		else
		{
			m_pBar->DrawEdge(hDC, rcPaint, m_bDropDownPressed ? BDR_SUNKENOUTER : BDR_RAISEDINNER, BF_RECT);
			rcPaint.Inflate(-1, -1);
			FillSolidRect(hDC, rcPaint, m_pBar->m_crMDIMenuBackground);
		}
		rcPaint.Inflate(-1, -1);
	}
	else
		rcPaint.Inflate(-2, -2);

	HDC hMemDC = GetMemDC();
	HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMenuExpand());

	int nHeight = rcPaint.Height()-4; 
	
	if (nHeight < 8)
		nHeight = 8;
	BitBltMasked(hDC, 
			     rcPaint.left + (rcPaint.Width() - nHeight)/2,
			     rcPaint.top + (rcPaint.Height() - nHeight)/2,
			     nHeight,
			     nHeight,
			     hMemDC,
			     0,
			     0,
			     9,
			     8,
			     m_pBar->m_crForeground);

	SelectBitmap(hMemDC, hBitmapOld);
}

//
// DrawClipMarker
//

BOOL CTool::DrawClipMarker(HDC hDC, CBar* pBar, CRect rcClipMarker, BOOL bVertical, BOOL bLargeIcons)
{
	HBITMAP hClipMarker = GetGlobals().GetBitmapEndMarker();
	if (NULL == hClipMarker)
		return FALSE;

	BOOL bResult;
	HDC hMemDC = GetMemDC();
	assert(hMemDC);
	if (NULL == hMemDC)
		return FALSE;

	HBITMAP hBitmapOld = SelectBitmap(hMemDC, hClipMarker);

	int nX = rcClipMarker.left + 2;
	int nY = rcClipMarker.top + 2;

	if (bVertical)
	{
		nX = rcClipMarker.right - 2 - eClipMarkerHeight * (bLargeIcons ? 2 : 1);
		bResult = BitBltMasked(hDC,
							   nX,
							   nY,
							   eClipMarkerHeight * (bLargeIcons ? 2 : 1),
							   eClipMarkerWidth * (bLargeIcons ? 2 : 1),
							   hMemDC,
							   eClipMarkerWidth,
							   0,
							   eClipMarkerHeight,
							   eClipMarkerWidth,
							   pBar->m_crForeground);
	}
	else
	{
		bResult= BitBltMasked(hDC,
							  nX,
							  nY,
							  eClipMarkerWidth * (bLargeIcons ? 2 : 1),
							  eClipMarkerHeight * (bLargeIcons ? 2 : 1),
							  hMemDC,
							  0,
							  0,
							  eClipMarkerWidth,
							  eClipMarkerHeight,
							  pBar->m_crForeground);
	}
	SelectBitmap(hMemDC, hBitmapOld);
	return TRUE;
}

//
// EnableMoreTools 
//

BOOL CTool::EnableMoreTools()
{
	if (!(ddBFCustomize & m_pBand->bpV1.m_dwFlags))
	{
		CTool* pTool;
		int nCount = m_pBand->m_pTools->GetVisibleToolCount();
		for (int nTool = 0; nTool < nCount; nTool++)
		{
			pTool = m_pBand->m_pTools->GetVisibleTool(nTool);

			if (pTool->IsVisibleOnPaint()  || ddTTSeparator == pTool->tpV1.m_ttTools || VARIANT_FALSE == pTool->tpV1.m_vbVisible)
				continue;
			return TRUE;
		}
		return FALSE;
	}
	return TRUE;
}

//
// DrawMoreTools
//

void CTool::DrawMoreTools(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical)
{
	BOOL bResult;
	CRect rcTemp;
	BOOL bEnable = EnableMoreTools();
	if (bEnable)
	{
		switch (btType)
		{
		case ddBTPopup:
			DrawBorder(hDC, rcPaint, rcTemp, btType, bVertical, m_bDropDownPressed);
			break;

		default:
			DrawBorder(hDC, rcPaint, rcTemp, btType, bVertical, m_bPressed);
			break;
		}
	}

	if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
	{
		if (m_pBar->m_pActiveTool == this)
		{
			HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
			if (aPen)
			{
				HPEN penOld = SelectPen(hDC, aPen);
				HBRUSH aBrush = CreateSolidBrush(m_pBar->m_crXPSelectedColor);
				if (aBrush)
				{
					HBRUSH brushOld = SelectBrush(hDC, aBrush);
	
					Rectangle(hDC, rcPaint.left, rcPaint.top, rcPaint.right, rcPaint.bottom);
					
					SelectBrush(hDC, brushOld);
					bResult = DeleteBrush(aBrush);
					assert(bResult);
				}
				SelectPen(hDC, penOld);
				bResult = DeletePen(aPen);
				assert(bResult);
			}
		}
		else if (m_bPressed && bEnable)
		{
			HPEN aPen = CreatePen(PS_SOLID, 1, m_pBand->m_pBar->m_crXPSelectedBorderColor);
			if (aPen)
			{
				HPEN penOld = SelectPen(hDC, aPen);
				
				CRect rcPopup;
				CRect rcTool;
				if (m_pSubBandWindow && m_pSubBandWindow->IsWindow())
				{
					m_pSubBandWindow->GetWindowRect(rcPopup);
					int nToolIndex = m_pBand->m_pTools->GetToolIndex(this);
					m_pBand->GetToolScreenRect(nToolIndex, rcTool);
				}

				if (bVertical)
				{
					MoveToEx(hDC, rcPaint.left, rcPaint.top, NULL);
					LineTo(hDC, rcPaint.right, rcPaint.top);
					LineTo(hDC, rcPaint.right, rcPaint.bottom - 1);
					LineTo(hDC, rcPaint.left, rcPaint.bottom - 1);
					if (rcTool.bottom < rcPopup.top || rcTool.top > rcPopup.bottom)
					{
						MoveToEx(hDC, rcPaint.left, rcPaint.top , NULL);
						LineTo(hDC, rcPaint.left, rcPaint.bottom);
					}
				}
				else
				{
					MoveToEx(hDC, rcPaint.left, rcPaint.bottom, NULL);
					LineTo(hDC, rcPaint.left, rcPaint.top);
					LineTo(hDC, rcPaint.right, rcPaint.top);
					LineTo(hDC, rcPaint.right, rcPaint.bottom + 1);
					if (rcTool.right < rcPopup.left || rcTool.left > rcPopup.right)
					{
						MoveToEx(hDC, rcPaint.left, rcPaint.bottom , NULL);
						LineTo(hDC, rcPaint.right, rcPaint.bottom);
					}
				}
				SelectPen(hDC, penOld);
				bResult = DeletePen(aPen);
				assert(bResult);
			}
		}
	}
	BOOL bLargeIcons = m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons;
	if (m_bMoreTools)
	{
		CRect rcClipMarker = rcPaint;
		if (bVertical)
		{
			rcClipMarker.top += (rcPaint.Height() - eClipMarkerWidth - CBand::eBevelBorder2) / 2;
			rcClipMarker.left = rcClipMarker.right - eClipMarkerHeight;
		}
		else
		{
			rcClipMarker.top += CBand::eBevelBorder2;
			rcClipMarker.bottom = rcClipMarker.top + eClipMarkerHeight;
			rcClipMarker.left += (rcPaint.Width() - eClipMarkerWidth - CBand::eBevelBorder2) / 2;
		}
		DrawClipMarker(hDC, m_pBar, rcClipMarker, bVertical, bLargeIcons);
	}
	rcTemp = rcPaint;
	int nSrcWidth = 0;
	int nAdjust = 0;
	if (bVertical)
	{
		int nHeight = 5 * (bLargeIcons ? 2 : 1);
		nAdjust = (rcPaint.Height() - nHeight) / 2;
		rcTemp.Offset(1, nAdjust);
		rcTemp.bottom = rcTemp.top + nHeight;
		rcTemp.right = rcTemp.left + 3 * (bLargeIcons ? 2 : 1);
		if (m_bPressed && bEnable)
			rcTemp.Offset(1, 1);
		nSrcWidth = bEnable ? 3 : 6;
	}
	else
	{
		int nWidth = 5 * (bLargeIcons ? 2 : 1);
		nAdjust = (rcTemp.Width() - nWidth) / 2;
		rcTemp.Offset(nAdjust, -2);
		rcTemp.right = rcTemp.left + nWidth;
		rcTemp.top = rcTemp.bottom - 3 * (bLargeIcons ? 2 : 1);
		if (m_bPressed && bEnable)
			rcTemp.Offset(1, 1);
		nSrcWidth = bEnable ? 0 : 15;
	}
	DrawExpandBitmap(hDC, rcTemp, bVertical ? GetGlobals().GetBitmapExpandHorz() : GetGlobals().GetBitmapExpandVert(), nSrcWidth, bVertical ? 3 : 5, bVertical ? 5 : 3);
}

//
// DrawAllTools
//

void CTool::DrawAllTools(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical)
{
	BOOL bResult;
	MAKE_TCHARPTR_FROMWIDE(szCaption, m_bstrCaption);
	HFONT hFontText = GetMyFont(btType, FALSE);
	HFONT hFontOld = SelectFont(hDC, hFontText);
	int nModeOld = SetBkMode(hDC, TRANSPARENT);

	CRect rcTemp = rcPaint;
	if (this == m_pBar->m_pActiveTool)
	{
		if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
		{
			rcTemp.bottom += 1;
			HPEN aPen = CreatePen(PS_SOLID, 1, m_pBar->m_crXPSelectedBorderColor);
			if (aPen)
			{
				HPEN penOld = SelectPen(hDC, aPen);
				HBRUSH aBrush = CreateSolidBrush(m_pBar->m_crXPBandBackground);
				if (aBrush)
				{
					HBRUSH brushOld = SelectBrush(hDC, aBrush);
					
					Rectangle(hDC, rcTemp.left, rcTemp.top, rcTemp.right, rcTemp.bottom);

					SelectBrush(hDC, brushOld);
					bResult = DeleteBrush(aBrush);
					assert(bResult);
				}
				SelectPen(hDC, penOld);
				bResult = DeletePen(aPen);
				assert(bResult);
			}
		}
		else
		{
			rcTemp.Inflate(-1, -1);
			rcTemp = rcPaint;
			rcTemp.Inflate(-1, -1);
			m_pBar->DrawEdge(hDC, rcTemp, BDR_SUNKENOUTER, BF_RECT);
		}
	}

	rcTemp = rcPaint;
	rcTemp.Inflate(-2, -2);
	
	rcTemp.right -= eToolWithSubBandWidth;
	
	COLORREF crTextOld = SetTextColor(hDC, m_pBar->m_crForeground);

	int nTextLen = lstrlen(szCaption);
	::DrawText(hDC, szCaption, nTextLen, &rcTemp, DT_LEFT|DT_BOTTOM|DT_SINGLELINE);
	
	SelectFont(hDC, hFontOld);
	SetBkMode(hDC, nModeOld);
	SetTextColor(hDC, crTextOld);

	rcTemp = rcPaint;
	rcTemp.Inflate(-2, -2);
	rcTemp.left = rcTemp.right - eToolWithSubBandWidth;
	DrawDropDownImage(hDC, rcTemp);
}

//
// DrawChecked
//

void CTool::DrawChecked(HDC hDC, const CRect& rcPaint)
{
	HDC hMemDC = GetMemDC();
	HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMenuCheck());
	
	BitBltMasked(hDC, 
			     rcPaint.left,
			     rcPaint.top,
			     eMenuCheckWidth,
			     eMenuCheckHeight,
			     hMemDC,
			     0,
			     0,
			     eMenuCheckWidth,
			     eMenuCheckHeight,
			     m_pBar->m_crForeground);

	SelectBitmap(hMemDC, hBitmapOld);
}

//
// GetShortCutString
//

LPTSTR CTool::GetShortCutString()
{
	static TCHAR szResBuffer[512];
	TCHAR szShortCut[120];
	szShortCut[0]= NULL;
	if (m_scTool.GetCount() > 0)
		m_scTool.Get(0)->GetShortCutDesc(szShortCut);

	if (NULL == szShortCut && m_bstrCaption)
	{
		WCHAR* ptr = m_bstrCaption;
		while (*ptr)
		{
			if (L'\t' == *ptr)
			{
			    long len =  (lstrlenW(ptr+1) + 1);
				WideCharToMultiByte(CP_ACP, 
									0, 
									ptr+1, 
									-1, 
									(LPSTR)szResBuffer, 
									len, 
									NULL, 
									NULL);
				return szResBuffer;
			}
			ptr++;
		}
	}
	else
	{
		if (NULL == m_pBar)
			return NULL;

		if (NULL == szShortCut)
			return NULL;

//		MAKE_WIDEPTR_FROMTCHAR(wShortCut, szShortCut);
//		long len = lstrlenW(wShortCut) + 1;
//		WideCharToMultiByte(CP_ACP, 0, wShortCut, -1, (LPSTR)szResBuffer, len, NULL, NULL);
		lstrcpy(szResBuffer, szShortCut);
		return szResBuffer;
	}
	return NULL;
}

//
// CalcImageSize
//

void CTool::CalcImageSize()
{
	if (!m_pBar->m_pImageMgr || FAILED(m_pBar->m_pImageMgr->Size(tpV1.m_nImageIds[ddITNormal], &m_sizeImage.cx, &m_sizeImage.cy)))
	{
		m_sizeImage.cx = CBar::eDefaultBarIconWidth;
		m_sizeImage.cy = CBar::eDefaultBarIconHeight;
	}
}

void CTool::DrawCombo(HDC hDC, const CRect& rcBound, BOOL bPressed, BOOL bDrawText)
{
	if (NULL == m_pBar)
		return;

	BOOL bActive = m_pBar->m_pActiveTool == this;
	CRect rcPaint = rcBound;
	rcPaint.top += (rcPaint.bottom - rcPaint.top - m_pBar->m_nControlFontHeight - 6 + 1/*lower if not even*/) / 2;
	rcPaint.bottom = rcPaint.top + m_pBar->m_nControlFontHeight + 6;
	CBList* pList = GetComboList();
	
	if (NULL != pList)
		rcPaint.Inflate(-2, -1);
	else
		rcPaint.Inflate(-2, -2);
	
	CRect rc = rcPaint;
	if (bActive)
		rc.right -= eComboDropWidth + 3;

	CRect rcText = rc;
	
	FillRect(hDC, &rc, m_pBar->m_hBrushHighLightColor);
	rc.Inflate(-1, -1);

	FillRect(hDC, &rc, VARIANT_TRUE == tpV1.m_vbEnabled ? (HBRUSH)(1+COLOR_WINDOW) : m_pBar->m_hBrushBackColor);
	rc.Inflate(1, 1);
	
	rc = rcPaint;
	rc.left = rc.right - eComboDropWidth - 1;
	--rc.right;
	++rc.top;
	--rc.bottom;
	FillRect(hDC, &rc, m_pBar->m_hBrushBackColor);

	int nOffset = 0;
	if (bPressed)
		nOffset = 1;
	
	if (VARIANT_TRUE == tpV1.m_vbEnabled)
	{
		PaintDropDownBitmap(hDC,
							rcPaint.right - eComboDropWidth - 1 + 2 + nOffset,
							rcPaint.top + (rcPaint.Height() - 3) / 2 + nOffset,
							GetGlobals().GetBitmapSMCombo(),
							m_pBar->m_crForeground);
	}
	else
	{
		PaintDropDownBitmap(hDC, 
							1 + rcPaint.right - eComboDropWidth - 1 + 2 + nOffset,
							1 + rcPaint.top + (rcPaint.Height() - 3) / 2 + nOffset, 
							GetGlobals().GetBitmapSMCombo(),
							GetSysColor(COLOR_BTNHIGHLIGHT));
		PaintDropDownBitmap(hDC,
							rcPaint.right - eComboDropWidth - 1 + 2 + nOffset,
							rcPaint.top + (rcPaint.Height() - 3) / 2 + nOffset,
							GetGlobals().GetBitmapSMCombo(),
							GetSysColor(COLOR_BTNSHADOW));
	}

	if (bActive)
	{
		rc.Inflate(1, 1);
		DrawEdge(hDC, &rc, bPressed ? BDR_SUNKENOUTER : BDR_RAISEDINNER, BF_RECT); // using r above (buttonrect)
		rcPaint.Inflate(2, 2);
		DrawEdge(hDC, &rcPaint, BDR_SUNKENOUTER, BF_RECT);
	}
	if (bDrawText)
	{
		if (!bActive)
			rcText.right -= eComboDropWidth + 3;
		DrawTextEdit(hDC, rcText);
	}
}

void CTool::DrawXPCombo(HDC hDC, const CRect& rcBound, BOOL bPressed, BOOL bDrawText)
{
	BOOL bResult;
	if (NULL == m_pBar)
		return;

	HBRUSH hbrushBackground = CreateSolidBrush(m_pBar->m_crXPBackground);
	BOOL bActive = m_pBar->m_pActiveTool == this;
	CRect rcPaint = rcBound;
	rcPaint.top += (rcPaint.Height() - m_pBar->m_nControlFontHeight - 6 + 1/*lower if not even*/) / 2;
	rcPaint.bottom = rcPaint.top + m_pBar->m_nControlFontHeight + 6;
	CBList* pList = GetComboList();
	
	CRect rc = rcPaint;
	if (bActive)
		rc.right -= eComboDropWidth + 3;

	CRect rcText = rc;
	
	FillRect(hDC, &rc, (HBRUSH)(1+COLOR_BTNHIGHLIGHT));
	rc.Inflate(-1, -1);

	FillRect(hDC, &rc, VARIANT_TRUE == tpV1.m_vbEnabled ? (HBRUSH)(1+COLOR_WINDOW) : hbrushBackground);
	rc.Inflate(1, 1);
	
	rc = rcPaint;
	rc.left = rc.right - eComboDropWidth - 1;
	--rc.right;
	++rc.top;
	--rc.bottom;
	if (bActive)
	{
		HPEN aPen = CreatePen(PS_SOLID, 1, m_pBand->m_pBar->m_crXPSelectedBorderColor);
		if (aPen)
		{
			HPEN penOld = SelectPen(hDC, aPen);

			rc.Inflate(1, 1);
			if (m_bPressed)
				FillSolidRect(hDC, rc, m_pBar->m_crXPPressed);
			else
				FillSolidRect(hDC, rc, m_pBar->m_crXPSelectedColor);

			CRect rc2 = rcPaint;
			MoveToEx(hDC, rc2.left, rc2.top, NULL);
			LineTo(hDC, rc2.right, rc2.top);
			LineTo(hDC, rc2.right, rc2.bottom);
			LineTo(hDC, rc2.left, rc2.bottom);
			LineTo(hDC, rc2.left, rc2.top);

			rc2.left = rc2.right - (eComboDropWidth + 3);
			MoveToEx(hDC, rc2.left, rc2.top, NULL);
			LineTo(hDC, rc2.left, rc2.bottom);

			SelectPen(hDC, penOld);
			bResult = DeletePen(aPen);
			assert(bResult);
		}
	}
	else
		FillRect(hDC, &rc, hbrushBackground);

	int nOffset = 0;
	if (bPressed)
		nOffset = 1;
	
	if (VARIANT_TRUE == tpV1.m_vbEnabled)
	{
		PaintDropDownBitmap(hDC,
							rcPaint.right - eComboDropWidth - 1 + 2 + nOffset,
							rcPaint.top + (rcPaint.Height() - 3) / 2 + nOffset,
							GetGlobals().GetBitmapSMCombo(),
							m_pBar->m_crForeground);
	}
	else
	{
		PaintDropDownBitmap(hDC, 
							1 + rcPaint.right - eComboDropWidth - 1 + 2 + nOffset,
							1 + rcPaint.top + (rcPaint.Height() - 3) / 2 + nOffset, 
							GetGlobals().GetBitmapSMCombo(),
							GetSysColor(COLOR_BTNHIGHLIGHT));
		PaintDropDownBitmap(hDC,
							rcPaint.right - eComboDropWidth - 1 + 2 + nOffset,
							rcPaint.top + (rcPaint.Height() - 3) / 2 + nOffset,
							GetGlobals().GetBitmapSMCombo(),
							GetSysColor(COLOR_BTNSHADOW));
	}

	if (bDrawText)
	{
		rcText.top += 2;
		if (!bActive)
			rcText.right -= eComboDropWidth + 3;
		DrawTextEdit(hDC, rcText);
	}
	bResult = DeleteBrush(hbrushBackground);
	assert(bResult);
}

//
// DrawImage
//

void CTool::DrawImage(HDC hDC, const CRect& rcBound)
{
	if (m_dwCustomFlag)
	{
		CCustomTool custTool(m_pDispCustom);
		custTool.SetHost((ICustomHost*)this);
		custTool.OnDraw((OLE_HANDLE)hDC, 
						 rcBound.left,
						 rcBound.top,
						 rcBound.Width(),
						 rcBound.Height());
	}
	else
	{
		DrawPict((OLE_HANDLE)hDC,
				 rcBound.left,
				 rcBound.top,
				 rcBound.Width(),
				 rcBound.Height(),
				 tpV1.m_vbEnabled);
	}
}

//
// DrawText
//
void CTool::DrawText(HDC		  hDC, 
					 const CRect& rcBound, 
					 BandTypes	  btType, 
					 BOOL		  bVertical, 
					 LPCTSTR	  szCaption, 
					 BSTR	      bstrCaption, 
					 BOOL		  bImage, 
					 BOOL		  bChecked, 
					 int		  nImageAdjustment,
					 BOOL         bShortCut)
{
	assert(szCaption);
	if (NULL == szCaption)
		return;

	BOOL bXPLook = VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook;
	BOOL bEnabled = VARIANT_TRUE == tpV1.m_vbEnabled;

	BOOL bInActiveMenu = m_pBand && ddBTMenuBar == m_pBand->bpV1.m_btBands && m_pBar && !m_pBar->m_bApplicationActive;

	HFONT hFontText = GetMyFont(btType, bVertical);
	if (VARIANT_TRUE == tpV1.m_vbDefault)
	{
		LOGFONT lf;
		GetObject(hFontText, sizeof(LOGFONT), &lf);
		lf.lfWeight = FW_BOLD;
		hFontText = CreateFontIndirect(&lf);
	}
	
	//
	// Set up colors
	//

	COLORREF crTextOld;
	CRect rc(rcBound);

	switch (btType)
	{
	case ddBTPopup:
		if (bImage)
			rc.left += nImageAdjustment + 1;

		if (ddSCheckedIconText == tpV1.m_tsStyle && (bImage || VARIANT_TRUE == tpV1.m_vbChecked))
			rc.left += nImageAdjustment + 1;

		if (m_bPressed)
		{
			//
			// Pressed means selected in this case for Popups
			//

			if (!bXPLook)
				FillSolidRect(hDC, rc, GetSysColor(COLOR_HIGHLIGHT));

			if (bEnabled || Customization())
			{
				if (bXPLook)
					crTextOld = SetTextColor(hDC, m_pBar->m_crForeground);
				else
					crTextOld = SetTextColor(hDC, GetSysColor(COLOR_HIGHLIGHTTEXT));
			}
			else
				crTextOld = SetTextColor(hDC, GetSysColor(COLOR_GRAYTEXT));
		}
		else 
		{
			if (bEnabled || Customization())
				crTextOld = SetTextColor(hDC, m_pBar->m_crForeground);
			else
				crTextOld = SetTextColor(hDC, GetSysColor(COLOR_BTNHIGHLIGHT));
		}

		if (!bImage)
			rc.left += nImageAdjustment + 1;
		
		if (ddSCheckedIconText == tpV1.m_tsStyle && !(bImage || VARIANT_TRUE == tpV1.m_vbChecked))
			rc.left += nImageAdjustment + 1;

		rc.left += 2;
		break;

	default:
		if (bInActiveMenu)
		{
			if (bEnabled || Customization())
				crTextOld = SetTextColor(hDC, GetSysColor(COLOR_GRAYTEXT));
			else
			{
				crTextOld = SetTextColor(hDC, GetSysColor(COLOR_BTNHIGHLIGHT));
				rc.Offset(1, 1);
			}	
		}
		else
		{
			if (bEnabled || Customization())
			{
				if (m_pBand && m_pBand->m_pParentBand && ddCBSSlidingTabs == m_pBand->m_pParentBand->bpV1.m_cbsChildStyle)
					crTextOld = SetTextColor(hDC, m_pBand->m_pParentBand->m_pChildBands->m_crToolForeColor);
				else
					crTextOld = SetTextColor(hDC, m_pBar->m_crForeground);
			}
			else
			{
				crTextOld = SetTextColor(hDC, GetSysColor(COLOR_BTNHIGHLIGHT));
				rc.Offset(1, 1);
			}	
		}
		break;
	}

	HFONT hFontOld = SelectFont(hDC, hFontText);
	int nModeOld = SetBkMode(hDC, TRANSPARENT);
	int nTextLen;

	if (m_bstrCaption)
	{
		if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
			nTextLen = wcslen(bstrCaption);
		else
			nTextLen = lstrlen(szCaption);

		switch (btType)
		{
		case ddBTPopup:
			if (!bEnabled && !Customization())
				rc.Offset(1, 1);

#ifdef _SCRIPTSTRING
			if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
				ScriptDrawText(hDC, 
							   bstrCaption, 
							   nTextLen, 
							   &rc, 
							   (bShortCut ? DT_RIGHT : DT_LEFT)|DT_SINGLELINE|DT_VCENTER);
			else
#endif
			{
				::DrawText(hDC, 
						   szCaption, 
						   nTextLen, 
						   &rc, 
						   (bShortCut ? DT_RIGHT : DT_LEFT)|DT_SINGLELINE|DT_VCENTER);
			}

			if (!bEnabled && !Customization())
				rc.Offset(-1, -1);

			if (!bEnabled && !m_bPressed && !Customization())
			{
				SetTextColor(hDC, GetSysColor(COLOR_BTNSHADOW));
#ifdef _SCRIPTSTRING
				if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
				{
					ScriptDrawText(hDC, 
								   bstrCaption, 
								   nTextLen, 
								   &rc, 
								   (bShortCut ? DT_RIGHT : DT_LEFT)|DT_SINGLELINE|DT_VCENTER);
				}
				else
#endif
				{
					::DrawText(hDC, 
							   szCaption, 
							   nTextLen, 
							   &rc, 
							   (bShortCut ? DT_RIGHT : DT_LEFT)|DT_SINGLELINE|DT_VCENTER);
				}
			}
			break;

		default:
			if (bVertical)
			{
				rc.top += 6;
				rc.bottom -= 6;
				rc.left += 2;
				rc.right -= 2;
				if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
					DrawVerticalText(hDC, szCaption, nTextLen, rc, bstrCaption);
				else
					DrawVerticalText(hDC, szCaption, nTextLen, rc);

				if (!bEnabled && !Customization())
				{
					rc.Offset(-1, -1);
					SetTextColor(hDC, GetSysColor(COLOR_BTNSHADOW));
					if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
						DrawVerticalText(hDC, szCaption, nTextLen, rc, bstrCaption);
					else
						DrawVerticalText(hDC, szCaption, nTextLen, rc);
				}
			}
			else
			{
				CRect rcCentered;
#ifdef _SCRIPTSTRING
				if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
					ScriptDrawText(hDC, bstrCaption, nTextLen, &rcCentered, DT_LEFT|DT_CALCRECT|DT_SINGLELINE);
				else
#endif
					::DrawText(hDC, szCaption, nTextLen, &rcCentered, DT_LEFT|DT_CALCRECT);
				
				int vCenterDY = (rc.Height() - rcCentered.Height())/2;
				rc.top += vCenterDY;
				rc.left += 2;
				rc.right -= 2;
					
				int nTextAlignFlag = 0;
				switch (abs(tpV1.m_taTools%3))
				{
				case eHTALeft:
					nTextAlignFlag |= DT_LEFT;
					break;

				case eHTACenter:
					nTextAlignFlag |= DT_CENTER;
					break;

				case eHTARight:
					nTextAlignFlag |= DT_RIGHT;
					break;

				default:
					assert(FALSE);
					break;
				}

#ifdef _SCRIPTSTRING
				if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
					ScriptDrawText(hDC, bstrCaption, nTextLen, &rc, nTextAlignFlag | DT_SINGLELINE);
				else
#endif
					::DrawText(hDC, szCaption, nTextLen, &rc, nTextAlignFlag);

				if (!bEnabled && !Customization())
				{
					rc.Offset(-1, -1);
					SetTextColor(hDC, GetSysColor(COLOR_BTNSHADOW));
#ifdef _SCRIPTSTRING
					if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
						ScriptDrawText(hDC, bstrCaption, nTextLen, &rc, nTextAlignFlag | DT_SINGLELINE);
					else
#endif
						::DrawText(hDC, szCaption, nTextLen, &rc, nTextAlignFlag);
				}
			}
			break;
		}
	}
	SelectFont(hDC, hFontOld);

	if (VARIANT_TRUE == tpV1.m_vbDefault)
	{
		BOOL bResult = DeleteFont(hFontText);
		assert(bResult);
	}

	SetBkMode(hDC, nModeOld);
	SetTextColor(hDC, crTextOld);
}

//
// DrawSubBandIndicator
//

void CTool::DrawSubBandIndicator(HDC hDC, const CRect& rcPaint)
{
	HDC hMemDC = GetMemDC();
	HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapSubMenu());
	
	COLORREF crColor;
	if (VARIANT_FALSE == tpV1.m_vbEnabled)
		crColor = GetSysColor(COLOR_GRAYTEXT);
	else if (m_bPressed && VARIANT_FALSE == m_pBar->bpV1.m_vbXPLook)
		crColor = GetSysColor(COLOR_HIGHLIGHTTEXT);
	else
		crColor = m_pBar->m_crForeground;

	BitBltMasked(hDC,
			     rcPaint.right - 8,
			     ((rcPaint.bottom + rcPaint.top) / 2) - 3,
			     4,
			     7,
			     hMemDC,
			     0,
			     0,
			     4,
			     7,
			     crColor);

	SelectBitmap(hMemDC, hBitmapOld);
}

//
// DrawDropDownImage
//

void CTool::DrawDropDownImage(HDC hDC, const CRect& rcBound, BOOL bDouble)
{
	CRect rcDropDownImage = rcBound;
	rcDropDownImage.left += (rcDropDownImage.Width() - (bDouble ? 10 : 5)) / 2;
	rcDropDownImage.top += (rcDropDownImage.Height() - (bDouble ? 6 : 3)) / 2;
	if (VARIANT_TRUE == tpV1.m_vbEnabled || Customization())
	{
		if (m_bDropDownPressed && VARIANT_FALSE == m_pBar->bpV1.m_vbXPLook)
			PaintDropDownBitmap(hDC,
								rcDropDownImage.left + 1,
								rcDropDownImage.top + 1,
								GetGlobals().GetBitmapSMCombo(),
								m_pBar->m_crForeground,
								bDouble);
		else
			PaintDropDownBitmap(hDC,
								rcDropDownImage.left,
								rcDropDownImage.top,
								GetGlobals().GetBitmapSMCombo(),
								m_pBar->m_crForeground,
								bDouble);
	}
	else
	{
		PaintDropDownBitmap(hDC,
							rcDropDownImage.left + 1,
							rcDropDownImage.top + 1,
							GetGlobals().GetBitmapSMCombo(),
							GetSysColor(COLOR_BTNHIGHLIGHT),
							bDouble);
		PaintDropDownBitmap(hDC,
							rcDropDownImage.left,
							rcDropDownImage.top,
							GetGlobals().GetBitmapSMCombo(),
							GetSysColor(COLOR_BTNSHADOW),
							bDouble);
	}
}

//
// DrawEdit
//

void CTool::DrawEdit(HDC hDC, const CRect& rcBound)
{
	if (NULL == m_pBar)
		return;

	BOOL bActive = m_pBar->m_pActiveTool == this && !Customization();
	CRect rcPaint = rcBound;
	
	rcPaint.Inflate(-2,-2);
	
	FillSolidRect(hDC, rcPaint, VARIANT_FALSE == m_pBar->bpV1.m_vbXPLook ? m_pBar->m_crHighLight : m_pBar->m_crXPHighLight);

	rcPaint.Inflate(-1, -1);

	if (VARIANT_FALSE == m_pBar->bpV1.m_vbXPLook)
		FillRect(hDC, &rcPaint, VARIANT_TRUE == tpV1.m_vbEnabled ? (HBRUSH)(1+COLOR_WINDOW) : m_pBar->m_hBrushBackColor);
	else
		FillRect(hDC, &rcPaint, VARIANT_TRUE == tpV1.m_vbEnabled ? (HBRUSH)(1+COLOR_WINDOW) : (HBRUSH)(1+COLOR_BTNFACE));

	rcPaint.Inflate(1, 1);
	
	CRect rcText = rcPaint;
	if (bActive)
	{
		if (VARIANT_FALSE == m_pBar->bpV1.m_vbXPLook)
		{
			rcPaint.Inflate(2, 2);
			DrawEdge(hDC, &rcPaint, BDR_SUNKENOUTER, BF_RECT);
		}
		else
		{
			rcPaint.Inflate(1, 1);
			MoveToEx(hDC, rcPaint.left, rcPaint.top, NULL);
			LineTo(hDC, rcPaint.right, rcPaint.top);
			LineTo(hDC, rcPaint.right, rcPaint.bottom);
			LineTo(hDC, rcPaint.left, rcPaint.bottom);
			LineTo(hDC, rcPaint.left, rcPaint.top);
		}
	}
	DrawTextEdit(hDC, rcText);
}

//
// DrawTextEdit
//

void CTool::DrawTextEdit(HDC hDC, const CRect& rcBound)
{
	if (NULL == m_pBar)
		return;

	HFONT hFontOld = SelectFont(hDC, m_pBar->GetControlFont());
	if (VARIANT_TRUE == tpV1.m_vbEnabled)
		SetTextColor(hDC, GetSysColor(COLOR_WINDOWTEXT));
	else
		SetTextColor(hDC, m_pBar->m_crShadow);

	SetBkMode(hDC, TRANSPARENT);
	MAKE_TCHARPTR_FROMWIDE(szText, m_bstrText);
	if (szText)
	{
		CRect rc = rcBound;
		rc.left += 2;
		::DrawText(hDC, szText, lstrlen(szText), &rc, DT_NOPREFIX|DT_EDITCONTROL);
	}
	SelectFont(hDC, hFontOld);
}

//
// GetMyFont
//

HFONT CTool::GetMyFont(BandTypes btType, BOOL bVertical)
{
	if (NULL == m_pBar)
		return NULL;

	if (ddBTPopup == btType || ddBTChildMenuBar == btType || ddBTMenuBar == btType)
		return m_pBar->GetMenuFont(bVertical, FALSE);
	return m_pBar->GetFont();
}

//
// CalcSizeEx
//

void CTool::CalcSizeEx(HDC		 hDC, 
					   BandTypes btType, 
					   BOOL		 bVertical, 
					   SIZE&	 sizeTotal, 
					   SIZE&	 sizeChecked, 
					   SIZE&	 sizeIcon, 
					   SIZE&	 sizeText, 
					   BOOL&	 bChecked, 
					   BOOL&	 bIcon, 
					   BOOL&	 bText)
{
	try
	{
		assert(m_pBar);
		sizeChecked.cx = sizeChecked.cy = 0;
		sizeText.cx = sizeText.cy = 0;
		if (m_dwCustomFlag > 1) 
		{
			CCustomTool custTool(m_pDispCustom);
			custTool.SetHost((ICustomHost*)this);
			int w,h;
			custTool.GetSize(&w,&h);
			if (SUCCEEDED(custTool.hr))
			{
				sizeIcon.cx = w;
				sizeIcon.cy = h;
				sizeTotal.cx = w + 2 * CBand::eHToolPadding;
				sizeTotal.cy = h + 2 * CBand::eVToolPadding;
				bIcon = TRUE;
				bText = FALSE;
				bChecked = FALSE;
				return;
			}
		}

		TCHAR szCaption[512];
		CleanCaption(szCaption, ddBTPopup == btType);

		HFONT hFontMy = GetMyFont(btType, bVertical);
		if (VARIANT_TRUE == tpV1.m_vbDefault)
		{
			LOGFONT lf;
			GetObject(hFontMy, sizeof(LOGFONT), &lf);
			lf.lfWeight = FW_BOLD;
			hFontMy = CreateFontIndirect(&lf);
		}

		HFONT hFontOld = SelectFont(hDC, hFontMy);

		sizeText.cx = sizeText.cy = 0;
		m_nComboNameOffset = 0;
		ToolTypes ttTools = tpV1.m_ttTools;
		BOOL bSlidingBand = m_pBand && ddCBSSlidingTabs == m_pBand->m_pRootBand->bpV1.m_cbsChildStyle;

		if (!bSlidingBand && bVertical && (ddTTCombobox == ttTools || ddTTEdit == ttTools))
			ttTools = ddTTButton;

		switch (ttTools)
		{
		case ddTTCombobox:
		case ddTTEdit:
			switch (btType)
			{
			case ddBTPopup:
				{
					int w = tpV1.m_nWidth;
					if (-1 == w)
						w = eDefaultComboWidth;
					
					if (m_bstrCaption && *m_bstrCaption)
					{
						CRect rcOut;
#ifdef _SCRIPTSTRING
						if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
							ScriptDrawText(hDC, m_bstrCaption, wcslen(m_bstrCaption), &rcOut, DT_LEFT|DT_CALCRECT|DT_SINGLELINE);
						else
#endif
							::DrawText(hDC, szCaption, lstrlen(szCaption), &rcOut, DT_LEFT|DT_CALCRECT);
						w += sizeText.cx = rcOut.Width() + 6;
						m_nComboNameOffset = sizeText.cx + 4;
						 if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
							m_nComboNameOffset += 16;
						sizeText.cy = rcOut.Height();
					}
					sizeTotal.cx = w;
					if (m_pBar)
					{
						 if (sizeText.cy < m_pBar->m_nControlFontHeight + 6)
							sizeText.cy = m_pBar->m_nControlFontHeight + 6; // depends on font
						 if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
						 {
							 sizeTotal.cx += 16;
						 }
					}
					else
						sizeText.cy = 22;

					sizeTotal.cy = sizeText.cy;
				}
				break;

			default:
				{
					if (-1 == tpV1.m_nWidth)
						sizeTotal.cx = eDefaultComboWidth;
					else
						sizeTotal.cx = tpV1.m_nWidth;

					if (m_pBar)
						sizeTotal.cy = m_pBar->m_nControlFontHeight + 6; // depends on font
					else
						sizeTotal.cy = 22;

					switch(tpV1.m_tsStyle)
					{
					case ddSText:
					case ddSIconText:
					case ddSCheckedIconText:
						if (m_bstrCaption && *m_bstrCaption)
						{
							CRect rcOut;
#ifdef _SCRIPTSTRING
							if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
								ScriptDrawText(hDC, m_bstrCaption, wcslen(m_bstrCaption), &rcOut, DT_LEFT|DT_CALCRECT|DT_SINGLELINE);
							else
#endif
							::DrawText(hDC, szCaption, lstrlen(szCaption), &rcOut, DT_LEFT|DT_CALCRECT);
							sizeTotal.cx += sizeText.cx = rcOut.Width() + 6;
							m_nComboNameOffset = sizeText.cx + 4;
						}
						break;
					}
				}
				break;
			}
			break;

		case ddTTLabel:
			switch (tpV1.m_lsLabelStyle)
			{
			case ddLSDate:
				{	
					TCHAR szDate[60];
					GetDateFormat(LOCALE_SYSTEM_DEFAULT, DATE_SHORTDATE, &m_theSystemTime, NULL, szDate, 60);
					MAKE_WIDEPTR_FROMTCHAR(wBuffer, szDate);
					put_Caption(wBuffer);
				}
				break;

			case ddLSTime:
				{	
					TCHAR szTime[60];
					GetTimeFormat(LOCALE_SYSTEM_DEFAULT, TIME_NOSECONDS, &m_theSystemTime, NULL, szTime, 60);
					MAKE_WIDEPTR_FROMTCHAR(wBuffer, szTime);
					put_Caption(wBuffer);
				}
				break;
			}
			// Fall thru

		case ddTTButton:
		case ddTTButtonDropDown:
			{
				// Button or Button Drop Down
				bIcon = FALSE;
				bText = FALSE;
				bChecked = FALSE;

				// Now bText and bIcon specify how to paint the butto
				if (1 == m_dwCustomFlag)
				{
					CCustomTool ctTool(m_pDispCustom);
					ctTool.SetHost((ICustomHost*)this);
					int w,h;
					ctTool.GetSize(&w, &h);
					if (SUCCEEDED(ctTool.hr))
					
						sizeIcon.cx = w;
						sizeIcon.cy = h;
						bIcon = TRUE;
					
				}
				else
					sizeIcon = m_sizeImage;

				if (tpV1.m_nImageWidth > -1)
					sizeIcon.cx = tpV1.m_nImageWidth;
				
				if (tpV1.m_nImageHeight > -1)
					sizeIcon.cy = tpV1.m_nImageHeight;

				switch(tpV1.m_tsStyle)
				{
				case ddSStandard:
					switch (btType)
					{
					case ddBTNormal:
						bIcon = TRUE;
						break;

					case ddBTPopup:
						bIcon = TRUE;
						bText = TRUE;
						break;

					default:
						bText = TRUE; 
					}
					break;

				case ddSText:
					bText = TRUE;
					break;

				case ddSIcon:
					bIcon = TRUE;
					break;

				case ddSIconText:
					bText = TRUE;
					bIcon = TRUE;
					break;

				case ddSCheckedIconText:
					bText = TRUE;
					bIcon = TRUE;
					bChecked = TRUE;
					sizeChecked = m_sizeImage;
					if (0 == sizeChecked.cx || 0 == sizeChecked.cy)
						sizeChecked.cx = sizeChecked.cy = 16;
					break;
				};
				
				if (bText && m_bstrCaption && *m_bstrCaption)
				{
					CRect rcOut;
#ifdef _SCRIPTSTRING
					if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
						ScriptDrawText(hDC, m_bstrCaption, wcslen(m_bstrCaption), &rcOut, DT_LEFT|DT_CALCRECT|DT_SINGLELINE);
					else
#endif
						::DrawText(hDC, szCaption, lstrlen(szCaption), &rcOut, DT_LEFT|DT_CALCRECT);
					sizeText.cx = rcOut.Width() + 8;
					sizeText.cy = rcOut.Height();
				}
				else
				{
					sizeText.cx = 0;
					sizeText.cy = m_pBar ? m_pBar->m_nFontHeight : 0;
				}

				switch (btType)
				{
				case ddBTChildMenuBar:
				case ddBTMenuBar:
					sizeText.cy += 4;
					sizeText.cx += 4; // 2+2 menu item space
					break;

				case ddBTPopup:
					if (sizeText.cy > 16)
						sizeIcon.cx = sizeIcon.cy = sizeText.cy;
					else
						sizeIcon.cx = sizeIcon.cy = 16;
					break;

				case ddBTStatusBar:
					break;

				default:
					sizeText.cy = max(sizeText.cy, 20);
					break;
				}

				// Check if bitmap exists
				if (NULL == m_pBar->m_pImageMgr && 0 == m_dwCustomFlag)
					bIcon = FALSE;

				CaptionPositionTypes cpos;
				int nHPadding = m_pBand ? m_pBand->bpV1.m_nToolsHPadding : CBand::eHToolPadding;
				int nVPadding = m_pBand ? m_pBand->bpV1.m_nToolsVPadding : CBand::eVToolPadding;

				if (ddBTPopup == btType)
				{
					// 3 + 3 above and below space
					sizeTotal.cx = sizeTotal.cy = sizeText.cy + 6;
					// 4 pix space between text and icon
					sizeTotal.cx += (bText ? sizeIcon.cx + sizeText.cx + 2 * CBand::eHToolPadding : 0); 
					sizeTotal.cx += (bChecked ? sizeIcon.cx + 2 * CBand::eHToolPadding : 0); 
					break;
				}

				if (bIcon && 0 == m_dwCustomFlag)
				{
					sizeIcon.cx += nHPadding * 2;
					sizeIcon.cy += nVPadding * 2;
				}

				if (bChecked)
				{
					sizeChecked.cx += nHPadding * 2;
					sizeChecked.cy += nVPadding * 2;
				}

				cpos = tpV1.m_cpTools;
				if (ddCPStandard == cpos)
					cpos = ddCPRight;

				// If we are Zooming the icons
				if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons)
				{
					sizeIcon.cx *= 2;
					sizeIcon.cy *= 2;
					sizeChecked.cx *= 2;
					sizeChecked.cy *= 2;
				}

				if (bText && bIcon && bChecked)
				{
					switch(cpos)
					{
					case ddCPLeft:
					case ddCPRight:
						sizeTotal.cx = sizeText.cx + sizeIcon.cx + sizeChecked.cx + 4;
						sizeTotal.cy = max(sizeText.cy, sizeIcon.cy);
						sizeTotal.cy = max(sizeTotal.cy, sizeChecked.cy);
						break;

					case ddCPAbove:
					case ddCPBelow:
						sizeTotal.cx = max(sizeText.cx, sizeIcon.cx) + sizeChecked.cx + 4;
						sizeTotal.cy = sizeText.cy + sizeIcon.cy + 2;
						break;

					case ddCPCenter:
						sizeTotal.cx = max(sizeText.cx, sizeIcon.cx) + sizeChecked.cx + 4;
						sizeTotal.cy = max(sizeText.cy, sizeIcon.cy);
						sizeTotal.cy = max(sizeTotal.cy, sizeChecked.cy);
						break;
					}
				}
				else if (bText && bIcon)
				{
					switch(cpos)
					{
					case ddCPLeft:
					case ddCPRight:
						sizeTotal.cx = sizeText.cx + sizeIcon.cx;
						sizeTotal.cy = max(sizeText.cy, sizeIcon.cy);
						break;

					case ddCPAbove:
					case ddCPBelow:
						sizeTotal.cx = max(sizeText.cx, sizeIcon.cx);
						sizeTotal.cy = sizeText.cy + sizeIcon.cy + 2;
						break;

					case ddCPCenter:
						sizeTotal.cx = max(sizeText.cx, sizeIcon.cx);
						sizeTotal.cy = max(sizeText.cy, sizeIcon.cy);
						break;
					}
					if (-1 != tpV1.m_nWidth && sizeTotal.cx > tpV1.m_nWidth)
					{
						sizeText.cx = tpV1.m_nWidth - sizeIcon.cx;
						sizeTotal.cx = tpV1.m_nWidth + 2 * CBand::eHToolPadding;
					}
				}
				else if (bText)
				{
					sizeTotal = sizeText;
					if (-1 != tpV1.m_nWidth && sizeTotal.cx > tpV1.m_nWidth)
						sizeTotal.cx = tpV1.m_nWidth + 2 * CBand::eHToolPadding;
					if (HasSubBand() && ddBTNormal == btType)
						sizeTotal.cx += eToolWithSubBandWidth + 4;
				}
				else if (bIcon)
				{				
					if (0 == m_dwCustomFlag)
						sizeTotal = sizeIcon;
					else
					{
						sizeTotal.cx = sizeIcon.cx + 2 * CBand::eHToolPadding;
						sizeTotal.cy = sizeIcon.cy + 2 * CBand::eVToolPadding;
					}
				}
				else if (bChecked)
					sizeTotal = sizeChecked;
				else
				{
					// empty tool
					sizeTotal.cx = 16 + 2 * CBand::eHToolPadding;		
					sizeTotal.cy = 16 + 2 * CBand::eVToolPadding;
				}
			}
			break;

		case ddTTSeparator:
			switch (btType)
			{
			case ddBTPopup:
				if (m_bstrCaption && *m_bstrCaption)
				{
					sizeTotal.cy = m_pBar->m_nControlFontHeight + eMenuSeparatorThickness * 2;
					CRect rcOut;
#ifdef _SCRIPTSTRING
					if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
						ScriptDrawText(hDC, m_bstrCaption, wcslen(m_bstrCaption), &rcOut, DT_LEFT|DT_CALCRECT|DT_SINGLELINE);
					else
#endif
						::DrawText(hDC, szCaption, lstrlen(szCaption), &rcOut, DT_LEFT|DT_CALCRECT);
					sizeTotal.cx = rcOut.Width() + eSeparatorThickness * 2;
				}
				else
				{
					if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
						sizeTotal.cy = eMenuSeparatorThickness * 2;
					else
						sizeTotal.cy = eMenuSeparatorThickness * 3;
					sizeTotal.cx = eMenuSeparatorThickness;
				}
				break;

			default:
				sizeTotal.cx = eSeparatorThickness;
				sizeTotal.cy = eSeparatorThickness;
				break;
			}
			break;
		
		case ddTTMDIButtons:
			if (m_pBar)
			{
				TEXTMETRIC txMetrics;
				GetTextMetrics(hDC, &txMetrics); 

				if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				{
					m_pBar->m_sizeSMButton.cx = m_pBar->m_sizeSMButton.cy = txMetrics.tmHeight;
					sizeTotal.cy = m_pBar->m_sizeSMButton.cy + 2;
					sizeTotal.cx = (3 * m_pBar->m_sizeSMButton.cx) + 2;
				}
				else
				{
					if (m_pBar->m_sizeSMButton.cy < txMetrics.tmHeight)
						m_pBar->m_sizeSMButton.cx = m_pBar->m_sizeSMButton.cy = txMetrics.tmHeight;
					sizeTotal.cy = m_pBar->m_sizeSMButton.cy;
					sizeTotal.cx = (3 * m_pBar->m_sizeSMButton.cx) + 2;
				}
			}
			break;

		case ddTTChildSysMenu:
			{
				TEXTMETRIC txMetrics;
				GetTextMetrics(hDC, &txMetrics); 
				if (txMetrics.tmHeight < 16)
					txMetrics.tmHeight = 16;
				sizeIcon.cx = sizeIcon.cy = txMetrics.tmHeight;
				sizeTotal = sizeIcon;
				sizeTotal.cx += 2;
				sizeTotal.cy += 2;
			}
			break;

		case ddTTMenuExpandTool:
			{
				TEXTMETRIC txMetrics;
				GetTextMetrics(hDC, &txMetrics); 
				sizeTotal.cx = sizeTotal.cy = txMetrics.tmHeight;
			}
			break;

		case ddTTMoreTools:
			if (bVertical)
			{
				sizeTotal.cy = (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 20 : 11);
				sizeTotal.cx = 11;
			}
			else
			{
				sizeTotal.cx = (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 20 : 11);
				sizeTotal.cy = 20;
			}
			break;
		
		case ddTTAllTools:
			sizeTotal.cx = eButtonDropDownWidth;
			if (m_bstrCaption && *m_bstrCaption)
			{
				MAKE_TCHARPTR_FROMWIDE(szCaption, m_bstrCaption);
				CRect rcOut;
#ifdef _SCRIPTSTRING
				if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbUseUnicode)
					ScriptDrawText(hDC, m_bstrCaption, wcslen(m_bstrCaption), &rcOut, DT_LEFT|DT_CALCRECT|DT_SINGLELINE);
				else
#endif
				::DrawText(hDC, szCaption, lstrlen(szCaption), &rcOut, DT_LEFT|DT_CALCRECT);
				sizeTotal.cx += sizeText.cx = rcOut.Width() + 6;
				sizeTotal.cy = rcOut.Height() + 4;
			}
			else
				sizeTotal.cx = sizeTotal.cy = 22;
			sizeTotal.cx += 2;
			sizeTotal.cy += 2;
			break;

		case ddTTWindowList:
			sizeTotal.cx = sizeTotal.cy = 20;
			break;

		case ddTTForm:
			if (!IsWindow(m_hWndActive))
			{
				sizeTotal.cx = 16 + 2 * CBand::eHToolPadding;		
				sizeTotal.cy = 16 + 2 * CBand::eVToolPadding;
			}
			else
			{
				CRect rcPlaceHolder;
				GetWindowRect(m_hWndActive, &rcPlaceHolder);
				sizeTotal = rcPlaceHolder.Size();
			}
			break;

		case ddTTControl:
			if (!IsWindow(m_hWndActive))
			{
				sizeTotal.cx = 16 + 2 * CBand::eHToolPadding;		
				sizeTotal.cy = 16 + 2 * CBand::eVToolPadding;
			}
			else
			{
				if (-1 == tpV1.m_nHeight || -1 == tpV1.m_nWidth)
				{
					CRect rcPlaceHolder;
					GetWindowRect(m_hWndActive, &rcPlaceHolder);
					sizeTotal = rcPlaceHolder.Size();
				}

				if (-1 == tpV1.m_nHeight)
					tpV1.m_nHeight = sizeTotal.cy;
				else
					sizeTotal.cy = tpV1.m_nHeight;

				if (-1 == tpV1.m_nWidth)
					tpV1.m_nWidth = sizeTotal.cx;
				else
					sizeTotal.cx = tpV1.m_nWidth;
			}
			break;
		}
		SelectFont(hDC, hFontOld);

		if (VARIANT_TRUE == tpV1.m_vbDefault)
			DeleteFont(hFontMy);

		if (ddTTButtonDropDown == ttTools)
			sizeTotal.cx += eButtonDropDownWidth * (VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 2 : 1);
		else if (HasSubBand() && ddBTStatusBar != btType && ddBTPopup != btType && ddBTChildMenuBar != btType && ddBTMenuBar != btType && bText && !bIcon)
			sizeTotal.cx += eToolWithSubBandWidth;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CalcSize
//

SIZE CTool::CalcSize(HDC hdc, BandTypes bandType, BOOL bVertical, SIZE* psizeIcon)
{
	SIZE sizeIcon = {16, 16}, sizeChecked = {0,0}, sizeText = {0,0}, sizeTotal = {0,0};
	BOOL bIcon = FALSE; 
	BOOL bText = FALSE;
	BOOL bChecked = FALSE;
	CalcSizeEx(hdc, bandType, bVertical, sizeTotal, sizeChecked, sizeIcon, sizeText, bChecked, bIcon, bText);
	if (psizeIcon)
		*psizeIcon = sizeIcon;
	return sizeTotal;
}

STDMETHODIMP CTool::Clone( ITool **pTool)
{
	CTool* pNewTool = CTool::CreateInstance(NULL);
	*pTool = (ITool*)pNewTool;
	if (NULL == pNewTool)
		return E_OUTOFMEMORY;

	pNewTool->SetBar(m_pBar);
	pNewTool->m_bClone = TRUE;

	HRESULT hResult = CopyTo((ITool**)&pNewTool);
	if (FAILED(hResult))
	{
		pNewTool->Release();
		*pTool = NULL;
		return hResult;
	}

	if (m_pBar)
		pNewTool->tpV1.m_dwIdentity = m_pBar->bpV1.m_dwToolIdentity++;
	
	return hResult;
}

STDMETHODIMP CTool::CopyTo(ITool** pDest)
{
	try
	{
		CTool* pTool = (CTool*)(ITool*)(*pDest);

		if (pTool == this)
			return NOERROR;

		HRESULT hResult = pTool->put_Caption(m_bstrCaption);
		if (FAILED(hResult))
			return hResult;
		
		hResult = pTool->put_Name(m_bstrName);
		if (FAILED(hResult))
			return hResult;
		
		hResult = pTool->put_Description(m_bstrDescription);
		if (FAILED(hResult))
			return hResult;
		
		hResult = pTool->put_TooltipText(m_bstrToolTipText);
		if (FAILED(hResult))
			return hResult;
		
		hResult = pTool->put_Category(m_bstrCategory);
		if (FAILED(hResult))
			return hResult;

		hResult = pTool->put_SubBand(m_bstrSubBand);
		if (FAILED(hResult))
			return hResult;

		hResult = pTool->put_Text(m_bstrText);
		if (FAILED(hResult))
			return hResult;

		hResult = VariantCopy(&pTool->m_vTag, &m_vTag);
		if (FAILED(hResult))
			return hResult;

		if (pTool->m_pComboList && m_pComboList)
			m_pComboList->CopyTo((IComboList**)&pTool->m_pComboList);
		else if (pTool->m_pComboList && NULL == m_pComboList)
		{
			pTool->m_pComboList->Release();
			pTool->m_pComboList = NULL;
		}
		else if (m_pComboList)
		{
			pTool->m_pComboList = CBList::CreateInstance(NULL);
			if (pTool->m_pComboList)
			{
				pTool->m_pComboList->SetTool(pTool);
				m_pComboList->CopyTo((IComboList**)&pTool->m_pComboList);
			}
		}

		pTool->ReleaseImages();
		memcpy (&pTool->tpV1, &tpV1, sizeof(tpV1));
		pTool->AddRefImages();

		m_scTool.CopyTo(pTool->m_scTool);
		if (pTool->m_pDispCustom)
			pTool->m_pDispCustom->Release();

		pTool->m_pDispCustom = m_pDispCustom;
		if (m_pDispCustom)
			m_pDispCustom->AddRef();

		pTool->m_dwCustomFlag = m_dwCustomFlag;
		pTool->m_sizeImage = m_sizeImage;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;
}

//
// Used by designer
//

STDMETHODIMP CTool::GetSize(short nBandType, int* cx, int* cy) 
{
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		// non vertical always
		SIZE s = CalcSize(hDC, (BandTypes)nBandType, FALSE); 
		ReleaseDC(NULL, hDC);
		*cx = s.cx;
		*cy = s.cy;
		return NOERROR;
	}
	*cx = 0;
	*cy = 0;
	return E_FAIL;
}

STDMETHODIMP CTool::ExtDraw(OLE_HANDLE hdc, int x,int y,int w,int h,short bandType,VARIANT_BOOL vbSel)
{
	if (NULL == m_pBar)
		return E_FAIL;

	CRect rcPaint (x, x + w, y, y + h);
	if (ddBTPopup == bandType && (ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
	{
		// need to calc combooffset
		// forces ::Draw to recalc m_nComboNameOffset
		m_nComboNameOffset = -1; 
	}

	BOOL bPrevPressed = m_bPressed;
	if (VARIANT_TRUE == vbSel)
		m_bPressed = TRUE;
	else 
		m_bPressed = FALSE;

	BOOL bPrevFlag = Customization();
	m_pBar->m_bCustomization = FALSE;

	BSTR bstrSaveText = m_bstrText;
	m_bstrText = NULL;

	VARIANT_BOOL vbSaveEnabled = tpV1.m_vbEnabled;
	tpV1.m_vbEnabled = VARIANT_TRUE;

	// Non Vertical Draw
	Draw((HDC)hdc, rcPaint,(BandTypes)bandType, FALSE, FALSE); 

	m_bstrText = bstrSaveText;
	tpV1.m_vbEnabled = vbSaveEnabled;
	m_bPressed = bPrevPressed;

	m_pBar->m_bCustomization = bPrevFlag;
	return NOERROR;
}

STDMETHODIMP CTool::get_Category(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = SysAllocString(m_bstrCategory);
	return NOERROR;
}

STDMETHODIMP CTool::put_Category(BSTR val)
{
	SysFreeString(m_bstrCategory);
	m_bstrCategory = SysAllocString(val);
	return NOERROR;
}

STDMETHODIMP CTool::GetMaskColor(OLE_COLOR* pocMaskColor)
{
	assert(m_pBar && m_pBar->m_pImageMgr);
	return m_pBar->m_pImageMgr->get_MaskColor(tpV1.m_nImageIds[ddITNormal], pocMaskColor);
}

//
// DragDropExchange
//

STDMETHODIMP CTool::DragDropExchange(IUnknown* pStreamIn, VARIANT_BOOL vbSave)
{
	IStream* pStream = (IStream*)pStreamIn;
	try
	{
		HRESULT hResult;
		HBITMAP hBitmap;
		BOOL    bResult;
		CDib theDib;
		if (VARIANT_TRUE == vbSave)
		{
			hResult = Exchange(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			HDC hDC = GetDC(NULL);
			if (hDC)
			{
				for (int nImage = 0; nImage < eNumOfImageTypes; nImage++)
				{
					if (-1 != tpV1.m_nImageIds[nImage])
					{
						hResult = m_pBar->m_pImageMgr->GetImageBitmap(tpV1.m_nImageIds[nImage], VARIANT_TRUE, (OLE_HANDLE*)&hBitmap);
						if (FAILED(hResult))
							return hResult;
						
						bResult = theDib.FromBitmap(hBitmap, hDC);
						assert(bResult);

						bResult = DeleteBitmap(hBitmap);
						assert(bResult);
						
						hResult = theDib.Write(pStream);
						if (FAILED(hResult))
							return hResult;

						hResult = m_pBar->m_pImageMgr->GetMaskBitmap(tpV1.m_nImageIds[nImage], (OLE_HANDLE*)&hBitmap);
						if (FAILED(hResult))
							return hResult;
						
						bResult = theDib.FromBitmap(hBitmap, hDC);
						assert(bResult);
						
						bResult = DeleteBitmap(hBitmap);
						assert(bResult);

						hResult = theDib.Write(pStream);
						if (FAILED(hResult))
							return hResult;
					}
				}
				ReleaseDC(NULL, hDC);
			}
		}
		else
		{
			hResult = Exchange(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			tpV1.m_nToolId = FindLastToolId(m_pBar) + 1;

			HDC hDC = GetDC(NULL);
			if (hDC)
			{
				BITMAP bmInfo;
				for (int nImage = 0; nImage < eNumOfImageTypes; nImage++)
				{
					if (-1 != tpV1.m_nImageIds[nImage])
					{
						tpV1.m_nImageIds[nImage] = -1;
						
						hResult = theDib.Read(pStream); 
						if (FAILED(hResult))
							return hResult;

						hBitmap = theDib.CreateBitmap(hDC);
						if (NULL == hBitmap)
							return E_FAIL;

						VARIANT_BOOL vbDesignerCreated = m_pBar && m_pBar->m_pDesigner ? VARIANT_TRUE : VARIANT_FALSE;

						GetObject(hBitmap, sizeof(bmInfo), &bmInfo);
						hResult = m_pBar->m_pImageMgr->CreateImageEx(&tpV1.m_nImageIds[nImage], bmInfo.bmWidth, bmInfo.bmHeight, vbDesignerCreated); 
						if (SUCCEEDED(hResult))
						{
							hResult = m_pBar->m_pImageMgr->PutImageBitmap(tpV1.m_nImageIds[nImage], vbDesignerCreated, (OLE_HANDLE)hBitmap);
							if (FAILED(hResult))
								return hResult;
						
							bResult = DeleteBitmap(hBitmap);
							assert(bResult);

							hResult = theDib.Read(pStream);
							if (FAILED(hResult))
								return hResult;

							hBitmap = theDib.CreateBitmap(hDC);
							if (NULL == hBitmap)
								return E_FAIL;

							hResult = m_pBar->m_pImageMgr->PutMaskBitmap(tpV1.m_nImageIds[nImage],  vbDesignerCreated, (OLE_HANDLE)hBitmap);

							bResult = DeleteBitmap(hBitmap);
							assert(bResult);
							if (FAILED(hResult))
								return hResult;
						}
					}
				}
				ReleaseDC(NULL, hDC);
			}
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
// Exchange
//
//
//	Every time a new case is added to this function the case constant-expression must be added to the dwIds
// array of the DragDropExchange function.
//

HRESULT CTool::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		long nStreamSize;
		short nSize;
		short nSize2;
		BOOL bResult;
		HRESULT hResult;
		BITMAP  bmInfo;
		HBITMAP hBitmap = NULL;
		HBITMAP hMaskBitmap = NULL;
		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving
			//

			if (m_pBar && m_pBar->m_bSaveImages)
			{
				try
				{
					//
					// Saving just images
					//

					long nTempId;
					for (int nImage = 0; nImage < eNumOfImageTypes; nImage++)
					{
						if (-1 == tpV1.m_nImageIds[nImage])
							continue;

						nTempId = -1;
						//
						// Get them from the normal image manager
						//

						hResult = m_pBar->m_pImageMgr->GetImageBitmap(tpV1.m_nImageIds[nImage], VARIANT_TRUE, (OLE_HANDLE*)&hBitmap);
						if (FAILED(hResult))
							return hResult;
							
						hResult = m_pBar->m_pImageMgr->GetMaskBitmap(tpV1.m_nImageIds[nImage], (OLE_HANDLE*)&hMaskBitmap);
						if (FAILED(hResult))
						{
							DeleteBitmap(hBitmap);
							return hResult;
						}
							
						VARIANT_BOOL vbDesignerCreated = m_pBar && m_pBar->m_pDesigner ? VARIANT_TRUE : VARIANT_FALSE;

						GetObject(hBitmap, sizeof(bmInfo), &bmInfo);

						//
						// Put the images into the Band Specific Image Manager
						//

						hResult = m_pBar->m_pImageMgrBandSpecific->CreateImageEx(&nTempId, bmInfo.bmWidth, bmInfo.bmHeight, vbDesignerCreated); 
						if (SUCCEEDED(hResult))
						{
							m_pBar->m_pImageMgrBandSpecific->m_mapConvertIds.SetAt((void*)tpV1.m_nImageIds[nImage], (void*)nTempId);
							
							hResult = m_pBar->m_pImageMgrBandSpecific->PutImageBitmap(nTempId, vbDesignerCreated, (OLE_HANDLE)hBitmap);
							if (FAILED(hResult))
								goto Cleanup;

							hResult = m_pBar->m_pImageMgrBandSpecific->PutMaskBitmap(nTempId, vbDesignerCreated, (OLE_HANDLE)hMaskBitmap);
							if (FAILED(hResult))
								goto Cleanup;
						}
Cleanup:
						if (hBitmap)
						{
							bResult = DeleteBitmap(hBitmap);
							assert(bResult);
							hBitmap = NULL;
						}
						if (hMaskBitmap)
						{
							bResult = DeleteBitmap(hMaskBitmap);
							assert(bResult);
							hMaskBitmap= NULL;
						}
						if (FAILED(hResult))
							return hResult;
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
					return E_FAIL;
				}
			}
			else
			{
				try
				{
					//
					// Normal Saving of Tools
					//

					nStreamSize = GetStringSize(m_bstrToolTipText);
					nStreamSize += GetStringSize(m_bstrDescription);
					nStreamSize += GetStringSize(m_bstrCategory);
					nStreamSize += GetStringSize(m_bstrCaption);
					nStreamSize += GetStringSize(m_bstrSubBand);
					nStreamSize += GetStringSize(m_bstrText);
					nStreamSize += GetStringSize(m_bstrName);
					nStreamSize += sizeof(nSize) + sizeof(tpV1);
					nStreamSize += GetVariantSize(m_vTag);

					hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
					if (FAILED(hResult))
						return hResult;

					hResult = StWriteBSTR(pStream, m_bstrToolTipText);
					if (FAILED(hResult))
						return hResult;
					
					hResult = StWriteBSTR(pStream, m_bstrDescription);
					if (FAILED(hResult))
						return hResult;
					
					hResult = StWriteBSTR(pStream, m_bstrCategory);
					if (FAILED(hResult))
						return hResult;
					
					hResult = StWriteBSTR(pStream, m_bstrCaption);
					if (FAILED(hResult))
						return hResult;
					
					hResult = StWriteBSTR(pStream, m_bstrSubBand);
					if (FAILED(hResult))
						return hResult;
					
					hResult = StWriteBSTR(pStream, m_bstrText);
					if (FAILED(hResult))
						return hResult;
					
					hResult = StWriteBSTR(pStream, m_bstrName);
					if (FAILED(hResult))
						return hResult;

					if (m_pBar && m_pBar->m_bBandSpecificConvertIds)
					{
						//
						// This is here because of the band specific save we must convert image ids form
						// one image manager to the other
						//

						DWORD nImageIds[eNumOfImageTypes];

						// Initializes the image ids
						memcpy(&nImageIds[0], &(tpV1.m_nImageIds[0]),sizeof(DWORD)*eNumOfImageTypes);

						for (int nImage = 0; nImage < eNumOfImageTypes; nImage++)
						{
							if (-1 == tpV1.m_nImageIds[nImage])
								continue;

							m_pBar->m_pImageMgrBandSpecific->m_mapConvertIds.Lookup((void*)tpV1.m_nImageIds[nImage], (void*&)nImageIds[nImage]);
						}
						
						DWORD nImageIds2[eNumOfImageTypes];
						memcpy(&nImageIds2[0], &(tpV1.m_nImageIds[0]), sizeof(DWORD)*eNumOfImageTypes);

						//
						// Save the converted image ids
						//
						
						memcpy(&(tpV1.m_nImageIds[0]), &nImageIds[0], sizeof(DWORD)*eNumOfImageTypes);
						
						nSize = sizeof(tpV1);
						hResult = pStream->Write(&nSize, sizeof(nSize), NULL);
						if (FAILED(hResult))
							return hResult;

						hResult = pStream->Write(&tpV1, nSize, NULL);
						if (FAILED(hResult))
							return hResult;

						//
						// Restore the image ids
						//

						memcpy(&(tpV1.m_nImageIds[0]), &nImageIds2[0], sizeof(DWORD)*eNumOfImageTypes);
					}
					else
					{
						nSize = sizeof(tpV1);
						hResult = pStream->Write(&nSize, sizeof(nSize), NULL);
						if (FAILED(hResult))
							return hResult;

						hResult = pStream->Write(&tpV1, nSize, NULL);
						if (FAILED(hResult))
							return hResult;
					}
					
					hResult = StWriteVariant(pStream, m_vTag);
					if (FAILED(hResult))
						return hResult;

					hResult = m_scTool.Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pStream->Write(&m_pComboList,sizeof(m_pComboList),NULL);
					if (FAILED(hResult))
						return hResult;
					
					if (m_pComboList)
					{
						hResult = hResult = m_pComboList->Exchange(pStream, vbSave);
						if (FAILED(hResult))
							return hResult;
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
					return E_FAIL;
				}
			}
		}
		else
		{ 
			try
			{
				//
				// Loading
				//

				ULARGE_INTEGER nBeginPosition;
				LARGE_INTEGER  nOffset;

				hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
				if (FAILED(hResult))
					return hResult;

				hResult = StReadBSTR(pStream, m_bstrToolTipText);
				if (FAILED(hResult))
					return hResult;
				
				nStreamSize -= GetStringSize(m_bstrToolTipText);
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = StReadBSTR(pStream, m_bstrDescription);
				if (FAILED(hResult))
					return hResult;

				nStreamSize -= GetStringSize(m_bstrDescription);
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = StReadBSTR(pStream, m_bstrCategory);
				if (FAILED(hResult))
					return hResult;
				
				nStreamSize -= GetStringSize(m_bstrCategory);
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = StReadBSTR(pStream, m_bstrCaption);
				if (FAILED(hResult))
					return hResult;
				
				nStreamSize -= GetStringSize(m_bstrCaption);
				if (nStreamSize <= 0)
					goto FinishedReading;

				if (m_szCleanCaptionNormal)
					free(m_szCleanCaptionNormal);
				if (m_szCleanCaptionPopup)
					free(m_szCleanCaptionPopup);

				m_szCleanCaptionNormal=CleanCaptionHelper(m_bstrCaption,FALSE);
				m_szCleanCaptionPopup=CleanCaptionHelper(m_bstrCaption,TRUE);

				hResult = StReadBSTR(pStream, m_bstrSubBand);
				if (FAILED(hResult))
					return hResult;
				
				nStreamSize -= GetStringSize(m_bstrSubBand);
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = StReadBSTR(pStream, m_bstrText);
				if (FAILED(hResult))
					return hResult;
				
				nStreamSize -= GetStringSize(m_bstrText);
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = StReadBSTR(pStream, m_bstrName);
				if (FAILED(hResult))
					return hResult;
				
				nStreamSize -= GetStringSize(m_bstrName);
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = pStream->Read(&nSize, sizeof(nSize), NULL);
				if (FAILED(hResult))
					return hResult;

				nStreamSize -= sizeof(nSize);
				if (nStreamSize <= 0)
					goto FinishedReading;

				nSize2 = sizeof(tpV1);
				hResult = pStream->Read(&tpV1, nSize < nSize2 ? nSize : nSize2, NULL);
				if (FAILED(hResult))
					return hResult;
				
				if (nSize2 < nSize)
				{
					nOffset.HighPart = 0;
					nOffset.LowPart = nSize - nSize2;
					hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
					if (FAILED(hResult))
						return hResult;
				}
				nStreamSize -= nSize;
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = StReadVariant(pStream, m_vTag);
				if (FAILED(hResult))
					return hResult;

				nStreamSize -= GetVariantSize(m_vTag);
				if (nStreamSize <= 0)
					goto FinishedReading;

				nOffset.HighPart = 0;
				nOffset.LowPart = nStreamSize;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
FinishedReading:
				hResult = m_scTool.Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;

				if (m_pComboList)
					m_pComboList->Release();

				hResult = pStream->Read(&m_pComboList,sizeof(m_pComboList),NULL);
				if (FAILED(hResult))
					return hResult;

				if (m_pComboList)
				{
					m_pComboList = CBList::CreateInstance(NULL);
					if (NULL == m_pComboList)
						return E_OUTOFMEMORY;

					m_pComboList->SetTool(this);
					hResult = m_pComboList->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}

				if (ddLSTime == tpV1.m_lsLabelStyle)
					GetLocalTime(&m_theSystemTime);
				else if (ddLSDate == tpV1.m_lsLabelStyle)
					GetLocalTime(&m_theSystemTime);

				if (m_pBar && m_pBar->m_pImageMgrBandSpecific)
				{
					try
					{
						//
						// Loading Images
						//

						for (int nImage = 0; nImage < eNumOfImageTypes; nImage++)
						{
							if (-1 == tpV1.m_nImageIds[nImage])
								continue;

							hResult = m_pBar->m_pImageMgrBandSpecific->GetImageBitmap(tpV1.m_nImageIds[nImage], VARIANT_TRUE, (OLE_HANDLE*)&hBitmap);
							if (FAILED(hResult))
								return hResult;
							
							hResult = m_pBar->m_pImageMgrBandSpecific->GetMaskBitmap(tpV1.m_nImageIds[nImage], (OLE_HANDLE*)&hMaskBitmap);
							if (FAILED(hResult))
								return hResult;

							GetObject(hBitmap, sizeof(bmInfo), &bmInfo);
							tpV1.m_nImageIds[nImage] = -1;
							
							VARIANT_BOOL vbDesignerCreated = m_pBar && m_pBar->m_pDesigner ? VARIANT_TRUE : VARIANT_FALSE;

							hResult = m_pBar->m_pImageMgr->CreateImageEx(&tpV1.m_nImageIds[nImage], bmInfo.bmWidth, bmInfo.bmHeight, vbDesignerCreated); 
							if (SUCCEEDED(hResult))
							{
								hResult = m_pBar->m_pImageMgr->PutImageBitmap(tpV1.m_nImageIds[nImage], vbDesignerCreated, (OLE_HANDLE)hBitmap);
								if (FAILED(hResult))
									goto Cleanup2;
		
								hResult = m_pBar->m_pImageMgr->PutMaskBitmap(tpV1.m_nImageIds[nImage], vbDesignerCreated, (OLE_HANDLE)hMaskBitmap);
								if (FAILED(hResult))
									goto Cleanup2;
							}
Cleanup2:
							if (hBitmap)
							{
								bResult = DeleteBitmap(hBitmap);
								assert(bResult);
								hBitmap = NULL;
							}
							if (hMaskBitmap)
							{
								bResult = DeleteBitmap(hMaskBitmap);
								assert(bResult);
								hMaskBitmap= NULL;
							}
							if (FAILED(hResult))
								return hResult;
						}
					}
					CATCH
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
						return E_FAIL;
					}
				}
				if (m_pBar && m_pBar->m_pImageMgr)
					CalcImageSize();
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return E_FAIL;
			}
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
// ExchangeConfig
//

HRESULT CTool::ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		short nSize;
		short nSize2;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			hResult = StWriteBSTR(pStream, m_bstrToolTipText);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteBSTR(pStream, m_bstrDescription);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteBSTR(pStream, m_bstrCategory);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteBSTR(pStream, m_bstrCaption);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteBSTR(pStream, m_bstrSubBand);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteBSTR(pStream, m_bstrText);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteBSTR(pStream, m_bstrName);
			if (FAILED(hResult))
				return hResult;

			nSize = sizeof(tpV1);
			hResult = pStream->Write(&nSize, sizeof(nSize),NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&tpV1, nSize, NULL);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteVariant(pStream, m_vTag);
			if (FAILED(hResult))
				return hResult;

			hResult = m_scTool.Exchange(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;
			
			hResult = pStream->Write(&m_pComboList,sizeof(m_pComboList),NULL);
			if (FAILED(hResult))
				return hResult;
			
			if (m_pComboList)
			{
				hResult = hResult = m_pComboList->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			hResult = StReadBSTR(pStream, m_bstrToolTipText);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StReadBSTR(pStream, m_bstrDescription);
			if (FAILED(hResult))
				return hResult;

			hResult = StReadBSTR(pStream, m_bstrCategory);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StReadBSTR(pStream, m_bstrCaption);
			if (FAILED(hResult))
				return hResult;
			
			if (m_szCleanCaptionNormal)
				free (m_szCleanCaptionNormal);
			if (m_szCleanCaptionPopup)
				free (m_szCleanCaptionPopup);

			m_szCleanCaptionNormal = CleanCaptionHelper(m_bstrCaption,FALSE);
			m_szCleanCaptionPopup = CleanCaptionHelper(m_bstrCaption,TRUE);

			hResult = StReadBSTR(pStream, m_bstrSubBand);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StReadBSTR(pStream, m_bstrText);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StReadBSTR(pStream, m_bstrName);
			if (FAILED(hResult))
				return hResult;
			
			hResult = pStream->Read(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nSize2 = sizeof(tpV1);
			hResult = pStream->Read(&tpV1, nSize < nSize2 ? nSize : nSize2, NULL);
			if (FAILED(hResult))
				return hResult;
			
			if (nSize2 < nSize)
			{
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}

			AddRefImages();

			hResult = StReadVariant(pStream, m_vTag);
			if (FAILED(hResult))
				return hResult;

			hResult = m_scTool.Exchange(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;
			
			if (m_pComboList)
				m_pComboList->Release();

			hResult = pStream->Read(&m_pComboList,sizeof(m_pComboList),NULL);
			if (FAILED(hResult))
				return hResult;

			if (m_pComboList)
			{
				m_pComboList = CBList::CreateInstance(NULL);
				if (NULL == m_pComboList)
					return E_OUTOFMEMORY;

				m_pComboList->SetTool(this);
				hResult = m_pComboList->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}

			switch (tpV1.m_lsLabelStyle)
			{
			case ddLSTime:
				GetLocalTime(&m_theSystemTime);
				break;

			case ddLSDate:
				GetLocalTime(&m_theSystemTime);
				break;
			}

			if (m_pBar && m_pBar->m_pImageMgr)
				CalcImageSize();
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

BOOL CTool::HasSubBand()
{
	if (NULL == m_bstrSubBand || NULL == *m_bstrSubBand)
		return FALSE;
	return TRUE;
}

void CTool::AddRefImages()
{
	if (NULL == m_pBar || NULL == m_pBar->m_pImageMgr)
		return;

	for (int nIndex = 0; nIndex < eNumOfImageTypes; nIndex++)
	{
		if (-1L != tpV1.m_nImageIds[nIndex])
			m_pBar->m_pImageMgr->AddRefImage(tpV1.m_nImageIds[nIndex]);
	}
}

void CTool::ReleaseImages()
{
	try
	{
		if (NULL == m_pBar || NULL == m_pBar->m_pImageMgr || (-1 != m_nImageLoad && m_pBar->m_pImageMgr->m_nImageLoad != m_nImageLoad))
			return;


		for (int nIndex = 0; nIndex < eNumOfImageTypes; nIndex++)
		{
			if (-1L != tpV1.m_nImageIds[nIndex])
				m_pBar->m_pImageMgr->ReleaseImage(&tpV1.m_nImageIds[nIndex]);
		}
	}
	CATCH
	{
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

void CTool::OnLButtonDown()
{
	m_pBand->m_nCurrentTool = m_pBand->GetIndexOfTool(this);
	POINT pt = {-1, -1};
	OnLButtonDown(0, pt);
}

STDMETHODIMP CTool::SetFocus()
{
	if (NULL == m_pBand || NULL == m_pBar)
		return E_FAIL;

	m_pBand->m_nCurrentTool = m_pBand->GetIndexOfTool(this);
	switch (tpV1.m_ttTools)
	{
	case ddTTCombobox:
		PostMessage(m_pBar->m_hWnd, GetGlobals().WM_SETTOOLFOCUS, 0, (LPARAM)this);
		break;

	case ddTTEdit:
		PostMessage(m_pBar->m_hWnd, GetGlobals().WM_SETTOOLFOCUS, 1, (LPARAM)this);
		break;

	case ddTTControl:
	case ddTTForm:
		if (IsWindow(m_hWndActive))
			::SetFocus(m_hWndActive);
		break;
	}
	return NOERROR;
}

//
// DrawSlidingEdge
//

void CTool::DrawSlidingEdge(HDC hDC, const CRect& rcEdge, BOOL bSunken)
{
	BOOL bResult;
	CRect rc = rcEdge;
	if (bSunken)
	{
		rc.bottom = rc.top + 1;
		bResult = FillSolidRect(hDC, rc, m_pBand->m_pParentBand->m_pChildBands->m_crShadowColor);
		
		rc = rcEdge;
		rc.right = rc.left + 1;
		bResult = FillSolidRect(hDC, rc, m_pBand->m_pParentBand->m_pChildBands->m_crShadowColor);
		
		rc = rcEdge;
		rc.top = rc.bottom - 1;
		bResult = FillSolidRect(hDC, rc, m_pBand->m_pParentBand->m_pChildBands->m_crHighLightColor);

		rc = rcEdge;
		rc.left = rc.right - 1;
		bResult = FillSolidRect(hDC, rc, m_pBand->m_pParentBand->m_pChildBands->m_crHighLightColor);
	}
	else
	{
		rc.bottom = rc.top + 1;
		bResult = FillSolidRect(hDC, rc, m_pBand->m_pParentBand->m_pChildBands->m_crHighLightColor);

		rc = rcEdge;
		rc.right = rc.left + 1;
		bResult = FillSolidRect(hDC, rc, m_pBand->m_pParentBand->m_pChildBands->m_crHighLightColor);

		rc = rcEdge;
		rc.top = rc.bottom - 1;
		bResult = FillSolidRect(hDC, rc, m_pBand->m_pParentBand->m_pChildBands->m_crShadowColor);

		rc = rcEdge;
		rc.left = rc.right - 1;
		bResult = FillSolidRect(hDC, rc, m_pBand->m_pParentBand->m_pChildBands->m_crShadowColor);
	}
}

//
// CheckForAlt
//

BOOL CTool::CheckForAlt(WCHAR& nKey)
{
	try
	{
		if (NULL == m_bstrCaption || NULL == *m_bstrCaption)
			return FALSE;

		WCHAR* pChar = m_bstrCaption;
		while (*pChar)
		{
			if ('&' == *pChar && '&' != *(pChar+1))
			{
				nKey = *(pChar+1);
				break;
			}
			pChar++;
		}
		if (NULL == *pChar)
			return FALSE;
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// OnLButtonUp
//

BOOL CTool::OnLButtonUp(UINT nFlags, POINT pt)
{
	try
	{
		ToolTypes ttTools = tpV1.m_ttTools;
		if ((ddDALeft == m_pBand->bpV1.m_daDockingArea || ddDARight == m_pBand->bpV1.m_daDockingArea) && 
			(ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
		{
			ttTools = ddTTButton;
		}

		switch (ttTools)
		{
		case ddTTAllTools:
		case ddTTButton:
			if (ddBTNormal == m_pBand->bpV1.m_btBands && ddDAPopup == m_pBand->bpV1.m_daDockingArea)
			{
				try
				{
					m_pBar->FireToolClick(reinterpret_cast<Tool*>(this));
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
			else if (m_bPressed)
			{
				try
				{
					m_bPressed = FALSE;
					if (VARIANT_FALSE == tpV1.m_vbAutoRepeat)
					{
						if (::IsWindow(m_pBar->m_hWnd))
						{
							AddRef();
							::PostMessage(m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)this);
							if (m_pBar->m_bIsVisualFoxPro)
							{
								AddRef();
								::PostMessage(m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)this);
							}
						}
						else
							m_pBar->FireToolClick(reinterpret_cast<Tool*>(this));
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
			break;

		case ddTTButtonDropDown:
			if (m_bPressed)
			{
				m_bPressed = FALSE;
				if (::IsWindow(m_pBar->m_hWnd))
				{
					AddRef();
					::PostMessage(m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)this);
					if (m_pBar->m_bIsVisualFoxPro)
					{
						AddRef();
						::PostMessage(m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)this);
					}
				}
				else
					m_pBar->FireToolClick(reinterpret_cast<Tool*>(this));
			}
			m_bDropDownPressed = FALSE;
			break;

		case ddTTLabel:
			if (m_bPressed)
			{
				m_bPressed = FALSE;
				m_pBar->FireToolClick(reinterpret_cast<Tool*>(this));
			}
			break;

		case ddTTMoreTools:
			m_bDropDownPressed = FALSE;
			m_bPressed = FALSE;
			break;
		}
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
// OnMenuLButtonUp
//

void CTool::OnMenuLButtonUp(BOOL bClicked, BOOL bDelayed)
{
	try
	{
		ToolTypes ttTools = tpV1.m_ttTools;
		if ((ddDALeft == m_pBand->bpV1.m_daDockingArea || ddDARight == m_pBand->bpV1.m_daDockingArea) && 
			(ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
		{
			ttTools = ddTTButton;
		}

		if (m_bPressed && bClicked)
		{
			if (bDelayed)
			{
				AddRef();
				PostMessage(m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)this);
			}
			else
				m_pBar->FireToolClick((Tool*)this);

			if (!IsWindow(m_pBar->m_hWnd))
				return;
		}
		m_bPressed = FALSE;
		m_bDropDownPressed = FALSE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// TrackToolLButtonDown
//

void CTool::TrackToolLButtonDown()
{
	try
	{
		AddRef();
		// Standard Tracking Loop
		m_pBar->SetActiveTool(NULL);

		CRect rcTool;
		HWND  hWndDock;
		BOOL bResult = m_pBand->GetToolBandRect(m_pBand->m_nCurrentTool, rcTool, hWndDock);
		assert(bResult);
		if (!bResult)
			return;

		CRect rcToolScreenRelative = rcTool;
		ClientToScreen(hWndDock, rcToolScreenRelative);
		rcToolScreenRelative.Inflate(CBand::eBevelBorder, CBand::eBevelBorder);

		UINT nTimer;
		if (VARIANT_TRUE == tpV1.m_vbAutoRepeat)
		{
			m_pBar->FireToolClick(reinterpret_cast<Tool*>(this));
			nTimer = SetTimer(hWndDock, eAutoRepeat, tpV1.m_nAutoRepeatInterval, NULL);
		}

		m_bFocus = TRUE;
		BOOL bProcessing = TRUE;
		MSG msg;
		SetCapture(hWndDock);
		POINT pt;
		goto MouseMoveEntryPoint;
		
		while (GetCapture() == hWndDock && bProcessing)
		{
			while (::PeekMessage(&msg, NULL, WM_PAINT, WM_PAINT, PM_NOREMOVE))
			{
				if (!GetMessage(&msg, NULL, WM_PAINT, WM_PAINT))
					break;

				DispatchMessage(&msg);
			}

			GetMessage(&msg, NULL, 0, 0);

			switch (msg.message)
			{
			// handle movement/accept messages
			case WM_CANCELMODE:
			case WM_LBUTTONUP:
				bProcessing = FALSE;
				break;

			case WM_TIMER:
				try
				{
					switch (msg.wParam)
					{
					case eAutoRepeat:
						switch (tpV1.m_ttTools)
						{
						case ddTTButton:
							if (m_bPressed)
							{
								try
								{
									m_pBar->FireToolClick(reinterpret_cast<Tool*>(this));
								}
								CATCH
								{
									assert(FALSE);
									REPORTEXCEPTION(__FILE__, __LINE__)
								}
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
				break;

			case WM_MOUSEMOVE:
			{
MouseMoveEntryPoint:
				GetCursorPos(&pt);
				if (PtInRect(&rcToolScreenRelative, pt))
				{
					if (!m_bPressed)
						m_bPressed = TRUE;
					m_pBand->InvalidateToolRect(this, rcTool);
				}
				else if (m_bPressed)
				{
					m_bPressed = FALSE;
					m_pBand->InvalidateToolRect(this, rcTool);
				}
			}
			break;

			default:
				DispatchMessage(&msg);
				break;
			}
		}
		if (hWndDock == GetCapture())
			ReleaseCapture();
		if (VARIANT_TRUE == tpV1.m_vbAutoRepeat)
			KillTimer(hWndDock, eAutoRepeat);
		m_bFocus = FALSE;
		if (m_bPressed)
			OnLButtonUp(msg.wParam, pt);
		m_pBand->InvalidateToolRect(this, rcTool);
		Release();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnMenuLButtonDown
//

void CTool::OnMenuLButtonDown(UINT nFlags, POINT pt)
{
	try
	{
		ToolTypes ttTools = tpV1.m_ttTools;
		if ((ddDALeft == m_pBand->bpV1.m_daDockingArea || ddDARight == m_pBand->bpV1.m_daDockingArea) && 
			(ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
		{
			ttTools = ddTTButton;
		}

		m_pBar->HideToolTips(TRUE);
		switch (ttTools)
		{
		case ddTTMoreTools:
		case ddTTAllTools:
		case ddTTButton:
			m_bPressed = TRUE;
			break;

		case ddTTButtonDropDown:
			{
				pt.x -= m_pBand->m_sizeEdgeOffset.cx;
				pt.y -= m_pBand->m_sizeEdgeOffset.cy;
				CRect rcTool = m_pBand->m_prcTools[m_pBand->m_nCurrentTool];
				if (pt.x > (rcTool.right - (eButtonDropDownWidth * (VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 2 : 1)) - 4))
					m_bDropDownPressed = TRUE;
				else
				{
					m_bPressed = TRUE;
					TrackToolLButtonDown();
				}
			}
			break;

		case ddTTLabel:
			m_bPressed = TRUE;
			break;

		case ddTTEdit:
			m_bPressed = TRUE;
			break;

		case ddTTCombobox:
			m_bPressed = TRUE;
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
// OnLButtonDown
//

void CTool::OnLButtonDown(UINT nFlags, POINT pt)
{
	try
	{
		ToolTypes ttTools = tpV1.m_ttTools;
		if ((ddDALeft == m_pBand->bpV1.m_daDockingArea || ddDARight == m_pBand->bpV1.m_daDockingArea) && 
			(ddTTCombobox == tpV1.m_ttTools || ddTTEdit == tpV1.m_ttTools))
		{
			ttTools = ddTTButton;
		}

		m_pBar->HideToolTips(TRUE);
		switch (ttTools)
		{
		case ddTTAllTools:
		case ddTTButton:
			m_bPressed = TRUE;
			if (HasSubBand())
				m_bDropDownPressed = TRUE;
			else
				TrackToolLButtonDown();
			break;

		case ddTTButtonDropDown:
			{
				pt.x -= m_pBand->m_sizeEdgeOffset.cx;
				pt.y -= m_pBand->m_sizeEdgeOffset.cy;
				CRect rcTool = m_pBand->m_prcTools[m_pBand->m_nCurrentTool];
				if (pt.x > (rcTool.right - (eButtonDropDownWidth * (VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 2 : 1)) - 4))
				{
					m_bDropDownPressed = TRUE;
				}
				else
				{
					m_bPressed = TRUE;
					TrackToolLButtonDown();
				}
			}
			break;

		case ddTTLabel:
			m_bPressed = TRUE;
			TrackToolLButtonDown();
			break;

		case ddTTEdit:
			{
				switch(tpV1.m_tsStyle)
				{				
				case ddSText:
				case ddSIconText:
				case ddSCheckedIconText:
					if (m_bstrCaption && *m_bstrCaption)
					{
						POINT pt2 = pt;
						pt2.x -= m_pBand->m_sizeEdgeOffset.cx;
						pt2.y -= m_pBand->m_sizeEdgeOffset.cy;
						CRect rcTool = m_pBand->m_prcTools[m_pBand->m_nCurrentTool];
						rcTool.right = rcTool.left + m_nComboNameOffset;
						if (PtInRect(&rcTool, pt2))
							return;
					}
					break;
				}
				HandleEdit(nFlags, pt);
			}
			break;

		case ddTTCombobox:
			{
				switch(tpV1.m_tsStyle)
				{	
				case ddSText:
				case ddSIconText:
				case ddSCheckedIconText:
					if (m_bstrCaption && *m_bstrCaption)
					{
						POINT pt2 = pt;
						pt2.x -= m_pBand->m_sizeEdgeOffset.cx;
						pt2.y -= m_pBand->m_sizeEdgeOffset.cy;
						CRect rcTool = m_pBand->m_prcTools[m_pBand->m_nCurrentTool];
						rcTool.right = rcTool.left + m_nComboNameOffset;
						if (PtInRect(&rcTool, pt2))
							return;
					}
					break;
				}
				HandleCombobox(nFlags, pt);
			}
			break;

		case ddTTMDIButtons:
			{
				BOOL bVertical = ddDALeft == m_pBand->bpV1.m_daDockingArea || ddDARight == m_pBand->bpV1.m_daDockingArea;
				CRect rcTool;
				HWND hWndBand;
				if (m_pBand->GetToolBandRect(m_pBand->m_nCurrentTool, rcTool, hWndBand))
				{
					if (ddDAFloat != m_pBand->bpV1.m_daDockingArea)
					{
						pt.x += m_pBand->m_rcDock.left;
						pt.y += m_pBand->m_rcDock.top;
					}
					HDC hDC = GetDC(hWndBand);
					if (hDC)
					{
						int nButton = CBar::eCloseWindowPosition;
						if (GetMDIButton(nButton, rcTool, pt, bVertical))
							TrackMDIButton(hDC, hWndBand, rcTool, nButton);
						ReleaseDC(hWndBand, hDC);
					}
				}
			}
			break;

		case ddTTChildSysMenu:
			{
				m_pBar->DoToolModal(TRUE);
				CRect rcTool;
				if (m_pBand->GetToolScreenRect(m_pBand->m_nCurrentTool, rcTool))
				{
					SIZE size;
					size.cx = rcTool.left;
					size.cy = rcTool.bottom;
					PixelToTwips(&size, &size);
					POINT ptScreen;
					ptScreen.x = size.cx;
					ptScreen.y = size.cy;
					CRect rc;
					m_pBar->OnMDIChildPopupMenu(ptScreen, rc);				
				}
				m_pBar->DoToolModal(FALSE);
			}
			break;

		case ddTTMoreTools:
			m_bDropDownPressed = TRUE;
			m_bPressed = TRUE;
			if (HasSubBand() && EnableMoreTools())
				GenerateMoreTools();
			break;
		}
	}
	CATCH
	{
		m_bPressed = FALSE;
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// GenerateMoreTools
//

void CTool::GenerateMoreTools()
{
	CTool* pTempTool;
	CTool* pTool;

	DDString strTemp;
	HRESULT hResult;
	int nCount = 0;
	if (ddBFCustomize & m_pBand->bpV1.m_dwFlags)
	{
		//
		// All Tools
		//

		if (m_pBar->m_pBands->Find(m_pBar->m_pMoreTools->m_bstrName))
			return;
		nCount = m_pBand->m_pTools->GetToolCount();
		hResult = m_pBar->m_pBands->InsertBand(-1, m_pBar->m_pAllTools);
		if (SUCCEEDED(hResult))
		{
			for (int nTool = 0; nTool < nCount; nTool++)
			{
				pTool = m_pBand->m_pTools->GetTool(nTool);

				if (ddTTMoreTools == pTool->tpV1.m_ttTools || ddTTSeparator == pTool->tpV1.m_ttTools)
					continue;

				hResult = m_pBar->m_pAllTools->m_pTools->Add(CBar::eToolIdToolToggle, pTool->m_bstrName, (Tool**)&pTempTool);
				if (SUCCEEDED(hResult))
				{
					hResult = pTool->CopyTo((ITool**)&pTempTool);
					if (SUCCEEDED(hResult))
					{
						pTempTool->tpV1.m_nUsage = 0;
						pTempTool->m_vTag.vt = VT_I4;
						pTempTool->m_vTag.pdispVal = (ITool*)pTool;
						pTempTool->tpV1.m_nToolId = CBar::eToolIdToolToggle;
						pTempTool->tpV1.m_ttTools = ddTTButton;
						pTempTool->put_SubBand(L"");
						pTempTool->tpV1.m_vbChecked = pTool->tpV1.m_vbVisible;
						pTempTool->tpV1.m_vbVisible = VARIANT_TRUE;
						pTempTool->tpV1.m_vbEnabled = VARIANT_TRUE;
						pTempTool->tpV1.m_tsStyle = ddSCheckedIconText;
						pTempTool->tpV1.m_mvMenuVisibility = ddMVAlwaysVisible;
					}
					pTempTool->Release();
				}
			}
			if (ddBFCustomize & m_pBand->bpV1.m_dwFlags)
			{
				hResult = m_pBar->m_pAllTools->m_pTools->Add(1, L"miSeparator", (Tool**)&pTempTool);
				if (SUCCEEDED(hResult))
				{
					pTempTool->tpV1.m_ttTools = ddTTSeparator;
					pTempTool->Release();
				}

				hResult = m_pBar->m_pAllTools->m_pTools->Add(2, L"miResetToolbar", (Tool**)&pTempTool);
				if (SUCCEEDED(hResult))
				{
					pTempTool->tpV1.m_nToolId = CBar::eToolIdResetToolbar;
					pTempTool->tpV1.m_nUsage = 0;
					LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTResetButton);
					if (szLocaleString)
						strTemp = szLocaleString;
					else
						strTemp.LoadString(IDS_RESETTOOLBAR);
					MAKE_WIDEPTR_FROMTCHAR(wCust, strTemp);
					hResult = pTempTool->put_Caption(wCust);
					pTempTool->tpV1.m_tsStyle = ddSCheckedIconText;
					pTempTool->m_vTag.vt = VT_BSTR;
					pTempTool->m_vTag.bstrVal = SysAllocString(m_pBand->m_bstrName);
					pTempTool->Release();
				}

				hResult = m_pBar->m_pAllTools->m_pTools->Add(3, L"miCustomize", (Tool**)&pTempTool);
				if (SUCCEEDED(hResult))
				{
					pTempTool->tpV1.m_nToolId = CBar::eToolIdCustomize;
					pTempTool->tpV1.m_nUsage = 0;
					LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTMenuCustomize);
					if (szLocaleString)
					{
						MAKE_WIDEPTR_FROMTCHAR(wCust, szLocaleString);
						hResult = pTempTool->put_Caption(wCust);
					}
					pTempTool->tpV1.m_tsStyle = ddSCheckedIconText;
					pTempTool->Release();
				}
			}
		}
	}

	//
	// More Tools
	//

	hResult = m_pBar->m_pBands->InsertBand(-1, m_pBar->m_pMoreTools);
	if (SUCCEEDED(hResult))
	{
		m_pBar->m_pMoreTools->bpV1.m_daDockingArea = ddDAPopup;
		if (m_bMoreTools)
		{
			nCount = m_pBand->m_pTools->GetVisibleToolCount();
			for (int nTool = 0; nTool < nCount; nTool++)
			{
				pTool = m_pBand->m_pTools->GetVisibleTool(nTool);

				if (pTool->IsVisibleOnPaint()  || ddTTSeparator == pTool->tpV1.m_ttTools || VARIANT_FALSE == pTool->tpV1.m_vbVisible)
					continue;

				hResult = m_pBar->m_pMoreTools->m_pTools->Add(pTool->tpV1.m_nToolId, pTool->m_bstrName, (Tool**)&pTempTool);
				if (SUCCEEDED(hResult))
				{
					hResult = pTool->CopyTo((ITool**)&pTempTool);
					if (SUCCEEDED(hResult))
					{
						pTempTool->m_hWndActive = pTool->m_hWndActive;
						
						switch (pTempTool->tpV1.m_ttTools)
						{
						case ddTTControl:
						case ddTTForm:
							pTempTool->tpV1.m_ttTools = ddTTButton;
							break;
						}
						pTempTool->tpV1.m_mvMenuVisibility = ddMVAlwaysVisible;
					}
					pTempTool->Release();
				}
			}
		}
		if (ddBFCustomize & m_pBand->bpV1.m_dwFlags)
		{
			hResult = m_pBar->m_pMoreTools->m_pTools->Add(1, L"miSeparator", (Tool**)&pTempTool);
			if (SUCCEEDED(hResult))
			{
				pTempTool->tpV1.m_ttTools = ddTTSeparator;
				pTempTool->Release();
			}

			strTemp.LoadString(IDS_ALLTOOLSNAME);
			MAKE_WIDEPTR_FROMTCHAR(wName, strTemp);
			hResult = m_pBar->m_pMoreTools->m_pTools->Add(2, wName, (Tool**)&pTempTool);
			if (SUCCEEDED(hResult))
			{
				LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTAddOrRemoveButton);
				MAKE_WIDEPTR_FROMTCHAR(wCaption, szLocaleString);
				hResult = pTempTool->put_Caption(wCaption);

				strTemp.LoadString(IDS_POPUPALLTOOLS);
				MAKE_WIDEPTR_FROMTCHAR(wSubBand, strTemp);
				hResult = pTempTool->put_SubBand(wSubBand);
				
				pTempTool->tpV1.m_ttTools = (ToolTypes)ddTTAllTools;
				pTempTool->Release();
			}
		}
	}
}

//
// GetMDIButton
//

BOOL CTool::GetMDIButton(int& nButton, CRect& rcButton, POINT pt, BOOL bVertical)
{
	nButton = -1;
	if (bVertical)
	{
		rcButton.Offset(-2, 0);
		rcButton.top = rcButton.bottom - m_pBar->m_sizeSMButton.cy;

		if (PtInRect(&rcButton, pt))
		{
			nButton = CBar::eCloseWindowPosition;
			return TRUE;
		}
		else
		{
			rcButton.Offset(0, -(m_pBar->m_sizeSMButton.cy + 2));
			if (PtInRect(&rcButton, pt))
			{
				nButton = CBar::eRestoreWindowPosition;
				return TRUE;
			}
			else
			{
				rcButton.Offset(0, -m_pBar->m_sizeSMButton.cy);
				if (PtInRect(&rcButton, pt))
				{
					nButton = CBar::eMinimizeWindowPosition;
					return TRUE;
				}
			}
		}
	}
	else
	{
		rcButton.top += 2;
		rcButton.left = rcButton.right - m_pBar->m_sizeSMButton.cx;
		rcButton.bottom = rcButton.top + m_pBar->m_sizeSMButton.cy;

		if (PtInRect(&rcButton, pt))
		{
			nButton = CBar::eCloseWindowPosition;
			return TRUE;
		}
		else
		{
			rcButton.Offset(-(m_pBar->m_sizeSMButton.cx+2), 0);
			if (PtInRect(&rcButton, pt))
			{
				nButton = CBar::eRestoreWindowPosition;
				return TRUE;
			}
			else
			{
				rcButton.Offset(-m_pBar->m_sizeSMButton.cx, 0);
				if (PtInRect(&rcButton, pt))
				{
					nButton = CBar::eMinimizeWindowPosition;
					return TRUE;
				}
			}
		}
	}
	return FALSE;
}

//
// TrackMDIButton
//

BOOL CTool::TrackMDIButton(HDC hDC, HWND hWndBand, const CRect& rcButton, int nButton)
{
	BOOL bResult;
	int nBitmapOffset = 0;
	UINT ndfcsCode;
	switch (nButton)
	{
	case CBar::eMinimizeWindowPosition:
		ndfcsCode = DFCS_CAPTIONMIN;
		break;
	case CBar::eRestoreWindowPosition:
		ndfcsCode = DFCS_CAPTIONRESTORE;
		nBitmapOffset = 8;
		break;
	case CBar::eCloseWindowPosition:
		if (!m_bMDICloseEnabled)
			return FALSE;
		ndfcsCode = DFCS_CAPTIONCLOSE;
		nBitmapOffset = 2 * 8;
		break;
	}

	BOOL bState = FALSE; // pressed
	BOOL bNewState;
	MSG msg;
	msg.message = WM_MOUSEMOVE;
	POINT pt;
	BOOL bProcessing = TRUE;
	SetCapture(hWndBand);
	
	m_pBar->m_bToolModal = true;

	goto StartLoop;

	while (GetCapture() == hWndBand && bProcessing)
	{
		GetMessage(&msg, NULL, 0, 0);

		switch (msg.message)
		{
		case WM_LBUTTONUP:
			Refresh();
			bProcessing = FALSE;
			GetCursorPos(&pt);
			ScreenToClient(hWndBand, &pt);
			if (PtInRect(&rcButton, pt))
			{
				HWND hWndActive = (HWND)SendMessage(m_pBar->m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
				if (hWndActive)
				{
					switch (nButton)
					{
					case CBar::eMinimizeWindowPosition:
						ShowWindow(hWndActive, SW_MINIMIZE);
						break;

					case CBar::eRestoreWindowPosition:
						ShowWindow(hWndActive, SW_RESTORE);
						break;

					case CBar::eCloseWindowPosition:
						if (m_bMDICloseEnabled)
							PostMessage(hWndActive, WM_SYSCOMMAND, SC_CLOSE, 0);
						break;
					}
				}
			}
			break;

		case WM_MOUSEMOVE:
		{
StartLoop:
			GetCursorPos(&pt);
			ScreenToClient(hWndBand, &pt);
			if (PtInRect(&rcButton, pt))
				bNewState = TRUE;
			else
				bNewState = FALSE;

			if (bNewState != bState)
			{
				bState = bNewState;
				if (VARIANT_FALSE == m_pBar->bpV1.m_vbXPLook)
				{
					DrawFrameControl(hDC,
									 (LPRECT)&rcButton,
									 DFC_CAPTION,
									 ndfcsCode | (bState ? DFCS_PUSHED : 0));
				}
				else
				{
					HBRUSH brushBackground = NULL;
					HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
					if (penBorder)
					{
						HPEN penOld = SelectPen(hDC, penBorder);

						if (bState) // IsPushed
							brushBackground = CreateSolidBrush(m_pBar->m_crXPPressed);
						else
							brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);

						if (brushBackground)
						{
							HBRUSH brushOld = SelectBrush(hDC, brushBackground);
							
							Rectangle(hDC, rcButton.left, rcButton.top, rcButton.right, rcButton.bottom);

							SelectBrush(hDC, brushOld);
							bResult = DeleteBrush(brushBackground);
							assert(bResult);
						}
						SelectPen(hDC, penBorder);
						bResult = DeletePen(penBorder);
						assert(bResult);
					}

					HDC hMemDC = GetMemDC();
					HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMDIButtons());

					BitBltMasked(hDC, 
								 rcButton.left + (rcButton.Width() - 8) / 2,
								 rcButton.top + (rcButton.Height() - 9) / 2,
								 8,
								 9,
								 hMemDC,
								 nBitmapOffset,
								 0,
								 8,
								 9,
								 m_pBar->m_crForeground);

					SelectBitmap(hMemDC, hBitmapOld);
				}
				m_pBar->HideToolTips(TRUE);
			}
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		break;
		
		default:
			// just dispatch rest of the messages
			TranslateMessage(&msg);
			DispatchMessage(&msg);
			break;
		}
	}
	ReleaseCapture();

	if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
	{
		FillSolidRect(hDC, rcButton, m_pBar->m_crXPBackground);
		HDC hMemDC = GetMemDC();
		HBITMAP hBitmapOld = SelectBitmap(hMemDC, GetGlobals().GetBitmapMDIButtons());

		BitBltMasked(hDC, 
					 rcButton.left + (rcButton.Width() - 8) / 2,
					 rcButton.top + (rcButton.Height() - 9) / 2,
					 8,
					 9,
					 hMemDC,
					 nBitmapOffset,
					 0,
					 8,
					 9,
					 m_pBar->m_crForeground);

		SelectBitmap(hMemDC, hBitmapOld);
		m_mdibActive = CTool::eNone;
	}
	m_pBar->m_bToolModal = false;
	return bState;
}

//
// HideHWnd
//

void CTool::HidehWnd()
{
	if (NULL == m_pBand || !IsWindow(m_hWndActive))
		return;

	ShowWindow(m_hWndActive, SW_HIDE);
}

//
// SetParent
//

void CTool::SetParent(HWND hWndParent)
{
	if (!IsWindow(m_hWndActive) || hWndParent == m_hWndParent)
		return;

	if (IsWindow(hWndParent))
	{
		m_hWndParent = hWndParent;
		if (NULL == m_hWndPrevParent)
		{
			ShowWindow(m_hWndActive, SW_HIDE);
			m_hWndPrevParent = ::SetParent(m_hWndActive, m_hWndParent);
			assert(m_hWndPrevParent);
		}
		else
		{
			ShowWindow(m_hWndActive, SW_HIDE);
			HWND hWndTmp = ::SetParent(m_hWndActive, m_hWndParent);
			assert(hWndTmp);
		}
	}
	else
	{
		if (IsWindow(m_hWndActive) && IsWindow(m_hWndPrevParent))
		{
			ShowWindow(m_hWndActive, SW_HIDE);
			::SetParent(m_hWndActive, m_hWndPrevParent);
			m_hWndParent = m_hWndPrevParent;
		}
	}
}

//
// HandleEdit
//

BOOL CTool::HandleEdit(int nFlags, POINT& pt)
{
	HWND hWndCapture = GetCapture();
	try
	{
		BOOL bHorz = ddDATop == m_pBand->bpV1.m_daDockingArea || ddDABottom == m_pBand->bpV1.m_daDockingArea || ddDAFloat == m_pBand->bpV1.m_daDockingArea || ddDAPopup == m_pBand->bpV1.m_daDockingArea;
		if (!Customization() && bHorz && VARIANT_TRUE == tpV1.m_vbEnabled)
		{
			m_pActiveEdit = m_pBar->GetActiveEdit();
			assert(m_pActiveEdit);
			if (NULL == m_pActiveEdit)
				return FALSE;
			m_pActiveEdit->SetTool(this);

			HWND hWnd;
			CRect rcTool;
			if (m_pBand->GetToolBandRect(m_pBand->m_nCurrentTool, rcTool, hWnd))
			{
				::SetParent(m_pActiveEdit->hWnd(), hWnd);

				rcTool.left += m_nComboNameOffset;

				rcTool.Offset(1, -1);
				
				if (ddBTPopup == m_pBand->bpV1.m_btBands && m_pBand->m_pPopupWin)
					m_pBand->m_pPopupWin->AdjustForScrolling(rcTool);

				m_hWndActive = m_pActiveEdit->hWnd();
				m_bPressed = TRUE;

				MAKE_TCHARPTR_FROMWIDE(szText, m_bstrText);
				m_pActiveEdit->SendMessage(WM_SETTEXT, 0, (LPARAM)szText);

				m_pActiveEdit->PostMessage(EM_SETSEL, tpV1.m_nSelStart, tpV1.m_nSelLength);
				m_pBar->m_pTabbedTool = this;


				if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
					rcTool.Inflate(-3, -3);
				m_pActiveEdit->SetWindowPos(NULL, 
											rcTool.left, 
											rcTool.top, 
											rcTool.Width(), 
											rcTool.Height(), 
											SWP_SHOWWINDOW | SWP_NOACTIVATE);
				
				if (ddDAFloat != m_pBand->bpV1.m_daDockingArea)
					rcTool.Offset(-m_pBand->m_rcDock.left, -m_pBand->m_rcDock.top);
				m_pActiveEdit->SetFocus();
				try
				{
					m_pBar->DoToolModal(TRUE);
					m_pBar->FireToolGotFocus((Tool*)this);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}

				BOOL bProcessing = TRUE;
				MSG msg;
				while (bProcessing)
				{
					if (NULL == m_pActiveEdit)
					{
						bProcessing = FALSE;
						continue;
					}

					if (!(WS_VISIBLE & GetWindowLong(m_pActiveEdit->hWnd(), GWL_STYLE)))
					{
						bProcessing = FALSE;
						continue;
					}

					try
					{
						if (PeekMessage(&msg, NULL, WM_SYSCOMMAND, WM_SYSCOMMAND, PM_NOREMOVE))
						{
							if ((msg.wParam & 0xFFF0) != SC_CLOSE || (GetKeyState(VK_F4) < 0 && GetKeyState(VK_MENU) < 0))
								bProcessing = FALSE;
							continue;
						}
					}
					catch (...)
					{
						assert(FALSE);
					}

					try
					{
						if (PeekMessage(&msg, NULL, WM_CLOSE, WM_CLOSE, PM_NOREMOVE))
						{
							bProcessing = FALSE;
							continue;
						}
					}
					catch (...)
					{
						assert(FALSE);
					}

					GetMessage(&msg, NULL, 0, 0);

					switch (msg.message)
					{
					case WM_CANCELMODE:
						bProcessing = FALSE;
						break;

					case WM_CLOSE:
						bProcessing = FALSE;
						break;

					case WM_KEYDOWN:
						try
						{
							if (NULL == m_pActiveEdit)
								break;

							switch (msg.wParam)
							{
							case VK_TAB:
								if (m_pActiveEdit->IsWindow())
								{
									m_pActiveEdit->ShowWindow(SW_HIDE);
									::SetParent(m_pActiveEdit->hWnd(), m_pBar->m_hWnd);
								}
								if (m_pBand->m_pPopupWin)
									m_pBand->m_pPopupWin->PostMessage(WM_KEYDOWN, msg.wParam, msg.lParam);
								bProcessing = FALSE;
								continue;

							case VK_ESCAPE:
								m_pActiveEdit->ShowWindow(SW_HIDE);
								::SetParent(m_pActiveEdit->hWnd(), m_pBar->m_hWnd);
								bProcessing = FALSE;
								break;
							}
							if (m_pActiveEdit)
							{
								int nLen = m_pActiveEdit->SendMessage(WM_GETTEXTLENGTH);
								if (VK_RETURN == msg.wParam && nLen >= 0)
								{
									try
									{
										m_pActiveEdit->Exchange();
										m_pBar->FireToolKeyDown((Tool*)this, (short*)&msg.wParam, OCXShiftState());
										m_pBar->FireToolKeyPress((Tool*)this, (long*)&msg.wParam);
										m_pBar->FireToolKeyUp((Tool*)this, (short*)&msg.wParam, OCXShiftState());
										if (m_pActiveEdit && m_pActiveEdit->IsWindow())
										{
											m_pActiveEdit->ShowWindow(SW_HIDE);
											::SetParent(m_pActiveEdit->hWnd(), m_pBar->m_hWnd);
										}
										bProcessing = FALSE;
										continue;
									}
									catch (...)
									{
										if (m_pActiveEdit && m_pActiveEdit->IsWindow())
										{
											m_pActiveEdit->ShowWindow(SW_HIDE);
											::SetParent(m_pActiveEdit->hWnd(), m_pBar->m_hWnd);
										}
										bProcessing = FALSE;
										assert(FALSE);
									}
								}
							}
							TranslateMessage(&msg);
							DispatchMessage(&msg);
						}
						catch (...)
						{
							assert(FALSE);
						}
						break;

					case WM_KEYUP:
						try
						{
							if (NULL == m_pActiveEdit)
								break;

							if (VK_TAB == msg.wParam)
								continue;

							if (VK_RETURN == msg.wParam)
							{
								m_pActiveEdit->ShowWindow(SW_HIDE);
								::SetParent(m_pActiveEdit->hWnd(), m_pBar->m_hWnd);
								bProcessing = FALSE;
							}
							else
							{
								TranslateMessage(&msg);
								DispatchMessage(&msg);
							}
						}
						catch (...)
						{
							assert(FALSE);
						}
						break;

					case WM_NCLBUTTONDOWN:
					case WM_LBUTTONDOWN:
						try
						{
							if (NULL == m_pActiveEdit)
								break;

							GetCursorPos(&pt);
							
							m_pActiveEdit->GetWindowRect(rcTool);
							if (!PtInRect(&rcTool, pt))
							{
								bProcessing = FALSE;
								m_pActiveEdit->ShowWindow(SW_HIDE);
								::SetParent(m_pActiveEdit->hWnd(), m_pBar->m_hWnd);
								m_pBar->DoToolModal(FALSE);
							}
							TranslateMessage(&msg);
							DispatchMessage(&msg);
						}
						catch (...)
						{
							assert(FALSE);
						}
						break;

					case WM_MOUSEMOVE:
						try
						{
							if (NULL == m_pActiveEdit)
								break;

							GetCursorPos(&pt);
							m_pActiveEdit->GetWindowRect(rcTool);
							if (PtInRect(&rcTool, pt))
								DispatchMessage(&msg);
						}
						catch (...)
						{
							assert(FALSE);
						}
						break;

					default:
						try
						{
							TranslateMessage(&msg);
							DispatchMessage(&msg);
						}
						catch (...)
						{
							assert(FALSE);
						}
						break;
					}
				}
				try
				{
					m_pBar->DoToolModal(FALSE);
					m_pBar->FireToolLostFocus((Tool*)this);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
		}
		try
		{
			m_pBar->m_hWndActiveCombo = NULL;

			if (::IsWindow(hWndCapture) && hWndCapture != GetCapture())
				SetCapture(hWndCapture);
		}
		catch (...)
		{
			assert(FALSE);
		}
		m_hWndActive = NULL;
	}
	CATCH
	{
		if (hWndCapture != GetCapture())
			SetCapture(hWndCapture);
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return TRUE;
}

//
// IsComboReadOnly
//

BOOL CTool::IsComboReadOnly()
{
	CBList* pList = GetComboList();
	if (NULL == pList)
		return TRUE;

	if (ddCBSReadOnly == pList->lpV1.m_nStyle || ddCBSSortedReadOnly == pList->lpV1.m_nStyle)
		return TRUE;
	return FALSE;
}

//
// HandleCombobox
//

BOOL CTool::HandleCombobox(int nFlags, POINT& pt)
{
	HWND hWndFocus = ::GetFocus();
	try
	{
		HWND hWndCapture = GetCapture();

		BOOL bHorz = ddDATop == m_pBand->bpV1.m_daDockingArea || ddDABottom == m_pBand->bpV1.m_daDockingArea || ddDAFloat == m_pBand->bpV1.m_daDockingArea || ddDAPopup == m_pBand->bpV1.m_daDockingArea;
		
		if (!Customization() && bHorz && VARIANT_TRUE == tpV1.m_vbEnabled)
		{
			DWORD dwStyle = WS_CHILDWINDOW | CBS_AUTOHSCROLL | WS_VSCROLL;
			CBList* pList = GetComboList();
			if (pList)
			{
				switch (pList->lpV1.m_nStyle)
				{
				case ddCBSReadOnly:
					dwStyle |= CBS_DROPDOWNLIST;
					break;

				case ddCBSSorted: 
					dwStyle |= CBS_DROPDOWN|CBS_SORT;
					break;

				case ddCBSSortedReadOnly:
					dwStyle |= CBS_SORT|CBS_DROPDOWNLIST;
					break;

				default:
					dwStyle |= CBS_DROPDOWN;
					break;
				}
			}
			m_pActiveCombo = m_pBar->GetActiveCombo(dwStyle);
			assert(m_pActiveCombo);
			if (NULL == m_pActiveCombo)
				return FALSE;
			m_pActiveCombo->SetTool(this);

			BSTR bstrTextOld = NULL;
			HRESULT hResult = get_Text(&bstrTextOld);
			MAKE_TCHARPTR_FROMWIDE(szText, bstrTextOld);

			CRect rcTool;
			HWND hWnd;
			if (m_pBand->GetToolBandRect(m_pBand->m_nCurrentTool, rcTool, hWnd))
			{
				rcTool.left += m_nComboNameOffset;
				rcTool.Offset(0, -1);

				if (ddBTPopup == m_pBand->bpV1.m_btBands && m_pBand->m_pPopupWin)
					m_pBand->m_pPopupWin->AdjustForScrolling(rcTool);

				::SetParent(m_pActiveCombo->hWnd(), hWnd);
				m_hWndActive = m_pBar->m_hWndActiveCombo = m_pActiveCombo->hWnd();

				if (pList)
				{
					//
					// Size the drop down height
					//

					int nItemHeight = m_pActiveCombo->SendMessage(CB_GETITEMHEIGHT);
					CRect rcCombobox;
					m_pActiveCombo->GetWindowRect(rcCombobox);
					
					int nComboHeight;
					if (pList->lpV1.m_nLines > 0)
					{
						nComboHeight = pList->lpV1.m_nLines * nItemHeight + rcCombobox.Height() + 2;

						::ScreenToClient(hWnd, rcCombobox);
						rcCombobox.bottom = rcCombobox.top + nComboHeight;
						rcCombobox.right = rcCombobox.left + 10;
						m_pActiveCombo->MoveWindow(rcCombobox, FALSE);
					}
					
					//
					// Set the width of the Drop Down
					//

					if (pList->lpV1.m_nWidth != 0)
						m_pActiveCombo->SendMessage(CB_SETDROPPEDWIDTH, pList->lpV1.m_nWidth, 0);

					//
					// Fill in Combo Contents
					//


					m_pActiveCombo->SendMessage(CB_RESETCONTENT);
					pList->m_hWndActive = m_pActiveCombo->hWnd();
					BSTR bstrItem;
					int nCount = pList->Count();
					for (int nItem = 0; nItem < nCount; nItem++)
					{
						bstrItem = pList->GetName(nItem);
						if (bstrItem)
						{
							MAKE_TCHARPTR_FROMWIDE(szItem, bstrItem);
							m_pActiveCombo->SendMessage(CB_ADDSTRING, 0, (LPARAM)szItem);
						}
					}
					switch (pList->lpV1.m_nStyle)
					{
					case ddCBSReadOnly:
					case ddCBSSortedReadOnly:
						{
							int nIndex = m_pActiveCombo->SendMessage(CB_FINDSTRING, (WPARAM)-1, (LPARAM)szText);
							if (CB_ERR != nIndex)
								m_pActiveCombo->SendMessage(CB_SETCURSEL, nIndex, 0);
						}
						break;

					default:
						m_pActiveCombo->SendMessage(CB_SETCURSEL, pList->lpV1.m_nListIndex);
						break;
					}
				}

				m_pActiveCombo->SendMessage(WM_SETTEXT, 0, (LPARAM)szText);


				m_pActiveCombo->PostMessage(CB_SETEDITSEL, tpV1.m_nSelStart, tpV1.m_nSelLength);
				m_pBar->m_pTabbedTool = this;
				m_pActiveCombo->SetWindowPos(NULL, rcTool.left, rcTool.top, rcTool.Width(), rcTool.Height(), SWP_SHOWWINDOW | SWP_NOACTIVATE);

				if (ddDAFloat != m_pBand->bpV1.m_daDockingArea)
					rcTool.Offset(-m_pBand->m_rcDock.left, -m_pBand->m_rcDock.top);
				if (pt.x != -1 && pt.y != -1)
				{
					pt.x -= rcTool.left;
					pt.y -= rcTool.top;
					m_pActiveCombo->PostMessage(WM_LBUTTONDOWN, nFlags, MAKELONG(pt.x, pt.y));
				}
				else
					m_pActiveCombo->SetFocus();

				try
				{
					m_pBar->DoToolModal(TRUE);
					m_pBar->FireToolGotFocus((Tool*)this);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				BOOL bDispatchLastMessage = FALSE;
				BOOL bProcessing = TRUE;
				MSG msg;
				while (bProcessing)
				{
					if (NULL == m_pActiveCombo)
					{
						bProcessing = FALSE;
						continue;
					}

					if (PeekMessage(&msg, NULL, WM_SYSCOMMAND, WM_SYSCOMMAND, PM_NOREMOVE))
					{
						if ((msg.wParam & 0xFFF0) != SC_CLOSE || (GetKeyState(VK_F4) < 0 && GetKeyState(VK_MENU) < 0))
							bProcessing = FALSE;
						continue;
					}

					if (!(WS_VISIBLE & GetWindowLong(m_pActiveCombo->hWnd(), GWL_STYLE)))
					{
						bProcessing = FALSE;
						continue;
					}


					GetMessage(&msg, NULL, 0, 0);

					switch (msg.message)
					{
					case WM_CLOSE:
						bProcessing = FALSE;
						DispatchMessage(&msg);
						break;

					case WM_CANCELMODE:
						bProcessing = FALSE;
						break;

					case WM_KEYDOWN:
						try
						{
							m_pBar->FireToolKeyDown((Tool*)this, (short*)&msg.wParam, OCXShiftState());
							TranslateMessage(&msg);
							DispatchMessage(&msg);
						}
						CATCH
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
						break;

					case WM_CHAR:
						try
						{
							m_pBar->FireToolKeyPress((Tool*)this, (long*)&msg.wParam);
							DispatchMessage(&msg);
						}
						CATCH
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
						break;

					case WM_KEYUP:
						try
						{
							m_pBar->FireToolKeyUp((Tool*)this, (short*)&msg.wParam, OCXShiftState());
						}
						CATCH
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}
						switch (msg.wParam)
						{
						case VK_ESCAPE:
							if (m_pActiveCombo->IsWindow())
							{
								m_pActiveCombo->ShowWindow(SW_HIDE);
								::SetParent(m_pActiveCombo->hWnd(), m_pBar->m_hWnd);
							}
							put_Text(bstrTextOld);
							bProcessing = FALSE;
							break;

						case VK_TAB:
							if (m_pActiveCombo->IsWindow())
							{
								m_pActiveCombo->ShowWindow(SW_HIDE);
								::SetParent(m_pActiveCombo->hWnd(), m_pBar->m_hWnd);
							}
							put_Text(bstrTextOld);
							if (m_pBand->m_pPopupWin)
								m_pBand->m_pPopupWin->PostMessage(WM_KEYDOWN, msg.wParam, msg.lParam);
							bProcessing = FALSE;
							break;
						}
						TranslateMessage(&msg);
						DispatchMessage(&msg);
						if (VK_RETURN == msg.wParam)
						{
							m_pActiveCombo->ShowWindow(SW_HIDE);
							::SetParent(m_pActiveCombo->hWnd(), m_pBar->m_hWnd);
							bProcessing = FALSE;
						}
						break;

					case WM_NCLBUTTONDOWN:
					case WM_LBUTTONDOWN:
						{
							if (m_pActiveCombo->IsWindow())
							{
								CRect rcTool;
								m_pActiveCombo->GetWindowRect(rcTool);

								POINT pt;
								GetCursorPos(&pt);
								CRect rcDropDown;
								BOOL bDropped = m_pActiveCombo->SendMessage(CB_GETDROPPEDSTATE);
								if (!PtInRect(&rcTool, pt) && !bDropped)
								{
									m_pActiveCombo->ShowWindow(SW_HIDE);
									::SetParent(m_pActiveCombo->hWnd(), m_pBar->m_hWnd);
									m_pBar->DoToolModal(FALSE);
									DispatchMessage(&msg);
									bProcessing = FALSE;
								}
								else if (bDropped)
								{
									rcTool.left = rcTool.right - 10;
									if (PtInRect(&rcTool, pt))
									{
										DispatchMessage(&msg);
										m_pActiveCombo->ShowWindow(SW_HIDE);
										::SetParent(m_pActiveCombo->hWnd(), m_pBar->m_hWnd);
										bProcessing = FALSE;
										m_pBar->DoToolModal(FALSE);
									}
									else
									{
										m_pActiveCombo->SendMessage(CB_GETDROPPEDCONTROLRECT, 0, (LPARAM)&rcDropDown);
										CRect rcCombo;
										m_pActiveCombo->GetWindowRect(rcCombo);
										int nCount = m_pActiveCombo->SendMessage(CB_GETCOUNT);
										int nItemHeight = m_pActiveCombo->SendMessage(CB_GETITEMHEIGHT);
										rcCombo.bottom += (nCount * nItemHeight);
										if (rcDropDown.bottom > rcCombo.bottom)
											rcDropDown = rcCombo;
										int nBottom = GetSystemMetrics(SM_CYVIRTUALSCREEN);
										CRect rcScreen;
										GetWindowRect(GetDesktopWindow(), &rcScreen);
										if (rcDropDown.bottom > nBottom)
											rcDropDown.Offset(0, -rcDropDown.Height());
										if (!PtInRect(&rcDropDown, pt) && !PtInRect(&rcTool, pt))
										{
											DispatchMessage(&msg);
											m_pActiveCombo->ShowWindow(SW_HIDE);
											::SetParent(m_pActiveCombo->hWnd(), m_pBar->m_hWnd);
											bProcessing = FALSE;
											m_pBar->DoToolModal(FALSE);
										}
									}
								}
								if (bProcessing)
									DispatchMessage(&msg);
							}
							else
								bProcessing = FALSE;
						}
						break;

					default:
						TranslateMessage(&msg);
						DispatchMessage(&msg);
						break;
					}
				}
				SysFreeString(bstrTextOld);
				m_pActiveCombo = NULL;
				m_pBar->m_hWndActiveCombo = NULL;
				try
				{
					m_pBar->DoToolModal(FALSE);
					m_pBar->FireToolLostFocus((Tool*)this);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
		}
		switch (m_pBar->m_eAppType)
		{
		case CBar::eMDIForm:
			{
//				HWND hWndActive = (HWND)SendMessage(m_pBar->m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
//				if (IsWindow(hWndActive))
//					::SetFocus(hWndActive);
			}
			break;

		case CBar::eSDIForm:
		case CBar::eClientArea:
			::SetFocus(m_pBar->m_hWnd);
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	m_hWndActive = NULL;
	::SetFocus(hWndFocus);
	return TRUE;
}

//
// Custom Tool Hosting Service
//

STDMETHODIMP CTool::get_hWin(OLE_HANDLE *retval)
{
	return E_FAIL;
}

STDMETHODIMP CTool::put_hWin(OLE_HANDLE val)
{
	return E_FAIL;
}

STDMETHODIMP CTool::get_State(short *retval)
{
	return E_FAIL;
}

STDMETHODIMP CTool::put_State(short val)
{
	return E_FAIL;
}

STDMETHODIMP CTool::Refresh()
{
	return NOERROR;
}

STDMETHODIMP CTool::Close()
{
	return NOERROR;
}

