//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <olectl.h>               
#include <stddef.h>       // for offsetof()
#include "Debug.h"
#include "Errors.h"
#include "Support.h"
#include "Flicker.h"
#include <LIMITS.H>
#include <MultiMon.h>
#include "Utility.h"
#include "Globals.h"
#include "IpServer.h"
#include "DropSource.h"
#include "Designer\DragDrop.h"
#include "Resource.h"
#include "DispIds.h"
#include "Dock.h"
#include "MiniWin.h"
#include "PopupWin.h"
#include "Tool.h"
#include "ChildBands.h"
#include "Bands.h"
#include "Localizer.h"
#include "Band.h"
#include "prj_i.c"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define DC_GRADIENT 0x0020
#define COLOR_GRADIENTACTIVECAPTION     27
#define COLOR_GRADIENTINACTIVECAPTION   28
#define SPI_GETGRADIENTCAPTIONS         0x1008

//
// DrawBtnFaceBorder 
//

static void DrawBtnFaceBorder(HDC hDC, const CRect& rcPaint, COLORREF crColor)
{
	// mask
	COLORREF crColorOld = SetBkColor(hDC, crColor);

	RECT rc = rcPaint;
	rc.bottom = rc.top + 1;
	FillSolidRect(hDC, rc, crColor);
	
	rc = rcPaint;
	rc.top = rc.bottom - 1;
	FillSolidRect(hDC, rc, crColor);
	
	rc = rcPaint;
	rc.right = rc.left + 1;
	FillSolidRect(hDC, rc, crColor);
	
	rc = rcPaint;
	rc.left = rc.right - 1;
	FillSolidRect(hDC, rc, crColor);

	SetBkColor(hDC, crColorOld);
}

#ifdef _DEBUG
//#define _DEBUGPAINTCOUNT
#endif

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CBand(IUnknown *pUnkOuter)
{
	return (IUnknown *)(IBand *)new CBand();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_Band,Band,_T("Band"),CreateFN_CBand,2,0,&IID_IBand,_T(""));
void *CBand::objectDef=&BandObject;
CBand *CBand::CreateInstance(IUnknown *pUnkOuter)
{
	return new CBand();
}
//{OBJECT CREATEFN}
CBand::CBand()
	: m_refCount(1)
	 ,m_pChildBands(NULL),
	  m_pnCustomControlIndexes(NULL),
	  m_pParentPopupWin(NULL),
	  m_pChildPopupWin(NULL),
	  m_pPopupWin(NULL),
	  m_bstrCaption(NULL),
	  m_prcTools(NULL),
	  m_prcPopupExpandArea(NULL),
	  m_pDockMgr(NULL),
	  m_bstrName(NULL),
	  m_pFloat(NULL),
	  m_pDock(NULL),
	  m_pToolMenuExpand(NULL),
	  m_pParentBand(NULL),
	  m_pRootBand(NULL),
	  m_pBar(NULL)
{
	InterlockedIncrement(&g_cLocks);

	m_daPrevDockingArea = ddDATop;
	m_nPrevDockOffset = 0;
	m_nPrevDockLine = 10;
	m_nCurrentTool = -1;
	m_dwLastDocked = 0;
	m_nFirstTool = 0;

	m_bMiniWinShowLock = FALSE;
	m_bChildBandChanging = FALSE;
	m_bPopupWinLock = FALSE;
	m_bDrawAnimated = FALSE;
	m_bVerticalMenu = FALSE;
	m_bDesignTime = FALSE;
	m_bCycleMark = FALSE;
	m_bLoading = FALSE;
	m_bToolRemoved = FALSE;
	m_vbCachedVisible = -2;
	m_sizeDragOffset.cx = m_sizeDragOffset.cy = 0;
	m_ptPercentOffset.x = m_ptPercentOffset.y = 0;
	m_nDockLineIndex = 0;
// {BEGIN INIT}
// {END INIT}
	m_pTools = CTools::CreateInstance(NULL);
	assert(m_pTools);
	if (m_pTools)
		m_pTools->SetBand(this);

	m_nPopupShortCutStringOffset = 0;
	m_nAdjustForGrabBar = 0;
	m_nTempDockOffset = -1;
	m_nToolMouseOver = -1;
	m_nPopupIndex = -1;
	m_nInBandOpen = 0;
	m_nDropLoc = -1;
	m_nCachedDockOffset = -1;
	VariantInit(&m_vTag);
	m_eExpandButtonState = eGrayed;
	m_sizePage.cx = m_sizePage.cy = 0;
	m_sizeBandEdge.cx = m_sizeBandEdge.cy = 0;
	m_sizeEdgeOffset.cx = m_sizeEdgeOffset.cy = 0;
	m_nFormControlCount = 0;
	m_nHostControlCount = 0;
	m_bExpandButtonVisible = FALSE;
	m_bCloseButtonVisible = FALSE; 
	m_bDontExpand = FALSE;
	m_bAllTools = FALSE;
	if (m_pBar && m_pBar->ActiveBand() == this)
		m_pBar->ActiveBand(NULL);
	OleTranslateColor(bpV1.m_ocPictureBackground, NULL, &m_crPictureBackground);
	TRACE1(1, "Created Band : %X\n",(DWORD)this);
	m_nLastPopupExpandedArea = -1;
	m_bSizerResized = FALSE;
	m_sizeMaxIcon.cx = m_sizeMaxIcon.cy = 16;
}

CBand::~CBand()
{
	TRACE1(1, "Deleting Band : %X\n",(DWORD)this);
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
// {END CLEANUP}
	if (m_pFloat)
	{
		m_pFloat->BandClosing(TRUE);
		HWND hWnd = m_pFloat->hWnd();
		if (IsWindow(hWnd))
		{

			try
			{
				::DestroyWindow(hWnd);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
		else
			delete m_pFloat;
	}
	SysFreeString(m_bstrCaption);
	SysFreeString(m_bstrName);
	if (m_pTools)
		m_pTools->Release();
	if (m_pChildBands)
		m_pChildBands->Release();
	if (m_pToolMenuExpand)
		m_pToolMenuExpand->Release();
	m_pRootBand = NULL;
	m_pBar = NULL;
	delete [] m_prcTools;
	m_prcTools = NULL;
	delete [] m_prcPopupExpandArea;
	m_prcPopupExpandArea = NULL;
	delete [] m_pnCustomControlIndexes;
	VariantClear(&m_vTag);
	m_bStatusBarSizerPresent = FALSE;
	TRACE(1, _T("CBand::~CBand()\n"));
}

#ifdef _DEBUG
void CBand::Dump(DumpContext& dc)
{
	MAKE_TCHARPTR_FROMWIDE(szCaption, m_bstrCaption);
	MAKE_TCHARPTR_FROMWIDE(szName, m_bstrName);
	_stprintf(dc.m_szBuffer,
			  _T("Band Name: %s Caption: %s\n"), 
			  szName,
			  szCaption);
	dc.Write();
	if (m_pChildBands)
		m_pChildBands->Dump(dc);
	if (m_pTools)
		m_pTools->Dump(dc);
}
void CBand::DumpLayoutInfo(DumpContext& dc)
{
	MAKE_TCHARPTR_FROMWIDE(szName, m_bstrName);
	_stprintf(dc.m_szBuffer,
			  _T("Band Name: %s, Dock Line: %i, Dock Line Index %i\n"), 
			  szName,
			  bpV1.m_nDockLine,
			  m_nDockLineIndex);
	dc.Write();
}
#endif

CBand::BandPropV1::BandPropV1()
{
	m_ghsGrabHandleStyle = ddGSNormal;
	m_tsMouseTracking = ddTSBevel;
	m_daDockingArea = ddDATop;
	m_cbsChildStyle = ddCBSNone;
	m_btBands = ddBTNormal;
	m_dwFlags = ddBFDockLeft|ddBFDockTop|ddBFDockRight|ddBFDockBottom|ddBFFloat|ddBFHide|ddBFCustomize;

	m_nDockOffset = 0;
	m_nDockLine = 0;
	m_nToolsHPadding = 2;
	m_nToolsVPadding = 2;
	m_nToolsHSpacing = 0;
	m_nToolsVSpacing = 0;
	m_nCreatedBy = ddCBApplication;

	m_vbWrapTools = VARIANT_FALSE;
	m_vbVisible = VARIANT_TRUE;
	m_vbDesignerCreated = VARIANT_FALSE;
	m_vbDetached = VARIANT_FALSE;

	m_rcDimension.Set(-1, -1, -1, -1);
	
	m_nDockedHorzMinWidth = 50;
	m_rcHorzDockForm.Set(0, 0, m_nDockedHorzMinWidth, m_nDockedHorzMinWidth);
	m_nDockedVertMinWidth = 50;
	m_rcVertDockForm.Set(0, 0, m_nDockedVertMinWidth, m_nDockedVertMinWidth);

	m_vbDisplayMoreToolsButton = VARIANT_TRUE;
	m_pbsPicture = ddPBBSNormal;
	m_ocPictureBackground = 0x80000000 + COLOR_BTNFACE;
	m_nScrollingSpeed = 100;
}

void CBand::ShutdownFloatWin()
{
	try
	{
		if (NULL == m_pFloat)
			return;

		try
		{
			if (m_pFloat->BandClosing())
				return;
			m_pFloat->BandClosing(TRUE);
			if (VARIANT_TRUE == bpV1.m_vbVisible)
				put_Visible(VARIANT_FALSE);
			HWND hWnd = m_pFloat->hWnd();
			if (IsWindow(hWnd))
			{
				try
				{
					::DestroyWindow(hWnd);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
			else
			{
				delete m_pFloat;
				m_pFloat = NULL;
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CBand::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IUnknown))
	{
		AddRef();
		*ppvObjOut=this;
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IDispatch))
	{
		*ppvObjOut=(void *)(IDispatch *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IBand))
	{
		*ppvObjOut=(void *)(IBand *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISupportErrorInfo))
	{
		*ppvObjOut=(void *)(ISupportErrorInfo *)this;
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
STDMETHODIMP_(ULONG) CBand::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CBand::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
STDMETHODIMP CBand::get_ScrollingSpeed(short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddBTPopup != bpV1.m_btBands)
		return E_FAIL;

	*retval = bpV1.m_nScrollingSpeed;
	return NOERROR;
}
STDMETHODIMP CBand::put_ScrollingSpeed(short val)
{
	if (ddBTPopup != bpV1.m_btBands)
		return E_FAIL;

	bpV1.m_nScrollingSpeed = val;
	return NOERROR;
}
STDMETHODIMP CBand::get_DockedVertMinWidth(short *retval)
{		
	if (NULL == retval)
		return E_INVALIDARG;

	if (!(ddBFSizer & bpV1.m_dwFlags))
	{
		if (m_pBar->AmbientUserMode())
		{
			m_pBar->m_theErrorObject.SendError(IDERR_SIZERFLAGMUSTBESET, NULL);
			return CUSTOM_CTL_SCODE(IDERR_SIZERFLAGMUSTBESET);
		}
		else
			return E_FAIL;
	}

	SIZE size = {bpV1.m_nDockedVertMinWidth, 0};
	PixelToTwips(&size, &size);

	*retval = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::put_DockedVertMinWidth(short val)
{
	SIZE size = {val, 0};
	TwipsToPixel(&size, &size);
	bpV1.m_nDockedVertMinWidth = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::get_DockedHorzMinWidth(short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (!(ddBFSizer & bpV1.m_dwFlags))
	{
		if (m_pBar->AmbientUserMode())
		{
			m_pBar->m_theErrorObject.SendError(IDERR_SIZERFLAGMUSTBESET, NULL);
			return CUSTOM_CTL_SCODE(IDERR_SIZERFLAGMUSTBESET);
		}
		else
			return E_FAIL;
	}

	SIZE size = {bpV1.m_nDockedHorzMinWidth, 0};
	PixelToTwips(&size, &size);

	*retval = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::put_DockedHorzMinWidth(short val)
{
	SIZE size = {val, 0};
	TwipsToPixel(&size, &size);
	bpV1.m_nDockedHorzMinWidth = (short)size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::MapPropertyToCategory( DISPID dispid,  PROPCAT*ppropcat)
{
	switch (dispid)
	{
	case DISPID_PICTURE:
	case DISPID_GRABHANDLESTYLE:
		*ppropcat = PROPCAT_Appearance;
		return S_OK;

	case DISPID_DOCKINGOFFSET:
	case DISPID_DOCKLINE:
	case DISPID_HEIGHT:
	case DISPID_WIDTH:
	case DISPID_LEFT:
	case DISPID_TOP:
	case DISPID_HPAD:
	case DISPID_VPAD:
	case DISPID_HSPACE:
	case DISPID_VSPACE:
		*ppropcat = PROPCAT_Position;
		return S_OK;

	case DISPID_ENABLED:
	case DISPID_VISIBLE:
	case DISPID_DOCKINGAREA:
	case DISPID_FLAGS:
	case DISPID_TYPE:
	case DISPID_AUTOSIZEFORMS:
	case DISPID_CHILDBANDSTYLE:
	case DISPID_DISPLAYMORETOOLSBUTTON:
	case DISPID_WRAPPABLE:
		*ppropcat = PROPCAT_Behavior;
		return S_OK;

	case DISPID_NAME:
	case DISPID_CAPTION:
		*ppropcat = PROPCAT_Text;
		return S_OK;

	case DISPID_TAG:
		*ppropcat = PROPCAT_Data;
		return S_OK;

	default:
		return E_FAIL;
	}
}
STDMETHODIMP CBand::GetCategoryName( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName)
{
	return E_FAIL;
}
STDMETHODIMP CBand::GetToolIndex( VARIANT *Index,  short *index)
{
	*index = m_pTools->GetPosOfItem(Index);
	return NOERROR;
}
STDMETHODIMP CBand::DragDropExchange(IUnknown* pStreamIn,  VARIANT_BOOL vbSave)
{
	IStream* pStream = (IStream*)pStreamIn;
	try
	{
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			hResult = StWriteBSTR(pStream, m_bstrCaption);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteBSTR(pStream, m_bstrName);
			if (FAILED(hResult))
				return hResult;
	
			hResult = pStream->Write(&bpV1, sizeof(bpV1), NULL);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteVariant(pStream, m_vTag);
			if (FAILED(hResult))
				return hResult;

			hResult = PersistPicture(pStream, &m_phPicture, vbSave);
			if (FAILED(hResult))
				return hResult;

			hResult = m_pTools->DragDropExchange(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&m_pChildBands, sizeof(m_pChildBands), NULL);
			if (FAILED(hResult))
				return hResult;

			if (m_pChildBands)
			{
				hResult = m_pChildBands->DragDropExchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
			//
			// Loading
			//

			hResult = StReadBSTR(pStream, m_bstrCaption);
			if (FAILED(hResult))
				return hResult;

			hResult = StReadBSTR(pStream, m_bstrName);
			if (FAILED(hResult))
				return hResult;
			
			hResult = pStream->Read(&bpV1, sizeof(bpV1), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = StReadVariant(pStream, m_vTag);
			if (FAILED(hResult))
				return hResult;

			hResult = PersistPicture(pStream, &m_phPicture, vbSave);
			if (FAILED(hResult))
				return hResult;

			hResult = m_pTools->DragDropExchange(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			bpV1.m_rcFloat.Offset(-bpV1.m_rcFloat.left, -bpV1.m_rcFloat.top);

			if (ddBTStatusBar == bpV1.m_btBands)
				m_pBar->StatusBand(this);

			//
			// Was there a ChildBands Collection?
			//

			CChildBands* pChildBandTemp;
			hResult = pStream->Read(&pChildBandTemp, sizeof(pChildBandTemp), NULL);
			if (FAILED(hResult))
				return hResult;

			if (pChildBandTemp)
			{
				//
				// The answer is yes
				//

				if (NULL == m_pChildBands)
				{
					m_pChildBands = CChildBands::CreateInstance(NULL);
					if (NULL == m_pChildBands)
						return E_OUTOFMEMORY;

					m_pChildBands->SetOwner(this);
				}

				hResult = m_pChildBands->DragDropExchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
			OleTranslateColor(bpV1.m_ocPictureBackground, NULL, &m_crPictureBackground);
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
STDMETHODIMP CBand::get_PopupBannerBackgroundStyle(PopupBannerBackgroundStyles *retval)
{
	if (!m_pBar->AmbientUserMode() && !(ddBTPopup == bpV1.m_btBands || m_pParentBand))
		return E_FAIL;

	*retval = bpV1.m_pbsPicture;
	return NOERROR;
}
STDMETHODIMP CBand::put_PopupBannerBackgroundStyle(PopupBannerBackgroundStyles val)
{
	bpV1.m_pbsPicture = val;
	return NOERROR;
}
STDMETHODIMP CBand::get_PopupBannerBackgroundColor(OLE_COLOR *retval)
{
	if (!m_pBar->AmbientUserMode() && !(ddBTPopup == bpV1.m_btBands || m_pParentBand))
		return E_FAIL;

	*retval = bpV1.m_ocPictureBackground;
	return NOERROR;
}
STDMETHODIMP CBand::put_PopupBannerBackgroundColor(OLE_COLOR val)
{
	bpV1.m_ocPictureBackground = val;
	HRESULT hResult = OleTranslateColor(bpV1.m_ocPictureBackground, NULL, &m_crPictureBackground);
	if (FAILED(hResult))
		return hResult;
	return NOERROR;
}
STDMETHODIMP CBand::get_Picture(LPPICTUREDISP *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	
	if (!m_pBar->AmbientUserMode() && !(ddBTPopup == bpV1.m_btBands || m_pParentBand))
		return E_FAIL;

	*retval = m_phPicture.GetPictureDispatch();
	return NOERROR;
}
STDMETHODIMP CBand::put_Picture(LPPICTUREDISP val)
{
	m_phPicture.SetPictureDispatch(val);
	return NOERROR;
}
STDMETHODIMP CBand::putref_Picture(LPPICTUREDISP* val)
{
	if (NULL == val)
		return E_INVALIDARG;
	return put_Picture(*val);
}
STDMETHODIMP CBand::get_DisplayMoreToolsButton(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (m_pBar && !m_pBar->AmbientUserMode() && ddDAPopup == bpV1.m_daDockingArea)
		return E_FAIL;

	*retval = bpV1.m_vbDisplayMoreToolsButton;
	
	return NOERROR;
}
STDMETHODIMP CBand::put_DisplayMoreToolsButton(VARIANT_BOOL val)
{
	if (bpV1.m_vbDisplayMoreToolsButton != val)
	{
		bpV1.m_vbDisplayMoreToolsButton = val;
		if (m_pTools->m_pMoreTools)
		{
			if (VARIANT_FALSE == val)
			{
				m_pTools->m_pMoreTools->SetVisibleOnPaint(FALSE);
				m_pTools->m_pMoreTools->m_bMoreTools = FALSE;
				m_pTools->m_pMoreTools->Release();
				m_pTools->m_pMoreTools = NULL;
			}
			else
			{
				m_pTools->m_pMoreTools->SetVisibleOnPaint(TRUE);
				m_pTools->m_pMoreTools->m_bMoreTools = TRUE;
			}
			if (m_pBar)
				m_pBar->RecalcLayout();
		}
	}
	return NOERROR;
}
STDMETHODIMP CBand::get_AutoSizeForms(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (!(ddBFSizer & bpV1.m_dwFlags))
	{
		if (m_pBar->AmbientUserMode())
		{
			m_pBar->m_theErrorObject.SendError(IDERR_SIZERFLAGMUSTBESET, NULL);
			return CUSTOM_CTL_SCODE(IDERR_SIZERFLAGMUSTBESET);
		}
		else
			return E_FAIL;
	}

	*retval = bpV1.m_vbAutoSizeForms;
	return NOERROR;
}
STDMETHODIMP CBand::put_AutoSizeForms(VARIANT_BOOL val)
{
	bpV1.m_vbAutoSizeForms = val;
	return NOERROR;
}
STDMETHODIMP CBand::IsChild( VARIANT_BOOL* IsChild)
{
	if (NULL == IsChild)
		return E_INVALIDARG;

	*IsChild = m_pParentBand ? VARIANT_TRUE : VARIANT_FALSE;
	return NOERROR;
}
STDMETHODIMP CBand::get_DockedHorzWidth(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (!(ddBFSizer & bpV1.m_dwFlags))
	{
		if (m_pBar->AmbientUserMode())
		{
			m_pBar->m_theErrorObject.SendError(IDERR_SIZERFLAGMUSTBESET, NULL);
			return CUSTOM_CTL_SCODE(IDERR_SIZERFLAGMUSTBESET);
		}
		else
			return E_FAIL;
	}

	*retval = bpV1.m_rcHorzDockForm.Width();
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_DockedHorzWidth(long val)
{
	if (val < 0)
		return E_INVALIDARG;
	SIZE size;
	size.cx = val;
	TwipsToPixel(&size, &size);

	if (val < bpV1.m_nDockedHorzMinWidth)
		return E_FAIL;

	if (size.cx < bpV1.m_nDockedHorzMinWidth)
		size.cx = bpV1.m_nDockedHorzMinWidth;

	int nDiff = size.cx - (bpV1.m_rcHorzDockForm.right - bpV1.m_rcHorzDockForm.left);
	if (m_pDock)
		m_pDock->AdjustSizableBands(this, nDiff, TRUE);	return NOERROR;
}
STDMETHODIMP CBand::get_DockedHorzHeight(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (!(ddBFSizer & bpV1.m_dwFlags))
	{
		if (m_pBar->AmbientUserMode())
		{
			m_pBar->m_theErrorObject.SendError(IDERR_SIZERFLAGMUSTBESET, NULL);
			return CUSTOM_CTL_SCODE(IDERR_SIZERFLAGMUSTBESET);
		}
		else
			return E_FAIL;
	}

	*retval = bpV1.m_rcHorzDockForm.Height();
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_DockedHorzHeight(long val)
{
	if (val < 0)
		return E_INVALIDARG;
	
	SIZE size = {val, 0};
	TwipsToPixel(&size, &size);
	
	bpV1.m_rcHorzDockForm.bottom = bpV1.m_rcHorzDockForm.top + size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::get_DockedVertWidth(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (!(ddBFSizer & bpV1.m_dwFlags))
	{
		if (m_pBar->AmbientUserMode())
		{
			m_pBar->m_theErrorObject.SendError(IDERR_SIZERFLAGMUSTBESET, NULL);
			return CUSTOM_CTL_SCODE(IDERR_SIZERFLAGMUSTBESET);
		}
		else
			return E_FAIL;
	}

	*retval = bpV1.m_rcVertDockForm.Width();
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_DockedVertWidth(long val)
{
	if (val < 0)
		return E_INVALIDARG;
	SIZE size;
	size.cx = val;
	TwipsToPixel(&size, &size);
	
	if (size.cx < bpV1.m_nDockedVertMinWidth)
		size.cx = bpV1.m_nDockedVertMinWidth;

	bpV1.m_rcVertDockForm.right = bpV1.m_rcVertDockForm.left + size.cx;
	
	return NOERROR;
}
STDMETHODIMP CBand::get_DockedVertHeight(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (!(ddBFSizer & bpV1.m_dwFlags))
	{
		if (m_pBar && m_pBar->AmbientUserMode())
		{
			m_pBar->m_theErrorObject.SendError(IDERR_SIZERFLAGMUSTBESET, NULL);
			return CUSTOM_CTL_SCODE(IDERR_SIZERFLAGMUSTBESET);
		}
		else
			return E_FAIL;
	}

	*retval = bpV1.m_rcVertDockForm.Height();
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_DockedVertHeight(long val)
{
	if (val < 0)
		return E_INVALIDARG;

	SIZE size;
	size.cx = val;
	TwipsToPixel(&size, &size);

	if (val < bpV1.m_nDockedVertMinWidth)
		return E_FAIL;

	int nDiff = size.cx - (bpV1.m_rcVertDockForm.bottom - bpV1.m_rcVertDockForm.top);
	if (m_pDock)
		m_pDock->AdjustSizableBands(this, nDiff, FALSE);
	return NOERROR;
}
STDMETHODIMP CBand::GetDisplayString( DISPID dispID, BSTR __RPC_FAR *pBstr)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBand::MapPropertyToPage( DISPID dispID, CLSID __RPC_FAR *pClsid)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBand::GetPredefinedStrings(  DISPID dispID,  CALPOLESTR __RPC_FAR *pCaStringsOut,  CADWORD __RPC_FAR *pCaCookiesOut)
{
	switch (dispID)
	{
	case DISPID_FLAGS:
		{
			pCaStringsOut->cElems = 15;
			pCaStringsOut->pElems = (LPOLESTR*)CoTaskMemAlloc(sizeof(LPOLESTR*)*pCaStringsOut->cElems);
			if (NULL == pCaStringsOut->pElems)
				return E_OUTOFMEMORY;

			pCaCookiesOut->cElems = pCaStringsOut->cElems;
			pCaCookiesOut->pElems = (DWORD*)CoTaskMemAlloc(sizeof(DWORD)*pCaStringsOut->cElems);
			if (NULL == pCaCookiesOut->pElems)
			{
				CoTaskMemFree(pCaStringsOut->pElems);
				return E_OUTOFMEMORY;
			}
			
			DDString strFlag;
			DWORD dwFlag = 1;
			for (DWORD nIndex = 0; nIndex < pCaStringsOut->cElems; nIndex++)
			{
				strFlag.LoadString(IDS_BANDFLAGS0+nIndex);
				pCaStringsOut->pElems[nIndex] = strFlag.AllocSysString();
				pCaCookiesOut->pElems[nIndex] = dwFlag;
				dwFlag = dwFlag << 1;
			}
		}
		return S_OK;
	}
	return E_NOTIMPL;
}
STDMETHODIMP CBand::GetPredefinedValue( DISPID dispID, DWORD dwCookie, VARIANT __RPC_FAR *pVarOut)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBand::GetType( DISPID dispID,  long *pnType)
{
	switch (dispID)
	{
	case DISPID_FLAGS:
		*pnType = 2;
		return S_OK;
	}
	return E_NOTIMPL;
}
STDMETHODIMP CBand::get_TagVariant(VARIANT *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	HRESULT hResult = VariantCopy(retval, &m_vTag);
	return hResult;
}
STDMETHODIMP CBand::put_TagVariant(VARIANT val)
{
	HRESULT hResult = VariantCopy(&m_vTag, &val);
	return hResult;
}
STDMETHODIMP CBand::get_Name(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval=SysAllocString(m_bstrName);
	return NOERROR;
}
STDMETHODIMP CBand::put_Name(BSTR val)
{
	if (NULL == val)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_NULLNAMENOTALLOWED, val);
		return CUSTOM_CTL_SCODE(IDERR_NULLNAMENOTALLOWED);
	}

	if (m_pBar && m_pRootBand == this)
	{
		CBand* pBand = m_pBar->m_pBands->Find(val);
		if (NULL != pBand)
		{
			m_pBar->m_theErrorObject.SendError(IDERR_DUPLICATEBANDNAME, val);
			return CUSTOM_CTL_SCODE(IDERR_DUPLICATEBANDNAME);
		}
	}

	SysFreeString(m_bstrName);
	m_bstrName=SysAllocString(val);
	return NOERROR;
}
STDMETHODIMP CBand::InterfaceSupportsErrorInfo( REFIID riid)
{
	return S_OK;
}
STDMETHODIMP CBand::get_GrabHandleStyle(GrabHandleStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddBTStatusBar == bpV1.m_btBands)
		return E_FAIL;
	*retval=bpV1.m_ghsGrabHandleStyle;
	return NOERROR;
}
STDMETHODIMP CBand::put_GrabHandleStyle(GrabHandleStyles val)
{
	if (ddBTStatusBar == bpV1.m_btBands)
		return E_FAIL;

	switch (val)
	{
	case ddGSNone:
	case ddGSNormal:
	case ddGSDot:
	case ddGSSlash:
	case ddGSStripe:
	case ddGSIE:
	case ddGSFlat:
	case ddGSCaption:
		bpV1.m_ghsGrabHandleStyle=val;
		break;

	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}
STDMETHODIMP CBand::get_IsDetached(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (!m_pBar->AmbientUserMode() || (m_pBar && VARIANT_TRUE == m_pBar->m_vbLibrary))
		return E_FAIL;

	*retval = bpV1.m_vbDetached;
	return NOERROR;
}
STDMETHODIMP CBand::put_IsDetached(VARIANT_BOOL val)
{
	bpV1.m_vbDetached = val;
	return NOERROR;
}
STDMETHODIMP CBand::get_ChildBands(ChildBands** retval)
{
	if (NULL == m_pChildBands)
	{
		m_pChildBands = CChildBands::CreateInstance(NULL);
		if (NULL == m_pChildBands)
			return E_OUTOFMEMORY;
		m_pChildBands->SetOwner(this);
	}
	if (NULL == retval)
	{
		*retval = NULL;
		return E_INVALIDARG;
	}
	m_pChildBands->AddRef();
	*retval = reinterpret_cast<ChildBands*>(m_pChildBands);
	return NOERROR;
}
STDMETHODIMP CBand::get_ChildBandStyle(ChildBandStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (m_pBar && !m_pBar->AmbientUserMode() && ddBTNormal != bpV1.m_btBands)
		return E_FAIL;

	if (ddBTStatusBar == bpV1.m_btBands)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_NOTONSTATUSBAND, 0);
		return CUSTOM_CTL_SCODE(IDERR_NOTONSTATUSBAND);
	}

	if (m_pParentBand && m_pBar)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_CHILDBANDSHAVETOBENORMAL, 0);
		return CUSTOM_CTL_SCODE(IDERR_CHILDBANDSHAVETOBENORMAL);
	}

	*retval=bpV1.m_cbsChildStyle;
	return NOERROR;
}
STDMETHODIMP CBand::put_ChildBandStyle(ChildBandStyles val)
{
	if (ddBTStatusBar == bpV1.m_btBands)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_NOTONSTATUSBAND, 0);
		return CUSTOM_CTL_SCODE(IDERR_NOTONSTATUSBAND);
	}

	if (m_pParentBand && m_pBar)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_CHILDBANDSHAVETOBENORMAL, 0);
		return CUSTOM_CTL_SCODE(IDERR_CHILDBANDSHAVETOBENORMAL);
	}

	switch (val)
	{
	case ddCBSNone:
		bpV1.m_cbsChildStyle = val;
		break;
		
	case ddCBSTopTabs:
	case ddCBSBottomTabs:
	case ddCBSLeftTabs:
	case ddCBSRightTabs:
	case ddCBSToolbarTopTabs:
	case ddCBSToolbarBottomTabs:
		bpV1.m_cbsChildStyle = val;
		break;

	case ddCBSSlidingTabs:
		bpV1.m_cbsChildStyle = val;
		put_Width(1500);
		put_Height(1500);
		break;

	default:
		return E_INVALIDARG;
	}

	if (ddCBSNone != bpV1.m_cbsChildStyle)
	{
		if (NULL == m_pChildBands)
		{
			m_pChildBands = CChildBands::CreateInstance(NULL);
			if (NULL == m_pChildBands)
				return E_OUTOFMEMORY;
			m_pChildBands->SetOwner(this);
		}
		m_pChildBands->SetCurrentChildBand(0);
	}
	return NOERROR;
}
STDMETHODIMP CBand::get_CreatedBy(CreatedByTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (!m_pBar->AmbientUserMode() || (m_pBar && VARIANT_TRUE == m_pBar->m_vbLibrary))
		return E_FAIL;

	*retval = (CreatedByTypes)bpV1.m_nCreatedBy;
	return NOERROR;
}
STDMETHODIMP CBand::put_CreatedBy(CreatedByTypes val)
{
	bpV1.m_nCreatedBy=val;
	return NOERROR;
}
STDMETHODIMP CBand::get_DockLine(short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddBTStatusBar == bpV1.m_btBands)
		return E_FAIL;

	*retval=bpV1.m_nDockLine;
	return NOERROR;
}
STDMETHODIMP CBand::put_DockLine(short val)
{
	if (val < -1)
		return E_INVALIDARG;

	bpV1.m_nDockLine=val;
	return NOERROR;
}
STDMETHODIMP CBand::get_DockingOffset(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	
	if (ddBTStatusBar == bpV1.m_btBands)
		return E_FAIL;

	*retval=bpV1.m_nDockOffset;
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_DockingOffset(long val)
{
	if (val < 0)
		return E_INVALIDARG;

	SIZE size;
	size.cx = val;
	if (val > 0)
		TwipsToPixel(&size, &size);
	bpV1.m_nDockOffset=size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::get_Flags(BandFlags *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval=(BandFlags)bpV1.m_dwFlags;
	return NOERROR;
}
STDMETHODIMP CBand::put_Flags(BandFlags val)
{
	if (ddBTStatusBar == bpV1.m_btBands)
	{
		if (!(val & ddBFDockBottom))
			return E_FAIL;
		if (!(val & ddBFStretch))
			return E_FAIL;
		if (!(val & ddBFFixed))
			return E_FAIL;
		if (val & ddBFSizer)
			return E_FAIL;
		if (val & ddBFDockTop)
			return E_FAIL;
		if (val & ddBFDockLeft)
			return E_FAIL;
		if (val & ddBFDockRight)
			return E_FAIL;
	}

	if (!(ddBFSizer & bpV1.m_dwFlags) && (ddBFSizer & val))
	{
		bpV1.m_rcVertDockForm = bpV1.m_rcDimension;
		bpV1.m_rcHorzDockForm = bpV1.m_rcDimension;
	}
	bpV1.m_dwFlags = (DWORD)val;
	return NOERROR;
}
STDMETHODIMP CBand::get_DockingArea(DockingAreaTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddBTStatusBar == bpV1.m_btBands)
		return E_FAIL;
	
	*retval=bpV1.m_daDockingArea;
	return NOERROR;
}
STDMETHODIMP CBand::put_DockingArea(DockingAreaTypes val)
{
	switch (bpV1.m_btBands)
	{
	case ddBTStatusBar:
		return E_FAIL;
	}

	switch (val)
	{
    case ddDATop:
	case ddDABottom:
	case ddDALeft:
	case ddDARight:
	case ddDAFloat:
	case ddDAPopup:
		bpV1.m_daDockingArea = val;
		break;

	default:
		return E_INVALIDARG;
	}

	return NOERROR;
}
STDMETHODIMP CBand::get_ToolsHPadding(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = bpV1.m_nToolsHPadding;
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_ToolsHPadding(long val)
{
	if (val < 0)
		return E_INVALIDARG;
	SIZE size = {val, 0};
	if (val > 0)
		TwipsToPixel(&size, &size);
	bpV1.m_nToolsHPadding=size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::get_ToolsVPadding(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = bpV1.m_nToolsVPadding;
	if (*retval > 0)
	{
		SIZE size = {*retval, 0};
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_ToolsVPadding(long val)
{
	if (val < 0)
		return E_INVALIDARG;
	SIZE size = {val, 0};
	if (val > 0)
		TwipsToPixel(&size, &size);
	bpV1.m_nToolsVPadding=size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::get_ToolsHSpacing(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	*retval = bpV1.m_nToolsHSpacing;
	if (*retval > 0)
	{
		SIZE size = {*retval, 0};
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_ToolsHSpacing(long val)
{
	if (val < 0)
		return E_INVALIDARG;
	SIZE size = {val, 0};
	if (val > 0)
		TwipsToPixel(&size, &size);
	bpV1.m_nToolsHSpacing=size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::get_ToolsVSpacing(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	
	*retval = bpV1.m_nToolsVSpacing;
	if (*retval > 0)
	{
		SIZE size = {*retval, 0};
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_ToolsVSpacing(long val)
{
	if (val < 0)
		return E_INVALIDARG;
	SIZE size = {val, 0};
	if (val > 0)
		TwipsToPixel(&size, &size);
	bpV1.m_nToolsVSpacing=size.cx;
	return NOERROR;
}
STDMETHODIMP CBand::get_MouseTracking(TrackingStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_tsMouseTracking;
	return NOERROR;
}
STDMETHODIMP CBand::put_MouseTracking(TrackingStyles val)
{
    switch (val)
	{
	case ddTSNone:
	case ddTSBevel:
	case ddTSColor:
		bpV1.m_tsMouseTracking=val;
		break;

	default:
		return E_FAIL;
	}
	return NOERROR;
}
STDMETHODIMP CBand::get_Height(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	switch(bpV1.m_daDockingArea)
	{
	case ddDAFloat:
		*retval = bpV1.m_rcFloat.Height();
		break;
	case ddDATop :
	case ddDABottom :
	case ddDALeft :
	case ddDARight:
		*retval = m_rcDock.Height();
		break;
	default:		
		*retval=bpV1.m_rcDimension.bottom;
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
STDMETHODIMP CBand::put_Height(long val)
{
	if (val < 0)
		return E_INVALIDARG;

	if (ddDAFloat == bpV1.m_daDockingArea)
	{
		HWND hWnd;
		CRect rcBand;
		GetBandRect(hWnd, rcBand);
		GetWindowRect(hWnd, &rcBand);
		bpV1.m_rcFloat.top = rcBand.top;
		bpV1.m_rcFloat.bottom = rcBand.bottom;
	}

	SIZE size = {val, 0};
	if (val > 0)
		TwipsToPixel(&size, &size);

	bpV1.m_rcFloat.bottom = bpV1.m_rcFloat.top + size.cx;
	bpV1.m_rcDimension.bottom = bpV1.m_rcDimension.top + size.cx;

	switch (bpV1.m_daDockingArea)
	{
	case ddDATop:
	case ddDABottom:
		bpV1.m_rcHorzDockForm.bottom = bpV1.m_rcHorzDockForm.top + size.cx;
		break;

	case ddDALeft:
	case ddDARight:
		bpV1.m_rcVertDockForm.right = bpV1.m_rcVertDockForm.left + size.cx;
		break;
	}
	DelayedFloatResize();
	return NOERROR;
}
STDMETHODIMP CBand::get_Width(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	
	switch(bpV1.m_daDockingArea)
	{
	case ddDAFloat:
		*retval=bpV1.m_rcFloat.Width();
		break;

	case ddDATop :
	case ddDABottom :
	case ddDALeft :
	case ddDARight:
		*retval = m_rcDock.Width();
		break;

	default:		
		*retval=bpV1.m_rcDimension.right;
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
STDMETHODIMP CBand::put_Width(long val)
{
	if (val < 0)
		return E_INVALIDARG;
	
	if (ddDAFloat == bpV1.m_daDockingArea)
	{
		HWND hWnd;
		CRect rcBand;
		GetBandRect(hWnd, rcBand);
		GetWindowRect(hWnd, &rcBand);
		bpV1.m_rcFloat.left = rcBand.left;
		bpV1.m_rcFloat.right = rcBand.right;
	}

	SIZE size = {0, 0};
	size.cx = val;
	if (val > 0)
		TwipsToPixel(&size, &size);
	
	bpV1.m_rcFloat.right = bpV1.m_rcFloat.left + size.cx;
	bpV1.m_rcDimension.right = bpV1.m_rcDimension.left + size.cx;

	switch (bpV1.m_daDockingArea)
	{
	case ddDATop:
	case ddDABottom:
		if (size.cx < bpV1.m_nDockedHorzMinWidth)
			size.cx = bpV1.m_nDockedHorzMinWidth;
	
		bpV1.m_rcHorzDockForm.right = bpV1.m_rcHorzDockForm.left + size.cx;
		break;

	case ddDALeft:
	case ddDARight:
		if (size.cx < bpV1.m_nDockedVertMinWidth)
			size.cx = bpV1.m_nDockedVertMinWidth;

		bpV1.m_rcVertDockForm.bottom = bpV1.m_rcVertDockForm.top + size.cx;
		break;
	}
	DelayedFloatResize();
	return NOERROR;
}
STDMETHODIMP CBand::get_Left(long *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	if (ddDAFloat == bpV1.m_daDockingArea)
		*retval=bpV1.m_rcFloat.left;
	else
		*retval=bpV1.m_rcDimension.left;
	
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_Left(long val)
{
	if (val < 0 && ddDAFloat != bpV1.m_daDockingArea)
		return E_INVALIDARG;

	if (ddDAFloat == bpV1.m_daDockingArea)
	{
		HWND hWnd;
		CRect rcBand;
		GetBandRect(hWnd, rcBand);
		GetWindowRect(hWnd, &bpV1.m_rcFloat);
	}

	SIZE size;
	size.cx = val;
	TwipsToPixel(&size, &size);
	
	bpV1.m_rcFloat.right = size.cx + bpV1.m_rcFloat.Width();
	bpV1.m_rcFloat.left = size.cx;
	bpV1.m_rcDimension.right = size.cx + bpV1.m_rcDimension.Width();
	bpV1.m_rcDimension.left = size.cx;

	DelayedFloatResize();
	return NOERROR;
}
STDMETHODIMP CBand::get_Top(long* retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	
	if (ddDAFloat == bpV1.m_daDockingArea)
		*retval = bpV1.m_rcFloat.top;
	else
		*retval = bpV1.m_rcDimension.top;
	
	if (*retval > 0)
	{
		SIZE size;
		size.cx = *retval;
		PixelToTwips(&size, &size);
		*retval = size.cx;
	}
	return NOERROR;
}
STDMETHODIMP CBand::put_Top(long val)
{
	if (val < 0 && ddDAFloat != bpV1.m_daDockingArea)
		return E_INVALIDARG;

	if (ddDAFloat == bpV1.m_daDockingArea)
	{
		HWND hWnd;
		CRect rcBand;
		GetBandRect(hWnd, rcBand);
		GetWindowRect(hWnd, &bpV1.m_rcFloat);
	}

	SIZE size;
	size.cx = val;
	TwipsToPixel(&size, &size);
	
	bpV1.m_rcFloat.bottom = size.cx + bpV1.m_rcFloat.Height();
	bpV1.m_rcFloat.top = size.cx;
	bpV1.m_rcDimension.bottom = size.cx + bpV1.m_rcDimension.Height();
	bpV1.m_rcDimension.top = size.cx;
	
	DelayedFloatResize();
	return NOERROR;
}
STDMETHODIMP CBand::get_Type(BandTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = bpV1.m_btBands;
	return NOERROR;
}
STDMETHODIMP CBand::put_Type(BandTypes val)
{
	if (m_pBar->m_pBands->CheckForStatusBand(val))
		return E_FAIL;

	switch (val)
	{
	case ddBTPopup:
		{
			switch (bpV1.m_btBands)
			{
			case ddBTStatusBar:
				m_pBar->StatusBand(NULL);
				break;
			}
			bpV1.m_daDockingArea = ddDAPopup;
		}
		break;

	case ddBTNormal:
		{
			switch (bpV1.m_btBands)
			{
			case ddBTStatusBar:
				m_pBar->StatusBand(NULL);
				break;
			}
		}
		break;

	case ddBTMenuBar:
		switch (bpV1.m_btBands)
		{
		case ddBTStatusBar:
			m_pBar->StatusBand(NULL);
			break;
		}
		m_pBar->BuildAccelators(m_pBar->GetMenuBand());
		break;

	case ddBTStatusBar:
		if (m_pBar)
			m_pBar->StatusBand(this);
		bpV1.m_dwFlags = ddBFDockBottom|ddBFStretch|ddBFFixed|ddBFCustomize;
		bpV1.m_ghsGrabHandleStyle = ddGSNone; 
		bpV1.m_daDockingArea = ddDABottom;
		bpV1.m_vbWrapTools = VARIANT_FALSE;
		break;

	case ddBTChildMenuBar:
		switch (bpV1.m_btBands)
		{
		case ddBTStatusBar:
			m_pBar->StatusBand(NULL);
			break;
		}
		bpV1.m_vbVisible = VARIANT_FALSE;
		break;

	default:
		return E_FAIL;
	}
	bpV1.m_btBands = val;
	if (m_pBar->m_pDesignerNotify)
	{
		m_pBar->m_pDesignerNotify->BandChanged((LPDISPATCH)(IBand*)this, 
											   (LPDISPATCH)(IBand*)(ddCBSNone != bpV1.m_cbsChildStyle ? m_pChildBands->GetCurrentChildBand() : NULL),
											   ddBandModified);
	}
	return NOERROR;
}
STDMETHODIMP CBand::get_Visible(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=bpV1.m_vbVisible;
	return NOERROR;
}
STDMETHODIMP CBand::put_Visible(VARIANT_BOOL val)
{
	bpV1.m_vbVisible = val;
	if (VARIANT_FALSE == bpV1.m_vbVisible)
	{
		TabbedWindowedTools(FALSE);
		ParentWindowedTools(NULL);
		if (m_pBar)
		{
			if (m_pBar->ActiveBand() == this)
				m_pBar->ActiveBand(NULL);
			m_pBar->FireBandClose((Band*)this);
		}
	}
	else
	{
		HWND hWnd;
		CRect rcBand;
		if (GetBandRect(hWnd, rcBand))
			ParentWindowedTools(hWnd);
		TabbedWindowedTools(TRUE);
		if (m_pBar)
		{
			IReturnBool* pReturnBool = CRetBool::CreateInstance(NULL);
			if (pReturnBool)
			{
				HRESULT hResult = pReturnBool->put_Value(VARIANT_FALSE);
				if (SUCCEEDED(hResult))
				{
					++m_nInBandOpen;
					m_pBar->FireBandOpen((Band*)this, (ReturnBool*)pReturnBool);
					--m_nInBandOpen;
					
					VARIANT_BOOL vbCancel;
					hResult = pReturnBool->get_Value(&vbCancel);
					if (SUCCEEDED(hResult) && VARIANT_TRUE == vbCancel)
						bpV1.m_vbVisible = VARIANT_FALSE;
				}
				pReturnBool->Release();
			}
		}
	}
//	if (m_pBar)  CR 2217
//		m_pBar->RecalcLayout();
	return NOERROR;
}
STDMETHODIMP CBand::get_WrapTools(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	
	if (ddBTStatusBar == bpV1.m_btBands)
		return E_FAIL;

	*retval=bpV1.m_vbWrapTools;
	return NOERROR;
}
STDMETHODIMP CBand::put_WrapTools(VARIANT_BOOL val)
{
	if (ddBTStatusBar == bpV1.m_btBands)
		return E_FAIL;

	bpV1.m_vbWrapTools=val;
	return NOERROR;
}
STDMETHODIMP CBand::get_ActiveTool(Tool **retval)
{
	if (m_pBar && !m_pBar->AmbientUserMode())
		return E_FAIL;

	if (NULL == retval)
		return E_INVALIDARG;

	int nCount = m_pTools->GetToolCount();
	for (int nTool = 0; nTool < nCount; nTool++)
	{
		if (m_nCurrentTool == nTool)
		{
			*retval = (Tool*)m_pTools->GetTool(nTool);
			reinterpret_cast<CTool*>(*retval)->AddRef();
			return NOERROR;
		}
	}
	return E_FAIL;
}
STDMETHODIMP CBand::put_ActiveTool(Tool *val)
{
	int nCount = m_pTools->GetToolCount();
	for (int nTool = 0; nTool < nCount; nTool++)
	{
		if (val == (Tool*)m_pTools->GetTool(nTool))
		{
			m_nCurrentTool = nTool;
			return NOERROR;
		}
	}
	return E_FAIL;
}
STDMETHODIMP CBand::putref_ActiveTool(Tool **val)
{
	return put_ActiveTool(*val);
}

STDMETHODIMP CBand::GetTypeInfoCount( UINT *pctinfo)
{
	if (NULL == pctinfo)
		return E_INVALIDARG;

	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CBand::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	if (NULL == pptinfo)
		return E_INVALIDARG;

	*pptinfo=GetObjectTypeInfo(lcid, objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CBand::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	if (!(riid==IID_NULL))
        return E_INVALIDARG;

	ITypeInfo* pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;

	HRESULT hr = pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgdispid);
	pTypeInfo->Release();
	return hr;
}
STDMETHODIMP CBand::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
{
	if (!(riid==IID_NULL))
        return E_INVALIDARG;

	ITypeInfo* pTypeInfo = GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;
	
#ifdef ERRORINFO_SUPPORT
	SetErrorInfo(0L, NULL); // should be called if ISupportErrorInfo is used
#endif

	HRESULT hr = pTypeInfo->Invoke((IDispatch *)(IBand*)this, 
								   dispidMember, 
								   wFlags,
								   pdispparams, 
								   pvarResult,
								   pexcepinfo, 
								   puArgErr);
    pTypeInfo->Release();
	return hr;
}

// IBand members

STDMETHODIMP CBand::get_Caption(BSTR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=SysAllocString(m_bstrCaption);
	return NOERROR;
}
STDMETHODIMP CBand::put_Caption(BSTR val)
{
	SysFreeString(m_bstrCaption);
	m_bstrCaption = SysAllocString(val);
	if (m_pFloat)
	{
		MAKE_TCHARPTR_FROMWIDE(szCaption,m_bstrCaption);
		SetWindowText(m_pFloat->hWnd(),szCaption);
	}
	return NOERROR;
}
STDMETHODIMP CBand::get_Tools(Tools** retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=(Tools*)m_pTools;
	m_pTools->AddRef();
	return NOERROR;
}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
//
// SetOwner
//

void CBand::SetOwner(CBar* pBar, BOOL bSetRoot)
{
	m_pBar = pBar;
	m_pDockMgr = m_pBar->m_pDockMgr; 
	m_pTools->SetBand(this);
	m_pTools->SetBar(m_pBar);
	if (bSetRoot)
		m_pRootBand = this;
}

void CBand::DelayedFloatResize()
{
	if (ddDAFloat == bpV1.m_daDockingArea && VARIANT_TRUE == bpV1.m_vbVisible && m_pFloat && m_pFloat->IsWindow())
		m_pFloat->SendMessage(WM_COMMAND, MAKELONG(CMiniWin::eFloatResize, 0));
}

//
// AdjustFloatRect
//

void CBand::AdjustFloatRect(const POINT& pt)
{
	SIZE size = bpV1.m_rcFloat.Size();
	if (0 == size.cx || 0 == size.cy)
	{
		bpV1.m_rcFloat = GetOptimalFloatRect(TRUE);
		size = bpV1.m_rcFloat.Size();
	}
	bpV1.m_rcFloat.left = pt.x;
	bpV1.m_rcFloat.top = pt.y;
	bpV1.m_rcFloat.right = pt.x + size.cx;
	bpV1.m_rcFloat.bottom = pt.y + size.cy;
}

//
// CalcLayout
//

BOOL CBand::CalcLayout(const CRect& rcBound, DWORD dwLayoutFlag, CRect& rcReturn, BOOL bCommit)
{
	BOOL bFloat = 0 != (dwLayoutFlag & eLayoutFloat);
	BOOL bLeftHandle = (0 != (dwLayoutFlag & eLayoutHorz) && IsGrabHandle());
	BOOL bTopHandle = (0 != (dwLayoutFlag & eLayoutVert) && IsGrabHandle());

	BOOL bResult = CalcLayoutEx(rcBound,
								dwLayoutFlag,
								bLeftHandle,
								bTopHandle,
								rcReturn,
								bCommit,
								IsWrappable() || bFloat);
	return bResult;
}

//
// CalcLayoutEx
//
	
BOOL CBand::CalcLayoutEx(const CRect& rcBound,
						 DWORD        dwLayoutFlag,
						 BOOL         bLeftHandle,
						 BOOL         bTopHandle,
						 CRect&       rcReturn,
						 BOOL         bCommit,
						 BOOL		  bWrapFlag)
{
	BOOL bResult;
	rcReturn.SetEmpty();

	HDC hDC = GetDC(m_pBar->m_hWnd);
	if (NULL == hDC)
		return FALSE;

	CRect rcTest;
	int  nTotalWidth = rcBound.Width();
	int  nTotalHeight = rcBound.Height();
	int  nLineHeight;
	int  nLineCount;
	int  nMaxWidth;

	// Vertical Layout
	if (dwLayoutFlag & eLayoutVert)
	{
		if ((ddBTChildMenuBar == bpV1.m_btBands || ddBTMenuBar == bpV1.m_btBands) && !(eLayoutFloat & dwLayoutFlag))
			CalcHorzLayout(hDC, rcBound, bCommit, nTotalHeight, nTotalWidth, TRUE, bLeftHandle, bTopHandle, rcReturn, nLineCount, dwLayoutFlag);
		else
		{
			//
			// Assume fixed height and try to build smallest width by filling max vertical.
			// do this by increasing number of tools on first line until all tools fit
			//

			nMaxWidth = 3 + 3;
			BOOL bContinue = TRUE;
			while (bContinue)
			{
				nLineHeight = 0;
				rcTest.SetEmpty();
				bResult = CalcHorzLayout(hDC, rcBound, FALSE, nMaxWidth, nTotalHeight, TRUE, bLeftHandle, bTopHandle, rcTest, nLineCount, dwLayoutFlag);
				if (!bResult)
				{
					ReleaseDC(m_pBar->m_hWnd, hDC);
					return bResult;
				}

				if (bWrapFlag && rcTest.Height() > nTotalHeight && 1 != nLineCount)
					nMaxWidth += 12;
				else
					bContinue = FALSE;
			}
			nTotalWidth = rcTest.Width();
			dwLayoutFlag |= eLayoutHorz;
			bWrapFlag = TRUE;
		}
	}

	if (dwLayoutFlag & eLayoutHorz)
	{
		rcReturn.SetEmpty();
		bResult = CalcHorzLayout(hDC,
							     rcBound,
								 bCommit,
							     nTotalWidth,
								 nTotalHeight,
							     bWrapFlag,
							     bLeftHandle,
							     bTopHandle,
							     rcReturn,
							     nLineCount,
							     dwLayoutFlag);
		if (!bResult)
		{
			ReleaseDC(m_pBar->m_hWnd, hDC);
			return bResult;
		}
	}
	if ((ddBFDetach & m_pRootBand->bpV1.m_dwFlags) && (ddDAPopup == m_pRootBand->bpV1.m_daDockingArea || ddBTPopup == m_pRootBand->bpV1.m_btBands))
	{
		if (bCommit && m_prcTools)
		{
 			int nCount = GetVisibleToolCount();
			for (int nTool = 0; nTool < nCount; nTool++)
				m_prcTools[nTool].Offset(0, eDetachBarHeight);
		}
		rcReturn.bottom += eDetachBarHeight;
	}
	ReleaseDC(m_pBar->m_hWnd, hDC);
	return TRUE;
}

//
// CalcHorzLayout
//
// Calculates the bounding rectangle with a given width restriction
//

BOOL CBand::CalcHorzLayout(HDC	      hDC,
						   const CRect& rcBound,
						   BOOL	      bCommit,
						   const int& nTotalWidth,
						   const int& nTotalHeight,
						   BOOL		  bWrapFlag,
						   BOOL		  bLeftHandle,
						   BOOL       bTopHandle,
						   CRect&     rcReturn,
						   int&       nReturnLineCount,
						   DWORD      dwLayoutFlag)
{ 
	BOOL bResult = TRUE;
	int nInternalWidth = nTotalWidth;

	//
	// In Customization Mode all Tools need to be Visible
	//
	
	if (m_pBar->m_bCustomization || m_pBar->m_bAltCustomization) 
		bWrapFlag = TRUE;

	//
	// Add space for the Grabhandles, Close Button, or Expand Button
	//

	BOOL bVertical = IsVertical();

	if (ddDAPopup != m_pRootBand->bpV1.m_daDockingArea && IsGrabHandle())
	{
		//
		// Float window do not get grab handles, close buttons, or expand buttons
		//

		if (ddGSCaption == m_pRootBand->bpV1.m_ghsGrabHandleStyle)
			m_nAdjustForGrabBar = m_pBar->GetSmallFontHeight(!bVertical) + 2 * eBevelBorder2;
		else
		{
			if (ddGSNormal == m_pRootBand->bpV1.m_ghsGrabHandleStyle && m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				if (ddBFClose & bpV1.m_dwFlags || ddBFExpand & bpV1.m_dwFlags)
					m_nAdjustForGrabBar = eCloseButtonSpace + eCloseButton;
				else
					m_nAdjustForGrabBar = 7;
			}
			else
				m_nAdjustForGrabBar = eGrabTotalHeight + eBevelBorder2;
		}

		if (bVertical)
			rcReturn.bottom = rcReturn.top + m_nAdjustForGrabBar;
		else
			rcReturn.right = rcReturn.left + m_nAdjustForGrabBar;

		if (!bVertical)
			nInternalWidth -= m_nAdjustForGrabBar;
	}
	else
		m_nAdjustForGrabBar = 0;

	//
	// Adjusting for the different types of bands
	//

	switch (bpV1.m_btBands)
	{
	case ddBTPopup:
		if (ddBFDetach & bpV1.m_dwFlags && ddDAPopup == bpV1.m_daDockingArea)
		{
			//
			// Add space above band for detachable popups
			//

			rcReturn.bottom += CBand::eDetachBarHeight;
		}
		nInternalWidth -= 2 * eBevelBorder2; 
		rcReturn.right += eBevelBorder2;
		rcReturn.bottom += eBevelBorder2;
		break;

	case ddBTStatusBar:
		{
			HWND hWnd = m_pBar->GetDockWindow();
			DWORD dwStyle = GetWindowLong(hWnd, GWL_STYLE);
			m_bStatusBarSizerPresent = (!(dwStyle & WS_MAXIMIZE) && (CBar::eMDIForm == m_pBar->m_eAppType || (IsWindow(hWnd) && dwStyle & WS_THICKFRAME)));
			nInternalWidth -= eBevelBorder2; 
			rcReturn.bottom += eBevelBorder2;
		}
		break;

	default:

		if (ddDAPopup != m_pRootBand->bpV1.m_daDockingArea)
		{
			//
			// Space for the outside edge
			//

			nInternalWidth -= 2 * eBevelBorder2; 
			rcReturn.right += eBevelBorder2;
			rcReturn.bottom += eBevelBorder2;
		}
		break;
	}

	if (ddCBSNone == bpV1.m_cbsChildStyle)
	{
		//
		// m_sizeBandEdge is just the space taken up by the bands edges, grabhandles, buttons.
		//

		m_sizeBandEdge.cx = rcReturn.Width();
		m_sizeBandEdge.cy = rcReturn.Height();
		
		//
		// m_sizeEdgeOffset is the space you need to add to get to the inside of a band.
		//

		m_sizeEdgeOffset = m_sizeBandEdge;

		InsideCalcHorzLayout(hDC,
							 rcBound,
						     bCommit,
						     nInternalWidth,
						     bWrapFlag,
						     bLeftHandle,
						     bTopHandle,
						     rcReturn,
						     nReturnLineCount,
						     dwLayoutFlag);
	}
	else
	{
		//
		// The updating of m_sizeBandEdge and m_sizeEdgeOffset is done 
		// in the ChildBands::CalcHorzLayout function to take in consideration
		// the dimensions of the different types of tabs
		//

		if (m_pChildBands)
		{
			m_pChildBands->CalcHorzLayout(hDC,
										  rcBound,
										  bCommit,
										  nInternalWidth,
										  bWrapFlag,
										  bLeftHandle,
										  bTopHandle,
										  rcReturn,
										  nReturnLineCount,
										  dwLayoutFlag);
		}
	}

	//
	// Add space for the outside edge 
	//

	if (ddDAPopup != m_pRootBand->bpV1.m_daDockingArea)
	{
		rcReturn.right += eBevelBorder2;
		rcReturn.bottom += eBevelBorder2;

		m_sizeBandEdge.cx += eBevelBorder2;
		m_sizeBandEdge.cy += eBevelBorder2;

		if (ddCBSSlidingTabs == bpV1.m_cbsChildStyle)
		{
			CRect rcTemp;
			if (ddDAFloat == m_pRootBand->bpV1.m_daDockingArea)
				rcTemp = rcBound;
			else
			{
				rcTemp = rcReturn;
				if (bVertical)
					rcTemp.bottom = nTotalHeight;
				else
					rcTemp.right = nTotalWidth;
			}
			AdjustChildBandsRect(rcTemp);
			m_pChildBands->LayoutSlidingTabs(bVertical, rcTemp);
		}
	}
	return bResult;
}

//
// AdjustChildBandsRect
//

void CBand::AdjustChildBandsRect(CRect& rcChildBand, BOOL* pbVertical, BOOL* pbVerticalPaint, BOOL* pbLeftHandle, BOOL* pbTopHandle)
{	
	BOOL bVertical = IsVertical();
	if (pbVertical)
		*pbVertical = bVertical;

	if (pbVerticalPaint)
		*pbVerticalPaint = (ddBTChildMenuBar == bpV1.m_btBands || ddBTMenuBar == bpV1.m_btBands) && bVertical;

	if (pbLeftHandle)
		*pbLeftHandle = !bVertical && IsGrabHandle();

	if (pbTopHandle)
		*pbTopHandle = bVertical && IsGrabHandle();

	//
	// Adjust for the bands GrabHandle Style
	//

	if (IsGrabHandle())
	{
		if (bVertical)
			rcChildBand.top += m_pRootBand->m_nAdjustForGrabBar;
		else
			rcChildBand.left += m_pRootBand->m_nAdjustForGrabBar;
	}

	//
	// Adjust for the bands edges
	//

	rcChildBand.Inflate(-eBevelBorder2, -eBevelBorder2);
	
	if (ddCBSSlidingTabs == bpV1.m_cbsChildStyle)
		rcChildBand.Inflate(-2 * eBevelBorder, -2 * eBevelBorder);
}

//
// Draw
//

#ifdef PERFORMACE_TEST
	int g_paintCount=0;
	DWORD g_totalPaintTime=0;
	BOOL measurePersistToDraw=FALSE;
	extern DWORD persisttick;
#endif

void CBand::Draw(HDC hDC, CRect rcPaint, BOOL bWrapFlag, const POINT& ptPaintOffset)
{
#ifdef PERFORMACE_TEST
	if (measurePersistToDraw)
	{
		DWORD end=GetTickCount();	
		char str[100];
		wsprintf(str,"persist to first draw time=%d\n",end-persisttick);
		measurePersistToDraw=FALSE;
	}

	++g_paintCount;
	DWORD start,end;
	start=GetTickCount();
#endif

	m_rcInsideBand = rcPaint;
	m_rcCurrentPaint = rcPaint;
	m_rcGrabHandle = rcPaint;
	BOOL  bVertical = IsVertical();
	BOOL bVerticalMenuPaint = (ddBTChildMenuBar == bpV1.m_btBands || ddBTMenuBar == bpV1.m_btBands) && bVertical;
	BOOL bLeftHandle = FALSE;
	BOOL bTopHandle = FALSE;
	HRGN hRgnChildBands = NULL;
	BOOL bResult;
	
	switch (bpV1.m_daDockingArea)
	{
	case ddDATop:
	case ddDABottom:
	case ddDALeft:
	case ddDARight:
		{
			switch (bpV1.m_btBands)
			{
			case ddBTStatusBar:
				{
					//
					// Draw the line on top of the status bar
					//

					if (m_pDock->m_nLineCount > 1 && CBar::eMDIForm == m_pBar->m_eAppType)
					{
						CRect rcLine = rcPaint;
						rcLine.bottom = rcLine.top + eBevelBorder2;
						if (VARIANT_TRUE != m_pBar->bpV1.m_vbXPLook)
							m_pBar->DrawEdge(hDC, rcLine, BDR_SUNKENOUTER, BF_RECT);
					}
					m_rcInsideBand.Inflate(0, -eBevelBorder);
					m_rcInsideBand.top += eBevelBorder2;
				}
				break;

			default:

				if (!(ddBFFlat & bpV1.m_dwFlags))
				{
					//
					// Draw the edges around band
					//
				
					if (VARIANT_TRUE != m_pBar->bpV1.m_vbXPLook)
						m_pBar->DrawEdge(hDC, rcPaint, BDR_RAISEDINNER, BF_RECT);
				}

				m_rcInsideBand.Inflate(-eBevelBorder2, -eBevelBorder2);

				m_rcGrabHandle = rcPaint;
				if (bVertical)
					m_rcGrabHandle.bottom = m_rcGrabHandle.top + m_pRootBand->m_nAdjustForGrabBar;
				else
					m_rcGrabHandle.right = m_rcGrabHandle.left + m_pRootBand->m_nAdjustForGrabBar;
				
				if (ddGSNone != m_pRootBand->bpV1.m_ghsGrabHandleStyle && ddGSCaption == bpV1.m_ghsGrabHandleStyle)
				{
					if (bVertical)
					{
						m_rcGrabHandle.top += 1;
						m_rcGrabHandle.Inflate(-4, 0);
					}
					else
					{
						m_rcGrabHandle.left += 1;
						m_rcGrabHandle.Inflate(0, -4);
					}

					if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
					{
						FillSolidRect(hDC, m_rcGrabHandle, m_pBar->m_crXPBackground);
					}
					else
					{
						if (g_fSysWin98Shell)
						{
							BOOL bGradient;
							if (SystemParametersInfo(SPI_GETGRADIENTCAPTIONS, 0, &bGradient, FALSE) && bGradient)
							{
								if (CreateGradient(hDC, NULL, m_ffGradient, (m_pBar->ActiveBand() == this ? GetSysColor(COLOR_ACTIVECAPTION) : GetSysColor(COLOR_INACTIVECAPTION)), (m_pBar->ActiveBand() == this ? GetSysColor(COLOR_GRADIENTACTIVECAPTION) : GetSysColor(COLOR_GRADIENTINACTIVECAPTION)), !bVertical, m_rcGrabHandle.Width(), m_rcGrabHandle.Height()))
									m_ffGradient.Paint(hDC, m_rcGrabHandle.left, m_rcGrabHandle.top);
								else
									FillRect(hDC, &m_rcGrabHandle, m_pBar->ActiveBand() == this ? (HBRUSH)(1+COLOR_ACTIVECAPTION) : (HBRUSH)(1+COLOR_INACTIVECAPTION));
							}
							else
								FillRect(hDC, &m_rcGrabHandle, m_pBar->ActiveBand() == this ? (HBRUSH)(1+COLOR_ACTIVECAPTION) : (HBRUSH)(1+COLOR_INACTIVECAPTION));
						}
						else
							FillRect(hDC, &m_rcGrabHandle, m_pBar->ActiveBand() == this ? (HBRUSH)(1+COLOR_ACTIVECAPTION) : (HBRUSH)(1+COLOR_INACTIVECAPTION));
					}
				}

				if (ddBFClose & bpV1.m_dwFlags)
				{
					//
					// Draw the close button
					//
					
					m_rcCloseButton = m_rcGrabHandle;
					if (bVertical)
					{
						m_rcCloseButton.left = m_rcCloseButton.right - eCloseButtonSpace - eCloseButton;
						m_rcGrabHandle.right -= eCloseButton + eCloseButtonSpace;
					}
					else
					{
						if (ddGSCaption == bpV1.m_ghsGrabHandleStyle)
						{
							m_rcCloseButton.left += eCloseButtonSpace;
							m_rcCloseButton.top = m_rcCloseButton.bottom - eCloseButtonSpace - eCloseButton;
							m_rcGrabHandle.bottom = m_rcCloseButton.top;
						}
						else
						{
							m_rcCloseButton.left += eCloseButtonSpace;
							m_rcGrabHandle.top += eCloseButton + eCloseButtonSpace;
						}
					}
					m_rcCloseButton.right = m_rcCloseButton.left + eCloseButton;
					m_rcCloseButton.top += eCloseButtonSpace;
					m_rcCloseButton.bottom = m_rcCloseButton.top + eCloseButton;
					if (bVertical)
					{
						if (m_rcCloseButton.left < m_rcGrabHandle.left)
							m_bCloseButtonVisible = FALSE;
						else
						{
							m_bCloseButtonVisible = TRUE;
							DrawCloseButton(hDC);
						}
					}
					else
					{
						if (ddGSCaption == bpV1.m_ghsGrabHandleStyle)
						{
							if (m_rcCloseButton.top < m_rcGrabHandle.top)
								m_bCloseButtonVisible = FALSE;
							else
							{
								m_bCloseButtonVisible = TRUE;
								DrawCloseButton(hDC);
							}
						}
						else
						{
							if (m_rcCloseButton.bottom > m_rcGrabHandle.bottom)
								m_bCloseButtonVisible = FALSE;
							else
							{
								m_bCloseButtonVisible = TRUE;
								DrawCloseButton(hDC);
							}
						}
					}
					m_rcCloseButton.Offset(ptPaintOffset.x, ptPaintOffset.y);
				}

				if (ddBFExpand & bpV1.m_dwFlags)
				{
					//
					// Draw the Expand button
					//
				
					m_rcExpandButton = m_rcGrabHandle;
					if (bVertical)
					{
						m_rcExpandButton.left = m_rcExpandButton.right - eExpandButton - eExpandButtonSpace;
						m_rcGrabHandle.right -= eExpandButton + eExpandButtonSpace;
					}
					else
					{
						if (ddGSCaption == bpV1.m_ghsGrabHandleStyle)
						{
							m_rcExpandButton.left += eExpandButtonSpace;
							m_rcExpandButton.top = m_rcExpandButton.bottom - eExpandButtonSpace - eExpandButton;
						}
						else
						{
							m_rcExpandButton.left += eExpandButtonSpace;
							m_rcGrabHandle.top += eExpandButton + eExpandButtonSpace;
						}
					}
					m_rcExpandButton.right = m_rcExpandButton.left + eExpandButton;
					m_rcExpandButton.top += eExpandButtonSpace;
					m_rcExpandButton.bottom = m_rcExpandButton.top + eExpandButton;
					if (bVertical)
					{
						if (m_rcExpandButton.left < m_rcGrabHandle.left)
							m_bExpandButtonVisible = FALSE;
						else
						{
							m_bExpandButtonVisible = TRUE;
							DrawExpandButton(hDC);
						}
					}
					else
					{
						if (ddGSCaption == bpV1.m_ghsGrabHandleStyle)
						{
							if (m_rcExpandButton.top < m_rcGrabHandle.top)
								m_bExpandButtonVisible = FALSE;
							else
							{
								m_bExpandButtonVisible = TRUE;
								DrawExpandButton(hDC);
							}
						}
						else
						{
							if (m_rcExpandButton.bottom > m_rcGrabHandle.bottom)
								m_bExpandButtonVisible = FALSE;
							else
							{
								m_bExpandButtonVisible = TRUE;
								DrawExpandButton(hDC);
							}
						}
					}
					m_rcExpandButton.Offset(ptPaintOffset.x, ptPaintOffset.y);
				}

				//
				// Draw grab handle
				//

				if (IsGrabHandle() && m_rcGrabHandle.bottom > m_rcGrabHandle.top && m_rcGrabHandle.right > m_rcGrabHandle.left)
					DrawGrabHandles(hDC, bVertical, m_rcGrabHandle, bLeftHandle, bTopHandle);

				//
				// Adjust for grab handles or close button or expand button
				//

				if (IsGrabHandle())
				{
					if (bVertical)
						m_rcInsideBand.top += m_nAdjustForGrabBar;
					else
						m_rcInsideBand.left += m_nAdjustForGrabBar;
				}
				break;
			}
		}
		break;

	case ddDAFloat:
		if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
		{
			if (IsMenuBand())
				DrawBtnFaceBorder(hDC, rcPaint, m_pBar->m_crXPBackground);
			else
				DrawBtnFaceBorder(hDC, rcPaint, m_pBar->m_crXPBandBackground);
		}
		else
			DrawBtnFaceBorder(hDC, rcPaint, m_pBar->m_crBackground);
		m_rcInsideBand.Inflate(-eBevelBorder2, -eBevelBorder2);
		break;

	case ddDAPopup:
		break;
	}

	switch (bpV1.m_cbsChildStyle)
	{
	case ddCBSSlidingTabs:

		//
		// Draw the inside edge for the Sliding Tabs
		//

		m_rcInsideBand.Inflate(-eBevelBorder, -eBevelBorder);

		DrawEdge(hDC, &m_rcInsideBand, BDR_SUNKENOUTER, BF_RECT);

		m_rcInsideBand.Inflate(-eBevelBorder, -eBevelBorder);
		break;
	}

	switch (bpV1.m_btBands)
	{
	case ddBTPopup:
		{
			CRect rcText = m_rcInsideBand;
			if (m_phPicture.m_pPict)
			{
				SIZEL sizePicture = {0, 0};
				m_phPicture.m_pPict->get_Width(&sizePicture.cx);
				m_phPicture.m_pPict->get_Height(&sizePicture.cy);
				HiMetricToPixel(&sizePicture, &sizePicture);
				CRect rcImage = m_rcInsideBand;
				rcImage.right = rcImage.left + sizePicture.cx;
				switch (bpV1.m_pbsPicture)
				{
				case ddPBBSNormal:
					bResult = FillSolidRect(hDC, rcImage, m_crPictureBackground);
					break;

				case ddPBBSGradient:
					{
						CFlickerFree ffBackground;
						bResult = CreateGradient(hDC, NULL, ffBackground, RGB(0, 0, 0), m_crPictureBackground, TRUE, rcImage.Width(), rcImage.Height());
						ffBackground.Paint(hDC, rcImage.left, rcImage.top);
					}
					break;
				}

				if (sizePicture.cy < rcImage.Height())
					rcImage.top = rcImage.bottom - sizePicture.cy;
				m_phPicture.Render(hDC, rcImage, rcImage);
				rcText.bottom = rcImage.top; 
			}
/*			if (m_bstrCaption && *m_bstrCaption)
			{
				HFONT hFont = m_pBar->GetMenuFont(TRUE, FALSE);
				COLORREF crTextOld = SetTextColor(hDC, m_pBar->m_crForeground);
				HFONT hFontOld = SelectFont(hDC, hFont);
				int nModeOld = SetBkMode(hDC, TRANSPARENT);
				MAKE_TCHARPTR_FROMWIDE(szCaption, m_bstrCaption);
				DrawVerticalText(hDC, szCaption, lstrlen(szCaption), rcText);
				SelectFont(hDC, hFontOld);
				SetBkMode(hDC, nModeOld);
				SetTextColor(hDC, crTextOld);
			}
*/		}
		break;
	}

	int nClipResult = 0;
	HRGN hRgnCurrent = CreateRectRgn(0,0,0,0);
	if (hRgnCurrent)
		nClipResult = GetClipRgn(hDC, hRgnCurrent);

	hRgnChildBands = CreateRectRgnIndirect(&m_rcInsideBand); 
	assert(hRgnChildBands);

	if (hRgnChildBands)
	{
		if (1 == nClipResult)
			CombineRgn(hRgnChildBands, hRgnChildBands, hRgnCurrent, RGN_AND);
		SelectClipRgn(hDC, hRgnChildBands); 
	}

	//
	// Drawing Tabs and Tools
	//

	if (ddCBSNone == bpV1.m_cbsChildStyle)
		InsideDraw(hDC, m_rcInsideBand, bWrapFlag, bVertical, bVerticalMenuPaint, bLeftHandle, bTopHandle, ptPaintOffset);
	else
		m_pChildBands->Draw(hDC, m_rcInsideBand, bWrapFlag, bVertical, bVerticalMenuPaint, bLeftHandle, bTopHandle, ptPaintOffset);

	if (hRgnChildBands)
	{
		if (1 == nClipResult)
			SelectClipRgn(hDC, hRgnCurrent); 
		else
			SelectClipRgn(hDC, NULL); 
		bResult = DeleteRgn(hRgnChildBands);   
		assert(bResult);
	}

	if (hRgnCurrent)
	{
		int bResult = DeleteRgn(hRgnCurrent);
		assert(bResult);
	}

	switch (bpV1.m_btBands)
	{
	case ddBTStatusBar:

		//
		// Draw the sizer for the status bar
		//

		if (m_bStatusBarSizerPresent)
		{
			m_rcSizer = rcPaint; 
			m_rcSizer.left = rcPaint.right - GetSystemMetrics(SM_CXSIZE);
			m_rcSizer.top = rcPaint.bottom - GetSystemMetrics(SM_CYSIZE);
			DrawSizer(hDC, m_rcSizer);
		}
		break;
	}
#ifdef PERFORMACE_TEST
	end=GetTickCount();
	char str[200];
	
	g_totalPaintTime+=end-start;

	wsprintf(str,"Total Band Draw [%d] total time=[%d]  , this time=%d\n",g_paintCount,g_totalPaintTime,end-start);
//	OutputDebugString(str);
#endif
}

//
// DrawSizer
//

BOOL CBand::DrawSizer(HDC hDC, const CRect& rcSizer)
{
	HPEN hLight = CreatePen(PS_SOLID, 1, m_pBar->m_crHighLight);
	HPEN hDark = CreatePen(PS_SOLID, 1, m_pBar->m_crShadow);
	int nROPOld = SetROP2(hDC, R2_COPYPEN);
	HPEN hPenOld = SelectPen(hDC, hLight);
	BOOL bResult;
	int nXStart = rcSizer.right;
	int nYStart = rcSizer.bottom;
	int nXEnd = rcSizer.right;
	int nYEnd = rcSizer.bottom;
	for (int nLine = 0; nLine < 3; nLine++)
	{
		// Draw Dark Line

		nXStart--;
		nYEnd--;

		SelectPen(hDC, hDark);
		MoveToEx(hDC, nXStart, nYStart, NULL);
		LineTo(hDC, nXEnd, nYEnd);

		nXStart--;
		nYEnd--;

		MoveToEx(hDC, nXStart, nYStart, NULL);
		LineTo(hDC, nXEnd, nYEnd);

		nXStart--;
		nYEnd--;

		MoveToEx(hDC, nXStart, nYStart, NULL);
		LineTo(hDC, nXEnd, nYEnd);

		nXStart--;
		nYEnd--;
		SelectPen(hDC, hLight);
		// Draw Light Line
		MoveToEx(hDC, nXStart, nYStart, NULL);
		LineTo(hDC, nXEnd, nYEnd);
		nXStart--;
		nYEnd--;
	}
	SelectPen(hDC, hPenOld);
	bResult = DeletePen(hLight);
	assert(bResult);
	bResult = DeletePen(hDark);
	assert(bResult);
	SetROP2(hDC, nROPOld);
	return TRUE;
}

//
// DrawFlatEdge
//

BOOL CBand::DrawFlatEdge(HDC hDC, const CRect& rcEdge, BOOL bVertical)
{
	BOOL bResult = TRUE;
	CRect rc = rcEdge;
	if (bVertical)
	{
		rc.bottom = rc.top + 1;
		bResult = FillSolidRect(hDC, rc, m_pBar->m_crHighLight);

		rc = rcEdge;
		rc.top++;
		rc.bottom = rc.top + 1;
		bResult = FillSolidRect(hDC, rc, m_pBar->m_crShadow);
	}
	else
	{
		rc.right = rc.left + 1;
		bResult = FillSolidRect(hDC, rc, m_pBar->m_crHighLight);

		rc = rcEdge;
		rc.left++;
		rc.right = rc.left + 1;
		bResult = FillSolidRect(hDC, rc, m_pBar->m_crShadow);
	}
	return bResult;
}

//
// DrawGrabHandles
//

void CBand::DrawGrabHandles(HDC   hDC, 
							BOOL  bVertical, 
							CRect rcGrabHandle, 
							BOOL& bLeftHandle, 
							BOOL& bTopHandle)
{
	switch (bpV1.m_ghsGrabHandleStyle)
	{
	case ddGSNormal:
		{
			if (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				HPEN aPen = CreatePen(PS_SOLID, 1, RGB(165, 162, 165));
				if (aPen)
				{
					HPEN hPenOld = SelectPen(hDC, aPen);
					int nX, nY;
					if (!bVertical)
					{
						rcGrabHandle.Inflate(-2, -4);
						nY = rcGrabHandle.top;
						while (nY < rcGrabHandle.bottom)
						{
							nX = rcGrabHandle.left;
							MoveToEx(hDC, nX, nY, NULL);
							nX = rcGrabHandle.right;
							LineTo(hDC, nX, nY);
							nY += 2;
						}
					}
					else
					{
						rcGrabHandle.Inflate(-4, -2);
						nX = rcGrabHandle.left;
						while (nX < rcGrabHandle.right)
						{
							nY = rcGrabHandle.top;
							MoveToEx(hDC, nX, nY, NULL);
							nY = rcGrabHandle.bottom;
							LineTo(hDC, nX, nY);
							nX += 2;
						}
					}
					SelectPen(hDC, hPenOld);
					BOOL bResult = DeletePen(aPen);
					assert(bResult);
				}
			}
			else
			{
				if (bVertical)
				{	
					// top handle
					rcGrabHandle.top += eGrabHandleX;
					rcGrabHandle.bottom = rcGrabHandle.top + eGrabHandleWidth;
					rcGrabHandle.left += eGrabHandleY;
					rcGrabHandle.right -= eGrabHandleY;
				} 
				else
				{	
					// left handle
					rcGrabHandle.top += eGrabHandleY;
					rcGrabHandle.bottom -= eGrabHandleY;
					rcGrabHandle.left += eGrabHandleX;
					rcGrabHandle.right = rcGrabHandle.left + eGrabHandleWidth;
					bLeftHandle = TRUE;
				}
				m_pBar->DrawEdge(hDC, rcGrabHandle, BDR_RAISEDINNER, BF_RECT);
				
				if (ddGSNormal == bpV1.m_ghsGrabHandleStyle)
				{
					rcGrabHandle.Offset(!bVertical ? eGrabHandleWidth : 0, bVertical ? eGrabHandleWidth : 0);
					
					m_pBar->DrawEdge(hDC, rcGrabHandle, BDR_RAISEDINNER, BF_RECT);
				}
			}
		}
		break;

	case ddGSIE:
		if (bVertical)
		{	
			// top handle
			rcGrabHandle.top += eGrabHandleX;
			rcGrabHandle.bottom = rcGrabHandle.top + eGrabHandleWidth;
			rcGrabHandle.left += eGrabHandleY;
			rcGrabHandle.right -= eGrabHandleY;
		} 
		else
		{	
			// left handle
			rcGrabHandle.top += eGrabHandleY;
			rcGrabHandle.bottom -= eGrabHandleY;
			rcGrabHandle.left += eGrabHandleX;
			rcGrabHandle.right = rcGrabHandle.left + eGrabHandleWidth;
			bLeftHandle = TRUE;
		}
		m_pBar->DrawEdge(hDC, rcGrabHandle, BDR_RAISEDINNER, BF_RECT);
		break;

	case ddGSFlat:
		if (bVertical)
		{	
			// top handle
			rcGrabHandle.top += eGrabHandleX;
			rcGrabHandle.bottom = rcGrabHandle.top + eGrabHandleWidth;
			rcGrabHandle.left += eGrabHandleY;
			rcGrabHandle.right -= eGrabHandleY;
		} 
		else
		{	
			// left handle
			rcGrabHandle.top += eGrabHandleY;
			rcGrabHandle.bottom -= eGrabHandleY;
			rcGrabHandle.left += eGrabHandleX;
			rcGrabHandle.right = rcGrabHandle.left + eGrabHandleWidth;
			bLeftHandle = TRUE;
		}		
		DrawFlatEdge(hDC, rcGrabHandle, bVertical);
		
		rcGrabHandle.Offset(!bVertical ? eGrabHandleWidth : 0, bVertical ? eGrabHandleWidth : 0);
		
		DrawFlatEdge(hDC, rcGrabHandle, bVertical);
		break;

	case ddGSDot:
	case ddGSSlash:
	case ddGSStripe:
		{
			int nIndex = bpV1.m_ghsGrabHandleStyle - 2;
			if (bVertical) 
			{	
				// top handle
				if (nIndex == ddGSSlash-1) 
					nIndex = ddGSStripe-1; // vertical /horz lines correction

				rcGrabHandle.top += eGrabHandleX;
				rcGrabHandle.bottom = rcGrabHandle.top + eGrabTotalHeight;
				rcGrabHandle.left += eGrabHandleY;
				rcGrabHandle.right -= eGrabHandleY;
			}
			else
			{
				// left handle
				rcGrabHandle.top += eGrabHandleY;
				rcGrabHandle.bottom -= eGrabHandleY;
				rcGrabHandle.left += eGrabHandleX;
				rcGrabHandle.right = rcGrabHandle.left + eGrabTotalWidth;
			}
			if (g_fSysWinNT)
				FillRect(hDC, &rcGrabHandle, GetGlobals().GetGrabHandleBrush(nIndex));
			else
			{
				HBITMAP bmp = (HBITMAP)GetGlobals().GetGrabHandleBrush(nIndex);
				assert(bmp);
				if (bmp)
				{
					BITMAP bmInfo;
					GetObject(bmp, sizeof(BITMAP), &bmInfo);
					FillRgnBitmap(hDC, 0, bmp, rcGrabHandle, bmInfo.bmWidth, bmInfo.bmHeight);
				}
			}
		}
		break;

	case ddGSCaption:
		{
			rcGrabHandle.Inflate(-1, -1);

			HFONT hFontSmall = m_pBar->GetSmallFont(!bVertical);
			HFONT hFontOld = SelectFont(hDC, hFontSmall);
			
			SetBkMode(hDC, TRANSPARENT);
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				SetTextColor(hDC, GetSysColor(COLOR_CAPTIONTEXT));
			else
			{
				if (g_fSysWin98Shell)
					SetTextColor(hDC, m_pBar->ActiveBand() == this ? GetSysColor(COLOR_CAPTIONTEXT) : GetSysColor(COLOR_INACTIVECAPTIONTEXT));
				else
					SetTextColor(hDC, m_pBar->ActiveBand() == this ? GetSysColor(COLOR_CAPTIONTEXT) : GetSysColor(COLOR_INACTIVECAPTIONTEXT));
			}

			MAKE_TCHARPTR_FROMWIDE(szCaption, m_bstrCaption);
			
			if (!bVertical)
				ExtTextOut(hDC, rcGrabHandle.right, rcGrabHandle.top, ETO_CLIPPED, &rcGrabHandle, szCaption, lstrlen(szCaption), NULL);
			else
				ExtTextOut(hDC, rcGrabHandle.left, rcGrabHandle.top, ETO_CLIPPED, &rcGrabHandle, szCaption, lstrlen(szCaption), NULL);
			
			SelectFont(hDC, hFontOld);
		}
		break;
	}
}


//
// DrawTool
//

void CBand::DrawTool(HDC hDC, CTool* pTool, CRect& rcTool, BOOL bVertical, const POINT& ptPaintOffset, BOOL bDrawControlsOrForms)
{
	switch (pTool->tpV1.m_ttTools)
	{
	case ddTTControl:
		if ((m_pBar->m_bCustomization || m_pBar->m_bAltCustomization) && m_pBar->m_diCustSelection.pTool == pTool)
			DSGDrawSelection((OLE_HANDLE)hDC, rcTool.left, rcTool.top, rcTool.Width(), -rcTool.Height());
		else if (IsWindow(pTool->m_hWndActive))
		{
			HWND hWnd;
			CRect rcControl;
			CRect rcTmp;
			if (GetToolBandRect(GetIndexOfTool(pTool), rcControl, hWnd))
			{
				pTool->SetParent(hWnd);
				BOOL bShow = VARIANT_TRUE == pTool->tpV1.m_vbVisible && !(m_pParentBand ? m_pParentBand->ChildBandChanging() : FALSE);

				GetWindowRect(pTool->m_hWndActive, &rcTmp);
				ScreenToClient(hWnd, rcTmp);
				if (rcTmp != rcControl || !((WS_VISIBLE & GetWindowLong(pTool->m_hWndActive, GWL_STYLE)) && bShow))
				{
					::SetWindowPos(pTool->m_hWndActive, 
								   NULL, 
								   rcControl.left, 
								   rcControl.top, 
								   rcControl.Width(), 
								   rcControl.Height(), 
								   SWP_NOACTIVATE | SWP_NOZORDER | (bShow ? SWP_SHOWWINDOW : SWP_HIDEWINDOW));

				}
			}
		}
		break;

	case ddTTForm:
		if ((m_pBar->m_bCustomization || m_pBar->m_bAltCustomization) && m_pBar->m_diCustSelection.pTool == pTool)
			DSGDrawSelection((OLE_HANDLE)hDC, rcTool.left, rcTool.top, rcTool.Width(), -rcTool.Height());
		else if (IsWindow(pTool->m_hWndActive))
		{
			HWND hWnd;
			CRect rcControl;
			CRect rcTmp;
			if (GetToolBandRect(GetIndexOfTool(pTool), rcControl, hWnd))
			{
				rcTool.Offset(ptPaintOffset.x, ptPaintOffset.y);
				pTool->SetParent(hWnd);
				BOOL bShow = VARIANT_TRUE == pTool->tpV1.m_vbVisible && !(m_pParentBand ? m_pParentBand->ChildBandChanging() : FALSE);
				
				GetWindowRect(pTool->m_hWndActive, &rcTmp);
				ScreenToClient(hWnd, rcTmp);

				if (rcTmp != rcControl || !((WS_VISIBLE & GetWindowLong(pTool->m_hWndActive, GWL_STYLE)) && bShow))
				{
					::SetWindowPos(pTool->m_hWndActive, 
								   NULL, 
								   rcControl.left, 
								   rcControl.top, 
								   rcControl.Width(), 
								   rcControl.Height(), 
								   SWP_NOACTIVATE | SWP_NOZORDER | (bShow ? SWP_SHOWWINDOW : SWP_HIDEWINDOW));
					::UpdateWindow(pTool->m_hWndActive);
				}
			}
		}
		break;

	default:
		//
		// Draw tool image and border
		//

		if (ddTSColor == bpV1.m_tsMouseTracking && pTool != m_pBar->m_pActiveTool && !pTool->m_bPressed)
			pTool->m_bGrayImage = TRUE;

		pTool->m_nShortCutOffset = m_nPopupShortCutStringOffset;

		//
		// Force normal
		//
		
		if (ddBTPopup == bpV1.m_btBands && VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
		{
			rcTool.Offset(1, 1);
			rcTool.Inflate(1, 1);
		}
		if (RectVisible(hDC,&rcTool))
		{
			if (ddBTPopup == m_pRootBand->bpV1.m_btBands && !m_bPopupWinLock)
				pTool->Draw(hDC, rcTool, ddBTNormal, IsVertical(), bVertical); 
			else
				pTool->Draw(hDC, rcTool, bpV1.m_btBands, IsVertical(), bVertical);

			if ((m_pBar->m_bCustomization || m_pBar->m_bAltCustomization) && m_pBar->m_diCustSelection.pTool == pTool)
				DSGDrawSelection((OLE_HANDLE)hDC, rcTool.left, rcTool.top, rcTool.Width(), -rcTool.Height());
		}
		pTool->m_bGrayImage = FALSE;
		break;
	}
}

//
// Refresh
//

STDMETHODIMP CBand::Refresh()
{
	try
	{
		if (VARIANT_FALSE == bpV1.m_vbVisible)
			return NOERROR;

		switch (bpV1.m_daDockingArea)
		{
		case ddDATop:
		case ddDABottom:
		case ddDALeft:
		case ddDARight:
			if (m_pDock && m_pDock->IsWindow())
				m_pDock->InvalidateRect(NULL, FALSE);
			break;

		case ddDAFloat:
			if (m_pFloat && m_pFloat->IsWindow())
				m_pFloat->InvalidateRect(NULL, FALSE);
			break;

		default:
			if (ddBTPopup == bpV1.m_btBands && m_pPopupWin && m_pPopupWin->IsWindow())
				m_pPopupWin->InvalidateRect(NULL, FALSE);
			break;
		}
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return E_FAIL;
}

//
// InvalidateRect
//

void CBand::InvalidateRect(CRect* prcUpdate, BOOL bBackground)
{
	try
	{
		switch (bpV1.m_daDockingArea)
		{
		case ddDAFloat:
			if (m_pFloat->IsWindow())
				m_pFloat->InvalidateRect(prcUpdate, bBackground);
			break;

		case ddDAPopup:
			if (m_pPopupWin->IsWindow())
				m_pPopupWin->InvalidateRect(prcUpdate, bBackground);
			break;

		default:
			if (m_pDock->IsWindow())
				m_pDock->InvalidateRect(prcUpdate, bBackground);
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
// InvalidateToolRect
//

void CBand::InvalidateToolRect(int nToolIndex)
{
	try
	{
		int nCount = m_pTools->GetVisibleToolCount();
		if (nToolIndex < 0 || nToolIndex >= nCount)
			return;
		CTool* pTool = m_pTools->GetVisibleTool(nToolIndex);
		HWND  hWnd;
		CRect rcTool;
		if (GetToolBandRect(nToolIndex, rcTool, hWnd))
			InvalidateToolRect(pTool, rcTool);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// InvalidateToolRect
//

void CBand::InvalidateToolRect(CTool* pTool, const CRect& rcTool)
{
	try
	{
		if (!pTool->IsVisibleOnPaint())
			return;

		switch (pTool->tpV1.m_ttTools)
		{
		case ddTTControl:
		case ddTTForm:
		case ddTTSeparator:
			if (NULL == m_pBar->m_pDesigner)
				return;
		}
		BOOL bResult;
		CRect rcPaint(rcTool);
		HWND hWnd = NULL;
		switch (m_pRootBand->bpV1.m_daDockingArea)
		{
		case ddDAFloat:
			hWnd = m_pRootBand->m_pFloat->hWnd();
			break;

		case ddDAPopup:
			hWnd = m_pRootBand->m_pPopupWin->hWnd();
			break;

		default:
			hWnd = m_pRootBand->m_pDock->hWnd();
			break;
		}

		HDC hDCDest = GetDC(hWnd);
		if (NULL == hDCDest)
			return;

		rcPaint.Inflate(eBevelBorder, eBevelBorder);
		CRect rcScreenBrush = rcPaint;
		CRect rcTotal = rcPaint;

		// For Menu bar tools I need to know how the tool is oriented, how the menu is flipped and if 
		// the sub menu is displayed

		BOOL bVertical = IsVertical();

		HDC hDC = m_pBar->m_ffObj.RequestDC(hDCDest, rcTotal.Width(), rcTotal.Height());
		if (NULL == hDC)
			hDC = hDCDest;
		else
			rcTotal.Offset(-rcTotal.left, -rcTotal.top);

		if (m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle)
		{
			//
			// Sliding Tabs
			//

			m_pParentBand->m_pChildBands->DrawToolSlidingTabBackground(hDC, 
																	   rcTotal, 
																	   m_prcTools[m_nToolMouseOver], 
																	   bVertical);
		}
		else
		{
			BOOL bTexture = m_pBar->HasTexture();
			switch (bpV1.m_btBands)
			{
			case ddBTPopup:
				if (bTexture)
					bTexture = ddBOPopups & m_pBar->bpV1.m_dwBackgroundOptions;
				break;

			case ddBTNormal:
				if (bTexture)
				{
					if (ddDAFloat == bpV1.m_daDockingArea)
						bTexture = ddBOFloat & m_pBar->bpV1.m_dwBackgroundOptions;
					else if (ddDAPopup == bpV1.m_daDockingArea)
						bTexture = ddBOPopups & m_pBar->bpV1.m_dwBackgroundOptions;
					else 
						bTexture = ddBODockAreas & m_pBar->bpV1.m_dwBackgroundOptions;
				}
				break;

			case ddBTMenuBar:
			case ddBTStatusBar:
			case ddBTChildMenuBar:
				if (bTexture)
				{
					if (ddDAFloat == bpV1.m_daDockingArea)
						bTexture = ddBOFloat & m_pBar->bpV1.m_dwBackgroundOptions;
					else
						bTexture = ddBODockAreas & m_pBar->bpV1.m_dwBackgroundOptions;
				}
				break;
			}
			
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				if (!(ddBTChildMenuBar == bpV1.m_btBands || ddBTMenuBar == bpV1.m_btBands || ddBTStatusBar == bpV1.m_btBands))
					FillSolidRect(hDC, rcTotal, m_pBar->m_crXPBandBackground);
				else
					FillSolidRect(hDC, rcTotal, m_pBar->m_crXPBackground);
			}
			else
			{
				if (bTexture)
				{
					if (m_pBar->m_hBrushTexture)
						UnrealizeObject(m_pBar->m_hBrushTexture);

					SetBrushOrgEx(hDC, -rcScreenBrush.left, -rcScreenBrush.top, NULL);
					m_pBar->FillTexture(hDC, rcTotal);
					SetBrushOrgEx(hDC, 0, 0, NULL);
				}
				else
					FillSolidRect(hDC, rcTotal, m_pBar->m_crBackground);
			}
		}

		//
		// Draw border of tool
		//
		
		rcTotal.Inflate(-eBevelBorder, -eBevelBorder);
		POINT ptPaintOffset = {0 , 0};

		DrawTool(hDC, 
				 pTool, 
				 rcTotal, 
				 (ddBTChildMenuBar == bpV1.m_btBands || ddBTMenuBar == bpV1.m_btBands) && bVertical,
				 ptPaintOffset,
				 FALSE);

		HRGN hRgnDest = NULL;
		int nResult = -1;
		HRGN hRgnPrev = NULL;
		if (m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle)
		{
			CRect rcTmpCurrentPage = m_rcCurrentPage;
			rcTmpCurrentPage.Inflate(-1, -1);

			hRgnDest = CreateRectRgnIndirect(&rcTmpCurrentPage); 
			assert(hRgnDest);
			if (NULL == hRgnDest)
				return;

			if (BTabs::eInactiveBottomButton != m_pParentBand->m_pChildBands->m_sbsBottomStyle)
			{
				HRGN hRgnBottomButton = CreateRectRgnIndirect(&m_pParentBand->m_pChildBands->m_rcBottomButton); 
				assert(hRgnBottomButton);
				if (NULL == hRgnBottomButton)
				{
					DeleteRgn(hRgnDest);
					ReleaseDC(hWnd, hDCDest);
					return;
				}

				CombineRgn(hRgnDest, hRgnDest, hRgnBottomButton, RGN_DIFF);
			
				bResult = DeleteRgn(hRgnBottomButton);
				assert(bResult);
			}

			if (BTabs::eInactiveTopButton != m_pParentBand->m_pChildBands->m_sbsTopStyle)
			{
				HRGN hRgnTopButton = CreateRectRgnIndirect(&m_pParentBand->m_pChildBands->m_rcTopButton); 
				assert(hRgnTopButton);
				if (NULL == hRgnTopButton)
				{
					DeleteRgn(hRgnDest);
					ReleaseDC(hWnd, hDCDest);
					return;
				}

				CombineRgn(hRgnDest, hRgnDest, hRgnTopButton, RGN_DIFF);

				bResult = DeleteRgn(hRgnTopButton);
				assert(bResult);
			}

			if (hRgnDest)
			{
				hRgnPrev = CreateRectRgn(0,0,0,0);
				if (hRgnPrev)
					nResult = GetClipRgn(hDCDest, hRgnPrev);

				int nResult = SelectClipRgn(hDCDest, hRgnDest); 
				assert(ERROR != nResult);
			}
		}

		if (hDC != hDCDest)
			m_pBar->m_ffObj.Paint(hDCDest, rcPaint.left, rcPaint.top);

		if (hRgnDest)
		{
			if (-1 == nResult)
				SelectClipRgn(hDCDest, NULL); 
			else
				SelectClipRgn(hDCDest, hRgnPrev);
			
			BOOL bResult = DeleteRgn(hRgnDest);
			assert(bResult);
		}

		if (hRgnPrev)
		{
			bResult = DeleteRgn(hRgnPrev);
			assert(bResult);
		}

		ReleaseDC(hWnd, hDCDest);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// XPLookPopupWinInvalidateToolRect
//

void CBand::XPLookPopupWinInvalidateToolRect(HWND hWnd, CTool* pTool, const CRect& rcTool)
{
	try
	{
		if (!pTool->IsVisibleOnPaint())
			return;

		switch (pTool->tpV1.m_ttTools)
		{
		case ddTTControl:
		case ddTTForm:
		case ddTTSeparator:
			if (NULL == m_pBar->m_pDesigner)
				return;
		}
		BOOL bResult;
		CRect rcPaint(rcTool);

		HDC hDCDest = GetDC(hWnd);
		if (NULL == hDCDest)
			return;

		rcPaint.Inflate(eBevelBorder, eBevelBorder);
		CRect rcScreenBrush = rcPaint;
		CRect rcTotal = rcPaint;

		// For Menu bar tools I need to know how the tool is oriented, how the menu is flipped and if 
		// the sub menu is displayed

		BOOL bVertical = IsVertical();

		HDC hDC = m_pBar->m_ffObj.RequestDC(hDCDest, rcTotal.Width(), rcTotal.Height());
		if (NULL == hDC)
			hDC = hDCDest;
		else
			rcTotal.Offset(-rcTotal.left, -rcTotal.top);

		if (m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle)
		{
			//
			// Sliding Tabs
			//

			m_pParentBand->m_pChildBands->DrawToolSlidingTabBackground(hDC, 
																	   rcTotal, 
																	   m_prcTools[m_nToolMouseOver], 
																	   bVertical);
		}
		else
		{
			BOOL bTexture = m_pBar->HasTexture();
			switch (bpV1.m_btBands)
			{
			case ddBTPopup:
				if (bTexture)
					bTexture = ddBOPopups & m_pBar->bpV1.m_dwBackgroundOptions;
				break;

			case ddBTNormal:
				if (bTexture)
				{
					if (ddDAFloat == bpV1.m_daDockingArea)
						bTexture = ddBOFloat & m_pBar->bpV1.m_dwBackgroundOptions;
					else if (ddDAPopup == bpV1.m_daDockingArea)
						bTexture = ddBOPopups & m_pBar->bpV1.m_dwBackgroundOptions;
					else 
						bTexture = ddBODockAreas & m_pBar->bpV1.m_dwBackgroundOptions;
				}
				break;

			case ddBTMenuBar:
			case ddBTStatusBar:
			case ddBTChildMenuBar:
				if (bTexture)
				{
					if (ddDAFloat == bpV1.m_daDockingArea)
						bTexture = ddBOFloat & m_pBar->bpV1.m_dwBackgroundOptions;
					else
						bTexture = ddBODockAreas & m_pBar->bpV1.m_dwBackgroundOptions;
				}
				break;
			}
			
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			{
				if (!(ddBTChildMenuBar == bpV1.m_btBands || ddBTMenuBar == bpV1.m_btBands || ddBTStatusBar == bpV1.m_btBands))
					FillSolidRect(hDC, rcTotal, m_pBar->m_crXPBandBackground);
				else
					FillSolidRect(hDC, rcTotal, m_pBar->m_crXPBackground);
			}
			else
			{
				if (bTexture)
				{
					if (m_pBar->m_hBrushTexture)
						UnrealizeObject(m_pBar->m_hBrushTexture);

					SetBrushOrgEx(hDC, -rcScreenBrush.left, -rcScreenBrush.top, NULL);
					m_pBar->FillTexture(hDC, rcTotal);
					SetBrushOrgEx(hDC, 0, 0, NULL);
				}
				else
					FillSolidRect(hDC, rcTotal, m_pBar->m_crBackground);
			}
		}

		//
		// Draw border of tool
		//
		
		rcTotal.Inflate(-eBevelBorder, -eBevelBorder);
		POINT ptPaintOffset = {0 , 0};

		DrawTool(hDC, 
				 pTool, 
				 rcTotal, 
				 (ddBTChildMenuBar == bpV1.m_btBands || ddBTMenuBar == bpV1.m_btBands) && bVertical,
				 ptPaintOffset,
				 FALSE);

		HRGN hRgnDest = NULL;
		int nResult = -1;
		HRGN hRgnPrev = NULL;
		if (m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle)
		{
			CRect rcTmpCurrentPage = m_rcCurrentPage;
			rcTmpCurrentPage.Inflate(-1, -1);

			hRgnDest = CreateRectRgnIndirect(&rcTmpCurrentPage); 
			assert(hRgnDest);
			if (NULL == hRgnDest)
				return;

			if (BTabs::eInactiveBottomButton != m_pParentBand->m_pChildBands->m_sbsBottomStyle)
			{
				HRGN hRgnBottomButton = CreateRectRgnIndirect(&m_pParentBand->m_pChildBands->m_rcBottomButton); 
				assert(hRgnBottomButton);
				if (NULL == hRgnBottomButton)
				{
					DeleteRgn(hRgnDest);
					ReleaseDC(hWnd, hDCDest);
					return;
				}

				CombineRgn(hRgnDest, hRgnDest, hRgnBottomButton, RGN_DIFF);
			
				bResult = DeleteRgn(hRgnBottomButton);
				assert(bResult);
			}

			if (BTabs::eInactiveTopButton != m_pParentBand->m_pChildBands->m_sbsTopStyle)
			{
				HRGN hRgnTopButton = CreateRectRgnIndirect(&m_pParentBand->m_pChildBands->m_rcTopButton); 
				assert(hRgnTopButton);
				if (NULL == hRgnTopButton)
				{
					DeleteRgn(hRgnDest);
					ReleaseDC(hWnd, hDCDest);
					return;
				}

				CombineRgn(hRgnDest, hRgnDest, hRgnTopButton, RGN_DIFF);

				bResult = DeleteRgn(hRgnTopButton);
				assert(bResult);
			}

			if (hRgnDest)
			{
				hRgnPrev = CreateRectRgn(0,0,0,0);
				if (hRgnPrev)
					nResult = GetClipRgn(hDCDest, hRgnPrev);

				int nResult = SelectClipRgn(hDCDest, hRgnDest); 
				assert(ERROR != nResult);
			}
		}

		if (hDC != hDCDest)
			m_pBar->m_ffObj.Paint(hDCDest, rcPaint.left, rcPaint.top);

		if (hRgnDest)
		{
			if (-1 == nResult)
				SelectClipRgn(hDCDest, NULL); 
			else
				SelectClipRgn(hDCDest, hRgnPrev);
			
			BOOL bResult = DeleteRgn(hRgnDest);
			assert(bResult);
		}

		if (hRgnPrev)
		{
			bResult = DeleteRgn(hRgnPrev);
			assert(bResult);
		}

		ReleaseDC(hWnd, hDCDest);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

HWND CBand::GetDockHandle()
{
	if (NULL == m_pRootBand)
		return NULL;
	switch (m_pRootBand->bpV1.m_daDockingArea)
	{
	case ddDAFloat:
		if (NULL == m_pRootBand->m_pFloat)
			return NULL;
		return m_pRootBand->m_pFloat->hWnd();

	case ddDAPopup:
		if (NULL == m_pRootBand->m_pPopupWin)
			return FALSE;

		return m_pRootBand->m_pPopupWin->hWnd();

	default:
		if (NULL == m_pRootBand->m_pDock)
			return FALSE;
			
		return m_pRootBand->m_pDock->hWnd();
	}
	return NULL;
}

//
// GetBandRect
//

BOOL CBand::GetBandRect(HWND& hWnd, CRect& rcBand)
{
	try
	{
		hWnd = NULL;
		rcBand.SetEmpty();

		switch (m_pRootBand->bpV1.m_daDockingArea)
		{
		case ddDAFloat:
			if (NULL == m_pRootBand->m_pFloat)
				return FALSE;

			hWnd = m_pRootBand->m_pFloat->hWnd();
			if (!IsWindow(hWnd))
				return FALSE;

			m_pRootBand->m_pFloat->GetClientRect(rcBand);
			break;

		case ddDAPopup:
			if (NULL == m_pRootBand->m_pPopupWin)
				return FALSE;

			hWnd = m_pRootBand->m_pPopupWin->hWnd();
			if (!IsWindow(hWnd))
				return FALSE;

			m_pRootBand->m_pPopupWin->GetClientRect(rcBand);
			rcBand.Inflate(-3, -3);
			break;

		default:
			{
				if (NULL == m_pRootBand->m_pDock)
					return FALSE;
				
				hWnd = m_pRootBand->m_pDock->hWnd();
				if (!IsWindow(hWnd))
					return FALSE;

				// Now translate window relative to client
				CRect rcDock;
				m_pRootBand->m_pDock->GetClientRect(rcDock);
				rcBand = m_pRootBand->m_rcDock;
				rcBand.Offset(rcDock.left, rcDock.top);
			}
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
// GetToolBandRect
//

BOOL CBand::GetToolBandRect(int nToolIndex, CRect& rcTool, HWND& hWnd)
{
	try
	{
		hWnd = NULL;
		rcTool.SetEmpty();

		if (NULL == m_prcTools)
			return FALSE;

		int nToolCount = GetVisibleToolCount();
		if (nToolIndex >= nToolCount || nToolIndex < 0)
			return FALSE;

		CRect rcBase;
		if (!GetBandRect(hWnd, rcBase))
			return FALSE;

		SIZE  sizeOffset;
		sizeOffset.cx = rcBase.left; 
		sizeOffset.cy = rcBase.top;
		switch (m_pRootBand->bpV1.m_daDockingArea)
		{
		case ddDATop:
		case ddDABottom:
		case ddDALeft:
		case ddDARight:
		case ddDAFloat:
			{
				if (m_pParentBand)
				{
					if (ddCBSNone == m_pParentBand->bpV1.m_cbsChildStyle)
					{
						sizeOffset.cx += m_sizeEdgeOffset.cx; 
						sizeOffset.cy += m_sizeEdgeOffset.cy;
					}
					else
					{
						sizeOffset.cx += m_pParentBand->m_sizeEdgeOffset.cx - m_pParentBand->m_pChildBands->m_nHorzSlidingOffset;
						sizeOffset.cy += m_pParentBand->m_sizeEdgeOffset.cy - m_pParentBand->m_pChildBands->m_nVertSlidingOffset;
					}
				}
				else
				{
					sizeOffset.cx += m_sizeEdgeOffset.cx; 
					sizeOffset.cy += m_sizeEdgeOffset.cy;
				}
			}
			break;

		case ddDAPopup:
			sizeOffset.cx += m_sizeEdgeOffset.cx; 
			sizeOffset.cy += m_sizeEdgeOffset.cy;
			break;
		}
		rcTool = m_prcTools[nToolIndex];
		rcTool.Offset(sizeOffset.cx, sizeOffset.cy);
		if (ddBTStatusBar == m_pRootBand->bpV1.m_btBands)
			rcTool.Offset(0, eBevelBorder);

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
// GetToolScreenRect
//

BOOL CBand::GetToolScreenRect(int nToolIndex, CRect& rcTool)
{
	HWND hWndBand;
	if (!GetToolBandRect(nToolIndex, rcTool, hWndBand))
		return FALSE;

	ClientToScreen(hWndBand, rcTool);
	return TRUE;
}

//
// ToolNotification
//

void CBand::ToolNotification(CTool* pTool, TOOLNF tnf)
{
	try
	{
		if (ddBTPopup == bpV1.m_btBands || 0 != m_nInBandOpen)
			return;
		
		if (ddCBSNone == bpV1.m_cbsChildStyle)
		{
			int nTool = m_pTools->GetVisibleToolIndex(pTool);
			if (-1 != nTool)
				ToolNotificationEx(pTool, nTool, tnf);
		}
		else
		{
			CBand* pChildBand = m_pChildBands->GetCurrentChildBand();
			if (pChildBand)
				pChildBand->ToolNotification(pTool, tnf);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
}

//
// ToolNotificationEx
//

void CBand::ToolNotificationEx(CTool* pTool, int nTool, TOOLNF tnf)
{
	try
	{
		if (VARIANT_FALSE == bpV1.m_vbVisible || 0 != m_nInBandOpen)
			return;

		if (ddBTPopup == bpV1.m_btBands)
			return;
		
		if (VARIANT_FALSE == pTool->tpV1.m_vbVisible)
			return;

		HWND hWndParent;
		switch(tnf)
		{
		case TNF_VIEWCHANGED:
			{
				CRect rcTool;
				if (GetToolBandRect(nTool, rcTool, hWndParent))
				{
					m_nToolMouseOver = nTool;
					InvalidateToolRect(nTool);
				}
			}
			break;

		case TNF_ACTIVETOOLCHECK:
			{
				if (m_pBar->m_bToolModal)
					break;

				POINT pt;
				GetCursorPos(&pt);
				CRect rcTool;
				if (GetToolBandRect(nTool, rcTool, hWndParent))
				{
					POINT pt2 = pt;
					ScreenToClient(hWndParent, &pt2);
					rcTool.Inflate(eBevelBorder2, eBevelBorder2);
					if (!PtInRect(&rcTool, pt2) && m_pBar->m_pActiveTool)
					{
						m_pBar->SetActiveTool(NULL);
						KillTimer(m_pBar->m_hWnd, CBar::eFlyByTimer);
					}
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
}

//
// DoFloat
//

BOOL CBand::DoFloat(BOOL bDoFloat)
{
	TRACE(3, "Floating\n");
	if (bDoFloat)
	{
		if (m_pBar && !m_pBar->AmbientUserMode() && NULL == m_pBar->m_pDesigner)
			return TRUE;

		HWND hWndParent = m_pBar->GetParentWindow();
		if (NULL == hWndParent)
		{
			bpV1.m_vbVisible = VARIANT_FALSE;
			return FALSE;
		}

		m_pFloat = new CMiniWin(this);
		if (NULL == m_pFloat)
		{
			bpV1.m_vbVisible = VARIANT_FALSE;
			return FALSE;
		}

		if (!m_pFloat->CreateWin(hWndParent))
		{
			// Failed so recover
			bpV1.m_vbVisible = VARIANT_FALSE;
			delete m_pFloat;
			m_pFloat = NULL;
			return FALSE;
		}
		m_rcDock.SetEmpty();
		m_pDock = NULL;
		bpV1.m_nDockLine = 0;
		bpV1.m_nDockOffset = 0;
		m_pFloat->SetWindowPos(HWND_TOP, 
							   0, 
							   0, 
							   0, 
							   0,
							   (bDoFloat && VARIANT_TRUE == bpV1.m_vbVisible ? SWP_SHOWWINDOW : SWP_HIDEWINDOW) |SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
	}
	else
	{
		m_pFloat->DestroyWindow();
		m_pFloat = NULL;
		bpV1.m_vbVisible = VARIANT_FALSE;
	}
	return TRUE;
}

//
// HitTestTools
//

CTool* CBand::HitTestTools(CBand*& pBandHitTest, POINT pt, int& nToolIndex)
{
	try
	{
		nToolIndex = -1;
		CBand* pChildBand = NULL;
		if (ddCBSNone != bpV1.m_cbsChildStyle && m_pChildBands)
			pChildBand = m_pChildBands->GetCurrentChildBand();
		if (ddCBSNone == bpV1.m_cbsChildStyle || pChildBand == this)
		{
			if (m_prcTools)
			{
				CTool* pTool;
				CRect rcTool;
				int nToolCount = GetVisibleToolCount();
				for (int nTool = GetFirstTool(); nTool < nToolCount; nTool++)
				{
					pTool = m_pTools->GetVisibleTool(nTool);

					if (!pTool->IsVisibleOnPaint())
						continue;

					if (m_pBar->m_bAltCustomization || m_pBar->m_bCustomization)
					{
						if (CBar::eToolIdSysCommand == pTool->tpV1.m_nToolId)
							continue;

						if (CBar::eToolIdMDIButtons == pTool->tpV1.m_nToolId)
							continue;

						if (CTool::ddTTMoreTools == pTool->tpV1.m_ttTools)
							continue;
					}

					rcTool = m_prcTools[nTool];
					rcTool.Inflate(eBevelBorder, eBevelBorder); 
					if (m_pParentBand)
					{
						if (ddCBSNone == m_pParentBand->bpV1.m_cbsChildStyle)
							rcTool.Offset(m_sizeEdgeOffset.cx, m_sizeEdgeOffset.cy); 
						else
							rcTool.Offset(m_pParentBand->m_sizeEdgeOffset.cx - m_pParentBand->m_pChildBands->m_nHorzSlidingOffset, 
										  m_pParentBand->m_sizeEdgeOffset.cy - m_pParentBand->m_pChildBands->m_nVertSlidingOffset); 
					}
					else
						rcTool.Offset(m_sizeEdgeOffset.cx, m_sizeEdgeOffset.cy); 

					// Check for clipping of tool
					if (PtInRect(&rcTool, pt)) 
					{
						nToolIndex = nTool;
						pBandHitTest = this;
						return pTool;
					}
				}
			}
			else
			{
				nToolIndex = -1;
				pBandHitTest = this;
			}
		}
		else if (pChildBand)
			return pChildBand->HitTestTools(pBandHitTest, pt, nToolIndex);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
	return NULL;
}

//
// OnLButtonDown
//

BOOL CBand::OnLButtonDown(UINT nFlags, POINT pt)
{
	if (m_pBar->m_bToolModal)
		return TRUE;

	if (0x8000 & GetKeyState(VK_MENU))
		m_pBar->m_bAltCustomization = TRUE;

	if (m_pBar->m_bCustomization || m_pBar->m_bWhatsThisHelp || m_pBar->m_bAltCustomization)
	{
		OnCustomMouseDown(WM_LBUTTONDOWN, pt);
		return TRUE;
	}

	if (ddBFClose & bpV1.m_dwFlags && m_bCloseButtonVisible)
	{
		if (TrackCloseButton(pt))
			return TRUE;
	}

	if (ddBFExpand & bpV1.m_dwFlags && eGrayed != m_eExpandButtonState && m_bExpandButtonVisible)
	{
		if (TrackExpandButton(pt))
			return TRUE;
	}

	m_bChildBandChanging = FALSE;
	if (ddCBSNone != bpV1.m_cbsChildStyle)
	{
		//
		// Checking if a tab was clicked
		//
		
		int nHit;
		HRESULT hResult = DSGCalcHitEx(pt.x, pt.y, &nHit);
		if (SUCCEEDED(hResult))
		{
			if (-1 != nHit)
			{
				if (nHit != m_pChildBands->BTGetCurSel())
				{
					m_bChildBandChanging = TRUE;
					OnChildBandChanged(nHit);
					m_bChildBandChanging = FALSE;
				}
				if (ddCBSSlidingTabs == bpV1.m_cbsChildStyle)
					SetCursor(GetGlobals().GetHandCursor());
			}
			return TRUE;
		}
	}

	CBand* pBandHitTest;
	CRect rcTool;
	int nToolIndex;
	CTool* pTool = HitTestTools(pBandHitTest, pt, nToolIndex);
	if (pTool && VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
	{
		try
		{
			if (m_pBar->m_pActiveTool != pTool)
				m_pBar->SetActiveTool(pTool);

			pBandHitTest->m_nCurrentTool = nToolIndex;
			pTool->OnLButtonDown(nFlags, pt);

			if (pTool->HasSubBand() && (pTool->tpV1.m_ttTools != ddTTButtonDropDown || pTool->m_bDropDownPressed) && pTool->tpV1.m_ttTools != CTool::ddTTChildSysMenu)
			{
				if (CTool::ddTTMoreTools != pTool->tpV1.m_ttTools || pTool->EnableMoreTools())
				{
					pBandHitTest->EnterMenuLoop(nToolIndex, pTool, ENTERMENULOOP_CLICK, 0);
					return TRUE;
				}
			}

			pBandHitTest->m_nCurrentTool = -1;
			if (nToolIndex < pBandHitTest->GetVisibleToolCount())
				pBandHitTest->InvalidateToolRect(nToolIndex);
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}
	else if (m_pDockMgr && m_pRootBand && !(ddBFFixed & m_pRootBand->bpV1.m_dwFlags) && ddDAPopup != m_pRootBand->bpV1.m_daDockingArea)
	{
		if (ddCBSSlidingTabs == m_pRootBand->bpV1.m_cbsChildStyle)
		{
			//
			// Don't start drag if it is inside of the edges of a band and not on a tool
			// 

			CRect rcTemp = m_rcInsideBand;
			rcTemp.Offset(-m_pRootBand->m_rcDock.left, -m_pRootBand->m_rcDock.top);
			if (PtInRect(&rcTemp, pt))
				return TRUE;
		}

		HWND hWndCapture;
		CRect rcBand;
		if (GetBandRect(hWndCapture, rcBand))
		{
			POINT ptDragStart = pt;

			BOOL bStartDrag = FALSE;
			BOOL bProcessing = TRUE;
			SetCapture(hWndCapture);
			MSG msg;
			DWORD nDragStartTick = GetTickCount();
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				SetCursor(LoadCursor(NULL, IDC_SIZEALL));

			while (GetCapture() == hWndCapture && bProcessing)
			{
				GetMessage(&msg, NULL, 0, 0);
				switch (msg.message)
				{
				case WM_KEYDOWN:
					switch (msg.wParam)
					{
					case VK_ESCAPE:
						bProcessing = FALSE;
						break;

					case VK_CONTROL:
						bStartDrag = TRUE;
						bProcessing = FALSE;
						break;
					}
					break;

				case WM_RBUTTONUP:
				case WM_LBUTTONUP:
					bProcessing = FALSE;
					break;

				case WM_MOUSEMOVE:
					{
						POINT pt = {LOWORD(msg.lParam), HIWORD(msg.lParam)};
						if ((GetTickCount()-nDragStartTick) < GetGlobals().m_nDragDelay &&
							abs(pt.x-ptDragStart.x) < GetGlobals().m_nDragDist &&
							abs(pt.y-ptDragStart.y) < GetGlobals().m_nDragDist)
						{
							break;
						}
						bStartDrag = TRUE;
						bProcessing = FALSE;
					}
					break;

				default:
					// Just dispatch rest of the messages
					DispatchMessage(&msg);
					break;
				}
			}
			if (GetCapture() == hWndCapture)
				ReleaseCapture();

			if (bStartDrag)
				m_pDockMgr->StartDrag(this, pt);
		}
	}
	return TRUE;
}

//
// OnLButtonDblClk
//

BOOL CBand::OnLButtonDblClk(UINT nFlags, POINT pt)
{
	if (m_bChildBandChanging)
		return TRUE;

	if (ddBFClose & m_pRootBand->bpV1.m_dwFlags)
	{
		//
		// Close button
		//

		CRect rcTemp = m_rcCloseButton;
		rcTemp.Offset(-m_rcDock.left, -m_rcDock.top);
		if (PtInRect(&rcTemp, pt))
			return FALSE;
	}

	if (ddBFExpand & m_pRootBand->bpV1.m_dwFlags)
	{
		//
		// Expand button
		//

		CRect rcTemp = m_rcExpandButton;
		rcTemp.Offset(-m_rcDock.left, -m_rcDock.top);
		if (PtInRect(&rcTemp, pt))
			return FALSE;
	}

	if (m_pRootBand && ddCBSNone != m_pRootBand->bpV1.m_cbsChildStyle)
	{
		if (ddCBSSlidingTabs == m_pRootBand->bpV1.m_cbsChildStyle)
		{
			//
			// Checking if a scroll button was clicked
			//

			int nHit = m_pChildBands->ScrollButtonHit(pt);
			if (BTabs::eSlidingNone != nHit)
			{
				TrackSlidingTabScrollButton(pt, nHit);
				return TRUE;
			}
		}

		//
		// Checking if a tab was clicked
		//
		
		int nHit;
		DSGCalcHitEx(pt.x, pt.y, &nHit);
		if (-1 != nHit)
		{
			if (nHit != m_pChildBands->BTGetCurSel())
				OnChildBandChanged(nHit);
			return TRUE;
		}
	}

	int nToolIndex;
	CBand* pBandHitTest;
	CTool* pTool = HitTestTools(pBandHitTest, pt, nToolIndex);
	if (pTool && VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
	{
		if (m_pBar->m_bCustomization || m_pBar->m_bWhatsThisHelp || m_pBar->m_bAltCustomization)
		{
			OnCustomMouseDown(WM_LBUTTONDOWN, pt);
			return TRUE;
		}

		//
		// Was a tool double clicked on?
		//

		if (VARIANT_TRUE == pTool->tpV1.m_vbEnabled && ddTTCombobox != pTool->tpV1.m_ttTools && ddTTEdit != pTool->tpV1.m_ttTools)
		{
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbFireDblClickEvent && (ddBTMenuBar != m_pRootBand->bpV1.m_btBands || ddBTChildMenuBar != m_pRootBand->bpV1.m_btBands))
				m_pBar->FireToolDblClick(reinterpret_cast<Tool*>(pTool));
			else
			{
				try
				{
					pBandHitTest->m_nCurrentTool = nToolIndex;
					pTool->OnLButtonDown(nFlags, pt);

					if (pTool->HasSubBand() && (pTool->tpV1.m_ttTools != ddTTButtonDropDown || pTool->m_bDropDownPressed) && pTool->tpV1.m_ttTools != CTool::ddTTChildSysMenu)
					{
						if (CTool::ddTTMoreTools != pTool->tpV1.m_ttTools || pTool->EnableMoreTools())
						{
							pBandHitTest->EnterMenuLoop(nToolIndex, pTool, ENTERMENULOOP_CLICK, 0);
							return TRUE;
						}
					}

					pBandHitTest->m_nCurrentTool = -1;
					if (nToolIndex < pBandHitTest->GetVisibleToolCount())
						pBandHitTest->InvalidateToolRect(nToolIndex);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
			return TRUE;
		}
	}

	if (!(ddBFFixed & m_pRootBand->bpV1.m_dwFlags) && ddBFFloat & m_pRootBand->bpV1.m_dwFlags && ddDAPopup != m_pRootBand->bpV1.m_daDockingArea)
	{
		//
		// Don't float the band if you double clicked on the inside of the band 
		//

		CRect rcTemp = m_rcInsideBand;
		rcTemp.Offset(-m_pRootBand->m_rcDock.left, -m_pRootBand->m_rcDock.top);
		if (PtInRect(&rcTemp, pt))
			return TRUE;

		try
		{
			//
			// Float the band
			//

			m_daPrevDockingArea = m_pRootBand->bpV1.m_daDockingArea;
			m_nPrevDockOffset = m_pRootBand->bpV1.m_nDockOffset;
			m_nPrevDockLine = m_pRootBand->bpV1.m_nDockLine;
			m_pRootBand->bpV1.m_daDockingArea = ddDAFloat;
			m_pBar->FireBandUndock((Band*)m_pRootBand);
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}			
		m_pBar->RecalcLayout();
	}
	return TRUE;
}

//
// OnLButtonUp
//

BOOL CBand::OnLButtonUp(UINT nFlags, POINT pt)
{
	if (m_pBar->m_bToolModal)
		return FALSE;

	if (m_pBar->m_bAltCustomization)
		m_pBar->m_bAltCustomization = FALSE;

	if (m_pBar->m_bCustomization)
	{
		OnCustomMouseDown(WM_LBUTTONUP, pt);
		return TRUE;
	}

	if (ddCBSNone != bpV1.m_cbsChildStyle)
	{
		//
		// Checking if the mouse is over a tab
		//
		
		if (ddCBSSlidingTabs == bpV1.m_cbsChildStyle && m_pChildBands)
		{
			CRect rcTab = m_rcCurrentPaint;
			rcTab.Offset(-rcTab.left, -rcTab.top);
			if (PtInRect(&rcTab, pt))
			{
				AdjustChildBandsRect(rcTab);
				int nHit = m_pChildBands->BTCalcHit(pt, rcTab);
				if (-1 != nHit)
				{
					SetCursor(GetGlobals().GetHandCursor());
					return TRUE;
				}
			}
		}
	}

	int nToolIndex;
	CBand* pBandHitTest;
	CTool* pTool = HitTestTools(pBandHitTest, pt, nToolIndex);
	if (pTool && VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
	{
		pBandHitTest->m_nCurrentTool = nToolIndex;
		pTool->OnLButtonUp(nFlags, pt);
		pBandHitTest->m_nCurrentTool = -1;
		pBandHitTest->InvalidateToolRect(nToolIndex);
		return TRUE;
	}
	return FALSE;
}

//
// OnRButtonDown
//

BOOL CBand::OnRButtonDown(UINT nFlags, POINT pt)
{
	if (m_pBar->m_bCustomization || m_pBar->m_bWhatsThisHelp || m_pBar->m_bAltCustomization)
	{
		OnCustomMouseDown(WM_RBUTTONDOWN, pt);
		return TRUE;
	}
	return FALSE;
}

//
// OnRButtonUp
//

BOOL CBand::OnRButtonUp(UINT nFlags, POINT pt)
{
	if (!m_pBar->m_bCustomization && !m_pBar->m_bWhatsThisHelp)
		m_pBar->OnBarContextMenu();
	
	if (m_pBar->m_bCustomization)
	{
		OnCustomMouseDown(WM_RBUTTONUP, pt);
		return TRUE;
	}

	return FALSE;
}

//
// OnMouseMove
//

BOOL CBand::OnMouseMove(UINT nFlags, POINT pt)
{
	try
	{
		if (m_pBar->m_bToolModal)
			return TRUE;

		if (m_pBar->m_bCustomization)
		{
			OnCustomMouseDown(WM_MOUSEMOVE, pt);
			return TRUE;
		}

		try
		{
			if (ddCBSNone != bpV1.m_cbsChildStyle)
			{
				//
				// Checking if the mouse is over a tab
				//
				
				if (ddCBSSlidingTabs == bpV1.m_cbsChildStyle && m_pChildBands)
				{
					CRect rcTab = m_rcCurrentPaint;
					rcTab.Offset(-rcTab.left, -rcTab.top);
					if (PtInRect(&rcTab, pt))
					{
						AdjustChildBandsRect(rcTab);
						try
						{
							int nHit = m_pChildBands->BTCalcHit(pt, rcTab);
							if (-1 != nHit)
							{
								SetCursor(GetGlobals().GetHandCursor());
								return TRUE;
							}
						}
						catch (...)
						{
							assert(FALSE);
						}
					}
				}
			}
		}
		catch (...)
		{
			assert(FALSE);
		}

		try
		{
			int nToolIndex;
			CBand* pBandHitTest;
			CTool* pTool = HitTestTools(pBandHitTest, pt, nToolIndex);
			if (pTool)
			{
				try
				{
					if (ddTTCombobox == pTool->tpV1.m_ttTools && !pTool->IsComboReadOnly() && VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
					{
						CRect rcTool;
						if (GetToolRect(nToolIndex, rcTool))
						{
							rcTool.Offset(m_sizeEdgeOffset.cx, m_sizeEdgeOffset.cy);
							rcTool.right -= 10;
							if (PtInRect(&rcTool, pt))
								SetCursor(LoadCursor(NULL, IDC_IBEAM));
							else
								SetCursor(LoadCursor(NULL, IDC_ARROW));
						}
						else
							SetCursor(LoadCursor(NULL, IDC_ARROW));
					}
					else if (ddTTEdit == pTool->tpV1.m_ttTools && VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
						SetCursor(LoadCursor(NULL, IDC_IBEAM));
					else
						SetCursor(LoadCursor(NULL, IDC_ARROW));

					pBandHitTest->m_nCurrentTool = nToolIndex;
					if ((ddBTChildMenuBar == bpV1.m_btBands || ddBTMenuBar == bpV1.m_btBands) && CBar::eToolIdMDIButtons == pTool->tpV1.m_nToolId)
					{
						try
						{
							//
							// If we are an MDI Project and the MDI Buttons are present Hittest for which MDI Button 
							// the mouse is over
							//

							if (pBandHitTest->m_prcTools)
							{
								CRect rcTool = pBandHitTest->m_prcTools[nToolIndex];
								rcTool.Offset(pBandHitTest->m_sizeEdgeOffset.cx, pBandHitTest->m_sizeEdgeOffset.cy); 

								int nButton;
								if (pTool->GetMDIButton(nButton, rcTool, pt, ddDALeft == m_pRootBand->bpV1.m_daDockingArea || ddDARight == m_pRootBand->bpV1.m_daDockingArea))
								{
									//
									// Found which buttom the mouse is over show the tooltip
									//

									CTool::MDIButtons mdibTemp = CTool::eNone;
									DDString strToolTip;
									switch (nButton)
									{
									case CBar::eMinimizeWindowPosition:
										strToolTip = m_pBar->Localizer()->GetString(ddLTMinimizeButton);
										mdibTemp = CTool::eMinimize;
										break;

									case CBar::eRestoreWindowPosition:
										strToolTip = m_pBar->Localizer()->GetString(ddLTRestoreButton);
										mdibTemp = CTool::eRestore;
										break;

									case CBar::eCloseWindowPosition:
										strToolTip = m_pBar->Localizer()->GetString(ddLTCloseWindowButton);
										mdibTemp = CTool::eClose;
										break;
									}
									
									if (pTool->m_mdibActive != mdibTemp)
									{
										pTool->m_mdibActive = mdibTemp;
										InvalidateToolRect(nToolIndex);
									}

									HWND hWndBand;
									if (pBandHitTest->GetToolBandRect(nToolIndex, rcTool, hWndBand))
									{
										BSTR bstrToolTipMsg = strToolTip.AllocSysString();
										if (bstrToolTipMsg)
										{
											ClientToScreen(hWndBand, rcTool);
											rcTool.Inflate(2, 2);
											m_pBar->ShowToolTip((DWORD)pTool + nButton, 
																bstrToolTipMsg, 
																rcTool, 
																ddDAPopup == m_pRootBand->bpV1.m_daDockingArea);
											SysFreeString(bstrToolTipMsg);
										}
									}
									pBandHitTest->m_nCurrentTool = -1;
								}
							}
						}
						catch (...)
						{
							assert(FALSE);
						}
					}
				}
				catch (...)
				{
					assert(FALSE);
				}

				if (m_pBar->m_pActiveTool == pTool)
				{
					// We are over the same tool so we don't have to worry about the tool tip
					return TRUE;
				}

				try
				{
					KillTimer(m_pBar->m_hWnd, CBar::eFlyByTimer);
					if (VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
					{
						//
						// Now start check timer for area leave
						//

						SetTimer(m_pBar->m_hWnd,
								 CBar::eFlyByTimer,
								 CBar::eFlyByTimerDuration,
								 0);

						m_pBar->SetActiveTool(pTool);
					}
					else if (m_pBar->m_pActiveTool)
						m_pBar->SetActiveTool(NULL);
				}
				catch (...)
				{
					assert(FALSE);
				}

				if ((ddBTNormal == bpV1.m_btBands || ddBTStatusBar == bpV1.m_btBands) && pTool->m_bstrToolTipText)
				{
					try
					{
						//
						// Display ToolTip
						//

						CRect rcTool;
						BOOL bResult = pBandHitTest->GetToolScreenRect(nToolIndex, rcTool);
						assert(bResult);
						if (bResult)
						{
							rcTool.Inflate(eBevelBorder, eBevelBorder);

							BSTR bstrToolTipMsg = SysAllocString(pTool->m_bstrToolTipText);
							if (bstrToolTipMsg)
							{
								LPTSTR szShortCut = pTool->GetShortCutString();
								if (m_pBar->bpV1.m_vbDisplayKeysInToolTip && szShortCut && lstrlen(szShortCut) > 0)
								{
									BSTRAppend(&bstrToolTipMsg, L" (");
									MAKE_WIDEPTR_FROMTCHAR(wStr, szShortCut);
									BSTRAppend(&bstrToolTipMsg, wStr);
									BSTRAppend(&bstrToolTipMsg, L")");
								}
								m_pBar->ShowToolTip((DWORD)pTool, 
													bstrToolTipMsg, 
													rcTool, 
													ddDAPopup == m_pRootBand->bpV1.m_daDockingArea);
								SysFreeString(bstrToolTipMsg);
							}
						}
						pBandHitTest->m_nCurrentTool = -1;
					}
					catch (...)
					{
						assert(FALSE);
					}
				}
				return TRUE;
			}

			HWND hWnd;
			CRect rcBand;
			GetBandRect(hWnd, rcBand);
			CRect rcGrabHandle = m_rcGrabHandle;
			rcGrabHandle.Offset(-rcBand.left, -rcBand.top);
			if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook && PtInRect(&rcGrabHandle, pt) && ddBTStatusBar != bpV1.m_btBands)
				SetCursor(LoadCursor(NULL, IDC_SIZEALL));
			else
				SetCursor(LoadCursor(NULL, IDC_ARROW));
		}
		catch (...)
		{
			assert(FALSE);
		}
		m_pBar->HideToolTips(FALSE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
	return FALSE;
}

//
// DSGCalcHitEx
//

STDMETHODIMP CBand::DSGCalcHitEx(int x, int y, int* pHit)
{
	POINT pt = {x, y};

	BOOL bVert = ddDALeft == m_pRootBand->bpV1.m_daDockingArea || ddDARight == m_pRootBand->bpV1.m_daDockingArea || ddDAFloat == m_pRootBand->bpV1.m_daDockingArea;
	CRect rcTab = m_rcCurrentPaint;
	rcTab.Offset(-rcTab.left, -rcTab.top);

	*pHit = -1;

	switch (bpV1.m_cbsChildStyle)
	{
	case ddCBSToolbarTopTabs:
		{
			if (bVert)
				rcTab.right = rcTab.left + m_rcChildBandArea.left;
			else
				rcTab.bottom = rcTab.top + m_rcChildBandArea.top;

			if (IsGrabHandle())
			{
				if (bVert)
					rcTab.top += m_nAdjustForGrabBar;
				else
					rcTab.left += m_nAdjustForGrabBar;
			}
			if (PtInRect(&rcTab, pt))
				*pHit = m_pChildBands->BTCalcHit(pt, rcTab);
		}
		break;

	case ddCBSToolbarBottomTabs:
		{
			if (bVert)
				rcTab.left = rcTab.right - m_rcChildBandArea.right;
			else
				rcTab.top = rcTab.bottom - m_rcChildBandArea.bottom;

			if (IsGrabHandle())
			{
				if (bVert)
					rcTab.top += m_nAdjustForGrabBar;
				else
					rcTab.left += m_nAdjustForGrabBar;
			}
			if (PtInRect(&rcTab, pt))
				*pHit = m_pChildBands->BTCalcHit(pt, rcTab);
		}
		break;

	case ddCBSSlidingTabs:
		{
			int nHit = m_pChildBands->ScrollButtonHit(pt);
			if (BTabs::eSlidingNone != nHit)
			{
				TrackSlidingTabScrollButton(pt, nHit);
				return NOERROR;
			}
			if (PtInRect(&rcTab, pt))
			{
				AdjustChildBandsRect(rcTab);
				*pHit = m_pChildBands->BTCalcHit(pt, rcTab);
			}
		}
		break;

	case ddCBSTopTabs:
		rcTab.bottom = rcTab.top + m_rcChildBandArea.top;

		if (IsGrabHandle())
		{
			if (bVert)
				rcTab.Offset(0, m_nAdjustForGrabBar);
			else
				rcTab.Offset(m_nAdjustForGrabBar, 0);
		}
		if (PtInRect(&rcTab, pt))
			*pHit = m_pChildBands->BTCalcHit(pt, rcTab);
		break;

	case ddCBSBottomTabs:
		rcTab.top = rcTab.bottom - m_rcChildBandArea.bottom;
		if (!bVert && IsGrabHandle())
			rcTab.Offset(m_nAdjustForGrabBar, 0);
		if (PtInRect(&rcTab, pt))
			*pHit = m_pChildBands->BTCalcHit(pt, rcTab);
		break;

	case ddCBSLeftTabs:
		rcTab.right = rcTab.left + m_rcChildBandArea.left;

		if (IsGrabHandle())
		{
			if (bVert)
				rcTab.Offset(0, m_nAdjustForGrabBar);
			else
				rcTab.Offset(m_nAdjustForGrabBar, 0);
		}
		if (PtInRect(&rcTab, pt))
			*pHit = m_pChildBands->BTCalcHit(pt, rcTab);
		break;

	case ddCBSRightTabs:
		rcTab.left = rcTab.right - m_rcChildBandArea.right;
		if (bVert && IsGrabHandle())
			rcTab.Offset(0, m_nAdjustForGrabBar);
		if (PtInRect(&rcTab, pt))
			*pHit = m_pChildBands->BTCalcHit(pt, rcTab);
		break;
	}
	if (-1 != *pHit)
		return NOERROR;
	return E_FAIL;
}
/*
//
// OnMouseMove
//

void CBand::OnMouseMove(CTool* pTool, POINT pt)
{
	try
	{
		assert(FALSE); // does anyone call this function anymore????

		if (m_pBar->m_pActiveTool == pTool || m_pBar->m_bToolModal)
			return;

		KillTimer(m_pBar->m_hWnd, CBar::eFlyByTimer);

		if (VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
		{
			m_pBar->SetActiveTool(pTool);

			//
			// Now start check timer for area leave
			//

			SetTimer(m_pBar->m_hWnd,
					 CBar::eFlyByTimer,
					 CBar::eFlyByTimerDuration,
					 NULL);
		}
		else if (m_pBar->m_pActiveTool)
			m_pBar->SetActiveTool(NULL);
		else
			m_pBar->FireMouseEnter((Tool*)pTool);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
}
*/
//
// CalcSqueezedSize
// 
// Min Width = Grabhandles + height or width of first tool + outer border space
//

SIZE CBand::CalcSqueezedSize() 
{
	SIZE s;
	BOOL bVertical = IsVertical();
	BOOL bHorz = !bVertical;
	BOOL bLeftHandle = bHorz && IsGrabHandle();
	BOOL bTopHandle = !bHorz && IsGrabHandle();

	int nVisibleToolCount = GetVisibleToolCount();
	if (0 == nVisibleToolCount)
	{
		s.cx = eEmptyBarWidth + (bLeftHandle ? eGrabTotalWidth : 0);
		s.cy = eEmptyBarHeight + (bTopHandle ? eGrabTotalWidth : 0);
		return s;
	}

	// Get first visible tool
	CTool* pTool = m_pTools->GetVisibleTool(0);

	// ddBTNormal since this is a docked area
	SIZE sizeTool;
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		sizeTool = pTool->CalcSize(hDC, ddBTNormal, bVertical); 
		ReleaseDC(NULL, hDC);
	}

	if (-1 != pTool->tpV1.m_nWidth)
		sizeTool.cx = pTool->tpV1.m_nWidth + 2 * eHToolPadding;

	if (-1 != pTool->tpV1.m_nHeight)
		sizeTool.cy = pTool->tpV1.m_nHeight + 2 * eVToolPadding;
	
	s.cx = sizeTool.cx + 2 * eBevelBorder2 + (bLeftHandle ? eGrabTotalWidth + eBevelBorder2 : 0);
	s.cy = sizeTool.cy + 2 * eBevelBorder2 + (bTopHandle ? eGrabTotalWidth + eBevelBorder2 : 0);

	if (1 != nVisibleToolCount)
	{
		if (bHorz)
			s.cx += eClipMarkerWidth + 3;
		else
			s.cy += eClipMarkerWidth + 3;
	}
	return s;
}

//
// GetOptimalFloatRect
//

CRect CBand::GetOptimalFloatRect(BOOL bCommit)
{
	CRect rcReturn;
	if (ddCBSNone != bpV1.m_cbsChildStyle)
	{

		CRect rcInit;
		rcReturn.SetEmpty();
		if (0 == bpV1.m_rcFloat.Width() || 0 == bpV1.m_rcFloat.Height())
		{
			if (-1 != bpV1.m_rcDimension.right || -1 != bpV1.m_rcDimension.bottom)
				rcInit = bpV1.m_rcDimension;
			else
				rcInit.right = 32767;

			// Set new layout
			CalcLayout(rcInit, eLayoutFloat | (ddCBSSlidingTabs != bpV1.m_cbsChildStyle ? eLayoutHorz : eLayoutVert), rcReturn, bCommit);
			rcInit.SetEmpty();
			CMiniWin::AdjustWindowRectEx(rcInit, FALSE);
			rcReturn.bottom += rcInit.Height();
			rcReturn.right += rcInit.Width();
		}
		else
		{
			rcInit = bpV1.m_rcFloat;
			rcReturn.SetEmpty();
			CMiniWin::AdjustWindowRectEx(rcReturn, FALSE);
			rcInit.left -= rcReturn.left;
			rcInit.top -= rcReturn.top;
			rcInit.right -= rcReturn.right;
			rcInit.bottom -= rcReturn.bottom;
			CalcLayout(rcInit, eLayoutFloat | (ddCBSSlidingTabs != bpV1.m_cbsChildStyle ? eLayoutHorz : eLayoutVert), rcReturn, bCommit);
		}
	}
	else
	{
		CRect rcInit;
		if (0 == bpV1.m_rcFloat.Width() || 0 == bpV1.m_rcFloat.Height())
		{
			if (-1 == bpV1.m_rcDimension.right || -1 == bpV1.m_rcDimension.bottom)
				rcInit.right = 32767;
			else
			{
				rcInit.right = bpV1.m_rcDimension.right;
				rcInit.bottom = bpV1.m_rcDimension.bottom;
				rcReturn.SetEmpty();

				CMiniWin::AdjustWindowRectEx(rcReturn, FALSE);

				rcInit.left -= rcReturn.left;
				rcInit.top -= rcReturn.top;
				rcInit.right -= rcReturn.right;
				rcInit.bottom -= rcReturn.bottom;
			}

			CalcLayout(rcInit, eLayoutFloat | eLayoutHorz, rcReturn, bCommit);
			
			// Now center to screen
			HWND hWndDock = m_pBar->GetDockWindow();
			CRect rcWin;
			GetWindowRect(hWndDock, &rcWin);
			if (0 != bpV1.m_rcFloat.left || 0 != bpV1.m_rcFloat.top)
				rcReturn.Offset(bpV1.m_rcFloat.left, bpV1.m_rcFloat.top);
			else
				// Now center to screen
				rcReturn.Offset((rcWin.Width() - rcReturn.Height()) / 2, (rcWin.Height() - rcReturn.Height()) / 2);
		}
		else
		{
			rcInit = bpV1.m_rcFloat;
			
			rcReturn.SetEmpty();

			CMiniWin::AdjustWindowRectEx(rcReturn, FALSE);
			
			rcInit.left -= rcReturn.left;
			rcInit.top -= rcReturn.top;
			rcInit.right -= rcReturn.right;
			rcInit.bottom -= rcReturn.bottom;
			
			CalcLayout(rcInit, eLayoutFloat|eLayoutHorz, rcReturn, bCommit);
		}
		CMiniWin::AdjustWindowRectEx(rcReturn, FALSE);
	}
	return rcReturn;
}

//
// Clone
//

STDMETHODIMP CBand::Clone(IBand** ppBand)
{
	if (NULL == ppBand)
		return E_INVALIDARG;

	CBand* pNewBand = CBand::CreateInstance(NULL);
	*ppBand = (IBand*)pNewBand;
	if (NULL == pNewBand)
		return E_OUTOFMEMORY;

	pNewBand->SetOwner(m_pBar);
	pNewBand->SetParent(m_pParentBand);
	pNewBand->m_pRootBand = this;
	return CopyTo(pNewBand);
}

//
// CopyTo
//

STDMETHODIMP CBand::CopyTo(IBand* pDest)
{
	if (NULL == pDest)
		return E_INVALIDARG;

	CBand* pBand = (CBand*)pDest;

	if (pBand == this)
		return NOERROR;

	SysFreeString(pBand->m_bstrCaption);
	pBand->m_bstrCaption = SysAllocString(m_bstrCaption);
	if (NULL == pBand->m_bstrCaption)
		return E_OUTOFMEMORY;

	SysFreeString(pBand->m_bstrName);
	pBand->m_bstrName = SysAllocString(m_bstrName);
	if (NULL == pBand->m_bstrName)
		return E_OUTOFMEMORY;

	memcpy (&pBand->bpV1, &bpV1, sizeof(bpV1));

	HRESULT hResult = VariantCopy(&pBand->m_vTag, &m_vTag);
	if (FAILED(hResult))
		return hResult;

	// Now copy all tools to destination
	if (m_pChildBands && pBand->m_pChildBands)
	{
		hResult = m_pChildBands->CopyTo(pBand->m_pChildBands);
		if (FAILED(hResult))
			return hResult;
	}
	hResult = m_pTools->CopyTo(pBand->m_pTools);
	return NOERROR;
}

//
// DSGCalcDropIndex
// 
//
STDMETHODIMP CBand::DSGCalcDropIndex(int x, int y, int* pnDropIndex, int* pnDirection) 
{
	try
	{
		if (ddCBSNone != bpV1.m_cbsChildStyle && m_pChildBands)
		{
			CBand* pChildBand = m_pChildBands->GetCurrentChildBand();
			if (pChildBand)
				return pChildBand->DSGCalcDropIndex(x, y, pnDropIndex, pnDirection);
		}

		if (NULL == m_prcTools)
		{
			*pnDropIndex = 0;
			*pnDirection = 0; 
			return E_FAIL;
		}

		int nToolCount = m_pTools->GetVisibleToolCount();
		if (0 == nToolCount)
		{
			*pnDropIndex = 0;
			*pnDirection = 0; 
			return NOERROR;
		}

		*pnDropIndex = -1;
		BOOL bAllowVertDropArea = (ddDALeft == m_pRootBand->bpV1.m_daDockingArea || ddDARight == m_pRootBand->bpV1.m_daDockingArea);
		BOOL bForceVertical = ddBTPopup == bpV1.m_btBands;
		if (bAllowVertDropArea)
			bForceVertical = TRUE;

		CRect rcTool;
		CTool* pTool;
		for (int nTool = 0; nTool < nToolCount; nTool++)
		{
			try
			{
				pTool = m_pTools->GetVisibleTool(nTool);

				if (pTool->tpV1.m_nToolId == CBar::eToolIdSysCommand)
					continue;

				if (pTool->tpV1.m_nToolId == CBar::eToolIdMDIButtons)
					continue;

				if (CTool::ddTTMoreTools == pTool->tpV1.m_ttTools)
					continue;

				rcTool = m_prcTools[nTool];
				rcTool.Inflate(eBevelBorder, eBevelBorder);
				if (m_pParentBand)
					rcTool.Offset(m_pParentBand->m_sizeEdgeOffset.cx, m_pParentBand->m_sizeEdgeOffset.cy); 
				else
					rcTool.Offset(m_sizeEdgeOffset.cx, m_sizeEdgeOffset.cy); 

				if (bForceVertical)
				{
					if (y > rcTool.top && y < rcTool.bottom)
					{
						if (y < ((rcTool.top + rcTool.bottom) / 2)) 
						{
							// left than midpoint?
							*pnDirection = eDropTop;
						}
						else
							*pnDirection = eDropBottom;
						*pnDropIndex = nTool;
						return NOERROR;
					}
					if (nTool == (nToolCount-1))
					{
						// nToolCount;
						*pnDropIndex = -1;
						return NOERROR;
					}
				}
				else
				{
					if (x >= rcTool.left && x <= rcTool.right && 
						y >= rcTool.top && y <= rcTool.bottom)
					{
						if (x < ((rcTool.left + rcTool.right)/2)) 
							*pnDirection = eDropLeft;
						else
							*pnDirection = eDropRight;
						*pnDropIndex = nTool;
						return NOERROR;
					}
					if (nTool == (nToolCount - 1) && 
						y >= rcTool.top && 
						y <= rcTool.bottom && 
						x >= rcTool.left)
					{
						// last tool ?
						*pnDropIndex = nToolCount;
						*pnDirection = eDropLeft;
						return NOERROR;
					}
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
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}			
	return NOERROR;
}

//
// DSGSetDropLoc
//

STDMETHODIMP CBand::DSGSetDropLoc(OLE_HANDLE hDC, 
								  int x, 
								  int y, 
								  int w, 
								  int h, 
								  int nDropIndex, 
								  int nDirection)
{
	CRect rcLocation(x, x + w, y, y + h);
	 // delete old one
	if (-1 != m_nDropLoc)
		DrawDropSign((OLE_HANDLE)hDC, 
					 rcLocation.left, 
					 rcLocation.top, 
					 rcLocation.Width(), 
					 rcLocation.Height(),
					 m_nDropLoc,
					 m_nDropLocDir);

	m_nDropLoc = nDropIndex;
	m_nDropLocDir = nDirection;
	
	if (-1 != m_nDropLoc)
		DrawDropSign((OLE_HANDLE)hDC, 
					 rcLocation.left, 
					 rcLocation.top, 
					 rcLocation.Width(), 
					 rcLocation.Height(),
					 m_nDropLoc,
					 m_nDropLocDir);
	return NOERROR;
}

//
// DrawDropSign
//

STDMETHODIMP CBand::DrawDropSign(OLE_HANDLE hDCScr, 
								 int		x, 
								 int		y, 
								 int		w, 
								 int		h, 
								 int		nDropIndex, 
								 int		nDirection)
{
	int nVisibleToolCount = GetVisibleToolCount();

	int nLastVisibleTool = nVisibleToolCount - 1;
	
	BOOL bVerticalDropSign = nDirection < eDropTop;
	
	if (eDropRight == nDirection || eDropBottom == nDirection)
		nDropIndex++;

	if (NULL == m_prcTools)
		return NOERROR;

	CRect rcTool;
	if (bVerticalDropSign) 
	{
		if (0 == nVisibleToolCount)
		{
			rcTool.top = 3;
			rcTool.bottom = rcTool.top + h - 6;
		}
		else if (nDropIndex <= nLastVisibleTool)
			rcTool = m_prcTools[nDropIndex];
		else
		{
			rcTool = m_prcTools[nLastVisibleTool];
			rcTool.left = rcTool.right;
		}
		if (ddBTMenuBar != m_pRootBand->bpV1.m_btBands)
		{
			rcTool.right = rcTool.left + eDropSignSize / 2;
			rcTool.left = rcTool.right - eDropSignSize;
		}
		else
		{
			rcTool.right = rcTool.left + eDropSignSize;
			rcTool.Inflate(0, -2);
		}
	}
	else
	{
		if (0 == nVisibleToolCount)
			rcTool.right = rcTool.left + eEmptyBarWidth;
		else if (nDropIndex <= nLastVisibleTool)
			rcTool = m_prcTools[nDropIndex];
		else
		{
			rcTool = m_prcTools[nLastVisibleTool];
			rcTool.top = rcTool.bottom;
		}
		if (0 == nDropIndex)
		{
			rcTool.top++;
			rcTool.bottom++;
		}
		else if (nDropIndex == nLastVisibleTool)
		{
			rcTool.top--;
			rcTool.bottom--;
		}
		if (ddBTMenuBar != m_pRootBand->bpV1.m_btBands)
		{
			rcTool.bottom = rcTool.top + eDropSignSize / 2;
			rcTool.top = rcTool.bottom - eDropSignSize;
		}
		else
		{
			rcTool.bottom = rcTool.top + eDropSignSize;
			rcTool.Inflate(-2, 0);
		}
	}

	rcTool.Offset(x, y);
	rcTool.Inflate(eBevelBorder, eBevelBorder);
	if (m_pParentBand)
		rcTool.Offset(m_pParentBand->m_sizeEdgeOffset.cx, m_pParentBand->m_sizeEdgeOffset.cy); 
	else
		rcTool.Offset(m_sizeEdgeOffset.cx, m_sizeEdgeOffset.cy); 

	HBRUSH hBrush = CreateHalftoneBrush(TRUE);
	if (NULL == hBrush)
		return E_FAIL;

	CRect rcBound = rcTool;
	HDC hDC = (HDC)hDCScr;
	if (hDC)
	{
		BOOL bResult;
		HBRUSH hBrushOld = SelectBrush(hDC, hBrush);
		COLORREF crTextOld = SetTextColor(hDC, m_pBar->m_crBackground);
		COLORREF crBackOld = SetBkColor(hDC, m_pBar->m_crBackground);

		if (bVerticalDropSign)
		{
			rcTool.bottom = rcTool.top + 2;
			bResult = PatBlt(hDC,
							 rcTool.left, 
							 rcTool.top, 
							 rcTool.Width(), 
							 rcTool.Height(), 
							 PATINVERT);

			rcTool = rcBound;
			rcTool.top = rcTool.bottom - 2;
			bResult = PatBlt(hDC,
						     rcTool.left, 
						     rcTool.top, 
						     rcTool.Width(), 
						     rcTool.Height(), 
						     PATINVERT);
			assert(bResult);

			rcTool = rcBound;
			rcTool.left += (rcTool.Width() - 2) / 2;
			rcTool.right = rcTool.left + 2;
			rcTool.top += 2;
			rcTool.bottom -= 2;
			bResult = PatBlt(hDC,
							 rcTool.left, 
							 rcTool.top, 
							 rcTool.Width(), 
							 rcTool.Height(), 
							 PATINVERT);
			assert(bResult);
		}
		else
		{
			rcTool.left += 2;
			rcTool.right -= 2;
			rcTool.top += eDropSignSize / 2;
			rcTool.bottom = rcTool.top + 2;
			bResult = PatBlt(hDC,
						     rcTool.left, 
						     rcTool.top, 
						     rcTool.Width(), 
						     rcTool.Height(), 
						     PATINVERT);
			assert(bResult);

			rcTool = rcBound;
			rcTool.right = rcTool.left + 2;
			bResult = PatBlt(hDC,
						     rcTool.left, 
						     rcTool.top, 
						     rcTool.Width(), 
						     rcTool.Height(), 
						     PATINVERT);
			assert(bResult);

			rcTool = rcBound;
			rcTool.left = rcTool.right - 2;
			bResult = PatBlt(hDC,
							 rcTool.left, 
							 rcTool.top, 
							 rcTool.Width(), 
							 rcTool.Height(), 
							 PATINVERT);
			assert(bResult);
		}
		
		SetTextColor(hDC, crTextOld);
		SetBkColor(hDC, crBackOld);
		
		SelectBrush(hDC, hBrushOld);
		
		bResult = DeleteBrush(hBrush);
		assert(bResult);
	}
	return NOERROR;
}

//
// DSGInsertTool
//

STDMETHODIMP CBand::DSGInsertTool(int nIndex, ITool* pTool)
{
	HRESULT hResult = m_pTools->InsertTool(nIndex, pTool, VARIANT_TRUE);
	if (SUCCEEDED(hResult) && m_pBar && m_pBar->m_pDesignerNotify)
	{
		m_pBar->m_pDesignerNotify->BandChanged((LPDISPATCH)(IBand*)this, 
											   (LPDISPATCH)(IBand*)(ddCBSNone != bpV1.m_cbsChildStyle ? m_pChildBands->GetCurrentChildBand() : NULL),
											   ddBandModified);
	}
	return hResult;
}

//
// DSGDrawSelection
// 
// Width is used when painting selection
//
//	rcTool rectangles for ddBTPopup else it should be called width -1 
//
// If nIndex < 0 then 
//		Height =- nIndex and don't use m_prcTools
//

STDMETHODIMP CBand::DSGDrawSelection(OLE_HANDLE ohDC, int x, int y, int nWidth, int nIndex) 
{
	CRect rcBound;
	HDC hDC = (HDC)ohDC;
	if (nIndex >= 0)
	{
		rcBound = m_prcTools[nIndex];
		rcBound.Offset(x, y);
		if (-1 != nWidth)
			rcBound.right = rcBound.left + nWidth;
	}
	else
	{
		rcBound.left = x;
		rcBound.top = y;
		rcBound.right = x + nWidth;
		rcBound.bottom = y - nIndex;
	}
	HBRUSH hBrush = CreateHalftoneBrush(TRUE);
	if (NULL == hBrush)
		return E_FAIL;
	
	BOOL bResult;
	HBRUSH hBrushOld = SelectBrush(hDC, hBrush);

	SetTextColor(hDC, m_pBar->m_crBackground);
	SetBkColor(hDC, m_pBar->m_crBackground);

	CRect rc = rcBound;
	rc.bottom = rc.top + 2;
	bResult = PatBlt(hDC, 
				     rc.left, 
				     rc.top, 
				     rc.Width(), 
				     rc.Height(), 
				     PATINVERT);
	if (!bResult)
		goto CleanUp;

	rc = rcBound;
	rc.top = rc.bottom - 2;
	bResult = PatBlt(hDC, 
				     rc.left, 
				     rc.top, 
				     rc.Width(), 
				     rc.Height(), 
				     PATINVERT);
	if (!bResult)
		goto CleanUp;

	rc = rcBound;
	rc.right = rc.left + 2;
	rc.top += 2;
	rc.bottom -= 2;
	bResult = PatBlt(hDC, 
					 rc.left, 
					 rc.top, 
					 rc.Width(), 
					 rc.Height(), 
					 PATINVERT);
	if (!bResult)
		goto CleanUp;

	rc = rcBound;
	rc.left = rc.right - 2;
	rc.top += 2;
	rc.bottom -= 2;
	bResult = PatBlt(hDC, 
					 rc.left, 
					 rc.top, 
					 rc.Width(),
					 rc.Height(), 
					 PATINVERT);

CleanUp:
	
	SelectBrush(hDC, hBrushOld);
	
	bResult = DeleteBrush(hBrush);
	assert(bResult);

	if (!bResult)
		return E_FAIL;
	
	return NOERROR;
}

//
// DSGDraw
//

STDMETHODIMP CBand::DSGDraw(OLE_HANDLE hDC, int x, int y, int w, int h)
{
	return NOERROR;
}

//
// DSGGetSize
//

STDMETHODIMP CBand::DSGGetSize(int dx, int dy, int bandType, int* w,  int* h) 
{
	return NOERROR;
}

//
// DSGCalcHit
//

STDMETHODIMP CBand::DSGCalcHit(int x, int y, int* pIndex)
{
	return NOERROR;
}

//
// Exchange
//

HRESULT CBand::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		int nWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
		int nHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
		long nStreamSize;
		short nSize;
		short nSize2;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving
			//

			if (m_pBar && m_pBar->m_bSaveImages)
			{
				//
				// Just images
				//

				hResult = m_pTools->Exchange(pStream, vbSave);

				if (m_pChildBands)
				{
					hResult = m_pChildBands->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}
			}
			else
			{
				nStreamSize = GetStringSize(m_bstrCaption) + GetStringSize(m_bstrName) + sizeof(nSize) + sizeof(bpV1);
				nStreamSize += GetVariantSize(m_vTag) + m_phPicture.GetSize();
				hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
				if (FAILED(hResult))
					return hResult;

				hResult = StWriteBSTR(pStream, m_bstrCaption);
				if (FAILED(hResult))
					return hResult;
				
				hResult = StWriteBSTR(pStream, m_bstrName);
				if (FAILED(hResult))
					return hResult;
		
				nSize = sizeof(bpV1);
				hResult = pStream->Write(&nSize, sizeof(nSize), NULL);
				if (FAILED(hResult))
					return hResult;

				CRect rcTemp = bpV1.m_rcFloat;
				bpV1.m_rcFloat.left = (10000 * bpV1.m_rcFloat.left)/nWidth;
				bpV1.m_rcFloat.right = (10000 * bpV1.m_rcFloat.right)/nWidth;
				bpV1.m_rcFloat.top = (10000 * bpV1.m_rcFloat.top)/nHeight;
				bpV1.m_rcFloat.bottom = (10000 * bpV1.m_rcFloat.bottom)/nHeight;

				hResult = pStream->Write(&bpV1, nSize, NULL);
				if (FAILED(hResult))
					return hResult;
				
				bpV1.m_rcFloat = rcTemp;

				hResult = StWriteVariant(pStream, m_vTag);
				if (FAILED(hResult))
					return hResult;

				hResult = PersistPicture(pStream, &m_phPicture, vbSave);
				if (FAILED(hResult))
					return hResult;

				hResult = m_pTools->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;

				hResult = pStream->Write(&m_pChildBands, sizeof(m_pChildBands), NULL);
				if (FAILED(hResult))
					return hResult;

				if (m_pChildBands)
				{
					hResult = m_pChildBands->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}
			}
		}
		else
		{
			//
			// Loading
			//

			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			m_bLoading = TRUE;

			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = StReadBSTR(pStream, m_bstrCaption);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= GetStringSize(m_bstrCaption);
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

			nSize2 = sizeof(bpV1);
			hResult = pStream->Read(&bpV1, nSize < nSize2 ? nSize : nSize2, NULL);
			if (FAILED(hResult))
				return hResult;

			bpV1.m_rcFloat.left = (nWidth * bpV1.m_rcFloat.left)/10000;
			bpV1.m_rcFloat.right = (nWidth * bpV1.m_rcFloat.right)/10000;
			bpV1.m_rcFloat.top = (nHeight * bpV1.m_rcFloat.top)/10000;
			bpV1.m_rcFloat.bottom = (nHeight * bpV1.m_rcFloat.bottom)/10000;

			if (ddBTStatusBar == bpV1.m_btBands && VARIANT_TRUE == bpV1.m_vbWrapTools)
				bpV1.m_vbWrapTools = VARIANT_FALSE;

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

			hResult = PersistPicture(pStream, &m_phPicture, vbSave);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= m_phPicture.GetSize();
			if (nStreamSize <= 0)
				goto FinishedReading;

			nOffset.HighPart = 0;
			nOffset.LowPart = nStreamSize;
			hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
			if (FAILED(hResult))
				return hResult;
FinishedReading:
//			bpV1.m_rcFloat.Offset(-bpV1.m_rcFloat.left, -bpV1.m_rcFloat.top);

			if (ddBTStatusBar == bpV1.m_btBands)
				m_pBar->StatusBand(this);

			hResult = m_pTools->Exchange(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			//
			// Was there a ChildBands Collection?
			//

			CChildBands* pChildBandTemp;
			hResult = pStream->Read(&pChildBandTemp, sizeof(pChildBandTemp), NULL);
			if (FAILED(hResult))
				return hResult;

			if (pChildBandTemp)
			{
				//
				// The answer is yes
				//

				if (NULL == m_pChildBands)
				{
					m_pChildBands = CChildBands::CreateInstance(NULL);
					if (NULL == m_pChildBands)
						return E_OUTOFMEMORY;

					m_pChildBands->SetOwner(this);
				}

				hResult = m_pChildBands->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
			OleTranslateColor(bpV1.m_ocPictureBackground, NULL, &m_crPictureBackground);
			m_bLoading = FALSE;
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

HRESULT CBand::ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		int nWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
		int nHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
		long nStreamSize;
		short nSize;
		short nSize2;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving
			//

			nStreamSize = GetStringSize(m_bstrCaption) + GetStringSize(m_bstrName) + sizeof(nSize) + sizeof(bpV1);
			nStreamSize += GetVariantSize(m_vTag) + m_phPicture.GetSize();
			hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = StWriteBSTR(pStream, m_bstrCaption);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteBSTR(pStream, m_bstrName);
			if (FAILED(hResult))
				return hResult;
			
			nSize = sizeof(bpV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			CRect rcTemp = bpV1.m_rcFloat;
			bpV1.m_rcFloat.left = (10000 * bpV1.m_rcFloat.left)/nWidth;
			bpV1.m_rcFloat.right = (10000 * bpV1.m_rcFloat.right)/nWidth;
			bpV1.m_rcFloat.top = (10000 * bpV1.m_rcFloat.top)/nHeight;
			bpV1.m_rcFloat.bottom = (10000 * bpV1.m_rcFloat.bottom)/nHeight;

			hResult = pStream->Write(&bpV1, nSize, NULL);
			if (FAILED(hResult))
				return hResult;
			
			hResult = StWriteVariant(pStream, m_vTag);
			if (FAILED(hResult))
				return hResult;

			hResult = PersistPicture(pStream, &m_phPicture, vbSave);
			if (FAILED(hResult))
				return hResult;

			hResult = m_pTools->ExchangeConfig(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&m_pChildBands,sizeof(m_pChildBands),NULL);
			if (FAILED(hResult))
				return hResult;

			if (m_pChildBands)
			{
				hResult = m_pChildBands->ExchangeConfig(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
			//
			// Loading
			//

			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			m_bLoading = TRUE;
			
			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = StReadBSTR(pStream, m_bstrCaption);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= GetStringSize(m_bstrCaption);
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

			nSize2 = sizeof(bpV1);
			hResult = pStream->Read(&bpV1, nSize < nSize2 ? nSize : nSize2, NULL);
			if (FAILED(hResult))
				return hResult;

			bpV1.m_rcFloat.left = (nWidth * bpV1.m_rcFloat.left)/10000;
			bpV1.m_rcFloat.right = (nWidth * bpV1.m_rcFloat.right)/10000;
			bpV1.m_rcFloat.top = (nHeight * bpV1.m_rcFloat.top)/10000;
			bpV1.m_rcFloat.bottom = (nHeight * bpV1.m_rcFloat.bottom)/10000;
			if (ddBTStatusBar == bpV1.m_btBands && VARIANT_TRUE == bpV1.m_vbWrapTools)
				bpV1.m_vbWrapTools = VARIANT_FALSE;

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

			hResult = PersistPicture(pStream, &m_phPicture, vbSave);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= m_phPicture.GetSize();
			if (nStreamSize <= 0)
				goto FinishedReading;

FinishedReading:
			if (ddBTStatusBar == bpV1.m_btBands)
				m_pBar->StatusBand(this);

			hResult = m_pTools->ExchangeConfig(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			//
			// Was there a ChildBands Collection?
			//

			CChildBands* pChildBandTemp;
			hResult = pStream->Read(&pChildBandTemp, sizeof(pChildBandTemp), NULL);
			if (FAILED(hResult))
				return hResult;

			if (pChildBandTemp)
			{
				if (NULL == m_pChildBands)
				{
					m_pChildBands = CChildBands::CreateInstance(NULL);
					if (NULL == m_pChildBands)
						return E_OUTOFMEMORY;

					m_pChildBands->SetOwner(this);
				}

				hResult = m_pChildBands->ExchangeConfig(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;

#ifdef _ABDLL
				
				if (ddCBSNone != bpV1.m_cbsChildStyle)
					OnChildBandChanged(0);
#endif

			}
			m_bLoading = FALSE;
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
// CalcPopupLayout
//

BOOL CBand::CalcPopupLayout(int& nWidth, int& nHeight, BOOL bCommit, BOOL bPopupMenu)
{
	//
	// Width = nMaxWidth of tool in menubar band + subband symbol if needed + picture if defined
	//

	m_sizeMaxIcon.cx = m_sizeMaxIcon.cy = 0;
	TypedArray<CTool*> aTools;
	CTool* pTool;
	m_bToolRemoved = FALSE;
	int nToolCount = m_pTools->GetVisibleTools(aTools);
	int nTool = 0;
	
	if (m_pBar && ddPMDisabled != m_pBar->bpV1.m_pmMenus && !bPopupMenu)
	{
		BOOL bMFU = FALSE;
		for (nTool = 0; nTool < nToolCount; nTool++)
		{
			pTool = aTools.GetAt(nTool);
			if ((pTool->tpV1.m_nUsage > CBar::eUsage || ddMVAlwaysVisible == pTool->tpV1.m_mvMenuVisibility) && ddTTSeparator != pTool->tpV1.m_ttTools)
			{
				bMFU = TRUE;
				break;
			}
		}

		if (bMFU && !(m_pBar->m_bCustomization || m_pBar->m_bWhatsThisHelp || m_pBar->m_bAltCustomization))
		{
			ToolTypes ttPrev = (ToolTypes)-1;
			BOOL bFirstTool = TRUE;
			nTool = 0;
			while (nTool < nToolCount)
			{
				pTool = aTools.GetAt(nTool);
				if ((pTool->tpV1.m_nUsage <= CBar::eUsage && ddMVAlwaysVisible != pTool->tpV1.m_mvMenuVisibility && ddTTSeparator != pTool->tpV1.m_ttTools) || (ddTTSeparator == pTool->tpV1.m_ttTools && ddTTSeparator == ttPrev) || (bFirstTool && ddTTSeparator == pTool->tpV1.m_ttTools))
				{
					// Trim out tools that fall below the freq
					m_bToolRemoved = TRUE;
					if (!m_pBar->m_bPopupMenuExpanded)
					{
						nToolCount--;
						pTool->m_bMFU = FALSE;
						aTools.RemoveAt(nTool);
					}
					else
						nTool++;
				}
				else
				{
					if (bFirstTool)
						bFirstTool = FALSE;
					nTool++;
					pTool->m_bMFU = TRUE;
					ttPrev = pTool->tpV1.m_ttTools;
				}
			}
			if (!m_pBar->m_bPopupMenuExpanded)
			{
				if (m_bToolRemoved)
				{
					if (NULL == m_pToolMenuExpand)
					{
						m_pToolMenuExpand = CTool::CreateInstance(NULL);
						if (m_pToolMenuExpand)
						{
							m_pToolMenuExpand->SetBand(this);
							m_pToolMenuExpand->tpV1.m_ttTools =	(ToolTypes)CTool::ddTTMenuExpandTool;
							m_pToolMenuExpand->tpV1.m_nToolId = CBar::eToolMenuExpand;
						}
					}
					if (m_pToolMenuExpand)
					{
						m_pToolMenuExpand->m_bPressed = FALSE;
						m_pToolMenuExpand->m_bDropDownPressed = FALSE;
						HRESULT hResult = aTools.Add(m_pToolMenuExpand);
						if (SUCCEEDED(hResult))
							nToolCount++;
					}
				}
			}
		}
	}

	if (nToolCount > 0)
	{
		HFONT hFont = m_pBar->GetMenuFont(FALSE, FALSE);
		assert(hFont);
		if (NULL == hFont)
			return FALSE;

		LOGFONT lf;
		GetObject(hFont, sizeof(LOGFONT), &lf);
		lf.lfWeight = FW_BOLD;
		HFONT hFontDefault = CreateFontIndirect(&lf);

		HDC hDC = GetDC(NULL);
		if (NULL == hDC)
			return FALSE;

		HFONT hFontDefaultOld;
		HFONT hFontOld = SelectFont(hDC, hFont);
		if (bCommit)
		{
			if (m_prcPopupExpandArea)
			{
				delete [] m_prcPopupExpandArea;
				m_prcPopupExpandArea = NULL;
			}
			if (m_prcTools)
			{
				delete [] m_prcTools;
				m_prcTools = NULL;
			}
			if (nToolCount > 0)
			{
				m_prcPopupExpandArea = new CRect [nToolCount];
				if (NULL == m_prcPopupExpandArea)
					return FALSE;
				m_prcTools = new CRect [nToolCount];
				assert(m_prcTools);
				if (NULL == m_prcTools)
					return FALSE;
			}
		}
	
		SIZEL sizePicture = {0,0};
		if (m_phPicture.m_pPict)
		{
			HRESULT hResult = m_phPicture.m_pPict->get_Width(&sizePicture.cx);
			hResult = m_phPicture.m_pPict->get_Height(&sizePicture.cy);
			HiMetricToPixel(&sizePicture, &sizePicture);
			sizePicture.cx += (m_pBar->m_bPopupMenuExpanded ? 3 : 0);
		}

		BOOL bResult;
		int nMaxShortCutWidth = 0;
		int nMaxToolWidth = 0;
		int nMaxWidth = 0;
		int nCurrentArea = -1;
		int nY = 0;
		for (nTool = 0; nTool < nToolCount; nTool++)
		{
			pTool = aTools.GetAt(nTool);
			if (NULL == pTool)
				continue;

			// Hook Tool to Band 
			pTool->m_pBand = this;  

			SIZE sizeIcon;
			SIZE sizeTool = pTool->CalcSize(hDC, ddBTPopup, FALSE, &sizeIcon);
			if (sizeTool.cx > nMaxToolWidth)
				nMaxToolWidth = sizeTool.cx;
			
			if (sizeIcon.cx > m_sizeMaxIcon.cx)
				m_sizeMaxIcon.cx = sizeIcon.cx;
			if (sizeIcon.cy > m_sizeMaxIcon.cy)
				m_sizeMaxIcon.cy = sizeIcon.cy;

			CRect rcShortcut;
			SIZE sizeTemp = {0,0};
			LPTSTR szKey = pTool->GetShortCutString();
			if (szKey)
			{
				if (VARIANT_TRUE == pTool->tpV1.m_vbDefault)
					hFontDefaultOld = SelectFont(hDC, hFontDefault);

				::DrawText(hDC, szKey, lstrlen(szKey), &rcShortcut, DT_LEFT|DT_CALCRECT);
				sizeTemp = rcShortcut.Size();
				if (sizeTemp.cx > nMaxShortCutWidth)
					nMaxShortCutWidth = sizeTemp.cx;

				if (VARIANT_TRUE == pTool->tpV1.m_vbDefault)
					SelectFont(hDC, hFontDefaultOld);
			}

			if ((sizeTool.cx + nMaxShortCutWidth) > nMaxWidth)
				nMaxWidth = sizeTool.cx + nMaxShortCutWidth;

			if (bCommit)
			{
				m_prcTools[nTool].left = 0;
				m_prcTools[nTool].bottom = m_prcTools[nTool].top = nY;
				m_prcTools[nTool].bottom += sizeTool.cy;
				m_prcTools[nTool].Offset(sizePicture.cx, 0);
				pTool->m_rcTemp = m_prcTools[nTool]; 
			}
			nY += sizeTool.cy;
		}
		
		if (nMaxShortCutWidth > 0)
		{
			m_nPopupShortCutStringOffset = nMaxWidth - nMaxShortCutWidth;
			nMaxWidth += 10;
		}
	
		if (bCommit)
		{
			int nOffset = 0;
			for (int nTool = 0; nTool < nToolCount; nTool++)
			{
				m_prcTools[nTool].right = m_prcTools[nTool].left + nMaxWidth;

				pTool = aTools.GetAt(nTool);
				if (NULL == pTool)
					continue;

				if (m_pBar && ddPMDisabled != m_pBar->bpV1.m_pmMenus && m_pBar->m_bPopupMenuExpanded && m_bToolRemoved)
				{
					if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
					{
						if (ddTTSeparator == pTool->tpV1.m_ttTools || pTool->tpV1.m_nUsage <= CBar::eUsage && ddMVAlwaysVisible != pTool->tpV1.m_mvMenuVisibility)
						{
							if (-1 == nCurrentArea && ddTTSeparator != pTool->tpV1.m_ttTools)
							{
								nCurrentArea = nTool;
								m_prcPopupExpandArea[nCurrentArea] = m_prcTools[nTool];
								m_prcPopupExpandArea[nCurrentArea].right = m_prcPopupExpandArea[nCurrentArea].left + m_sizeMaxIcon.cx;
								m_prcPopupExpandArea[nCurrentArea].Inflate(1, 1);
								m_prcPopupExpandArea[nCurrentArea].Offset(1, nOffset + 1);
								nY++;
								nOffset++;
							}
							else if (-1 != nCurrentArea)
								m_prcPopupExpandArea[nCurrentArea].bottom = m_prcTools[nTool].bottom + nOffset + 1;
							m_prcTools[nTool].Offset(1, nOffset);
						}
						else
						{
							nCurrentArea = -1;
							nY += 2;
							nOffset += 2;
							m_prcTools[nTool].Offset(1, nOffset);
						}
					}
					else
					{
						if (ddTTSeparator == pTool->tpV1.m_ttTools || pTool->tpV1.m_nUsage > CBar::eUsage || ddMVAlwaysVisible == pTool->tpV1.m_mvMenuVisibility)
						{
							if (-1 == nCurrentArea && ddTTSeparator != pTool->tpV1.m_ttTools)
							{
								nCurrentArea = nTool;
								m_prcPopupExpandArea[nCurrentArea] = m_prcTools[nTool];
								if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
								{
									m_prcPopupExpandArea[nCurrentArea].right = m_prcPopupExpandArea[nCurrentArea].left + m_sizeMaxIcon.cx;
								}
								m_prcPopupExpandArea[nCurrentArea].Inflate(1, 1);
								m_prcPopupExpandArea[nCurrentArea].Offset(1, nOffset + 1);
								nY++;
								nOffset++;
							}
							else if (-1 != nCurrentArea)
								m_prcPopupExpandArea[nCurrentArea].bottom = m_prcTools[nTool].bottom + nOffset + 1;
							m_prcTools[nTool].Offset(1, nOffset);
						}
						else
						{
							nCurrentArea = -1;
							nY += 2;
							nOffset += 2;
							m_prcTools[nTool].Offset(1, nOffset);
						}
					}
				}
			}
			m_nLastPopupExpandedArea = nCurrentArea;
		}
					
		if (m_pBar && ddPMDisabled != m_pBar->bpV1.m_pmMenus && m_pBar->m_bPopupMenuExpanded && m_bToolRemoved)
		{
			nMaxWidth += 2;
			nY++;
		}
		nWidth = nMaxWidth + sizePicture.cx;
		nHeight = nY;
		SelectFont(hDC, hFontOld);
		if (hFontDefault)
		{
			bResult = DeleteFont(hFontDefault);
			assert(bResult);
		}
		ReleaseDC(NULL, hDC);
	}
	else
	{
		nWidth = eEmptyBarWidth * 6;
		nHeight = eEmptyBarHeight;
	}
	
	if (!bPopupMenu && (ddBFDetach & m_pRootBand->bpV1.m_dwFlags) && (ddDAPopup == m_pRootBand->bpV1.m_daDockingArea || ddBTPopup == m_pRootBand->bpV1.m_btBands))
	{
		if (bCommit && m_prcTools)
		{
			int nCount = nToolCount;
			for (int nTool = 0; nTool < nCount; nTool++)
				m_prcTools[nTool].Offset(0, eDetachBarHeight);
		}
		nHeight += eDetachBarHeight;
	}
	if (bCommit)
		m_pTools->CommitVisibleTools(aTools);
	return TRUE;
}

//
// IsAltGR
//

static BOOL IsAltGR()
{
	BOOL bAltGR = FALSE;
	HKL hkl = GetKeyboardLayout(GetCurrentThreadId());
	long nId = LOWORD(hkl);
	switch (nId)
	{
	case 0x0407:
	case 0x0807:
	case 0x0c07:
	case 0x1007:
	case 0x1407:
	case 0x0405:
		bAltGR = TRUE;
		break;
	default:
		bAltGR = FALSE;
		break;
	}
	return bAltGR;
}

//
// EnterMenuLoop
//
// There is a problem with the WM_MOUSEMOVE message. When a new
// popup window is created the MOUSEMOVE message is sent which causes the loop
// to process it. We eliminate this problem by recording the cursor position at the point 
// where the key is pressed
//
// entrymode=0 for keyboard based entry
// entrymode=1 for click entry
//
//

void CBand::EnterMenuLoop(int nToolIndex, CTool* pTool, int nEntryMode, UINT nInitKey)
{
	CBand* pDetachBand = NULL; 
	POINT  pt;
	CRect  rcTool;
	DWORD  dwStartTime = GetTickCount();
	HWND   hWndPrevTarget = NULL;
	HWND   hWndPopup;
	HWND   hWndBand;
	BOOL   bDispatchLastMessage = TRUE;
	BOOL   bButtonIsDown = FALSE;
	BOOL   bOpenEnabled = TRUE;
	BOOL   bIsKeyUsed;
	BOOL   bLButtonDownOutOfArea = FALSE;
	static BOOL bMenuItemActive= FALSE;
	MSG    msg;
	int    nTool = 0;
	int    nNewToolIndex;

	TRACE(1, _T("Enter menu loop\n"));
	if (NULL == m_prcTools)
		return;

	SetCursor(LoadCursor(NULL, IDC_ARROW));

	assert(m_pBar);
	m_pBar->m_bMenuLoop = m_pBar->m_bClickLoop = TRUE;

	HWND hWndTemp = m_pBar->m_hWnd;
	KillTimer(hWndTemp, CBar::eFlyByTimer);
	
	// prep stage
	int nToolCount = m_pTools->GetVisibleToolCount(); 
	
	CRect rcDockAreaScreenRelative;
	if (!GetBandRect(hWndBand, rcDockAreaScreenRelative))
	{
		assert(FALSE);
		return;
	}
	
	if (ddDAFloat == bpV1.m_daDockingArea && m_pFloat)
		m_pFloat->ClientToScreen(rcDockAreaScreenRelative);
	else
	{
		CRect rcWin;
		GetWindowRect(hWndBand, &rcWin);
		rcDockAreaScreenRelative.Offset(rcWin.left, rcWin.top);
	}

	POINT ptCursorAtKeyDown;
	
	m_pBar->m_bFirstPopup = FALSE;
	try
	{
		// VK_MENU
		switch (nEntryMode)
		{
		case ENTERMENULOOP_KEY:
			{
				GetCursorPos(&ptCursorAtKeyDown);
				if (VK_MENU == nInitKey || VK_F10 == nInitKey)
				{
					bOpenEnabled = FALSE;
					
					BOOL bMaximized = FALSE;
					if (NULL != m_pBar->m_hWndMDIClient)
						::SendMessage(m_pBar->m_hWndMDIClient, WM_MDIGETACTIVE, 0, (LPARAM)&bMaximized);

					nTool = 0;
					if (bMaximized)
						nTool = 1;

					pTool = m_pTools->GetVisibleTool(nTool);
					if (nTool != nToolCount)
					{
						nToolIndex = nTool;
						try
						{
							m_pBar->FireMouseEnter((Tool*)pTool);
						}
						CATCH
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}

						//
						// Just Alt was pressed so bOpenEnabled = FALSE, Tool HighLighted
						//
						
						m_pBar->SetActiveTool(pTool); 
					}
					else
						pTool = NULL;
				}
				else
				{ 
					//
					// Some SysKeyDown was pressed
					//

					nNewToolIndex = CheckMenuKey(nInitKey);
					if (-1 == nNewToolIndex)
					{
						// not a valid key exit
						m_pBar->m_bMenuLoop = m_pBar->m_bClickLoop = FALSE; 
						return;
					}
					else 
					{
						//
						// Was a valid ALT+X key
						//

						nToolIndex = nNewToolIndex;
						pTool = m_pTools->GetVisibleTool(nToolIndex);
						if (pTool)
						{
							try
							{
								m_pBar->FireMouseEnter((Tool*)pTool); //*new
								pTool->OnMenuLButtonDown(0, ptCursorAtKeyDown);
								bButtonIsDown = TRUE;
								InvalidateToolRect(nToolIndex);
							}
							CATCH
							{
								assert(FALSE);
								REPORTEXCEPTION(__FILE__, __LINE__)
							}
						}


						if (pTool->HasSubBand())
						{
							SetPopupIndex(nToolIndex, NULL, TRUE);

							//
							// This is necessary because we are using a hook to get the keyboard messages
							// we got in here because of a WM_SYSKEYDOWN.  Now we need to get rid of it before 
							// we continue.
							//

							::PeekMessage(&msg, NULL, WM_SYSKEYDOWN, WM_SYSKEYDOWN, PM_REMOVE);
						}
						else
						{   
							//
							// Total exit fire click event the root menu item was clicked
							//

							try
							{
								pTool->OnMenuLButtonUp(TRUE);
								InvalidateToolRect(nToolIndex);
								m_pBar->FireMouseExit((Tool*)pTool); 
							}
							CATCH
							{
								assert(FALSE);
								REPORTEXCEPTION(__FILE__, __LINE__)
							}
							m_pBar->m_bClickLoop = FALSE;
							m_pBar->m_bMenuLoop = FALSE;
							return;				
						}
					}
				}
			}
			break;

		case ENTERMENULOOP_CLICK:
			{	
				//
				// Click
				//
				
				m_pBar->SetActiveTool(NULL);

				//
				// pTool was passed in as a parameter
				//

				assert(pTool);
				if (pTool)
				{
					try
					{
						m_pBar->FireMouseEnter((Tool*)pTool); 
					}
					CATCH
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}
				}

				bButtonIsDown = TRUE;

				SetPopupIndex(nToolIndex);
			}
			break;

		case ENTERMENULOOP_FIRSTITEM:
		case ENTERMENULOOP_LASTITEM:
			{
				if (ENTERMENULOOP_FIRSTITEM == nEntryMode)
					nNewToolIndex = 0;
				else
					nNewToolIndex = nToolCount - 1;

				nToolIndex = nNewToolIndex;
				pTool = m_pTools->GetVisibleTool(nToolIndex);
				if (pTool)
				{
					try
					{
						m_pBar->FireMouseEnter((Tool*)pTool);
					}
					CATCH
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}

					GetCursorPos(&pt);
					pTool->OnMenuLButtonDown(0, pt);
				}
				bButtonIsDown = TRUE;
				InvalidateToolRect(nToolIndex);
				SetPopupIndex(nToolIndex, NULL, TRUE);
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}

	//
	// Status Band Enter
	//
	
	if (ddBTChildMenuBar == m_pRootBand->bpV1.m_btBands || ddBTMenuBar == m_pRootBand->bpV1.m_btBands)
		m_pBar->StatusBandEnter(pTool);

	//
	// Start Message Loop
	//

	HWND hWndDockWindow = m_pBar->GetParentWindow();

	SetCapture(hWndBand);

	try
	{
		//
		// Go a head and get rid of the paint messages before we go into the loop
		//

		while (::PeekMessage(&msg, NULL, WM_PAINT, WM_PAINT, PM_NOREMOVE))
		{
			if (!GetMessage(&msg, NULL, WM_PAINT, WM_PAINT))
				return;
			DispatchMessage(&msg);
		}

		m_pBar->AddRef();

		//
		// Entering the menu loop
		//

		POINT ptInit;
		GetCursorPos(&ptInit);
		DWORD dwInitTime = GetTickCount();
		int nDblClkDistanceCx = GetSystemMetrics(SM_CXDOUBLECLK);
		int nDblClkDistanceCy = GetSystemMetrics(SM_CYDOUBLECLK);
		while (TRUE)
		{
			if (m_pBar->m_bWhatsThisHelp)
				SetCursor(LoadCursor(NULL, IDC_HELP));

			if (hWndBand != GetCapture())
			{
				goto ExitPoint;
			}

			if (!m_pBar->m_bMenuLoop)
			{
				bDispatchLastMessage = FALSE;
				goto ExitPoint;
			}

			if (!GetMessage(&msg, NULL, 0, 0))
				goto ExitPoint;

			// Check if TotalExit
			if (!IsWindow(hWndDockWindow))
				return;
			
			if (!m_pBar->IsForeground())
				goto ExitPoint;
			
			switch(msg.message)
			{

			// 
			// KEYBOARD
			// 
			// Handle initial ALT down + ALT up
			//

			case WM_SYSKEYUP:  
				TRACE1(3, "X SYSKEYUP =%X\n",msg.wParam);
				if (VK_MENU == msg.wParam || VK_F10 == msg.wParam)
				{
					if (!bMenuItemActive)
					{
						bMenuItemActive = TRUE;
						m_pBar->SetActiveTool(pTool);
					}
					else
					{
						bMenuItemActive = FALSE;
						m_pBar->SetActiveTool(NULL);
						goto ExitPoint;
					}
				}
				if (VK_CONTROL == msg.wParam && IsAltGR())
				{
					bMenuItemActive = FALSE;
					m_pBar->SetActiveTool(NULL);
					goto ExitPoint;
				}
				continue;

			case WM_SYSKEYDOWN: 
				TRACE1(3, "X SYSKEYDOWN=%X\n", msg.wParam);

				if ((VK_MENU == msg.wParam || VK_F10 == msg.wParam) && bMenuItemActive) 
				{
					//
					// Handle ALT toggle from active to passive state on the up
					//
					continue;
				}

				if (-1 != m_nPopupIndex && m_pChildPopupWin)
				{
					// send key to submenu
					bIsKeyUsed = FALSE;
					if (bMenuItemActive)
						msg.message = WM_KEYDOWN;

					m_pChildPopupWin->SendMessage(msg.message, msg.wParam, (LPARAM)&bIsKeyUsed);
					if (!bIsKeyUsed)
					{ 
						// key was not used so shutdown this popup
						TRACE(3, "ExitP Unused key\n");
						goto ExitPoint;
					}

					if (NULL == m_pChildPopupWin)
					{
						TRACE(3, "ExitP syskey used\n");

						bDispatchLastMessage = FALSE;
						goto ExitPoint;
					}
				}
				else if (NULL == pTool)
				{
					nNewToolIndex = CheckMenuKey(msg.wParam);
					if (-1 == nNewToolIndex)
					{
						nToolIndex = nNewToolIndex;
						bOpenEnabled = TRUE;
						pTool = m_pTools->GetVisibleTool(nToolIndex);
						if (pTool->HasSubBand())
						{
							pTool->OnMenuLButtonDown(0, pt);
							bButtonIsDown = TRUE;
							InvalidateToolRect(nToolIndex);
							SetPopupIndex(nToolIndex);
						}
						bMenuItemActive = TRUE;
					}
					else
					{
						if (VK_CONTROL != msg.wParam && VK_SHIFT != msg.wParam)
						{
							TRACE(3, "ExitP __keyx\n");
							goto ExitPoint;
						}
					}
				}
				else
				{
					TRACE(3, "ExitP key2\n");
					goto ExitPoint;
				}
				continue;

			case WM_KEYDOWN:
				TRACE1(1, "X KEYDOWN=%X\n", msg.wParam);

				GetCursorPos(&ptCursorAtKeyDown);

				// shutdown current popup
				if (VK_ESCAPE == msg.wParam)
				{
					if (bOpenEnabled)
					{
						if (pTool)
						{
							if (bButtonIsDown)
							{
								pTool->OnMenuLButtonUp();
								bButtonIsDown = FALSE;
							}
							InvalidateToolRect(nToolIndex);
							m_pBar->m_pToolShadow->Hide();
							SetPopupIndex(-1);
							bOpenEnabled = FALSE;
							m_pBar->SetActiveTool(pTool);
						}
						else
						{
							TRACE(1, "ExitP ESC key1\n");
							m_pBar->SetActiveTool(NULL);
							goto ExitPoint;
						}
					}
					else
					{
						TRACE(1, "ExitP ESC key2\n");
						m_pBar->SetActiveTool(NULL);
						goto ExitPoint;
					}

					bIsKeyUsed = FALSE;
					if (-1 != m_nPopupIndex && m_pChildPopupWin)
					{
						m_pChildPopupWin->SendMessage(msg.message, msg.wParam, (LPARAM)&bIsKeyUsed);
						if (NULL == m_pChildPopupWin)
						{
							TRACE(1, "ExitP key3\n");
							bDispatchLastMessage = FALSE;
							goto ExitPoint;
						}
					}
				}
					
				if (!bOpenEnabled)
				{
					if (VK_RETURN == msg.wParam || VK_DOWN == msg.wParam || VK_UP == msg.wParam)
					{
						bOpenEnabled = TRUE;
						m_pBar->SetActiveTool(NULL);
						if (pTool->HasSubBand())
						{
							pTool->OnLButtonDown();
							bButtonIsDown = TRUE;
							bDispatchLastMessage = FALSE; // CR 4045
						}
						else
							m_pBar->SetActiveTool(pTool);

						InvalidateToolRect(nToolIndex);
						if (pTool->HasSubBand())
							SetPopupIndex(nToolIndex, NULL, TRUE);
						if (VK_DOWN == msg.wParam)
							continue;
						break;
					}
					else
					{
						nNewToolIndex = CheckMenuKey(msg.wParam);
						if (-1 != nNewToolIndex)
						{
							if (pTool)
							{
								if (bButtonIsDown)
								{
									pTool->OnLButtonDown();
									bButtonIsDown = FALSE;
								}
								else
									m_pBar->SetActiveTool(NULL);
								InvalidateToolRect(nToolIndex);
							}
							nToolIndex = nNewToolIndex;

							pTool = (CTool*)m_pTools->GetVisibleTool(nToolIndex);
							if (pTool->HasSubBand())
							{
								SetPopupIndex(nToolIndex, 0, TRUE);
								pTool->OnLButtonDown();
								bButtonIsDown = TRUE;
								bOpenEnabled = TRUE;
							}
							else
								m_pBar->SetActiveTool(pTool);

							InvalidateToolRect(nToolIndex);
							break;
						}
					}
				}

				bIsKeyUsed = FALSE;
				if (-1 != m_nPopupIndex && m_pChildPopupWin)
				{
					m_pChildPopupWin->SendMessage(msg.message, msg.wParam, (LPARAM)&bIsKeyUsed);
					if (NULL == m_pChildPopupWin)
					{
						TRACE(5, "ExitP key3\n");
						PeekMessage(&msg, hWndBand, WM_KEYUP, WM_KEYUP, PM_REMOVE);
						bDispatchLastMessage = FALSE;
						goto ExitPoint;
					}
				}

				if (!bIsKeyUsed)
				{
					nNewToolIndex = nToolIndex;
					switch (msg.wParam)
					{
					case VK_LEFT:
						nNewToolIndex--;
						break;

					case VK_RIGHT:
						nNewToolIndex++;
						
					case VK_CONTROL:
					case VK_SHIFT:
						break;

					default:
						break;
					}

					//
					// Left or Right key was pressed so process it:
					//
					
					if (nNewToolIndex != nToolIndex)
					{
						//
						// switch active top level popups
						//
						
						if (nNewToolIndex >= nToolCount)
							nNewToolIndex = 0;
						else if (nNewToolIndex < 0)
						{
							nNewToolIndex = nToolCount - 1;
							if (nNewToolIndex < 0)
								nNewToolIndex = nToolIndex;
						}
						
						// Shutdown current popup
						if (pTool)
						{
							if (bButtonIsDown)
							{
								pTool->OnMenuLButtonUp(FALSE);
								bButtonIsDown = FALSE;
							}
							else
								m_pBar->SetActiveTool(NULL);

							try
							{
								m_pBar->FireMouseExit((Tool*)pTool);
							}
							CATCH
							{
								assert(FALSE);
								REPORTEXCEPTION(__FILE__, __LINE__)
							}
							InvalidateToolRect(nToolIndex);
						}

						SetPopupIndex(-1);
						m_pBar->m_pToolShadow->Hide();
						// Remove the shadow around the tool
						// Open Popup
 						nToolIndex = nNewToolIndex;
						pTool = m_pTools->GetVisibleTool(nToolIndex);
						
						if (bOpenEnabled && VARIANT_TRUE == pTool->tpV1.m_vbEnabled)
						{
							try
							{
								m_pBar->FireMouseEnter((Tool*)pTool);
							}
							CATCH
							{
								assert(FALSE);
								REPORTEXCEPTION(__FILE__, __LINE__)
							}

							if (pTool->HasSubBand())
							{
								POINT ptOp;
								rcTool = m_prcTools[nToolIndex];
								rcTool.Offset(rcDockAreaScreenRelative.left, rcDockAreaScreenRelative.top);
								ptOp.x = rcTool.right;
								m_nCurrentTool = nToolIndex;
								pTool->OnMenuLButtonDown(0, ptOp);
								m_nCurrentTool = -1;
								bButtonIsDown = TRUE;
							}
							else
								m_pBar->SetActiveTool(pTool);

							InvalidateToolRect(nToolIndex);
							if (pTool->HasSubBand())
							{
								SetPopupIndex(nToolIndex, 0, TRUE);
							}
						}
						else
							m_pBar->SetActiveTool(pTool);
					}
				}
				continue;

			case WM_MOUSEMOVE:
				{
					GetCursorPos(&pt);
					// mouse still at same pos so ignore
					if (0 == memcmp(&pt, &ptCursorAtKeyDown, sizeof(POINT)))
						break; 
		
					QueryPopupWin(hWndPopup, pt);
					if (IsWindow(hWndPrevTarget) && hWndPrevTarget && hWndPrevTarget != hWndPopup) 
					{
						// Send Old Target a mouse message so that it deactivates selection
						POINT ptClient = pt;
						ScreenToClient(hWndPrevTarget, &ptClient);
						SendMessage(hWndPrevTarget, WM_MOUSEMOVE, 0, MAKELONG(ptClient.x, ptClient.y));
						hWndPrevTarget = NULL;
					}
					
					if (hWndPopup)
					{
						POINT ptClient = pt;
						ScreenToClient(hWndPopup, &ptClient);
						m_pBar->m_nPopupExitAction = 0;
						SendMessage(hWndPopup, WM_MOUSEMOVE, 0, MAKELONG(ptClient.x, ptClient.y));
						hWndPrevTarget = hWndPopup;
						break;
					}

					GetCursorPos(&pt);
					if (PtInRect(&rcDockAreaScreenRelative, pt))
					{
						// Still in the Menubar. 
						CRect rcTool;
						// find tool under cursor
						CTool* pNewTool = NULL;
						nToolCount = m_pTools->GetVisibleToolCount(); 
						for (int nTool = 0; nTool < nToolCount; nTool++)
						{
							pNewTool = m_pTools->GetVisibleTool(nTool);
							if (!pNewTool->IsVisibleOnPaint())
								continue;

							GetToolScreenRect(nTool, rcTool);
							if (PtInRect(&rcTool, pt))
								break;
						}

						if (nTool == nToolCount || pTool == pNewTool)
							break;

						if (pTool && pTool != pNewTool)
						{
							// Hey the cursor left the active pTool. Close the current popup
							m_pBar->m_pToolShadow->Hide();
							if (bOpenEnabled)
							{
								if (bButtonIsDown)
								{
									pTool->OnMenuLButtonUp(FALSE);
									bButtonIsDown = FALSE;
								}
								else
									m_pBar->SetActiveTool(NULL);

								InvalidateToolRect(nToolIndex);
								try
								{
									m_pBar->FireMouseExit((Tool*)pTool);
								}
								CATCH
								{
									assert(FALSE);
									REPORTEXCEPTION(__FILE__, __LINE__)
								}
							}
							SetPopupIndex(-1);
							pTool = NULL;
						}

						if (pNewTool && VARIANT_TRUE == pNewTool->tpV1.m_vbEnabled)
						{
							//
							// We have entered yet another popup subband
							//

							pTool = pNewTool;
							nToolIndex = nTool;
							rcTool = m_prcTools[nToolIndex];
							rcTool.Offset(rcDockAreaScreenRelative.left, rcDockAreaScreenRelative.top);
							if (bOpenEnabled)
							{
								POINT ptOp;
								ptOp.x = rcTool.right;
								
								if (ddBTMenuBar == m_pRootBand->bpV1.m_btBands || ddBTChildMenuBar == m_pRootBand->bpV1.m_btBands)
									m_pBar->StatusBandUpdate(pTool);

								try
								{
									m_pBar->FireMouseEnter((Tool*)pTool);
								}
								CATCH
								{
									assert(FALSE);
									REPORTEXCEPTION(__FILE__, __LINE__)
								}

								if (pTool->HasSubBand()) 
								{
									m_nCurrentTool = nToolIndex;
									if (CTool::ddTTMoreTools == pTool->tpV1.m_ttTools)
										pTool->GenerateMoreTools();

									pTool->OnMenuLButtonDown(0, ptOp);
									
									m_nCurrentTool = -1;
									bButtonIsDown=TRUE;
								}
								else
									m_pBar->SetActiveTool(pTool);

								m_nToolMouseOver = nToolIndex;
								InvalidateToolRect(nToolIndex);
								if (pTool->HasSubBand())
								{
									SetPopupIndex(nToolIndex);
									break;
								}
							}
							else
								m_pBar->SetActiveTool(pTool);
						}
					}
				}	
				break;
			
			case WM_LBUTTONUP:
				{
					pt = msg.pt; // Testing for Karl
					if (GetTickCount() > dwStartTime + eMenuMouseUpDelay)
					{
						QueryPopupWin(hWndPopup, pt);
						if (hWndPopup)
						{
							POINT ptClient = pt;
							ScreenToClient(hWndPopup, &ptClient);
							m_pBar->m_nPopupExitAction = 0;
							SendMessage(hWndPopup, WM_LBUTTONUP, 0, MAKELONG(ptClient.x, ptClient.y));
							if (NULL == m_pChildPopupWin)
							{
								TRACE(1, "ExitP LButtonUp popup closed\n");
								bDispatchLastMessage = FALSE;
								goto ExitPoint;
							}
							break;
						}
					}
					else
						continue;
					
					if (!PtInRect(&rcDockAreaScreenRelative, pt))
					{
						TRACE(1, "ExitP LButton up invalid region\n");
						
						//
						// TODO get this to work
						//

						bLButtonDownOutOfArea = TRUE;
						goto ExitPoint;
					}
					
					//
					// If Tool does not have a SubBand exit and send command message
					//

					CTool* pExitTool;
					CRect rcTool;
					for (nTool = 0; nTool < nToolCount; nTool++)
					{
						pExitTool = m_pTools->GetVisibleTool(nTool);
						if (!pExitTool->IsVisibleOnPaint())
							continue;

						GetToolScreenRect(nTool, rcTool);
						rcTool.Inflate(2, 2);
						if (!PtInRect(&rcTool, pt))
							continue;

						if (pExitTool == pTool && (pExitTool->m_bPressed || pExitTool->m_bDropDownPressed))
						{
							if (pTool->m_bPressed && (ddTTButtonDropDown == pTool->tpV1.m_ttTools || !pTool->HasSubBand()))
							{
								pExitTool->OnMenuLButtonUp(TRUE, TRUE);
								InvalidateToolRect(nToolIndex);
								bButtonIsDown = FALSE;
								TRACE(1, "ExitP buttondropdown1\n");
								goto ExitPoint;
							}
						}
						break;
					}
					if (nTool == nToolCount)
					{
						TRACE(1, "ExitP LButtonup no hit\n");
						goto ExitPoint; // no hit
					}
				}
				break;

			case WM_LBUTTONDBLCLK:
				if (!m_pBar->m_bPopupMenuExpanded)
				{
					m_pBar->m_bPopupMenuExpanded = TRUE;
					if (pTool && pTool->HasSubBand() && m_pChildPopupWin && m_pChildPopupWin->m_pBand->m_bToolRemoved)
					{
						if (bOpenEnabled && bButtonIsDown)
						{
							//
							// This happens when you click on a menu and then click again
							//

							pTool->OnMenuLButtonUp(FALSE);
							SetPopupIndex(-1);
							bButtonIsDown = FALSE;
						}
					}
				}
			case WM_LBUTTONDOWN:
				GetCursorPos(&pt);
				if (WM_LBUTTONDBLCLK != msg.message && abs(pt.x - ptInit.x) < nDblClkDistanceCx && abs(pt.y - ptInit.y) < nDblClkDistanceCy)
				{
					// check time elapsed
					if ((GetTickCount() - dwInitTime) <= GetDoubleClickTime())
					{
						msg.message = WM_LBUTTONDBLCLK;
						if (!m_pBar->m_bPopupMenuExpanded)
						{
							m_pBar->m_bPopupMenuExpanded = TRUE;
							if (pTool && pTool->HasSubBand() && m_pChildPopupWin && m_pChildPopupWin->m_pBand->m_bToolRemoved)
							{
								if (bOpenEnabled && bButtonIsDown)
								{
									//
									// This happens when you click on a menu and then click again
									//

									pTool->OnMenuLButtonUp(FALSE);
									SetPopupIndex(-1);
									bButtonIsDown = FALSE;
								}
							}
						}
					}
				}

				GetCursorPos(&pt);
				QueryPopupWin(hWndPopup, pt);
				if (hWndPopup)
				{
					POINT ptClient = pt;
					ScreenToClient(hWndPopup, &ptClient);
					m_pBar->m_nPopupExitAction = 0;
					SendMessage(hWndPopup, WM_LBUTTONDOWN, 0, MAKELONG(ptClient.x, ptClient.y));
					if (NULL == m_pChildPopupWin)
					{
						TRACE(1, "ExitP LButtonDown popup closed\n");
						goto ExitPoint;
					}

					if (CPopupWin::PEA_DETACH == m_pBar->m_nPopupExitAction)
					{
						CPopupWin* pPopupWin = NULL;
						SendMessage(hWndPopup, GetGlobals().WM_POPUPWINMSG, 1, (LPARAM)&pPopupWin);
						if (pPopupWin)
						{
							HRESULT hResult = pPopupWin->m_pBand->Clone((IBand**)&pDetachBand);
							m_pBar->DestroyDetachBandOf(pPopupWin->m_pBand);
							if (pDetachBand)
							{
								pDetachBand->bpV1.m_vbDetached = VARIANT_TRUE;
								pDetachBand->m_daPrevDockingArea = bpV1.m_daDockingArea;
								pDetachBand->bpV1.m_nDockLine = 0;
								pDetachBand->bpV1.m_nDockOffset = 0;
								pDetachBand->bpV1.m_daDockingArea = ddDAFloat;
								pDetachBand->bpV1.m_vbVisible = VARIANT_TRUE;
								pDetachBand->bpV1.m_btBands = ddBTNormal;
								pDetachBand->bpV1.m_nCreatedBy = ddCBDetached;
								pDetachBand->OnChildBandChanged(0);
								pDetachBand->bpV1.m_rcFloat = pDetachBand->GetOptimalFloatRect(FALSE);
								pDetachBand->bpV1.m_vbDisplayMoreToolsButton = FALSE;

								GetCursorPos(&m_pBar->m_ptPea);

								pDetachBand->bpV1.m_rcFloat.Offset(-pDetachBand->bpV1.m_rcFloat.left + m_pBar->m_ptPea.x,
																   -pDetachBand->bpV1.m_rcFloat.top + m_pBar->m_ptPea.y);

								m_pBar->m_ptPea.x = 5;
								m_pBar->m_ptPea.y = - CMiniWin::eMiniWinCaptionHeight / 2;
								hResult = m_pBar->m_pBands->InsertBand(-1, pDetachBand);
								m_pBar->InternalRecalcLayout();
								bDispatchLastMessage = FALSE;
							}
							TRACE(1, "ExitP LButtonDown popup detached\n");
						}
						goto ExitPoint;
					}
					bDispatchLastMessage = FALSE;
					break;
				}

				if (!PtInRect(&rcDockAreaScreenRelative, pt))
				{
					bDispatchLastMessage = FALSE;
					goto ExitPoint;
				}

				if (pTool)
				{
					if (bOpenEnabled && bButtonIsDown)
					{
						//
						// This happens when you click on a menu and then click again
						//

						pTool->OnMenuLButtonUp(FALSE);
						bButtonIsDown = FALSE;
						InvalidateToolRect(nToolIndex);
						bDispatchLastMessage = FALSE;
						goto ExitPoint;
					}
					else if (ddTTButtonDropDown == pTool->tpV1.m_ttTools)
					{
						pTool->OnMenuLButtonUp(FALSE);
						bButtonIsDown = FALSE;
						InvalidateToolRect(nToolIndex);
						bDispatchLastMessage = FALSE;
						goto ExitPoint;
					}
					else
					{
						if (bButtonIsDown)
						{
							pTool->OnMenuLButtonUp(FALSE);
							InvalidateToolRect(nToolIndex);
						}
						
						bButtonIsDown = FALSE;
						m_pBar->SetActiveTool(NULL);
						bOpenEnabled = TRUE;

						rcTool = m_prcTools[nToolIndex];
						rcTool.Offset(rcDockAreaScreenRelative.left, rcDockAreaScreenRelative.top);

						POINT ptOp;
						ptOp.x = rcTool.right;

						m_nCurrentTool = nToolIndex;
						
						pTool->OnMenuLButtonDown(0, ptOp);
						
						m_nCurrentTool = -1;
						
						bButtonIsDown = TRUE;

						InvalidateToolRect(nToolIndex);
						if (pTool->HasSubBand())
						{
							SetPopupIndex(nToolIndex);
							break;
						}
					}
				}
				break;

			case WM_RBUTTONUP:
			case WM_RBUTTONDOWN:
				continue;

			case WM_CANCELMODE:
				DispatchMessage(&msg);
				TRACE(1, "ExitP CancelMode\n");
				goto ExitPoint;

			default:
				DispatchMessage(&msg);
				break;
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}

ExitPoint:
	try
	{
		m_pBar->m_bFirstPopup = FALSE;

		m_pBar->m_pToolShadow->Hide();

		m_pBar->m_bPopupMenuExpanded = FALSE;

		if (hWndBand == GetCapture())
			ReleaseCapture();

		SetPopupIndex(-1);

		if (bLButtonDownOutOfArea)
		{
			HWND hWnd = WindowFromPoint(pt);
			if (IsWindow(hWnd))
				PostMessage(hWnd, WM_NCHITTEST, pt.x, pt.y);
		}

		if (pTool)
		{
			pTool->m_bDropDownPressed = FALSE;
			if (bButtonIsDown)
			{
				pTool->OnMenuLButtonUp(FALSE);
				if (nToolIndex < GetVisibleToolCount())
					InvalidateToolRect(nToolIndex);
			}
			else
				m_pBar->SetActiveTool(NULL);
	
			try
			{
				m_pBar->FireMouseExit((Tool*)pTool);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}

			if (ENTERMENULOOP_CLICK == nEntryMode)
			{
				GetCursorPos(&pt);
				::ScreenToClient(hWndBand, &pt);
				OnMouseMove(1, pt);
			}
		}
		
		if (ddBTChildMenuBar == m_pRootBand->bpV1.m_btBands || ddBTMenuBar == m_pRootBand->bpV1.m_btBands)
			m_pBar->StatusBandExit();
		m_pBar->Release();

		TRACE(1, "ExitPoint release capture\n");
		if (!IsWindow(hWndTemp))
		{
			m_pBar->m_bClickLoop = FALSE;
			m_pBar->m_bMenuLoop = FALSE;
			return;
		}

		m_pBar->m_bMenuLoop = m_pBar->m_bClickLoop = FALSE;

		if (bDispatchLastMessage && !m_pBar->m_bWhatsThisHelp)
			DispatchMessage(&msg); 

		if (pDetachBand)
		{
			m_pDockMgr->StartDrag(pDetachBand, m_pBar->m_ptPea);
			pDetachBand->Release();
			m_pBar->m_nPopupExitAction = 0;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// QueryPopupWin
//
// ptReturn is returned relative to hWnd
//

void CBand::QueryPopupWin(HWND& hWnd, POINT& ptReturn) 
{
	POINT pt;
	GetCursorPos(&pt);
	ptReturn = pt;

	hWnd = NULL;
	HWND hWndPrev = NULL;
	HWND hWndTemp = WindowFromPoint(pt);
	while (hWndTemp)
	{
		hWndPrev = hWndTemp;
		hWndTemp = GetParent(hWndTemp);
	}

	BOOL bVal = FALSE;
	if (IsWindow(hWndPrev))
	{
		SendMessage(hWndPrev, GetGlobals().WM_POPUPWINMSG, 0, (LPARAM)&bVal);
		if (bVal)
			hWnd = hWndPrev;
	}
}

//
// CheckMenuKey
//

int CBand::CheckMenuKey(WORD nVirtKey)
{
	if (nVirtKey >= VK_NUMPAD0 && nVirtKey <= VK_NUMPAD9)
		nVirtKey = nVirtKey - '0'; 
	WCHAR  ch;
	CTool* pTool;
	int nToolCount = m_pTools->GetVisibleToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = m_pTools->GetVisibleTool(nTool);
		if (pTool->CheckForAlt(ch))
		{
			TCHAR ch2 = (TCHAR)ch;
			
			short nCode = VkKeyScan((TCHAR)ch);
			if (-1 == nCode)
				nCode = ch;

			UINT nShift = (HIBYTE(nCode) & 1) != 0;
			
			UINT nKey = LOBYTE(nCode);
			
			BOOL bShiftState = (GetKeyState(VK_SHIFT)&0x8000) != 0;
			
			BOOL bLetter = CharUpper((TCHAR*)MAKELONG(ch, 0)) != CharLower((TCHAR*)MAKELONG(ch2, 0));
			
			if (nVirtKey == nKey && (bShiftState == (0 != nShift) || bLetter))
				return nTool;
		}
	}
	return -1;
}

//
// PopupMenu
//

STDMETHODIMP CBand::PopupMenu(VARIANT vFlags, VARIANT x, VARIANT y)
{
	if (VT_ERROR == x.vt || VT_EMPTY == x.vt)
		x.iVal = -1;
	if (VT_ERROR == y.vt || VT_EMPTY==y.vt)
		y.iVal = -1;
	if (VT_ERROR == vFlags.vt)
		vFlags.iVal = ddPopupMenuLeftAlign;
	HRESULT hResult = TrackPopupEx(vFlags.iVal, x.iVal, y.iVal, NULL);
	HRESULT hResult2 = put_Visible(VARIANT_FALSE);
	return hResult;
}

//
// PopupMenuEx
//

STDMETHODIMP CBand::PopupMenuEx(int nFlags, int x, int y, int left, int top, int right, int bottom, VARIANT_BOOL* retval)
{
	assert(retval);
	CRect rc;
	rc.left = left;
	rc.top = top;
	rc.right = right;
	rc.bottom = bottom;
	BOOL bDblClick;
	HRESULT hResult = TrackPopupEx(nFlags, x, y, &rc, &bDblClick);
	*retval = -bDblClick;
	return hResult;
}

//
// TrackPopupEx
//

STDMETHODIMP CBand::TrackPopupEx(int nFlags, int x, int y, CRect* prcTrack, BOOL* pbDblClick)
{
	try
	{
		CBand* pDetachBand = NULL; 
		DWORD dwInitTime = GetTickCount();

		POINT ptInit;
		GetCursorPos(&ptInit);

		m_pBar->SetActiveTool(NULL);

		if (!CheckPopupOpen())
			return NOERROR;

		if (m_pPopupWin)
			m_pPopupWin->DestroyWindow();
		
		CPopupWin* pPopup = new CPopupWin(this, FALSE);
		if (NULL == pPopup)
			return E_OUTOFMEMORY;

		// This is to cheat and have the root popup band to be drawn animated it 
		// used in the PopupWin Draw function
		if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
			m_pParentPopupWin = pPopup;

		if (x == -1 || y == -1)
		{
			x = ptInit.x;
			y = ptInit.y;
		}
		else
		{
			SIZE size;
			size.cx = x;
			size.cy = y;
			TwipsToPixel(&size, &size);
			x = size.cx;
			y = size.cy;
			if (prcTrack)
			{
				size.cx = prcTrack->left;
				size.cy = prcTrack->top;
				TwipsToPixel(&size, &size);
				prcTrack->left = size.cx;
				prcTrack->top = size.cy;

				size.cx = prcTrack->right;
				size.cy = prcTrack->bottom;
				TwipsToPixel(&size, &size);
				prcTrack->right = size.cx;
				prcTrack->bottom = size.cy;
			}
		}

		m_bPopupWinLock = TRUE;
		CRect rcBound;
		rcBound.left = rcBound.right = x;
		rcBound.top = rcBound.bottom = y;

		VARIANT_BOOL vbDetach = bpV1.m_vbDetached;
		bpV1.m_vbDetached = VARIANT_FALSE;

		pPopup->m_bPopupMenu = TRUE;
		if (!pPopup->PreCreateWin(NULL, rcBound, CPopupWin::eDirectionFlipHort, FALSE, bpV1.m_dwFlags & ddBFPopupFlipUp, nFlags))
		{
			delete pPopup;
			return E_FAIL;
		}

		if (!pPopup->CreateWin(m_pBar->GetDockWindow()))
		{
			delete pPopup;
			return E_FAIL;
		}

		HWND hWndPopup = pPopup->hWnd();

		if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
		{
			pPopup->Shadow().SetWindowPos(NULL, 
										  0, 
										  0, 
										  0, 
										  0,
										  SWP_SHOWWINDOW|SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
		}

		pPopup->SetWindowPos(HWND_TOPMOST, // CR 4134
							 0, 
							 0, 
							 0, 
							 0,
							 SWP_SHOWWINDOW  | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);

		pPopup->UpdateWindow();

		// make sure our thread is foreground
		//if (hWndPopup != GetActiveWindow())  CR 4134
		//	SetTopMost ForegroundWindow(hWndPopup);

		SetCapture(hWndPopup);

		MSG msg;
		HWND hWndPrevTarget = NULL;
		if (pbDblClick)
			*pbDblClick = FALSE;

		int nDblClkDistanceCx = GetSystemMetrics(SM_CXDOUBLECLK);
		int nDblClkDistanceCy = GetSystemMetrics(SM_CYDOUBLECLK);

		m_pBar->m_bMenuLoop = TRUE;
		POINT pt;
		BOOL bFirst = TRUE;
		BOOL bReleaseCapture = TRUE;
		BOOL bDispLastMessage = TRUE;
		while (TRUE)
		{
			if (GetCapture() != hWndPopup || !m_pBar->m_bMenuLoop)
			{
				TRACE(2, "TrackPopupEx exit on lost capture\n");
				bDispLastMessage = FALSE;
				bReleaseCapture = FALSE;
				break;
			}

			if (m_pBar->m_bWhatsThisHelp)
				SetCursor(LoadCursor(NULL, IDC_HELP));

			GetMessage(&msg, NULL, 0, 0);
			TRACE1(1, "TrackPopUpMsg %X\n", msg.message);

			if (VARIANT_FALSE == bpV1.m_vbVisible)
				break;
			
			switch (msg.message)
			{
			// handle movement/accept messages
			case WM_CANCELMODE:
				bDispLastMessage=FALSE;
				goto ExitLoop;

			case WM_LBUTTONDOWN:
				TRACE(2, "WM_LBUTTONDOWN\n");
				
				GetCursorPos(&pt);
				if (abs(pt.x - ptInit.x) < nDblClkDistanceCx && abs(pt.y - ptInit.y) < nDblClkDistanceCy)
				{
					// check time elapsed
					if ((GetTickCount() - dwInitTime) <= GetDoubleClickTime())
						msg.message = WM_LBUTTONDBLCLK;
				}

			case WM_LBUTTONUP:
			case WM_RBUTTONUP:
			case WM_RBUTTONDOWN:
			case WM_MOUSEMOVE:
			case WM_NCMOUSEMOVE:
			case WM_LBUTTONDBLCLK:
				{
					HWND hWndTarget;
					
					//
					// pt is returned relative to hWndTarget
					//
					
					QueryPopupWin(hWndTarget, pt); 

					//
					// Send old target a mouse message so that it deactivates selection
					//

					if (hWndPrevTarget && hWndPrevTarget != hWndTarget) 
					{
						POINT ptClient = pt;
						ScreenToClient(hWndPrevTarget, &ptClient);
						SendMessage(hWndPrevTarget,
									WM_MOUSEMOVE,
									0,
									MAKELONG(ptClient.x, ptClient.y));
					}

					if (hWndTarget)
					{
						POINT ptClient = pt;
						ScreenToClient(hWndTarget, &ptClient);
						if (pPopup->Flip() && 
							bFirst && 
							(WM_MOUSEMOVE == msg.message || WM_LBUTTONUP == msg.message) &&	CBar::eSysMenu & m_pBar->m_dwMdiButtons)
						{
							bFirst = FALSE;
							break;
						}
						
						m_pBar->m_nPopupExitAction = 0;
						if (IsWindow(hWndTarget))
						{
							SendMessage(hWndTarget,
										msg.message,
										0,
										MAKELONG(ptClient.x, ptClient.y));
						}
						if (NULL == hWndPopup || !IsWindow(hWndPopup))
						{
							bDispLastMessage=FALSE;
							goto ExitLoop;
						}

						if (WM_LBUTTONDOWN == msg.message)
						{
							if (CPopupWin::PEA_DETACH == m_pBar->m_nPopupExitAction)
							{
								CPopupWin* pPopupWin = NULL;
								SendMessage(hWndTarget, GetGlobals().WM_POPUPWINMSG, 1, (LPARAM)&pPopupWin);
								if (pPopupWin)
								{
									HRESULT hResult = pPopupWin->m_pBand->Clone((IBand**)&pDetachBand);
									m_pBar->DestroyDetachBandOf(pPopupWin->m_pBand);
									if (pDetachBand)
									{
										pDetachBand->m_pTools->Release();
										pDetachBand->m_pTools = pPopupWin->m_pBand->m_pTools;
										pDetachBand->m_pTools->AddRef();
										pDetachBand->bpV1.m_vbDetached = VARIANT_TRUE;
										pDetachBand->m_daPrevDockingArea = bpV1.m_daDockingArea;
										pDetachBand->bpV1.m_nDockLine = 0;
										pDetachBand->bpV1.m_nDockOffset = 0;
										pDetachBand->bpV1.m_daDockingArea = ddDAFloat;
										pDetachBand->bpV1.m_vbVisible = VARIANT_TRUE;
										pDetachBand->bpV1.m_btBands = ddBTNormal;
										pDetachBand->bpV1.m_nCreatedBy = 2;
										pDetachBand->OnChildBandChanged(0);
										pDetachBand->bpV1.m_rcFloat = pDetachBand->GetOptimalFloatRect(FALSE);

										GetCursorPos(&m_pBar->m_ptPea);

										pDetachBand->bpV1.m_rcFloat.Offset(-pDetachBand->bpV1.m_rcFloat.left + m_pBar->m_ptPea.x,
																		   -pDetachBand->bpV1.m_rcFloat.top + m_pBar->m_ptPea.y);

										m_pBar->m_ptPea.x = 5;
										m_pBar->m_ptPea.y = - CMiniWin::eMiniWinCaptionHeight / 2;
										hResult = m_pBar->m_pBands->InsertBand(-1, pDetachBand);
										m_pBar->InternalRecalcLayout();
										bDispLastMessage = FALSE;
									}
									TRACE(2, "ExitP LButtonDown popup detached\n");
								}
								goto ExitLoop;
							}
						}

						BOOL bVal = FALSE;
						if (!IsWindow(hWndTarget))
						{
							bDispLastMessage = FALSE;
							goto ExitLoop;
						}

						SendMessage(hWndTarget, GetGlobals().WM_POPUPWINMSG, 0, (LPARAM)&bVal);
						if (!bVal)
						{
							bDispLastMessage = FALSE;
							goto ExitLoop;
						}
					}
					else if (WM_MOUSEMOVE != msg.message)
					{
						POINT ptCursor;
						GetCursorPos(&ptCursor);
						if (WM_LBUTTONDBLCLK == msg.message)
						{
							if (prcTrack && (PtInRect(prcTrack, ptCursor)))
							{
								if (pbDblClick)
									*pbDblClick = TRUE;
								goto ExitLoop;
							}
						}
						
						if ((WM_LBUTTONDOWN == msg.message || WM_RBUTTONDOWN == msg.message) && 
							(NULL == prcTrack || (!PtInRect(prcTrack, ptCursor))))						
						{
							goto ExitLoop;
						}
					}
					hWndPrevTarget = hWndTarget;
				}
				break;

			case WM_KEYDOWN:
				{
					BOOL bKeyUsed = FALSE;
					switch (msg.wParam)
					{
					case VK_ESCAPE:
						bDispLastMessage = FALSE;
						goto ExitLoop;

					default:
						SendMessage(pPopup->m_hWnd, msg.message, msg.wParam,(LPARAM)&bKeyUsed);
						if (NULL == hWndPopup)
							goto ExitLoop;

						if (bKeyUsed)
						{
							if (pPopup->m_pBand->m_pChildPopupWin)
								::PeekMessage(&msg, NULL, WM_MOUSEMOVE, WM_MOUSEMOVE, PM_REMOVE);
							bDispLastMessage = FALSE;
						}
						break;
					}
				}
				break;

			// Just dispatch rest of the messages
			default:
				DispatchMessage(&msg);
				if (NULL == m_pPopupWin)
					goto ExitLoop;
				break;
			}
		}

ExitLoop:
		m_pBar->m_bMenuLoop = FALSE;
		if (GetCapture() == hWndPopup)
			ReleaseCapture();
		
		bpV1.m_vbDetached = vbDetach;

		switch (msg.message)
		{
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			{
				POINT p;
				GetCursorPos(&p);
				HWND hWndClickTarget = WindowFromPointEx(p);
				if (IsWindowVisible(hWndClickTarget) && IsWindowEnabled(hWndClickTarget) && bDispLastMessage)
				{
					if (hWndClickTarget == m_pBar->GetDockWindow())
						DispatchMessage(&msg);	
					else
					{
						ScreenToClient(hWndClickTarget,&p);
						if (m_pBar->m_hWndModifySelection != hWndClickTarget)
							::PostMessage(hWndClickTarget, msg.message, msg.wParam, MAKELONG(p.x,p.y));
					}
				}
			}
			break;

		default:
			if (bDispLastMessage)
				DispatchMessage(&msg);
			break;
		}

		if (IsWindow(hWndPopup))
		{
			pPopup->DestroyWindow();
			pPopup = NULL;
		}
		m_bPopupWinLock = FALSE;

		if (pDetachBand)
		{
			m_pDockMgr->StartDrag(pDetachBand, m_pBar->m_ptPea);
			pDetachBand->Release();
			m_pBar->m_nPopupExitAction = 0;
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

//
// OnChildBandChanged
//

void CBand::OnChildBandChanged(short nNewPage)
{
	if (NULL == m_pChildBands)
		return;

	m_pChildBands->SetCurrentChildBand(nNewPage);
	
	//
	// Get the current page for the band
	//

	CBand* pChildBand = m_pChildBands->GetCurrentChildBand();
	if (NULL == pChildBand)
		return;

	//
	// If we are load don't RecalcLayout and don't fire the ChildBandChange Event
	//

	if (m_bLoading)
		return;

	try
	{
		m_pBar->FireChildBandChange((Band*)pChildBand);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}

	switch (bpV1.m_cbsChildStyle)
	{
	case ddCBSSlidingTabs:
		if (0 == pChildBand->m_nFormControlCount)
			m_bDrawAnimated = TRUE;
		break;
	
	default:
		break;
	}
	PostMessage(m_pBar->m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0); 
}

//
// GetActiveTools
//

void CBand::GetActiveTools(CTools** pTools)
{
	assert(pTools);
	try
	{
		if (ddCBSNone == bpV1.m_cbsChildStyle)
			*pTools = m_pTools;
		else
			m_pChildBands->GetCurrentChildBand()->get_Tools((Tools**)pTools);
	}
	CATCH
	{
		*pTools = NULL;
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// SetPopupIndex
//
// pToolToBeDropped is filled in at customization time
//

void CBand::SetPopupIndex(int nNewIndex, TypedArray<CTool*>* pToolsToBeDropped, BOOL bSelectFirstItem)
{
	try
	{
		static BOOL bReEntry = FALSE;
		if (bReEntry)
			return;

		m_pBar->HideToolTips(TRUE);

		if (VARIANT_FALSE == bpV1.m_vbVisible)
			return;

		if (-1 == nNewIndex && m_pBar->m_pPopupRoot == this)
			m_pBar->m_pPopupRoot = NULL;

		if (ddBTPopup != bpV1.m_btBands && 
			ddDAPopup != bpV1.m_daDockingArea &&
			m_pBar->m_pPopupRoot &&
			m_pBar->m_pPopupRoot != this)
		{
			m_pBar->m_pPopupRoot->SetPopupIndex(-1);
		}

		if (nNewIndex >= GetVisibleToolCount())
			nNewIndex = GetVisibleToolCount() - 1;
		
		if (nNewIndex == m_nPopupIndex)
			return;

		if (-1 != m_nPopupIndex && m_pChildPopupWin)
		{
			m_nPopupIndex = -1;
			bReEntry = TRUE;
			if (m_pChildPopupWin->IsWindow())
			{
				if (m_pBar && m_pBar->m_pDesigner)
					m_pChildPopupWin->PostMessage(WM_CLOSE);
				else
					BOOL bDestroy = m_pChildPopupWin->DestroyWindow();
			}
			m_pChildPopupWin = NULL;
			bReEntry = FALSE;
		}

		if (-1 != nNewIndex)
		{
			CTool* pTool = m_pTools->GetVisibleTool(nNewIndex);
			
			if (m_pBar->m_pCustomizeDragLock == pTool)
				return;

			assert(pTool);
			if (NULL == pTool)
				return;

			// If the current tool is being dragged in customization and has subbands. 
			// We cannot drop it into subbands so we have prevent opening popups

			if (pTool->HasSubBand())
			{
				if (m_bstrName && 0 == wcscmp(pTool->m_bstrSubBand, m_bstrName))
					return;
				
				if ((m_pBar->m_bAltCustomization || m_pBar->m_bCustomization || m_pBar->m_bWhatsThisHelp) && 
					m_pBar->NeedsCycleShutDown(pTool, pToolsToBeDropped))
				{
					return;
				}

				CBand* pSubBand = m_pBar->FindSubBand(pTool->m_bstrSubBand);
				if (pSubBand)
				{
					CRect rcTool;
					if (GetToolScreenRect(nNewIndex, rcTool) && pSubBand->CheckPopupOpen())
					{
						try
						{
							// The BandOpen event might call Recalcbands which 
							// sets pressed=false for tools (precaution agains lots of tools changing)

							BOOL bPrevPressed = pTool->m_bPressed; 
							m_nPopupIndex = nNewIndex;
							m_pChildPopupWin = new CPopupWin(pSubBand, IsVertical());
							assert(m_pChildPopupWin);
							if (m_pChildPopupWin)
							{
								int nDirection = ddPopupMenuLeftAlign;
								if (CTool::ddTTMoreTools == pTool->tpV1.m_ttTools)
								{
									nDirection = ddPopupMenuRightAlign;
									if (IsVertical())
										rcTool.Offset(-rcTool.Width(), 0);
									else
										rcTool.Offset(rcTool.Width() + eBevelBorder2, 0);
								}
								if (!m_pChildPopupWin->PreCreateWin(pTool,
																    rcTool,
																    (CPopupWin::FlipDirection)(ddDALeft == bpV1.m_daDockingArea || ddDARight == bpV1.m_daDockingArea || NULL != m_pPopupWin),
																	ddPMDisabled != m_pBar->bpV1.m_pmMenus,
																	bpV1.m_dwFlags & ddBFPopupFlipUp,
																	nDirection))
								{
									m_nPopupIndex = -1;
									delete m_pChildPopupWin;
									m_pChildPopupWin = NULL;
								}

								InvalidateToolRect(nNewIndex);

								if (!m_pChildPopupWin->CreateWin(m_pBar->GetParentWindow()))
								{
									m_nPopupIndex = -1;
									delete m_pChildPopupWin;
									m_pChildPopupWin = NULL;
								}
								else
								{
									m_pBar->m_theToolStack.Push(pTool);
									if (bSelectFirstItem)
										m_pChildPopupWin->m_nCurSel = 0;

									if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
									{
										if (ddBTNormal == pTool->m_pBand->bpV1.m_btBands && ddTTButton == pTool->tpV1.m_ttTools)
											pSubBand->m_pParentPopupWin = m_pChildPopupWin;
										if (GetToolScreenRect(m_nPopupIndex, rcTool))
											m_pBar->m_pToolShadow->Show(pTool, rcTool);
										m_pChildPopupWin->Shadow().SetWindowPos(NULL, 
																			    0, 
																			    0, 
																			    0, 
																			    0,
																			    SWP_SHOWWINDOW|SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
									}
									m_pChildPopupWin->SetWindowPos(HWND_TOPMOST, 
																   0, 
																   0, 
																   0, 
																   0,
																   SWP_SHOWWINDOW|SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOSIZE);
									if (ddBTPopup != bpV1.m_btBands)
										m_pBar->m_pPopupRoot = this;
								}
							}
							pTool->m_bPressed = bPrevPressed;
						}
						CATCH
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}			
					}
				}						
			}
		}
		m_nPopupIndex = nNewIndex;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
}

//
// AddSubBandsToArray
// 
// returns true if cycle detected
//

BOOL CBand::AddSubBandsToArray(TypedArray<CBand*>* pArray, 
							   CBand**			   ppParentOfProblemBand, 
							   BOOL				   bAddIfVisible)
{
	if (m_bCycleMark)
	{
		TRACE1(1, "cycle item0 %ls",m_bstrName);
		return TRUE;
	}

	m_bCycleMark = TRUE;
	CBand* pBand;
	CTool* pTool;
	int nToolCount = m_pTools->GetVisibleToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = m_pTools->GetVisibleTool(nTool);
		if (!pTool->HasSubBand())
			continue;

		pBand = m_pBar->FindBand(pTool->m_bstrSubBand);
		if (NULL == pBand)
			continue;

		if (!bAddIfVisible || VARIANT_TRUE == bpV1.m_vbVisible)
		{
			// Eliminate infinite loop recursive call
			if (!pBand->m_bCycleMark) 
			{
	 			if (pBand->AddSubBandsToArray(pArray,
											  ppParentOfProblemBand,
											  bAddIfVisible))
				{
					if (ppParentOfProblemBand && NULL == *ppParentOfProblemBand)
						*ppParentOfProblemBand = pBand;
					return TRUE; 
				}
			}
			else
			{
				if (ppParentOfProblemBand && NULL == *ppParentOfProblemBand)
					*ppParentOfProblemBand = this;

				TRACE1(1, "cycle item0 %ls",m_bstrName);
				return TRUE;
			}
		}
	}
	HRESULT hResult = pArray->Add(this);
	assert(SUCCEEDED(hResult));
	return FALSE;
}

//
// CheckPopupOpen
//

BOOL CBand::CheckPopupOpen()
{
	// If this band is already in use on a DockingArea don't open Popup
	if (ddBTPopup == bpV1.m_btBands)
		return TRUE;

	switch (bpV1.m_daDockingArea)
	{
	case ddDALeft:
	case ddDARight:
	case ddDATop:
	case ddDABottom:
		m_pBar->m_theErrorObject.SendError(IDERR_OPENINVALIDPOPUP, 0);
		return FALSE;
	}
	return TRUE;
}

//
// ContainsSubBandOfToolTree 
//

BOOL CBand::ContainsSubBandOfToolTree(CTool* pTool, CBand** ppParentOfProblemBand)
{
	// Create subband list of this band and the pTool->subband. If there is any collision return TRUE
	m_pBar->ResetCycleMarks();
	
	if (!pTool->HasSubBand())
		return FALSE;

	CBand* pSubBand = m_pBar->FindSubBand(pTool->m_bstrSubBand);
	if (NULL == pSubBand)
		return FALSE;

	// First mark all recursive subbands of this tool.
	TypedArray<CBand*> faSubBandArray;
	if (pSubBand->AddSubBandsToArray(&faSubBandArray, ppParentOfProblemBand, TRUE))
		return TRUE;

	// Now move through self(droptarget) and try to find a problem area
	if (AddSubBandsToArray(&faSubBandArray, ppParentOfProblemBand, TRUE))
		return TRUE;

	return FALSE;
}

//
// GetVisibleToolCount
//

int CBand::GetVisibleToolCount()
{
	if (m_bDesignTime)
		return m_pTools->GetToolCount();
	else
		return m_pTools->GetVisibleToolCount();
}

//
// ResetToolPressedState
//

void CBand::ResetToolPressedState()
{
	if (ddBTMenuBar == bpV1.m_btBands || ddBTChildMenuBar == bpV1.m_btBands)
		return;

	CTool* pTool;
	int nCount = m_pTools->GetToolCount();
	for (int nTool = 0; nTool < nCount; nTool++)
	{
		pTool = m_pTools->GetTool(nTool);
		if (m_nPopupIndex != nTool && pTool != m_pBar->m_pActiveTool)
			pTool->m_bPressed = FALSE;
	}
}

//
// CalcDropLoc
//

void CBand::CalcDropLoc(POINT pt, CBar::DropInfo* pDropInfo)
{
	try
	{
		memset(pDropInfo, 0, sizeof(CBar::DropInfo));
		pDropInfo->nDropIndex = -1;

		if (!(ddBFCustomize & bpV1.m_dwFlags) && NULL == m_pBar->m_pDesigner)
			return;

		pDropInfo->pTool = HitTestTools(pDropInfo->pBand, pt, pDropInfo->nToolIndex);
		if (NULL == pDropInfo->pBand)
			pDropInfo->pBand = this;

		if (NULL == pDropInfo->pTool && -1 == pDropInfo->nToolIndex)
			pDropInfo->ptTarget = pt;

		switch (bpV1.m_daDockingArea)
		{
		case ddDATop:
		case ddDABottom:
		case ddDALeft:
		case ddDARight:
			DSGCalcDropIndex(pt.x, 
							 pt.y, 
							 &pDropInfo->nDropIndex, 
							 &pDropInfo->nDropDir);
			pDropInfo->rcBand = m_pRootBand->m_rcDock;
			pDropInfo->bFloating = FALSE;
			break;

		case ddDAFloat:
			DSGCalcDropIndex(pt.x, 
							 pt.y, 
							 &pDropInfo->nDropIndex, 
							 &pDropInfo->nDropDir);
			m_pFloat->GetClientRect(pDropInfo->rcBand);
			pDropInfo->bFloating = TRUE;
			break;

		case ddDAPopup:
			DSGCalcDropIndex(pt.x,
							 pt.y,
							 &pDropInfo->nDropIndex,
							 &pDropInfo->nDropDir);
			m_pPopupWin->GetClientRect(pDropInfo->rcBand);
			pDropInfo->rcBand.Inflate(-3, -3);
			pDropInfo->bPopup = TRUE;
			break;
		}
		pDropInfo->ptTarget = pt;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
}

//
// This is used for creating the Tool Customiztion Popup
//

struct CustomEditTool
{
	UINT nNameResId;
	UINT nCaptionResId;
} theEditTools[]=
{
	{IDS_TOOLRESET,			IDS_TOOLRESETCAPTION},
	{IDS_TOOLDELETE,		IDS_TOOLDELETECAPTION},
	{0,						0},
	{IDS_TOOLEDITNAME,		IDS_TOOLEDITNAMECAPTION},
	{0,						0},
	{IDS_TOOLCOPYIMAGE,     IDS_TOOLCOPYIMAGECAPTION},
	{IDS_TOOLPASTEIMAGE,    IDS_TOOLPASTEIMAGECAPTION},
	{0,						0},
	{IDS_TOOLDEFAULTSTYLE,  IDS_TOOLDEFAULTSTYLECAPTION},
	{IDS_TOOLTEXTONLY,		IDS_TOOLTEXTONLYCAPTION},
	{IDS_TOOLIMAGEONLY,		IDS_TOOLIMAGEONLYCAPTION},
	{IDS_TOOLIMAGEANDTEXT,  IDS_TOOLIMAGEANDTEXTCAPTION}
};

//
// SetupEditToolPopup
//

BOOL CBand::SetupEditToolPopup(const CTool* pToolIn)
{
	try
	{
		CTool* pTool;
		if (NULL == m_pBar->m_pEditToolPopup)
		{
			m_pBar->m_pEditToolPopup = CBand::CreateInstance(NULL);
			if (NULL == m_pBar->m_pEditToolPopup)
				return FALSE;

			m_pBar->m_pEditToolPopup->put_Caption(L"Edit Tool");
			m_pBar->m_pEditToolPopup->put_Name(L"popEditTool");
			m_pBar->m_pEditToolPopup->SetOwner(m_pBar, TRUE);
			m_pBar->m_pEditToolPopup->bpV1.m_btBands = ddBTPopup;
			m_pBar->m_pEditToolPopup->bpV1.m_nCreatedBy = ddCBSystem;
			m_pBar->m_pEditToolPopup->bpV1.m_daDockingArea = ddDAPopup;
			m_pBar->m_pEditToolPopup->bpV1.m_dwFlags = ddBFDockLeft|ddBFDockTop|ddBFDockRight|ddBFDockBottom|ddBFFloat;

			m_pBar->m_bToolCreateLock = TRUE;

			int nSize = sizeof(theEditTools)/sizeof(CustomEditTool);
			for (int nTool = 0; nTool < nSize; nTool++)
			{
				m_pBar->m_pEditToolPopup->m_pTools->CreateTool((ITool**)&pTool);
				if (pTool)
				{
					pTool->tpV1.m_nToolId = CBar::eEditToolId + nTool;
					
					LPTSTR szCustom = NULL;
					if (0 != theEditTools[nTool].nNameResId)
					{
						szCustom = LoadStringRes(theEditTools[nTool].nNameResId);
						MAKE_WIDEPTR_FROMTCHAR(wCustom, szCustom);
						pTool->put_Name(wCustom);
						if (0 != theEditTools[nTool].nCaptionResId)
						{
							szCustom = LoadStringRes(theEditTools[nTool].nCaptionResId);
							MAKE_WIDEPTR_FROMTCHAR(wCustom, szCustom);
							pTool->put_Caption(wCustom);
						}
					}
					else
					{
						pTool->put_Name(L"miSeparator");
						pTool->tpV1.m_ttTools = ddTTSeparator;
					}

					m_pBar->m_pEditToolPopup->m_pTools->InsertTool(-1, pTool, VARIANT_FALSE);
					pTool->Release();
				}
			}
		}
		
		m_pBar->m_bToolCreateLock = TRUE;

		int nSize = m_pBar->m_pEditToolPopup->m_pTools->GetToolCount();
		for (int nTool = 0; nTool < nSize; nTool++)
		{
			pTool = m_pBar->m_pEditToolPopup->m_pTools->GetTool(nTool);
			switch (theEditTools[nTool].nNameResId)
			{
			case IDS_TOOLEDITNAME:
				pTool->tpV1.m_ttTools = ddTTEdit;
				pTool->put_Text(pToolIn->m_bstrCaption);
				break;
			
			case IDS_TOOLCOPYIMAGE:
				if (-1 == pToolIn->tpV1.m_nImageIds[0])
					pTool->tpV1.m_vbEnabled = VARIANT_FALSE;
				break;
			
			case IDS_TOOLPASTEIMAGE:
				pTool->tpV1.m_vbEnabled = IsClipboardFormatAvailable(CF_BITMAP) ? VARIANT_TRUE : VARIANT_FALSE;
				break;
			
			case IDS_TOOLDEFAULTSTYLE:
				pTool->tpV1.m_vbChecked = pToolIn->tpV1.m_tsStyle == ddSStandard ? VARIANT_TRUE : VARIANT_FALSE;
				break;
			
			case IDS_TOOLTEXTONLY:
				pTool->tpV1.m_vbChecked = pToolIn->tpV1.m_tsStyle == ddSText ? VARIANT_TRUE : VARIANT_FALSE;
				break;
			
			case IDS_TOOLIMAGEONLY:
				pTool->tpV1.m_vbChecked = pToolIn->tpV1.m_tsStyle == ddSIcon ? VARIANT_TRUE : VARIANT_FALSE;
				break;
			
			case IDS_TOOLIMAGEANDTEXT:
				pTool->tpV1.m_vbChecked = pToolIn->tpV1.m_tsStyle == ddSIconText ? VARIANT_TRUE : VARIANT_FALSE;
				break;
			}
		}
		m_pBar->m_bToolCreateLock = FALSE;
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
// CustomEditTools
//

void CBand::CustomEditTools(const POINT& ptScreen, const CTool* pToolIn)
{
	try
	{
		if (SetupEditToolPopup(pToolIn))
		{
			m_pBar->m_mcInfo.pBand = m_pBar->m_diCustSelection.pBand;
			m_pBar->m_mcInfo.pTool = m_pBar->m_diCustSelection.pTool;
			VARIANT x;
			VARIANT y;
			VARIANT vFlags;
			x.vt = VT_I2;
			y.vt = VT_I2;
			vFlags.vt = VT_I2;
			x.iVal = (short)ptScreen.x;
			y.iVal = (short)ptScreen.y;
			vFlags.iVal = (int)ddPopupMenuLeftAlign;
			m_pBar->m_pEditToolPopup->PopupMenu(vFlags, x, y);
			m_pBar->m_pEditToolPopup->Release();	
			m_pBar->m_pEditToolPopup = NULL;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
}

//
// OnCustomMouseDown
//

void CBand::OnCustomMouseDown(UINT nMsg, POINT pt)
{
	try
	{
		if (WM_LBUTTONDOWN == nMsg || WM_RBUTTONDOWN == nMsg)
		{
			TRACE(5, "Here - 1\n");
			CBar::DropInfo diOld;
			memcpy(&diOld, &m_pBar->m_diCustSelection, sizeof(CBar::DropInfo));

			CalcDropLoc(pt, &m_pBar->m_diCustSelection);
				
			try
			{
				if (diOld.pTool != m_pBar->m_diCustSelection.pTool)
				{
					if (diOld.pTool)
					{
						m_pBar->m_pToolShadow->Hide();
						if (diOld.bPopup)
						{
							if (diOld.pBand->m_pPopupWin)
								diOld.pBand->m_pPopupWin->RefreshTool(diOld.pTool);
						}
						else
							diOld.pBand->ToolNotification(diOld.pTool, TNF_VIEWCHANGED);
					}
					if (m_pBar->m_diCustSelection.pTool)
					{
						if (m_pBar->m_diCustSelection.bPopup)
						{
							if (m_pBar->m_diCustSelection.pBand->m_pPopupWin)
								m_pBar->m_diCustSelection.pBand->m_pPopupWin->RefreshTool(m_pBar->m_diCustSelection.pTool);
						}
						ToolNotification(m_pBar->m_diCustSelection.pTool, TNF_VIEWCHANGED);
						if (m_pBar->m_pDesignerNotify)
						{
							m_pBar->m_pDesignerNotify->ToolSelected(m_bstrName, 
																	(ddCBSNone != bpV1.m_cbsChildStyle ? m_pChildBands->GetCurrentChildBand()->m_bstrName : NULL), 
																	m_pBar->m_diCustSelection.pTool->tpV1.m_nToolId);
						}
					}
					if (m_pBar->m_diCustSelection.pBand)
					{
						if (NULL == m_pBar->m_diCustSelection.pTool)
							m_pBar->m_diCustSelection.pBand->SetPopupIndex(-1);

						if (m_pBar->m_diCustSelection.pTool && WM_LBUTTONDOWN == nMsg)
							m_pBar->m_diCustSelection.pBand->SetPopupIndex(m_pBar->m_diCustSelection.nToolIndex);
					}
				}
				else if (diOld.pBand && (ddBTMenuBar == diOld.pBand->bpV1.m_btBands || ddBTChildMenuBar == diOld.pBand->bpV1.m_btBands))
				{
					if (diOld.pTool && diOld.pTool->HasSubBand())
					{
						TRACE(5, "Here - 2\n");
						if (diOld.pTool->m_pBand->m_pChildPopupWin)
						{
							SetPopupIndex(-1);
							m_pBar->m_pToolShadow->Hide();
						}
						else
							SetPopupIndex(diOld.nToolIndex);
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}			
		}

		if (m_pBar->m_diCustSelection.pTool)
		{
			HWND hWndCapture;
			CRect rcBand;
			if (m_pBar->m_hWndModifySelection && ::IsWindow(m_pBar->m_hWndModifySelection))
				EnableWindow(m_pBar->m_hWndModifySelection, TRUE);
			try 
			{
				switch (nMsg)
				{
				case WM_MOUSEMOVE:
					{
						switch (m_pBar->m_diCustSelection.pTool->tpV1.m_ttTools)
						{
						case ddTTCombobox:
						case ddTTEdit:
							{
								CRect rcTool;
								POINT ptCursor;
								GetCursorPos(&ptCursor);
								if (GetToolScreenRect(m_pBar->m_diCustSelection.nToolIndex, rcTool) && PtInRect(&rcTool, ptCursor))
								{
									if (ptCursor.x < rcTool.left + 4 || ptCursor.x > rcTool.right - 4)
										SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDC_COMBOSPLIT)));
									else
										SetCursor(LoadCursor(NULL, IDC_ARROW));
								}
							}
							break;
						}
					}
					break;

				case WM_LBUTTONDOWN:
					{
						switch (m_pBar->m_diCustSelection.pTool->tpV1.m_ttTools)
						{
						case ddTTCombobox:
						case ddTTEdit:
							{
								CRect rcBound;
								GetWindowRect(m_pBar->GetParentWindow(), &rcBound);
								CRect rcTool;
								POINT ptCursor;
								GetCursorPos(&ptCursor);
								if (GetToolScreenRect(m_pBar->m_diCustSelection.nToolIndex, rcTool) && PtInRect(&rcTool, ptCursor))
								{
									BOOL bEnd = FALSE;
									if (ptCursor.x < rcTool.left + 4 || ptCursor.x > rcTool.right - 4)
									{
										if (ptCursor.x > rcTool.right - 4)
											bEnd = TRUE;


										POINT ptCursorEnd;
										SetCursor(LoadCursor(g_hInstance, MAKEINTRESOURCE(IDC_COMBOSPLIT)));
										if (GetBandRect(hWndCapture, rcBand))
										{
											HDC hDC = GetDC(hWndCapture);

											SetCapture(hWndCapture);
											BOOL bProcessing = TRUE;
											MSG msg;
											while (GetCapture() == hWndCapture && bProcessing)
											{
												GetMessage(&msg, NULL, 0, 0);
												switch (msg.message)
												{
												case WM_KEYDOWN:
													if (VK_ESCAPE == msg.wParam)
														bProcessing = FALSE;
													break;

												case WM_RBUTTONUP:
													bProcessing = FALSE;
													break;

												case WM_MOUSEMOVE:
													GetCursorPos(&ptCursorEnd);
													if (ptCursorEnd.x > rcBound.right - 20)
														ptCursorEnd.x = rcBound.right - 20;
													else if (ptCursorEnd.x < rcBound.left + 20)
														ptCursorEnd.x = rcBound.left + 20;
													DSGDrawSelection((OLE_HANDLE)hDC, rcTool.left, rcTool.top, rcTool.Width(), -rcTool.Height());
													if (bEnd)
														rcTool.right = ptCursorEnd.x;
													else
														rcTool.left = ptCursorEnd.x;
													DSGDrawSelection((OLE_HANDLE)hDC, rcTool.left, rcTool.top, rcTool.Width(), -rcTool.Height());
													break;

												case WM_LBUTTONUP:
													{
														bProcessing = FALSE;
														m_pBar->m_diCustSelection.pTool->tpV1.m_nWidth = rcTool.Width();
														m_pBar->RecalcLayout();
														DispatchMessage(&msg);
													}
													break;

												default:
													DispatchMessage(&msg);
													break;
												}
											}
											if (GetCapture() == hWndCapture)
												ReleaseCapture();
											ReleaseDC(hWndCapture, hDC);
										}
										return;
									}
									else
										SetCursor(LoadCursor(NULL, IDC_ARROW));
								}
							}
							break;
						}

						m_pBar->FireCustomizeToolClick((Tool*)m_pBar->m_diCustSelection.pTool);

						if (GetBandRect(hWndCapture, rcBand))
						{
							POINT ptDragStart;
							GetCursorPos(&ptDragStart);
							ScreenToClient(hWndCapture, &ptDragStart);

							BOOL bProcessing = TRUE;
							SetCapture(hWndCapture);
							MSG msg;
							DWORD nDragStartTick = GetTickCount();
							while (GetCapture() == hWndCapture && bProcessing)
							{
								GetMessage(&msg, NULL, 0, 0);
								switch (msg.message)
								{
								case WM_KEYDOWN:
									if (VK_ESCAPE == msg.wParam)
										bProcessing = FALSE;
									break;

								case WM_RBUTTONUP:
								case WM_LBUTTONUP:
									bProcessing = FALSE;
									break;

								case WM_MOUSEMOVE:
									{
										try
										{

											POINT pt = {LOWORD(msg.lParam), HIWORD(msg.lParam)};
											if ((GetTickCount()-nDragStartTick) < GetGlobals().m_nDragDelay && abs(pt.x-ptDragStart.x) < GetGlobals().m_nDragDist && abs(pt.y-ptDragStart.y) < GetGlobals().m_nDragDist)
												break;

											CToolDropSource* pSource = new CToolDropSource;
											if (NULL == pSource)
												return;

											TypedArray<ITool*> aTools;
											CTool* pTool = m_pBar->m_diCustSelection.pTool;
											HRESULT hResult = aTools.Add(pTool);
											if (SUCCEEDED(hResult))
											{
												CToolDataObject* pDataObject = new CToolDataObject(CToolDataObject::eActiveBarDragDropId, 
																								   (IActiveBar2*)m_pBar, 
																								   aTools, 
																								   m_bstrName, 
																								   (ddCBSNone != bpV1.m_cbsChildStyle ? m_pChildBands->GetCurrentChildBand()->m_bstrName : NULL));
												if (NULL == pDataObject)
												{
													m_pBar->m_pCustomizeDragLock = NULL; 
													pSource->Release();
													return;
												}

												m_pBar->m_pCustomizeDragLock = pTool; 
												DWORD dwEffect;
												HRESULT hResult = GetGlobals().m_pDragDropMgr->DoDragDrop((LPUNKNOWN)pDataObject, 
																										  (LPUNKNOWN)pSource, 
																										  DROPEFFECT_MOVE|DROPEFFECT_COPY, 
																										  &dwEffect);
												m_pBar->m_pCustomizeDragLock = NULL; 
												BOOL bFirst = DRAGDROP_S_CANCEL != hResult && E_UNEXPECTED != hResult;
												BOOL bSecond = (DROPEFFECT_MOVE == dwEffect || DROPEFFECT_NONE == dwEffect);
												if (bFirst && bSecond)
												{
													// Tool has moved so we need to delete it from this band
													SetPopupIndex(-1);
													if (ddCBSNone == bpV1.m_cbsChildStyle)
														hResult = m_pTools->DeleteTool(pTool);
													else
													{
														CBand* pChildBand = m_pChildBands->GetCurrentChildBand();
														assert(pChildBand);
														if (pChildBand)
															hResult = pChildBand->m_pTools->DeleteTool(pTool);
													}
													if (m_pBar->m_pDesigner)
													{
														m_pBar->m_pDesigner->SetDirty();
														if (m_pBar->m_pDesignerNotify)
														{
															m_pBar->m_pDesignerNotify->BandChanged((LPDISPATCH)(IBand*)this, 
																								   (LPDISPATCH)(IBand*)(ddCBSNone != bpV1.m_cbsChildStyle ? m_pChildBands->GetCurrentChildBand() : NULL),
																								   ddBandModified);
														}
													}
												}
												if (DRAGDROP_S_CANCEL != hResult)
												{
													m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
													m_pBar->RecalcLayout();
													Refresh();
												}
												pDataObject->Release();
												pSource->Release();
											}
										}
										CATCH
										{
											assert(FALSE);
											REPORTEXCEPTION(__FILE__, __LINE__)
										}			
									}
									break;

								default:
									// Just dispatch rest of the messages
									DispatchMessage(&msg);
									break;
								}
							}
							if (GetCapture() == hWndCapture)
								ReleaseCapture();
						}
					}
					break;

				case WM_RBUTTONDOWN:
					{
						CBar::DropInfo diOld;
						memcpy(&diOld, &m_pBar->m_diCustSelection, sizeof(CBar::DropInfo));

						CalcDropLoc(pt, &m_pBar->m_diCustSelection);
							
						try
						{
							if (diOld.pTool != m_pBar->m_diCustSelection.pTool)
							{
								if (diOld.pTool)
								{
									if (diOld.bPopup)
									{
										if (diOld.pBand->m_pPopupWin)
											diOld.pBand->m_pPopupWin->RefreshTool(diOld.pTool);
									}
									else
										diOld.pBand->ToolNotification(diOld.pTool, TNF_VIEWCHANGED);
								}
								if (m_pBar->m_diCustSelection.pTool)
								{
									if (m_pBar->m_diCustSelection.bPopup)
									{
										if (m_pBar->m_diCustSelection.pBand->m_pPopupWin)
											m_pBar->m_diCustSelection.pBand->m_pPopupWin->RefreshTool(m_pBar->m_diCustSelection.pTool);
									}
									ToolNotification(m_pBar->m_diCustSelection.pTool, TNF_VIEWCHANGED);
									try
									{
										if (m_pBar->m_pDesignerNotify)
										{
											m_pBar->m_pDesignerNotify->ToolSelected(m_bstrName, 
																					(ddCBSNone != bpV1.m_cbsChildStyle ? m_pChildBands->GetCurrentChildBand()->m_bstrName : NULL), 
																					m_pBar->m_diCustSelection.pTool->tpV1.m_nToolId);
										}
									}
									CATCH
									{
										assert(FALSE);
										REPORTEXCEPTION(__FILE__, __LINE__)
									}			
								}
								if (m_pBar->m_diCustSelection.pBand)
								{
									if (NULL == m_pBar->m_diCustSelection.pTool)
										m_pBar->m_diCustSelection.pBand->SetPopupIndex(-1);

									if (m_pBar->m_diCustSelection.pTool && WM_LBUTTONDOWN == nMsg)
										m_pBar->m_diCustSelection.pBand->SetPopupIndex(m_pBar->m_diCustSelection.nToolIndex);
								}
							}
						}
						CATCH
						{
							assert(FALSE);
							REPORTEXCEPTION(__FILE__, __LINE__)
						}			
						m_pBar->FireCustomizeToolClick((Tool*)m_pBar->m_diCustSelection.pTool);
						if (m_pBar && NULL == m_pBar->m_pDesigner)
						{
							POINT ptScreen = {-1, -1};
							CustomEditTools(ptScreen, m_pBar->m_diCustSelection.pTool);
						}
					}
					break;
				}

				if (m_pBar->m_bAltCustomization)
					m_pBar->m_bAltCustomization = FALSE;
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}			
		}
		else
		{
			if (NULL != m_pBar->m_hWndModifySelection && IsWindow(m_pBar->m_hWndModifySelection))
				EnableWindow(m_pBar->m_hWndModifySelection, FALSE);

			if (WM_LBUTTONDOWN == nMsg || WM_RBUTTONDOWN == nMsg)
			{
				if (m_pBar->m_diCustSelection.pBand)
				{
					CBand* pBand = m_pBar->m_diCustSelection.pBand;

					// First check if one of the tabs got a hit for customization
					if (ddCBSNone != pBand->bpV1.m_cbsChildStyle)
					{
						int nHit;
						pBand->DSGCalcHitEx(m_pBar->m_diCustSelection.ptTarget.x,
											m_pBar->m_diCustSelection.ptTarget.y, 
											&nHit);
						if (-1 != nHit)
						{
							if (nHit != pBand->m_pChildBands->BTGetCurSel())
								pBand->OnChildBandChanged(nHit);

							if (m_pBar->m_pDesignerNotify)
							{
								m_pBar->m_pDesignerNotify->ToolSelected(m_bstrName, 
																		m_pChildBands->GetCurrentChildBand()->m_bstrName, 
																		0);
							}
							return;
						}
						else if (m_pBar->m_pDesignerNotify)
							m_pBar->m_pDesignerNotify->ToolSelected(m_bstrName, NULL, 0);
					}
					else if (m_pBar->m_pDesignerNotify)
						m_pBar->m_pDesignerNotify->ToolSelected(m_bstrName, NULL, 0);
				}
				if (WM_LBUTTONDOWN == nMsg && m_pBar->m_diCustSelection.pBand && ddBTPopup != m_pBar->m_diCustSelection.pBand->bpV1.m_btBands)
					m_pDockMgr->StartDrag(m_pBar->m_diCustSelection.pBand, pt);

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
// DrawAnimatedSlidingTabs
//

void CBand::DrawAnimatedSlidingTabs(CFlickerFree& ffObj, const HDC& hDC, CRect rcBand)
{
	if (!m_bDrawAnimated)
		return;
	m_bDrawAnimated = FALSE;

	const DWORD dwWaitTime = 30;
	DWORD dwTickCountStart;
	DWORD dwPaintTime;
	BOOL  bVertical = IsVertical();

	int nChildBands = m_pChildBands->GetChildBandCount();
	int nLineHeight = m_pChildBands->BTGetFontHeight();
	
	int nCurSpace = ((m_pChildBands->tpV1.m_nCurTab + 1) * nLineHeight);
	int nCurEndSpace = (nChildBands - m_pChildBands->tpV1.m_nCurTab - 1) * nLineHeight;
	int nPrevSpace = ((m_pChildBands->m_nPrevTab + 1) * nLineHeight);
	int nPrevEndSpace = (nChildBands - m_pChildBands->m_nPrevTab - 1) * nLineHeight;

	AdjustChildBandsRect(rcBand);
	if (m_pChildBands->m_nPrevTab < m_pChildBands->tpV1.m_nCurTab)
	{
		if (bVertical)
		{
			int nWidth = rcBand.Width();
			int nHeight = rcBand.Height();
			int nDelta = (nHeight - nPrevSpace - nPrevEndSpace) / 6;
			for (int nY = 1; nY < 6; nY++)
			{
				dwTickCountStart = GetTickCount(); 
				
				ffObj.Paint(hDC, 
							rcBand.left,
							rcBand.top + nHeight - nPrevEndSpace - (nY * nDelta), 
							nWidth,
							nCurSpace - nPrevSpace + nY * nDelta, 
							rcBand.left,
							rcBand.top + nPrevSpace);
				
				dwPaintTime = GetTickCount() - dwTickCountStart;
				
				if (dwPaintTime < dwWaitTime)
					Sleep(dwWaitTime-dwPaintTime);
			}
		}
		else
		{
			int nWidth = rcBand.Width();
			int nHeight = rcBand.Height();
			int nDelta = (nWidth - nPrevSpace - nPrevEndSpace) / 6;
			for (int nX = 1; nX < 6; nX++)
			{
				dwTickCountStart = GetTickCount(); 
				
				ffObj.Paint(hDC, 
							rcBand.left + nWidth - nPrevEndSpace - (nX * nDelta), 
							rcBand.top, 
							nCurSpace - nPrevSpace + (nX * nDelta),
							nHeight,
							rcBand.left + nPrevSpace,
							rcBand.top);
				
				dwPaintTime = GetTickCount()-dwTickCountStart;
				
				if (dwPaintTime < dwWaitTime)
					Sleep(dwWaitTime-dwPaintTime);
			}
		} 
	}
	else
	{ 
		if (bVertical)
		{
			int nHeight = rcBand.Height();
			int nHeight2 = nHeight - nPrevSpace - nPrevEndSpace - nPrevEndSpace + nCurEndSpace;
			int nWidth = rcBand.Width();
			int nDelta = (nHeight - nPrevSpace - nPrevEndSpace) / 6;
			for (int nY = 1; nY < 6; nY++)
			{
				dwTickCountStart = GetTickCount(); 
				
				ffObj.Paint(hDC, 
							rcBand.left, 
							rcBand.top,
							nWidth,
							nHeight - nCurEndSpace,
							rcBand.left,
							rcBand.top);
				
				m_ffOldBand.Paint(hDC, 
								  rcBand.left, 
								  rcBand.top + nCurSpace + (nY * nDelta), 
								  nWidth,
								  nHeight2 - (nY * nDelta),
								  0,
								  nCurSpace);
				
				dwPaintTime = GetTickCount() - dwTickCountStart;
				if (dwPaintTime < dwWaitTime)
					Sleep(dwWaitTime - dwPaintTime);
			}
		}
		else
		{
			int nHeight = rcBand.Height();
			int nWidth = rcBand.Width();
			int nWidth2 = nWidth - nPrevSpace - nPrevEndSpace;
			int nDelta = nWidth2 / 6;
			for (int nX = 1; nX < 6; nX++)
			{
				dwTickCountStart = GetTickCount(); 
				
				ffObj.Paint(hDC, 
							rcBand.left, 
							rcBand.top,
							nWidth - nCurEndSpace,
							nHeight,
							rcBand.left,
							rcBand.top);

				m_ffOldBand.Paint(hDC, 
								  rcBand.left + nCurSpace + (nX * nDelta), 
								  rcBand.top, 
								  nWidth - nCurSpace - nPrevEndSpace - (nX * nDelta),
								  nHeight,
								  nCurSpace,
								  0);
				
				dwPaintTime = GetTickCount()-dwTickCountStart;
				
				if (dwPaintTime < dwWaitTime)
					Sleep(dwWaitTime-dwPaintTime);
			}
		}
	}
	m_pChildBands->GetCurrentChildBand()->m_pTools->ShowTabbedWindowedTools();
}

//
// GetCustomToolSizes
//

BOOL CBand::GetCustomToolSizes(VARIANT& vWidth, VARIANT& vHeight)
{
	int nTool = 0;
	if (ddCBSNone != bpV1.m_cbsChildStyle)
	{
		CBand* pChildBand = m_pChildBands->GetCurrentChildBand();
		return pChildBand->GetCustomToolSizes(vWidth, vHeight);
	}

	TypedArray<CTool*> aCustomTools;
	BOOL    bResult = TRUE;
	CTool* pTool;

	//
	// Find custom tools
	//

	HRESULT hResult;
	int nCount = m_pTools->GetVisibleToolCount();
	for (nTool = 0; nTool < nCount; nTool++)
	{
		pTool = m_pTools->GetVisibleTool(nTool);

		if (ddTTForm != pTool->tpV1.m_ttTools && ddTTControl != pTool->tpV1.m_ttTools)
			continue;

		hResult = aCustomTools.Add(pTool);
		if (FAILED(hResult))
			continue;
	}
	nCount = aCustomTools.GetSize();

	//
	// Create the safearrays
	//

    SAFEARRAYBOUND rgsabound[1];    
	rgsabound[0].lLbound = 0;
    rgsabound[0].cElements = nCount;
	
	vWidth.vt = VT_ARRAY|VT_I4;
	vWidth.parray = SafeArrayCreate(VT_I4, 1, rgsabound);
	if (NULL == vWidth.parray)
		return FALSE;

	vHeight.vt = VT_ARRAY|VT_I4;
	vHeight.parray = SafeArrayCreate(VT_I4, 1, rgsabound);
	if (NULL == vHeight.parray)
	{
		SafeArrayDestroyDescriptor(vWidth.parray);
		return FALSE;
	}

	long* pWidth;
	hResult = SafeArrayAccessData(vWidth.parray, (void**)&pWidth);
	if (FAILED(hResult))
	{
		SafeArrayDestroyDescriptor(vWidth.parray);
		return FALSE;
	}

	long* pHeight;
	hResult = SafeArrayAccessData(vHeight.parray, (void**)&pHeight);
	if (FAILED(hResult))
	{
		SafeArrayUnaccessData(vWidth.parray);
		SafeArrayDestroyDescriptor(vWidth.parray);
		SafeArrayDestroyDescriptor(vHeight.parray);
		return FALSE;
	}

	//
	// Get the size of the custom tools
	//

	CRect rcTool;
	SIZE size;
	for (nTool = 0; nTool < nCount; nTool++)
	{
		pTool = aCustomTools.GetAt(nTool);
		size = pTool->m_rcTemp.Size();	
		PixelToTwips(&size, &size);
		pWidth[nTool] = size.cx;
		pHeight[nTool] = size.cy;
	}

	SafeArrayUnaccessData(vWidth.parray);
	SafeArrayUnaccessData(vHeight.parray);
	
	return bResult;
}

//
// SetCustomToolSizes
//

BOOL CBand::SetCustomToolSizes(SIZE size)
{
	if (ddCBSNone != bpV1.m_cbsChildStyle)
	{
		CBand* pChildBand = m_pChildBands->GetCurrentChildBand();
		if (NULL == pChildBand)
			return FALSE;
		return pChildBand->SetCustomToolSizes(size);
	}

	if (NULL == m_prcTools)
		return TRUE;

	HRESULT hResult;
	VARIANT vWidth;
	vWidth.vt = VT_I4 | VT_ARRAY;
	VARIANT vHeight;
	vHeight.vt = VT_I4 | VT_ARRAY;
	long* pWidth = NULL;
	long* pHeight = NULL;
	if (VARIANT_FALSE == m_pRootBand->bpV1.m_vbAutoSizeForms)
	{
		GetCustomToolSizes(vWidth, vHeight);

		SIZE sizeTwips;
		PixelToTwips(&size, &sizeTwips);
		m_pBar->FireBandResize((Band*)m_pRootBand, &vWidth, &vHeight, sizeTwips.cx-1, sizeTwips.cy-1);

		hResult = SafeArrayAccessData(vWidth.parray, (void**)&pWidth);
		if (FAILED(hResult))
			return FALSE;

		hResult = SafeArrayAccessData(vHeight.parray, (void**)&pHeight);
		if (FAILED(hResult))
		{
			SafeArrayUnaccessData(vWidth.parray);
			return FALSE;
		}
	}

	int nTool = 0;
	int nCustomTool = 0;
	BOOL bResult = TRUE;
	CTool* pTool;
	int nLast = 0;
	int nCount = m_pTools->GetVisibleToolCount();
	for (nTool = 0; nTool < nCount; nTool++)
	{
		pTool = m_pTools->GetVisibleTool(nTool);
		
		if (ddTTForm == pTool->tpV1.m_ttTools)
			continue;

		switch (m_pRootBand->bpV1.m_daDockingArea)
		{
		case ddDATop:
		case ddDABottom:
		case ddDAFloat:
			size.cx -= m_prcTools[nTool].Width() + 3;
			break;

		case ddDALeft:
		case ddDARight:
			if (nLast < m_prcTools[nTool].bottom)
			{
				nLast = m_prcTools[nTool].bottom;
				size.cy -= m_prcTools[nTool].Height() + eBevelBorder2;
			}
			break;
		}
	}

	int nOffset = 0;
	int nTemp = 0;
	int nTemp2 = 0;
	for (nTool = 0; nTool < nCount; nTool++)
	{
		pTool = m_pTools->GetVisibleTool(nTool);

		if (VARIANT_TRUE == m_pRootBand->bpV1.m_vbAutoSizeForms)
		{
			switch (m_pRootBand->bpV1.m_daDockingArea)
			{
			case ddDATop:
			case ddDAFloat:
			case ddDABottom:
				if (nOffset)
					m_prcTools[nTool].Offset(nOffset, 0);
				break;

			case ddDALeft:
			case ddDARight:
				if (nOffset)
					m_prcTools[nTool].Offset(0, nOffset);
				break;
			}
			if (ddTTForm == pTool->tpV1.m_ttTools)
			{
				switch (m_pRootBand->bpV1.m_daDockingArea)
				{
				case ddDATop:
				case ddDAFloat:
				case ddDABottom:
					nTemp = m_prcTools[nTool].right;
					nTemp2 = m_prcTools[nTool].right = m_prcTools[nTool].left + (size.cx / m_nFormControlCount) - eBevelBorder2 - 1;
					m_prcTools[nTool].bottom = m_prcTools[nTool].top + size.cy - eBevelBorder2;
					nOffset += nTemp2 - nTemp; 
					break;

				case ddDALeft:
				case ddDARight:
					m_prcTools[nTool].right = m_prcTools[nTool].left + size.cx - eBevelBorder2;
					nTemp = m_prcTools[nTool].bottom;
					nTemp2 = m_prcTools[nTool].bottom = m_prcTools[nTool].top + (size.cy / m_nFormControlCount) - eBevelBorder2 - 1; 
					nOffset += nTemp2 - nTemp;
					break;
				}
			}
			else if (ddTTControl == pTool->tpV1.m_ttTools)
			{
				switch (m_pRootBand->bpV1.m_daDockingArea)
				{
				case ddDATop:
				case ddDAFloat:
				case ddDABottom:
					if (m_prcTools[nTool].Height() > size.cy)
						m_prcTools[nTool].bottom = m_prcTools[nTool].top + size.cy - eBevelBorder2;
					break;

				case ddDALeft:
				case ddDARight:
					if (m_prcTools[nTool].Width() > size.cx)
						m_prcTools[nTool].right = m_prcTools[nTool].left + size.cx - eBevelBorder2;
					break;
				}
			}
		}
		else
		{
			if (ddTTForm != pTool->tpV1.m_ttTools && ddTTControl != pTool->tpV1.m_ttTools)
				continue;

			size.cx = pWidth[nCustomTool];
			size.cy = pHeight[nCustomTool];
			TwipsToPixel(&size, &size);
			m_prcTools[nTool].right = m_prcTools[nTool].left + size.cx;
			m_prcTools[nTool].bottom = m_prcTools[nTool].top + size.cy;
		}
		nCustomTool++;
	}

	if (VARIANT_FALSE == m_pRootBand->bpV1.m_vbAutoSizeForms)
	{
		hResult = SafeArrayUnaccessData(vWidth.parray);
		assert(SUCCEEDED(hResult));
		VariantClear(&vWidth);
		hResult = SafeArrayUnaccessData(vHeight.parray);
		assert(SUCCEEDED(hResult));
		VariantClear(&vHeight);
	}
	return bResult;
}

//
// TrackSlidingTabScrollButton
//

void CBand::TrackSlidingTabScrollButton(const POINT& pt, int nHit)
{
	HWND hWndCapture;
	CRect rcBand;
	GetBandRect(hWndCapture, rcBand);
	HDC hDC = GetDC(hWndCapture);
	if (NULL == hDC)
		return;

	CRect rcButton;
	BTabs::SCROLLBUTTONSTYLES sbsPressed;
	BTabs::SCROLLBUTTONSTYLES sbsCurrent;
	switch (nHit)
	{
	case BTabs::eSlidingBottom:
		sbsCurrent = m_pChildBands->GetBottomButtonStyle();
		if (BTabs::eInactiveRightButton == sbsCurrent || BTabs::eInactiveBottomButton == sbsCurrent)
			return;
		
		m_pChildBands->SetBottomPaintLock(TRUE);
		switch (sbsCurrent)
		{
		case BTabs::eRightButton:
			sbsPressed = BTabs::ePressedRightButton;
			break;

		case BTabs::eBottomButton:
			sbsPressed = BTabs::ePressedBottomButton;
			break;
		}
		rcButton = m_pChildBands->GetBottomButtonLocation();
		m_pChildBands->m_pCurChildBand->IncrementFirstTool();
		break;

	case BTabs::eSlidingTop:
		sbsCurrent = m_pChildBands->GetTopButtonStyle();
		if (BTabs::eInactiveTopButton == sbsCurrent || BTabs::eInactiveLeftButton == sbsCurrent)
			return;

		m_pChildBands->SetTopPaintLock(TRUE);
		switch (sbsCurrent)
		{
		case BTabs::eLeftButton:
			sbsPressed = BTabs::ePressedLeftButton;
			break;

		case BTabs::eTopButton:
			sbsPressed = BTabs::ePressedTopButton;
			break;
		}
		rcButton = m_pChildBands->GetTopButtonLocation();
		m_pChildBands->m_pCurChildBand->DecrementFirstTool();
		break;
	}

	BOOL bKillTimer = FALSE;
	m_pChildBands->InteractivePaintScrollButton(hDC, sbsPressed);
	SetCapture(hWndCapture);
	BTabs::SCROLLBUTTONSTYLES theStyle;
	const UINT nTimerId = 1000;
	UINT nTimer = SetTimer(hWndCapture, nTimerId, 250, 0); 
	MSG msg;
  	while (GetCapture() == hWndCapture)
	{
		GetMessage(&msg, NULL, 0, 0);

		switch (msg.message)
		{
		case WM_TIMER:
			switch (sbsPressed)
			{
			case BTabs::ePressedTopButton:
			case BTabs::ePressedLeftButton:
				theStyle = m_pChildBands->GetTopButtonStyle();
				if (BTabs::eInactiveLeftButton != theStyle &&
					BTabs::eInactiveTopButton != theStyle)
				{
					m_pChildBands->m_pCurChildBand->DecrementFirstTool();
					Refresh();
				}
				else
				{
					bKillTimer = TRUE;
					KillTimer(hWndCapture, nTimer);
				}
				break;

			case BTabs::ePressedRightButton:
			case BTabs::ePressedBottomButton:
				theStyle = m_pChildBands->GetBottomButtonStyle();
				if (BTabs::eInactiveRightButton != theStyle &&
					BTabs::eInactiveBottomButton != theStyle)
				{
					m_pChildBands->m_pCurChildBand->IncrementFirstTool();
					Refresh();
				}
				else
				{
					bKillTimer = TRUE;
					KillTimer(hWndCapture, nTimer);
				}
				break;
			}
			break;

		case WM_LBUTTONUP:
			goto ExitLoop;

		default:
			DispatchMessage(&msg);
			break;
		}
	}
ExitLoop:
	if (!bKillTimer)
		KillTimer(hWndCapture, nTimer);
	ReleaseDC(hWndCapture, hDC);

	if (GetCapture() == hWndCapture)
		ReleaseCapture();

	switch (sbsPressed)
	{
	case BTabs::ePressedLeftButton:
	case BTabs::ePressedTopButton:
		m_pChildBands->SetTopPaintLock(FALSE);
		break;

	case BTabs::ePressedRightButton:
	case BTabs::ePressedBottomButton:
		m_pChildBands->SetBottomPaintLock(FALSE);
		break;
	}
	m_pBar->RecalcLayout();
}

//
// DrawCloseButton
//

void CBand::DrawCloseButton(HDC hDC, BOOL bPressed, BOOL bHover)
{
	if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
	{
		if (bPressed || bHover)
		{
			HBRUSH brushBackground = NULL;
			HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
			if (penBorder)
			{
				HPEN penOld = SelectPen(hDC, penBorder);

				if (bPressed) // IsPushed
					brushBackground = CreateSolidBrush(m_pBar->m_crXPPressed);
				else
					brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);

				if (brushBackground)
				{
					HBRUSH brushOld = SelectBrush(hDC, brushBackground);

					CRect rcTemp = m_rcCloseButton;
					rcTemp.Inflate(1, 1);
					Rectangle(hDC, rcTemp.left, rcTemp.top, rcTemp.right, rcTemp.bottom);

					SelectBrush(hDC, brushOld);
					BOOL bResult = DeleteBrush(brushBackground);
					assert(bResult);
				}
				SelectPen(hDC, penBorder);
				BOOL bResult = DeletePen(penBorder);
				assert(bResult);
			}
		}
		else
			FillSolidRect(hDC, m_rcCloseButton, m_pBar->m_crXPBackground);
		DrawCross(hDC, m_rcCloseButton.left+1, m_rcCloseButton.top+1, m_rcCloseButton.right-1, m_rcCloseButton.bottom-1, 2, GetSysColor(COLOR_BTNTEXT));
	}
	else
	{
		FillSolidRect(hDC, m_rcCloseButton, m_pBar->m_crBackground);
		DrawCross(hDC, m_rcCloseButton.left + 2, m_rcCloseButton.top + 2, m_rcCloseButton.right - 3, m_rcCloseButton.bottom - 3, 2, GetSysColor(COLOR_BTNTEXT));
		if (bPressed)
			m_pBar->DrawEdge(hDC, m_rcCloseButton, BDR_SUNKENOUTER, BF_RECT);
		else
			m_pBar->DrawEdge(hDC, m_rcCloseButton, BDR_RAISEDINNER, BF_RECT);
	}
}

//
// TrackCloseButton
//

BOOL CBand::TrackCloseButton(const POINT& pt)
{
	CRect rcTemp = m_rcCloseButton;
	rcTemp.Offset(-m_rcDock.left, -m_rcDock.top);
	if (!PtInRect(&rcTemp, pt))
		return FALSE;

	HWND hWndCapture;
	CRect rcBand;
	GetBandRect(hWndCapture, rcBand);
	
	HDC hDC = GetDC(hWndCapture);

	DrawCloseButton(hDC, TRUE, FALSE);

	rcTemp = m_rcCloseButton;
	ClientToScreen(hWndCapture, rcTemp);
	SetCapture(hWndCapture);
	MSG msg;
  	while (GetCapture() == hWndCapture)
	{
		GetMessage(&msg, NULL, 0, 0);

		switch (msg.message)
		{
		case WM_CANCELMODE:
			goto ExitLoop;

		case WM_LBUTTONUP:
			if (PtInRect(&rcTemp, msg.pt))
			{
				try
				{
					put_Visible(VARIANT_FALSE);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
			else
				DrawCloseButton(hDC);
			goto ExitLoop;

		case WM_MOUSEMOVE:
			if (PtInRect(&rcTemp, msg.pt))
				DrawCloseButton(hDC, TRUE);
			else
				DrawCloseButton(hDC, FALSE, TRUE);
			break;

		default:
			DispatchMessage(&msg);
			break;
		}
	}
ExitLoop:
	if (VARIANT_FALSE == bpV1.m_vbVisible)
	{
		m_pBar->InternalRecalcLayout();
		m_pDock->InvalidateRect(NULL, FALSE);
	}

	ReleaseCapture();
	ReleaseDC(hWndCapture, hDC);
	return TRUE;
}

//
// DrawExpandButton
//

void CBand::DrawExpandButton(HDC hDC, BOOL bPressed, BOOL bHover)
{
	BOOL bVertical = IsVertical();
	if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
	{
		if (bPressed || bHover)
		{
			HBRUSH brushBackground = NULL;
			HPEN penBorder = CreatePen(PS_SOLID, 1, m_pBar->m_crXPMenuSelectedBorderColor);
			if (penBorder)
			{
				HPEN penOld = SelectPen(hDC, penBorder);

				if (bPressed) // IsPushed
					brushBackground = CreateSolidBrush(m_pBar->m_crXPPressed);
				else
					brushBackground = CreateSolidBrush(m_pBar->m_crXPSelectedColor);

				if (brushBackground)
				{
					HBRUSH brushOld = SelectBrush(hDC, brushBackground);
					
					CRect rcTemp = m_rcExpandButton;
					rcTemp.Inflate(1,1);
					Rectangle(hDC, rcTemp.left, rcTemp.top, rcTemp.right, rcTemp.bottom);

					SelectBrush(hDC, brushOld);
					BOOL bResult = DeleteBrush(brushBackground);
					assert(bResult);
				}
				SelectPen(hDC, penBorder);
				BOOL bResult = DeletePen(penBorder);
				assert(bResult);
			}
		}
		else
			FillSolidRect(hDC, m_rcExpandButton, m_pBar->m_crXPBackground);
	}
	else
	{
		FillSolidRect(hDC, m_rcExpandButton, m_pBar->m_crBackground);
	
		if (bPressed)
			m_pBar->DrawEdge(hDC, m_rcExpandButton, BDR_SUNKENOUTER, BF_RECT);
		else
			m_pBar->DrawEdge(hDC, m_rcExpandButton, BDR_RAISEDINNER, BF_RECT);
	}
	CRect rcBitmap = m_rcExpandButton;
	if (bVertical)
	{
		rcBitmap.left = rcBitmap.left + (rcBitmap.Width() - 5) / 2;
		rcBitmap.right = rcBitmap.left + 5;
		rcBitmap.top = rcBitmap.top + (rcBitmap.Height() - 3) / 2;
		rcBitmap.bottom = rcBitmap.top + 3;
	}
	else
	{
		rcBitmap.left = rcBitmap.left + (rcBitmap.Width() - 3) / 2;
		rcBitmap.right = rcBitmap.left + 3;
		rcBitmap.top = rcBitmap.top + (rcBitmap.Height() - 5) / 2;
		rcBitmap.bottom = rcBitmap.top + 5;
	}

	switch (m_eExpandButtonState)
	{
	case eGrayed:
		DrawExpandBitmap(hDC, rcBitmap, bVertical ? GetGlobals().GetBitmapExpandVert() : GetGlobals().GetBitmapExpandHorz(), bVertical ? 10 : 6, bVertical ? 5 : 3, bVertical ? 3 : 5);
		break;

	case eContracted:
		DrawExpandBitmap(hDC, rcBitmap, bVertical ? GetGlobals().GetBitmapExpandVert() : GetGlobals().GetBitmapExpandHorz(), bVertical ? 5 : 3, bVertical ? 5 : 3, bVertical ? 3 : 5);
		break;

	case eExpanded:
		DrawExpandBitmap(hDC, rcBitmap, bVertical ? GetGlobals().GetBitmapExpandVert() : GetGlobals().GetBitmapExpandHorz(), 0, bVertical ? 5 : 3, bVertical ? 3 : 5);
		break;
	}
}

//
// TrackExpandButton
//

BOOL CBand::TrackExpandButton(const POINT& pt)
{
	BOOL bStateChanged = FALSE;
	CRect rcTemp = m_rcExpandButton;
	rcTemp.Offset(-m_rcDock.left, -m_rcDock.top);
	if (!PtInRect(&rcTemp, pt))
		return FALSE;

	HWND hWndCapture;
	CRect rcBand;
	GetBandRect(hWndCapture, rcBand);
	
	HDC hDC = GetDC(hWndCapture);

	DrawExpandButton(hDC, TRUE);

	rcTemp = m_rcExpandButton;
	ClientToScreen(hWndCapture, rcTemp);
	SetCapture(hWndCapture);
	MSG msg;
  	while (GetCapture() == hWndCapture)
	{
		GetMessage(&msg, NULL, 0, 0);

		switch (msg.message)
		{
		case WM_CANCELMODE:
			goto ExitLoop;

		case WM_LBUTTONUP:
			if (PtInRect(&rcTemp, msg.pt))
			{
				bStateChanged = TRUE;
				switch (m_eExpandButtonState) 
				{
				case eContracted:
					if (m_pDock)
						m_pDock->MakeLineContracted(m_nDockLineIndex);
					m_eExpandButtonState = eExpanded;
					break;

				case eExpanded:
					m_eExpandButtonState = eExpandedToContracted;
					break;
				}
			}
			DrawExpandButton(hDC);
			goto ExitLoop;

		case WM_MOUSEMOVE:
			if (PtInRect(&rcTemp, msg.pt))
				DrawExpandButton(hDC, TRUE);
			else
				DrawExpandButton(hDC, FALSE, TRUE);
			break;

		default:
			DispatchMessage(&msg);
			break;
		}
	}

ExitLoop:
	if (bStateChanged)
	{
		m_pBar->InternalRecalcLayout();
		m_pDock->InvalidateRect(NULL, FALSE);
	}
	ReleaseCapture();
	ReleaseDC(hWndCapture, hDC);
	return TRUE;
}

//
// InsideCalcHorzLayout
//

BOOL CBand::InsideCalcHorzLayout(HDC    hDC,
								 const CRect& rcBound,
							     BOOL   bCommit,
							     int    nTotalWidth,
							     BOOL   bWrapFlag,
							     BOOL   bLeftHandle,
							     BOOL   bTopHandle,
							     CRect& rcReturn,
							     int&   nReturnLineCount,
							     DWORD  dwLayoutFlag)
{
	TypedArray<CTool*> aTools;
	
	try
	{
		BOOL bVertical = (dwLayoutFlag & CBand::eLayoutVert && !(dwLayoutFlag & CBand::eLayoutFloat)) || (dwLayoutFlag & CBand::eLayoutVert && m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle); 
		BOOL bVerticalMenu = bVertical && (ddBTMenuBar == m_pRootBand->bpV1.m_btBands || ddBTChildMenuBar == m_pRootBand->bpV1.m_btBands);
		if (bVerticalMenu)
			bWrapFlag = TRUE;
		if (m_pParentBand && ddCBSNone != m_pParentBand->bpV1.m_cbsChildStyle)
			bWrapFlag = (VARIANT_TRUE == bpV1.m_vbWrapTools);

		nReturnLineCount = 1;
		m_aLineHeight.RemoveAll();
		m_aLineWidth.RemoveAll();
		
		int nToolCount = m_pTools->GetVisibleTools(aTools);
		if (bCommit)
		{
			if (m_prcTools)
			{
				delete [] m_prcTools;
				m_prcTools = NULL;
			}

			if (nToolCount > 0)
			{
				m_prcTools = new CRect[nToolCount];
				if (NULL == m_prcTools)
					return FALSE;
			}
		}

		m_nHostControlCount = 0;
		m_nFormControlCount = 0;
		if (m_pnCustomControlIndexes)
		{
			delete [] m_pnCustomControlIndexes;
			m_pnCustomControlIndexes = NULL;
		}

		if (nToolCount < 1)
		{
			if (ddBTStatusBar == m_pRootBand->bpV1.m_btBands && m_bStatusBarSizerPresent)
			{
				rcReturn.right += eEmptyStatusBarWidth;
				rcReturn.bottom += eEmptyStatusBarHeight;
			}
			else if (ddBFSizer & m_pRootBand->bpV1.m_dwFlags)
			{
				if (bVertical)
				{
					if (0 == m_pRootBand->bpV1.m_rcVertDockForm.Width() && 0 == m_pRootBand->bpV1.m_rcVertDockForm.Height())
						m_pRootBand->bpV1.m_rcVertDockForm.Set(0, 0, eEmptyDockWidth, eEmptyDockHeight);

					rcReturn.right += m_pRootBand->bpV1.m_rcVertDockForm.Width();
					rcReturn.bottom += m_pRootBand->bpV1.m_rcVertDockForm.Height();
				}
				else
				{
					if (0 == m_pRootBand->bpV1.m_rcHorzDockForm.Width() && 0 == m_pRootBand->bpV1.m_rcHorzDockForm.Height())
						m_pRootBand->bpV1.m_rcHorzDockForm.Set(0, 0, eEmptyDockWidth, eEmptyDockHeight);

					rcReturn.right += m_pRootBand->bpV1.m_rcHorzDockForm.Width();
					rcReturn.bottom += m_pRootBand->bpV1.m_rcHorzDockForm.Height();
				}
			}
			else if (m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle)
			{
				if (ddDAFloat == m_pRootBand->bpV1.m_daDockingArea && m_pRootBand->GetFloat() && IsWindow(m_pRootBand->GetFloat()->hWnd()))
				{
					rcReturn = rcBound;
					m_pRootBand->AdjustChildBandsRect(rcReturn);
				}
				else
				{
					if (-1 == m_pParentBand->bpV1.m_rcDimension.right || -1 == m_pParentBand->bpV1.m_rcDimension.bottom)
					{
						rcReturn.right += 30;
						rcReturn.bottom += 30;
						m_pParentBand->bpV1.m_rcDimension.Set(0, 0, rcReturn.right, rcReturn.bottom);
					}
					if (bVertical)
					{
						rcReturn.right = m_pParentBand->bpV1.m_rcDimension.Width();
						rcReturn.bottom = rcBound.Height();
					}
					else
					{
						rcReturn.right = rcBound.Width();
						rcReturn.bottom = m_pParentBand->bpV1.m_rcDimension.Height();
					}
				}
			}
			else
			{
				//
				// All other types of bands 
				//

				rcReturn.right += eEmptyBarWidth;
				rcReturn.bottom += eEmptyBarHeight;
			}
			if (bCommit)
				m_pTools->CommitVisibleTools(aTools);
			return TRUE; 
		}

		CTool* pTool;
		CTool* pLastVisibleTool;
		CRect  rcTool;
		SIZE   sizeTool;

		int nY = 0;
		int nTool = 0; 
		int nWidth;
		int nXOffset = 0;
		int nMaxWidth = 0;
		int nBackTrackPos;
		int nPrevMaxWidth = 0;
		int nMaxHeight = 0;
		int nLineHeight = 0;

		BandTypes btCorrectedBandType = bpV1.m_btBands;
		if (ddBTPopup == btCorrectedBandType && !m_bPopupWinLock)
			btCorrectedBandType = ddBTNormal;

		//
		// Adjusting for the tools
		//
		// Bounding box needs to at least as large as the largest tool width + borders
		// eliminate infinite wrapping of tool due to small TotalWidth
		//

		m_sizeLargestTool.cx = 0;
		m_sizeLargestTool.cy = 0;
		for (nTool = 0; nTool < nToolCount; nTool++)
		{
			try
			{
				pTool = aTools.GetAt(nTool);

				switch (pTool->tpV1.m_ttTools)
				{
				case ddTTForm:
					m_nFormControlCount++;
					break;

				case ddTTControl:
					m_nHostControlCount++;
					break;

				case ddTTSeparator:
					pTool->m_bVerticalSeparator = FALSE;
					break;
				}

				//
				// Hook Tool to Band 
				//

				pTool->m_pBand = this;  
				pTool->m_pBar = m_pBar;  

				sizeTool = pTool->CalcSize(hDC, btCorrectedBandType, bVertical);

				if (!(bVertical && (ddTTCombobox == pTool->tpV1.m_ttTools || ddTTEdit == pTool->tpV1.m_ttTools)))
				{
					if (-1 != pTool->tpV1.m_nWidth && VARIANT_TRUE != m_pBar->bpV1.m_vbLargeIcons && sizeTool.cx < pTool->tpV1.m_nWidth)
						sizeTool.cx = pTool->tpV1.m_nWidth + 2 * CBand::eHToolPadding;

					if (-1 != pTool->tpV1.m_nHeight && VARIANT_TRUE != m_pBar->bpV1.m_vbLargeIcons && sizeTool.cy < pTool->tpV1.m_nHeight)
						sizeTool.cy = pTool->tpV1.m_nHeight + 2 * CBand::eVToolPadding;

					if (m_sizeLargestTool.cx < sizeTool.cx)
						m_sizeLargestTool.cx = sizeTool.cx;
					
					if (m_sizeLargestTool.cy < sizeTool.cy)
						m_sizeLargestTool.cy = sizeTool.cy;
				}

				if (bVertical && CTool::ddTTMoreTools == pTool->tpV1.m_ttTools)
				{
					if (m_pParentBand && VARIANT_TRUE == bpV1.m_vbWrapTools)
						sizeTool.cx = m_sizeLargestTool.cx;
					else if (NULL == m_pParentBand && VARIANT_TRUE != m_pRootBand->bpV1.m_vbWrapTools)
						sizeTool.cx = m_sizeLargestTool.cx;
				}

				pTool->m_sizeTool = sizeTool;

				nWidth = nXOffset + sizeTool.cx + CBand::eBevelBorder2;
				if (nWidth > nTotalWidth)	
					nTotalWidth = nWidth;

				pLastVisibleTool = pTool;
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}
		}
		
		//
		// Do the layout
		//
		
		CTool* pBackTrackTool;
		CTool* pPrevTool;
		BOOL bBackTrackSeparator = FALSE;
		BOOL bFirstVisible = TRUE;
		int nLastSeparatorPos = 0;
		int nRequiredWidth;
	
		if (m_nFormControlCount > 0)
		{
			m_pnCustomControlIndexes = new int[m_nFormControlCount];
			if (NULL == m_pnCustomControlIndexes)
				return FALSE;
		}

		HRESULT hResult;
		BOOL bAutoSizeTool = FALSE;	
		for (nTool = 0; nTool < nToolCount; nTool++)
		{
SeparatorReturnPoint:
			try
			{
				pTool = aTools.GetAt(nTool);

				pTool->m_bAdjusted = FALSE;

				//
				// Calculate Tool Size
				//

				sizeTool = pTool->m_sizeTool;

				if (ddBFStretch & bpV1.m_dwFlags && NULL == m_pFloat && VARIANT_FALSE == bpV1.m_vbWrapTools)
				{
					// CR 1981
					switch (pTool->tpV1.m_asAutoSize)
					{
					case ddTASSpringWidth:
						bAutoSizeTool = TRUE;
						if (bVertical)
							sizeTool.cx = m_sizeLargestTool.cx;
						break;

					case ddTASSpringHeight:
						if (!bVertical)
							sizeTool.cy = m_sizeLargestTool.cy;
						break;

					case ddTASSpringBoth:
						bAutoSizeTool = TRUE;
						if (bVertical)
							sizeTool.cx = m_sizeLargestTool.cx;
						else
							sizeTool.cy = m_sizeLargestTool.cy;
						break;
					}
				}

				//
				// Calculate the required width
				//
				
				nRequiredWidth = nXOffset + sizeTool.cx + CBand::eBevelBorder2;
				if (pTool != pLastVisibleTool)
					nRequiredWidth += bpV1.m_nToolsHSpacing;

				if (nTotalWidth < nRequiredWidth)
				{
					// Check if wrappable or this is a floating band
					if (bWrapFlag)
					{
						if (!bFirstVisible)
						{
							// 
							// Now we need to back track to find a separator 
							// tool to break the line at.
							//

							nBackTrackPos = nTool;
							if (nBackTrackPos == nToolCount)
								nBackTrackPos--;

							while (nBackTrackPos >= 0)
							{
								pBackTrackTool = aTools.GetAt(nBackTrackPos);

								if (ddTTSeparator == pBackTrackTool->tpV1.m_ttTools)
								{
									if (0 != nBackTrackPos)
									{
										pPrevTool = aTools.GetAt(nBackTrackPos - 1);

										if (pPrevTool->m_rcTemp.top == pBackTrackTool->m_rcTemp.top)
										{
											//
											// Is't not the first tool on the line, make the Separator Tool a Vertical Separator
											//
											
											pBackTrackTool->m_bVerticalSeparator = TRUE;

											//
											// Break the line at this point 
											//
											
											nY += nLineHeight + bpV1.m_nToolsVSpacing;
											hResult = m_aLineHeight.Add(nLineHeight);
											assert(SUCCEEDED(hResult));
											hResult = m_aLineWidth.Add(nXOffset);
											assert(SUCCEEDED(hResult));
											nReturnLineCount++;
											nLineHeight = 0;
											nXOffset = 0;
											nMaxWidth = pBackTrackTool->m_nTempMaxWidth;

											//
											// Calculate the layout from nBackTrackPos
											//
	
											nTool = nBackTrackPos;
											goto SeparatorReturnPoint;
										}
									}
									break;
								}
								nBackTrackPos--;
							} // End While
						}

						//
						// Wrap to the next line
						//

						if (ddTTSeparator == pTool->tpV1.m_ttTools)
							pTool->m_bVerticalSeparator = TRUE;

						if (nXOffset > nMaxWidth)
							nMaxWidth = nXOffset;

						nY += nLineHeight + bpV1.m_nToolsVSpacing;
						
						hResult = m_aLineHeight.Add(nLineHeight);
						assert(SUCCEEDED(hResult));
						hResult = m_aLineWidth.Add(nXOffset);
						assert(SUCCEEDED(hResult));
						nReturnLineCount++;
						nLineHeight = 0;
						nXOffset = 0;
					} 
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}

			bFirstVisible = FALSE;

			if (pTool->m_bVerticalSeparator)
			{
				sizeTool.cy = sizeTool.cx;
				sizeTool.cx = 0;
			}

			rcTool.top    = nY + CBand::eBevelBorder;  
			rcTool.bottom = rcTool.top + sizeTool.cy;
			rcTool.left   = nXOffset + CBand::eBevelBorder;
			rcTool.right  = rcTool.left + sizeTool.cx;
			pTool->m_nBandLine = nReturnLineCount;
			switch (pTool->tpV1.m_ttTools)
			{
			case CTool::ddTTMDIButtons:
				{
					//
					// Adjusting the MDI Buttons on a Menu Band docked in the Top Dock Area
					//

					if (ddDATop == m_pRootBand->bpV1.m_daDockingArea)
					{
						int nWidth = rcTool.Width();
						rcTool.left = nTotalWidth - nWidth;
						rcTool.right = rcTool.left + nWidth;
					}
				}
				break;
			}
			
			pTool->m_rcTemp = rcTool;

			//
			// m_nTempMaxWidth is used during back tracking
			//
			
			pTool->m_nTempMaxWidth = nMaxWidth;
			
			//
			// Commit the Tool's Layout
			//
			
			if (bCommit && m_prcTools)
				m_prcTools[nTool] = rcTool;
			
			//
			// Calculate the new line height
			//

			if (rcTool.Height() + CBand::eBevelBorder > nLineHeight)
				nLineHeight = rcTool.Height() + CBand::eBevelBorder2;

			if (bVertical && m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle && !bWrapFlag)
			{
				nY += rcTool.Height() + CBand::eBevelBorder + bpV1.m_nToolsVSpacing;
				if (pTool != pLastVisibleTool)
					nXOffset += bpV1.m_nToolsHSpacing;
				if (nXOffset > nMaxWidth)
					nMaxWidth = nXOffset; 
				nXOffset = 0;
			}
			else
			{
				nXOffset += rcTool.Width() + CBand::eBevelBorder2;

				//
				// Don't use the last m_nToolsHSpacing (see line below)
				//
			
				if (nXOffset > nMaxWidth)
					nMaxWidth = nXOffset; 

				if (pTool != pLastVisibleTool)
					nXOffset += bpV1.m_nToolsHSpacing;

				//
				// Only applicable if free form sizeable toolbar
				//
				
				if ((nY + sizeTool.cy) > nMaxHeight) 
					nMaxHeight = nY + sizeTool.cy;

				if (pTool->m_bVerticalSeparator)
				{
					nY += nLineHeight + bpV1.m_nToolsVSpacing;
					nLineHeight = 0;
					nXOffset = 0;
				}
			}
		}
		hResult = m_aLineHeight.Add(nLineHeight);
		assert(SUCCEEDED(hResult));
		hResult = m_aLineWidth.Add(nXOffset);
		assert(SUCCEEDED(hResult));

		//
		// Outer Border and Bevel Border for the end of the band
		//
		
		nMaxHeight += eBevelBorder2; 

		//
		// Finally return the result rectangle
		//

		m_sizePage.cx = nMaxWidth;
		m_sizePage.cy = nMaxHeight;
		if (ddBFSizer & m_pRootBand->bpV1.m_dwFlags)
		{
			if (ddDAFloat == m_pRootBand->bpV1.m_daDockingArea && m_pRootBand->GetFloat() && m_pRootBand->GetFloat()->IsWindow())
			{
				rcReturn = rcBound;
				m_pRootBand->AdjustChildBandsRect(rcReturn);
			}
			else
			{
				if (bVertical)
				{
					if ((0 == m_pRootBand->bpV1.m_rcVertDockForm.Width() && 0 == m_pRootBand->bpV1.m_rcVertDockForm.Height()))
						m_pRootBand->bpV1.m_rcVertDockForm.Set(0, 0, nMaxWidth, nMaxHeight);

					rcReturn.right += m_pRootBand->bpV1.m_rcVertDockForm.Width();
					rcReturn.bottom += m_pRootBand->bpV1.m_rcVertDockForm.Height();
				}
				else
				{
					if (0 == m_pRootBand->bpV1.m_rcHorzDockForm.Width() && 0 == m_pRootBand->bpV1.m_rcHorzDockForm.Height())
						m_pRootBand->bpV1.m_rcHorzDockForm.Set(0, 0, nMaxWidth, nMaxHeight);

					rcReturn.right += m_pRootBand->bpV1.m_rcHorzDockForm.Width();
					rcReturn.bottom += m_pRootBand->bpV1.m_rcHorzDockForm.Height();
				}
			}
		}
		else if (m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle)
		{
			if (ddDAFloat == m_pRootBand->bpV1.m_daDockingArea && m_pRootBand->GetFloat() && IsWindow(m_pRootBand->GetFloat()->hWnd()))
			{
				rcReturn = rcBound;
				m_pRootBand->AdjustChildBandsRect(rcReturn);
			}
			else
			{
				if (-1 == m_pParentBand->bpV1.m_rcDimension.right || -1 == m_pParentBand->bpV1.m_rcDimension.bottom)
				{
					rcReturn.right += nMaxWidth;
					rcReturn.bottom += nMaxHeight;
					m_pParentBand->bpV1.m_rcDimension.Set(0, 0, rcReturn.right, rcReturn.bottom);
				}
				if (bVertical)
				{
					rcReturn.right = m_pParentBand->bpV1.m_rcDimension.Width();
					rcReturn.bottom = rcBound.Height();
				}
				else
				{
					rcReturn.right = rcBound.Width();
					rcReturn.bottom = m_pParentBand->bpV1.m_rcDimension.Height();
				}
			}
		}
		else
		{
			if (bVerticalMenu)
			{
				if (bCommit)
				{
					// m_prcTools rotate 90deg and flip horizontally
					for (nTool = nToolCount - 1; nTool >= 0; nTool--)
					{
						rcTool = m_prcTools[nTool];
						m_prcTools[nTool].left = nMaxHeight - rcTool.bottom;
						m_prcTools[nTool].right = nMaxHeight - rcTool.top;
						m_prcTools[nTool].top = rcTool.left;
						m_prcTools[nTool].bottom = rcTool.right;
					}
				}
				int nTemp = nMaxWidth;
				nMaxWidth = nMaxHeight;
				nMaxHeight = nTemp;
			}
			rcReturn.right += nMaxWidth;
			rcReturn.bottom += nMaxHeight;
		}
		if (bCommit)
			m_pTools->CommitVisibleTools(aTools);
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
// InsideDraw
//

BOOL CBand::InsideDraw(HDC hDC, CRect rcPaint, BOOL bWrapFlag, BOOL bVertical, BOOL bVerticalPaint, BOOL bLeftHandle, BOOL bRightHandle, const POINT& ptPaintOffset, BOOL bDrawControlsOrForms)
{
	try
	{
		if (m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle)
			bWrapFlag = (VARIANT_TRUE == bpV1.m_vbWrapTools);

		m_rcCurrentPage = rcPaint;
		m_rcCurrentPage.Offset(ptPaintOffset.x, ptPaintOffset.y);

		//
		// Painting Tools  
		//

		// There are no tools
		if (NULL == m_prcTools) 
			return TRUE;

		//
		// Drawing Tools  
		//

		CTool* pTool;
		CRect  rcTool;
		CRect  rcLastTool;
		int    nHeight;
		BOOL   bExceededLimit = FALSE;
		BOOL   bLastTool = FALSE;
		BOOL   bContinue = TRUE;
		int    nToolCount = m_pTools->GetVisibleToolCount();
		int    nLastVisibleTool = nToolCount - 1;
		int    nLimit = bVertical ? rcPaint.bottom : rcPaint.right;

		if (nToolCount > 0)
		{
			rcLastTool = m_prcTools[nLastVisibleTool];
			rcLastTool.Offset(rcPaint.left, rcPaint.top);
		}

		if (m_pTools->m_pMoreTools)
		{
			m_pTools->m_pMoreTools->SetVisibleOnPaint(TRUE);
			m_pTools->m_pMoreTools->m_bMoreTools = FALSE;
		}

		CRect rcTemp;
		for (int nTool = GetFirstTool(); nTool < nToolCount && bContinue; nTool++)
		{
			pTool = m_pTools->GetVisibleTool(nTool);

			if (m_pBar->m_bPopupMenuExpanded && m_prcPopupExpandArea && !m_prcPopupExpandArea[nTool].IsEmpty())
			{
				//
				// Drawing the background rectangle around the MRU Menus
				//

				rcTemp = m_prcPopupExpandArea[nTool];
				rcTemp.Offset(3, 3);
				rcTemp.Inflate(2, 0);
				if (nTool == GetFirstTool())
				{
					//
					// This is the first expanded area
					//

					rcTemp.top -= 2;
				}
				else if (nTool == m_nLastPopupExpandedArea)
				{
					//
					// This is the last expanded area
					//

					rcTemp.bottom += 2;
				}
				
				if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
				{
					FillSolidRect(hDC, rcTemp, m_pBar->m_crXPMRUBackground);
				}
				else
				{
					m_pBar->DrawEdge(hDC, rcTemp, BDR_RAISEDINNER, BF_RECT);
					
					rcTemp.Inflate(-1, -1);
					if (m_pBar->HasTexture() && ddBOPopups & m_pBar->bpV1.m_dwBackgroundOptions)
					{
						if (m_pBar->m_hBrushTexture)
							UnrealizeObject(m_pBar->m_hBrushTexture);

						SetBrushOrgEx(hDC, -rcTemp.left, -rcTemp.top, NULL);
						m_pBar->FillTexture(hDC, rcTemp);
						SetBrushOrgEx(hDC, 0, 0, NULL);
					}
					else
						FillSolidRect(hDC, rcTemp, m_pBar->m_crBackground);
				}
			}

			if (ddBTPopup != bpV1.m_btBands)
			{
				//
				// Sizing special types of tools
				//

				switch (pTool->tpV1.m_ttTools)
				{
				case ddTTSeparator:
					if (pTool->m_bVerticalSeparator)
					{
						m_prcTools[nTool].left = 0;
						m_prcTools[nTool].right = rcPaint.Width();
					}
					else
						m_prcTools[nTool].bottom = m_prcTools[nTool].top + m_aLineHeight.GetAt(pTool->m_nBandLine-1);
					break;

				case CTool::ddTTMoreTools:
					if (!bLastTool)
					{
						if (VARIANT_TRUE == m_pRootBand->bpV1.m_vbWrapTools || !bVertical)
							m_prcTools[nTool].bottom = m_prcTools[nTool].top + m_aLineHeight.GetAt(pTool->m_nBandLine-1) - CBand::eBevelBorder2;
					}
					break;

				case ddTTCombobox:
				case ddTTEdit:
					{
						if (!bVertical && !pTool->m_bAdjusted)
						{
							pTool->m_bAdjusted = TRUE;
							if (ddBFStretch & bpV1.m_dwFlags && NULL == m_pFloat && VARIANT_FALSE == bpV1.m_vbWrapTools)
							{
								switch (pTool->tpV1.m_asAutoSize)
								{
								case ddTASSpringWidth:
									{
										nHeight = m_prcTools[nTool].Height();
										m_prcTools[nTool].top += (m_aLineHeight.GetAt(pTool->m_nBandLine-1) - nHeight) / 2;
										m_prcTools[nTool].bottom = m_prcTools[nTool].top + nHeight;

										int nWidth = rcPaint.Width() - m_prcTools[nToolCount-1].right;
										if (nWidth > 0)
										{
											if (ddBTStatusBar == bpV1.m_btBands && m_bStatusBarSizerPresent)
												nWidth -= GetSystemMetrics(SM_CXSIZE);
											m_prcTools[nTool].right = m_prcTools[nTool].left + nWidth + m_prcTools[nTool].Width();
											for (int nToolAutoSize = nTool+1; nToolAutoSize < nToolCount; nToolAutoSize++)
												m_prcTools[nToolAutoSize].Offset(nWidth, 0);
										}
									}
									break;

								case ddTASSpringHeight:
									m_prcTools[nTool].bottom = m_prcTools[nTool].top + m_sizeLargestTool.cy;
									break;

								case ddTASSpringBoth:
									{
										//
										// Height
										//

										m_prcTools[nTool].bottom = m_prcTools[nTool].top + m_sizeLargestTool.cy;

										//
										// Width
										//

										int nWidth = rcPaint.Width() - m_prcTools[nToolCount-1].right - eBevelBorder2;
										if (nWidth > 0)
										{
											if (ddBTStatusBar == bpV1.m_btBands && m_bStatusBarSizerPresent)
												nWidth -= GetSystemMetrics(SM_CXSIZE);
											m_prcTools[nTool].right = m_prcTools[nTool].left + nWidth + m_prcTools[nTool].Width();
											for (int nToolAutoSize = nTool+1; nToolAutoSize < nToolCount; nToolAutoSize++)
												m_prcTools[nToolAutoSize].Offset(nWidth, 0);
										}
									}
									break;

								default:
									{
										nHeight = m_prcTools[nTool].Height();
										m_prcTools[nTool].top += (m_aLineHeight.GetAt(pTool->m_nBandLine-1) - nHeight) / 2;
										m_prcTools[nTool].bottom = m_prcTools[nTool].top + nHeight;
									}
									break;
								}
							}
							else
							{
								nHeight = m_prcTools[nTool].Height();
								m_prcTools[nTool].top += (m_aLineHeight.GetAt(pTool->m_nBandLine-1) - nHeight) / 2;
								m_prcTools[nTool].bottom = m_prcTools[nTool].top + nHeight;
							}
						}
					}
					break;

				default:
					{
						if (!pTool->m_bAdjusted)
						{
							int nAdjustWidth = -1;
							int nAdjustHeight = -1;
							pTool->m_bAdjusted = TRUE;
							BOOL bMenuBarButtons = (ddBTMenuBar == bpV1.m_btBands || ddBTChildMenuBar == bpV1.m_btBands) && CBar::eSysMenu & m_pBar->m_dwMdiButtons;
							if (ddBFStretch & bpV1.m_dwFlags && NULL == m_pFloat && VARIANT_FALSE == bpV1.m_vbWrapTools)
							{
								switch (pTool->tpV1.m_asAutoSize)
								{
								case ddTASSpringWidth:
									{
										if (!bVertical)
										{
											int nWidth = rcPaint.Width() - m_prcTools[nToolCount-1].right - eBevelBorder2;
											if (bMenuBarButtons)
												nWidth = rcPaint.Width() - m_prcTools[nToolCount-2].right - m_prcTools[nToolCount-1].Width() - eBevelBorder2;
											if (ddBTStatusBar == bpV1.m_btBands && m_bStatusBarSizerPresent)
												nWidth -= GetSystemMetrics(SM_CXSIZE);
											if (nWidth < 0)
												nWidth = 0;
											m_prcTools[nTool].right = m_prcTools[nTool].left + nWidth + m_prcTools[nTool].Width();
											nAdjustWidth = m_prcTools[nTool].Width();
											for (int nToolAutoSize = nTool+1; nToolAutoSize < (bMenuBarButtons ? nToolCount - 1 : nToolCount); nToolAutoSize++)
												m_prcTools[nToolAutoSize].Offset(nWidth, 0);
										}
									}
									break;

								case ddTASSpringHeight:
									{
										if (bVertical)
										{
											if (nToolCount > 1)
											{
												//
												// Height
												//

												int nWidth = rcPaint.Height() - m_prcTools[nToolCount-1].bottom - eBevelBorder2;					
												if (bMenuBarButtons)
													nWidth = rcPaint.Height() - m_prcTools[nToolCount-2].bottom - m_prcTools[nToolCount-1].Height() - eBevelBorder2;
												if (nWidth < 0)
													nWidth = 0;

												m_prcTools[nTool].bottom = m_prcTools[nTool].top + nWidth + m_prcTools[nTool].Height();
												nAdjustHeight = m_prcTools[nTool].Height();
												for (int nToolAutoSize = nTool+1; nToolAutoSize < (bMenuBarButtons ? nToolCount - 1 : nToolCount); nToolAutoSize++)
													m_prcTools[nToolAutoSize].Offset(0, nWidth);
											}
											else
											{
												m_prcTools[nTool].right = m_prcTools[nTool].left + rcPaint.Width() - 20;
												m_prcTools[nTool].bottom = m_prcTools[nTool].top + rcPaint.Height() - 20;
											}
										}
										else
										{
											m_prcTools[nTool].bottom = m_prcTools[nTool].top + m_sizeLargestTool.cy;
											nAdjustHeight = m_prcTools[nTool].Height();
										}
									}
									break;

								case ddTASSpringBoth:
									{
										if (bVertical)
										{
											if (nToolCount > 1)
											{
												//
												// Width
												//

												m_prcTools[nTool].right = m_prcTools[nTool].left + m_sizeLargestTool.cx;
												nAdjustWidth = m_prcTools[nTool].Width();

												//
												// Height
												//

												int nWidth = rcPaint.Height() - m_prcTools[nToolCount-1].bottom - eBevelBorder2;					
												if (bMenuBarButtons)
													nWidth = rcPaint.Height() - m_prcTools[nToolCount-2].bottom - m_prcTools[nToolCount-1].Height() - eBevelBorder2;
												if (nWidth < 0)
													nWidth = 0;

												m_prcTools[nTool].bottom = m_prcTools[nTool].top + nWidth + m_prcTools[nTool].Height();
												nAdjustHeight = m_prcTools[nTool].Height();
												for (int nToolAutoSize = nTool+1; nToolAutoSize < (bMenuBarButtons ? nToolCount - 1 : nToolCount); nToolAutoSize++)
													m_prcTools[nToolAutoSize].Offset(0, nWidth);
											}
											else
											{
												m_prcTools[nTool].right = m_prcTools[nTool].left + rcPaint.Width() - 20;
												m_prcTools[nTool].bottom = m_prcTools[nTool].top + rcPaint.Height() - 20;
											}
										}
										else
										{
											//
											// Height
											//

											m_prcTools[nTool].bottom = m_prcTools[nTool].top + m_sizeLargestTool.cy;
											nAdjustHeight = m_prcTools[nTool].Height();

											//
											// Width
											//

											int nWidth = rcPaint.Width() - m_prcTools[nToolCount-1].right - eBevelBorder2;
											if (bMenuBarButtons)
												nWidth = rcPaint.Width() - m_prcTools[nToolCount-2].right - m_prcTools[nToolCount-1].Width() - eBevelBorder2;
											if (ddBTStatusBar == bpV1.m_btBands && m_bStatusBarSizerPresent)
												nWidth -= GetSystemMetrics(SM_CXSIZE);
											if (nWidth < 0)
												nWidth = 0;
											m_prcTools[nTool].right = m_prcTools[nTool].left + nWidth + m_prcTools[nTool].Width();
											nAdjustWidth = m_prcTools[nTool].Width();
											for (int nToolAutoSize = nTool+1; nToolAutoSize < (bMenuBarButtons ? nToolCount - 1 : nToolCount); nToolAutoSize++)
												m_prcTools[nToolAutoSize].Offset(nWidth, 0);
										}
									}
									break;
								}
							}
							if (ddTTControl == pTool->tpV1.m_ttTools && pTool->m_pDispCustom)
							{
								HRESULT hResult;
								VARIANT vProperty;
								vProperty.vt = VT_R4;
								SIZE size = {nAdjustWidth, nAdjustHeight};
								PixelToTwips(&size, &size);
								if (-1 != nAdjustWidth)
								{
									vProperty.fltVal = (float)size.cx;
									hResult = PropertyPut(pTool->m_pDispCustom, L"Width", vProperty);
								}
								if (-1 != nAdjustHeight)
								{
									vProperty.fltVal = (float)size.cy;
									hResult = PropertyPut(pTool->m_pDispCustom, L"Height", vProperty);
								}
							}
						}
					}
					break;
				}
			}
			
			rcTool = m_prcTools[nTool];
			rcTool.Offset(rcPaint.left, rcPaint.top);
			if (m_pParentBand)
			{
				switch (m_pParentBand->bpV1.m_cbsChildStyle)
				{
				case ddCBSToolbarTopTabs:
				case ddCBSToolbarBottomTabs:
					if ((bVertical ? rcTool.right > rcPaint.right : rcTool.bottom > rcPaint.bottom))
					{
						pTool->SetVisibleOnPaint(FALSE);
						continue;
					}
					else
						goto DrawTool;
					break;

				case ddCBSSlidingTabs:
					rcTool.Offset(-m_pParentBand->m_pChildBands->m_nHorzSlidingOffset, 
							      -m_pParentBand->m_pChildBands->m_nVertSlidingOffset);
					break;
				}
			}

			//
			// Check for clipping at end
			//
			
			if ((!bWrapFlag && (ddDAFloat != m_pRootBand->bpV1.m_daDockingArea || ddCBSSlidingTabs == m_pRootBand->bpV1.m_cbsChildStyle) && ddDAPopup != m_pRootBand->bpV1.m_daDockingArea))
			{
				if (bLastTool)
				{
					if (pTool != m_pTools->m_pMoreTools)
						pTool->SetVisibleOnPaint(FALSE);
					continue;
				}

				if (bVertical)
				{
					if (m_pTools->m_pMoreTools)
						bExceededLimit = ((nTool != nLastVisibleTool && (rcTool.bottom + (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 20 : 11) + eBevelBorder2) > nLimit) || (nTool == nLastVisibleTool && rcTool.bottom > nLimit));
					else
						bExceededLimit = ((nTool != nLastVisibleTool && (rcTool.bottom + (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 18 : 10)) > nLimit) || (nTool == nLastVisibleTool && rcTool.bottom > nLimit));
				}
				else
				{
					bExceededLimit = (nTool == nLastVisibleTool && rcTool.right > nLimit);
					if (!bExceededLimit)
					{
						if (m_pTools->m_pMoreTools)
							bExceededLimit = (nTool != nLastVisibleTool && (rcTool.right + (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 20 : 11) + eBevelBorder2) > nLimit);
						else
							bExceededLimit = (nTool != nLastVisibleTool && (rcTool.right + (m_pBar && VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons ? 18 : 10)) > nLimit);
					}
				}

				if (bExceededLimit)
				{
					if (!bLastTool)
					{
						bLastTool = TRUE;
						if (m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle)
						{
							if (bVertical)
								m_pParentBand->m_pChildBands->SetBottomButtonStyle(BTabs::eBottomButton);
							else
								m_pParentBand->m_pChildBands->SetBottomButtonStyle(BTabs::eRightButton);

							goto DrawTool;
						}
						pTool->SetVisibleOnPaint(FALSE);
						if (m_pTools->m_pMoreTools)
						{
							m_pTools->m_pMoreTools->m_bMoreTools = TRUE;
							rcTool = m_pTools->m_pMoreTools->m_rcTemp;
							rcTool.Offset(rcPaint.left, rcPaint.top);
							if (bVertical)
							{
								rcTool.left = rcPaint.left + eBevelBorder2;
								rcTool.right = rcPaint.right - eBevelBorder2;
								int nHeight = rcTool.Height();
								rcTool.bottom = rcPaint.bottom - eBevelBorder2;
								rcTool.top = rcTool.bottom - nHeight;
							}
							else
							{
								int nWidth = rcTool.Width();
								rcTool.right = rcPaint.right - eBevelBorder2;
								rcTool.left = rcTool.right - nWidth;
								rcTool.top = rcPaint.top + eBevelBorder2;
								rcTool.bottom = rcPaint.bottom - eBevelBorder2;
							}
							m_prcTools[nToolCount-1] = rcTool;
							m_prcTools[nToolCount-1].Offset(-rcPaint.left, -rcPaint.top);
							m_pTools->m_pMoreTools->SetVisibleOnPaint(TRUE);
							DrawTool(hDC, m_pTools->m_pMoreTools, rcTool, bVerticalPaint, ptPaintOffset, bDrawControlsOrForms);
						}
						else if (VARIANT_FALSE == bpV1.m_vbWrapTools && ddBTStatusBar != bpV1.m_btBands)
						{
							rcTool = rcPaint;
							if (bVertical)
								rcTool.top = rcTool.bottom - eClipMarkerHeight - eBevelBorder2;
							else
								rcTool.left = rcTool.right - eClipMarkerHeight - eBevelBorder2;
							CTool::DrawClipMarker(hDC, m_pBar, rcTool, bVertical, VARIANT_TRUE == m_pBar->bpV1.m_vbLargeIcons);
						}
						continue;
					}
				}
			} 
DrawTool:
			pTool->SetVisibleOnPaint(TRUE);
			DrawTool(hDC, pTool, rcTool, bVerticalPaint, ptPaintOffset, bDrawControlsOrForms);
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
// ParentWindowedTools
//

void CBand::ParentWindowedTools(HWND hWndParent)
{
	m_pTools->ParentWindowedTools(hWndParent);
	if (m_pChildBands)
		m_pChildBands->ParentWindowedTools(hWndParent);
}

//
// HideWindowedTools
//

void CBand::HideWindowedTools()
{
	m_pTools->HideWindowedTools();
	if (m_pChildBands)
		m_pChildBands->HideWindowedTools();
}

//
// TabbedWindowedTools
//

void CBand::TabbedWindowedTools(BOOL bShow)
{
	try
	{
		m_pTools->TabbedWindowedTools(bShow);
		if (m_pChildBands)
			m_pChildBands->TabbedWindowedTools(bShow);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// IncrementFirstTool
//

void CBand::IncrementFirstTool()
{
	if (m_nFirstTool < m_pTools->GetVisibleToolCount() - 1)
		m_nFirstTool++;
}

//
// ExchangeMenuUsageData
//

HRESULT CBand::ExchangeMenuUsageData(IStream* pStream, VARIANT_BOOL vbSave)
{
	HRESULT hResult;
	CTool* pTool;
	int nToolCount = m_pTools->GetToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = m_pTools->GetTool(nTool);
		if (VARIANT_TRUE == vbSave)
			hResult = pStream->Write(&pTool->tpV1.m_nUsage, sizeof(pTool->tpV1.m_nUsage), NULL);
		else
			hResult = pStream->Read(&pTool->tpV1.m_nUsage, sizeof(pTool->tpV1.m_nUsage), NULL);
		if (FAILED(hResult))
			return hResult;
	}
	if (m_pChildBands)
	{
		CBand* pBand;
		int nChildBandCount = m_pChildBands->GetChildBandCount();
		for (int nBand = 0; nBand < nChildBandCount; nBand++)
		{
			pBand = m_pChildBands->GetChildBand(nBand);
			hResult = pBand->ExchangeMenuUsageData(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;
		}
	}
	return NOERROR;
}

//
// ClearMenuUsageData
//

HRESULT CBand::ClearMenuUsageData()
{
	HRESULT hResult;
	int nToolCount = m_pTools->GetToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
		m_pTools->GetTool(nTool)->tpV1.m_nUsage = 0;
	if (m_pChildBands)
	{
		CBand* pBand;
		int nChildBandCount = m_pChildBands->GetChildBandCount();
		for (int nBand = 0; nBand < nChildBandCount; nBand++)
		{
			pBand = m_pChildBands->GetChildBand(nBand);
			hResult = pBand->ClearMenuUsageData();
			if (FAILED(hResult))
				return hResult;
		}
	}
	return NOERROR;
}

int CBand::Height()
{
	int nHeight = 0;
	switch(bpV1.m_daDockingArea)
	{
	case ddDAFloat:
		nHeight = bpV1.m_rcFloat.Height();
		break;
	case ddDATop :
	case ddDABottom :
	case ddDALeft :
	case ddDARight:
		nHeight = m_rcDock.Height();
		break;
	default:		
		nHeight=bpV1.m_rcDimension.bottom;
		break;
	}
	return nHeight;
}
