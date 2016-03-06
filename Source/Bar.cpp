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
#include <exception>
#include <stddef.h>       // for offsetof()
#include "IpServer.h"
#include "Support.h"
#include "Resource.h"
#include "Debug.h"
#include "XEvents.h"
#include "Streams.h"
#include "Dispids.h"
#include "Guids.h"
#include "Errors.h"
#include "Tool.h"
#include "BTabs.h"
#include "ChildBands.h"
#include "Bands.h"
#include "Custom.h"
#include "CBList.h"
#include "ImageMgr.h"
#include "ShortCuts.h"
#include "Returnbool.h"
#include "Returnstring.h"
#include "hlp\ActiveBar2Ref.hh"
#include "Dib.h"
#include "FWnd.h"
#include "Dock.h"
#include "Globals.h"
#include "FDialog.h"
#include "MiniWin.h"
#include "ToolTip.h"
#include "PopupWin.h"
#include "StaticLink.h"
#include "WindowProc.h"
#include "Interfaces.h"
#include "Localizer.h"
#include "Designer\DesignerInterfaces.h"
#include "TpPopup.h"
#include "Bar.h"

//
// license entry:{F89413C0-5F92-11D1-AA4D-0060081C43D9}
//

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//{OLE VERBS}
OLEVERB CBarVerbs[]={
	{OLEIVERB_SHOW,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_HIDE,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_INPLACEACTIVATE,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_PRIMARY,(LPWSTR)NULL,0x0,0x0},
	{OLEIVERB_UIACTIVATE,(LPWSTR)NULL,0x0,0x0},
	{ID_VERB_DESIGNER,(LPWSTR)IDS_VERB_DESIGNER,0x0,0x3},
};
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CBar(IUnknown *pUnkOuter)
{
	CBar *pObject;
	pObject=new CBar(pUnkOuter);
	if (!pObject)
		return NULL;
	else
		return pObject->PrivateUnknown();
}

const CLSID *ActiveBar2_PropPages[]={
&CLSID_DesignerPage
};

DEFINE_CONTROLOBJECT(&CLSID_ActiveBar2,ActiveBar2,_T("ActiveBar2 Class"),CreateFN_CBar,2,0,&IID_IActiveBar2,_T(""),
	&DIID_IActiveBar2Events,
	OLEMISC_RECOMPOSEONRESIZE|
	OLEMISC_CANTLINKINSIDE|
	OLEMISC_INSIDEOUT|
	OLEMISC_ACTIVATEWHENVISIBLE|
	OLEMISC_ALIGNABLE|
	OLEMISC_SIMPLEFRAME|
	OLEMISC_SETCLIENTSITEFIRST,
0,
	FALSE,
	FALSE,
	1,	1,ActiveBar2_PropPages,0,NULL
	);
void *CBar::objectDef=&ActiveBar2Object;
CBar *CBar::CreateInstance(IUnknown *pUnkOuter)
{
	return new CBar(pUnkOuter);
}
//{OBJECT CREATEFN}
CBar::CBar(IUnknown *pUnkOuter)
	: m_pUnkOuter(pUnkOuter==NULL ? &m_UnkPrivate : pUnkOuter)
	  ,m_bstrDataPath(NULL),
	  m_pFreezeBand(NULL),
	  m_bstrHelpFile(NULL),
	  m_hFontMenuVert(NULL),
	  m_hFontMenuHorz(NULL),
	  m_hFontToolTip(NULL),
	  m_hFontSmall(NULL),
	  m_hFontSmallVert(NULL),
	  m_hFontChildBand(NULL),
	  m_hFontChildBandVert(NULL),
	  m_hBrushTexture(NULL),
	  m_hBrushForeColor(NULL),
	  m_hBrushBackColor(NULL),
	  m_hBrushShadowColor(NULL),
	  m_hBrushHighLightColor(NULL),
	  m_pCustomizeDragLock(NULL),
	  m_pvbStatusBandToolsVisible(NULL),
	  m_pDesignerSettingNotify(NULL),
	  m_pStatusBandTool(NULL),
	  m_pStatusBand(NULL),
	  m_pPopupRoot(NULL),
	  m_pDesignerNotify(NULL),
	  m_pActiveBand(NULL),
	  m_pActiveTool(NULL),
	  m_pCustomize(NULL),
	  m_hWndMDIClient(NULL),
	  m_pImageMgrBandSpecific(NULL),
	  m_hWndDock(NULL),
	  m_pMainProc(NULL),
	  m_pMDIClientProc(NULL),
	  m_pTip(NULL),
	  m_pDragDropManager(NULL),
	  m_hAccel(NULL),
	  m_pDesigner(NULL),
	  m_pColorQuantizer(NULL),
	  m_pDesignerImpl(NULL),
	  m_hWndDesignerShutDownAdvise(NULL),
	  m_hWndMenuChild(NULL),
	  m_hWndModifySelection(NULL),
	  m_pEditToolPopup(NULL),
	  m_pRelocate(NULL),
	  m_cached_fhFont(NULL),
	  m_cachedControlFont(NULL),
	  m_hStatusbarThread(NULL),
	  m_pTabbedTool(NULL),
	  m_bstrWindowClass(NULL),
	  m_pEdit(NULL),
	  m_pCombo(NULL)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
	m_pClientSite=NULL;
	m_pControlSite=NULL;
	m_pInPlaceSiteWndless=NULL;
	m_pSimpleFrameSite=NULL;
	m_pOleAdviseHolder=NULL;
	m_pInPlaceSite=NULL;
	m_fDirty=FALSE;
	m_hWnd=NULL;
	m_hWndParent=NULL;
	m_hWndReflect=NULL;
	m_pInPlaceFrame=NULL;
    m_pInPlaceUIWindow=NULL;
	m_pDispAmbient=NULL;
	m_Size.cx=32;
	m_Size.cy=32;
	m_fInPlaceActive=FALSE;
	m_fUIActive=FALSE;
	m_hRgn=NULL;
	m_fUsingWindowRgn=FALSE;
    m_fViewAdvisePrimeFirst = FALSE;
    m_fViewAdviseOnlyOnce = FALSE;
	m_pViewAdviseSink=NULL;
	m_fSaveSucceeded = FALSE;
	m_cpEvents=new CConnectionPoint(this,SINK_TYPE_EVENT);
	m_cpPropNotify=new CConnectionPoint(this,SINK_TYPE_PROPNOTIFY);
// {END INIT}
	OnFontHeightChanged();
	
	m_pBands = CBands::CreateInstance(NULL);
	assert(m_pBands);
	if (m_pBands)
		m_pBands->SetOwner(this);
	
	m_dwThreadId = -1;

	m_pTools = CTools::CreateInstance(NULL);
	assert(m_pTools);
	if (m_pTools)
		m_pTools->SetBar(this);
	
	m_pImageMgr = CImageMgr::CreateInstance(NULL);
	assert(m_pImageMgr);
	
	m_pShortCuts = new CShortCuts;
	assert(m_pShortCuts);
	if (m_pShortCuts)
		m_pShortCuts->SetOwner(this);

	m_pDockMgr = new CDockMgr(this);
	assert(m_pDockMgr);

	m_theChildMenus.SetBar(this);

	m_bstrHelpFile = SysAllocString(L"Activebar20.hlp");
	assert(m_bstrHelpFile);

	m_bstrWindowClass = SysAllocString(L"MDICLIENT");

	m_fhFont.InitializeFont(&GetGlobals()._fdDefault);
	m_fhControl.InitializeFont(&GetGlobals()._fdDefaultControl);
	m_fhChildBand.InitializeFont(&GetGlobals()._fdDefaultControl);
	m_fhFont.SetPropertyNotifySink(&m_xPropNotify);
	m_fhControl.SetPropertyNotifySink(&m_xPropNotify);
	m_fhChildBand.SetPropertyNotifySink(&m_xPropNotify);

	m_nChildBandFontHeight = 0;
	m_nChildBandFontHeightVert = 0;
	m_dwMdiButtons = 0;
	m_bIsVisualFoxPro = FALSE;
	m_bStartedToolBars = FALSE;
	m_bApplicationActive = TRUE;
	m_bMenuLoop = FALSE;
	m_bToolModal = FALSE;
	m_bClickLoop = FALSE;
	m_bHasTexture = FALSE;
	m_bCustomization = FALSE;
	m_bWhatsThisHelp = FALSE;
	m_bBandDestroyLock = FALSE;
	m_bNeedRecalcLayout = FALSE;
	m_bPopupMenuExpanded = FALSE;
	m_bFirstRecalcLayout = FALSE;
	m_bInitialRecalcLayout = FALSE;
	m_bWindowPosChangeMessage = FALSE;
	m_bStopStatusBarThread = FALSE;
	m_bToolCreateLock = FALSE;
	m_bIE = FALSE;
	m_bFirstPopup = FALSE;
	m_vbCustomizeModified = VARIANT_FALSE;
	m_nDragStartTick = 0;
	m_bFreezeEvents = FALSE;
	m_ptDragStart.x = m_ptDragStart.y = 0;
	m_vbLibrary = VARIANT_FALSE;

	m_bUserMode = -1;

	memset(&m_diCustSelection, 0, sizeof(m_diCustSelection));

	m_strMsgTitle.LoadString(IDS_ACTIVEBAR);

	CacheColors();

	m_eAppType = eSDIForm;

	m_fInPlaceActive = FALSE;
	m_bAltCustomization = FALSE;
	m_bSaveImages = FALSE;
	m_bBandSpecificConvertIds = FALSE;
	m_bStatusBandLock = FALSE;

	m_theErrorObject.m_pEvents = m_cpEvents->m_pControl;
	m_theErrorObject.objectDef = objectDef;
	m_theErrorObject.m_pErrorTable = m_etErrors;
	m_theErrorObject.m_nErrorTableSize = 31;
	m_theErrorObject.m_bstrHelpFile = SysAllocString(m_bstrHelpFile);
	m_dwLastDocked = 0;
	m_pColorQuantizer = new CColorQuantizer(236, 6);
	assert(m_pColorQuantizer);
	
	m_hStatusBarEvent = CreateEvent(NULL, FALSE, TRUE, NULL);

	m_nFlags = 0;

	m_bShutdownLock = FALSE;

	m_pToolShadow = new CPopupWinToolAndShadow();
	assert(m_pToolShadow);
	if (m_pToolShadow)
		m_pToolShadow->Bar(this);

	m_pLocalizer = new CLocalizer;
	assert(m_pLocalizer);
	if (m_pLocalizer)
	{
		static const long nLocalizableStrings[] = {IDS_CUSTOMIZE, 
												   IDS_TOOLBAR, 
												   -1,
												   -1,
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   IDS_COMMANDS, 
												   IDS_OPTIONS, 
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   -1,
												   -1, 
												   -1, 
												   -1, 
												   IDS_KEYBOARD, 
												   IDS_CLOSEBUTTON, 
												   -1, 
												   -1,
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   -1, 
												   IDS_CONTEXTCUSTOMIZE,
												   -1,
												   -1, 
												   -1,
												   -1,
												   -1,
												   -1,
												   -1,
												   IDS_MOREBUTTONS,
												   -1,
												   -1,
												   -1,
												   -1,
												   IDS_MINIMIZEWINDOW,
												   IDS_RESTOREWINDOW,
												   IDS_CLOSEWINDOW,
												   IDS_ALLTOOLSCAPTION};
		m_pLocalizer->Add(50, nLocalizableStrings);
	}

	m_pMoreTools = CBand::CreateInstance(NULL);
	assert(m_pMoreTools);
	if (m_pMoreTools)
	{
		DDString strTemp;
		strTemp.LoadString(IDS_POPUPMORETOOLS);
		MAKE_WIDEPTR_FROMTCHAR(wPopup, strTemp);
		m_pMoreTools->put_Name(wPopup);
		m_pMoreTools->bpV1.m_daDockingArea = ddDAPopup;
		m_pMoreTools->bpV1.m_vbWrapTools = VARIANT_TRUE;
		m_pMoreTools->put_Width(400);
		m_pMoreTools->put_Height(2000);
		m_pMoreTools->SetOwner(this, TRUE);
	}

	m_vAlign.vt = VT_I2;
	m_vAlign.iVal = 0;

	m_pAllTools = CBand::CreateInstance(NULL);
	assert(m_pAllTools);
	if (m_pAllTools)
	{
		DDString strTemp;
		strTemp.LoadString(IDS_POPUPALLTOOLS);
		MAKE_WIDEPTR_FROMTCHAR(wPopup, strTemp);
		m_pAllTools->put_Name(wPopup);
		m_pAllTools->bpV1.m_btBands = ddBTPopup;
		m_pAllTools->SetOwner(this, TRUE);
		m_pAllTools->m_bDontExpand = TRUE;
		m_pAllTools->m_bAllTools = TRUE;
	}
	
#ifdef _DEBUG
	TRACE1(3, _T("CBar: %i\n"), sizeof(CBar));
	TRACE1(3, _T("CBand: %i\n"), sizeof(CBand));
	TRACE1(3, _T("CBands: %i\n"), sizeof(CBands));
	TRACE1(3, _T("CTools: %i\n"), sizeof(CTools));
	TRACE1(3, _T("CTool: %i\n"), sizeof(CTool));
	TRACE1(3, _T("CChildBands %i\n"), sizeof(CChildBands));
	TRACE1(3, _T("CShortCut %i\n"), sizeof(CShortCut));
	TRACE1(3, _T("CShortCuts: %i\n"), sizeof(CShortCuts));
	TRACE1(3, _T("CImageMgr %i\n"), sizeof(CImageMgr));
	TRACE1(3, _T("ImageSizeEntry: %i\n"), sizeof(ImageSizeEntry));
	TRACE1(3, _T("ImageEntry: %i\n"), sizeof(ImageEntry));
	TRACE1(3, _T("ErrorObject: %i\n"), sizeof(ErrorObject));
	TRACE1(3, _T("CMiniWin: %i\n"), sizeof(CMiniWin));
	TRACE1(3, _T("CPopupWin: %i\n"), sizeof(CPopupWin));
	TRACE1(3, _T("CDock: %i\n"), sizeof(CDock));
	TRACE1(3, _T("CDockMgr: %i\n"), sizeof(CDockMgr));
	TRACE1(3, _T("CToolTip: %i\n"), sizeof(CToolTip));
	TRACE1(3, _T("CConnectionPoint: %i\n"), sizeof(CConnectionPoint));
#endif
}

CBar::~CBar()
{
	if (m_pClientSite)
	{
		// Powerbuilder bug. 
		// If g_cLocks=0 at this point the WindowProc Chain Crashes 
		// in unloaded DLL memory area
		Detach(); 
	}
	else
		InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
	if (m_pClientSite) m_pClientSite->Release();
	if (m_pControlSite) m_pControlSite->Release();
	if (m_pSimpleFrameSite) m_pSimpleFrameSite->Release();
	if (m_pOleAdviseHolder) m_pOleAdviseHolder->Release();
	if (m_pInPlaceSite) m_pInPlaceSite->Release();
	if (m_pInPlaceFrame) m_pInPlaceFrame->Release();
    if (m_pInPlaceUIWindow) m_pInPlaceUIWindow->Release();
    if (m_pDispAmbient) m_pDispAmbient->Release();
	if (m_pInPlaceSiteWndless) m_pInPlaceSiteWndless->Release();
	if (m_hRgn) DeleteObject(m_hRgn);
	if (m_cpEvents)
		m_cpEvents->Release();
	if (m_cpPropNotify)
		m_cpPropNotify->Release();
// {END CLEANUP}
	if (m_pEdit && m_pEdit->IsWindow())
	{
		// Self Destroyed
		m_pEdit->DestroyWindow();
	}

	if (m_pCombo && m_pCombo->IsWindow())
	{
		// Self Destroyed
		m_pCombo->DestroyWindow();
	}

	m_pDragDropManager = NULL;
	m_cpEvents = NULL;
	m_cpPropNotify = NULL;

	m_mapMenuBar.RemoveAll();
	delete m_pDockMgr;
	delete m_pRelocate;
	m_pStatusBandTool = NULL;
	m_pStatusBand = NULL;

	if (m_pBands)
		m_pBands->Release();

	if (m_pTools)
		m_pTools->Release();

	if (m_pMoreTools)
		m_pMoreTools->Release();

	if (m_pAllTools)
		m_pAllTools->Release();

	delete m_pShortCuts;

	delete m_pCustomize;

	delete m_pColorQuantizer;

	if (m_pLocalizer)
		m_pLocalizer->Release();

	if (m_hAccel)
		DestroyAcceleratorTable(m_hAccel);
	
	if (m_pImageMgr)
	{
		m_pImageMgr->Release();
		m_pImageMgr = NULL;
	}

	BOOL bResult;
	if (m_hFontMenuVert)
	{
		bResult = DeleteFont(m_hFontMenuVert);
		assert(bResult);
	}

	if (m_hFontMenuHorz)
	{
		bResult = DeleteFont(m_hFontMenuHorz);
		assert(bResult);
	}

	if (m_hFontToolTip)
	{
		bResult = DeleteFont(m_hFontToolTip);
		assert(bResult);
	}

	if (m_hFontSmall)
	{
		bResult = DeleteFont(m_hFontSmall);
		assert(bResult);
	}

	if (m_hFontSmallVert)
	{
		bResult = DeleteFont(m_hFontSmallVert);
		assert(bResult);
	}

	if (m_hFontChildBandVert)
	{
		bResult = DeleteFont(m_hFontChildBandVert);
		assert(bResult);
	}

	if (m_pTip)
		m_pTip->DestroyWindow();

	if (m_hBrushTexture)
	{
		bResult = DeleteBrush(m_hBrushTexture);
		assert(bResult);
	}
	
	if (m_hBrushForeColor)
	{
		bResult = DeleteBrush(m_hBrushForeColor);
		assert(bResult);
	}
	
	if (m_hBrushBackColor)
	{
		bResult = DeleteBrush(m_hBrushBackColor);
		assert(bResult);
	}
	
	if (m_hBrushHighLightColor)
	{
		bResult = DeleteBrush(m_hBrushHighLightColor);
		assert(bResult);
	}

	if (m_hBrushShadowColor)
	{
		bResult = DeleteBrush(m_hBrushShadowColor);
		assert(bResult);
	}

	if (m_bstrDataPath)
		SysFreeString(m_bstrDataPath);

	if (m_bstrHelpFile)
		SysFreeString(m_bstrHelpFile);

	if (m_bstrWindowClass)
		SysFreeString(m_bstrWindowClass);

	if (m_pDesignerImpl)
		m_pDesignerImpl->Release();

	if (m_pToolShadow)
	{
		if (m_pToolShadow->IsWindow())
			m_pToolShadow->DestroyWindow();
		delete m_pToolShadow;
	}

	CloseHandle(m_hStatusBarEvent);
}

//{DEF IUNKNOWN MEMBERS}
//=--------------------------------------------------------------------------=
// CBar::CPrivateUnknownObject::QueryInterface
//=--------------------------------------------------------------------------=
inline CBar *CBar::CPrivateUnknownObject::m_pMainUnknown(void)
{
    return (CBar *)((LPBYTE)this - offsetof(CBar, m_UnkPrivate));
}

STDMETHODIMP CBar::CPrivateUnknownObject::QueryInterface(REFIID riid,void **ppvObjOut)
{
    CHECK_POINTER(ppvObjOut);
    if (DO_GUIDS_MATCH(riid,IID_IUnknown)) 
	{
        m_cRef++;
        *ppvObjOut = (IUnknown *)this;
        return S_OK;
    } 
	else
        return m_pMainUnknown()->InternalQueryInterface(riid, ppvObjOut);
}

ULONG CBar::AddRef(void)
{
    return 0;
}
ULONG CBar::Release(void)
{
	return 0;
}
STDMETHODIMP CBar::QueryInterface(REFIID riid,void **ppvObjOut)
{
	return NULL;
}
ULONG CBar::CPrivateUnknownObject::AddRef(void)
{
    return ++m_cRef;
}

ULONG CBar::CPrivateUnknownObject::Release(void)
{
    ULONG cRef = --m_cRef;
    if (!m_cRef)
        delete m_pMainUnknown();
    return cRef;
}
HRESULT CBar::InternalQueryInterface(REFIID riid,void  **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IOleObject))
	{
		*ppvObjOut=(void *)(IOleObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleWindow))
	{
		*ppvObjOut=(void *)(IOleWindow *)(IOleInPlaceObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleInPlaceObject))
	{
		*ppvObjOut=(void *)(IOleInPlaceObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleInPlaceActiveObject))
	{
		*ppvObjOut=(void *)(IOleInPlaceActiveObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IViewObject))
	{
		*ppvObjOut=(void *)(IViewObject *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IViewObject2))
	{
		*ppvObjOut=(void *)(IViewObject2 *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPersist))
	{
		*ppvObjOut=(void *)(IPersist *)(IPersistPropertyBag *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPersistPropertyBag))
	{
		*ppvObjOut=(void *)(IPersistPropertyBag *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPersistStreamInit))
	{
		*ppvObjOut=(void *)(IPersistStreamInit *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IPersistStorage))
	{
		*ppvObjOut=(void *)(IPersistStorage *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IConnectionPointContainer))
	{
		*ppvObjOut=(void *)(IConnectionPointContainer *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISpecifyPropertyPages))
	{
		*ppvObjOut=(void *)(ISpecifyPropertyPages *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IProvideClassInfo))
	{
		*ppvObjOut=(void *)(IProvideClassInfo *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IBarPrivate))
	{
		*ppvObjOut=(void *)(IBarPrivate *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IOleControl))
	{
		*ppvObjOut=(void *)(IOleControl *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISupportErrorInfo))
	{
		*ppvObjOut=(void *)(ISupportErrorInfo *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ICategorizeProperties))
	{
		*ppvObjOut=(void *)(ICategorizeProperties *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IDispatch))
	{
		*ppvObjOut=(void *)(IDispatch *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IActiveBar2))
	{
		*ppvObjOut=(void *)(IActiveBar2 *)this;
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
	#ifdef _DEBUG
	TCHAR tmp[141];
	wsprintf(tmp,_T("QI failed for {0x%08X,0x%X,0x%X,{0x%hX,0x%hX,0x%hX,0x%hX,0x%hX,0x%hX,0x%hX,0x%hX}}\n"),
		(unsigned int)riid.Data1,(unsigned int)riid.Data2,(unsigned int)riid.Data3,(unsigned int)riid.Data4[0],(unsigned int)riid.Data4[1],
		(unsigned int)riid.Data4[2],(unsigned int)riid.Data4[3],(unsigned int)riid.Data4[4],(unsigned int)riid.Data4[5],
		(unsigned int)riid.Data4[6],(unsigned int)riid.Data4[7]);
	TRACET(tmp);
	#endif
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
//{ENDDEF IUNKNOWN MEMBERS}










STDMETHODIMP CBar::SetClientSite( IOleClientSite *pClientSite)
{
	m_bUserMode = -1;

	TRACE(1, "IOleObject::SetClientSite\n");
	if (m_pClientSite)
	{
		m_pClientSite->Release();
		m_pClientSite = NULL;
	} 
	if (m_pControlSite)
	{
		m_pControlSite->Release();
		m_pControlSite = NULL;
	}
	if (m_pSimpleFrameSite)
	{
		m_pSimpleFrameSite->Release();
		m_pSimpleFrameSite = NULL;
	}
	if (m_pInPlaceSite)
	{
		m_pInPlaceSite->Release();
		m_pInPlaceSite = NULL;
	}
	if (NULL == pClientSite)
	{
		if (m_pDesignerSettingNotify)
		{
			delete m_pDesignerSettingNotify;
			m_pDesignerSettingNotify = NULL;
		}
		if (m_pDesignerImpl)
			m_pDesignerImpl->CloseDesigner();
		if (m_bCustomization)
			DoCustomization(FALSE);
		Detach();
		if (m_pDispAmbient) 
		{
			m_pDispAmbient->Release();
			m_pDispAmbient = NULL;
		}
	}
	m_pClientSite = pClientSite;
    if (m_pClientSite) 
	{
        m_pClientSite->AddRef();
        m_pClientSite->QueryInterface(IID_IOleControlSite, (void **)&m_pControlSite);
        if (((CONTROLOBJECTDESC*)objectDef)->dwOleMiscFlags & OLEMISC_SIMPLEFRAME)
            m_pClientSite->QueryInterface(IID_ISimpleFrameSite, (void **)&m_pSimpleFrameSite);

		//
		// Get the parent form's Background Color
		//

		OLE_COLOR ocBackColor;
		if (GetAmbientProperty(DISPID_AMBIENT_BACKCOLOR, VT_I4, &ocBackColor))
			OleTranslateColor(ocBackColor, 0, &m_crAmbientBackground);

		Attach(NULL);

		if (!AmbientUserMode())
		{
			m_pDesignerSettingNotify = new CDesignTimeSystemSettingNotify(this);
			if (m_pDesignerSettingNotify)
				m_pDesignerSettingNotify->Create();
		}

		IOleControlSite* pControlSite;
		HRESULT hResult = m_pClientSite->QueryInterface(IID_IOleControlSite, (void**)&pControlSite);
		if (SUCCEEDED(hResult))
		{
			LPDISPATCH pDispatch;
			hResult = pControlSite->GetExtendedControl(&pDispatch);
			pControlSite->Release();
			if (SUCCEEDED(hResult))
			{
				hResult = PropertyGet(pDispatch, L"Align", m_vAlign);
				pDispatch->Release();
			}
		}
    }
	return S_OK;
}
STDMETHODIMP CBar::GetClientSite( IOleClientSite **ppClientSite)
{
	TRACE(1, "IOleObject::GetClientSite\n");
	*ppClientSite=m_pClientSite;
	if (m_pClientSite)
		m_pClientSite->AddRef();
	return S_OK;
}
STDMETHODIMP CBar::SetHostNames( LPCOLESTR szContainerApp, LPCOLESTR szContainerObj)
{
	return S_OK;
}
STDMETHODIMP CBar::Close(DWORD dwSaveOption)
{
	TRACE(1, "IOleObject::Close\n");

	#ifdef DEF_IOLEINPLACEOBJECT
	if (m_pDesignerImpl)
		m_pDesignerImpl->CloseDesigner();

	if (IsWindow(m_hWnd) && m_pClientSite)
		m_pClientSite->OnShowWindow(FALSE);

    if (m_fInPlaceActive) 
	{
        HRESULT hr = InPlaceDeactivate();
        RETURN_ON_FAILURE(hr);
    }
	#endif
	if (m_pClientSite)
		m_pClientSite->OnShowWindow(FALSE);
    // handle the save flag.
    if ((dwSaveOption == OLECLOSE_SAVEIFDIRTY || dwSaveOption == OLECLOSE_PROMPTSAVE) && m_fDirty) 
	{
        if (m_pClientSite) 
			m_pClientSite->SaveObject();
        if (m_pOleAdviseHolder) 
			m_pOleAdviseHolder->SendOnSave();
    }
	if (m_pInPlaceSite)
	{
		m_pInPlaceSite->Release();
		m_pInPlaceSite=NULL;
	}
    if (m_pViewAdviseSink)
	{
		m_pViewAdviseSink->Release();
		m_pViewAdviseSink = NULL;
	}
    return S_OK;
}
STDMETHODIMP CBar::SetMoniker( DWORD dwWhichMoniker, IMoniker *pmk)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBar::GetMoniker( DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBar::InitFromData( IDataObject *pDataObject, BOOL fCreation, DWORD reserved)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBar::GetClipboardData( DWORD dwReserved, IDataObject **ppDataObject)
{
	*ppDataObject=NULL;
	return E_NOTIMPL;
}
STDMETHODIMP CBar::DoVerb( LONG lVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect)
{
	TRACE1(1, "IOleObject::DoVerb %d\n",lVerb);
	HRESULT hr;
	switch(lVerb)
	{
		case OLEIVERB_PRIMARY:
		case OLEIVERB_SHOW:
		case OLEIVERB_INPLACEACTIVATE:
		case OLEIVERB_UIACTIVATE:
#ifdef DEF_IOLEINPLACEOBJECT
			hr = InPlaceActivate(lVerb);
#else
			hr=NOERROR;
#endif
			OnVerb(lVerb);
			return (hr);
		case OLEIVERB_HIDE:
			#ifdef DEF_IOLEINPLACEOBJECT
			if (m_fUIActive)
				UIDeactivate();
			#endif
			if (m_fInPlaceVisible) 
				SetInPlaceVisible(FALSE);
			OnVerb(lVerb);
			return S_OK;

		case ID_VERB_DESIGNER:
		case ID_VERB_STANDALONE:
			{
				if (m_pDesignerImpl)
				{
					m_pDesignerImpl->SetFocus();
					return NOERROR;
				}

				HRESULT hResult = CoCreateInstance(CLSID_Designer, 
												   NULL, 
												   CLSCTX_INPROC_SERVER, 
												   IID_IDesigner, 
												   (LPVOID*)&m_pDesignerImpl);
				if (FAILED(hResult))
				{
					DDString strError;
					strError.LoadString(IDS_ERR_DESIGNERFAILEDTOOPEN);
					MessageBox(strError);
					return hResult;
				}
				m_pDesignerImpl->OpenDesigner(ID_VERB_STANDALONE == lVerb ? (OLE_HANDLE)hwndParent : (OLE_HANDLE)GetParentWindow(), (LPDISPATCH)this, (ID_VERB_STANDALONE == lVerb ? VARIANT_TRUE : VARIANT_FALSE));
				return NOERROR;
			}
			break;

		case CTLIVERB_PROPERTIES:
		case OLEIVERB_PROPERTIES:
			try
			{
				// show the frame ourselves if the hose can't.
				if (m_pControlSite) 
				{
					try
					{
						hr = m_pControlSite->ShowPropertyFrame();
						if (hr != E_NOTIMPL)
							return hr;
					}
					CATCH
					{
						DDString strError;
						strError.LoadString(IDS_ERR_DESIGNERFAILEDTOOPEN);
						MessageBox(strError);
						REPORTEXCEPTION(__FILE__, __LINE__)
						return E_FAIL;
					}
				}
				IUnknown* pUnk = (IUnknown*)(IOleObject*)this;
				MAKE_WIDEPTR_FROMTCHAR(pwsz, NAMEOFOBJECT(m_objectIndex));
				ModalDialog(TRUE);
				hr = OleCreatePropertyFrame(GetActiveWindow(),
											GetSystemMetrics(SM_CXSCREEN) / 2,
											GetSystemMetrics(SM_CYSCREEN) / 2,
											pwsz,
											1,
											&pUnk,
											PROPPAGECOUNT(objectDef),
											(LPCLSID)*(PROPPAGES(objectDef)),
											g_lcidLocale,
											NULL, 
											NULL);
				ModalDialog(FALSE);
				return hr;
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
			return S_OK;

	default:
        // if it's a derived-control defined verb, pass it on to them
        if (lVerb > 0) 
		{
            hr = OnVerb(lVerb);
            if (OLEOBJ_S_INVALIDVERB == hr) 
			{
                // unrecognised verb -- just do the primary verb and activate
                hr = InPlaceActivate(OLEIVERB_PRIMARY);
                return (FAILED(hr)) ? hr : OLEOBJ_S_INVALIDVERB;
            } 
			else
				return hr;
        } 
		else 
		{
			TRACE(1, "Unrecognized negative verb in DoVerb().\n");
			return E_NOTIMPL;
        }
        break;
	}
	// dead code
}
STDMETHODIMP CBar::EnumVerbs( IEnumOLEVERB **ppEnumOleVerb)
{
	TRACE(1, "IOleObject::EnumVerbs\n");
	int cVerbs=sizeof(CBarVerbs)/sizeof(OLEVERB);
	if (0 == cVerbs)
        return OLEOBJ_E_NOVERBS;

	OLEVERB *rgVerbs;
	if (!(rgVerbs = (OLEVERB *)HeapAlloc(g_hHeap, 0, cVerbs * sizeof(OLEVERB))))
		return E_OUTOFMEMORY;

	memcpy(rgVerbs,CBarVerbs,cVerbs * sizeof(OLEVERB));
	*ppEnumOleVerb = (IEnumOLEVERB *)(IEnumX *) new CEnumX(IID_IEnumOLEVERB,
														   cVerbs, 
														   sizeof(OLEVERB), 
														   rgVerbs, 
														   CopyOLEVERB);

	if (!*ppEnumOleVerb)
        return E_OUTOFMEMORY;
	return S_OK;
}
STDMETHODIMP CBar::Update()
{
	return S_OK;
}
STDMETHODIMP CBar::IsUpToDate()
{
	return S_OK;
}
STDMETHODIMP CBar::GetUserClassID( CLSID *pClsid)
{
#ifdef DEF_PERSIST
	return GetClassID(pClsid); // use IPersist impl
#else
	*pClsid=*(((UNKNOWNOBJECTDESC *)(g_objectTable[0].objectDefPtr))->rclsid);
#endif
	return NOERROR;
}
STDMETHODIMP CBar::GetUserType( DWORD dwFormOfType, LPOLESTR *pszUserType)
{
	*pszUserType=OLESTRFROMANSI(((UNKNOWNOBJECTDESC *)(g_objectTable[0].objectDefPtr))->pszObjectName);
	return (*pszUserType) ? S_OK : E_OUTOFMEMORY;
}
STDMETHODIMP CBar::SetExtent( DWORD dwDrawAspect,  SIZEL *psizel)
{
	TRACE(1, "IOleObject::SetExtent\n");
	if (!(dwDrawAspect&DVASPECT_CONTENT))
		return DV_E_DVASPECT;

	SIZEL pixsizel;
	HiMetricToPixel(psizel,&pixsizel);

	BOOL fAcceptSizing=OnSetExtent(&pixsizel);
	if (!fAcceptSizing)
		return E_FAIL;

	m_Size=pixsizel;

	RECT rect;
	if (!m_pInPlaceSiteWndless) 
	{
		if (m_fInPlaceActive) 
		{
			// theoretically, one should not need to call OnPosRectChange
			// here, but there appear to be a few host related issues that
            // will make us keep it here.  we won't, however, both with
            // windowless ole controls, since they are all new hosts who
            // should know better
			GetWindowRect(m_hWnd, &rect);
			MapWindowPoints(0, m_hWndParent, (LPPOINT)&rect, 2);
			rect.right = rect.left + m_Size.cx;
			rect.bottom = rect.top + m_Size.cy;
			
			if (m_pInPlaceSite)
				m_pInPlaceSite->OnPosRectChange(&rect);

			if (m_hWnd) 
			{
				// just go and resize
                if (m_hWndReflect)
					SetWindowPos(m_hWndReflect, 0, 0, 0, m_Size.cx, m_Size.cy,
                                     SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                SetWindowPos(m_hWnd, 0, 0, 0, m_Size.cx, m_Size.cy,
	                         SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
			} 
		}
		else if (m_hWnd) 
			SetWindowPos(m_hWnd, NULL, 0, 0, m_Size.cx, m_Size.cy, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
		else 
			ViewChanged();
	} 
	else
		if (m_pInPlaceSite) 
			m_pInPlaceSite->OnPosRectChange(&rect);
	return S_OK;
}
STDMETHODIMP CBar::GetExtent( DWORD dwDrawAspect,  SIZEL *psizel)
{
	TRACE(1, "IOleObject::GetExtent\n");
	PixelToHiMetric((const SIZEL *)&m_Size, psizel);
	return S_OK;
}
STDMETHODIMP CBar::Advise( IAdviseSink *pAdvSink,  DWORD *pdwConnection)
{
	TRACE(1, "IOleObject::Advise\n");
    HRESULT hr;
    if (!m_pOleAdviseHolder) 
	{
        hr = CreateOleAdviseHolder(&m_pOleAdviseHolder);
        RETURN_ON_FAILURE(hr);
    }
	return m_pOleAdviseHolder->Advise(pAdvSink, pdwConnection);
}
STDMETHODIMP CBar::Unadvise(DWORD dwConnection)
{
	TRACE(1, "IOleObject::Unadvise\n");
    if (!m_pOleAdviseHolder) 
	{
		TRACE(1, "IOleObject::Unadvise call without Advise");
		return CONNECT_E_NOCONNECTION;
    }
    return m_pOleAdviseHolder->Unadvise(dwConnection);
}
STDMETHODIMP CBar::EnumAdvise(IEnumSTATDATA **ppenumAdvise)
{
	TRACE(1, "IOleObject::EnumAdvise\n");
	if (!m_pOleAdviseHolder) 
	{
        TRACE(1, "IOleObject::EnumAdvise call without Advise");
        *ppenumAdvise = NULL;
        return E_FAIL;
    }
    return m_pOleAdviseHolder->EnumAdvise(ppenumAdvise);
}
STDMETHODIMP CBar::GetMiscStatus( DWORD dwAspect,  DWORD *pdwStatus)
{
	TRACE(1, "IOleObject::GetMiscStatus\n");
	if (dwAspect!=DVASPECT_CONTENT)
		return DV_E_DVASPECT;
	*pdwStatus=((CONTROLOBJECTDESC*)objectDef)->dwOleMiscFlags;
	return NOERROR;
}
STDMETHODIMP CBar::SetColorScheme( LOGPALETTE *pLogpal)
{
	return S_OK;
}

//{IOleObject Support Code}
static CBar *s_pLastControlCreated;
LPTSTR szWndClass_CBar=DD_WNDCLASS("DynaBarCls");

void CBar::ViewChanged()
{
    if (m_pViewAdviseSink) 
	{
        m_pViewAdviseSink->OnViewChange(DVASPECT_CONTENT, -1);
        if (m_fViewAdviseOnlyOnce)
            SetAdvise(DVASPECT_CONTENT, 0, NULL);
    }
}

BOOL CBar::RegisterClassData(void)
{
	TRACE(1, "RegisterClassData\n");

	WNDCLASS wndclass;
	memset(&wndclass, 0, sizeof(WNDCLASS));
    wndclass.style          = CS_VREDRAW|CS_HREDRAW| CS_DBLCLKS;
    wndclass.lpfnWndProc    = CBar::ControlWindowProc;
    wndclass.hInstance      = g_hInstance;
    wndclass.hCursor        = LoadCursor(NULL, IDC_ARROW);
    wndclass.hbrBackground  = NULL;
    wndclass.lpszClassName  = szWndClass_CBar;
	((CONTROLOBJECTDESC *)objectDef)->szWndClass=szWndClass_CBar;
    return RegisterClass(&wndclass);
}

LRESULT CALLBACK CBar::ControlWindowProc(HWND hwnd,UINT msg,WPARAM wParam,LPARAM lParam)
{
    
	CBar *pCtl = ControlFromHwnd(hwnd);
    HRESULT hr;
    LRESULT lResult;
    DWORD   dwCookie;

    // if the value isn't a positive value, then it's in some special
    // state [creation or destruction]  this is safe because under win32,
    // the upper 2GB of an address space aren't available.
    //
    if ((LONG)pCtl == 0) 
	{
        pCtl = s_pLastControlCreated;
        SetWindowLong(hwnd, GWL_USERDATA, (LONG)pCtl);
        pCtl->m_hWnd = hwnd;
    }
	else if ((ULONG)pCtl == 0xffffffff) 
	{
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    // message preprocessing
    //
    if (pCtl->m_pSimpleFrameSite) 
	{
        hr = pCtl->m_pSimpleFrameSite->PreMessageFilter(hwnd, msg, wParam, lParam, &lResult, &dwCookie);
        if (hr == S_FALSE) return lResult;
    }

    // for certain messages, do not call the user window proc. instead,
    // we have something else we'd like to do.
    //

	pCtl->AddRef();
    lResult = pCtl->WindowProc(msg, wParam, lParam);

    // message postprocessing
    switch (msg) 
	{
	case WM_DESTROY:
		TRACE(1, _T("WM_DESTROY\n"));
		break;

	case WM_NCDESTROY:
        // after this point, the window doesn't exist any more
        pCtl->m_hWnd = NULL;
        break;

    case WM_MOUSEMOVE:
    case WM_NCMOUSEMOVE:
		if (pCtl->m_pActiveTool)
		{
			pCtl->ToolNotification(pCtl->m_pActiveTool, TNF_ACTIVETOOLCHECK);
			pCtl->m_pActiveTool = NULL;
		}
        break;

    case WM_SETFOCUS:
		TRACE(1, "BARCTL SETFOCUS\n");
        // give the control site focus notification

		if (!IsWindowVisible(pCtl->m_hWnd))
			break;

		if (pCtl->m_fInPlaceActive)
		{
			IOleObject* pOleObject;
			HRESULT hResult = pCtl->QueryInterface(IID_IOleObject, (void**)&pOleObject);
			if (SUCCEEDED(hResult))
			{
				CRect m_rcPos;
				pOleObject->DoVerb(OLEIVERB_UIACTIVATE, NULL, pCtl->m_pClientSite, 0, hwnd, &m_rcPos);
				pOleObject->Release();
			}

			if (pCtl->m_pControlSite)
			{
				TRACE(1, "BARCTL SETFOCUS - 2\n");
				pCtl->m_pControlSite->OnFocus(TRUE);
			}
		}

		if (pCtl->m_fInPlaceActive && pCtl->m_pControlSite)
			pCtl->m_pControlSite->OnFocus(TRUE);
        break;

	case WM_KILLFOCUS:
		TRACE(1, "BARCTL KILLFOCUS\n");
        // give the control site focus notification
        if (pCtl->m_fInPlaceActive && pCtl->m_pControlSite)
		{
			TRACE(1, "BARCTL KILLFOCUS - 2\n");
            pCtl->m_pControlSite->OnFocus(FALSE);
		}
        break;

	case WM_SIZE:
        // a change in size is a change in view
        if (!pCtl->m_fCreatingWindow)
		{
			// Why I did this I don't know right now
            pCtl->ViewChanged();
        }
		break;
    }

    // lastly, simple frame postmessage processing
    if (pCtl->m_pSimpleFrameSite)
        pCtl->m_pSimpleFrameSite->PostMessageFilter(hwnd, msg, wParam, lParam, &lResult, dwCookie);

	pCtl->Release();
    return lResult;

}

const BYTE g_rgcbDataTypeSize[] = {
    0,                      // VT_EMPTY= 0,
    0,                      // VT_NULL= 1,
    sizeof(short),          // VT_I2= 2,
    sizeof(long),           // VT_I4 = 3,
    sizeof(float),          // VT_R4  = 4,
    sizeof(double),         // VT_R8= 5,
    sizeof(CURRENCY),       // VT_CY= 6,
    sizeof(DATE),           // VT_DATE = 7,
    sizeof(BSTR),           // VT_BSTR = 8,
    sizeof(IDispatch *),    // VT_DISPATCH    = 9,
    sizeof(SCODE),          // VT_ERROR    = 10,
    sizeof(VARIANT_BOOL),   // VT_BOOL    = 11,
    sizeof(VARIANT),        // VT_VARIANT= 12,
    sizeof(IUnknown *),     // VT_UNKNOWN= 13,
};

BOOL CBar::GetAmbientProperty(DISPID  dispid, VARTYPE vt, void* pData)
{
    DISPPARAMS dispparams;
    VARIANT v, v2;
    HRESULT hr;
    v.vt = VT_EMPTY;
    v.lVal = 0;

    v2.vt = VT_EMPTY;
    v2.lVal = 0;
    
	// get a pointer to the source of ambient properties.
    if (!m_pDispAmbient) 
	{
        if (m_pClientSite)
            m_pClientSite->QueryInterface(IID_IDispatch, (void **)&m_pDispAmbient);
        if (!m_pDispAmbient)
            return FALSE;
    }

    // now go and get the property into a variant.
    memset(&dispparams, 0, sizeof(DISPPARAMS));
    hr = m_pDispAmbient->Invoke(dispid, 
								IID_NULL, 
								0, 
								DISPATCH_PROPERTYGET, 
								&dispparams,
                                &v, 
								0, 
								0);
    if (FAILED(hr)) 
		return FALSE;
    
	// we've got the variant, so now go an coerce it to the type that the user
    // wants.  if the types are the same, then this will copy the stuff to
    // do appropriate ref counting ...

    hr = VariantChangeType(&v2, &v, 0, vt);
    if (FAILED(hr)) 
	{
        VariantClear(&v);
        return FALSE;
    }
    
	// copy the data to where the user wants it
    CopyMemory(pData, &(v2.lVal), g_rgcbDataTypeSize[vt]);
    VariantClear(&v);
    return TRUE;
}

BOOL CBar::BeforeCreateWindow(DWORD *pdwWindowStyle,DWORD *pdwExWindowStyle,LPTSTR  pszWindowTitle)
{
	*pdwWindowStyle |= WS_VISIBLE;
    return TRUE;
}

UINT WINAPI StatusBarThread(LPVOID pParameter)
{
//	g_cLocks++;
	SYSTEMTIME theSystemTime;
	BOOL bRefresh = FALSE;
	CBar* pBar = (CBar*)pParameter;
	try
	{
		while (!pBar->m_bStopStatusBarThread)
		{
			if (pBar->m_pStatusBand)
			{
				CTool* pTool;
				CTools* pTools = pBar->m_pStatusBand->m_pTools;
				try
				{
					int nCount = pTools->GetVisibleToolCount();
					for (int nTool = 0; nTool < nCount; nTool++)
					{
						pTool = pTools->GetVisibleTool(nTool);
						if (pTool && ddTTLabel == pTool->tpV1.m_ttTools && (ddLSDate == pTool->tpV1.m_lsLabelStyle || ddLSTime == pTool->tpV1.m_lsLabelStyle))
						{
							GetLocalTime(&theSystemTime);
							if (ddLSTime == pTool->tpV1.m_lsLabelStyle)
							{
								if (theSystemTime.wHour != pTool->m_theSystemTime.wHour || theSystemTime.wMinute != pTool->m_theSystemTime.wMinute)
								{
									pTool->m_theSystemTime = theSystemTime;
									bRefresh = TRUE;
								}
							}
							else
							{
								if (theSystemTime.wYear != pTool->m_theSystemTime.wYear || theSystemTime.wMonth != pTool->m_theSystemTime.wMonth || theSystemTime.wDay != pTool->m_theSystemTime.wDay)
								{
									pTool->m_theSystemTime = theSystemTime;
									bRefresh = TRUE;
								}
							}
						}
					}
					if (bRefresh)
					{
						pBar->m_pStatusBand->Refresh();
						bRefresh = FALSE;
					}
				}
				catch (...)
				{
					assert(FALSE);
				}
			}
//			if (g_cLocks < 2)
//				break;

			WaitForSingleObject(pBar->m_hStatusBarEvent, 300);

//			if (g_cLocks < 2)
//				break;
		}
//		g_cLocks--;
	}
	catch (...)
	{
//		g_cLocks--;
		assert(FALSE);
	}
	_endthreadex(0);
	return -1;
}

BOOL CBar::AfterCreateWindow()
{
	CreateDockWindows();
	if (AmbientUserMode())
	{
		//
		// We need to set a map for the Accelator hook
		//

		HWND hWndTmp = GetDockWindow();
		LPVOID pTmp;
		if (!GetGlobals().m_pmapAccelator->Lookup((LPVOID)hWndTmp, (void*&)pTmp))
			GetGlobals().m_pmapAccelator->SetAt(hWndTmp, (LPVOID)this);

		if (m_bIE && m_bstrDataPath && *m_bstrDataPath)
			PostMessage(m_hWnd, WM_COMMAND, MAKELONG(eStartDownload,0), 0);
	}
	return TRUE;
}

void CBar::BeforeDestroyWindow()
{
	// used for subclassed ones
	if (IsWindow(m_hWnd))
	{
		KillTimer(m_hWnd, eFlyByTimer);
		HideToolTips(TRUE);
	}
	if (AmbientUserMode() && eClientArea != m_eAppType)
		StopStatusBandThread();
	BOOL bResult = GetGlobals().m_pmapAccelator->RemoveKey(GetDockWindow());
	m_theFormsAndControls.ShutDown();
	if (m_pRelocate)
	{
		delete m_pRelocate;
		m_pRelocate = NULL;
	}
	if (m_pDockMgr)
		m_pDockMgr->DestroyDockAreasWnds();
}

LRESULT CALLBACK CBar::ReflectWindowProc(HWND hwnd,UINT msg, WPARAM  wParam,LPARAM  lParam)
{
    CBar *pCtl;
    switch (msg) 
	{
        case WM_COMMAND:
        case WM_NOTIFY:
        case WM_CTLCOLORBTN:
        case WM_CTLCOLORDLG:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX:
        case WM_CTLCOLORMSGBOX:
        case WM_CTLCOLORSCROLLBAR:
        case WM_CTLCOLORSTATIC:
        case WM_DRAWITEM:
        case WM_MEASUREITEM:
        case WM_DELETEITEM:
        case WM_VKEYTOITEM:
        case WM_CHARTOITEM:
        case WM_COMPAREITEM:
        case WM_HSCROLL:
        case WM_VSCROLL:
        case WM_PARENTNOTIFY:
        case WM_SETFOCUS:
        case WM_SIZE:
            pCtl = (CBar *)GetWindowLong(hwnd, GWL_USERDATA);
            if (pCtl)
                return SendMessage(pCtl->m_hWnd, OCM__BASE + msg, wParam, lParam);
            break;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}
//
//		if (!AmbientUserMode())
//			GetInPlaceSite()->OnUIActivate(); 
HRESULT CBar::InPlaceActivate(LONG lVerb)
{
    BOOL f;
    SIZEL sizel;
    IOleInPlaceSiteEx *pIPSEx = NULL;
    HRESULT hr;
    BOOL    fNoRedraw = FALSE;

	TRACE1(1, "InPlaceActivate %d\n",lVerb);

    if (!m_pClientSite) // check if client site exists
        return S_OK;
    // get an InPlace site pointer.
    if (!GetInPlaceSite()) 
	{
        // if they want windowless support, then we want IOleInPlaceSiteWindowless
        if (((CONTROLOBJECTDESC*)objectDef)->fWindowless)
            m_pClientSite->QueryInterface(IID_IOleInPlaceSiteWindowless, (void **)&m_pInPlaceSiteWndless);
        // if we're not able to do windowless siting, then we'll just get an
        // IOleInPlaceSite pointer.
        if (!m_pInPlaceSiteWndless) 
		{
            hr = m_pClientSite->QueryInterface(IID_IOleInPlaceSite, (void **)&m_pInPlaceSite);
            RETURN_ON_FAILURE(hr);
        }
    }

    // now, we want an IOleInPlaceSiteEx pointer for windowless and flicker free
    // activation.  if we're windowless, we've already got it, else we need to
    // try and get it
    //
    if (m_pInPlaceSiteWndless) 
	{
        pIPSEx = (IOleInPlaceSiteEx *)m_pInPlaceSiteWndless;
        pIPSEx->AddRef();
    }
	else
        m_pClientSite->QueryInterface(IID_IOleInPlaceSiteEx, (void **)&pIPSEx);

    // if we're not already active, go and do it.
    if (!m_fInPlaceActive) 
	{
        OLEINPLACEFRAMEINFO InPlaceFrameInfo;
        RECT rcPos, rcClip;

        // if we have a windowless site, see if we can go in-place windowless active
        hr = S_FALSE;
        if (m_pInPlaceSiteWndless) 
		{
            hr = m_pInPlaceSiteWndless->CanWindowlessActivate();
            CLEANUP_ON_FAILURE(hr);
		    // if they refused windowless, we'll try windowed
            if (S_OK != hr) 
			{
                RELEASE_OBJECT(m_pInPlaceSiteWndless);
                hr = m_pClientSite->QueryInterface(IID_IOleInPlaceSite, (void **)&m_pInPlaceSite);
                CLEANUP_ON_FAILURE(hr);
            }
        }
        // just try regular windowed in-place activation
        if (hr != S_OK) 
		{
            hr = m_pInPlaceSite->CanInPlaceActivate();
            if (hr != S_OK) 
			{
                hr = (FAILED(hr)) ? E_FAIL : hr;
                goto CleanUp;
            }
        }

        // if we are here, then we have permission to go in-place active.
        // now, announce our intentions to actually go ahead and do this.
        hr = (pIPSEx) ? pIPSEx->OnInPlaceActivateEx(&fNoRedraw, (m_pInPlaceSiteWndless) ? ACTIVATE_WINDOWLESS : 0)
                       : m_pInPlaceSite->OnInPlaceActivate();
        CLEANUP_ON_FAILURE(hr);

        // if we're here, we're ready to go in-place active.  we just need
        // to set up some flags, and then create the window [if we have one]
   
		m_fInPlaceActive = TRUE;

        // we need to get some information about our location in the parent
        // window, as well as some information about the parent
        //
        InPlaceFrameInfo.cb = sizeof(OLEINPLACEFRAMEINFO);
		HWND prevParent=m_hWndParent;
        hr = GetInPlaceSite()->GetWindow(&m_hWndParent);
        if (SUCCEEDED(hr))
            hr = GetInPlaceSite()->GetWindowContext(&m_pInPlaceFrame, &m_pInPlaceUIWindow, &rcPos, &rcClip, &InPlaceFrameInfo);
        CLEANUP_ON_FAILURE(hr);

        // make sure we'll display ourselves in the correct location with the correct size
        sizel.cx = rcPos.right - rcPos.left;
        sizel.cy = rcPos.bottom - rcPos.top;
        f = OnSetExtent(&sizel);
        if (f) m_Size = sizel;
#ifdef DEF_IOLEINPLACEOBJECT
        SetObjectRects(&rcPos, &rcClip);
#endif
        // finally, create our window if we have to!
        if (!m_pInPlaceSiteWndless) 
		{
			// SetInPlaceParent(m_hwndParent);
			ASSERT(!m_pInPlaceSiteWndless, "this should execute for windowed OLE controls only");
			if (prevParent!= m_hWndParent)
			{
				if (m_hWnd)
				    SetParent(GetOuterWindow(),m_hWndParent);
			}
			// end set inplace parent

            // create the window, and display it.  
            if (!CreateInPlaceWindow(rcPos.left, rcPos.top, fNoRedraw)) 
			{
                hr = E_FAIL;
                goto CleanUp;
            }
        }
    }

    RELEASE_OBJECT(pIPSEx);

    // if we're not inplace visible yet, do so now.
    if (!m_fInPlaceVisible)
        SetInPlaceVisible(TRUE);

    // if we weren't asked to UIActivate, then we're done.
    if (lVerb != OLEIVERB_PRIMARY && lVerb != OLEIVERB_UIACTIVATE)
        return S_OK;

    // if we're not already UI active, do sow now.
    if (!m_fUIActive) 
	{
        m_fUIActive = TRUE;
        // inform the container of our intent
		if (!AmbientUserMode())
			GetInPlaceSite()->OnUIActivate(); 
        // take the focus  [which is what UI Activation is all about !]
        SetFocus(TRUE);
        // set ourselves up in the host.
        m_pInPlaceFrame->SetActiveObject((IOleInPlaceActiveObject *)this, NULL);
        if (m_pInPlaceUIWindow)
            m_pInPlaceUIWindow->SetActiveObject((IOleInPlaceActiveObject *)this, NULL);
        // we have to explicitly say we don't wany any border space.
        m_pInPlaceFrame->SetBorderSpace(NULL);
        if (m_pInPlaceUIWindow)
            m_pInPlaceUIWindow->SetBorderSpace(NULL);
    }
    return S_OK;

  CleanUp:
    // something bad happened 
    if (pIPSEx) 
		pIPSEx->Release();
    m_fInPlaceActive = FALSE;
    return hr;
}

HWND CBar::CreateInPlaceWindow(int x, int y, BOOL fNoRedraw)
{
    BOOL    fVisible;
    HRESULT hr;
    DWORD   dwWindowStyle, dwExWindowStyle;
    TCHAR szWindowTitle[128];

	TRACE(1, "CreateInPlaceWindow\n");

    // if we've already got a window, do nothing.
    if (m_hWnd)
        return m_hWnd;
    // get the user to register the class if it's not already
    // been done.  we have to critical section this since more than one thread
    // can be trying to create this control
    //
#ifdef DEF_CRITSECTION
    EnterCriticalSection(&g_CriticalSection);
#endif

    if (((CONTROLOBJECTDESC *)objectDef)->szWndClass==NULL) 
	{
		RegisterClassData();
/*        if (!RegisterClassData()) 
		{
#ifdef DEF_CRITSECTION
            LeaveCriticalSection(&g_CriticalSection);
#endif
            return NULL;
        } 
*/
    }
#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif

    // let the user set up things like the window title, the
    // style, and anything else they feel interested in fiddling
    // with.
    dwWindowStyle = dwExWindowStyle = 0;
    szWindowTitle[0] = '\0';
    if (!BeforeCreateWindow(&dwWindowStyle, &dwExWindowStyle, szWindowTitle))
        return NULL;

    dwWindowStyle |= (WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN);

    // create window visible if parent hidden (common case)
    // otherwise, create hidden, then shown.  this is a little subtle, but
    // it makes sense eventually.
    //
    if (!m_hWndParent)
        m_hWndParent = GetParkingWindow();

    fVisible = IsWindowVisible(m_hWndParent);

    // This one kinda sucks -- if a control is subclassed, and we're in
    // a host that doesn't support Message Reflecting, we have to create
    // the user window in another window which will do all the reflecting.
    // VERY blech. [don't however, bother in design mode]
    //
    if (((CONTROLOBJECTDESC *)objectDef)->pfnSubClass
		&& (m_hWndParent != GetParkingWindow())) 
	{
        // determine if the host supports message reflecting.
        if (!m_fCheckedReflecting) 
		{
            VARIANT_BOOL f;
            hr = GetAmbientProperty(DISPID_AMBIENT_MESSAGEREFLECT, VT_BOOL, &f);
            if (FAILED(hr) || !f)
                m_fHostReflects = FALSE;
            m_fCheckedReflecting = TRUE;
        }
        // if the host doesn't support reflecting, then we have to create
        // an extra window around the control window, and then parent it
        // off that.
        if (!m_fHostReflects) 
		{
            ASSERT(m_hWndReflect == NULL, "Where'd this come from?");
            m_hWndReflect = CreateReflectWindow(!fVisible, m_hWndParent, x, y, &m_Size,ReflectWindowProc);
            if (!m_hWndReflect)
                return NULL;
            SetWindowLong(m_hWndReflect, GWL_USERDATA, (long)this);
            dwWindowStyle |= WS_VISIBLE;
        }
    } 
	else 
	{
        if (!fVisible)
            dwWindowStyle |= WS_VISIBLE;
    }

    // we have to mutex the entire create window process since we need to use
    // the s_pLastControlCreated to pass in the object pointer.  nothing too
    // serious
    //
#ifdef DEF_CRITSECTION
    EnterCriticalSection(&g_CriticalSection);
#endif

    s_pLastControlCreated = this;
    m_fCreatingWindow = TRUE;

    //
	// finally, go create the window, parenting it as appropriate.
    //
   
	m_hWnd = CreateWindowEx(dwExWindowStyle,
                            ((CONTROLOBJECTDESC *)objectDef)->szWndClass,
                            szWindowTitle,
                            dwWindowStyle,
                            (m_hWndReflect) ? 0 : x,
                            (m_hWndReflect) ? 0 : y,
							m_Size.cx,
							m_Size.cy,
                            (m_hWndReflect) ? m_hWndReflect : m_hWndParent,
                            NULL, g_hInstance, NULL);

    //
	// clean up some variables, and leave the critical section
    //
    
	m_fCreatingWindow = FALSE;
    s_pLastControlCreated = NULL;
#ifdef DEF_CRITSECTION
    LeaveCriticalSection(&g_CriticalSection);
#endif
    if (m_hWnd) 
	{
        // let the derived-control do something if they so desire
        if (!AfterCreateWindow()) 
		{
            BeforeDestroyWindow();
            SetWindowLong(m_hWnd, GWL_USERDATA, 0xFFFFFFFF);
            DestroyWindow(m_hWnd);
            m_hWnd = NULL;
            return m_hWnd;
        }

        // if we didn't create the window visible, show it now.
        if (fVisible)
		{
			if (GetParent(m_hWnd) != m_hWndParent)
				// SetWindowPos fails if parent hwnd is passed in so keep
				// this behaviour in those cases without ripping. 
				SetWindowPos(m_hWnd, m_hWndParent, 0, 0, 0, 0,
							 SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_SHOWWINDOW | ((fNoRedraw) ? SWP_NOREDRAW : 0));
		}
    }
    // finally, tell the host of this
    if (m_pClientSite)
        m_pClientSite->ShowObject();
    return m_hWnd;
}

BOOL CBar::SetFocus(BOOL fGrab)
{
    HRESULT hr;
    HWND    hwnd;
	
	TRACE(1, "SetFocus\n");

    // first thing to do is check out UI Activation state, and then set
    // focus [either with windows api, or via the host for windowless
    // controls]
    if (m_pInPlaceSiteWndless) 
	{
        if (!m_fUIActive && fGrab)
            if (FAILED(InPlaceActivate(OLEIVERB_UIACTIVATE))) return FALSE;

        hr = m_pInPlaceSiteWndless->SetFocus(fGrab);
        return (hr == S_OK) ? TRUE : FALSE;
    }
	else 
	{
	    // we've got a window.
        if (m_fInPlaceActive) 
		{
            hwnd= (fGrab) ? m_hWnd : m_hWndParent;
            if (!m_fUIActive && fGrab)
                return SUCCEEDED(InPlaceActivate(OLEIVERB_UIACTIVATE));
            else
			{
				// Notes:
				//    we do this because some controls host non-ole window hierarchies, like
				// the Netscape plugin ocx.  in such cases, the control may need to be UIActive
				// to function properly in the document, but cannot take the windows focus
				// away from one of its child windows.  such controls may override this method
				// and respond as appropriate.
				return (::SetFocus(hwnd) == hwnd);
			}
        }
		else
            return FALSE;
    }
}

void CBar::SetInPlaceVisible(BOOL fShow)
{
	BOOL fVisible;
    m_fInPlaceVisible = fShow;

	TRACE1(1, "SetInPlaceVisible %d\n",(int)fShow);

    // don't do anything if we don't have a window.  otherwise, set it
    if (m_hWnd) 
	{
        fVisible = ((GetWindowLong(GetOuterWindow(), GWL_STYLE) & WS_VISIBLE) != 0);
        if (fVisible && !fShow)
            ShowWindow(GetOuterWindow(), SW_HIDE);
        else if (!fVisible && fShow)
            ShowWindow(GetOuterWindow(), SW_SHOWNA);
    }
}

//
// OnSetExtent
//

BOOL CBar::OnSetExtent(SIZEL* sizel)
{
	switch (m_eAppType)
	{
	case eMDIForm:
		if (m_rcFrame.Size().cx > 0 && m_rcFrame.Size().cy > 0)
			*sizel = m_rcFrame.Size();
		break;

	case eSDIForm:
		if (IsWindow(m_hWnd))
		{
			if (m_pDockMgr)
				m_pDockMgr->ResizeDockableForms();
			SendMessage(m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
		}
		if (m_rcFrame.Size().cx > 0 && m_rcFrame.Size().cy > 0)
			*sizel = m_rcFrame.Size();
		break;
	}
	return TRUE;
}

//
// ModalDialog
//

void CBar::ModalDialog(BOOL fShow)
{
    if (m_pInPlaceFrame)
        m_pInPlaceFrame->EnableModeless(!fShow);
}

//
// ConvertPtToHiMetric
//

static void ConvertPtToHiMetric(LPARAM lParam, SIZEL& sizeHi)
{
	SIZEL sizePixel;
	sizePixel.cx = LOWORD(lParam);
	sizePixel.cy = HIWORD(lParam);
	PixelToHiMetric(&sizePixel, &sizeHi);
}

//
// OCXMouseState
//

short OCXMouseState(WPARAM wParam)
{
    BOOL bLeft = (wParam & MK_LBUTTON);
    BOOL bRight  = (wParam & MK_RBUTTON);
    BOOL bMiddle   = (wParam & MK_MBUTTON);
    return (short)(bLeft + (bRight << 1) + (bMiddle << 2));
}

//
// WindowProc
//

extern BOOL g_fMachineHasLicense;
extern BOOL g_bDemoDlgDisplayed;

LRESULT CBar::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		POINT pt;
		DWORD dwCurTime;
		static DWORD dwInitTime = 0;
		static POINT ptInit = {0, 0};

		if (GetGlobals().WM_SETTOOLFOCUS == nMsg)
		{
			POINT pt = {-1, -1};
			CTool* pTool = (CTool*)lParam;
			if (0 == wParam)
				pTool->HandleCombobox(0, pt);
			else
				pTool->HandleEdit(0, pt);
			return 0;
		}

		if (GetGlobals().WM_KILLWINDOW == nMsg)
		{
			if (IsWindow((HWND)wParam))
				DestroyWindow((HWND)wParam);
			return 0;
		}
		
		if (GetGlobals().WM_REFRESHMENUBAND == nMsg)
		{
			CBand* pBand = GetMenuBand();
			if (pBand && VARIANT_TRUE == pBand->bpV1.m_vbVisible)
				RecalcLayout();
			return 0;
		}

		if (GetGlobals().WM_SIZEPARENT == nMsg)
		{
			InternalRecalcLayout();
			return 0;
		}

		if (GetGlobals().WM_RECALCLAYOUT == nMsg) 
		{
			RecalcLayout();
			return 0;
		}

		if (GetGlobals().WM_ACTIVEBARTEXTCHANGE == nMsg)
		{
			CTool* pTool = reinterpret_cast<CTool*>(lParam);
			if (pTool)
			{
				FireTextChange(reinterpret_cast<Tool*>(pTool));
				pTool->Release();
			}
			return 0;
		}

		if (GetGlobals().WM_ACTIVEBARCOMBOSELCHANGE == nMsg)
		{
			CTool* pTool = reinterpret_cast<CTool*>(lParam);
			if (pTool)
			{
				FireComboSelChange(reinterpret_cast<Tool*>(pTool));
				TRACE(5, "FireComboSelChange\n");
				pTool->Release();
			}
			return 0;
		}

		if (GetGlobals().WM_ACTIVEBARCLICK == nMsg)
		{
			CTool* pTool = reinterpret_cast<CTool*>(lParam);
			if (pTool)
			{
				FireToolClick(reinterpret_cast<Tool*>(pTool));
				pTool->Release();
			}
			return 0;
		}

		switch(nMsg)
		{
		case WM_LBUTTONDBLCLK:
			FireDblClick();
			return 0;

		case WM_MBUTTONDBLCLK:
			FireDblClick();
			return 0;
		
		case WM_RBUTTONDBLCLK:
			FireDblClick();
			return 0;

		case WM_TIMER:
			OnTimer(wParam);
			return 0;

		case WM_MOUSEMOVE:
			if (AmbientUserMode() && !(m_bCustomization || m_bWhatsThisHelp))
			{
				SIZE size;
				size.cx = LOWORD(lParam);
				size.cy = HIWORD(lParam);
				PixelToTwips(&size, &size);
				FireMouseMove(OCXMouseState(wParam), OCXShiftState(), (float)size.cx, (float)size.cy);
			}
			break;

		case WM_LBUTTONUP:
			if (AmbientUserMode() && !(m_bCustomization || m_bWhatsThisHelp))
			{
				SIZE size;
				size.cx = LOWORD(lParam);
				size.cy = HIWORD(lParam);
				PixelToTwips(&size, &size);
				FireMouseUp(vbLeftButton, OCXShiftState(), (float)size.cx, (float)size.cy);
				GetCursorPos(&pt);
				dwCurTime = GetTickCount();
				if (abs(pt.x - ptInit.x) < GetSystemMetrics(SM_CXDOUBLECLK) && abs(pt.y - ptInit.y) < GetSystemMetrics(SM_CYDOUBLECLK))
				{
					// check time elapsed
					if ((GetTickCount() - dwInitTime) > GetDoubleClickTime())
						FireClick();
				}
				else
					FireClick();
				ptInit = pt;
				dwInitTime = dwCurTime;
			}
			break;

		case WM_LBUTTONDOWN:
			if (m_bCustomization || m_bWhatsThisHelp)
			{
				ClearCustomization();
				return 0;
			}
			if (AmbientUserMode())
			{
				if (m_pActiveBand)
				{
					m_pActiveBand = NULL;
					Refresh();
				}
				SIZE size;
				size.cx = LOWORD(lParam);
				size.cy = HIWORD(lParam);
				PixelToTwips(&size, &size);
				FireMouseDown(vbLeftButton, OCXShiftState(), (float)size.cx, (float)size.cy);
			}
			break;

		case WM_RBUTTONUP:
			if (AmbientUserMode() && !(m_bCustomization || m_bWhatsThisHelp))
			{
				SIZE size;
				size.cx = LOWORD(lParam);
				size.cy = HIWORD(lParam);
				PixelToTwips(&size, &size);
				FireMouseUp(vbRightButton, OCXShiftState(), (float)size.cx, (float)size.cy);
				GetCursorPos(&pt);
				dwCurTime = GetTickCount();
				if (abs(pt.x - ptInit.x) < GetSystemMetrics(SM_CXDOUBLECLK) && abs(pt.y - ptInit.y) < GetSystemMetrics(SM_CYDOUBLECLK))
				{
					// check time elapsed
					if ((GetTickCount() - dwInitTime) > GetDoubleClickTime())
						FireClick();
				}
				else
					FireClick();
				ptInit = pt;
				dwInitTime = dwCurTime;
			}
			break;

		case WM_RBUTTONDOWN:
			if (m_bCustomization || m_bWhatsThisHelp)
			{
				ClearCustomization();
				return 0;
			}
			if (AmbientUserMode())
			{
				if (m_pActiveBand)
				{
					m_pActiveBand = NULL;
					Refresh();
				}
				SIZE size;
				size.cx = LOWORD(lParam);
				size.cy = HIWORD(lParam);
				PixelToTwips(&size, &size);
				FireMouseDown(vbRightButton, OCXShiftState(), (float)size.cx, (float)size.cy);
			}
			break;

		case WM_MBUTTONDOWN:
			if (m_bCustomization || m_bWhatsThisHelp)
			{
				ClearCustomization();
				return 0;
			}
			if (AmbientUserMode())
			{
				if (m_pActiveBand)
				{
					m_pActiveBand = NULL;
					Refresh();
				}
				SIZE size;
				size.cx = LOWORD(lParam);
				size.cy = HIWORD(lParam);
				PixelToTwips(&size, &size);
				FireMouseDown(vbMiddleButton, OCXShiftState(), (float)size.cx, (float)size.cy);
			}
			break;

		case WM_MBUTTONUP:
			if (AmbientUserMode() && !(m_bCustomization || m_bWhatsThisHelp))
			{
				SIZE size;
				size.cx = LOWORD(lParam);
				size.cy = HIWORD(lParam);
				PixelToTwips(&size, &size);
				FireMouseUp(vbMiddleButton, OCXShiftState(), (float)size.cx, (float)size.cy);
				GetCursorPos(&pt);
				dwCurTime = GetTickCount();
				if (abs(pt.x - ptInit.x) < GetSystemMetrics(SM_CXDOUBLECLK) && abs(pt.y - ptInit.y) < GetSystemMetrics(SM_CYDOUBLECLK))
				{
					// check time elapsed
					if ((GetTickCount() - dwInitTime) > GetDoubleClickTime())
						FireClick();
				}
				else
					FireClick();
				ptInit = pt;
				dwInitTime = dwCurTime;
			}
			break;

		case WM_KEYDOWN:
			if (eMDIForm == m_eAppType)
			{
				::SetFocus(NULL);
				SetFocus(FALSE);
				break;
			}
			FireKeyDown(wParam, OCXShiftState());
			SetFocus(TRUE);
			break;
		
		case WM_KEYUP:
			if (eMDIForm == m_eAppType)
			{
				::SetFocus(NULL);
				SetFocus(FALSE);
				break;
			}
			FireKeyUp(wParam, OCXShiftState());
			SetFocus(TRUE);
			break;

		case WM_WINDOWPOSCHANGING:
			switch (m_eAppType)
			{
			case eMDIForm:
				if (AmbientUserMode())
				{
					LPWINDOWPOS pWP = (LPWINDOWPOS)lParam;
					pWP->x = 0;
					pWP->y = 0;
					pWP->cx = 0;
					pWP->cy = 0;
					return 0;
				}
				else
					PostMessage(m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
				break;
			}
			break;
		
		case WM_SHOWWINDOW:
			switch (m_eAppType)
			{
			case eSDIForm:
			case eClientArea:
				if ((BOOL)wParam)
					PostMessage(m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
				break;
			}
			break;

		case WM_SYSCOLORCHANGE:
		case WM_SETTINGCHANGE:
			OnSysColorChanged();
			break;

		case WM_ERASEBKGND:
			return 0;

		case WM_PARENTNOTIFY:
			{
				switch (LOWORD(wParam))
				{
				case WM_CREATE:
				case WM_DESTROY:
					{
						TRACE(5, "WM_PARENTNOTIFY - 1\n");
					}
					break;

				case WM_LBUTTONDOWN:
				case WM_MBUTTONDOWN:
				case WM_RBUTTONDOWN:
					{
						TRACE(5, "WM_PARENTNOTIFY - 2\n");
					}
					break;
				}
			}
			break;

		case WM_COMMAND:
			switch(LOWORD(wParam))
			{
			case CTool::eEdit:
			case CTool::eCombobox:
				//
				// Message reflection for the built in Combobox and Edit Tool.
				//
				switch (HIWORD(wParam))
				{
				case EN_SETFOCUS:
				case EN_KILLFOCUS:
				case EN_CHANGE:
				case CBN_CLOSEUP:
				case CBN_KILLFOCUS:
				case CBN_DROPDOWN:
//				case CBN_SELCHANGE:
				case CBN_EDITCHANGE:
					::SendMessage((HWND)lParam, WM_COMMAND, wParam, lParam);
					break;
				}
				break;

			case eStartDownload:
				{
					// start async download
					if (NULL == m_downloadItem.m_pClientSite)
						m_downloadItem.Init(this, m_pClientSite, &DownloadCallback);

					m_downloadItem.CancelDownload();
					if (FAILED(m_downloadItem.StartDownload(m_bstrDataPath)))
						DataReady(&m_downloadItem, FALSE);
				}
				break;

			case eDownloadError:
				break;
			}
			break;

		case WM_PAINT:
			{
				// call the user's OnDraw routine.
				PAINTSTRUCT ps;
				HDC         hDC;
				// if we're given an HDC, then use it
				if (!wParam)
					hDC = BeginPaint(m_hWnd, &ps);
				else
					hDC = (HDC)wParam;

				RECT rc;
				GetClientRect(m_hWnd, &rc);
				
				PaintVBBackground(m_hWnd, hDC, m_eAppType == eMDIForm);

				OnDraw(DVASPECT_CONTENT, hDC, (RECTL*)&rc, NULL, NULL, TRUE);

				if (!g_fMachineHasLicense) 
				{
					if (!g_bDemoDlgDisplayed)
						DemoDlg();
				}

				if (!wParam)
					EndPaint(m_hWnd, &ps);
			}
			return 0;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
	return DefWindowProc(m_hWnd, nMsg, wParam, lParam);
}
//{IOleObject Support Code}

// IOleWindow members

STDMETHODIMP CBar::GetWindow( HWND *phwnd)
{
    if (m_pInPlaceSiteWndless)
        return E_FAIL;
    *phwnd = GetOuterWindow();
    return (*phwnd) ? S_OK : E_UNEXPECTED;
}
STDMETHODIMP CBar::ContextSensitiveHelp( BOOL fEnterMode)
{
	return E_NOTIMPL;
}

// IOleInPlaceObject members

STDMETHODIMP CBar::InPlaceDeactivate()
{
	TRACE(1, "InplaceDeactivate\n");
    if (!m_fInPlaceActive)
        return S_OK;
    // transition from UIActive back to active
    if (m_fUIActive)
        UIDeactivate();
    m_fInPlaceActive = FALSE;
    m_fInPlaceVisible = FALSE;
    // if we have a window, tell it to go away.
    if (IsWindow(m_hWnd)) 
	{
        ASSERT(!m_pInPlaceSiteWndless, "internal state really messed up");
        // so our window proc doesn't crash.
        BeforeDestroyWindow();
        SetWindowLong(m_hWnd, GWL_USERDATA, 0xFFFFFFFF);
		if (IsWindow(m_hWnd))
			DestroyWindow(m_hWnd);
        m_hWnd = NULL;

        if (m_hWndReflect) 
		{
            SetWindowLong(m_hWndReflect, GWL_USERDATA, 0);
            DestroyWindow(m_hWndReflect);
            m_hWndReflect = NULL;
        }
    }
    RELEASE_OBJECT(m_pInPlaceFrame);
    RELEASE_OBJECT(m_pInPlaceUIWindow);
	if (GetInPlaceSite())
		GetInPlaceSite()->OnInPlaceDeactivate();
	RELEASE_OBJECT(m_pInPlaceSite);
    return S_OK;
}
STDMETHODIMP CBar::UIDeactivate()
{
	if (m_pDesignerImpl && !m_bShutdownLock)
	{
		DoCustomization(FALSE);
		m_pDesignerImpl->UIDeactivateCloseDesigner();
	}
	if (!m_fUIActive)
		return S_OK;
    m_fUIActive = FALSE;
    // notify frame windows, if appropriate, that we're no longer ui-active.
    if (m_pInPlaceUIWindow) 
		m_pInPlaceUIWindow->SetActiveObject(NULL, NULL);
	if (m_pInPlaceFrame)
		m_pInPlaceFrame->SetActiveObject(NULL, NULL);
    // we don't need to explicitly release the focus here since somebody else grabbing it
	if (GetInPlaceSite())
		GetInPlaceSite()->OnUIDeactivate(FALSE);
    return S_OK;
}
STDMETHODIMP CBar::SetObjectRects( LPCRECT lprcPosRect, LPCRECT lprcClipRect)
{
    BOOL fRemoveWindowRgn;
    // move our window to the new location and handle clipping. not applicable
    // for windowless controls, since the container will be responsible for all
    // clipping.
    if (m_hWnd) 
	{
        fRemoveWindowRgn = m_fUsingWindowRgn;
        if (lprcClipRect) 
		{
            // the container wants us to clip, so figure out if we really
            // need to
            //
            RECT rcIXect;
            if ( IntersectRect(&rcIXect, lprcPosRect, lprcClipRect) ) {
                if (!EqualRect(&rcIXect, lprcPosRect)) {
                    OffsetRect(&rcIXect, -(lprcPosRect->left), -(lprcPosRect->top));

                    HRGN tempRgn = CreateRectRgnIndirect(&rcIXect);
					if (0 != tempRgn)
					{
	                    SetWindowRgn(GetOuterWindow(), tempRgn, TRUE);

		                if (m_hRgn != NULL)
			               DeleteObject(m_hRgn);
				        m_hRgn = tempRgn;
	                    m_fUsingWindowRgn = TRUE;
		                fRemoveWindowRgn  = FALSE;
					}
					else
					{
						m_hRgn = 0;
	                    m_fUsingWindowRgn = FALSE;
		                fRemoveWindowRgn  = TRUE;
						TRACE2(1, _T("CreateRectRgnIndirect failed: %s %i"), __FILE__, __LINE__)
					}
                }
            }
        }

        if (fRemoveWindowRgn) 
		{
            SetWindowRgn(GetOuterWindow(), NULL, TRUE);
            if (m_hRgn != 0)
            {
               DeleteObject(m_hRgn);
               m_hRgn = 0;
            }
            m_fUsingWindowRgn = FALSE;
        }

        // set our control's location, but don't change it's size at all
        // [people for whom zooming is important should set that up here]
        DWORD dwFlag=0;
//        OnSetObjectRectsChangingWindowPos(&dwFlag);
        int cx, cy;
        cx = lprcPosRect->right - lprcPosRect->left;
        cy = lprcPosRect->bottom - lprcPosRect->top;
        SetWindowPos(GetOuterWindow(), NULL, lprcPosRect->left, lprcPosRect->top, cx, cy, dwFlag | SWP_NOZORDER | SWP_NOACTIVATE);
		if (m_pDockMgr)
			m_pDockMgr->ResizeDockableForms();
		RecalcLayout();
    }
    // save out our current location.  windowless controls want this 
    m_rcLocation = *lprcPosRect;
    return S_OK;
}
STDMETHODIMP CBar::ReactivateAndUndo()
{
	return E_NOTIMPL;
}

// IOleInPlaceActiveObject members

STDMETHODIMP CBar::TranslateAccelerator( LPMSG lpmsg)
{
	try
	{
		extern short _SpecialKeyState();
		// see if we want it or not.
		if (OnSpecialKey(lpmsg))
			return S_OK;
		// if not, then we want to forward it back to the site for further processing
		if (m_pControlSite)
			return m_pControlSite->TranslateAccelerator(lpmsg, _SpecialKeyState());
		// we didn't want it.
	}
	CATCH
	{
		//
		// This is here because when a user uses an accelator key that fires an event that closes the form
		// ActiveBar is release before we make it back.  One big crash!!!!!
		//
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
    return S_FALSE;
}
STDMETHODIMP CBar::OnFrameWindowActivate( BOOL fActivate)
{
    return InPlaceActivate(OLEIVERB_UIACTIVATE);
}
STDMETHODIMP CBar::OnDocWindowActivate( BOOL fActivate)
{
    return InPlaceActivate(OLEIVERB_UIACTIVATE);
}
STDMETHODIMP CBar::ResizeBorder( LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow)
{
	// we have no border
    return S_OK;
}
STDMETHODIMP CBar::EnableModeless( BOOL fEnable)
{
    return S_OK;
}

//{IOleInPlaceActiveObject Support Code}
short _SpecialKeyState()
{
    // don't appear to be able to reduce number of calls to GetKeyState
    //
    BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
    BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
    BOOL bAlt   = (GetKeyState(VK_MENU) < 0);

    return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}

BOOL CBar::OnSpecialKey(LPMSG pmsg)
{
	if (!(GetKeyState(VK_SHIFT) & 0x8000) && 
		!(GetKeyState(VK_CONTROL) & 0x8000) && 
		!(GetKeyState(VK_MENU) & 0x8000))
	{
		HWND hWndFocus = GetFocus();
		if (m_theFormsAndControls.IsChild(hWndFocus))
		{
			switch (pmsg->wParam)
			{
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
				SendMessage(hWndFocus, pmsg->message, pmsg->wParam, pmsg->lParam);
				return TRUE;
			}
		}

	}
	return FALSE;
}

//
// {IOleInPlaceActiveObject Support Code}
// IViewObject members
//
typedef BOOL (STDMETHODCALLTYPE* PFNCONTINUE)(DWORD dwContinue);

STDMETHODIMP CBar::Draw(DWORD dwDrawAspect, 
						LONG   lindex, 
						void*  pvAspect, 
						DVTARGETDEVICE* ptd, 
						HDC    hicTargetDev, 
						HDC    hdcDraw, 
						LPCRECTL lprcBounds, 
						LPCRECTL lprcWBounds, 
						PFNCONTINUE pfnContinue, 
						DWORD dwContinue)
{
    HRESULT hr;
    RECTL rc;
    POINT pVp, pW;
    BOOL  fOptimize = FALSE;
    int iMode;
    BYTE fMetafile = FALSE;
    BYTE fDeleteDC = FALSE;

    // support the aspects required for multi-pass drawing
    switch (dwDrawAspect) 
	{
        case DVASPECT_CONTENT:
        case DVASPECT_OPAQUE:
        case DVASPECT_TRANSPARENT:
		    break;
        default:
            return DV_E_DVASPECT;
    }

    // first, have to do a little bit to support printing.
    if (GetDeviceCaps(hdcDraw, TECHNOLOGY) == DT_METAFILE) 
	{
	    // We are dealing with a metafile.
        fMetafile = TRUE;
        // If attributes DC is NULL, create one, based on ptd.
        if (!hicTargetDev) 
		{
		    // Does _CreateOleDC have to return an hDC
            // or can it be flagged to return an hIC 
            // for this particular case?
            hicTargetDev = _CreateOleDC(ptd);
            fDeleteDC = TRUE;
        }
    }
    // check to see if we have any flags passed in the pvAspect parameter.
    if (pvAspect && ((DVASPECTINFO *)pvAspect)->cb == sizeof(DVASPECTINFO))
        fOptimize = (((DVASPECTINFO *)pvAspect)->dwFlags & DVASPECTINFOFLAG_CANOPTIMIZE) ? TRUE : FALSE;
    // if we are windowless, then we just pass this on to the end control code.
    if (m_fInPlaceActive) 
	{
	    // give them a rectangle with which to draw
        //ASSERT(!m_fInPlaceActive || !lprcBounds, "Inplace active and somebody passed in prcBounds!!!");
        if (lprcBounds)
			memcpy(&rc, lprcBounds, sizeof(rc));
		else
			memcpy(&rc, &m_rcLocation, sizeof(rc));
    } 
	else
	{
	    // first -- convert the DC back to MM_TEXT mapping mode so that the
        // window proc and OnDraw can share the same painting code.  save
        // some information on it, so we can restore it later [without using
        // a SaveDC/RestoreDC]
        rc = *lprcBounds;
        // Don't do anything to hdcDraw if it's a metafile.
        // The control's Draw method must make the appropriate
        // accomodations for drawing to a metafile
        if (!fMetafile) 
		{
            LPtoDP(hdcDraw, (POINT *)&rc, 2);
            SetViewportOrgEx(hdcDraw, 0, 0, &pVp);
            SetWindowOrgEx(hdcDraw, 0, 0, &pW);
            iMode = SetMapMode(hdcDraw, MM_TEXT);
        }
    }

    // prcWBounds is NULL and not used if we are not dealing with a metafile.
    // For metafiles, we pass on rc as *prcBounds, we should also include
    // prcWBounds
    hr = OnDraw(dwDrawAspect, hdcDraw, &rc, lprcWBounds, hicTargetDev, fOptimize);
    // clean up the DC when we're done with it, if appropriate.
    if (!m_fInPlaceActive) 
	{
        SetViewportOrgEx(hdcDraw, pVp.x, pVp.y, NULL);
        SetWindowOrgEx(hdcDraw, pW.x, pW.y, NULL);
        SetMapMode(hdcDraw, iMode);
    }

    if (fDeleteDC) 
		DeleteDC(hicTargetDev);
    return hr;
}
STDMETHODIMP CBar::GetColorSet(DWORD dwDrawAspect, 
							   LONG lindex, 
							   void* pvAspect, 
							   DVTARGETDEVICE* ptd, 
							   HDC hicTargetDev, 
							   LOGPALETTE** ppColorSet)
{
	if (dwDrawAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    *ppColorSet = NULL;
    return (OnGetPalette(hicTargetDev, ppColorSet)) ? ((*ppColorSet) ? S_OK : S_FALSE) : E_NOTIMPL;
}

STDMETHODIMP CBar::Freeze( DWORD dwDrawAspect, LONG lindex, void *pvAspect, DWORD *pdwFreeze)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBar::Unfreeze( DWORD dwFreeze)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBar::SetAdvise( DWORD dwAspects, DWORD dwAdviseFlags, IAdviseSink *pAdviseSink)
{
	TRACE(1, "IViewObject::SetAdvise\n");
    if (!(dwAspects & DVASPECT_CONTENT))
        return DV_E_DVASPECT;
    m_fViewAdvisePrimeFirst = (dwAdviseFlags & ADVF_PRIMEFIRST) ? TRUE : FALSE;
    m_fViewAdviseOnlyOnce = (dwAdviseFlags & ADVF_ONLYONCE) ? TRUE : FALSE;

	// cache advise sink interface
    if (m_pViewAdviseSink)
		m_pViewAdviseSink->Release();
    m_pViewAdviseSink = pAdviseSink;
	if (m_pViewAdviseSink)
		m_pViewAdviseSink->AddRef();
    
    // prime them if they want it [we need to store this so they can get flags later]
    if (m_fViewAdvisePrimeFirst)
        ViewChanged();
    return S_OK;
}
STDMETHODIMP CBar::GetAdvise( DWORD *pdwAspects, DWORD *pdwAdviseFlags, IAdviseSink **ppAdviseSink)
{
	TRACE(1, "IViewObject::GetAdvise\n");
    if (pdwAspects)
        *pdwAspects = DVASPECT_CONTENT;
    if (pdwAdviseFlags) 
	{
        *pdwAdviseFlags = 0;
        if (m_fViewAdviseOnlyOnce) 
			*pdwAdviseFlags |= ADVF_ONLYONCE;
        if (m_fViewAdvisePrimeFirst) 
			*pdwAdviseFlags |= ADVF_PRIMEFIRST;
    }
    if (ppAdviseSink) 
	{
        *ppAdviseSink = m_pViewAdviseSink;
        if (*ppAdviseSink)
			(*ppAdviseSink)->AddRef();
    }
    return S_OK;
}

//{IViewObject Support Code}
STDMETHODIMP CBar::OnDraw(DWORD dvAspect, 
						  HDC hdcDraw, 
						  LPCRECTL prcBounds, 
						  LPCRECTL prcWBounds, 
						  HDC hicTargetDev, 
						  BOOL fOptimize)
{
	if (!AmbientUserMode())
	{
		BSTR bstrName = NULL;
		BOOL bResult = GetAmbientProperty(DISPID_AMBIENT_DISPLAYNAME, VT_BSTR, &bstrName);
		if (bResult)
		{
			HFONT hFontOld = SelectFont(hdcDraw, GetFont());
			COLORREF crTextOld = SetTextColor(hdcDraw, RGB(0,0,0));
			COLORREF crBackOld = SetBkColor(hdcDraw, m_crBackground);
			CRect rcBounds(prcBounds->left, prcBounds->right, prcBounds->top, prcBounds->bottom);
			rcBounds.bottom -= m_pDockMgr->GetDockRect(ddDABottom).Height();
			rcBounds.right -= m_pDockMgr->GetDockRect(ddDARight).Width();
			SetBkMode(hdcDraw, TRANSPARENT);
			MAKE_TCHARPTR_FROMWIDE(szName, bstrName);
			DrawText(hdcDraw, 
					 szName, 
					 lstrlen(szName), 
					 (LPRECT)&rcBounds, 
					 DT_BOTTOM|DT_RIGHT|DT_SINGLELINE);
			SetTextColor(hdcDraw, crTextOld);
			SetBkColor(hdcDraw, crBackOld);
			SelectFont(hdcDraw, hFontOld);
		}
	}
	return NOERROR;
}

BOOL CBar::OnGetPalette(HDC hicTargetDevice,LOGPALETTE **ppColorSet)
{
	// use if control supports palettes
    return FALSE;
}

#ifdef CBar_SUBCLASS
HRESULT CBar::DoSuperClassPaint(HDC hdc,LPCRECTL prcBounds)
{
    HWND hwnd;
    RECT rcClient;
    int  iMapMode;
    POINT ptWOrg, ptVOrg;
    SIZE  sWOrg, sVOrg;

    // make sure we have a window.
    //
    hwnd = CreateInPlaceWindow(0,0, FALSE);
    if (!hwnd)
        return E_FAIL;

    GetClientRect(hwnd, &rcClient);

    if ((rcClient.right - rcClient.left != prcBounds->right - prcBounds->left)
        && (rcClient.bottom - rcClient.top != prcBounds->bottom - prcBounds->top)) 
	{

        iMapMode = SetMapMode(hdc, MM_ANISOTROPIC);
        SetWindowExtEx(hdc, rcClient.right, rcClient.bottom, &sWOrg);
        SetViewportExtEx(hdc, prcBounds->right - prcBounds->left, prcBounds->bottom - prcBounds->top, &sVOrg);
    }

    SetWindowOrgEx(hdc, 0, 0, &ptWOrg);
    SetViewportOrgEx(hdc, prcBounds->left, prcBounds->top, &ptVOrg);

#if STRICT
    CallWindowProc((WNDPROC)SUBCLASSWNDPROCOFCONTROL2(), hwnd, (TRUE) ? WM_PRINT : WM_PAINT, (WPARAM)hdc, (LPARAM)(TRUE ? PRF_CHILDREN | PRF_CLIENT : 0));
#else
    CallWindowProc((FARPROC)SUBCLASSWNDPROCOFCONTROL2(), hwnd, (TRUE) ? WM_PRINT : WM_PAINT, (WPARAM)hdc, (LPARAM)(TRUE ? PRF_CHILDREN | PRF_CLIENT : 0));
#endif // STRICT

    return S_OK;
}
#endif
//{IViewObject Support Code}

// IViewObject2 members

STDMETHODIMP CBar::GetExtent( DWORD dwDrawAspect, LONG lindex, DVTARGETDEVICE *ptd, LPSIZEL lpsizel)
{
	TRACE(1, "IViewObject2::GetExtent\n");
	return GetExtent(dwDrawAspect, lpsizel); // use IOleObject impl.
}

// IPersist members

STDMETHODIMP CBar::GetClassID( LPCLSID lpClassID)
{
	*lpClassID=*(((UNKNOWNOBJECTDESC *)(g_objectTable[0].objectDefPtr))->rclsid);
	return NOERROR;
}

// IPersistPropertyBag members

STDMETHODIMP CBar::Load( IPropertyBag *pPropBag, IErrorLog *pErrorLog)
{
	return PersistText(pPropBag,pErrorLog, FALSE);
}
STDMETHODIMP CBar::Save( IPropertyBag *pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties)
{
	HRESULT hResult = PersistText(pPropBag, NULL, TRUE);
	if (FAILED(hResult))
		return hResult;

	if (fClearDirty)
        m_fDirty = FALSE;

    if (m_pOleAdviseHolder)
        m_pOleAdviseHolder->SendOnSave();

	return NOERROR;
}

//{IPersistPropertyBag Support Code}
HRESULT CBar::PersistText(IPropertyBag* pPropBag, IErrorLog* pErrorLog, BOOL save)
{
	VARIANT_BOOL vbSave = save ? VARIANT_TRUE : VARIANT_FALSE;
	HRESULT hResult;
	SIZEL slHiMetric;

    // Persist Standard properties
	if (VARIANT_TRUE == vbSave)
	    PixelToHiMetric(&m_Size, &slHiMetric);

	BOOL bOldVersion = FALSE;
	long nVersion = eLayoutVersion;

	hResult = PersistBagI4(pPropBag, L"_LayoutVersion", pErrorLog, (LONG*)&nVersion, vbSave);
	if (FAILED(hResult))
	{
		bOldVersion = TRUE;
	}

	hResult = PersistBagI4(pPropBag, L"_ExtentX", pErrorLog, (LONG*)&slHiMetric.cx, vbSave);
	if (FAILED(hResult))
		return hResult;

	hResult = PersistBagI4(pPropBag, L"_ExtentY", pErrorLog, (LONG*)&slHiMetric.cy, vbSave);
	if (FAILED(hResult))
		return hResult;

	if (!bOldVersion)
	{
		hResult = PersistBagBSTR(pPropBag, L"_DataPath", pErrorLog, &m_bstrDataPath, vbSave);
		if (FAILED(hResult))
			return hResult;
	}

	if (VARIANT_FALSE == vbSave)
		HiMetricToPixel(&slHiMetric, &m_Size);

	//{PERSISTBAG}
	//{PERSISTBAG}
	IStream* pStream;
	hResult = CreateStreamOnHGlobal(NULL, TRUE, &pStream);
	if (FAILED(hResult))
		return hResult;

	ULONG nDataSize = 0;
	if (VARIANT_TRUE == vbSave)
	{
		hResult = PersistState(pStream, vbSave);
		if (FAILED(hResult))
			return hResult;

		//
		// Get DataSize
		//
		
		LARGE_INTEGER moveAmount; 
		moveAmount.LowPart = moveAmount.HighPart = 0;
		ULARGE_INTEGER curPos;
		pStream->Seek(moveAmount,STREAM_SEEK_CUR,&curPos);
		nDataSize = curPos.LowPart;
	}	

	VARIANT vStream;
	vStream.vt = VT_UNKNOWN;
	vStream.punkVal = new MemStream(pStream, nDataSize);
	if (NULL == vStream.punkVal)
		return E_OUTOFMEMORY;

	hResult = PersistBagVariant(pPropBag, L"Bands", pErrorLog, &vStream, vbSave);
	if (FAILED(hResult))
	{
		pStream->Release();
		return hResult;
	}

	if (VARIANT_FALSE == vbSave)
	{
		// Move to Start of Stream
		LARGE_INTEGER moveAmount; 
		moveAmount.LowPart = moveAmount.HighPart = 0;
		ULARGE_INTEGER curPos;
		pStream->Seek(moveAmount, STREAM_SEEK_SET, &curPos);
		hResult = PersistState(pStream, vbSave);
		if (FAILED(hResult))
		{
			pStream->Release();
			return hResult;
		}
	}
	VariantClear(&vStream);
	CacheColors();
	CacheTexture();
	BuildAccelators(GetMenuBand());
	OnFontHeightChanged();

	return NOERROR;
}
//{IPersistPropertyBag Support Code}

// IPersistStreamInit members

STDMETHODIMP CBar::IsDirty()
{
    return (m_fDirty) ? S_OK : S_FALSE;
}
STDMETHODIMP CBar::Load( LPSTREAM pStm)
{
    // first thing to do is read in standard properties the user don't
    // persist themselves.
    HRESULT hr = LoadStandardState(pStm);
    RETURN_ON_FAILURE(hr);
    // load in the user properties.  this method is one they -have- to implement
    // themselves.
    hr = PersistBinaryState(pStm, VARIANT_FALSE);
    return hr;
}
STDMETHODIMP CBar::Save( LPSTREAM pStm, BOOL fClearDirty)
{
    // use our helper routine that we share with the IStorage persistence
    // code.
    HRESULT hr = m_SaveToStream(pStm);
    RETURN_ON_FAILURE(hr);
    // clear out dirty flag [if appropriate] and notify that we're done
    // with save.
    if (fClearDirty)
        m_fDirty = FALSE;
    if (m_pOleAdviseHolder)
        m_pOleAdviseHolder->SendOnSave();
    return S_OK;
}
STDMETHODIMP CBar::GetSizeMax( ULARGE_INTEGER *pCbSize)
{
    return E_NOTIMPL;
}
STDMETHODIMP CBar::InitNew()
{
    BOOL f = InitializeNewState();
    return (f) ? S_OK : E_FAIL;
}

//{IPersistStreamInit Support Code}
#define STREAMHDR_SIGNATURE 0x55441234

typedef struct tagSTREAMHDR {

    DWORD  dwSignature;     // Signature.
    size_t cbWritten;       // Number of bytes written

} STREAMHDR;

HRESULT CBar::LoadStandardState(IStream *pStream)
{
    STREAMHDR stmhdr;

    HRESULT hr = pStream->Read(&stmhdr, sizeof(STREAMHDR), NULL);
    RETURN_ON_FAILURE(hr);

    if (stmhdr.dwSignature != STREAMHDR_SIGNATURE)
        return E_UNEXPECTED;

    if (stmhdr.cbWritten != sizeof(m_Size))
        return E_UNEXPECTED;

    SIZEL   slHiMetric;
    hr = pStream->Read(&slHiMetric, sizeof(slHiMetric), NULL);
    RETURN_ON_FAILURE(hr);

    HiMetricToPixel(&slHiMetric, &m_Size);
    return S_OK;
}

#ifdef PERFORMACE_TEST
	DWORD persisttick;
	extern BOOL measurePersistToDraw;
#endif

HRESULT CBar::PersistBinaryState(IStream *pStream, BOOL save)
{
	VARIANT_BOOL vbSave = save ? VARIANT_TRUE : VARIANT_FALSE;
	//{PERSIST BINARY}
	//{PERSIST BINARY}
	
#ifdef PERFORMACE_TEST
	DWORD startTick,tickCount;
	if (save)
	{
//		OutputDebugString("Saving\n");
		startTick=lastTick=tickCount=GetTickCount();
	}
	else
	{
//		OutputDebugString("Loading\n");
		startTick=lastTick=tickCount=GetTickCount();
	}
#endif
	HRESULT hResult = PersistState(pStream, vbSave);

	CacheColors();
	CacheTexture();
	BuildAccelators(GetMenuBand());
	OnFontHeightChanged();
#ifdef PERFORMACE_TEST
	char str[100];
	tickCount=GetTickCount();
	wsprintf(str,"Load/Save complete in %d \n",tickCount-startTick);
//	OutputDebugString(str);
	persisttick=tickCount;
	measurePersistToDraw=TRUE;
#endif
	return hResult;
}

HRESULT CBar::SaveStandardState(IStream *pStream)
{
    STREAMHDR streamhdr = { STREAMHDR_SIGNATURE, sizeof(SIZEL) };
    HRESULT hr;
    SIZEL   slHiMetric;

    hr = pStream->Write(&streamhdr, sizeof(STREAMHDR), NULL);
    RETURN_ON_FAILURE(hr);

    PixelToHiMetric(&m_Size, &slHiMetric);

    hr = pStream->Write(&slHiMetric, sizeof(slHiMetric), NULL);
    return hr;
}

//{IPersistStreamInit Support Code}

// IPersistStorage members

STDMETHODIMP CBar::InitNew( IStorage *pStg)
{
	TRACE(1, "IPersistStorage::InitNew\n");
    // call the overridable function to do this work
    BOOL bInitSuccess = InitializeNewState();
    return (bInitSuccess) ? S_OK : E_FAIL;
}
STDMETHODIMP CBar::Load( IStorage *pStg)
{
	TRACE(1, "IPersistStorage::Load\n");
    // we're going to use IPersistStream::Load from the CONTENTS stream.
	IStream *pStream;
    HRESULT  hr = pStg->OpenStream(wszCtlSaveStream, 0, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);
    RETURN_ON_FAILURE(hr);
	#ifdef DEF_PERSISTSTREAM
    hr = Load(pStream);
	#endif
    pStream->Release();
    return hr;
}

STDMETHODIMP CBar::Save( IStorage *pStgSave, BOOL fSameAsLoad)
{
	TRACE(1, "IPersistStorage::Save\n");
    IStream* pStream;
    // we're just going to save out to the CONTENTES stream.
    HRESULT hr = pStgSave->CreateStream(wszCtlSaveStream, 
										STGM_WRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE,
                                        0, 
										0, 
										&pStream);
    RETURN_ON_FAILURE(hr);
    hr = m_SaveToStream(pStream);
    m_bSaveSucceeded = (FAILED(hr)) ? FALSE : TRUE;
    pStream->Release();
    return hr;
}
STDMETHODIMP CBar::SaveCompleted( IStorage *pStgNew)
{
	TRACE(1, "IPersistStorage::SaveCompleted\n");
    // if our save succeeded, then we can do our post save work.
    if (m_bSaveSucceeded) 
	{
        m_fDirty = FALSE;
        if (m_pOleAdviseHolder)
            m_pOleAdviseHolder->SendOnSave();
    }
    return S_OK;
}
STDMETHODIMP CBar::HandsOffStorage()
{
	TRACE(1, "IPersistStorage::HandsOffStorage\n");
	return S_OK; // only used when caching IStorage pointer
}


//{IPersistStorage Support Code}

BOOL CBar::InitializeNewState()
{
	//{PERSIST INIT}
	//{PERSIST INIT}
	return TRUE;
}

HRESULT CBar::m_SaveToStream(IStream *pStream)
{
	HRESULT hr = SaveStandardState(pStream);
    RETURN_ON_FAILURE(hr);
	hr = PersistBinaryState(pStream,VARIANT_TRUE);
    return hr;
}
//{IPersistStorage Support Code}

//                                                                                                                               members

STDMETHODIMP CBar::EnumConnectionPoints( IEnumConnectionPoints **ppEnum)
{
    TRACE(1, "Enum connection points\n");
	IConnectionPoint **rgConnectionPoints;

    // HeapAlloc an array of connection points [since our standard enum
    // assumes this and HeapFree's it later ]
    //
    rgConnectionPoints = (IConnectionPoint **)HeapAlloc(g_hHeap, 0, sizeof(IConnectionPoint *) * 2);
    RETURN_ON_NULLALLOC(rgConnectionPoints);

    // we support the event interface for this dude as well as IPropertyNotifySink
    //
    rgConnectionPoints[0] = (IConnectionPoint*)m_cpEvents;
    rgConnectionPoints[1] = (IConnectionPoint*)m_cpPropNotify;

    *ppEnum = (IEnumConnectionPoints *)(IEnumX *) new CEnumX(IID_IEnumConnectionPoints,
															 2, 
															 sizeof(IConnectionPoint*), 
															 (void *)rgConnectionPoints,
															 CopyAndAddRefObject);
    if (!*ppEnum) 
	{
        HeapFree(g_hHeap, 0, rgConnectionPoints);
        return E_OUTOFMEMORY;
    }

    return S_OK;
}
STDMETHODIMP CBar::FindConnectionPoint( REFIID riid,  IConnectionPoint **ppCP)
{
	TRACE(1, "FindConnectionPoint\n");
    // we support the event interface, and IDispatch for it, and we also
    // support IPropertyNotifySink.
    //
    if ((EVENTREFIIDOFCONTROL(0)!=NULL && 
		DO_GUIDS_MATCH(riid, (*EVENTREFIIDOFCONTROL(0)))) || DO_GUIDS_MATCH(riid, IID_IDispatch))
        *ppCP = (IConnectionPoint*)m_cpEvents;
    else if (DO_GUIDS_MATCH(riid, IID_IPropertyNotifySink))
        *ppCP = (IConnectionPoint*)m_cpPropNotify;
    else
        return E_NOINTERFACE;

    (*ppCP)->AddRef();
    return S_OK;
}

// ISpecifyPropertyPages members

STDMETHODIMP CBar::GetPages( CAUUID *pPages)
{
    const GUID **pElems;
    void *pv;
    WORD  x;

	pPages->cElems=PROPPAGECOUNT(objectDef);
	if (pPages->cElems==0)
	{
		pPages->pElems=NULL;
		return S_OK;
	}

	pv = CoTaskMemAlloc(sizeof(GUID) * (pPages->cElems));
    if (!pv)
		return E_OUTOFMEMORY;

    pPages->pElems = (GUID *)pv;

    pElems = PROPPAGES(objectDef);
    for (x = 0; x < pPages->cElems; x++)
        pPages->pElems[x] = *(pElems[x]);

	return S_OK;
}

// IProvideClassInfo members

STDMETHODIMP CBar::GetClassInfo( ITypeInfo **ppTI)
{
	ITypeLib *pTypeLib;

	*ppTI = NULL;
    HRESULT hr;
    hr = LoadRegTypeLib(LIBID_PROJECT, 
						((CONTROLOBJECTDESC*)objectDef)->automationDesc.versionMajor,
						((CONTROLOBJECTDESC*)objectDef)->automationDesc.versionMinor,
    					LANGIDFROMLCID(g_lcidLocale),
						&pTypeLib);
    RETURN_ON_FAILURE(hr);

    // got the typelib.  get typeinfo for our coclass.
    hr = pTypeLib->GetTypeInfoOfGuid((REFIID)CLSIDOFOBJECT(objectDef), ppTI);
    pTypeLib->Release();
    RETURN_ON_FAILURE(hr);
    return S_OK;
}

//
// StartCustomization
//

void CBar::StartCustomization()
{
	FireCustomizeBegin();
	if (VARIANT_FALSE == bpV1.m_vbUserDefinedCustomization)
	{
		DoCustomization(TRUE);
		m_pCustomize = new CCustomize(GetDockWindow(), this);
		if (m_pCustomize)
		{
			HWND hWndPropertySheet = m_pCustomize->ShowSheet();
			MSG msg;
			while (GetMessage(&msg, NULL, 0, 0))
			{
				if (!IsWindow(hWndPropertySheet))
					break;
				if (NULL == m_pCustomize)
					break;
				if (PropSheet_IsDialogMessage(hWndPropertySheet, &msg))
					continue;
				DispatchMessage(&msg);
			}
		}
	}
}

//
// FireToolClick
//

void CBar::FireToolClick( Tool *tool)
{
	try
	{
		CTool* pTool = reinterpret_cast<CTool*>(tool);
		if (VARIANT_FALSE == pTool->tpV1.m_vbEnabled)
			return;

		pTool->tpV1.m_nUsage++;
		pTool->m_nImageLoad = m_pImageMgr->m_nImageLoad;

		long nId = pTool->tpV1.m_nToolId;
		if (nId & eSpecialToolId)
		{
			switch (nId)
			{
			case eToolIdCustomize:
				StartCustomization();
				return;

			case eToolIdResetToolbar:
				FireReset(pTool->m_vTag.bstrVal);
				return;

			case eToolIdRestore:
			case eToolIdMininize:
			case eToolIdClose:
				{
					HWND hWndActive = (HWND)SendMessage(m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
					if (hWndActive)
					{
						switch (nId)
						{
						case eToolIdRestore:
							ShowWindow(hWndActive, SW_RESTORE);
							break;
						case eToolIdMininize:
							ShowWindow(hWndActive, SW_MINIMIZE);
							break;
						case eToolIdClose:
							SendMessage(hWndActive, WM_SYSCOMMAND, SC_CLOSE, 0);
							break;
						}
					}
				}
				return;

			case eToolIdCascade:
				CascadeWindows(m_hWndMDIClient, 0, NULL, 0, NULL); 
				return;

			case eToolIdTileHorz:
				TileWindows(m_hWndMDIClient, MDITILE_HORIZONTAL, NULL, 0, NULL); 
				return;

			case eToolIdTileVert:
				TileWindows(m_hWndMDIClient, MDITILE_VERTICAL, NULL, 0, NULL); 
				return;

			case eEditToolId: // Reset
				FireToolReset((Tool*)m_mcInfo.pTool);
				return;

			case eEditToolId+1: // Delete
				m_mcInfo.pBand->m_pTools->DeleteTool(m_mcInfo.pTool);
				m_vbCustomizeModified = VARIANT_TRUE;
				if (m_pDesignerNotify)
				{
					m_pDesignerNotify->BandChanged((LPDISPATCH)(IBand*)m_mcInfo.pBand, 
												   (LPDISPATCH)(IBand*)(ddCBSNone != m_mcInfo.pBand->bpV1.m_cbsChildStyle ? m_mcInfo.pBand->m_pChildBands->GetCurrentChildBand() : NULL),
												   ddBandModified);
					m_pDesigner->SetDirty();
				}
				RecalcLayout();
				return;

			case eEditToolId+5: // CopyImage
				{
					if (!OpenClipboard(m_hWnd))
					{
						return;
					}

					EmptyClipboard();
					
					HBITMAP hBmp = NULL;
					HDC hDC = GetDC(NULL);
					if (hDC)
					{
						HDC hDCBitmap = CreateCompatibleDC(hDC);
						if (hDCBitmap)
						{
							hBmp = CreateCompatibleBitmap(hDC, m_mcInfo.pTool->m_sizeImage.cx, m_mcInfo.pTool->m_sizeImage.cy);
							if (hBmp)
							{
								HBITMAP hBitOld = SelectBitmap(hDCBitmap, hBmp);
								CRect rc;
								rc.right = m_mcInfo.pTool->m_sizeImage.cx;
								rc.bottom = m_mcInfo.pTool->m_sizeImage.cy;
								FillRect(hDCBitmap,&rc,(HBRUSH)(1+COLOR_BTNFACE));
								m_mcInfo.pTool->DrawPict((OLE_HANDLE)hDCBitmap, 
														 0,
														 0,
														 m_mcInfo.pTool->m_sizeImage.cx,
														 m_mcInfo.pTool->m_sizeImage.cy, 
														 VARIANT_FALSE);
								hBitOld = SelectBitmap(hDCBitmap, hBitOld);
								DeleteDC(hDCBitmap);
							}
						}
						ReleaseDC(NULL, hDC);
					}

					HANDLE hBM = SetClipboardData(CF_BITMAP, hBmp);
					if (NULL == hBM)
					{
						BOOL bResult = DeleteBitmap(hBmp);
						assert(bResult);
						CloseClipboard();
						return;
					}

					if (!CloseClipboard())
					{
						return;
					}
				}
				return;

			case eEditToolId+6: // Paste Image
				{
					if (!OpenClipboard(m_hWnd))
					{
						return;
					}

					HANDLE hBM = ::GetClipboardData(CF_BITMAP);
					if (hBM)
					{
						HRESULT hResult = m_mcInfo.pTool->put_Bitmap(ddITNormal, (OLE_HANDLE)hBM);
						if (SUCCEEDED(hResult))
						{
							m_vbCustomizeModified = VARIANT_TRUE;
							HDC hDC = GetDC(m_hWnd);
							if (hDC)
							{
								BITMAP bmInfo;
								GetObject(hBM, sizeof(BITMAP), &bmInfo);
								int nWidth = bmInfo.bmWidth;
								int nHeight = bmInfo.bmHeight;
								HBITMAP hMask = CreateMaskBitmap(hDC, (HBITMAP)hBM, nWidth, nHeight);
								if (hMask)
								{
									hResult = m_mcInfo.pTool->put_MaskBitmap(ddITNormal, (OLE_HANDLE)hMask);
									BOOL bResult = DeleteBitmap(hMask);
									assert(bResult);
								}
								ReleaseDC(m_hWnd, hDC);

							}
							RecalcLayout();
						}
					}

					if (!CloseClipboard())
					{
						return;
					}
				}
				return;

			case eEditToolId+8: // Default
				m_mcInfo.pTool->tpV1.m_tsStyle = ddSStandard;
				m_vbCustomizeModified = VARIANT_TRUE;
				RecalcLayout();
				return;

			case eEditToolId+9: // Text Only
				m_mcInfo.pTool->tpV1.m_tsStyle = ddSText;
				m_vbCustomizeModified = VARIANT_TRUE;
				RecalcLayout();
				return;
			
			case eEditToolId+10: // Icon Only
				m_mcInfo.pTool->tpV1.m_tsStyle = ddSIcon;
				m_vbCustomizeModified = VARIANT_TRUE;
				RecalcLayout();
				return;

			case eEditToolId+11: // Icon & Text
				m_mcInfo.pTool->tpV1.m_tsStyle = ddSIconText;
				m_vbCustomizeModified = VARIANT_TRUE;
				RecalcLayout();
				return;
			
			default:
				switch (nId)
				{
				case eToolIdWindowList:
					if (IsWindow(m_hWndMDIClient))
					{
						HWND hWnd = (HWND)SendMessage(m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
						if (pTool->m_hWndActive != hWnd)
						{
							hWnd = pTool->m_hWndActive;
							SendMessage(hWnd, WM_SYSCOMMAND, SC_RESTORE, -1);
							SendMessage(m_hWndMDIClient, WM_MDIACTIVATE, (WPARAM)hWnd, 0);
							// Let it fire a tool click for Neal 
						}
					}
					break;

				case eToolIdToolToggle:
					{
						CTool* pToolToggle = (CTool*)(ITool*)pTool->m_vTag.pdispVal;
						if (pToolToggle)
						{
							pToolToggle->tpV1.m_vbVisible = pTool->tpV1.m_vbChecked == VARIANT_TRUE ? VARIANT_FALSE : VARIANT_TRUE;
							pTool->tpV1.m_vbChecked = (VARIANT_TRUE == pTool->tpV1.m_vbChecked ? VARIANT_FALSE : VARIANT_TRUE);
							PostMessage(m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
							if (m_pToolShadow)
								m_pToolShadow->Hide(TRUE);
						}
					}
					return;

				case eToolIdBandToggle:
					{
						CBand* pBand = FindBand(pTool->m_vTag.bstrVal);
						if (pBand)
						{
							pBand->put_Visible(pBand->bpV1.m_vbVisible == VARIANT_TRUE ? VARIANT_FALSE : VARIANT_TRUE);
							PostMessage(m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
							if (pBand->bpV1.m_vbVisible == VARIANT_TRUE)
								RecalcLayout();
							else
								PostMessage(m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
							return;
						}
					}
					break;
				}
				break;
			}
		}
		try 
		{
			AddRef();
			if (m_cpEvents) 
				m_cpEvents->DoInvoke(DISPID_TOOLCLICK,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
			Release();
		}
		CATCH
		{
			//
			// If we fire an event that causes VB to shutdown we will crash.  So catch it here.
			//
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}			
	}
	CATCH
	{
		//
		// If we fire an event that causes VB to shutdown we will crash.  So catch it here.
		//
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
}

void CBar::FireTextChange( Tool *tool)
{
	if (eEditToolId + 3 == reinterpret_cast<CTool*>(tool)->tpV1.m_nToolId)
	{
		m_mcInfo.pTool->put_Caption(reinterpret_cast<CTool*>(tool)->m_bstrText);
		m_vbCustomizeModified = VARIANT_TRUE;
		RecalcLayout();
		return;
	}
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TEXTCHANGE,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireReset( BSTR BandName)
{
	if (m_pCustomize)
		m_pCustomize->Clear();
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_RESET,EVENT_PARAM(VTS_WBSTR VTS_NONE),BandName);
	if (m_pCustomize)
		m_pCustomize->Reset();
}

//{EVENT WRAPPERS}

void CBar::FireNewToolbar( ReturnString *Name)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_NEWTOOLBAR,EVENT_PARAM(VTS_DISPATCH VTS_NONE),Name);
}

void CBar::FireComboDrop( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_COMBODROP,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireComboSelChange( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_COMBOSELCHANGE,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireBandOpen( Band *band,  ReturnBool *Cancel)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_OPEN,EVENT_PARAM(VTS_DISPATCH VTS_DISPATCH VTS_NONE),band,Cancel);
}

void CBar::FireBandClose( Band *band)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_BANDCLOSE,EVENT_PARAM(VTS_DISPATCH VTS_NONE),band);
}

void CBar::FireToolDblClick( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLDBLCLK,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireBandMove( Band *band)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_MOVE,EVENT_PARAM(VTS_DISPATCH VTS_NONE),band);
}

void CBar::FireBandResize( Band *band, VARIANT *Widths, VARIANT *Heights,  long BandWidth,  long BandHeight)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_BANDRESIZE,EVENT_PARAM(VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_I4 VTS_I4 VTS_NONE),band,Widths,Heights,BandWidth,BandHeight);
}

void CBar::FireError( short Number, ReturnString *Description, long Scode, BSTR Source, BSTR HelpFile, long HelpContext, ReturnBool *CancelDisplay)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_ERROR,EVENT_PARAM(VTS_I2 VTS_DISPATCH VTS_I4 VTS_WBSTR VTS_WBSTR VTS_I4 VTS_DISPATCH VTS_NONE),Number,Description,Scode,Source,HelpFile,HelpContext,CancelDisplay);
}

void CBar::FireDataReady()
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_DATAREADY,EVENT_PARAM(VTS_NONE));
}

void CBar::FireMouseEnter( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_MOUSEENTER,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireMouseExit( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_MOUSEEXIT,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireMenuItemEnter( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_MENUITEMENTER,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireMenuItemExit( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_MENUITEMEXIT,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireKeyDown( short keycode, short shift)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_KEYDOWN,EVENT_PARAM(VTS_I2 VTS_I2 VTS_NONE),keycode,shift);
}

void CBar::FireKeyUp( short keycode, short shift)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_KEYUP,EVENT_PARAM(VTS_I2 VTS_I2 VTS_NONE),keycode,shift);
}

void CBar::FireChildBandChange( Band *band)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_CHILDBANDCHANGE,EVENT_PARAM(VTS_DISPATCH VTS_NONE),band);
}

void CBar::FireCustomizeBegin()
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_CUSTOMIZEBEGIN,EVENT_PARAM(VTS_NONE));
}

void CBar::FireFileDrop( Band *band,  BSTR FileName)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_FILEDROP,EVENT_PARAM(VTS_DISPATCH VTS_WBSTR VTS_NONE),band,FileName);
}

void CBar::FireCustomizeEnd( VARIANT_BOOL bModified)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_CUSTOMIZEEND,EVENT_PARAM(VTS_BOOL VTS_NONE),bModified);
}

void CBar::FireMouseDown( short Button,  short Shift,  float x,  float y)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_MOUSEDOWN,EVENT_PARAM(VTS_I2 VTS_I2 VTS_R4 VTS_R4 VTS_NONE),Button,Shift,x,y);
}

void CBar::FireMouseMove( short Button,  short Shift,  float x,  float y)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_MOUSEMOVE,EVENT_PARAM(VTS_I2 VTS_I2 VTS_R4 VTS_R4 VTS_NONE),Button,Shift,x,y);
}

void CBar::FireMouseUp( short Button,  short Shift,  float x,  float y)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_MOUSEUP,EVENT_PARAM(VTS_I2 VTS_I2 VTS_R4 VTS_R4 VTS_NONE),Button,Shift,x,y);
}

void CBar::FireResize( long Left,  long Top,  long Width,  long Height)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_RESIZE,EVENT_PARAM(VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_NONE),Left,Top,Width,Height);
}

void CBar::FireBandDock( Band *band)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_BANDDOCK,EVENT_PARAM(VTS_DISPATCH VTS_NONE),band);
}

void CBar::FireBandUndock( Band *band)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_BANDUNDOCK,EVENT_PARAM(VTS_DISPATCH VTS_NONE),band);
}

void CBar::FireToolReset(Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLRESET,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireToolGotFocus( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_GOTFOCUS,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireToolLostFocus( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_LOSTFOCUS,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireWhatsThisHelp( Band *band, Tool *tool,  long HelpId)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_WHATSTHISHELPEVENT,EVENT_PARAM(VTS_DISPATCH VTS_DISPATCH VTS_I4 VTS_NONE),band,tool,HelpId);
}

void CBar::FireCustomizeHelp( short ControlId)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_CUSTOMIZEHELP,EVENT_PARAM(VTS_I2 VTS_NONE),ControlId);
}

void CBar::Fire_ToolKeyDown( short keycode, short shift)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLKEYDOWN,EVENT_PARAM(VTS_I2 VTS_I2 VTS_NONE),keycode,shift);
}

void CBar::Fire_ToolKeyUp( Tool*tool,  short keycode, short shift)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLKEYUP,EVENT_PARAM(VTS_DISPATCH VTS_I2 VTS_I2 VTS_NONE),tool,keycode,shift);
}

void CBar::FireClick()
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_CLICK,EVENT_PARAM(VTS_NONE));
}

void CBar::FireDblClick()
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_DOUBLECLICK,EVENT_PARAM(VTS_NONE));
}

void CBar::FireCustomizeToolClick( Tool *tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_CUSTOMIZETOOLCLICK,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::Fire_ToolKeyPress( Tool*tool,  long keyascii )
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLKEYPRESS,EVENT_PARAM(VTS_DISPATCH VTS_I4 VTS_NONE),tool,keyascii );
}

void CBar::Fire_ToolComboClose( Tool*tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLCOMBOCLOSE,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireToolKeyDown( Tool*tool,  short*keycode, short shift)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLKEYDOWN2,EVENT_PARAM(VTS_DISPATCH VTS_PI2 VTS_I2 VTS_NONE),tool,keycode,shift);
}

void CBar::FireToolKeyUp( Tool*tool,  short*keycode, short shift)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLKEYUP2,EVENT_PARAM(VTS_DISPATCH VTS_PI2 VTS_I2 VTS_NONE),tool,keycode,shift);
}

void CBar::FireToolKeyPress( Tool*tool,  long*keyascii )
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLKEYPRESS2,EVENT_PARAM(VTS_DISPATCH VTS_PI4 VTS_NONE),tool,keyascii );
}

void CBar::FireToolComboClose( Tool*tool)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_TOOLCOMBOCLOSE2,EVENT_PARAM(VTS_DISPATCH VTS_NONE),tool);
}

void CBar::FireQueryUnload( short*Cancel)
{
	if (m_cpEvents) m_cpEvents->DoInvoke(DISPID_QUERYUNLOAD,EVENT_PARAM(VTS_PI2 VTS_NONE),Cancel);
}
//{EVENT WRAPPERS}
//{PROPERTY NOTIFICATION SINK}

CBar *CBar::CPropertyNotifySink::m_pMainUnknown() 
{
	return (CBar *)((LPBYTE)this - offsetof(CBar, m_xPropNotify));
}

STDMETHODIMP CBar::CPropertyNotifySink::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (riid==IID_IUnknown || riid==IID_IPropertyNotifySink)
	{
		*ppvObjOut=this;
		AddRef();
		return NOERROR;
	}
	return E_NOTIMPL;
}

STDMETHODIMP CBar::OnRequestEdit(DISPID dispID)
{
	return NOERROR;
}

STDMETHODIMP CBar::OnChanged(DISPID dispID)
{
	BOOL bResult;
	if (m_hFontMenuVert)
	{
		bResult = DeleteFont(m_hFontMenuVert);
		assert(bResult);
		m_hFontMenuVert = NULL;
	}
	if (m_hFontMenuHorz)
	{
		bResult = DeleteFont(m_hFontMenuHorz);
		assert(bResult);
		m_hFontMenuHorz = NULL;
	}
	if (m_hFontChildBand)
		m_hFontChildBand = NULL;
	if (m_hFontChildBandVert)
	{
		bResult = DeleteFont(m_hFontChildBandVert);
		assert(bResult);
		m_hFontChildBandVert = NULL;
	}
	if (m_cached_fhFont)
		m_cached_fhFont=NULL;

	if (m_cachedControlFont)
		m_cachedControlFont=NULL;
	OnFontHeightChanged();
	ViewChanged();
	return NOERROR;
}

//{PROPERTY NOTIFICATION SINK}

HRESULT CBar::OnVerb(LONG lVerb)
{
	return OLEOBJ_S_INVALIDVERB;
}

STDMETHODIMP CBar::GetControlInfo(CONTROLINFO *pCI)
{
	pCI->hAccel = NULL;
	pCI->cAccel = 0;    
	pCI->dwFlags = 0;  

	if (eClientArea == m_eAppType)
	{
		TypedArray<WORD> faChar;
		BuildAccelators(GetMenuBand(), &faChar);

		int nCount = faChar.GetSize();
		if (nCount > 0)
		{
			ACCEL* pAccel = new ACCEL[nCount];
			for (int nIndex = 0; nIndex < nCount; nIndex++)
			{
				pAccel[nIndex].fVirt = FVIRTKEY | FALT;
				pAccel[nIndex].key = LOBYTE(VkKeyScan((TCHAR)faChar.GetAt(nIndex)));
				pAccel[nIndex].cmd = 1;
			}
			if (m_hAccel)
				DestroyAcceleratorTable(m_hAccel);
			m_hAccel = CreateAcceleratorTable(pAccel, nCount);  
			if (m_hAccel != NULL)  
			{
				pCI->hAccel = m_hAccel;
				pCI->cAccel = nCount;    
			}
			delete [] pAccel;
		}
		return NOERROR;
	}
    return E_NOTIMPL;
}
STDMETHODIMP CBar::OnMnemonic( MSG *pMsg)
{
	if (eClientArea == m_eAppType && WM_SYSKEYDOWN == pMsg->message && DoMenuAccelator(pMsg->wParam))
		return NOERROR;
	return E_NOTIMPL;
}

STDMETHODIMP CBar::OnAmbientPropertyChange( DISPID dispID)
{
	switch (dispID)
	{
	case DISPID_UNKNOWN:
        m_fModeFlagValid = FALSE;
		break;

	case DISPID_AMBIENT_DISPLAYNAME:
		InvalidateRect(m_hWnd, NULL, TRUE);
		break;

	case DISPID_BACKCOLOR:
		if (GetAmbientProperty(DISPID_AMBIENT_BACKCOLOR, VT_I4, &m_ocAmbientBackColor))
			OleTranslateColor(m_ocAmbientBackColor, NULL, &m_crAmbientBackColor);
		break;
	}

    AmbientPropertyChanged(dispID);
    return S_OK;
}
STDMETHODIMP CBar::FreezeEvents( BOOL bFreeze)
{
	m_bFreezeEvents = bFreeze;
	if (!bFreeze)
	{
		if (m_pFreezeBand)
		{
			FireBandClose((Band*)m_pFreezeBand);
			m_pFreezeBand->Release();
			m_pFreezeBand = NULL;
		}
	}
    return S_OK;
}

//{IOleControl Support Code}
void CBar::AmbientPropertyChanged(DISPID dispID)
{
	m_fDirty = TRUE;
    // do nothing
}
//{IOleControl Support Code}


// IPerPropertyBrowsing
STDMETHODIMP CBar::GetDisplayString( DISPID dispID, BSTR __RPC_FAR *pBstr)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBar::MapPropertyToPage(DISPID dispID, CLSID __RPC_FAR* pClsid)
{
	return PERPROP_E_NOPAGEAVAILABLE;
}
STDMETHODIMP CBar::GetPredefinedStrings(DISPID dispID, CALPOLESTR __RPC_FAR *pCaStringsOut, CADWORD __RPC_FAR* pCaCookiesOut)
{
/*	switch (dispID)
	{
	case DISPID_BACKGROUNDOPTION:
		{
			if (NULL == m_pDesigner || NULL == pCaStringsOut)
				return E_NOTIMPL;

			pCaStringsOut->cElems = 4;
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
				strFlag.LoadString(IDS_BACKGROUNDOPTIONS+nIndex);
				pCaStringsOut->pElems[nIndex] = strFlag.AllocSysString();
				pCaCookiesOut->pElems[nIndex] = dwFlag;
				dwFlag = dwFlag << 1;
			}
		}
		return S_OK;
	}
*/	return E_NOTIMPL;
}	
STDMETHODIMP CBar::GetPredefinedValue( DISPID dispID, DWORD dwCookie, VARIANT __RPC_FAR *pVarOut)
{
	return E_NOTIMPL;
}
STDMETHODIMP CBar::GetType( DISPID dispID,  long *pnType)
{
	switch (dispID)
	{
	case DISPID_BACKGROUNDOPTION:
		*pnType = 2;
		return S_OK;
	}
	return E_NOTIMPL;
}

//
// InterfaceSupportsErrorInfo
//

STDMETHODIMP CBar::InterfaceSupportsErrorInfo( REFIID riid)
{
	return S_OK;
}

//
// IsForeground
//

BOOL CBar::IsForeground(HWND hWnd, BOOL* pIsEnabled)
{
	if (NULL == hWnd)
		hWnd = GetDockWindow();

	if (NULL == hWnd && pIsEnabled)
	{
		*pIsEnabled = FALSE;
		return FALSE;
	}
	HWND hWndParent;
	HWND hWndTopLevel = hWndParent = hWnd;
	while (hWndParent)
	{
		hWndTopLevel = hWndParent;
		hWndParent = GetParent(hWndParent);
	}
	if (pIsEnabled)
	{
		if (0 != (GetWindowLong(hWndTopLevel, GWL_STYLE) & WS_DISABLED))
			*pIsEnabled = FALSE;
		else
			*pIsEnabled = TRUE;
	}
	return (GetActiveWindow() != NULL);
}

//{IConnectionPointContainer Support Code}
STDMETHODIMP CBar::CConnectionPoint::QueryInterface
(
    REFIID riid,
    void **ppvObjOut
)
{
    if (DO_GUIDS_MATCH(riid, IID_IConnectionPoint) || DO_GUIDS_MATCH(riid, IID_IUnknown)) {
        *ppvObjOut = (IConnectionPoint *)this;
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG CBar::CConnectionPoint::AddRef
(
    void
)
{
	return ++m_refCount;
}

ULONG CBar::CConnectionPoint::Release
(
    void
)
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}

STDMETHODIMP CBar::CConnectionPoint::GetConnectionInterface
(
    IID *piid
)
{
	TRACE(1, "GetConnectionInterface\n");
    if (m_bType == SINK_TYPE_EVENT)
        *piid = *EVENTREFIIDOFCONTROL(0);
    else
        *piid = IID_IPropertyNotifySink;

    return S_OK;
}

STDMETHODIMP CBar::CConnectionPoint::GetConnectionPointContainer
(
    IConnectionPointContainer **ppCPC
)
{
    return m_pControl->ExternalQueryInterface(IID_IConnectionPointContainer, (void **)ppCPC);
}

STDMETHODIMP CBar::CConnectionPoint::Advise
(
    IUnknown *pUnk,
    DWORD    *pdwCookie
)
{
	TRACE(1, "IConnectionPoint::Advise\n");
    HRESULT    hr;
    void      *pv;

    CHECK_POINTER(pdwCookie);

    // first, make sure everybody's got what they thinks they got
    //
    if (m_bType == SINK_TYPE_EVENT) {

        // CONSIDER: 12.95 -- this theoretically is broken -- if they do a find
        // connection point on IDispatch, and they just happened to also support
        // the Event IID, we'd advise on that.  this is not awesome, but will
        // prove entirely acceptable short term.
        //
        hr = pUnk->QueryInterface(*EVENTREFIIDOFCONTROL(0), &pv);
        if (FAILED(hr))
            hr = pUnk->QueryInterface(IID_IDispatch, &pv);
    }
    else
        hr = pUnk->QueryInterface(IID_IPropertyNotifySink, &pv);
    RETURN_ON_FAILURE(hr);

    // finally, add the sink.  it's now been cast to the correct type and has
    // been AddRef'd.
    //
    return AddSink(pv, pdwCookie);
}

HRESULT CBar::CConnectionPoint::AddSink
(
    void  *pv,
    DWORD *pdwCookie
)
{
	TRACE(1, "IConnectionPoint::AddSink\n");
    IUnknown **rgUnkNew;
    int        i = 0;

    // we optimize the case where there is only one sink to not allocate
    // any storage.  turns out very rarely is there more than one.
    //
    switch (m_cSinks) {

        case 0:
            ASSERT(!m_rgSinks, "this should be null when there are no sinks");
            m_rgSinks = (IUnknown **)pv;
            break;

        case 1:
            // go ahead and do the initial allocation.  we'll get 8 at a time
            //
            rgUnkNew = (IUnknown **)HeapAlloc(g_hHeap, 0, 8 * sizeof(IUnknown *));
            RETURN_ON_NULLALLOC(rgUnkNew);
            rgUnkNew[0] = (IUnknown *)m_rgSinks;
            rgUnkNew[1] = (IUnknown *)pv;
            m_rgSinks = rgUnkNew;
            break;

        default:
            // if we're out of sinks, then we have to increase the size
            // of the array
            //
            if (!(m_cSinks & 0x7)) {
                rgUnkNew = (IUnknown **)HeapReAlloc(g_hHeap, 0, m_rgSinks, (m_cSinks + 8) * sizeof(IUnknown *));
                RETURN_ON_NULLALLOC(rgUnkNew);
                m_rgSinks = rgUnkNew;
            } else
                rgUnkNew = m_rgSinks;

            rgUnkNew[m_cSinks] = (IUnknown *)pv;
            break;
    }

    *pdwCookie = (DWORD)pv;
    m_cSinks++;
    return S_OK;
}


STDMETHODIMP CBar::CConnectionPoint::Unadvise
(
    DWORD dwCookie
)
{
	TRACE(1, "IConnectionPoint::Unadvise\n");
    IUnknown *pUnk;
    int       x;

    if (!dwCookie)
        return S_OK;

    // see how many sinks we've currently got, and deal with things based
    // on that.
    //
    switch (m_cSinks) {
        case 1:
            // it's the only sink.  make sure the ptrs are the same, and
            // then free things up
            //
            if ((DWORD)m_rgSinks != dwCookie)
                return CONNECT_E_NOCONNECTION;
            m_rgSinks = NULL;
            break;

        case 2:
            // there are two sinks.  go back down to one sink scenario
            //
            if ((DWORD)m_rgSinks[0] != dwCookie && (DWORD)m_rgSinks[1] != dwCookie)
                return CONNECT_E_NOCONNECTION;

            pUnk = ((DWORD)m_rgSinks[0] == dwCookie)
                   ? m_rgSinks[1]
                   : ((DWORD)m_rgSinks[1] == dwCookie) ? m_rgSinks[0] : NULL;

            if (!pUnk) return CONNECT_E_NOCONNECTION;

            HeapFree(g_hHeap, 0, m_rgSinks);
            m_rgSinks = (IUnknown **)pUnk;
            break;

        default:
            // there are more than two sinks.  just clean up the hole we've
            // got in our array now.
            //
            for (x = 0; x < m_cSinks; x++) {
                if ((DWORD)m_rgSinks[x] == dwCookie)
                    break;
            }
            if (x == m_cSinks) return CONNECT_E_NOCONNECTION;
            if (x < m_cSinks - 1) 
                memcpy(&(m_rgSinks[x]), &(m_rgSinks[x + 1]), (m_cSinks -1 - x) * sizeof(IUnknown *));
            else
                m_rgSinks[x] = NULL;
            break;
    }


    // we're happy
    //
    m_cSinks--;
    ((IUnknown *)dwCookie)->Release();
    return S_OK;
}

STDMETHODIMP CBar::CConnectionPoint::EnumConnections
(
    IEnumConnections **ppEnumOut
)
{
	TRACE(1, "IConnectionPoint::EnumConnections\n");
    CONNECTDATA *rgConnectData = NULL;
    int i;

    if (m_cSinks) {
        // allocate some memory big enough to hold all of the sinks.
        //
        rgConnectData = (CONNECTDATA *)HeapAlloc(g_hHeap, 0, m_cSinks * sizeof(CONNECTDATA));
        RETURN_ON_NULLALLOC(rgConnectData);

        // fill in the array
        //
        if (m_cSinks == 1) {
            rgConnectData[0].pUnk = (IUnknown *)m_rgSinks;
            rgConnectData[0].dwCookie = (DWORD)m_rgSinks;
        } else {
            // loop through all available sinks.
            //
            for (i = 0; i < m_cSinks; i++) {
                rgConnectData[i].pUnk = m_rgSinks[i];
                rgConnectData[i].dwCookie = (DWORD)m_rgSinks[i];
            }
        }
    }

    // create yon enumerator object.
    //
    *ppEnumOut = (IEnumConnections *)(IEnumX *)new CEnumX(IID_IEnumConnections,
                        m_cSinks, sizeof(CONNECTDATA), rgConnectData, CopyAndAddRefObject);
    if (!*ppEnumOut) 
	{
        HeapFree(g_hHeap, 0, rgConnectData);
        return E_OUTOFMEMORY;
    }

    return S_OK;
}

CBar::CConnectionPoint::~CConnectionPoint ()
{
	DisconnectAll();
}

void CBar::CConnectionPoint::DisconnectAll()
{
    int x;
    // clean up some memory stuff
    //
    if (!m_cSinks)
        return;
    else if (m_cSinks == 1)
        ((IUnknown *)m_rgSinks)->Release();
    else {
        for (x = 0; x < m_cSinks; x++)
            QUICK_RELEASE(m_rgSinks[x]);
        HeapFree(g_hHeap, 0, m_rgSinks);
    }
	m_cSinks=0;
}

void CBar::CConnectionPoint::DoInvoke(DISPID dispid,BYTE* pbParams,...)
{
	va_list argList;
	va_start(argList, pbParams);
	++inFireEvent;
	DoInvokeV(dispid,pbParams,argList);
	--inFireEvent;
	va_end(argList);
}

void CBar::CConnectionPoint::DoInvokeV(DISPID dispid,BYTE* pbParams,va_list argList)
{
	TRACE1(1, "IConnectionPoint::DoInvoke, %i\n", dispid);
	switch (m_cSinks)
	{
	case 0:
		break;

    case 1:
		{
			HRESULT hResult = DispatchHelper(((IDispatch*)m_rgSinks), dispid, DISPATCH_METHOD, VT_EMPTY, NULL, pbParams, argList); 
			assert(SUCCEEDED(hResult));
		}
		break;

	default:
        for (int iConnection = 0; iConnection < m_cSinks; iConnection++)
            DispatchHelper(((IDispatch*)m_rgSinks[iConnection]), dispid, DISPATCH_METHOD, VT_EMPTY,NULL, pbParams, argList); 
		break;
	}
}

void CBar::CConnectionPoint::DoOnChanged
(
    DISPID dispid
)
{
	TRACE(1, "IConnectionPoint::DoOnChanged\n");
    int iConnection;

    // if we don't have any sinks, then there's nothing to do.
    //
    if (m_cSinks == 0)
        return;
    else if (m_cSinks == 1)
        ((IPropertyNotifySink *)m_rgSinks)->OnChanged(dispid);
    else
        for (iConnection = 0; iConnection < m_cSinks; iConnection++)
            ((IPropertyNotifySink *)m_rgSinks[iConnection])->OnChanged(dispid);
}

BOOL CBar::CConnectionPoint::DoOnRequestEdit
(
    DISPID dispid
)
{
	TRACE(1, "IConnectionPoint::DoOnRequestEdit\n");
    HRESULT hr;
    int     iConnection;

    // if we don't have any sinks, then there's nothing to do.
    //
    if (m_cSinks == 0)
        hr = S_OK;
    else if (m_cSinks == 1)
        hr =((IPropertyNotifySink *)m_rgSinks)->OnRequestEdit(dispid);
    else {
        for (iConnection = 0; iConnection < m_cSinks; iConnection++) {
            hr = ((IPropertyNotifySink *)m_rgSinks[iConnection])->OnRequestEdit(dispid);
            if (hr != S_OK) break;
        }
    }

    return (hr == S_OK) ? TRUE : FALSE;
}

void CBar::BoundPropertyChanged(DISPID dispid)
{
	m_cpPropNotify->DoOnChanged(dispid);
}
//{IConnectionPointContainer Support Code}

void CBar::DataReady(AsyncDownloadItem* pItem, BOOL bSuccessfull)
{
	if (!bSuccessfull)
	{
		m_downloadItem.m_readyState = ASYNC_STATE_ERROR;
		PostMessage(m_hWnd, WM_COMMAND, MAKELONG(eDownloadError, 0), 0);
		return;
	}
	
	m_theErrorObject.m_nAsyncError++;
	LoadData(m_downloadItem.m_pStream, FALSE, NULL);
	FireDataReady();
	m_theErrorObject.m_nAsyncError--;
}

//
// This is used for creating the MDI Button Tools
//

struct MdiChildItem
{
	UINT		 nResStrId;
	short		 nShortcut;
	UINT		 nBitmapId;
	VARIANT_BOOL vbEnabled;
} MdiChildArray[]=
{
	{IDS_SYSRESTORE, 0, IDB_SYSRESTORE,	 VARIANT_TRUE},
	{IDS_SYSMOVE,    0, 0,				 VARIANT_FALSE},
	{IDS_SYSSIZE,    0, 0,				 VARIANT_FALSE},
	{IDS_SYSMIN,     0, IDB_SYSMINIMIZE, VARIANT_TRUE},
	{IDS_SYSMAX,     0, IDB_SYSMAX,		 VARIANT_FALSE},
	{0,				 0, 0,				 VARIANT_TRUE},
	{IDS_SYSCLOSE,   0, IDB_SYSCLOSE,	 VARIANT_TRUE}
};

//
// OnMDIChildPopupMenu
//

void CBar::OnMDIChildPopupMenu(POINT ptScreen, CRect& rcScreen)
{
	CBand* pContextBand = CBand::CreateInstance(NULL);
	if (!pContextBand)
		return;

	pContextBand->SetOwner(this, TRUE);
	pContextBand->put_Type(ddBTPopup);
	pContextBand->put_CreatedBy(ddCBSystem);
	pContextBand->put_Caption(L"SysMDI");
	pContextBand->put_Name(L"SysMDI");
	pContextBand->put_DockingArea(ddDAPopup);

	CTool* pTool;
	HBITMAP hBmp;
	int nSize = sizeof(MdiChildArray)/sizeof(MdiChildItem);
	for (int nTool = 0; nTool < nSize; nTool++)
	{
		pContextBand->m_pTools->CreateTool((ITool**)&pTool);
		if (pTool)
		{
			pTool->tpV1.m_nToolId = eToolIdRestore + nTool;
			LPTSTR szCustom = NULL;
			if (0 != MdiChildArray[nTool].nResStrId)
				szCustom = LoadStringRes(MdiChildArray[nTool].nResStrId);
			else
				pTool->tpV1.m_ttTools = ddTTSeparator;

			if (0 != MdiChildArray[nTool].nBitmapId)
			{
				hBmp = LoadBitmap(g_hInstance, MAKEINTRESOURCE(MdiChildArray[nTool].nBitmapId));
				if (hBmp)
				{
					pTool->put_Bitmap(ddITNormal, (OLE_HANDLE)hBmp);
					pTool->put_MaskBitmap(ddITNormal, (OLE_HANDLE)hBmp);
					DeleteBitmap(hBmp);
				}
			}
			pTool->tpV1.m_vbEnabled = MdiChildArray[nTool].vbEnabled;
			if (ddTTSeparator != pTool->tpV1.m_ttTools)
			{
				MAKE_WIDEPTR_FROMTCHAR(wCustom, szCustom);
				pTool->put_Caption(wCustom);
			}
			pContextBand->m_pTools->InsertTool(-1, pTool, VARIANT_FALSE);
			pTool->Release();
		}
	}
	
	BOOL bDblClicked;
	pContextBand->TrackPopupEx((int)ddPopupMenuLeftAlign, ptScreen.x, ptScreen.y, &rcScreen, &bDblClicked);
	if (bDblClicked)
	{
		HWND hWndActive = (HWND)SendMessage(m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
		if (hWndActive && IsWindow(hWndActive))
			PostMessage(hWndActive, WM_SYSCOMMAND, SC_CLOSE, 0);		
	}
	pContextBand->Release();	
}

void ActiveBarErrorObject::AsyncError(short		     nNumber, 
									  IReturnString* pDescription, 
									  long		     nSCode, 
									  BSTR		     bstrSource, 
									  BSTR		     bstrHelpFile, 
									  long		     nHelpContext, 
									  IReturnBool*   pCancelDisplay)
{
#ifdef _UNICODE
	((CBar*)m_pEvents)->FireError(nNumber, 
								  (ReturnString*)pDescription, 
								  nSCode, 
								  NAMEOFOBJECT(0), 
								  HELPFILEOFOBJECT(0), 
								  nHelpContext, 
								  (ReturnBool*)pCancelDisplay);
#else
	WCHAR Source[256];
	MultiByteToWideChar(CP_ACP, 0, NAMEOFOBJECT(0), -1, Source, 256);
	MAKE_WIDEPTR_FROMANSI(wszHelpFile, HELPFILEOFOBJECT(0));
	((CBar*)m_pEvents)->FireError(nNumber, 
								  (ReturnString*)pDescription, 
								  nSCode, 
								  Source, 
								  wszHelpFile, 
								  nHelpContext, 
								  (ReturnBool*)pCancelDisplay);
#endif
	BSTR bstrDescription;
	pDescription->get_Value(&bstrDescription);
	VARIANT_BOOL vbCancelDisplay;
	pCancelDisplay->get_Value(&vbCancelDisplay);
	if (VARIANT_FALSE == vbCancelDisplay && !((CBar*)m_pEvents)->AmbientUserMode())
	{
		TCHAR szError[1024];
		MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDescription);
		wsprintf(szError, _T("Error %d : %s"), nNumber, szDesc);
		ErrorDialog dlgError(szError);
		dlgError.DoModal(((CBar*)m_pEvents)->m_hWnd);
	}
	SysFreeString(bstrDescription);
}

ErrorTable CBar::m_etErrors[]=
{
	{0x0, IDS_ERR_UNKNOWN,-1}, // this entry needs to be here all the time
	{IDERR_CREATEFILE,IDS_ERR_CANNOTCREATEFILE,csMSaveActiveBar2},
	{IDERR_OPENFILE,IDS_ERR_CANNOTOPENFILE,csMLoadActiveBar2},
	{IDERR_INVALIDFILEFORMAT,IDS_ERR_INVALIDFILEFORMAT,csMLoadActiveBar2},
	{IDERR_BANDNOTFOUND,IDS_ERR_BANDNOTFOUND,csCtlBand},
	{IDERR_CYCLICPOPUP,IDS_ERR_CYCLICPOPUP,csCtlBand},
	{IDERR_OPENINVALIDPOPUP,IDS_ERR_OPENINVALIDPOPUP,csCtlBand},
	{IDERR_BADTOOLSINDEX,IDS_ERR_BADTOOLSINDEX,csCtlTool},
	{IDERR_BADBANDSINDEX,IDS_ERR_BADBANDSINDEX,csCtlBand},
	{IDERR_BADCOLLINDEX ,IDS_ERR_BADCOLLINDEX,-1},  // 10
	{IDERR_INVALIDPROPERTYVALUE,IDS_ERR_INVALIDPROPERTYVALUE,-1},
	{IDERR_INVALIDPICTURE,IDS_ERR_INVALIDPICTURE,csCtlTool},
	{IDERR_DUPLICATEBANDNAME,IDS_ERR_DUPLICATEBANDNAME,csCtlBand},
	{IDERR_DUPLICATETOOLID,IDS_ERR_DUPLICATETOOLID,csCtlTool},
	{IDERR_IMAGEBITMAPNOTSET,IDS_ERR_IMAGEBITMAPNOTSET,csCtlTool},
	{IDERR_MASKINVALIDSIZE,IDS_ERR_MASKINVALIDSIZE,csCtlTool},
	{IDERR_FAILEDTOGETWINDOWHANDLE,IDS_ERR_FAILEDTOGETWINDOWHANDLE,-1},
	{IDERR_INPROPERBORDERSTYLE,IDS_ERR_INPROPERBORDERSTYLE,csCtlTool},
	{IDERR_MUSTBEARRAYOFSTRINGS,IDS_ERR_INPROPERBORDERSTYLE,csCtlTool},
	{IDERR_WRONGTOOLTYPE,IDS_ERR_WRONGTOOLTYPE,csCtlTool},  // 20
	{IDERR_INVALIDSHORTCUT, IDS_ERR_INVALIDSHORTCUT,csCtlTool},
	{IDERR_WINDOWISNOTVALID, IDS_ERR_WINDOWISNOTVALID, -1},
	{IDERR_INVALIDPICTURETYPE, IDS_ERR_INVALIDPICTURETYPE, -1},
	{IDERR_ONLYONETOOLCANHAVEAUTOSIZE, IDS_ERR_ONLYONETOOLCANHAVEAUTOSIZE, csCtlTool},
	{IDERR_SIZERFLAGMUSTBESET, IDS_ERR_SIZERFLAGMUSTBESET, csCtlTool},
	{IDERR_ONLYFORUSEWITHSDIAPPLICATIONS, IDS_ERR_ONLYFORUSEWITHSDIAPPLICATIONS, csCtlActiveBar2},
	{IDERR_NOTONSTATUSBAND, IDS_ERR_NOTONSTATUSBAND, csCtlBand},
	{IDERR_CHILDBANDSHAVETOBENORMAL, IDS_ERR_CHILDBANDSHAVETOBENORMAL, csCtlBand},
	{IDERR_OUTOFMEMORY, IDS_ERR_OUTOFMEMORY, csCtlTool},
	{IDERR_NULLNAMENOTALLOWED, IDS_ERR_NULLNAMENOTALLOWED, csCtlBand},
	{IDERR_COMBOBOXISREADONLY, IDS_ERR_COMBOBOXISREADONLY, csCtlTool}  // 30
};

CBar::BarPropV1::BarPropV1()
{
	m_asChildren  = ddASProportional;
	m_mfsMenuFont = ddMSSystem;
	m_msMenuStyle = ddMSAnimateSlide;
	m_ctCustomize = ddCTCustomizeStop;
	m_pmMenus = ddPMDisabled;

	m_vbDisplayKeysInToolTip = VARIANT_FALSE;
	m_vbWhatsThisHelpMode = VARIANT_FALSE;
	m_vbDisplayToolTips = VARIANT_TRUE;
	m_vbAlignToForm = VARIANT_FALSE;
	m_vbLargeIcons = VARIANT_FALSE;
	m_vbDuplicateToolIds = VARIANT_FALSE;
	
	m_ocForeground = 0x80000000 + COLOR_BTNTEXT;
	m_ocBackground = 0x80000000 + COLOR_BTNFACE;
	m_ocHighLight = 0x80000000 + COLOR_BTNHIGHLIGHT;
	m_ocShadow = 0x80000000 + COLOR_BTNSHADOW;
	m_oc3DDarkShadow = 0x80000000 + COLOR_3DDKSHADOW;
	m_oc3DLight = 0x80000000 + COLOR_3DLIGHT;
	
	m_vbAutoUpdateStatusbar = VARIANT_FALSE;
	m_vbUserDefinedCustomization = VARIANT_FALSE;
	m_vbFireDblClickEvent = VARIANT_FALSE;

	m_dwBackgroundOptions = ddBODockAreas|ddBOFloat|ddBOPopups|ddBOClientArea;

	m_dwToolIdentity = 1;

	m_vbUseUnicode = VARIANT_FALSE;
	m_vbXPLook = VARIANT_FALSE;
}

//
// Palette
//

HPALETTE& CBar::Palette()
{
	return m_pImageMgr->m_hPal;
}


//
// GetMenuFont
//
// The menu font is calculated according to the system metrics. If the menu font 
// is nontrue type a truetype font is substituted since the menu text will have 
// to be rotated 90deg.
//

HFONT CBar::GetMenuFont(BOOL bVertical, BOOL bForce)
{
	BOOL bResult;
	if (bForce)
	{
		if (bVertical)
		{
			bResult = DeleteFont(m_hFontMenuVert);
			m_hFontMenuVert = NULL;
		}
		else
		{
			bResult = DeleteFont(m_hFontMenuHorz);
			m_hFontMenuHorz = NULL;
		}
	}
	if (bVertical)
	{
		if (m_hFontMenuVert)
			return m_hFontMenuVert;
	}
	else
	{
		if (m_hFontMenuHorz)
			return m_hFontMenuHorz;
	}
	
	LOGFONT lfMenuFont;
	if (ddMSSystem == bpV1.m_mfsMenuFont)
	{
		if (g_fSysWin95Shell)
		{
			NONCLIENTMETRICS nm;
			nm.cbSize = sizeof(NONCLIENTMETRICS);
			if (!SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, &nm, 0))
				GetObject(GetStockFont(SYSTEM_FONT), sizeof(LOGFONT), &lfMenuFont);
			else
				lfMenuFont = nm.lfMenuFont;
		}
		else
			GetObject(GetStockFont(SYSTEM_FONT), sizeof(LOGFONT), &lfMenuFont);
	}
	else
	{
		HFONT hFontMy = m_fhFont.GetFontHandle();
		GetObject(hFontMy, sizeof(lfMenuFont), &lfMenuFont);
	}

	if (bVertical)
	{
		lfMenuFont.lfEscapement = -900;
		lfMenuFont.lfOrientation = -900;
		lfMenuFont.lfClipPrecision = CLIP_LH_ANGLES;
		lfMenuFont.lfOutPrecision = OUT_TT_ONLY_PRECIS;
	}
	
	HFONT hFontNew = CreateFontIndirect(&lfMenuFont);
	if (bVertical)
		m_hFontMenuVert = hFontNew;
	else
		m_hFontMenuHorz = hFontNew;
	return hFontNew;
}

HFONT CBar::GetControlFont()
{
	if (!m_cachedControlFont)
		m_cachedControlFont=m_fhControl.GetFontHandle();

	return m_cachedControlFont;
}

//
// GetSmallFont
//

HFONT CBar::GetSmallFont(BOOL bVert)
{
	if (!bVert && m_hFontSmall)
		return m_hFontSmall;

	if (bVert && m_hFontSmallVert)
		return m_hFontSmallVert;

	LOGFONT lf;
	memset(&lf, 0, sizeof(LOGFONT));
	MAKE_ANSIPTR_FROMWIDE(szFontName, TABS_NORM_FONT);
	lstrcpy(lf.lfFaceName, szFontName);

#ifndef JAPBUILD
	if (GetGlobals().m_bUseDBCSUI)
		lstrcpy(lf.lfFaceName, _T("System"));
#else
	lf.lfCharSet = SHIFTJIS_CHARSET;
#endif

	if (bVert)
	{
		MAKE_ANSIPTR_FROMWIDE(szFontName, TABS_ROT_FONT);
		lf.lfHeight = TABS_ROT_SIZE;
		if (!g_fSysWinNT)
			lstrcpy(lf.lfFaceName, szFontName);
		lf.lfEscapement = -900;
		lf.lfOrientation = -900;
		lf.lfOutPrecision = OUT_TT_ONLY_PRECIS;
	}
	else
		lf.lfHeight = TABS_NORM_SIZE;	

	HFONT hFontNew = CreateFontIndirect(&lf);
	if (bVert)
		m_hFontSmallVert = hFontNew;
	else
		m_hFontSmall = hFontNew;

	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		HFONT hFontOld = SelectFont(hDC, hFontNew);
		TEXTMETRIC tm;
		GetTextMetrics(hDC, &tm);
		if (bVert)
			m_nSmallFontHeightVert = tm.tmHeight;
		else
			m_nSmallFontHeight = tm.tmHeight;

		SelectFont(hDC, hFontOld);
		ReleaseDC(NULL, hDC);
	}
	return m_hFontSmall;
}

//
// GetChildBandFont
//

HFONT CBar::GetChildBandFont(BOOL bVert)
{
	if (!bVert)
	{
		if (m_hFontChildBand)
			return m_hFontChildBand;

		m_hFontChildBand = m_fhChildBand.GetFontHandle();
		HDC hDC = GetDC(NULL);
		if (hDC)
		{
			HFONT hFontOld = SelectFont(hDC, m_hFontChildBand);
			TEXTMETRIC tm;
			GetTextMetrics(hDC, &tm);
			m_nChildBandFontHeight = tm.tmHeight + tm.tmExternalLeading;
			SelectFont(hDC, hFontOld);
			ReleaseDC(NULL, hDC);
		}
		return m_hFontChildBand;
	}

	if (m_hFontChildBandVert)
		return m_hFontChildBandVert;

	HFONT hFont = m_fhChildBand.GetFontHandle();

	LOGFONT lfChildBandVert;
	int nResult = GetObject(hFont, sizeof(LOGFONT), &lfChildBandVert); 
	if (!g_fSysWinNT)
	{
		MAKE_TCHARPTR_FROMWIDE(szFontName, TABS_ROT_FONT);
		lstrcpy(lfChildBandVert.lfFaceName, szFontName);
	}
	lfChildBandVert.lfEscapement = -900;
	lfChildBandVert.lfOrientation = -900;
	lfChildBandVert.lfOutPrecision = OUT_TT_ONLY_PRECIS;
	m_hFontChildBandVert = CreateFontIndirect(&lfChildBandVert);

	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		HFONT hFontOld = SelectFont(hDC, m_hFontChildBandVert);
		TEXTMETRIC tm;
		GetTextMetrics(hDC, &tm);
		m_nChildBandFontHeightVert = tm.tmHeight + tm.tmExternalLeading;
		SelectFont(hDC, hFontOld);
		ReleaseDC(NULL, hDC);
	}
	return m_hFontChildBandVert;
}

//
// FindSubBand
//

CBand* CBar::FindSubBand(BSTR bstrName)
{
	if (NULL == bstrName)
		return NULL;

	CBand* pBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (pBand->m_bstrName && 0 == wcscmp(pBand->m_bstrName, bstrName) && (ddBTPopup == pBand->bpV1.m_btBands || (ddBTNormal == pBand->bpV1.m_btBands && ddDAPopup == pBand->bpV1.m_daDockingArea)))
			return pBand;
	}
	return NULL;
}

//
// FindBand
//

CBand* CBar::FindBand(BSTR bstrName)
{
	if (NULL == bstrName)
		return NULL;

	CBand* pBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (pBand->m_bstrName && 0 == wcscmp(pBand->m_bstrName, bstrName))
			return pBand;
	}
	return NULL;
}

//
// GetMenuBand
//

CBand* CBar::GetMenuBand()
{
	CBand* pBand;
	int nBandCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nBandCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if ((ddBTChildMenuBar == pBand->bpV1.m_btBands || ddBTMenuBar == pBand->bpV1.m_btBands) && VARIANT_TRUE == pBand->bpV1.m_vbVisible)
			return pBand;
	}
	return NULL;
}

//
// GetPopupBands
//

void CBar::GetPopupBands(CALPOLESTR& caPopups)
{
	HRESULT hResult;
	CBand* pBand;
	BSTR bstrName;
	TypedArray<BSTR> aBSTRs;
	int nBandCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nBandCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (ddBTPopup == pBand->bpV1.m_btBands || (ddBTNormal == pBand->bpV1.m_btBands && ddDAPopup == pBand->bpV1.m_daDockingArea))
		{
			if (SUCCEEDED(pBand->get_Name(&bstrName)))
			{
				hResult = aBSTRs.Add(bstrName);
				if (FAILED(hResult))
				{
					SysFreeString(bstrName);
					assert(FALSE);
					return;
				}
			}
		}
	}

	caPopups.cElems = aBSTRs.GetSize();
	if (caPopups.cElems > 0)
	{
		caPopups.pElems = (LPOLESTR*)CoTaskMemAlloc(sizeof(LPOLESTR*)*caPopups.cElems);
		for (UINT nBand = 0; nBand < caPopups.cElems; nBand++)
			caPopups.pElems[nBand] = aBSTRs.GetAt(nBand);
	}
}

void CBar::EnterMenuLoop(HWND hWndParent, UINT nKey)
{
	CBand* pMenuBand = GetMenuBand();
	if (NULL == pMenuBand)
		return;
	
	if (0 != pMenuBand->m_pTools->GetToolCount())
	{
		pMenuBand->EnterMenuLoop(0, NULL, CBand::ENTERMENULOOP_KEY, nKey);
	}
}

//
// StatusBandEnter
//

void CBar::StatusBandEnter(CTool* pToolIn)
{
	if (m_bStatusBandLock)
		return;

	m_bStatusBandLock = TRUE;

	if (NULL == m_pStatusBand || VARIANT_FALSE == bpV1.m_vbAutoUpdateStatusbar)
		return;

	short nCount = m_pStatusBand->m_pTools->GetToolCount();
	m_pvbStatusBandToolsVisible = new VARIANT_BOOL[nCount];
	if (m_pvbStatusBandToolsVisible)
	{
		CTool* pTool;
		for (int nTool = 0; nTool < nCount; nTool++)
		{
			pTool = m_pStatusBand->m_pTools->GetTool(nTool);
			if (NULL == pTool)
				continue;

			m_pvbStatusBandToolsVisible[nTool] = pTool->tpV1.m_vbVisible;
			pTool->tpV1.m_vbVisible = VARIANT_FALSE;
		}
	}

	//
	// I need to create a tool and stick it at the end.
	//

	BSTR bstrTemp = SysAllocString(L"miUpdate");
	assert(bstrTemp);
	HRESULT hResult = m_pStatusBand->m_pTools->Add(eToolIdStatusbar, bstrTemp, (Tool**)&m_pStatusBandTool);
	SysFreeString(bstrTemp);
	if (FAILED(hResult))
		return;

	if (pToolIn)
		hResult = m_pStatusBandTool->put_Caption(pToolIn->m_bstrDescription);

	m_pStatusBandTool->tpV1.m_ttTools = ddTTLabel;
	m_pStatusBandTool->tpV1.m_tsStyle = ddSText;
	m_pStatusBandTool->tpV1.m_taTools = ddALeftCenter;
	m_pStatusBandTool->tpV1.m_cpTools = ddCPLeft;
	m_pStatusBandTool->tpV1.m_asAutoSize = ddTASSpringWidth;
	CRect rcBound, rcReturn;
	m_pStatusBand->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, TRUE);
	m_pStatusBand->Refresh();
}

//
// StatusBandUpdate
//

void CBar::StatusBandUpdate(CTool* pToolIn)
{
	if (NULL == m_pStatusBand || NULL == m_pStatusBandTool || NULL == pToolIn || VARIANT_FALSE == bpV1.m_vbAutoUpdateStatusbar)
		return;

	m_pStatusBandTool->put_Caption(pToolIn->m_bstrDescription);
	CRect rcBound, rcReturn;
	m_pStatusBand->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, TRUE);
	m_pStatusBand->Refresh();
}

//
// StatusBandExit
//

void CBar::StatusBandExit()
{
	if (NULL == m_pStatusBand || VARIANT_FALSE == bpV1.m_vbAutoUpdateStatusbar)
		return;

	if (m_pStatusBandTool)
	{
		m_pStatusBand->m_pTools->DeleteTool(m_pStatusBandTool);
		m_pStatusBandTool->Release();
		m_pStatusBandTool = NULL;
	}

	short nCount = m_pStatusBand->m_pTools->GetToolCount();
	if (nCount > 0 && m_pvbStatusBandToolsVisible)
	{
		CTool* pTool;
		for (int nTool = 0; nTool < nCount; nTool++)
		{
			pTool = m_pStatusBand->m_pTools->GetTool(nTool);
			if (NULL == pTool)
				continue;

			pTool->tpV1.m_vbVisible = m_pvbStatusBandToolsVisible[nTool];
		}
	}

	delete [] m_pvbStatusBandToolsVisible;
	m_pvbStatusBandToolsVisible = NULL;
	CRect rcBound, rcReturn;
	m_pStatusBand->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, TRUE);
	m_pStatusBand->Refresh();
	m_bStatusBandLock = FALSE;
}

//
// NeedsCycleShutDown
//

BOOL CBar::NeedsCycleShutDown(CTool* pTool, TypedArray<CTool*>* pfaToolsToBeDropped)
{
	// If the tool being dropped will cause a cyclic reference don't open subband
	if (NULL == pfaToolsToBeDropped)
		return FALSE;
		
	CTool* pDropTool;
	int nDropToolCount = pfaToolsToBeDropped->GetSize();
	for (int nTool = 0; nTool < nDropToolCount; nTool++)
	{
		pDropTool = pfaToolsToBeDropped->GetAt(nTool);
		if (pTool && pDropTool->tpV1.m_nToolId == pTool->tpV1.m_nToolId)
			return TRUE;
	
		if ((m_pPopupRoot && -1 != m_pPopupRoot->m_nPopupIndex))
		{
			CTool* pTargetBandTool = m_pPopupRoot->m_pTools->GetTool(m_pPopupRoot->m_nPopupIndex);
			if (pTargetBandTool->HasSubBand())
			{
				CBand* pTargetBand = FindBand(pTargetBandTool->m_bstrSubBand);
				CBand* pParentOfProblemBand = NULL;
				if (pTargetBand && 
					pTargetBand->ContainsSubBandOfToolTree(pDropTool, &pParentOfProblemBand))
				{
					if (NULL == pParentOfProblemBand)
					{
						TRACE(1, "Problem Band = NULL\n");
						m_pPopupRoot->SetPopupIndex(-1);
					}
					else
					{
						TRACE1(1, "Problem Band = %ls\n", pParentOfProblemBand);
						pParentOfProblemBand->SetPopupIndex(-1);
					}
					TRACE(1, "Cycle\n\n")
					return TRUE;
				}
			}
		}
	}
	return FALSE;
}

//
// DrawEdge
//

BOOL CBar::DrawEdge(HDC hDC, const CRect& rcEdge, UINT nEdge, UINT nFlags)
{
	BOOL bResult = TRUE;
	CRect rc;
	switch (nEdge)
	{
	case BDR_RAISEDOUTER:
		rc = rcEdge;
		rc.bottom = rc.top + 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_cr3DLight : m_crXP3DLight);

		rc = rcEdge;
		rc.right = rc.left + 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_cr3DLight : m_crXP3DLight);

		rc = rcEdge;
		rc.top = rc.bottom - 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_cr3DDarkShadow : m_crXP3DDarkShadow);

		rc = rcEdge;
		rc.left = rc.right - 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_cr3DDarkShadow : m_crXP3DDarkShadow);
		break;

	case BDR_RAISEDINNER:
		rc = rcEdge;
		rc.bottom = rc.top + 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_crHighLight : m_crXPHighLight);

		rc = rcEdge;
		rc.right = rc.left + 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_crHighLight : m_crXPHighLight);

		rc = rcEdge;
		rc.top = rc.bottom - 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_crShadow : m_crXPShadow);

		rc = rcEdge;
		rc.left = rc.right - 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_crShadow : m_crXPShadow);
		break;

	case BDR_SUNKENOUTER:
		rc = rcEdge;
		rc.bottom = rc.top + 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_crShadow : m_crXPShadow);
		
		rc = rcEdge;
		rc.right = rc.left + 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_crShadow : m_crXPShadow);
		
		rc = rcEdge;
		rc.top = rc.bottom - 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_crHighLight : m_crXPHighLight);

		rc = rcEdge;
		rc.left = rc.right - 1;
		bResult = FillSolidRect(hDC, rc, VARIANT_FALSE == bpV1.m_vbXPLook ? m_crHighLight : m_crXPHighLight);
		break;

	default:
		assert(FALSE);
		break;
	}
	return bResult;
}

//
// BuildAccelators
//

void CBar::BuildAccelators(CBand* pBand, TypedArray<WORD>* pfaChar)
{
	if (NULL == pBand)
		return;

	HRESULT hResult;
	WCHAR  nKey;
	CTool* pTool;
	m_mapMenuBar.RemoveAll();
	int nCount = pBand->m_pTools->GetToolCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pTool = pBand->m_pTools->GetTool(nBand);
		if (pTool->CheckForAlt(nKey))
		{
			if (_istlower(nKey))
				nKey = _toupper(nKey);
			m_mapMenuBar.SetAt(nKey, pTool);
			if (pfaChar)
			{
				hResult = pfaChar->Add(nKey);
				if (FAILED(hResult))
				{
					assert(FALSE);
					return;
				}
			}
		}
	}
}

//
// DoMenuAccelator
//

BOOL CBar::DoMenuAccelator(WORD nKey)
{
	CTool* pTool;
	if (nKey >= VK_NUMPAD0 && nKey <= VK_NUMPAD9)
		nKey = nKey - '0'; 

	if (!m_mapMenuBar.Lookup(nKey, pTool))
		return FALSE;

	if (VARIANT_FALSE == pTool->tpV1.m_vbEnabled)
		return TRUE;

	CBand* pBand = GetMenuBand();
	if (NULL == pBand)
		return FALSE;

	if (m_bMenuLoop || m_bClickLoop)
	{
		// Send End Event
//		NotifyWinEvent(EVENT_SYSTEM_MENUEND, m_hWnd, OBJID_MENU, CHILDID_SELF);
	}
	// Send Start Event
//	NotifyWinEvent(EVENT_OBJECT_FOCUS, m_hWnd, OBJID_MENU, CHILDID_SELF);

	pBand->EnterMenuLoop(0, NULL, CBand::ENTERMENULOOP_KEY, nKey);
	
	// Send End Event
//	NotifyWinEvent(EVENT_SYSTEM_MENUEND, m_hWnd, OBJID_MENU, CHILDID_SELF);
	return TRUE;
}

//
// CacheTexture
//

void CBar::CacheTexture()
{
	BOOL bResult;
	if (m_hBrushTexture)
	{
		bResult = DeleteBrush(m_hBrushTexture);
		assert(bResult);
	}

	m_hBrushTexture = NULL;
	m_bHasTexture = FALSE;

	if (NULL == m_phPict.m_pPict)
		return;

	short nType = m_phPict.GetType();
	if (!(PICTYPE_UNINITIALIZED == nType || PICTYPE_NONE == nType))
	{
		HBITMAP hBmp = GetTextureBitmap();
		if (hBmp)
		{
			if (g_fSysWinNT)
			{
				m_hBrushTexture = CreatePatternBrush(hBmp);
				if (m_hBrushTexture)
					m_bHasTexture = TRUE;
			}
			else
			{
				BITMAP bmInfo;
				GetObject(hBmp,sizeof(BITMAP),&bmInfo);
				m_nTextureWidth = bmInfo.bmWidth;
				m_nTextureHeight = bmInfo.bmHeight;
				m_bHasTexture = TRUE;
			}
		}
	}
}	

//
// CacheColors
//

void CBar::CacheColors()
{
	BOOL bResult;
	OleTranslateColor(bpV1.m_ocForeground, NULL, &m_crForeground);
	if (m_hBrushForeColor)
	{
		bResult = DeleteBrush(m_hBrushForeColor);
		assert(bResult);
	}

	m_hBrushForeColor = CreateSolidBrush(m_crForeground);
	
	OleTranslateColor(bpV1.m_ocBackground, NULL, &m_crBackground);
	if (m_hBrushBackColor)
	{
		bResult = DeleteBrush(m_hBrushBackColor);
		assert(bResult);
	}
	
	m_hBrushBackColor = CreateSolidBrush(m_crBackground);

	OleTranslateColor(bpV1.m_ocHighLight, NULL, &m_crHighLight);
	if (m_hBrushHighLightColor)
	{
		bResult = DeleteBrush(m_hBrushHighLightColor);
		assert(bResult);
	}
	
	m_hBrushHighLightColor = CreateSolidBrush(m_crHighLight);
	
	OleTranslateColor(bpV1.m_ocShadow, NULL, &m_crShadow);
	if (m_hBrushShadowColor)
	{
		bResult = DeleteBrush(m_hBrushShadowColor);
		assert(bResult);
	}
	
	m_hBrushShadowColor = CreateSolidBrush(m_crShadow);

	OleTranslateColor(bpV1.m_oc3DDarkShadow, NULL, &m_cr3DDarkShadow);
	OleTranslateColor(bpV1.m_oc3DLight, NULL, &m_cr3DLight);

	if (GetAmbientProperty(DISPID_AMBIENT_BACKCOLOR, VT_I4, &m_ocAmbientBackColor))
		OleTranslateColor(m_ocAmbientBackColor, NULL, &m_crAmbientBackColor);

	m_crMDIMenuBackground = RGB((10 * GetRValue(m_crBackground)) / 9, (10 * GetGValue(m_crBackground)) / 9, (10 * GetBValue(m_crBackground)) / 9);

	HDC hDC = GetDC(m_hWnd);
	if (hDC)
	{
//		m_crMDIMenuBackground = NewColor(hDC, m_crBackground, 60);

		m_crXPFloatCaptionBackground = NewColor(hDC, GetSysColor(COLOR_APPWORKSPACE), 20);

		m_crXPBackground = GetSysColor(COLOR_BTNFACE);
		
		m_crXPMRUBackground = NewColor(hDC, m_crXPBackground, -40);

		m_crXPBandBackground = NewColor(hDC, m_crXPBackground, 20);
		
		m_crXPMenuBackground = NewColor(hDC, m_crXPBackground, 86);
		m_crXPMenuSelectedBorderColor = GetSysColor(COLOR_3DDKSHADOW);
		
		m_crXPSelectedBorderColor = GetSysColor(COLOR_HIGHLIGHT);

		m_crXPSelectedColor = NewColor(hDC, GetSysColor(COLOR_HIGHLIGHT), 68);

		int nBBRed = GetRValue(m_crXPBandBackground);
		int nBBGreen = GetGValue(m_crXPBandBackground);
		int nBBBlue = GetBValue(m_crXPBandBackground);

		int nSRed = GetRValue(m_crXPSelectedColor);
		int nSGreen = GetGValue(m_crXPSelectedColor);
		int nSBlue = GetBValue(m_crXPSelectedColor);

//		if ((nSRed > nBBRed - 10 && nSRed <= nBBRed + 10) && 
//			(nSGreen > nBBGreen - 10 && nSGreen <= nBBGreen + 10) && 
//			(nSBlue > nBBBlue - 10 && nSBlue <= nBBBlue + 10))
//		{
//			m_crXPSelectedColor = RGB(255, 255, 255);
//		}

		m_crXPPressed = NewColor(hDC, m_crXPSelectedBorderColor, 50);
		
		m_crXPSelectedCrossColor = NewColor(hDC, m_crXPSelectedBorderColor, 40);
		
		m_crXPCheckedColor = NewColor(hDC, m_crXPSelectedColor, 50);
		
		m_crXPMenuBorderShadow = NewColor(hDC, m_crXPSelectedBorderColor, 80);

		m_crXPHighLight = GetSysColor(COLOR_BTNHIGHLIGHT);
		m_crXPShadow = GetSysColor(COLOR_BTNSHADOW);
		m_crXP3DDarkShadow = GetSysColor(COLOR_3DDKSHADOW);
		m_crXP3DLight = GetSysColor(COLOR_3DLIGHT);

		ReleaseDC(m_hWnd, hDC);
	}
}

//
// ResetMenuFonts
//

void CBar::ResetMenuFonts()
{
	BOOL bResult;
	if (m_hFontMenuVert)
	{
		bResult = DeleteFont(m_hFontMenuVert);
		assert(bResult);
		m_hFontMenuVert = NULL;
	}
	if (m_hFontMenuHorz)
	{
		bResult = DeleteFont(m_hFontMenuHorz);
		assert(bResult);
		m_hFontMenuHorz = NULL;
	}
}

//
// FillTexture
//

void CBar::FillTexture(HDC hDC, const CRect& rc)
{
	if (g_fSysWinNT)
		FillRect(hDC, &rc, m_hBrushTexture);
	else
		FillRgnBitmap(hDC, 0, GetTextureBitmap(), rc, m_nTextureWidth, m_nTextureHeight);
}

//
// GetTextureBitmap
//

HBITMAP CBar::GetTextureBitmap()
{
	HBITMAP hBmp;
	m_phPict.m_pPict->get_Handle((OLE_HANDLE*)&hBmp);
	return hBmp;
}

//
// ModifyTool
//
// Generic function to visit each tool in a all of the tool collections and calls a user 
// defined function which can modify the tools value.
//

void CBar::ModifyTool(ULONG nToolId, void* pData, PFNMODIFYTOOL pToolFunction)
{
	try
	{
		// Look through all bands and tools 
		CBand* pBand;
		CTool* pTool;
		CBand* pChildBand;
		int nTool, nToolCount;
		int nPage, nPageCount;
		BOOL bBandChanged = FALSE;

		int nBandCount = m_pBands->GetBandCount();
		for (int nBand = 0; nBand < nBandCount; nBand++)
		{
			pBand = m_pBands->GetBand(nBand);

			if (ddCBSNone == pBand->bpV1.m_cbsChildStyle)
			{
				nToolCount = pBand->m_pTools->GetToolCount();
				for (nTool = 0; nTool < nToolCount; nTool++)
				{
					pTool = pBand->m_pTools->GetTool(nTool);
					if (NULL == pTool)
						continue;

					if (pTool->tpV1.m_nToolId == nToolId)
					{
						(*pToolFunction)(pBand, pTool, pData);
						bBandChanged = TRUE;
					}
					else if (-1 == nToolId)
						(*pToolFunction)(pBand, pTool, pData);
				}
			}
			else
			{
				CChildBands* pChildBands = pBand->m_pChildBands;
				if (NULL == pChildBands)
					continue;

				nPageCount = pChildBands->GetChildBandCount();
				for (nPage = 0; nPage < nPageCount; nPage++)
				{
					pChildBand = pChildBands->GetChildBand(nPage);
					if (NULL == pChildBand)
						continue;

					nToolCount = pChildBand->m_pTools->GetToolCount();
					for (nTool = 0; nTool < nToolCount; nTool++)
					{
						pTool = pChildBand->m_pTools->GetTool(nTool);
						if (NULL == pTool)
							continue;

						if (pTool->tpV1.m_nToolId == nToolId)
						{
							(*pToolFunction)(pBand, pTool, pData);
							bBandChanged = TRUE;
						}
						else if (-1 == nToolId)
							(*pToolFunction)(pBand, pTool, pData);
					}
				}
			}
			if (bBandChanged)
				pBand->Refresh();
		}
		
		nToolCount = m_pTools->GetToolCount();
		for (nTool = 0; nTool < nToolCount; nTool++)
		{
			pTool = m_pTools->GetTool(nTool);
			if (pTool->tpV1.m_nToolId == nToolId)
				(*pToolFunction)(NULL, pTool, pData);
			else if (-1 == nToolId)
				(*pToolFunction)(NULL, pTool, pData);
		}
	}
	catch (...)
	{
	}
}

//
// Custom
//

struct Custom
{
	LPDISPATCH pCustomTool;
	DWORD      dwFlag;
	static void SetCustom(CBand* pBand, CTool* pTool, void* pData);
};

//
// SetCustom
//

void Custom::SetCustom(CBand* pBand, CTool* pTool, void* pData)
{
	Custom* pCustom = (Custom*)pData;
	if (pTool->m_pDispCustom)
		pTool->m_pDispCustom->Release();

	if (pCustom->pCustomTool)
	{
		pCustom->pCustomTool->AddRef();
		pTool->m_pDispCustom = pCustom->pCustomTool;
		pTool->m_dwCustomFlag = pCustom->dwFlag;
	}
}

//
// SetCustomInterface
//

void CBar::SetCustomInterface(ULONG nToolId, LPDISPATCH pCustomTool, DWORD dwFlag)
{
	Custom theCustom;
	theCustom.pCustomTool = pCustomTool;
	theCustom.dwFlag = dwFlag;
	ModifyTool(nToolId, &theCustom, Custom::SetCustom);

}

//
// ToolState
//

struct ToolState
{
	CBar::ToolStateOps tsOperation;
	VARIANT_BOOL vbNewVal;
	BOOL bRuntime;
	static void Change(CBand* pBand, CTool* pTool, void* pData);
};

//
// Change
//

void ToolState::Change(CBand* pBand, CTool* pTool, void* pData)
{
	ToolState* pChangeToolState = (ToolState*)pData;
	switch(pChangeToolState->tsOperation)
	{
	case CBar::TSO_ENABLE:
		if (pTool->tpV1.m_vbEnabled != pChangeToolState->vbNewVal)
		{
			pTool->tpV1.m_vbEnabled = pChangeToolState->vbNewVal;
			if (pBand && pChangeToolState->bRuntime)
				pBand->ToolNotification(pTool, TNF_VIEWCHANGED);
		}
		break;

	case CBar::TSO_VISIBLE:
		if (pTool->tpV1.m_vbVisible != pChangeToolState->vbNewVal)
			pTool->tpV1.m_vbVisible = pChangeToolState->vbNewVal;
		break;

	case CBar::TSO_CHECK:
		if (pTool->tpV1.m_vbChecked != pChangeToolState->vbNewVal)
		{
			pTool->tpV1.m_vbChecked = pChangeToolState->vbNewVal;
			if (pBand && pChangeToolState->bRuntime)
				pBand->ToolNotification(pTool, TNF_VIEWCHANGED); 
		}
		break;
	}
}

//
// State Transition Control
//

void CBar::ChangeToolState(ULONG nToolId, ToolStateOps op, VARIANT_BOOL vbNewVal)
{
	ToolState theState;
	theState.vbNewVal = vbNewVal;
	theState.tsOperation = op;
	theState.bRuntime = AmbientUserMode();
	ModifyTool(nToolId, &theState, ToolState::Change);
}

void CBar::EnableButton(ULONG nToolId, BOOL bEnable)
{
	ChangeToolState(nToolId, TSO_ENABLE, bEnable);
}

void CBar::SetVisibleButton(ULONG nToolId, BOOL vbVisible)
{
	ChangeToolState(nToolId, TSO_VISIBLE, vbVisible);
}

void CBar::CheckButton(ULONG nToolId, BOOL bChecked)
{
	ChangeToolState(nToolId, TSO_CHECK, bChecked);
}

struct ToolText
{
	BSTR bstrNew;
	static void Change(CBand* pBand, CTool* pTool, void* pData);
};

void ToolText::Change(CBand* pBand, CTool* pTool, void* pData)
{
	ToolText* pToolText = (ToolText*)pData;
	SysFreeString(pTool->m_bstrText);
	pTool->m_bstrText = SysAllocString(pToolText->bstrNew);
}

void CBar::ChangeToolText(ULONG nToolId, BSTR bstrNewText)
{
	ToolText theToolText;
	theToolText.bstrNew = bstrNewText;
	ModifyTool(nToolId, &theToolText, ToolText::Change);
}

struct ToolShortCut
{
	CToolShortCut* pToolShortCut;
	static void Change(CBand* pBand, CTool* pTool, void* pData);
};

void ToolShortCut::Change(CBand* pBand, CTool* pTool, void* pData)
{
	ToolShortCut* pToolShortCut = (ToolShortCut*)pData;
	if (pToolShortCut)
		pToolShortCut->pToolShortCut->CopyTo(pTool->m_scTool);
}

void CBar::SetToolShortCut(ULONG nToolId, CToolShortCut* pToolShortCutIn)
{
	ToolShortCut theToolShortCut;
	theToolShortCut.pToolShortCut = pToolShortCutIn;
	if (theToolShortCut.pToolShortCut)
		ModifyTool(nToolId, &theToolShortCut, ToolShortCut::Change);
}

struct RemoveShortCutFromTools
{
	ShortCutStore* pShortCutStore;
	static void Change(CBand* pBand, CTool* pTool, void* pData);
};

void RemoveShortCutFromTools::Change(CBand* pBand, CTool* pTool, void* pData)
{
	RemoveShortCutFromTools* pRemoveShortCutFromTools = (RemoveShortCutFromTools*)pData;
	if (pRemoveShortCutFromTools)
		pTool->m_scTool.Remove(pRemoveShortCutFromTools->pShortCutStore);
}

void CBar::RemoveShortCutFromAllTools(ULONG nToolId, ShortCutStore* pShortCutStoreIn)
{
	RemoveShortCutFromTools theRemoveShortCutFromTools;
	theRemoveShortCutFromTools.pShortCutStore = pShortCutStoreIn;
	if (theRemoveShortCutFromTools.pShortCutStore)
		ModifyTool(nToolId, &theRemoveShortCutFromTools, RemoveShortCutFromTools::Change);
}

struct ToolComboList
{
	IComboList* pToolComboList;
	static void Change(CBand* pBand, CTool* pTool, void* pData);
};

void ToolComboList::Change(CBand* pBand, CTool* pTool, void* pData)
{
	ToolComboList* pToolComboList = (ToolComboList*)pData;
	IComboList* pComboList = pTool->GetComboList();
	if (pComboList)
		pToolComboList->pToolComboList->CopyTo(&pComboList);
}

void CBar::SetToolComboList(ULONG nToolId, IComboList* pToolComboListIn)
{
	ToolComboList theToolComboList;
	theToolComboList.pToolComboList = pToolComboListIn;
	if (theToolComboList.pToolComboList)
		ModifyTool(nToolId, &theToolComboList, ToolComboList::Change);
}

//
// GetCanonicalTool 
//

void CBar::GetCanonicalTool(ULONG nToolId, CTool*& pTool)
{
	int nToolCount = m_pTools->GetToolCount();
	for (int nTool = 0; nTool < nToolCount; nTool++)
	{
		pTool = m_pTools->GetTool(nTool);
		if (pTool->tpV1.m_nToolId == nToolId)
			return;
	}
	pTool = NULL;
}

void CBar::ToolNotification(CTool* pTool, TOOLNF tnf)
{
	CBand* pBand;
	int nBandCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nBandCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (pBand)
			pBand->ToolNotification(pTool, tnf);
	}
}

void CBar::OnTimer(UINT nId)
{
	try
	{
		switch (nId)
		{
		case eFlyByTimer:
			if (m_pActiveTool && !m_bMenuLoop)
				ToolNotification(m_pActiveTool, TNF_ACTIVETOOLCHECK);

			if (NULL == m_pActiveTool)
				KillTimer(m_hWnd, eFlyByTimer);
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

void DownloadCallback(void* pCookie, AsyncDownloadItem* pItem, DWORD dwReason)
{
	CBar* pControl = (CBar*)pCookie;
	if (NULL == pControl)
		return;

	switch (dwReason)
	{
	case ASYNC_MSG_DATAREADY:
		pControl->DataReady(pItem, TRUE);
		break;

	case ASYNC_MSG_ERROR:
		pControl->DataReady(pItem, FALSE);
		break;
	}
}

void CBar::ResetCycleMarks()
{
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
		m_pBands->GetBand(nBand)->m_bCycleMark = FALSE;
}

void CBar::DestroyDetachBandOf(CBand* pBand)
{
	if (NULL == pBand->m_bstrName)
		return;

	CBand* pCmpBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; ++nBand)
	{
		pCmpBand = m_pBands->GetBand(nBand);

		if (pCmpBand->m_bstrName && 
			0 == wcscmp(pCmpBand->m_bstrName, pBand->m_bstrName) &&
			VARIANT_TRUE == pCmpBand->bpV1.m_vbDetached)
		{
			if (pCmpBand->m_pPopupWin)
				DestroyWindow(pCmpBand->m_pPopupWin->hWnd());

			if (pCmpBand->m_pFloat)
				pCmpBand->DoFloat(FALSE);

			m_pBands->RemoveEx(pCmpBand);
			RecalcLayout();
			break;
		}
	}
}

//
// Customize
//

void CBar::DoCustomization(BOOL bFlag, BOOL bRecalcLayout)
{
	if (m_bCustomization == bFlag)
		return;
	
	m_bCustomization = bFlag;
	SetActiveTool(NULL);
	HideToolTips(TRUE);
	memset(&m_diCustSelection, 0, sizeof(DropInfo));

	if (bFlag)
	{
		m_vbCustomizeModified = VARIANT_FALSE;
		
		GetGlobals().m_pCustomizeActiveBar = this;			

		if (bRecalcLayout)
			RecalcLayout();
	}
	else
	{
		FireCustomizeEnd(m_vbCustomizeModified);
		
		GetGlobals().m_pCustomizeActiveBar = NULL;
		
		if (VARIANT_FALSE == bpV1.m_vbUserDefinedCustomization && m_pCustomize)
		{
			if (IsWindow(m_pCustomize->hWnd()))
				DestroyWindow(m_pCustomize->hWnd());
			delete m_pCustomize;
			m_pCustomize = NULL;
		}

		if (m_diCustSelection.pBand)
			m_diCustSelection.pBand->Refresh();

		if (m_pPopupRoot)
			m_pPopupRoot->SetPopupIndex(-1);

		if (bRecalcLayout)
		{
			//RecalcLayout();
			PostMessage(m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
		}
	}
}

//
// SetActiveTool
//

void CBar::SetActiveTool(CTool* pNewActiveTool)
{
	try
	{
		if (m_bToolModal)
			return;

		CTool* pPrevActiveTool = m_pActiveTool;
		if (pPrevActiveTool)
		{
			try
			{
				if (m_pActiveTool)
				{
					FireMouseExit(reinterpret_cast<Tool*>(m_pActiveTool));
					m_pActiveTool = NULL;
				}
				if (CTool::ddTTMDIButtons == pPrevActiveTool->tpV1.m_ttTools)
				{
					pPrevActiveTool->m_mdibActive = CTool::eNone;
					pPrevActiveTool->m_pBand->InvalidateToolRect(pPrevActiveTool->m_pBand->m_pTools->GetToolIndex(pPrevActiveTool));
				}
				else
					ToolNotification(pPrevActiveTool, TNF_VIEWCHANGED);	
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
		if (pNewActiveTool)
		{
			try
			{
				m_pActiveTool = pNewActiveTool;
				ToolNotification(m_pActiveTool, TNF_VIEWCHANGED);
				FireMouseEnter(reinterpret_cast<Tool*>(m_pActiveTool));
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
		else
			HideToolTips(FALSE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CalcFontHeight
//

static int CalcFontHeight(HDC hDC, CFontHolder* pFontHolder)
{
	if (NULL == pFontHolder)
		return 12;

	HFONT hFont = pFontHolder->GetFontHandle();
	if (NULL == hFont)
		return 12;

	HFONT hFontOld = SelectFont(hDC, hFont);
	TEXTMETRIC tm;
	if (!GetTextMetrics(hDC, &tm))
		tm.tmHeight = 12;
	SelectFont(hDC, hFontOld);
	return tm.tmHeight;
}

//
// OnFontHeightChanged
//

void CBar::OnFontHeightChanged()
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return;
	m_nFontHeight = CalcFontHeight(hDC, &m_fhFont);
	m_nControlFontHeight = CalcFontHeight(hDC, &m_fhControl);
	m_nChildBandFontHeight = CalcFontHeight(hDC, &m_fhChildBand);
	ReleaseDC(NULL, hDC);
}

//
// HideMiniWins
//

void CBar::HideMiniWins()
{
	CBand* pBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (pBand && pBand->m_pFloat)
			pBand->m_pFloat->ShowWindow(SW_HIDE);
	}
}

//
// DestroyMiniWins
//

void CBar::DestroyMiniWins()
{
	CBand* pBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		try
		{
			pBand = m_pBands->GetBand(nBand);
			if (pBand && pBand->m_pFloat)
			{
				pBand->m_pFloat->BandClosing(TRUE);
				if (pBand->m_pFloat->IsWindow())
					::SendMessage(m_hWnd, GetGlobals().WM_KILLWINDOW, (WPARAM)pBand->m_pFloat->hWnd(), 0);
//					pBand->m_pFloat->SendMessage(WM_CLOSE);
				else
					delete pBand->m_pFloat;
			}
		}
		catch (...)
		{
		}
	}
}

//
// DestroyPopups
//

void CBar::DestroyPopups()
{
	CBand* pBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (pBand && pBand->m_pPopupWin && pBand->m_pPopupWin->IsWindow())
			::PostMessage(m_hWnd, GetGlobals().WM_KILLWINDOW, (WPARAM)pBand->m_pPopupWin->hWnd(), 0);
	}
}

//
// ShowPopups
//

void CBar::ShowPopups()
{
	CBand* pBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (pBand && pBand->m_pPopupWin && pBand->m_pPopupWin->IsWindow())
			pBand->m_pPopupWin->ShowWindow(SW_SHOW);
	}
}

//
// HidePopups
//

void CBar::HidePopups()
{
	CBand* pBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (pBand && pBand->m_pPopupWin && pBand->m_pPopupWin->IsWindow())
			pBand->m_pPopupWin->ShowWindow(SW_HIDE);
	}
}

//
// OnBarContextMenu
//
//		This functions dynamically builds the customization menu 
// for customization mode
//
//

BOOL CBar::OnBarContextMenu()
{
	HRESULT hResult;
	
	CBand* pContextBand = CBand::CreateInstance(NULL);
	if (NULL == pContextBand)
		return FALSE;

	pContextBand->SetOwner(this, TRUE);
	
	hResult = pContextBand->put_Type(ddBTPopup);
	if (FAILED(hResult))
		return FALSE;
	
	hResult = pContextBand->put_CreatedBy(ddCBSystem);
	if (FAILED(hResult))
		return FALSE;
	
	hResult = pContextBand->put_Caption(L"SysCustomize");
	if (FAILED(hResult))
		return FALSE;
	
	hResult = pContextBand->put_Name(L"SysCustomize");
	if (FAILED(hResult))
		return FALSE;

	hResult = pContextBand->put_DockingArea(ddDAPopup);
	if (FAILED(hResult))
		return FALSE;

	CBand* pBand;
	CTool* pTool;
	BOOL bCustomize = FALSE;
	int nBandCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nBandCount; ++nBand)
	{
		pBand = m_pBands->GetBand(nBand);
		assert(pBand);
		if (NULL == pBand)
			continue;

		if (ddBTNormal != pBand->bpV1.m_btBands)
			continue;
		
		if (ddDAPopup == pBand->bpV1.m_daDockingArea)
			continue;

		if (VARIANT_TRUE == pBand->bpV1.m_vbDetached)
			continue;
		
		if (!(pBand->bpV1.m_dwFlags & ddBFHide))
			continue;

		if (!bCustomize && (pBand->bpV1.m_dwFlags & ddBFCustomize))
			bCustomize = TRUE;

		hResult = pContextBand->m_pTools->CreateTool((ITool**)&pTool);
		if (FAILED(hResult))
			continue;

		pTool->tpV1.m_nToolId = eToolIdBandToggle;

		hResult = pTool->put_Caption(pBand->m_bstrCaption);
		if (FAILED(hResult))
			return FALSE;

		VARIANT vName;
		vName.vt = VT_BSTR;
		vName.bstrVal = SysAllocString(pBand->m_bstrName);
		hResult = pTool->put_TagVariant(vName);
		VariantClear(&vName);
		if (FAILED(hResult))
			return FALSE;
		
		if (pBand->bpV1.m_vbVisible)
			pTool->tpV1.m_vbChecked = VARIANT_TRUE;
		
		pTool->tpV1.m_mvMenuVisibility = ddMVAlwaysVisible;

		hResult = pContextBand->m_pTools->InsertTool(-1, pTool, VARIANT_FALSE);
		if (FAILED(hResult))
			return FALSE;
		
		pTool->Release();
	}
	
	if (pContextBand->m_pTools->GetToolCount() > 0 && bCustomize)
	{
		//
		// Add separator
		//

		hResult = pContextBand->m_pTools->CreateTool((ITool**)&pTool);
		if (SUCCEEDED(hResult))
		{
			pTool->tpV1.m_nToolId = eToolIdSeparator;
			pTool->tpV1.m_ttTools = ddTTSeparator;
			pTool->tpV1.m_mvMenuVisibility = ddMVAlwaysVisible;
			hResult = pContextBand->m_pTools->InsertTool(-1, pTool, VARIANT_FALSE);
			if (FAILED(hResult))
				return FALSE;
			pTool->Release();
		}
	}

	if (bCustomize)
	{
		//
		// Add customize item
		//
		
		hResult = pContextBand->m_pTools->CreateTool((ITool**)&pTool);
		if (SUCCEEDED(hResult))
		{
			pTool->tpV1.m_nToolId = eToolIdCustomize;
			LPCTSTR szLocaleString = Localizer()->GetString(ddLTMenuCustomize);
			if (szLocaleString)
			{
				MAKE_WIDEPTR_FROMTCHAR(wCust, szLocaleString);
				hResult = pTool->put_Caption(wCust);
			}
			
			pTool->tpV1.m_mvMenuVisibility = ddMVAlwaysVisible;

			hResult = pContextBand->m_pTools->InsertTool(-1, pTool, VARIANT_FALSE);
			if (FAILED(hResult))
				return FALSE;
			pTool->Release();
		}
	}

	//
	// Display the popup
	//

	if (pContextBand->m_pTools->GetToolCount() > 0 && bCustomize)
	{
		hResult = pContextBand->TrackPopupEx(ddPopupMenuLeftAlign, -1, -1, NULL);
		if (FAILED(hResult))
			return FALSE;
	}

	pContextBand->Release();
	return TRUE;
}

//
// CreateDockWindows
//

void CBar::CreateDockWindows()
{
	HWND hWndParent = GetParentWindow();
	m_pDockMgr->CreateDockAreasWnds(hWndParent);
	if (eMDIForm == m_eAppType)
		InternalRecalcLayout();
}

//
// LoadData
//

HRESULT CBar::LoadData(IStream* pStream, BOOL bConfiguration, BSTR bstrBandName)
{
	HRESULT hResult;
	if (NULL == bstrBandName || NULL == *bstrBandName)
	{
		m_pPopupRoot = NULL;
		DestroyPopups();
	}

	if (m_pStatusBand)
	{
		StatusBandExit();
		m_pStatusBand = NULL;
	}
	if (bConfiguration)
		hResult = PersistConfig(pStream, VARIANT_FALSE);
	else
	{
		if (NULL == bstrBandName || NULL == *bstrBandName)
		{
			m_pDockMgr->CleanupLines();
			m_pBands->RemoveAll();
			m_pTools->RemoveAll();
		}
		hResult = PersistState(pStream, VARIANT_FALSE, bstrBandName);
	}

	CBand* pBand;
	int nCount = m_pBands->GetBandCount();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_pBands->GetBand(nBand);
		if (ddBTStatusBar == pBand->bpV1.m_btBands)
		{
			StatusBand(pBand);
			break;
		}
	}
	return hResult;
}

//
// SaveData
//

HRESULT CBar::SaveData(IStream* pStream, BOOL bConfiguration, BSTR bstrBandName)
{
	if (bstrBandName && *bstrBandName)
	{
		CBand* pBand = FindBand(bstrBandName);
		if (NULL == pBand)
		{
			m_theErrorObject.SendError(IDERR_BANDNOTFOUND, bstrBandName);
			return E_FAIL;
		}
	}

	HRESULT hResult;
	if (bConfiguration)
		hResult = PersistConfig(pStream, VARIANT_TRUE);
	else
		hResult = PersistState(pStream, VARIANT_TRUE, bstrBandName);

	return hResult;
}

//
// PersistState
//

HRESULT CBar::PersistState(IStream* pStream, VARIANT_BOOL vbSave, BSTR bstrBandName)
{
	try
	{
		long nStreamSize;
		short nSize;
		short nSize2;
		HRESULT hResult;
		short nBandCount = 1;
		BOOL bAll = NULL == bstrBandName || NULL == *bstrBandName;

		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving 
			//

			int nLayoutVersion = eLayoutVersion;
			hResult = pStream->Write(&nLayoutVersion, sizeof(nLayoutVersion), NULL);
			if (FAILED(hResult))
				goto Cleanup;

			if (bAll)
				nBandCount = m_pBands->GetBandCount();

			hResult = pStream->Write(&nBandCount, sizeof(nBandCount), NULL);
			if (FAILED(hResult))
				goto Cleanup;

			if (bAll)
			{
				hResult = m_pImageMgr->Exchange(this, pStream, vbSave);
				if (FAILED(hResult))
					goto Cleanup;
			}
			else
			{
				//
				// Create a Band Specific ImageMgr so the Band's Images can be copied to it.
				//

				m_pImageMgrBandSpecific = CImageMgr::CreateInstance(NULL);
				if (NULL == m_pImageMgrBandSpecific)
				{
					hResult = E_OUTOFMEMORY;
					goto Cleanup;
				}

				//
				// Find the band
				//

				CBand* pBand = FindBand(bstrBandName);
				if (NULL == pBand)
				{
					hResult = E_FAIL;
					goto Cleanup;
				}

				//
				// Copy the images to the Band Specific Image Manager
				//

				m_bSaveImages = TRUE;
				
				hResult = pBand->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				m_bSaveImages = FALSE;

				//
				// Save the Band's images
				//

				hResult = m_pImageMgrBandSpecific->Exchange(this, pStream, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				m_bBandSpecificConvertIds = TRUE;
			}
				
			hResult = m_pBands->Exchange(pStream, vbSave, bstrBandName, nBandCount);
			if (FAILED(hResult))
				goto Cleanup;

			if (bAll)
			{
				//
				// Bar Properties
				//

				hResult = m_pTools->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				nStreamSize = GetStringSize(m_bstrHelpFile) + m_phPict.GetSize() + m_fhFont.GetSize() + m_fhControl.GetSize() + m_fhChildBand.GetSize() + sizeof(nSize) + sizeof(bpV1);
				hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
				if (FAILED(hResult))
					goto Cleanup;
				
				hResult = StWriteBSTR(pStream, m_bstrHelpFile);
				if (FAILED(hResult))
					goto Cleanup;

				hResult = PersistPicture(pStream, &m_phPict, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				hResult = PersistFont(pStream, &m_fhFont, &GetGlobals()._fdDefault, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				hResult = PersistFont(pStream, &m_fhControl, &GetGlobals()._fdDefault, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				hResult = PersistFont(pStream, &m_fhChildBand, &GetGlobals()._fdDefault, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				nSize = sizeof(bpV1);
				hResult = pStream->Write(&nSize, sizeof(nSize), NULL);
				if (FAILED(hResult))
					goto Cleanup;

				hResult = pStream->Write(&bpV1, nSize, NULL);
				if (FAILED(hResult))
					goto Cleanup;
			}
		}
		else
		{
			//
			// Loading
			//

			int nCheckVersion;
			hResult = pStream->Read(&nCheckVersion, sizeof(nCheckVersion), NULL);
			if (FAILED(hResult))
				goto Cleanup;

			hResult = pStream->Read(&nBandCount, sizeof(nBandCount), NULL);
			if (FAILED(hResult))
				goto Cleanup;

			if (bAll)
			{
				hResult = m_pImageMgr->Exchange(this, pStream, vbSave);
				if (FAILED(hResult))
					goto Cleanup;
			}
			else
			{
				m_pImageMgrBandSpecific = CImageMgr::CreateInstance(NULL);
				if (NULL == m_pImageMgrBandSpecific)
				{
					hResult = E_OUTOFMEMORY;
					goto Cleanup;
				}

				hResult = m_pImageMgrBandSpecific->Exchange(this, pStream, vbSave);
				if (FAILED(hResult))
					goto Cleanup;
			}

 			hResult = m_pBands->Exchange(pStream, vbSave, bstrBandName, nBandCount);
			if (FAILED(hResult))
				goto Cleanup;

			if (bAll)
			{
				ULARGE_INTEGER nBeginPosition;
				LARGE_INTEGER  nOffset;

				//
				// Bar Properties
				//

				hResult = m_pTools->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
				if (FAILED(hResult))
					goto Cleanup;

				hResult = StReadBSTR(pStream, m_bstrHelpFile);
				if (FAILED(hResult))
					goto Cleanup;

				nStreamSize -= GetStringSize(m_bstrHelpFile);
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = PersistPicture(pStream, &m_phPict, vbSave);
				if (FAILED(hResult))
					goto Cleanup;
				
				nStreamSize -= m_phPict.GetSize();
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = PersistFont(pStream, &m_fhFont, &GetGlobals()._fdDefault, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				nStreamSize -= m_fhFont.GetSize();
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = PersistFont(pStream, &m_fhControl, &GetGlobals()._fdDefault, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				nStreamSize -= m_fhControl.GetSize();
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = PersistFont(pStream, &m_fhChildBand, &GetGlobals()._fdDefault, vbSave);
				if (FAILED(hResult))
					goto Cleanup;

				nStreamSize -= m_fhChildBand.GetSize();
				if (nStreamSize <= 0)
					goto FinishedReading;

				hResult = pStream->Read(&nSize, sizeof(nSize), NULL);
				if (FAILED(hResult))
					goto Cleanup;

				nStreamSize -= sizeof(nSize);
				if (nStreamSize <= 0)
					goto FinishedReading;

				nSize2 = sizeof(bpV1);
				hResult = pStream->Read(&bpV1, nSize < nSize2 ? nSize : nSize2, NULL);
				if (FAILED(hResult))
					goto Cleanup;

				if (nSize2 < nSize)
				{
					nOffset.HighPart = 0;
					nOffset.LowPart = nSize - nSize2;
					hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
					if (FAILED(hResult))
						goto Cleanup;
				}

				nStreamSize -= nSize;
				if (nStreamSize <= 0)
					goto FinishedReading;

				//
				// This will read the slack
				//

				nOffset.HighPart = 0;
				nOffset.LowPart = nStreamSize;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					goto Cleanup;

FinishedReading:
				if (ddASProportional == bpV1.m_asChildren && NULL == m_pRelocate)
				{
					m_pRelocate = new CRelocate(this);
					if (NULL == m_pRelocate)
						return E_OUTOFMEMORY;
				}

				if (8 == GetGlobals().m_nBitDepth)
				{
					m_pColorQuantizer->AddColor(GetRValue(m_crBackground), 
												GetGValue(m_crBackground), 
												GetBValue(m_crBackground));
					m_pColorQuantizer->AddColor(GetRValue(m_crForeground), 
												 GetGValue(m_crForeground), 
												 GetBValue(m_crForeground));
					m_pColorQuantizer->AddColor(GetRValue(m_crHighLight), 
												GetGValue(m_crHighLight), 
												GetBValue(m_crHighLight));
					m_pColorQuantizer->AddColor(GetRValue(m_crShadow), 
												GetGValue(m_crShadow), 
												GetBValue(m_crShadow));
					m_pColorQuantizer->AddColor(GetRValue(m_cr3DLight), 
												GetGValue(m_cr3DLight), 
												GetBValue(m_cr3DLight));
					m_pColorQuantizer->AddColor(GetRValue(m_cr3DDarkShadow), 
												GetGValue(m_cr3DDarkShadow), 
												GetBValue(m_cr3DDarkShadow));
					m_pColorQuantizer->AddColor(GetRValue(m_crMDIMenuBackground), 
												GetGValue(m_crMDIMenuBackground), 
												GetBValue(m_crMDIMenuBackground));
				}

				if (VARIANT_TRUE == bpV1.m_vbAlignToForm)
				{
					Detach();
					Attach(NULL);
				}
			}
			if (8 == GetGlobals().m_nBitDepth)
				m_pImageMgr->put_Palette((OLE_HANDLE)m_pColorQuantizer->CreatePalette());
		}
Cleanup:
		if (m_pImageMgrBandSpecific)
		{
			m_pImageMgrBandSpecific->Release();
			m_pImageMgrBandSpecific = NULL;
		}
		if (m_bBandSpecificConvertIds)
			m_bBandSpecificConvertIds = FALSE;
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
// PersistConfig
//

HRESULT CBar::PersistConfig(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		long nStreamSize;
		short nSize;
		short nSize2;
		HRESULT hResult;
		short nBandCount;

		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving 
			//

			nStreamSize = GetStringSize(m_bstrHelpFile) + m_phPict.GetSize() + m_fhFont.GetSize() + m_fhControl.GetSize() + m_fhChildBand.GetSize() + sizeof(nSize) + sizeof(bpV1);
			hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			int nLayoutVersion = eConfigLayoutVersion;
			hResult = pStream->Write(&nLayoutVersion, sizeof(nLayoutVersion), NULL);
			if (FAILED(hResult))
				return hResult;

			nBandCount = m_pBands->GetBandCount();
			hResult = pStream->Write(&nBandCount, sizeof(nBandCount), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = m_pImageMgr->ExchangeConfig(this, pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			hResult = m_pBands->ExchangeConfig(pStream, vbSave, nBandCount);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&m_pRelocate, sizeof(m_pRelocate), NULL);
			if (FAILED(hResult))
				return hResult;

			if (m_pRelocate)
				m_pRelocate->ExchangeConfig(pStream, vbSave);

			nSize = sizeof(bpV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&bpV1, nSize, NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = m_pTools->ExchangeConfig(pStream, vbSave);
			if (FAILED(hResult))
				return hResult;
		}
		else
		{
			//
			// Loading
			//

			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			int nCheckVersion;
			hResult = pStream->Read(&nCheckVersion, sizeof(nCheckVersion), NULL);
			if (FAILED(hResult))
				return hResult;

/*			if (eConfigLayoutVersion < nCheckVersion)
			{
				m_theErrorObject.m_nAsyncError++;
				m_theErrorObject.SendError(IDERR_INVALIDFILEFORMAT, NULL);
				m_theErrorObject.m_nAsyncError--;
				return CUSTOM_CTL_SCODE(IDERR_INVALIDFILEFORMAT);
			}
*/
			hResult = pStream->Read(&nBandCount, sizeof(nBandCount), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = m_pImageMgr->ExchangeConfig(this, pStream, vbSave);
			if (FAILED(hResult))
				return hResult;

			hResult = m_pBands->ExchangeConfig(pStream, vbSave, nBandCount);
			if (FAILED(hResult))
				return hResult;

			CRelocate* pRelocate;
			hResult = pStream->Read(&pRelocate, sizeof(pRelocate), NULL);
			if (FAILED(hResult))
				return hResult;

			if (pRelocate)
			{
				if (NULL == m_pRelocate)
				{
					m_pRelocate = new CRelocate(this);
					if (NULL == m_pRelocate)
						return E_OUTOFMEMORY;
					m_pRelocate->GetControlInfo();
				}
				else if (m_pRelocate->Count() < 1 && ddASProportional == bpV1.m_asChildren && (eSDIForm == m_eAppType || eClientArea == m_eAppType))
					m_pRelocate->GetControlInfo();
				m_pRelocate->ExchangeConfig(pStream, vbSave);
			}

			hResult = pStream->Read(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nSize2 = sizeof(bpV1);
			hResult = pStream->Read(&bpV1, nSize < nSize2 ? nSize : nSize2, NULL);
			if (FAILED(hResult))
				return hResult;


			if (nSize2 < nSize)
			{
				ULARGE_INTEGER nBeginPosition;
				LARGE_INTEGER  nOffset;
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}

			if (nCheckVersion > eConfigLayoutVersionOld)
			{
				hResult = m_pTools->ExchangeConfig(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}

		}
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
// DrawDropSign
//

void CBar::DrawDropSign(DropInfo& di)
{
	try
	{
		assert(di.pBand);
		if (NULL == di.pBand)
			return;

		HWND hWndBand;
		if (di.bFloating)
		{
			if (NULL == di.pBand->m_pRootBand->GetFloat())
				return;
			hWndBand = di.pBand->m_pRootBand->GetFloat()->hWnd();
		}
		else if (di.bPopup)
		{
			if (NULL == di.pBand->m_pRootBand->GetPopup())
				return;
			hWndBand = di.pBand->m_pRootBand->GetPopup()->hWnd();
		}
		else
			hWndBand = di.pBand->m_pRootBand->m_pDock->hWnd();

		// Span complete width if Popup
		int nWidth = -1;
		if (ddBTPopup == di.pBand->m_pRootBand->bpV1.m_btBands)
			nWidth = di.rcBand.Width();
			
		HDC hDC = GetDC(hWndBand);
		if (hDC)
		{
			di.pBand->DSGSetDropLoc((OLE_HANDLE)hDC,
									di.rcBand.left,
									di.rcBand.top,
									nWidth,
									di.rcBand.Height(),
									di.nDropIndex,
									di.nDropDir);
			ReleaseDC(hWndBand, hDC);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}			
}

//
// ClearCustomization
//

void CBar::ClearCustomization()
{
	if (NULL != m_hWndModifySelection && IsWindow(m_hWndModifySelection))
		EnableWindow(m_hWndModifySelection, FALSE);
	DropInfo diOld;
	memcpy(&diOld, &m_diCustSelection, sizeof(DropInfo));
	memset(&m_diCustSelection, NULL, sizeof(m_diCustSelection));
	if (diOld.pTool)
	{
		if (diOld.bPopup)
		{
			if (diOld.pBand->m_pPopupWin)
				diOld.pBand->m_pPopupWin->RefreshTool(diOld.pTool);
		}
		else
			diOld.pBand->ToolNotification(diOld.pTool, TNF_VIEWCHANGED);
		if (m_pPopupRoot)
			m_pPopupRoot->SetPopupIndex(-1);
	}
}

//
// PaintVBBackground
//

void CBar::PaintVBBackground(HWND m_hWnd, HDC hDC, BOOL bIsMDI)
{
	try
	{
		CRect rcClient;
		GetClientRect(m_hWnd, &rcClient);

		if (bIsMDI && AmbientUserMode()) 
		{
			// MDI
			return;
		}
		else if (!bIsMDI)
		{
			// SDI
			if (AmbientUserMode())
			{
				if (HasTexture() && ddBOClientArea & bpV1.m_dwBackgroundOptions)
					FillTexture(hDC, rcClient);
				else
					FillSolidRect(hDC, rcClient, m_crBackground);
				return;
			}
			else
			{
				if (HasTexture() && ddBOClientArea & bpV1.m_dwBackgroundOptions)
				{
					FillTexture(hDC, rcClient);
					return;
				}
			}
		}

		// no texture, no runtime+mdi so we have to paint the grid

		BOOL bDrawGrid = TRUE;
		HKEY  hKey = NULL;
		DWORD dwType = REG_SZ;
		DWORD dwSize = 120;
		TCHAR* szData[120];
		long lResult = RegOpenKeyEx(HKEY_CURRENT_USER, 
												_T("Software\\Microsoft\\Visual Basic\\6.0"),
												0,
												KEY_ALL_ACCESS, 
												&hKey);
		if (ERROR_SUCCESS != lResult) 
		{
			lResult = RegOpenKeyEx(HKEY_CURRENT_USER, 
									   _T("Software\\Microsoft\\Visual Basic\\5.0"),
									   0,
									   KEY_ALL_ACCESS, 
									   &hKey);
		}
		if (ERROR_SUCCESS == lResult) 
		{
			lResult = RegQueryValueEx(hKey, 
									_T("ShowGrid"), 
									0,
									&dwType,
									(LPBYTE)&szData, 
									&dwSize);
			if (ERROR_SUCCESS == lResult && 0 != lstrcmp((const TCHAR*)szData, _T("1"))) 
				bDrawGrid = FALSE;
		}
		
		RegCloseKey(hKey);

		if (bDrawGrid) 
		{
			RECT rcClipBox;
			if (GetClipBox(hDC, &rcClipBox) == ERROR)
				rcClipBox = rcClient;

			SetTextColor(hDC, bIsMDI ? m_crAmbientBackground : m_crBackground);
			SetBkColor(hDC, RGB(0, 0, 0));
			HBITMAP bmp = (HBITMAP)GetGlobals().GetBrushPattern();
			assert(bmp);
			if (bmp)
			{
				BITMAP bmInfo;
				RECT rcTemp;
				GetObject(bmp, sizeof(BITMAP), &bmInfo);
				
				if (bmInfo.bmHeight==0 || bmInfo.bmWidth==0)
					goto recoverPaint;
				
				int x,y;
				RECT r;
				
				for (y=rcClient.top;y<rcClient.bottom ;y+=bmInfo.bmHeight)
				{
					for (x=rcClient.left;x<rcClient.right ;x+=bmInfo.bmWidth)
					{
						r.top=y;
						r.bottom=r.top+1;
						r.left=x;
						r.right=x+1;
						
						if (IntersectRect(&rcTemp,&r,&rcClipBox))
							FillSolidRect(hDC, r, RGB(0,0,0));
					}
				}

				r = rcClient;
				for (x = rcClient.left; x < rcClient.right; x += bmInfo.bmWidth)
				{
					r.left = x + 1;
					r.right = r.left + bmInfo.bmWidth - 1;
					if (IntersectRect(&rcTemp, &r, &rcClipBox))
						FillSolidRect(hDC, r,  bIsMDI ? m_crAmbientBackground : m_crBackground);
				}

				r = rcClient;
				for (y = rcClient.top;y < rcClient.bottom; y += bmInfo.bmHeight)
				{
					r.top = y+1;
					r.bottom = r.top + bmInfo.bmHeight-1;
					if (IntersectRect(&rcTemp, &r, &rcClipBox))
						FillSolidRect(hDC, r,  bIsMDI ? m_crAmbientBackground : m_crBackground);
				}

				return;
			}
		}

recoverPaint:
		FillSolidRect(hDC, rcClient, m_crAmbientBackColor);
	}
	catch (...)
	{
	}
}

//
// StartStatusBandThread
//

BOOL CBar::StartStatusBandThread()
{
	m_hStatusbarThread = (HANDLE)_beginthreadex(NULL, 0, StatusBarThread, this, 0, &m_dwThreadId);
	return (NULL != m_hStatusbarThread);
}

//
// StopStatusBandThread
//

BOOL CBar::StopStatusBandThread()
{
	if (m_hStatusbarThread)
	{
		m_bStopStatusBarThread=TRUE;
		SetEvent(m_hStatusBarEvent);
		WaitForSingleObject(m_hStatusbarThread, 2000);
		CloseHandle(m_hStatusbarThread);
		m_hStatusbarThread =NULL;
	}
	return TRUE;
}

//
// CDesignTimeSystemSettingNotify
//
// This is so the designer can be notified that a system setting has changed
//

CDesignTimeSystemSettingNotify::CDesignTimeSystemSettingNotify(CBar* pBar)
	: m_pBar(pBar)
{
	BOOL bResult = RegisterWindow(DD_WNDCLASS("DesignTimeSystemSettingNotify"),
								  CS_SAVEBITS,
								  NULL,
								  ::LoadCursor(NULL, IDC_ARROW));
}

//
// ~CDesignTimeSystemSettingNotify()
//

CDesignTimeSystemSettingNotify::~CDesignTimeSystemSettingNotify()
{
	if (IsWindow())
		DestroyWindow();
}

//
// Create
//

BOOL CDesignTimeSystemSettingNotify::Create()
{
	FWnd::CreateEx(WS_EX_TRANSPARENT,
				   _T(""),
				   WS_OVERLAPPED,
				   0,
				   0,
				   0,
				   0,
				   (HWND)NULL);
	return (m_hWnd ? TRUE : FALSE);
}

//
// WindowProc 
//

LRESULT CDesignTimeSystemSettingNotify::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_SYSCOLORCHANGE:
	case WM_SETTINGCHANGE:
		m_pBar->OnSysColorChanged();
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// CChildRelocate
//

CChildRelocate::CChildRelocate()
	: m_pDispatch(NULL),
	  m_hWnd(NULL)
{
	crProperies.m_eControlType = eNormal;
	crProperies.m_nHeight = 0;
	crProperies.m_bHasFontProperty = FALSE;
}

CChildRelocate::~CChildRelocate()
{
}

//
// ExchangeConfig
//

HRESULT CChildRelocate::ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave)
{
	HRESULT hResult;

	if (VARIANT_TRUE == vbSave)
	{
		//
		// Saving
		//
		hResult = pStream->Write(&crProperies, sizeof(crProperies), NULL);
	}
	else
	{
		//
		// Loading
		//
		hResult = pStream->Read(&crProperies, sizeof(crProperies), NULL);
	}

	return hResult;
}

void CChildRelocate::CacheFontSize()
{
	LPDISPATCH pExtender = NULL;
	LPFONTDISP pFontDisp = NULL;
	LPFONT pFont = NULL;
	DISPID dispidFont;
	assert(!crProperies.m_bHasFontProperty);
	crProperies.m_bHasFontProperty = FALSE;
	HRESULT hResult;
	static WCHAR* szFont = L"Font";
	try
	{
		if (SUCCEEDED(m_pDispatch->GetIDsOfNames(IID_NULL,(LPWSTR *)&szFont, 1, LOCALE_SYSTEM_DEFAULT,&dispidFont)))
		{
			if (!GetAmbientProperty(m_pDispatch, dispidFont, VT_DISPATCH, &pFontDisp))
				goto Cleanup;
		}

		// We might have to use extender
		pExtender = GetExtenderDisp(m_pDispatch);
		if (pExtender)
		{
			if (SUCCEEDED(m_pDispatch->GetIDsOfNames(IID_NULL,(LPWSTR *)&szFont, 1, LOCALE_SYSTEM_DEFAULT,&dispidFont)))
			{
				// yes use extender
				m_pDispatch = pExtender;
			}
		}

		if (!GetAmbientProperty(m_pDispatch, dispidFont, VT_DISPATCH, &pFontDisp))
			goto Cleanup;

		hResult = pFontDisp->QueryInterface(IID_IFont,(LPVOID *)&pFont);
		if (FAILED(hResult))
			goto Cleanup;

		hResult=pFont->get_Size(&crProperies.m_sizeBaseFont);
		if (FAILED(hResult))
			goto Cleanup;

		// Success: record font size
		crProperies.m_bHasFontProperty=TRUE;
	}
	catch(...)
	{
		assert(FALSE);
	}

Cleanup:
	if (pFont)
		pFont->Release();
	if (pFontDisp)
		pFontDisp->Release();
	if (pExtender)
		pExtender->Release();
}

//
// CRelocate
//

CRelocate::CRelocate(CBar* pBar)
	: m_pBar(pBar)
{
}

CRelocate::~CRelocate()
{
	Cleanup();
}

void CRelocate::Cleanup()
{
	int nCount = m_aChildRelocate.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		delete m_aChildRelocate.GetAt(nIndex);
	m_aChildRelocate.RemoveAll();
}

//
// LayoutClient
//

BOOL CRelocate::LayoutClient(HWND hWnd, CRect rc)
{
	if (!m_pBar->AmbientUserMode() || m_aChildRelocate.GetSize() < 1)
		return TRUE;

	int nParentWidth = rc.Width();
	int nParentHeight = rc.Height();
//	if (nParentWidth < 1 || nParentHeight < 1)  Commented out for AR2 CR 3559
//		return TRUE;

	CRect rcChild = rc;

	SIZE size= {rcChild.left, rcChild.top};
	PixelToTwips(&size, &size);

	SIZE size2 = {rcChild.Width(), rcChild.Height()};
	PixelToTwips(&size2, &size2);
	
	CTool* pTool;
	VARIANT vProperty;
	vProperty.vt = VT_R4;
	CChildRelocate* pChild;
	int nCount = m_aChildRelocate.GetSize();
	for (int nControl = 0; (int)nControl < nCount; nControl++)
	{
		pChild = m_aChildRelocate.GetAt(nControl);

		if (IsWindow(pChild->m_hWnd) && m_pBar->m_theFormsAndControls.FindControl(pChild->m_hWnd, pTool))
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
			continue;
		}
		else if (m_pBar->m_theFormsAndControls.FindControl(pChild->m_pDispatch))
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
			continue;
		}

		if (CChildRelocate::eHWnd == pChild->crProperies.m_eControlType)
		{
			SetWindowPos(pChild->m_hWnd, NULL, rcChild.left, rcChild.top, rcChild.Width(), rcChild.Height(), SWP_SHOWWINDOW | SWP_NOACTIVATE);
		}
		else
		{
			vProperty.fltVal = (float)size.cx;
			HRESULT hResult = PropertyPut(pChild->m_pDispatch, L"Left", vProperty);
			if (FAILED(hResult))
				return FALSE;

			vProperty.fltVal = (float)size.cy;
			hResult = PropertyPut(pChild->m_pDispatch, L"Top", vProperty);
			if (FAILED(hResult))
				return FALSE;

			vProperty.fltVal = (float)size2.cx;
			hResult = PropertyPut(pChild->m_pDispatch, L"Width", vProperty);
			if (FAILED(hResult))
				return FALSE;

			vProperty.fltVal = (float)size2.cy;
			hResult = PropertyPut(pChild->m_pDispatch, L"Height", vProperty);
			if (FAILED(hResult))
				return FALSE;
		}
	}
	return TRUE;
}

HRESULT CRelocate::AddControl(HWND hWnd)
{
	if (!IsWindow(hWnd))
		return E_INVALIDARG;
	CChildRelocate* pChildRelocate = new CChildRelocate;
	if (NULL == pChildRelocate)
		return E_OUTOFMEMORY;
	Cleanup();
	pChildRelocate->m_hWnd = hWnd;
	pChildRelocate->crProperies.m_eControlType = CChildRelocate::eHWnd;
	return m_aChildRelocate.Add(pChildRelocate);
}

HRESULT CRelocate::AddControl(LPDISPATCH pDispatch)
{
	if (NULL == pDispatch)
		return E_INVALIDARG;
	CChildRelocate* pChildRelocate = new CChildRelocate;
	if (NULL == pChildRelocate)
		return E_OUTOFMEMORY;
	Cleanup();
	pChildRelocate->m_pDispatch = pDispatch;
	pChildRelocate->CacheFontSize();
	return m_aChildRelocate.Add(pChildRelocate);
}

//
// LayoutProportional
//

#define EPSILON 1e-6f

BOOL CRelocate::LayoutProportional(HWND hWnd, CRect rc)
{
	if (!m_pBar->AmbientUserMode())
		return TRUE;

	CRect rcChild;
	int nParentWidth = rc.Width();
	int nParentHeight = rc.Height();
	if (nParentWidth < 1 || nParentHeight < 1)
		return TRUE;

	TEXTMETRIC txMetrics;
	VARIANT vProperty;
	HRESULT hResult;
	HFONT hFontOld;
	HFONT hFont;
	HDC hDC;
	vProperty.vt = VT_R4;
	int nRows;
	SIZE size;
	SIZE size2;

	CTool* pTool;
	CChildRelocate* pChild;
	int nCount = m_aChildRelocate.GetSize();
	for (int nControl = 0; (int)nControl < nCount; nControl++)
	{
		pChild = m_aChildRelocate.GetAt(nControl);

		if (IsWindow(pChild->m_hWnd) && m_pBar->m_theFormsAndControls.FindControl(pChild->m_hWnd, pTool))
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
			continue;
		}
		else if (m_pBar->m_theFormsAndControls.FindControl(pChild->m_pDispatch))
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
			continue;
		}

		//
		// Placing the controls on a Non MDI Form during run time
		//

		rcChild.left = MulDiv(pChild->crProperies.m_rcPerOfParent.left, nParentWidth, 100);
		rcChild.top = MulDiv(pChild->crProperies.m_rcPerOfParent.top, nParentHeight, 100);
		rcChild.right = rcChild.left + MulDiv(pChild->crProperies.m_rcPerOfParent.right, nParentWidth, 100);

		switch (pChild->crProperies.m_eControlType)
		{
		case CChildRelocate::eToolBar:
			continue;

		case CChildRelocate::eListbox:
			{
				hFont = (HFONT)SendMessage(pChild->m_hWnd, WM_GETFONT, 0, 0);
				hDC = GetDC(NULL);
				if (hDC)
				{
					// This is not working
					hFontOld = SelectFont(hDC, hFont);
					GetTextMetrics(hDC, &txMetrics);
					pChild->crProperies.m_nHeight = txMetrics.tmHeight + txMetrics.tmExternalLeading + 2;
					nRows = (((pChild->crProperies.m_rcPerOfParent.bottom * nParentHeight)/100) - 6) / pChild->crProperies.m_nHeight;
					rcChild.bottom = rcChild.top + (nRows * pChild->crProperies.m_nHeight + 6);
					SelectFont(hDC, hFontOld);
					ReleaseDC(NULL, hDC);
				}
			}
			break;

		case CChildRelocate::eCombobox:
			rcChild.bottom = rcChild.top + pChild->crProperies.m_nHeight;
			break;

		default:
			rcChild.bottom = rcChild.top + MulDiv(pChild->crProperies.m_rcPerOfParent.bottom, nParentHeight, 100);
		}

		rcChild.Offset(rc.left, rc.top);

		size.cx = rcChild.left;
		size.cy = rcChild.top;
		PixelToTwips(&size, &size);

		size2.cx = rcChild.Width();
		size2.cy = rcChild.Height();
		PixelToTwips(&size2, &size2);

		if (pChild->crProperies.m_bHasFontProperty)
		{
			IFontDisp* pDispFont=NULL;
			IFont* pFont=NULL;
			float fBasePoint = ((float)pChild->crProperies.m_sizeBaseFont.Lo)/10000.0f;
			if (!(pChild->crProperies.m_fBaseHeight < EPSILON))
			{
				float newPoint = fBasePoint * (size2.cy / pChild->crProperies.m_fBaseHeight);
				if (fBasePoint != newPoint)
				{
					CY newSize;
					DISPID dispidFont;
					newSize.Hi=0;newSize.Lo = (int)(newPoint*10000.0f);
					WCHAR* szFont = L"Font";
					if (SUCCEEDED(pChild->m_pDispatch->GetIDsOfNames(IID_NULL,(LPWSTR *)&szFont, 1, LOCALE_SYSTEM_DEFAULT,&dispidFont)))
					{
						if (GetAmbientProperty(pChild->m_pDispatch, dispidFont,VT_DISPATCH, &pDispFont))
						{
							if (SUCCEEDED(pDispFont->QueryInterface(IID_IFont,(LPVOID *)&pFont)))
							{
								if (newSize.Lo>0 && newSize.Lo<4000000)
									pFont->put_Size(newSize);
								pFont->Release();
							}
							pDispFont->Release();
						}
					}
				}		
			}
		}

		try
		{
			vProperty.fltVal = (float)size.cx;
			hResult = PropertyPut(pChild->m_pDispatch, L"Left", vProperty);
			if (FAILED(hResult))
			{
				pChild->crProperies.m_eControlType = CChildRelocate::eToolBar;
				continue;
			}

			vProperty.fltVal = (float)size.cy;
			hResult = PropertyPut(pChild->m_pDispatch, L"Top", vProperty);
			if (FAILED(hResult))
			{
				pChild->crProperies.m_eControlType = CChildRelocate::eToolBar;
				continue;
			}

			vProperty.fltVal = (float)size2.cx;
			hResult = PropertyPut(pChild->m_pDispatch, L"Width", vProperty);
			if (FAILED(hResult))
			{
				pChild->crProperies.m_eControlType = CChildRelocate::eToolBar;
				continue;
			}

			vProperty.fltVal = (float)size2.cy;
			hResult = PropertyPut(pChild->m_pDispatch, L"Height", vProperty);
			if (FAILED(hResult))
			{
				pChild->crProperies.m_eControlType = CChildRelocate::eToolBar;
				continue;
			}
		}
		catch (...)
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
		}
	}
	return TRUE;
}

//
// GetControlInfo
//

BOOL CRelocate::GetControlInfo()
{
	if (!m_pBar->AmbientUserMode())
		return FALSE;

	int nCount = m_aChildRelocate.GetSize();
	if (NULL == m_pBar->m_pClientSite || nCount > 0)
		return FALSE;

	const ULONG nNumberOfUnknowns = 10;
	TypedArray<LPDISPATCH> aChildControls;
	LPOLECONTROLSITE pControlSite = NULL;
	LPOLECLIENTSITE pClientSite = NULL;
	IVBGetControl* pVBGetControl = NULL;
	LPENUMUNKNOWN pEnumObject = NULL;
	LPOLEOBJECT pObject = NULL;
	LPDISPATCH pDispatch;
	LPUNKNOWN pControlUnknown[nNumberOfUnknowns];
	VARIANT vProperty;
	ULONG nNumberReturned;
	ULONG nControl;
	DWORD dwStatus;
	BOOL  bResult = TRUE;

	HRESULT hResult = m_pBar->m_pClientSite->QueryInterface(IID_IVBGetControl, (void**)&pVBGetControl);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	hResult = pVBGetControl->EnumControls(OLECONTF_EMBEDDINGS | OLECONTF_OTHERS, GC_WCH_CONTAINED, &pEnumObject);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}
 
	hResult = pEnumObject->Next(nNumberOfUnknowns, (IUnknown**)&pControlUnknown, &nNumberReturned);
	if (FAILED(hResult))
	{
		bResult = FALSE;
		goto Cleanup;
	}

	if (nNumberReturned < 1)
	{
		bResult = FALSE;
		goto Cleanup;
	}

	while (nNumberReturned > 0)
	{
		for (nControl = 0; nControl < nNumberReturned; nControl++)
		{
			hResult = pControlUnknown[nControl]->QueryInterface(IID_IOleObject, (void**)&pObject);
			if (FAILED(hResult))
			{
				//
				// This is an internal control
				//

				hResult = pControlUnknown[nControl]->QueryInterface(IID_IDispatch, (void**)&pDispatch);
				pControlUnknown[nControl]->Release();
				if (FAILED(hResult))
					continue;
			}
			else
			{
				pControlUnknown[nControl]->Release();

				hResult = pObject->GetMiscStatus(DVASPECT_CONTENT, &dwStatus);
				if (FAILED(hResult))
				{
					pObject->Release();
					continue;
				}

				if (OLEMISC_INVISIBLEATRUNTIME & dwStatus)
				{
					pObject->Release();
					continue;
				}

				hResult = pObject->QueryInterface(IID_IDispatch, (void**)&pDispatch);
				if (FAILED(hResult))
				{
					pObject->Release();
					continue;
				}

				pDispatch->Release();

				hResult = pObject->GetClientSite(&pClientSite);
				pObject->Release();
				if (FAILED(hResult))
					continue;

				hResult = pClientSite->QueryInterface(IID_IOleControlSite, (void**)&pControlSite);
				pClientSite->Release();
				if (FAILED(hResult))
					continue;

				hResult = pControlSite->GetExtendedControl(&pDispatch);
				pControlSite->Release();
				if (FAILED(hResult))
					continue;
			}

			if (m_pBar->m_theFormsAndControls.FindControl(pDispatch))
			{
				pDispatch->Release();
				continue;
			}

			if (OLEMISC_ALIGNABLE & dwStatus)
			{
				hResult = PropertyGet(pDispatch, L"Align", vProperty);
				if (SUCCEEDED(hResult))
				{
					if (0 != vProperty.iVal)
					{
						//
						// Align property is something other than None skip this one and continue
						//

						pDispatch->Release();
						continue;
					}
				}
			}
			hResult = aChildControls.Add(pDispatch);
		}
		hResult = pEnumObject->Next(nNumberOfUnknowns, (IUnknown**)&pControlUnknown, &nNumberReturned);
		if (FAILED(hResult))
			nNumberReturned = 0;
	}
	Cleanup();
	nCount = aChildControls.GetSize();
	if (0 == nCount)
		goto Cleanup;
	
	CChildRelocate* pChild;
	for (nControl = 0; (int)nControl < nCount; nControl++)
	{
		pChild = new CChildRelocate;
		assert(pChild);
		if (NULL == pChild)
			continue;

		pChild->m_pDispatch = aChildControls.GetAt(nControl); 

		hResult = PropertyGet(pChild->m_pDispatch, L"hWnd", vProperty);
		if (SUCCEEDED(hResult) && IsWindow((HWND)vProperty.lVal))
			pChild->m_hWnd = (HWND)vProperty.lVal;

		pChild->m_pDispatch->Release();
		hResult = m_aChildRelocate.Add(pChild);
		pChild->CacheFontSize();
	}
Cleanup:
	if (pVBGetControl)
		pVBGetControl->Release();
	if (pEnumObject)
		pEnumObject->Release();

	return bResult;
}

//
// GetLayoutInfo
//

int CRelocate::GetLayoutInfo(HWND hWnd, const CRect& rc)
{
	BOOL bResult = GetControlInfo();
	if (!bResult)
		return bResult;

	HRESULT hResult;
	const int nLen = 40;
	TCHAR szClassName[nLen];
	CRect rcChild;
	int nParentWidth = rc.Width();
	int nParentHeight = rc.Height();
	if (nParentWidth < 1 || nParentHeight < 1)
		return TRUE;

	VARIANT vProperty;
	ULONG nControl;

	DWORD dwStyle;
	SIZE size;
	CChildRelocate* pChild;
	int nCount = m_aChildRelocate.GetSize();
	for (nControl = 0; (int)nControl < nCount; nControl++)
	{
		vProperty.vt = VT_EMPTY;

		pChild = m_aChildRelocate.GetAt(nControl);

		hResult = PropertyGet(pChild->m_pDispatch, L"Left", vProperty);
		if (FAILED(hResult))
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
			continue;
		}

		hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
		if (FAILED(hResult))
			continue;

		size.cx = vProperty.lVal;

		hResult = PropertyGet(pChild->m_pDispatch, L"Top", vProperty);
		if (FAILED(hResult))
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
			continue;
		}

		hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
		if (FAILED(hResult))
			continue;

		size.cy = vProperty.lVal;

		TwipsToPixel(&size, &size);
		
		rcChild.left = size.cx;
		rcChild.top = size.cy;

		hResult = PropertyGet(pChild->m_pDispatch, L"Width", vProperty);
		if (FAILED(hResult))
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
			continue;
		}

		pChild->crProperies.m_fBaseWidth = vProperty.fltVal;

		hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
		if (FAILED(hResult))
		{
			assert(FALSE);
			continue;
		}

		size.cx = vProperty.lVal;
		
		hResult = PropertyGet(pChild->m_pDispatch, L"Height", vProperty);
		if (FAILED(hResult))
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
			continue;
		}

		pChild->crProperies.m_fBaseHeight = vProperty.fltVal;

		hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
		if (FAILED(hResult))
		{
			assert(FALSE);
			continue;
		}

		size.cy = vProperty.lVal;
		
		TwipsToPixel(&size, &size);
		rcChild.right = rcChild.left + size.cx;
		rcChild.bottom = rcChild.top + size.cy;
		
		rcChild.Offset(-rc.left, -rc.top);
		pChild->crProperies.m_rcPerOfParent.left = MulDiv(rcChild.left, 100, nParentWidth);
		pChild->crProperies.m_rcPerOfParent.top = MulDiv(rcChild.top, 100, nParentHeight);
		pChild->crProperies.m_rcPerOfParent.right = MulDiv(rcChild.Width(), 100, nParentWidth);
		pChild->crProperies.m_rcPerOfParent.bottom = MulDiv(rcChild.Height(), 100, nParentHeight);

		if (IsWindow(pChild->m_hWnd))
		{
			//
			// Windowed Control
			//
			// Special checking for Comboboxes and Listboxes
			//

			if (CChildRelocate::eToolBar != pChild->crProperies.m_eControlType)
			{
				GetClassName(pChild->m_hWnd, szClassName, nLen);
				if (0 == lstrcmpi(szClassName, _T("ThunderComboBox")))
				{
					pChild->crProperies.m_eControlType = CChildRelocate::eCombobox;
					pChild->crProperies.m_nHeight = rcChild.Height();
				}
				else if (0 == lstrcmpi(szClassName, _T("ThunderListBox")) || 
						 0 == lstrcmpi(szClassName, _T("ThunderDirListBox")) ||
						 0 == lstrcmpi(szClassName, _T("ThunderDriveListBox")))
				{
					dwStyle = GetWindowLong(pChild->m_hWnd, GWL_STYLE);
					if (!(dwStyle & LBS_NOINTEGRALHEIGHT))
						pChild->crProperies.m_eControlType = CChildRelocate::eListbox;
				}
			}
		}
	}
	return bResult;
}

//
// FindControls
//

BOOL CRelocate::FindControl(LPDISPATCH pDispatch)
{
	CChildRelocate* pChild;
	int nCount = m_aChildRelocate.GetSize();
	for (int nControl = 0; (int)nControl < nCount; nControl++)
	{
		pChild = m_aChildRelocate.GetAt(nControl);
		if (pChild->m_pDispatch == pDispatch)
		{
			m_aChildRelocate.RemoveAt(nControl);
			delete pChild;
			nControl--;
			nCount--;
		}
	}
	return FALSE;
}

//
// ExchangeConfig
//

HRESULT CRelocate::ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave)
{
	HRESULT hResult;
	CChildRelocate* pChild;
	int nCount = m_aChildRelocate.GetSize();

	if (VARIANT_TRUE == vbSave)
	{
		//
		// Saving
		//

		for (int nControl = 0; (int)nControl < nCount; nControl++)
		{
			pChild = m_aChildRelocate.GetAt(nControl);
			hResult = pChild->ExchangeConfig(pStream, vbSave);
		}
	}
	else
	{
		//
		// Loading
		//

		for (int nControl = 0; (int)nControl < nCount; nControl++)
		{
			pChild = m_aChildRelocate.GetAt(nControl);
			hResult = pChild->ExchangeConfig(pStream, vbSave);
		}
	}

	return hResult;
}

//
// FormsAndControls
//
//
// Remove
//

BOOL FormsAndControls::Remove(CTool* pTool)
{
	try
	{
		int nSize = m_aFormsAndControls.GetSize();
		for (int nTool = 0; nTool < nSize; nTool++)
		{
			if (pTool == m_aFormsAndControls.GetAt(nTool))
			{
				if (IsWindow(pTool->m_hWndActive))
					ShowWindow(pTool->m_hWndActive, SW_HIDE);
				m_aFormsAndControls.RemoveAt(nTool);
				return TRUE;
			}
		}
	}
	catch (...)
	{
		assert(FALSE);
	}
	return FALSE;
}

//
// FindChildFromPoint
//

HWND FormsAndControls::FindChildFromPoint(POINT ptScreen, CTool** ppTool)
{
	CTool* pTool;
	CRect rcControl;
	int nSize = m_aFormsAndControls.GetSize();
	for (int nTool = 0; nTool < nSize; nTool++)
	{
		pTool = m_aFormsAndControls.GetAt(nTool);
		if (!IsWindow(pTool->m_hWndActive))
			continue;

		if (!GetWindowRect(pTool->m_hWndActive, &rcControl))
			continue;

		if (PtInRect(&rcControl, ptScreen))
		{
			if (ppTool)
				*ppTool = pTool;
			return pTool->m_hWndActive;
		}
	}
	if (ppTool)
		*ppTool = NULL;
	return NULL;
}

//
// IsChild
//

BOOL FormsAndControls::IsChild(HWND hWnd)
{
	int nSize = m_aFormsAndControls.GetSize();
	for (int nTool = 0; nTool < nSize; nTool++)
	{
		if (::IsChild(m_aFormsAndControls.GetAt(nTool)->m_hWndActive, hWnd))
			return TRUE;
	}
	return FALSE;
}

//
// FindControl
//

BOOL FormsAndControls::FindControl(LPDISPATCH pDispatch)
{
	int nSize = m_aFormsAndControls.GetSize();
	for (int nTool = 0; nTool < nSize; nTool++)
	{
		if (m_aFormsAndControls.GetAt(nTool)->m_pDispCustom == pDispatch && NULL != pDispatch)
			return TRUE;
	}
	return FALSE;
}

//
// FindControl
//

BOOL FormsAndControls::FindControl(HWND hWnd, CTool*& pTool)
{
	int nSize = m_aFormsAndControls.GetSize();
	for (int nTool = 0; nTool < nSize; nTool++)
	{
		if (m_aFormsAndControls.GetAt(nTool)->m_hWndActive == hWnd)
		{
			pTool = m_aFormsAndControls.GetAt(nTool);
			return TRUE;
		}
	}
	return FALSE;
}

//
// EnableForms
//

void FormsAndControls::EnableForms(BOOL bEnable)
{
	CTool* pTool;
	int nSize = m_aFormsAndControls.GetSize();
	for (int nTool = 0; nTool < nSize; nTool++)
	{
		pTool = m_aFormsAndControls.GetAt(nTool);
		if (ddTTForm && pTool->tpV1.m_ttTools)
			EnableWindow(pTool->m_hWndActive, bEnable);
	}
}

//
// ShutDown
//

void FormsAndControls::ShutDown()
{
	try
	{
		CTool* pTool;
		int nSize = m_aFormsAndControls.GetSize();
		while (nSize > 0)
		{
			pTool = m_aFormsAndControls.GetAt(0);
			//
			// put_Custom removes this tool from the m_aFormsAndControls array
			//
			if (IsWindow(pTool->m_hWndActive))
				PostMessage(pTool->m_hWndActive, WM_CLOSE, 0, 0);
			if (pTool->m_pDispCustom)
				pTool->put_Custom(NULL);
			else
				pTool->put_hWnd(NULL);
			--nSize;
		}
	}
	catch (...)
	{
		assert(FALSE);
	}
}

//
// CChildMenu
//

CChildMenu::CChildMenu()
	: m_pBar(NULL),
	  m_pMainBand(NULL)
{
}

CChildMenu::~CChildMenu()
{
	HWND hWnd;
	BSTR bstrBandName;
	FPOSITION posMap = m_theChildMenus.GetStartPosition();
	while (posMap)
	{
		bstrBandName = NULL;
		m_theChildMenus.GetNextAssoc(posMap, hWnd, bstrBandName);
		SysFreeString(bstrBandName);
	}
	m_theChildMenus.RemoveAll();
}

//
// SetBar
//

void CChildMenu::SetBar(CBar* pBar)
{
	m_pBar = pBar;
}

//
// SwitchMenus
//

BOOL CChildMenu::SwitchMenus(HWND hWndChild)
{
	try
	{
		CBand* pBand = m_pBar->GetMenuBand();
		if (NULL == hWndChild)
		{
			if (pBand)
			{
				if (ddBTMenuBar == pBand->bpV1.m_btBands && NULL == m_pMainBand)
					m_pMainBand = pBand;
				else if (m_pMainBand != pBand)
					pBand->bpV1.m_vbVisible = VARIANT_FALSE;
			}

			if (m_pMainBand)
			{
				m_pMainBand->bpV1.m_vbVisible = VARIANT_TRUE;
				if (ddDAFloat == m_pMainBand->bpV1.m_daDockingArea)
					m_pMainBand->bpV1.m_rcFloat = pBand->bpV1.m_rcFloat;
				m_pMainBand->bpV1.m_daDockingArea = pBand->bpV1.m_daDockingArea;
				m_pMainBand->bpV1.m_nDockOffset = pBand->bpV1.m_nDockOffset;
				m_pMainBand->bpV1.m_nDockLine = pBand->bpV1.m_nDockLine;
				m_pBar->BuildAccelators(m_pMainBand);
			}
			m_pBar->RecalcLayout();
		}
		else
		{
			pBand = m_pBar->GetMenuBand();
			if (NULL == pBand)
				return FALSE;

			if (NULL == m_pMainBand)
				m_pMainBand = pBand;

			BSTR bstrBandName = NULL;
			if (!m_theChildMenus.Lookup(hWndChild, bstrBandName) && (NULL == bstrBandName || NULL == *bstrBandName))
				return FALSE;

			CBand* pNewBand = m_pBar->FindBand(bstrBandName);
			if (pNewBand && pNewBand != pBand)
			{
				pBand->bpV1.m_vbVisible = VARIANT_FALSE;
				pNewBand->bpV1.m_vbVisible = VARIANT_TRUE;
				pNewBand->bpV1.m_daDockingArea = pBand->bpV1.m_daDockingArea;
				if (ddDAFloat == pNewBand->bpV1.m_daDockingArea)
					pNewBand->bpV1.m_rcFloat = pBand->bpV1.m_rcFloat;
				pNewBand->bpV1.m_nDockOffset = pBand->bpV1.m_nDockOffset;
				pNewBand->bpV1.m_nDockLine = pBand->bpV1.m_nDockLine;
				m_pBar->BuildAccelators(pNewBand);
				m_pBar->RecalcLayout();
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
// RegisterChildWindow
//

BOOL CChildMenu::RegisterChildWindow(HWND hWndChild, BSTR bstrChildMenuBandName)
{
	BSTR bstrBandName = SysAllocString(bstrChildMenuBandName);
	if (NULL == bstrBandName)
		return FALSE;

	m_theChildMenus.SetAt(hWndChild, bstrBandName);
	return TRUE;
}

//
// UnregisterChildWindow
//

BOOL CChildMenu::UnregisterChildWindow(HWND hWndChild)
{
	BSTR bstrBandName = NULL;
	if (!m_theChildMenus.Lookup(hWndChild, bstrBandName))
		return FALSE;

	SysFreeString(bstrBandName);

	m_theChildMenus.RemoveKey(hWndChild);

	return TRUE;
}

//
// AddWindowList
//

BOOL CBar::AddWindowList(CTool* pWindowListTool, BandTypes btType, TypedArray<CTool*>& aTools)
{
	int nCount = m_aChildWindows.GetSize();
	if (nCount > 0)
	{
		HRESULT hResult;
		CTool* pTool;
		if (ddBTPopup == btType)
		{
			pTool = CTool::CreateInstance(NULL);
			if (NULL == pTool)
				return FALSE;

			pTool->tpV1.m_vbVisible = VARIANT_TRUE;
			pTool->tpV1.m_ttTools = ddTTSeparator;
			pTool->tpV1.m_nToolId = eToolIdSeparator;
			pTool->m_pBar = this;
			pTool->m_bCreatedInternally = TRUE;
			hResult = aTools.Add(pTool);
		}
		HWND hWndActive = (HWND)SendMessage(m_hWndMDIClient, WM_MDIGETACTIVE, 0, 0);
		HWND hWnd;
		TCHAR szBuffer[MAX_PATH];
		TCHAR szBuffer2[MAX_PATH];
		for (int nWnd = 0; nWnd < nCount; nWnd++)
		{
			hWnd = m_aChildWindows.GetAt(nWnd);

			if (!(WS_VISIBLE & GetWindowLong(hWnd, GWL_STYLE)))
				continue;

			if (!IsChild(m_hWndMDIClient, hWnd))
				continue;

			if (GetWindowText(hWnd, szBuffer, MAX_PATH) < 1)
				continue;

			pTool = CTool::CreateInstance(NULL);
			if (NULL == pTool)
				continue;

			pTool->m_bCreatedInternally = TRUE;
			pTool->m_pBar = this;
			pTool->tpV1.m_nToolId = eToolIdWindowList;
			pTool->tpV1.m_tsStyle = ddSIconText;
			hResult = pTool->put_Name(L"miChildWindow");
			
			pTool->tpV1.m_mvMenuVisibility = ddMVAlwaysVisible;
			if (hWnd == hWndActive)
				hResult = pTool->put_Checked(VARIANT_TRUE);

			pTool->m_hWndActive = hWnd;
			hResult = aTools.Add(pTool);
			if (ddBTNormal == btType)
			{
				MAKE_WIDEPTR_FROMTCHAR(wTitle, szBuffer);
				hResult = pTool->put_Caption(wTitle);

				// Try to load the small icon
				HICON hIcon = (HICON)SendMessage(hWnd, WM_GETICON, (WPARAM)ICON_SMALL, 0);
				if (hIcon)
				{
					pTool->tpV1.m_nImageWidth = GetSystemMetrics(SM_CXSMICON);
					pTool->tpV1.m_nImageHeight = GetSystemMetrics(SM_CYSMICON);
				}
				else
				{
					// Try to load big one
					hIcon = (HICON)SendMessage(hWnd, WM_GETICON, (WPARAM)ICON_BIG, 0);
					if (NULL == hIcon)
						// Last chance is the Window's icon
						hIcon = LoadIcon(NULL, MAKEINTRESOURCE(IDI_WINLOGO));
				}
				if (hIcon)
				{
					if (-1 != pWindowListTool->tpV1.m_nImageWidth)
						pTool->tpV1.m_nImageWidth = pWindowListTool->tpV1.m_nImageWidth;
					
					if (-1 != pWindowListTool->tpV1.m_nImageHeight)
						pTool->tpV1.m_nImageHeight = pWindowListTool->tpV1.m_nImageHeight;

					ICONINFO info;
					memset(&info, 0, sizeof(ICONINFO));
					info.fIcon = TRUE;
					if (GetIconInfo(hIcon, &info))
					{
						hResult = pTool->put_Bitmap(ddITNormal, (OLE_HANDLE)info.hbmColor);
						if (SUCCEEDED(hResult))
							hResult = pTool->put_MaskBitmap(ddITNormal, (OLE_HANDLE)info.hbmMask);
						BOOL bResult = DeleteBitmap(info.hbmColor);
						assert(bResult);
						bResult = DeleteBitmap(info.hbmMask);
						assert(bResult);
					}
				}
			}
			else
			{
				_stprintf(szBuffer2, _T("&%i %s"), nWnd + 1, szBuffer);
				MAKE_WIDEPTR_FROMTCHAR(wTitle, szBuffer2);
				hResult = pTool->put_Caption(wTitle);
			}
		}
	}
	return TRUE;
}

STDMETHODIMP CBar::MapPropertyToCategory( DISPID dispid,  PROPCAT*ppropcat)
{
	if (NULL == ppropcat) 
		return E_POINTER;

	switch (dispid) 
	{
	case DISPID_FORECOLOR:   
	case DISPID_BACKCOLOR:
	case DISPID_HIGHLIGHTCOLOR:
	case DISPID_SHADOWCOLOR:
	case DISPID_3DLIGHT:
	case DISPID_3DDARKSHADOW:
	case DISPID_PICTURE:
		*ppropcat = PROPCAT_Appearance;
		return S_OK;

	case DISPID_DCDATAPATH:
		*ppropcat = PROPCAT_Data;
		return S_OK;
		
	case DISPID_AUTOUPDATESTATUSBAR:
	case DISPID_DISPLAYKEYS:
	case DISPID_DISPLAYTOOLTIPS:
	case DISPID_FIREDBLCLICKEVENT:
	case DISPID_ENABLED:
	case DISPID_LARGEICONS:
	case DISPID_MENUANIMATION:
	case DISPID_MENUFONTSTYLE:
	case DISPID_PERSONALIZEDMENU:
	case DISPID_USERDEFINEDCUSTOMIZE:
		*ppropcat = PROPCAT_Behavior;
		return S_OK;

	case DISPID_FONT:
	case DISPID_CONTROLFONT:
	case DISPID_CHILDBANDFONT:
		*ppropcat = PROPCAT_Font;
		return S_OK;

	default:
		return E_FAIL;
	}
}
STDMETHODIMP CBar::GetCategoryName( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName)
{
	return E_FAIL;
}

//
// TabForward
//

BOOL CBar::TabForward(CBand*& pBand, HWND& hWndBand, CTool*& pTool, int& nTool)
{
	if (NULL == m_pTabbedTool)
		return FALSE;

	if (NULL == m_pActiveTool)
		m_pActiveTool = m_pTabbedTool;

	pBand = m_pActiveTool->m_pBand;
	if (NULL == pBand)
		return FALSE;

	CRect rcBand;
	hWndBand = NULL;
	if (!pBand->GetBandRect(hWndBand, rcBand))
		return FALSE;

	nTool = pBand->m_pTools->GetVisibleToolIndex(m_pActiveTool);

	int nCount = pBand->m_pTools->GetVisibleToolCount();

	if (nTool < nCount - 1)
		nTool++;
	else if (nCount > 0)
		nTool = 0;

	pTool = pBand->m_pTools->GetVisibleTool(nTool);
	while (ddTTSeparator == pTool->tpV1.m_ttTools || VARIANT_FALSE == pTool->tpV1.m_vbEnabled)
	{
		if (nTool >= nCount)
			nTool = 0;
		else
			nTool++;
		pTool = pBand->m_pTools->GetVisibleTool(nTool);
	}
	return TRUE;
}

//
// TabBackward
//

BOOL CBar::TabBackward(CBand*& pBand, HWND& hWndBand, CTool*& pTool, int& nTool)
{
	if (NULL == m_pTabbedTool)
		return FALSE;

	if (NULL == m_pActiveTool)
		m_pActiveTool = m_pTabbedTool;

	pBand = m_pActiveTool->m_pBand;
	if (NULL == pBand)
		return FALSE;

	CRect rcBand;
	hWndBand = NULL;
	if (!pBand->GetBandRect(hWndBand, rcBand))
		return FALSE;

	nTool = pBand->m_pTools->GetVisibleToolIndex(m_pActiveTool);

	int nCount = pBand->m_pTools->GetVisibleToolCount();

	//
	// Go in reverse
	//

	if (nTool > 0)
		nTool--;
	else
		nTool = nCount - 1;

	pTool = pBand->m_pTools->GetVisibleTool(nTool);
	while (ddTTSeparator == pTool->tpV1.m_ttTools || VARIANT_FALSE == pTool->tpV1.m_vbEnabled)
	{
		if (nTool > 1)
			nTool--;
		else
			nTool = nCount - 1;

		pTool = pBand->m_pTools->GetVisibleTool(nTool);
	}
	return TRUE;
}

//
// UpdateTools
//

BOOL CBar::UpdateTools(CBand* pBand, CTool* pTool, int& nTool, HWND hWndBand)
{
	if (NULL == pTool)
		return FALSE;

	if (m_pActiveTool)
	{
		switch (m_pActiveTool->tpV1.m_ttTools)
		{
		case ddTTEdit:
		case ddTTCombobox:
			break;

		case ddTTControl: 
			break;
		}
	}

	HideToolTips(TRUE);
	KillTimer(m_hWnd, eFlyByTimer);
	MSG msg;
	while (::PeekMessage(&msg, NULL, WM_MOUSEMOVE, WM_MOUSEMOVE, PM_NOREMOVE))
	{
		if (!GetMessage(&msg, NULL, WM_MOUSEMOVE, WM_MOUSEMOVE))
			break;
	}

	m_pActiveTool = pTool;

	if (-1 != pBand->m_nCurrentTool)
		pBand->InvalidateToolRect(pBand->m_nCurrentTool);

	pBand->m_nToolMouseOver = pBand->m_nCurrentTool = nTool;

	if (m_pActiveTool)
	{
		switch (m_pActiveTool->tpV1.m_ttTools)
		{
		case ddTTControl:
			::SetFocus(m_pActiveTool->m_hWndActive);
			break;

		default:
			pBand->InvalidateToolRect(pBand->m_nCurrentTool);
		}
	}
	
	return TRUE;
}

//
// PositionAlignableControls
//

BOOL CBar::PositionAlignableControls(SIZEPARENTPARAMS& sppLayout)
{
	if (NULL == m_pClientSite)
		return FALSE;

	const ULONG nNumberOfUnknowns = 10;
	LPOLECONTROLSITE pControlSite = NULL;
	LPOLECLIENTSITE pClientSite = NULL;
	IVBGetControl* pVBGetControl = NULL;
	LPENUMUNKNOWN pEnumObject = NULL;
	LPOLEOBJECT pObject = NULL;
	LPDISPATCH pDispatch;
	IActiveBar2* pBar = NULL;
	LPUNKNOWN pControlUnknown[nNumberOfUnknowns];
	ULONG nNumberReturned;
	ULONG nControl;
	DWORD dwStatus;
	BOOL  bResult = TRUE;
	VARIANT vProperty;
	CRect rcControl;
	SIZE size;
	int nAlign;

	try
	{
		HRESULT hResult = m_pClientSite->QueryInterface(IID_IVBGetControl, (void**)&pVBGetControl);
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hResult = pVBGetControl->EnumControls(OLECONTF_EMBEDDINGS | OLECONTF_OTHERS, GC_WCH_SIBLING, &pEnumObject);
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}
 
		hResult = pEnumObject->Next(nNumberOfUnknowns, (IUnknown**)&pControlUnknown, &nNumberReturned);
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}

		if (nNumberReturned < 1)
		{
			bResult = FALSE;
			goto Cleanup;
		}

		while (nNumberReturned > 0)
		{
			for (nControl = 0; nControl < nNumberReturned; nControl++)
			{
				hResult = pControlUnknown[nControl]->QueryInterface(IID_IOleObject, (void**)&pObject);
				if (FAILED(hResult))
				{
					//
					// This is an internal control
					//

					hResult = pControlUnknown[nControl]->QueryInterface(IID_IDispatch, (void**)&pDispatch);
					pControlUnknown[nControl]->Release();
					if (FAILED(hResult))
						continue;
				}
				else
				{
					pControlUnknown[nControl]->Release();

					hResult = pObject->QueryInterface(IID_IActiveBar2, (void**)&pBar);
					if (SUCCEEDED(hResult))
					{
						if ((CBar*)pBar == this)
						{
							pBar->Release();
							pObject->Release();
							continue;
						}
						pBar->Release();
					}

					hResult = pObject->GetMiscStatus(DVASPECT_CONTENT, &dwStatus);
					if (FAILED(hResult))
					{
						pObject->Release();
						continue;
					}

					if (OLEMISC_INVISIBLEATRUNTIME & dwStatus || !(OLEMISC_ALIGNABLE & dwStatus))
					{
						pObject->Release();
						continue;
					}

					hResult = pObject->QueryInterface(IID_IDispatch, (void**)&pDispatch);
					if (FAILED(hResult))
					{
						pObject->Release();
						continue;
					}

					pDispatch->Release();

					hResult = pObject->GetClientSite(&pClientSite);
					pObject->Release();
					if (FAILED(hResult))
						continue;

					hResult = pClientSite->QueryInterface(IID_IOleControlSite, (void**)&pControlSite);
					pClientSite->Release();
					if (FAILED(hResult))
						continue;

					hResult = pControlSite->GetExtendedControl(&pDispatch);
					pControlSite->Release();
					if (FAILED(hResult))
						continue;
				}

				hResult = PropertyGet(pDispatch, L"Visible", vProperty);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				if (VARIANT_FALSE == vProperty.boolVal)
				{
					pDispatch->Release();
					continue;
				}

				hResult = PropertyGet(pDispatch, L"Align", vProperty);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				nAlign = vProperty.iVal;

				hResult = PropertyGet(pDispatch, L"Left", vProperty);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				size.cx = vProperty.lVal;

				hResult = PropertyGet(pDispatch, L"Top", vProperty);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				size.cy = vProperty.lVal;

				TwipsToPixel(&size, &size);

				rcControl.left = size.cx;
				rcControl.top = size.cy;

				hResult = PropertyGet(pDispatch, L"Width", vProperty);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}
				
				hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				size.cx = vProperty.lVal;

				hResult = PropertyGet(pDispatch, L"Height", vProperty);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				hResult = VariantChangeType(&vProperty, &vProperty, VARIANT_NOVALUEPROP, VT_I4);
				if (FAILED(hResult))
				{
					pDispatch->Release();
					continue;
				}

				size.cy = vProperty.lVal;

				TwipsToPixel(&size, &size);

				rcControl.right = rcControl.left + size.cx;
				rcControl.bottom = rcControl.top + size.cy;

				switch (nAlign)
				{
				case 0: // None
					{
					}
					break;

				case 1: // Top
					{
						if (0 != rcControl.Height() && 0 != rcControl.Width())
						{
							sppLayout.rc.top += rcControl.Height();
							sppLayout.sizeTotal.cy += rcControl.Height();
						}
					}
					break;

				case 2: // Bottom
					{
						if (0 != rcControl.Height() && 0 != rcControl.Width())
						{
							sppLayout.rc.bottom -= rcControl.Height();
							sppLayout.sizeTotal.cy += rcControl.Height();
						}
					}
					break;

				case 3: // Left
					{
						if (0 != rcControl.Height() && 0 != rcControl.Width())
						{
							sppLayout.rc.left += rcControl.Width();
							sppLayout.sizeTotal.cx += rcControl.Width();
						}
					}
					break;

				case 4: // Right
					{
						if (0 != rcControl.Height() && 0 != rcControl.Width())
						{
							sppLayout.rc.right -= rcControl.Width();
							sppLayout.sizeTotal.cx += rcControl.Width();
						}
					}
					break;
				}
				pDispatch->Release();
			}
			hResult = pEnumObject->Next(nNumberOfUnknowns, (IUnknown**)&pControlUnknown, &nNumberReturned);
			if (FAILED(hResult))
				nNumberReturned = 0;
		}
	Cleanup:
		if (pVBGetControl)
			pVBGetControl->Release();
		if (pEnumObject)
			pEnumObject->Release();
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return FALSE;
}


ActiveEdit* CBar::GetActiveEdit()
{
	if (NULL == m_pEdit || !m_pEdit->IsWindow())
	{
		m_pEdit = new ActiveEdit();
		if (NULL == m_pEdit)
			return NULL;
		CRect rc;
		m_pEdit->Create(m_hWnd, rc, VARIANT_TRUE == bpV1.m_vbXPLook);
		HFONT hFontCtrl = m_fhControl.GetFontHandle();
		if (hFontCtrl)
			m_pEdit->SendMessage(WM_SETFONT, (WPARAM)hFontCtrl, FALSE);
	}
	return m_pEdit;
}

ActiveCombobox* CBar::GetActiveCombo(DWORD dwStyle)
{
	if (NULL != m_pCombo)
	{
		if (m_pCombo->IsWindow())
		{
			if (m_pCombo->Style() != dwStyle)
				m_pCombo->DestroyWindow();
			m_pCombo = NULL;
		}
		else
			m_pCombo = NULL;
	}
	if (NULL == m_pCombo)
	{
		m_pCombo = new ActiveCombobox();
		if (NULL == m_pCombo)
			return NULL;
		CRect rc;
		m_pCombo->Create(m_hWnd, rc, dwStyle);
		HFONT hFontCtrl = m_fhControl.GetFontHandle();
		if (hFontCtrl)
			m_pCombo->SendMessage(WM_SETFONT, (WPARAM)hFontCtrl, FALSE);
	}
	return m_pCombo;
}
















