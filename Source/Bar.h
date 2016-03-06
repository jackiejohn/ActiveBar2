#ifndef __CBAR_H__
#define __CBAR_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Interfaces.h"
#include "PrivateInterfaces.h"
#include "Support.h"
#include "Flicker.h"
#include "DDAsync.h"
#include "Map.h"

#define GC_WCH_SIBLING     0x00000001L
#define GC_WCH_CONTAINER    0x00000002L   // no FONLYNEXT/PREV
#define GC_WCH_CONTAINED    0x00000003L   // no FONLYNEXT/PREV
#define GC_WCH_ALL     0x00000004L
#define GC_WCH_FREVERSEDIR  0x08000000L   // OR'd with others
#define GC_WCH_FONLYNEXT    0x10000000L   // OR'd with others
#define GC_WCH_FONLYPREV    0x20000000L   // OR'd with others
#define GC_WCH_FSELECTED    0x40000000L   // OR'd with others

DEFINE_GUID(IID_IVBGetControl, 0x40A050A0L, 0x3C31, 0x101B, 0xA8, 0x2E, 0x08, 0x00, 0x2B, 0x2B, 0x23, 0x37);

DECLARE_INTERFACE_(IVBGetControl, IUnknown)
{
    // *** IUnknown methods ****
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;

    // *** IVBGetControl methods ****
    STDMETHOD(EnumControls)(THIS_ DWORD dwOleContF, DWORD dwWhich,
                            LPENUMUNKNOWN FAR *ppenumUnk) PURE;
};

interface IDesigner;
struct SIZEPARENTPARAMS;
class CPopupWinToolAndShadow;
class CColorQuantizer;
class CToolShortCut;
class CDragDropMgr;
class ShortCutStore;
//class CAccessible;
class CShortCuts;
class ActiveEdit;
class ActiveCombobox;
class CImageMgr;
class CToolTip;
class CDockMgr;
class CTools;
class CTool;
class CBands;
class CBand;
class CDock;
class CBar;

//
// ToolStack
//

class ToolStack
{
public:
	BOOL Push(CTool* pTool);
	CTool* Pop();
private:
	TypedArray<CTool*> m_aTools;
};

inline BOOL ToolStack::Push(CTool* pTool)
{
	return SUCCEEDED(m_aTools.Add(pTool));
}

inline CTool* ToolStack::Pop()
{
	int nCount = m_aTools.GetSize();
	if (nCount < 1)
		return NULL;
	nCount--;
	CTool* pTool = m_aTools.GetAt(nCount);
	m_aTools.RemoveAt(nCount);
	return pTool;
}

//
// ActiveBarErrorObject
//

class ActiveBarErrorObject : public ErrorObject
{
	virtual void AsyncError(short		   nNumber, 
						    IReturnString* pDescription, 
						    long		   nSCode, 
						    BSTR		   bstrSource, 
						    BSTR		   bstrHelpFile, 
						    long		   nHelpContext, 
						    IReturnBool*   pCancelDisplay);
};

//
// CChildRelocate
//

class CChildRelocate
{
public:
	CChildRelocate();
	~CChildRelocate();

	void CacheFontSize();

	HRESULT ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave);

	enum ControlType
	{
		eNormal,
		eToolBar,
		eCombobox,
		eListbox,
		eHWnd,
	};

#pragma pack(push)
#pragma pack (1)
	struct RelocatePropV1
	{
		ControlType m_eControlType;
		CRect       m_rcPerOfParent;
		int         m_nHeight;
		
		// Font Resizing
		BOOL m_bHasFontProperty;
		CY m_sizeBaseFont;
		float m_fBaseWidth;
		float m_fBaseHeight;
	} crProperies;
#pragma pack(pop)
	HWND		m_hWnd;
	LPDISPATCH  m_pDispatch; 
};

//
// CRelocate
//

class CRelocate
{
public:
	CRelocate(CBar* pBar);
	~CRelocate();

	BOOL LayoutProportional (HWND hWnd, CRect rc);
	BOOL LayoutClient(HWND hWnd, CRect rc);
	BOOL GetLayoutInfo (HWND hWnd, const CRect& rc);
	BOOL GetControlInfo();
	HRESULT ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave);
	BOOL FindControl(LPDISPATCH pDispatch);
	int Count();

	HRESULT AddControl(LPDISPATCH pDispatch);
	HRESULT AddControl(HWND hWnd);
	void Cleanup();

private:
	TypedArray<CChildRelocate*> m_aChildRelocate;
	CBar* m_pBar;
};

inline int CRelocate::Count()
{
	return m_aChildRelocate.GetSize();
}

//
// FormsAndControls
//

class FormsAndControls
{
public:
	BOOL Add(CTool* pTool);
	int  Count();
	BOOL Remove(CTool* pTool);
	BOOL FindControl(LPDISPATCH pDispatch);
	BOOL FindControl(HWND hWnd, CTool*& pTool);
	HWND FindChildFromPoint(POINT ptScreen, CTool** ppTool = NULL);
	BOOL IsChild(HWND hWnd);
	void ShutDown();
	void EnableForms(BOOL bEnable);

private:
	TypedArray<CTool*> m_aFormsAndControls;
};

//
// Count
//

inline int FormsAndControls::Count()
{
	return m_aFormsAndControls.GetSize();
}

//
// Add
//

inline BOOL FormsAndControls::Add(CTool* pTool)
{
	HRESULT hResult = m_aFormsAndControls.Add(pTool);
	return SUCCEEDED(hResult);
}

//
// CChildMenu
//

class CChildMenu
{
public:
	CChildMenu();
	~CChildMenu();

	void SetBar(CBar* pBar);
	void MainBand(CBand* pBand);
	CBand* MainBand();

	BOOL SwitchMenus(HWND hWndChild);
	BOOL RegisterChildWindow(HWND hWndChild, BSTR bstrChildMenuBandName);
	BOOL UnregisterChildWindow(HWND hWndChild);

private:
	TypedMap<HWND, BSTR> m_theChildMenus;
	CBand* m_pMainBand;
	CBar*  m_pBar;
};

inline void CChildMenu::MainBand(CBand* pBand)
{
	m_pMainBand = pBand;
}

inline CBand* CChildMenu::MainBand()
{
	return m_pMainBand;
}

enum TOOLNF
{
	TNF_VIEWCHANGED,
	TNF_ACTIVETOOLCHECK
};

class CDesignTimeSystemSettingNotify : public FWnd
{
public:
	CDesignTimeSystemSettingNotify(CBar* pBar);
	~CDesignTimeSystemSettingNotify();

	BOOL Create();
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	CBar* m_pBar;
};

class CBar : public IOleObject,public IOleInPlaceObject,public IOleInPlaceActiveObject,public IViewObject2,public IPersistPropertyBag,public IPersistStreamInit,public IPersistStorage,public IConnectionPointContainer,public ISpecifyPropertyPages,public IProvideClassInfo,public IBarPrivate,public IOleControl,public ISupportErrorInfo,public ICategorizeProperties,public IActiveBar2,public IDDPerPropertyBrowsing
{
public:
	CBar(IUnknown *pUnkOuter);
	int m_objectIndex;
	~CBar();
	
	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CBar *CreateInstance(IUnknown *pUnkOuter);
	// IUnknown members
	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
	STDMETHOD_(ULONG, AddRef)(void);
	STDMETHOD_(ULONG, Release)(void);
//	DECLARE_STANDARD_UNKNOWN();

	// IOleObject members

	STDMETHOD(SetClientSite)( IOleClientSite *pClientSite);
	STDMETHOD(GetClientSite)( IOleClientSite **ppClientSite);
	STDMETHOD(SetHostNames)( LPCOLESTR szContainerApp, LPCOLESTR szContainerObj);
	STDMETHOD(Close)( DWORD dwSaveOption);
	STDMETHOD(SetMoniker)( DWORD dwWhichMoniker, IMoniker *pmk);
	STDMETHOD(GetMoniker)( DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk);
	STDMETHOD(InitFromData)( IDataObject *pDataObject, BOOL fCreation, DWORD reserved);
	STDMETHOD(GetClipboardData)( DWORD dwReserved, IDataObject **ppDataObject);
	STDMETHOD(DoVerb)( LONG lVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect);
	STDMETHOD(EnumVerbs)( IEnumOLEVERB **ppEnumOleVerb);
	STDMETHOD(Update)();
	STDMETHOD(IsUpToDate)();
	STDMETHOD(GetUserClassID)( CLSID *pClsid);
	STDMETHOD(GetUserType)( DWORD dwFormOfType, LPOLESTR *pszUserType);
	STDMETHOD(SetExtent)( DWORD dwDrawAspect,  SIZEL *psizel);
	STDMETHOD(GetExtent)( DWORD dwDrawAspect,  SIZEL *psizel);
	STDMETHOD(Advise)( IAdviseSink *pAdvSink,  DWORD *pdwConnection);
	STDMETHOD(Unadvise)(DWORD dwConnection);
	STDMETHOD(EnumAdvise)(IEnumSTATDATA **ppenumAdvise);
	STDMETHOD(GetMiscStatus)( DWORD dwAspect,  DWORD *pdwStatus);
	STDMETHOD(SetColorScheme)( LOGPALETTE *pLogpal);
	// DEFS for IOleObject
	IOleClientSite *m_pClientSite;
	IOleControlSite *m_pControlSite;
	ISimpleFrameSite *m_pSimpleFrameSite;
	IOleAdviseHolder *m_pOleAdviseHolder;
	IOleInPlaceFrame   *m_pInPlaceFrame;
    IOleInPlaceUIWindow *m_pInPlaceUIWindow;
	IDispatch *m_pDispAmbient;

	unsigned m_fDirty:1;  // control need to be resaved?
    unsigned m_fInPlaceActive:1;  
    unsigned m_fInPlaceVisible:1; // in place visible ?
    unsigned m_fUIActive:1; // UI active ?
    unsigned m_fCreatingWindow:1; // indicates if in CreateWindowEx
	unsigned m_fCheckedReflecting:1;
	unsigned m_fHostReflects:1;
	
	// inplace activation stuff
	HWND m_hWnd;
	HWND m_hWndParent;
	HWND m_hWndReflect;
	SIZEL m_Size; // current size of control
	IOleInPlaceSite    *m_pInPlaceSite;
	IOleInPlaceSiteWindowless *m_pInPlaceSiteWndless;
	void ViewChanged();	
	BOOL OnSetExtent(SIZEL *);
	HRESULT InPlaceActivate(LONG lVerb);
	HWND CreateInPlaceWindow(int  x,int  y,BOOL fNoRedraw);
	virtual BOOL RegisterClassData(void);
	static LRESULT CALLBACK ReflectWindowProc(HWND hwnd,UINT msg,WPARAM  wParam,LPARAM  lParam);
	static LRESULT CALLBACK ControlWindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
    static CBar *ControlFromHwnd(HWND hwnd) {return (CBar *) GetWindowLong(hwnd, GWL_USERDATA);}
    virtual LRESULT WindowProc(UINT msg, WPARAM wParam, LPARAM lParam);

	BOOL BeforeCreateWindow(DWORD *pdwWindowStyle,DWORD *pdwExWindowStyle,LPTSTR  pszWindowTitle);
	BOOL AfterCreateWindow();
	void BeforeDestroyWindow();
	BOOL SetFocus(BOOL fGrab);
	void SetInPlaceVisible(BOOL);
	HRESULT OnVerb(LONG lVerb);
	BOOL GetAmbientProperty(DISPID  dispid,VARTYPE vt,void *pData);

	inline HWND GetOuterWindow(void) {return (m_hWndReflect) ? m_hWndReflect : m_hWnd;}
    inline BOOL Windowless(void) {return !m_fInPlaceActive || m_pInPlaceSiteWndless;}
    inline IOleInPlaceSite *GetInPlaceSite(void) {return (IOleInPlaceSite *)(m_pInPlaceSiteWndless ? m_pInPlaceSiteWndless : m_pInPlaceSite);}

	void ModalDialog(BOOL fShow);
	
	

	// IOleWindow members

	STDMETHOD(GetWindow)( HWND *phwnd);
	STDMETHOD(ContextSensitiveHelp)( BOOL fEnterMode);
	// DEFS for IOleWindow
	#define DEF_IOLEWINDOW
	
	

	// IOleInPlaceObject members

	STDMETHOD(InPlaceDeactivate)();
	STDMETHOD(UIDeactivate)();
	STDMETHOD(SetObjectRects)( LPCRECT lprcPosRect, LPCRECT lprcClipRect);
	STDMETHOD(ReactivateAndUndo)();
	// DEFS for IOleInPlaceObject
	#define DEF_IOLEINPLACEOBJECT
	unsigned m_fUsingWindowRgn:1; 
    HRGN m_hRgn;
	
	

	// IOleInPlaceActiveObject members

	STDMETHOD(TranslateAccelerator)( LPMSG lpmsg);
	STDMETHOD(OnFrameWindowActivate)( BOOL fActivate);
	STDMETHOD(OnDocWindowActivate)( BOOL fActivate);
	STDMETHOD(ResizeBorder)( LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow);
	STDMETHOD(EnableModeless)( BOOL fEnable);
	// DEFS for IOleInPlaceActiveObject
	virtual BOOL OnSpecialKey(LPMSG);	
	

	// IViewObject members

	STDMETHOD(Draw)( DWORD dwDrawAspect, LONG lindex, void *pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDev, HDC hdcDraw, LPCRECTL lprcBounds, LPCRECTL lprcWBounds, BOOL ( STDMETHODCALLTYPE *pfnContinue )(DWORD dwContinue), DWORD dwContinue);
	STDMETHOD(GetColorSet)( DWORD dwDrawAspect, LONG lindex, void *pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDev, LOGPALETTE **ppColorSet);
	STDMETHOD(Freeze)( DWORD dwDrawAspect, LONG lindex, void *pvAspect, DWORD *pdwFreeze);
	STDMETHOD(Unfreeze)( DWORD dwFreeze);
	STDMETHOD(SetAdvise)( DWORD dwAspects, DWORD dwAdviseFlags, IAdviseSink *pAdviseSink);
	STDMETHOD(GetAdvise)( DWORD *pdwAspects, DWORD *pdwAdviseFlags, IAdviseSink **ppAdviseSink);
	// DEFS for IViewObject
	RECT m_rcLocation; // set by IOleViewObject::Draw
    unsigned m_fViewAdvisePrimeFirst: 1;
    unsigned m_fViewAdviseOnlyOnce: 1;
	IAdviseSink *m_pViewAdviseSink;
	STDMETHOD(OnDraw)(DWORD dvAspect, HDC hdcDraw, LPCRECTL prcBounds, LPCRECTL prcWBounds, HDC hicTargetDev, BOOL fOptimize);
	BOOL OnGetPalette(HDC hicTargetDevice,LOGPALETTE **ppColorSet);
#ifdef CBar_SUBCLASS
	HRESULT DoSuperClassPaint(HDC hdc,LPCRECTL prcBounds);
#endif
		
	

	// IViewObject2 members

	STDMETHOD(GetExtent)( DWORD dwDrawAspect, LONG lindex, DVTARGETDEVICE *ptd, LPSIZEL lpsizel);
	// DEFS for IViewObject2
	
	

	// IPersist members

	STDMETHOD(GetClassID)( LPCLSID lpClassID);
	// DEFS for IPersist
	#define DEF_PERSIST
	
	

	// IPersistPropertyBag members

	STDMETHOD(Load)( IPropertyBag *pPropBag, IErrorLog *pErrorLog);
	STDMETHOD(Save)( IPropertyBag *pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties);
	// DEFS for IPersistPropertyBag
	HRESULT PersistText(IPropertyBag *pPropBag, IErrorLog *pErrorLog,BOOL save);	
	

	// IPersistStreamInit members

	STDMETHOD(IsDirty)();
	STDMETHOD(Load)( LPSTREAM pStm);
	STDMETHOD(Save)( LPSTREAM pStm, BOOL fClearDirty);
	STDMETHOD(GetSizeMax)( ULARGE_INTEGER *pCbSize);
	STDMETHOD(InitNew)();
	// DEFS for IPersistStreamInit
	#define DEF_PERSISTSTREAM
	HRESULT LoadStandardState(IStream *pStream);
	HRESULT PersistBinaryState(IStream *pStream,BOOL save);
	
	

	// IPersistStorage members

	STDMETHOD(InitNew)( IStorage *pStg);
	STDMETHOD(Load)( IStorage *pStg);
	STDMETHOD(Save)( IStorage *pStgSave, BOOL fSameAsLoad);
	STDMETHOD(SaveCompleted)( IStorage *pStgNew);
	STDMETHOD(HandsOffStorage)();
	// DEFS for IPersistStorage
	unsigned m_fSaveSucceeded:1;
	BOOL InitializeNewState();
	HRESULT m_SaveToStream(IStream *pStream);
	HRESULT SaveStandardState(IStream *pStream);
		
	

	// IConnectionPointContainer members

	STDMETHOD(EnumConnectionPoints)( IEnumConnectionPoints **ppEnum);
	STDMETHOD(FindConnectionPoint)( REFIID riid,  IConnectionPoint **ppCP);
	// DEFS for IConnectionPointContainer
	#define SINK_TYPE_EVENT      0
	#define SINK_TYPE_PROPNOTIFY 1
	
	class CConnectionPoint : public IConnectionPoint 
	{
	public:
        IUnknown **m_rgSinks;
        BYTE   m_bType;
		int inFireEvent;
        // IUnknown methods
        //
        STDMETHOD(QueryInterface)(THIS_ REFIID riid, LPVOID FAR* ppvObj) ;
        STDMETHOD_(ULONG,AddRef)(THIS) ;
        STDMETHOD_(ULONG,Release)(THIS) ;

        // IConnectionPoint methods
        //
        STDMETHOD(GetConnectionInterface)(IID FAR* pIID);
        STDMETHOD(GetConnectionPointContainer)(IConnectionPointContainer FAR* FAR* ppCPC);
        STDMETHOD(Advise)(LPUNKNOWN pUnkSink, DWORD FAR* pdwCookie);
        STDMETHOD(Unadvise)(DWORD dwCookie);
        STDMETHOD(EnumConnections)(LPENUMCONNECTIONS FAR* ppEnum);

        void    DoInvoke(DISPID dispid,BYTE* pbParams,...);
        void    DoInvokeV(DISPID dispid,BYTE* pbParams,va_list argList);
        void    DoOnChanged(DISPID dispid);
        BOOL    DoOnRequestEdit(DISPID dispid);
        HRESULT AddSink(void *, DWORD *);

        CBar *m_pControl;
        CConnectionPoint(CBar *pControl,BYTE type)
		{
			m_refCount=1;
			m_pControl=pControl;
            m_rgSinks = NULL;
            m_cSinks = 0;
			inFireEvent=0;
			m_bType=type;
        }
        ~CConnectionPoint();
		void DisconnectAll();
      private:
        short  m_cSinks;
		ULONG m_refCount;

    } *m_cpEvents, *m_cpPropNotify;
	
	void BoundPropertyChanged(DISPID dispid);

    friend CConnectionPoint;
	
	

	// ISpecifyPropertyPages members

	STDMETHOD(GetPages)( CAUUID *pPages);
	// DEFS for ISpecifyPropertyPages
	
	

	// IProvideClassInfo members

	STDMETHOD(GetClassInfo)( ITypeInfo **ppTI);
	// DEFS for IProvideClassInfo
	
	

	// IBarPrivate members

	STDMETHOD(RegisterBandChange)( IDesignerNotifications*pDesignerNotify);
	STDMETHOD(RevokeBandChange)();
	STDMETHOD(DesignerInitialize)(IDragDropManager*pDragDropManager,  IDesignerNotify*pDesigner);
	STDMETHOD(DesignerShutdown)();
	STDMETHOD(SetDesignerModified)();
	STDMETHOD(ExchangeToolById)( IUnknown*pStream,  VARIANT_BOOL vbSave, IDispatch**ppTool);
	STDMETHOD(ExchangeToolByBandChildBandToolId)( IUnknown*pStream,  BSTR bstrBand, BSTR bstrPage,  VARIANT_BOOL vbSave, IDispatch**ppTool);
	STDMETHOD(ShutDownAdvise)( OLE_HANDLE hWnd);
	STDMETHOD(get_Library)(VARIANT_BOOL *retval);
	STDMETHOD(put_Library)(VARIANT_BOOL val);
	STDMETHOD(put_CustomizeDragLock)(LPDISPATCH val);
	STDMETHOD(VerifyAndCorrectImages)();
	STDMETHOD(Attach)( OLE_HANDLE hWndParent);
	STDMETHOD(Detach)();
	STDMETHOD(DeactivateWindow)();
	STDMETHOD(get_PrivateHwnd)(OLE_HANDLE *retval);
	STDMETHOD(ExchangeToolByIdentity)( IUnknown*pStream,  VARIANT_BOOL vbSave, IDispatch**ppTool);
	STDMETHOD(ExchangeToolByBandChildBandToolIdentity)( IUnknown*pStream,  BSTR bstrBand, BSTR bstrPage,  VARIANT_BOOL vbSave, IDispatch**ppTool);
	// DEFS for IBarPrivate
	
	

	// IOleControl members

	STDMETHOD(GetControlInfo)( CONTROLINFO *pCI);
	STDMETHOD(OnMnemonic)( MSG *pMsg);
	STDMETHOD(OnAmbientPropertyChange)( DISPID dispID);
	STDMETHOD(FreezeEvents)( BOOL bFreeze);
	// DEFS for IOleControl
	virtual void AmbientPropertyChanged(DISPID dispid);
	BOOL m_fModeFlagValid;	
	

	// ISupportErrorInfo members

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid);
	// DEFS for ISupportErrorInfo
	
	

	// ICategorizeProperties members

	STDMETHOD(MapPropertyToCategory)( DISPID dispid,  PROPCAT*ppropcat);
	STDMETHOD(GetCategoryName)( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName);
	// DEFS for ICategorizeProperties
	
	

	// IDispatch members

	STDMETHOD(GetTypeInfoCount)( UINT *pctinfo);
	STDMETHOD(GetTypeInfo)( UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
	STDMETHOD(GetIDsOfNames)( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid);
	STDMETHOD(Invoke)( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr);
	// DEFS for IDispatch
	
	

	// IActiveBar2 members

	STDMETHOD(get_Bands)(Bands * *retval);
	STDMETHOD(RecalcLayout)();
	STDMETHOD(get_DisplayToolTips)(VARIANT_BOOL *retval);
	STDMETHOD(put_DisplayToolTips)(VARIANT_BOOL val);
	STDMETHOD(get_DisplayKeysInToolTip)(VARIANT_BOOL *retval);
	STDMETHOD(put_DisplayKeysInToolTip)(VARIANT_BOOL val);
	STDMETHOD(get_Tools)(Tools * *retval);
	STDMETHOD(get_ActiveBand)(Band * *retval);
	STDMETHOD(put_ActiveBand)(Band * val);
	STDMETHOD(Customize)( CustomizeTypes nCustomize);
	STDMETHOD(ReleaseFocus)();
	STDMETHOD(Save)( BSTR BandName,  BSTR FileName,  SaveOptionTypes SerializeOptions,  VARIANT*vData);
	STDMETHOD(Load)( BSTR BandName,  VARIANT *vData,  SaveOptionTypes SerializeOptions);
	STDMETHOD(SaveLayoutChanges)( BSTR FileName,  SaveOptionTypes SerializeOptions,  VARIANT *vData);
	STDMETHOD(LoadLayoutChanges)( VARIANT *vData,  SaveOptionTypes SerializeOptions);
	STDMETHOD(get_Font)(LPFONTDISP *retval);
	STDMETHOD(put_Font)(LPFONTDISP val);
	STDMETHOD(putref_Font)(LPFONTDISP *val);
	STDMETHOD(About)();
	STDMETHOD(get_DataPath)(BSTR *retval);
	STDMETHOD(put_DataPath)(BSTR val);
	STDMETHOD(get_MenuFontStyle)(MenuFontStyles *retval);
	STDMETHOD(put_MenuFontStyle)(MenuFontStyles val);
	STDMETHOD(get_Picture)(LPPICTUREDISP *retval);
	STDMETHOD(put_Picture)(LPPICTUREDISP val);
	STDMETHOD(putref_Picture)(LPPICTUREDISP *val);
	STDMETHOD(get_BackColor)(OLE_COLOR *retval);
	STDMETHOD(put_BackColor)(OLE_COLOR val);
	STDMETHOD(get_ForeColor)(OLE_COLOR *retval);
	STDMETHOD(put_ForeColor)(OLE_COLOR val);
	STDMETHOD(get_HighlightColor)(OLE_COLOR *retval);
	STDMETHOD(put_HighlightColor)(OLE_COLOR val);
	STDMETHOD(get_ShadowColor)(OLE_COLOR *retval);
	STDMETHOD(put_ShadowColor)(OLE_COLOR val);
	STDMETHOD(get_ControlFont)(LPFONTDISP *retval);
	STDMETHOD(put_ControlFont)(LPFONTDISP val);
	STDMETHOD(putref_ControlFont)(LPFONTDISP *val);
	STDMETHOD(Refresh)();
	STDMETHOD(OnSysColorChanged)();
	STDMETHOD(get_ChildBandFont)(LPFONTDISP *retval);
	STDMETHOD(put_ChildBandFont)(LPFONTDISP val);
	STDMETHOD(putref_ChildBandFont)(LPFONTDISP *val);
	STDMETHOD(PlaySound)(BSTR szSound, SoundTypes stType);
	STDMETHOD(get_MenuAnimation)(MenuStyles *retval);
	STDMETHOD(put_MenuAnimation)(MenuStyles val);
	STDMETHOD(get_LargeIcons)(VARIANT_BOOL *retval);
	STDMETHOD(put_LargeIcons)(VARIANT_BOOL val);
	STDMETHOD(get_ImageManager)(VARIANT *retval);
	STDMETHOD(put_ImageManager)(VARIANT val);
	STDMETHOD(get_WhatsThisHelpMode)(VARIANT_BOOL *retval);
	STDMETHOD(put_WhatsThisHelpMode)(VARIANT_BOOL val);
	STDMETHOD(get_AlignToForm)(VARIANT_BOOL *retval);
	STDMETHOD(put_AlignToForm)(VARIANT_BOOL val);
	STDMETHOD(get_AutoSizeChildren)(AutoSizeChildrenTypes *retval);
	STDMETHOD(put_AutoSizeChildren)(AutoSizeChildrenTypes val);
	STDMETHOD(Localize)(LocalizationTypes Index, BSTR LocaleString);
	STDMETHOD(RegisterChildMenu)( OLE_HANDLE hWndChild,  BSTR strChildMenuName);
	STDMETHOD(get_Version)(BSTR *retval);
	STDMETHOD(get_AutoUpdateStatusBar)(VARIANT_BOOL *retval);
	STDMETHOD(put_AutoUpdateStatusBar)(VARIANT_BOOL val);
	STDMETHOD(get_PersonalizedMenus)(PersonalizedMenuTypes *retval);
	STDMETHOD(put_PersonalizedMenus)(PersonalizedMenuTypes val);
	STDMETHOD(SaveMenuUsageData)( BSTR szFileName,  SaveOptionTypes SerializeOptions,  VARIANT *vData);
	STDMETHOD(LoadMenuUsageData)( VARIANT *vData,  SaveOptionTypes SerializeOptions);
	STDMETHOD(ClearMenuUsageData)();
	STDMETHOD(putref_ClientAreaControl)(LPDISPATCH *retval);
	STDMETHOD(put_ClientAreaControl)(LPDISPATCH val);
	STDMETHOD(get_UserDefinedCustomization)(VARIANT_BOOL *retval);
	STDMETHOD(put_UserDefinedCustomization)(VARIANT_BOOL val);
	STDMETHOD(GetToolFromPosition)( long x,  long y,  Tool**tool);
	STDMETHOD(GetBandFromPosition)( long x,  long y,  Band**band);
	STDMETHOD(get_FireDblClickEvent)(VARIANT_BOOL *retval);
	STDMETHOD(put_FireDblClickEvent)(VARIANT_BOOL val);
	STDMETHOD(get_ThreeDLight)(OLE_COLOR *retval);
	STDMETHOD(put_ThreeDLight)(OLE_COLOR val);
	STDMETHOD(get_ThreeDDarkShadow)(OLE_COLOR *retval);
	STDMETHOD(put_ThreeDDarkShadow)(OLE_COLOR val);
	STDMETHOD(ApplyAll)( Tool *tool);
	STDMETHOD(get_BackgroundOption)(short *retval);
	STDMETHOD(put_BackgroundOption)(short val);
	STDMETHOD(get_ClientAreaLeft)(long *retval);
	STDMETHOD(get_ClientAreaTop)(long *retval);
	STDMETHOD(get_ClientAreaWidth)(long *retval);
	STDMETHOD(get_ClientAreaHeight)(long *retval);
	STDMETHOD(get_SDIChildWindowClass)(BSTR *retval);
	STDMETHOD(put_SDIChildWindowClass)(BSTR val);
	STDMETHOD(put_ClientAreaHWnd)(OLE_HANDLE val);
	STDMETHOD(get_UseUnicode)(VARIANT_BOOL *retval);
	STDMETHOD(put_UseUnicode)(VARIANT_BOOL val);
	STDMETHOD(get_XPLook)(VARIANT_BOOL *retval);
	STDMETHOD(put_XPLook)(VARIANT_BOOL val);
	// DEFS for IActiveBar2
	
	

	// IPerPropertyBrowsing members

	STDMETHOD(GetDisplayString)( DISPID dispID, BSTR __RPC_FAR *pBstr);
	STDMETHOD(MapPropertyToPage)( DISPID dispID, CLSID __RPC_FAR *pClsid);
	STDMETHOD(GetPredefinedStrings)(  DISPID dispID,  CALPOLESTR __RPC_FAR *pCaStringsOut,  CADWORD __RPC_FAR *pCaCookiesOut);
	STDMETHOD(GetPredefinedValue)( DISPID dispID, DWORD dwCookie, VARIANT __RPC_FAR *pVarOut);
	// DEFS for IPerPropertyBrowsing
	
	

	// IDDPerPropertyBrowsing members

	STDMETHOD(GetType)( DISPID dispID,  long *pnType);
	// DEFS for IDDPerPropertyBrowsing
	
	
	// Events 
	void FireToolClick( Tool *tool);
	void FireReset( BSTR BandName);
	void FireNewToolbar( ReturnString *Name);
	void FireComboDrop( Tool *tool);
	void FireComboSelChange( Tool *tool);
	void FireTextChange( Tool *tool);
	void FireBandOpen( Band *band,  ReturnBool *Cancel);
	void FireBandClose( Band *band);
	void FireToolDblClick( Tool *tool);
	void FireBandMove( Band *band);
	void FireBandResize( Band *band, VARIANT *Widths, VARIANT *Heights,  long BandWidth,  long BandHeight);
	void FireError( short Number, ReturnString *Description, long Scode, BSTR Source, BSTR HelpFile, long HelpContext, ReturnBool *CancelDisplay);
	void FireDataReady();
	void FireMouseEnter( Tool *tool);
	void FireMouseExit( Tool *tool);
	void FireMenuItemEnter( Tool *tool);
	void FireMenuItemExit( Tool *tool);
	void FireKeyDown( short keycode, short shift);
	void FireKeyUp( short keycode, short shift);
	void FireChildBandChange( Band *band);
	void FireCustomizeBegin();
	void FireFileDrop( Band *band,  BSTR FileName);
	void FireCustomizeEnd( VARIANT_BOOL bModified);
	void FireMouseDown( short Button,  short Shift,  float x,  float y);
	void FireMouseMove( short Button,  short Shift,  float x,  float y);
	void FireMouseUp( short Button,  short Shift,  float x,  float y);
	void FireResize( long Left,  long Top,  long Width,  long Height);
	void FireBandDock( Band *band);
	void FireBandUndock( Band *band);
	void FireToolReset(Tool *tool);
	void FireToolGotFocus( Tool *tool);
	void FireToolLostFocus( Tool *tool);
	void FireWhatsThisHelp( Band *band, Tool *tool,  long HelpId);
	void FireCustomizeHelp( short ControlId);
	void Fire_ToolKeyDown( short keycode, short shift);
	void Fire_ToolKeyUp( Tool*tool,  short keycode, short shift);
	void FireClick();
	void FireDblClick();
	void FireCustomizeToolClick( Tool *tool);
	void Fire_ToolKeyPress( Tool*tool,  long keyascii );
	void Fire_ToolComboClose( Tool*tool);
	void FireToolKeyDown( Tool*tool,  short*keycode, short shift);
	void FireToolKeyUp( Tool*tool,  short*keycode, short shift);
	void FireToolKeyPress( Tool*tool,  long*keyascii );
	void FireToolComboClose( Tool*tool);
	void FireQueryUnload( short*Cancel);
	void FireToolHelp( Band *band, Tool *tool,  long HelpId);
	//{END INTERFACEDEFS}
	class CPropertyNotifySink : public IPropertyNotifySink
	{
		CBar *m_pMainUnknown();
		STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut);
		STDMETHOD_(ULONG, AddRef)(void) {return 1;};
		STDMETHOD_(ULONG, Release)(void) {return 1;};
		STDMETHOD(OnChanged)(DISPID dispID) {return m_pMainUnknown()->OnChanged(dispID);};
		STDMETHOD(OnRequestEdit)(DISPID dispID) {return m_pMainUnknown()->OnChanged(dispID);};
	} m_xPropNotify;
	
	STDMETHOD(OnChanged)(DISPID dispID);
	STDMETHOD(OnRequestEdit)(DISPID dispID);
	
	friend class CPropertyNotifySink;
	
//{AGGREGATION SUPPORT}
	virtual HRESULT InternalQueryInterface(REFIID riid, void **ppvObjOut);
	inline IUnknown *PrivateUnknown (void) {
	    return &m_UnkPrivate;
	}
protected:
	HRESULT ExternalQueryInterface(REFIID riid, void **ppvObjOut) {
		return m_pUnkOuter->QueryInterface(riid, ppvObjOut);
	}
	ULONG ExternalAddRef(void) {
		return m_pUnkOuter->AddRef();
	}
	ULONG ExternalRelease(void) {
		return m_pUnkOuter->Release();
	}
	IUnknown* m_pUnkOuter;
private:
	class CPrivateUnknownObject : public IUnknown {
	public:
		STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut);
		STDMETHOD_(ULONG, AddRef)(void);
		STDMETHOD_(ULONG, Release)(void);
		CPrivateUnknownObject() : m_cRef(1) {}
	private:
		CBar *m_pMainUnknown();
		ULONG m_cRef;
	} m_UnkPrivate;

	friend class CPrivateUnknownObject;
	//{AGGREGATION SUPPORT}

public:
#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif
	void OnNewDragDropStarted(void* pData);
	void DataReady(AsyncDownloadItem* pItem, BOOL bSuccessfull);
	void SetModified();
	HWND GetParentWindow();
	void ActiveBand(CBand* pActiveBand);
	CBand* ActiveBand();

	BOOL RemoveWindowFromWindowList(HWND hWndChild);

	BOOL AmbientUserMode();
	HWND GetDockWindow();
	void CreateDockWindows();
	BOOL IsForeground(HWND hWnd = NULL, BOOL* pbEnabled = NULL);
	STDMETHODIMP InternalRecalcLayout(BOOL bJustRecalc = FALSE);
	void NotifyFloatingWindows(DWORD dwFlags);

	void StartCustomization();

	friend class CCustomizeListbox;
	friend class CCustomize;
	friend class CBand;

	enum 
	{
		eStartDownload = 5700,
		eDownloadError = 5701,
		eCustomizationTimer = 432,
		eCustomizationTimerDuration = 200,
		eFlyByTimer = 555,
		eFlyByTimerDuration = 100,
		eDefaultBarIconWidth = 16,
		eDefaultBarIconHeight = 16,
		eLayoutVersion = 1,
		eConfigLayoutVersion = 10001,
		eConfigLayoutVersionOld = 10000,
		vbLeftButton = 1,
		vbRightButton = 2,
		vbMiddleButton = 4,
		eUsage = 0
	};

	enum SpecialToolIds
	{
		eSpecialToolId = 0x80000000,

		eToolIdCustomize =  0x80000000,
		eToolIdResetToolbar = 0x80000001,
		eToolIdBandToggle = 0x80000002,
		eToolIdToolToggle = 0x80000003,
		
		eToolIdSysCommand = 0x80001000,
		eToolIdMDIButtons = 0x80001001,

		eToolIdMininize = 0x80002003,
		eToolIdRestore =  0x80002000,
		eToolIdClose =    0x80002006,

		eToolIdStatusbar = 0x80003000,
		
		eToolIdWindow =     0x80004000,
		eToolIdCascade =    0x80004001,
		eToolIdTileHorz =   0x80004002,
		eToolIdTileVert =   0x80004003,
		eToolIdSeparator =  0x80004004,
		eToolIdWindowList = 0x80004005,
		
		eEditToolId = 0x80005000,

		eToolMenuExpand = 0x80006000,
		eToolMoreTools =  0x80006001
	};

	enum
	{
		eMDIForm,
		eSDIForm,
		eClientArea
	} m_eAppType;

	enum MDIButtonPosition
	{
		eMinimizeWindowPosition = 0,
		eRestoreWindowPosition = 1,
		eCloseWindowPosition = 2
	};

	enum MDIButton
	{
		eCloseWindow = 1,
		eRestoreWindow = 2,
		eMinimizeWindow = 4,
		eSysMenu = 8
	};

	IDragDropManager* m_pDragDropManager;
	IDesignerNotify*  m_pDesigner;
	CShortCuts*       m_pShortCuts;
	CImageMgr*        m_pImageMgrBandSpecific;
	CBand*			  m_pEditToolPopup;

	CPopupWinToolAndShadow* m_pToolShadow;

	BOOL m_bFirstPopup;

	//
	// Properties
	//

	BSTR m_bstrHelpFile;

	CPictureHolder m_phPict;
	CFontHolder    m_fhFont;
	CFontHolder    m_fhControl;
	CFontHolder    m_fhChildBand;

	CImageMgr* m_pImageMgr;
	CDockMgr*  m_pDockMgr;
	CBands*    m_pBands;
	CTools*    m_pTools;

	CDesignTimeSystemSettingNotify* m_pDesignerSettingNotify;

#pragma pack(push)
#pragma pack (1)
	struct BarPropV1
	{
		BarPropV1();

		VARIANT_BOOL m_vbDisplayKeysInToolTip:1;
		VARIANT_BOOL m_vbWhatsThisHelpMode:1;
		VARIANT_BOOL m_vbDisplayToolTips:1;
		VARIANT_BOOL m_vbAlignToForm:1;
		VARIANT_BOOL m_vbLargeIcons:1;
		VARIANT_BOOL m_vbDuplicateToolIds:1;

		AutoSizeChildrenTypes m_asChildren:4;
		MenuFontStyles m_mfsMenuFont:4;
		CustomizeTypes m_ctCustomize:4;
		MenuStyles	   m_msMenuStyle:4;

		OLE_COLOR m_ocBackground;
		OLE_COLOR m_ocForeground;
		OLE_COLOR m_ocHighLight;
		OLE_COLOR m_ocShadow;

		VARIANT_BOOL m_vbAutoUpdateStatusbar:1;
		PersonalizedMenuTypes m_pmMenus:3;
		VARIANT_BOOL m_vbUserDefinedCustomization:1;
		VARIANT_BOOL m_vbFireDblClickEvent:1;

		OLE_COLOR m_oc3DLight;
		OLE_COLOR m_oc3DDarkShadow;
		DWORD m_dwBackgroundOptions:4;
		DWORD m_dwToolIdentity;
		VARIANT_BOOL m_vbUseUnicode:1;
		VARIANT_BOOL m_vbXPLook:1;
	} bpV1;

	VARIANT_BOOL  m_vbCustomizeModified:1;

	BOOL m_bBandSpecificConvertIds:1;
	BOOL m_bWindowPosChangeMessage:1;
	BOOL m_bInitialRecalcLayout:1;
	BOOL m_bPopupMenuExpanded:1; 
	BOOL m_bApplicationActive:1;
	BOOL m_bFirstRecalcLayout:1;
	BOOL m_bNeedRecalcLayout:1;
	BOOL m_bAltCustomization:1;
	BOOL m_bStartedToolBars:1;
	BOOL m_bBandDestroyLock:1;
	BOOL m_bIsVisualFoxPro:1;
	BOOL m_bSaveSucceeded:1;
	BOOL m_bCustomization:1;
	BOOL m_bWhatsThisHelp:1;
	BOOL m_bSaveImages:1;
	BOOL m_bHasTexture:1;
	BOOL m_bToolModal:1;
	BOOL m_bClickLoop:1;
	BOOL m_bMenuLoop:1;
	BOOL m_bStatusBandLock:1;
	BOOL m_bToolCreateLock:1;
	BOOL m_bFreezeEvents:1;
	BOOL m_bStopStatusBarThread:1;
	BOOL m_bIE:1;
	BOOL m_bShutdownLock:1;
	VARIANT_BOOL m_vbLibrary:1;
	CBand* m_pFreezeBand;
#pragma pack(pop)

	ToolStack m_theToolStack;

	TypedMap<WORD, CTool*> m_mapMenuBar;
	TypedArray<HWND> m_aChildWindows;
	typedef void (*PFNMODIFYTOOL) (CBand* pBand, CTool* pTool, void* pData);

	long m_nFlags;

	//
	// Texture brush
	//
	
	HBRUSH m_hBrushTexture; 
	int    m_nTextureWidth;
	int    m_nTextureHeight;

	//
	// Colors
	//

	OLE_COLOR m_ocAmbientBackColor;
	COLORREF m_crAmbientBackColor;
	COLORREF m_crBackground;
	COLORREF m_crForeground;
	COLORREF m_crHighLight;
	COLORREF m_crShadow;
	COLORREF m_cr3DLight;
	COLORREF m_cr3DDarkShadow;
	COLORREF m_crMDIMenuBackground;

	COLORREF m_crXPPressed;
	COLORREF m_crXPBackground;
	COLORREF m_crXPMRUBackground;
	COLORREF m_crXPCheckedColor;
	COLORREF m_crXPSelectedColor;
	COLORREF m_crXPBandBackground;
	COLORREF m_crXPMenuBackground;
	COLORREF m_crXPMenuBorderShadow;
	COLORREF m_crXPSelectedCrossColor;
	COLORREF m_crXPSelectedBorderColor;
	COLORREF m_crXPMenuSelectedBorderColor;
	COLORREF m_crXPFloatCaptionBackground;
	COLORREF m_crXPShadow;
	COLORREF m_crXP3DLight;
	COLORREF m_crXPHighLight;
	COLORREF m_crXP3DDarkShadow;

	HBRUSH m_hBrushForeColor;
	HBRUSH m_hBrushBackColor;
	HBRUSH m_hBrushHighLightColor;
	HBRUSH m_hBrushShadowColor;

	struct DropInfo
	{
		DropInfo()
			: pBand(NULL),
			  pTool(NULL)
			 
		{
		}
		POINT  ptTarget;
		// Band Info
		CBand* pBand;
		CRect  rcBand; // Relative to Dock Window or Floating Window
		BOOL   bFloating;
		BOOL   bPopup;
		// Tool Info
		CTool* pTool;
		int    nToolIndex;
		int    nDropIndex;
		int    nDropDir;
	};
	DropInfo m_diDropInfo;
	DropInfo m_diCustSelection;

	struct MenuCustomization
	{
		MenuCustomization()
			: pBand(NULL),
			  pTool(NULL)
		{
		}

		CBand* pBand;
		CTool* pTool;
	} m_mcInfo;


	CBand* m_pPopupRoot;
	CBand* m_pMoreTools;
	CBand* m_pAllTools;

	DWORD m_dwMdiButtons; 

	int MessageBox(LPCTSTR szMsg, UINT nType);

	DWORD& LastDocked();
	DWORD& IncrementLastDocked();

	void CheckButton(ULONG nToolId, BOOL bChecked);
	void EnableButton(ULONG nToolId, BOOL bEnable);
	void SetVisibleButton(ULONG nToolId, BOOL bVisible);
	void SetToolShortCut(ULONG nToolId, CToolShortCut* pToolShortCutIn);
	void SetToolComboList(ULONG nToolId, IComboList* pToolComboListIn);
	void RemoveShortCutFromAllTools(ULONG nToolId, ShortCutStore* ShortCutStoreIn);

	void UpdateChildButtonState();
	void CacheSmButtonSize();
	void ClearCustomization();
	BOOL OnBarContextMenu();

	enum ToolStateOps 
	{
		TSO_ENABLE,
		TSO_VISIBLE,
		TSO_CHECK,
	};

	BOOL EnterWhatsThisHelpMode(HWND hWnd);
	void OnMDIChildPopupMenu(POINT ptScreen, CRect& rcScreen);
	void ChangeToolText(ULONG nToolId, BSTR newText);
	void GetCanonicalTool(ULONG nToolId, CTool*& pTool);
	void ChangeToolState(ULONG nToolId, ToolStateOps op, VARIANT_BOOL vbNewVal);
	void SetCustomInterface(ULONG nToolId, LPDISPATCH val, DWORD dwFlag);

	BOOL TabForward(CBand*& pBand, HWND& hWndBand, CTool*& pTool, int& nTool);
	BOOL TabBackward(CBand*& pBand, HWND& hWndBand, CTool*& pTool, int& nTool);
	BOOL UpdateTools(CBand* pBand, CTool* pTool, int& nTool, HWND hWndBand);

	HFONT GetFont();
	HFONT GetMenuFont(BOOL bVertical, BOOL bForce);
	HFONT GetControlFont();
	
	HFONT GetSmallFont(BOOL bVert);
	int GetSmallFontHeight(BOOL bVert);

	HFONT GetChildBandFont(BOOL bVert);
	int GetChildBandFontHeight(BOOL bVert);

	CBand* FindSubBand(BSTR bstrName);
	CBand* FindBand(BSTR bstrName);
	CBand* GetMenuBand();
	void GetPopupBands(CALPOLESTR& caPopups);
	void EnterMenuLoop(HWND hWndParent, UINT nKey);

	void StatusBandEnter(CTool* pToolIn);
	void StatusBandUpdate(CTool* pToolIn);
	void StatusBandExit();

	HBITMAP GetTextureBitmap();
	BOOL HasTexture();
	void FillTexture(HDC hDC, const CRect& rc);

	BOOL NeedsCycleShutDown(CTool* pTool, TypedArray<CTool*>* pfaToolsToBeDropped);
	
	BOOL DrawEdge(HDC hDC, const CRect& rcEdge, UINT nEdge, UINT nFlags);
	
	void BuildAccelators(CBand* pBand, TypedArray<WORD>* pfaChar = NULL);
	BOOL DoMenuAccelator(WORD nKey);

	void ModifyTool(ULONG nToolId, void* pData, PFNMODIFYTOOL theToolFunction);

	void ResetMenuFonts();

	HPALETTE& Palette();

	CLocalizer* Localizer();
	AsyncDownloadItem m_downloadItem;

	void ToolNotification(CTool* pTool, TOOLNF tnf);

	CBands* GetBands();
	CBand* StatusBand();
	void StatusBand(CBand* pStatusBand);

	void ShowToolTip(DWORD dwCookie, BSTR bstrText, const CRect& rcBound, BOOL bTopMost = FALSE);
	void HideToolTips(BOOL bResetTimer);
	void DestroyDetachBandOf(CBand* pBand);
	void ResetCycleMarks();
	void DoCustomization(BOOL bStart = TRUE, BOOL bRecalcLayout = TRUE);
	void SetActiveTool(CTool* pNewActiveTool);

	void DoToolModal(BOOL bVal);
	CRect& MDIClientRect();
	CRect& FrameRect();
	WNDPROC FrameProc();
	WNDPROC MDIClientProc();

	SIZE m_sizeSMButton;

	// PopupExit action detach point.
	POINT m_ptPea; 

	WNDPROC m_pMDIClientProc;
	WNDPROC m_pMainProc;

	ActiveBarErrorObject m_theErrorObject;
	CFlickerFree m_ffObj; 

	IDesignerNotifications* m_pDesignerNotify;
	CToolTip*	  m_pTip;
	CTool*		  m_pCustomizeDragLock;
	CTool*		  m_pActiveTool; 
	CTool*        m_pTabbedTool;

	CRect m_rcMDIClient;
	CRect m_rcFrame;
	CRect m_rcInsideFrame;

	DWORD m_nDragStartTick;

	HWND m_hWndModifySelection;
	HWND m_hWndMDIClient;
	HWND m_hWndMenuChild;
	HWND m_hWndDock;
	HWND m_hWndActiveCombo;

	HANDLE m_hStatusBarEvent;

	int m_nNewCustomizedFloatBandCounter;
	int m_nPopupExitAction; 

	POINT m_ptDragStart;
	POINT m_ptNewCustomizedFloatBandPos;

	void HideMiniWins();
	void DestroyMiniWins();

	void DestroyPopups();
	void ShowPopups();
	void HidePopups();
	void DrawDropSign(DropInfo& pInfo);

	FormsAndControls m_theFormsAndControls;
	CColorQuantizer* m_pColorQuantizer;
	CChildMenu m_theChildMenus;

	BOOL AddWindowList(CTool* pWindowListTool, BandTypes btType, TypedArray<CTool*>& aTools);
	CBand*		  m_pStatusBand;

	VARIANT m_vAlign;

	ActiveEdit* GetActiveEdit();

	ActiveCombobox* GetActiveCombo(DWORD dwStyle);

	ActiveCombobox* GetActiveCombo();

protected:
	void OnTimer(UINT nId);
	void OnFontHeightChanged();

	BOOL PositionAlignableControls(SIZEPARENTPARAMS& sppLayout);

	HRESULT PersistState(IStream* pStream, VARIANT_BOOL vbSave, BSTR bstrBandName = NULL);
	HRESULT PersistConfig(IStream* pStream, VARIANT_BOOL vbSave);
	HRESULT SaveData(IStream* pStream, BOOL bConfiguration, BSTR bstrName);
	HRESULT LoadData(IStream* pStream, BOOL bConfiguration, BSTR bstrName);

	HWND m_hWndDesignerShutDownAdvise;

	//
	// Color
	//
	
	void CacheColors();
	void CacheTexture();

	COLORREF m_crAmbientBackground;

	// Async downloading
	BSTR m_bstrDataPath;
	
	BSTR m_bstrWindowClass;
	//
	// Status Bar Stuff
	//
	
	IDesigner* m_pDesignerImpl;
	VARIANT_BOOL* m_pvbStatusBandToolsVisible;
	CCustomize*   m_pCustomize;
	CLocalizer*   m_pLocalizer;
	CTool*		  m_pStatusBandTool;
	CBand*		  m_pActiveBand;

	CRelocate* m_pRelocate;

	ActiveEdit* m_pEdit;
	ActiveCombobox* m_pCombo;

	HACCEL m_hAccel;

	HFONT m_hFontMenuVert;
	HFONT m_hFontMenuHorz;
	HFONT m_hFontToolTip;

	// used for etc
	HFONT m_hFontSmall;
	HFONT m_hFontSmallVert; 
	int m_nSmallFontHeight;
	int m_nSmallFontHeightVert;
	HFONT m_cached_fhFont;
	HFONT m_cachedControlFont;

	HFONT  m_hFontChildBand;
	HFONT  m_hFontChildBandVert; 
	int m_nChildBandFontHeight;
	int m_nChildBandFontHeightVert;
	int m_nFontHeight;
	int m_nControlFontHeight;

	DDString m_strMsgTitle;
	DWORD m_dwLastDocked;
	friend class CTool;
	static ErrorTable m_etErrors[];
	friend LRESULT CALLBACK FrameWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);
	
	HANDLE m_hStatusbarThread;
	UINT m_dwThreadId;

	void PaintVBBackground(HWND m_hWnd,HDC hdc,BOOL bIsMDI);
	int m_bUserMode; // -1 not cached yet, 0 designtime 1 runtime
					 // set to -1 in constructor and setclientsite

	BOOL StartStatusBandThread();
	BOOL StopStatusBandThread();
};

inline BOOL CBar::AmbientUserMode()
{
	if (-1 == m_bUserMode)
	{
		VARIANT_BOOL vbUserMode;
		if (GetAmbientProperty(DISPID_AMBIENT_USERMODE, VT_BOOL, &vbUserMode))
			m_bUserMode = VARIANT_TRUE == vbUserMode ? 1 : 0;
		else
			m_bUserMode = 0;
	}
	return m_bUserMode;
}

inline DWORD& CBar::LastDocked()
{
	return m_dwLastDocked;
}

inline DWORD& CBar::IncrementLastDocked()
{
	return ++m_dwLastDocked;
}

//
// GetParentWindow()
//

inline HWND CBar::GetParentWindow()
{
	switch (m_eAppType)
	{
	case eMDIForm:
		return (AmbientUserMode() ? GetDockWindow() : m_hWnd);
	default:
		return m_hWnd;
	}
}

//
// SetModified
//

inline void CBar::SetModified()
{
	m_fDirty = TRUE;
}

//
// GetSmallFontHeight
//

inline int CBar::GetSmallFontHeight(BOOL bVert)
{
	if (bVert)
	{
		if (NULL == m_hFontSmallVert)
			GetSmallFont(bVert);
		return m_nSmallFontHeightVert;
	}
	if (NULL == m_hFontSmall)
		GetSmallFont(bVert);
	return m_nSmallFontHeight;
}

//
// GetPageFontHeight
//

inline int CBar::GetChildBandFontHeight(BOOL bVert)
{
	if (bVert)
	{
		if (NULL == m_hFontChildBandVert)
			GetChildBandFont(bVert);
		return m_nChildBandFontHeightVert;
	}
	if (NULL == m_hFontChildBandVert)
		GetChildBandFont(bVert);
	return m_nChildBandFontHeight;
}

//
// GetFont
//

inline HFONT CBar::GetFont()
{
	if (!m_cached_fhFont)
		m_cached_fhFont=m_fhFont.GetFontHandle();

	return m_cached_fhFont;
}

//
// HasTexture
//

inline BOOL CBar::HasTexture() 
{
	return m_bHasTexture;
}

//
// MessageBox
//

inline int CBar::MessageBox(LPCTSTR szMsg, UINT nType = MB_OK)
{
	IsWindow(m_hWnd);
	return ::MessageBox(m_hWnd, szMsg, m_strMsgTitle, nType);
}

//
// GetBands
//

inline CBands* CBar::GetBands() 
{
	return m_pBands;
}

//
// DoToolModal
//

inline void CBar::DoToolModal(BOOL bVal) 
{
	m_bToolModal = bVal;
}

//
// MDIClientRect
//

inline CRect& CBar::MDIClientRect()
{
	return m_rcMDIClient;
}

//
// FrameRect
//

inline CRect& CBar::FrameRect()
{
	return m_rcFrame;
}

//
// FrameProc
//

inline WNDPROC CBar::FrameProc()
{
	return m_pMainProc;
}

//
// MDIClientProc
//

inline WNDPROC CBar::MDIClientProc()
{
	return m_pMDIClientProc;
}

//
// OnNewDragDropStarted
//

inline void CBar::OnNewDragDropStarted(void* pData) 
{
	NeedsCycleShutDown(NULL, (TypedArray<CTool*>*)pData);
}

//
// StatusBand
//

inline CBand* CBar::StatusBand()
{
	return m_pStatusBand;
}

//
// StatusBand
//

inline void CBar::StatusBand(CBand* pStatusBand)
{
	m_pStatusBand = pStatusBand;
	if (NULL == m_hStatusbarThread && m_pStatusBand && AmbientUserMode())
		StartStatusBandThread();
	else if (m_hStatusbarThread && NULL == m_pStatusBand)
		StopStatusBandThread();
}

//
// ActiveBand
//

inline void CBar::ActiveBand(CBand* pActiveBand)
{
	m_pActiveBand = pActiveBand;
}

//
// ActiveBand
//

inline CBand* CBar::ActiveBand()
{
	return m_pActiveBand;
}

inline ActiveCombobox* CBar::GetActiveCombo()
{
	return m_pCombo;
}

//
// Localizer
//

inline CLocalizer* CBar::Localizer()
{
	return m_pLocalizer;
}

//
// DownloadCallback
//

void DownloadCallback(void *cookie,AsyncDownloadItem *item,DWORD reason);

extern ITypeInfo* GetObjectTypeInfoEx(LCID lcid,REFIID iid);
extern ITypeInfo* GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif

