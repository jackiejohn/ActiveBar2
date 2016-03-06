#ifndef __CBar_H__
#define __CBar_H__
#include "fontholder.h"

class CBar : public IOleObject,public IOleInPlaceObject,public IOleInPlaceActiveObject,public IViewObject2,public IPersistPropertyBag,public IPersistStreamInit,public IPersistStorage,public IConnectionPointContainer,public ISpecifyPropertyPages,public IProvideClassInfo,public IBarPrivate,public IOleControl,public IPerPropertyBrowsing,public ISupportErrorInfo,public IActiveBar2
{
public:
	CBar(IUnknown *pUnkOuter);
	~CBar();

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CBar *CreateInstance(IUnknown *pUnkOuter);
	// IUnknown members
	DECLARE_STANDARD_UNKNOWN();

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

        CBar *m_pOleControl();
        CConnectionPoint()
		{
            m_rgSinks = NULL;
            m_cSinks = 0;
			inFireEvent=0;
        }
        ~CConnectionPoint();

      private:
        short  m_cSinks;

    } m_cpEvents, m_cpPropNotify;
	
	void BoundPropertyChanged(DISPID dispid);

    friend CConnectionPoint;
	
	

	// ISpecifyPropertyPages members

	STDMETHOD(GetPages)( CAUUID *pPages);
	// DEFS for ISpecifyPropertyPages
	
	

	// IProvideClassInfo members

	STDMETHOD(GetClassInfo)( ITypeInfo **ppTI);
	// DEFS for IProvideClassInfo
	
	

	// IBarPrivate members

	STDMETHOD(DesignerShutDown)();
	STDMETHOD(CreateBand)( IBand **pBand);
	STDMETHOD(CreateToolsCollection)( ITools **pTools);
	STDMETHOD(ShutDownAdvise)( OLE_HANDLE handle);
	// DEFS for IBarPrivate
	
	

	// IOleControl members

	STDMETHOD(GetControlInfo)( CONTROLINFO *pCI);
	STDMETHOD(OnMnemonic)( MSG *pMsg);
	STDMETHOD(OnAmbientPropertyChange)( DISPID dispID);
	STDMETHOD(FreezeEvents)( BOOL bFreeze);
	// DEFS for IOleControl
	virtual void AmbientPropertyChanged(DISPID dispid);
	BOOL m_fModeFlagValid;	
	

	// IPerPropertyBrowsing members

	STDMETHOD(GetDisplayString)( DISPID dispID, BSTR __RPC_FAR *pBstr);
	STDMETHOD(MapPropertyToPage)( DISPID dispID, CLSID __RPC_FAR *pClsid);
	STDMETHOD(GetPredefinedStrings)(  DISPID dispID,  CALPOLESTR __RPC_FAR *pCaStringsOut,  CADWORD __RPC_FAR *pCaCookiesOut);
	STDMETHOD(GetPredefinedValue)( DISPID dispID, DWORD dwCookie, VARIANT __RPC_FAR *pVarOut);
	// DEFS for IPerPropertyBrowsing
	
	

	// ISupportErrorInfo members

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid);
	// DEFS for ISupportErrorInfo
	
	

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
	STDMETHOD(Customize)(CustomizeTypes nCustomize);
	STDMETHOD(ReleaseFocus)();
	STDMETHOD(Attach)();
	STDMETHOD(Detach)();
	STDMETHOD(OnKeyDown)( short KeyCode, short Shift, VARIANT_BOOL *retval);
	STDMETHOD(OnKeyUp)( short KeyCode, short Shift, VARIANT_BOOL *retval);
	STDMETHOD(Save)( BSTR FileName, BSTR BandName);
	STDMETHOD(Load)( BSTR FileName, BSTR BandName);
	STDMETHOD(GetLayoutData)( BSTR BandName, VARIANT *data);
	STDMETHOD(SetLayoutData)( BSTR bandName,  VARIANT *data);
	STDMETHOD(get_Font)(LPFONTDISP *retval);
	STDMETHOD(put_Font)(LPFONTDISP val);
	STDMETHOD(putref_Font)(LPFONTDISP *val);
	STDMETHOD(About)();
	STDMETHOD(get_ColorDepth)(BitDepth *retval);
	STDMETHOD(put_ColorDepth)(BitDepth val);
	STDMETHOD(get_DataPath)(BSTR *retval);
	STDMETHOD(put_DataPath)(BSTR val);
	STDMETHOD(AttachEx)( OLE_HANDLE hWnd);
	STDMETHOD(get_MenuFontStyle)(MenuFontStyles *retval);
	STDMETHOD(put_MenuFontStyle)(MenuFontStyles val);
	STDMETHOD(get_LocaleString)(BSTR *retval);
	STDMETHOD(put_LocaleString)(BSTR val);
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
	STDMETHOD(get_Palette)(OLE_HANDLE *retval);
	STDMETHOD(put_Palette)(OLE_HANDLE val);
	STDMETHOD(get_ControlFont)(LPFONTDISP *retval);
	STDMETHOD(put_ControlFont)(LPFONTDISP val);
	STDMETHOD(putref_ControlFont)(LPFONTDISP *val);
	STDMETHOD(Refresh)();
	STDMETHOD(OnSysColorChanged)();
	STDMETHOD(get_PageFont)(LPFONTDISP *retval);
	STDMETHOD(put_PageFont)(LPFONTDISP val);
	STDMETHOD(putref_PageFont)(LPFONTDISP *val);
	STDMETHOD(PlaySound)(BSTR strSound, SoundTypes stType);
	STDMETHOD(get_MenuAnimation)(MenuStyles *retval);
	STDMETHOD(put_MenuAnimation)(MenuStyles val);
	STDMETHOD(get_LargeIcons)(VARIANT_BOOL *retval);
	STDMETHOD(put_LargeIcons)(VARIANT_BOOL val);
	STDMETHOD(get_ImageManager)(VARIANT *retval);
	STDMETHOD(put_ImageManager)(VARIANT val);
	STDMETHOD(ResetHooks)();
	STDMETHOD(get_WhatsThisHelpMode)(VARIANT_BOOL *retval);
	STDMETHOD(put_WhatsThisHelpMode)(VARIANT_BOOL val);
	// DEFS for IActiveBar2
	
	CFontHolder m_Font;
	
	// Events 
	void FireClick( Tool *tool);
	void FireReset( BSTR BandName);
	void FireNewToolbar( ReturnString *Name);
	void FireComboDrop( Tool *tool);
	void FireComboSelChange( Tool *tool);
	void FireTextChange( Tool *tool);
	void FireBandOpen( Band *band);
	void FireBandClose( Band *band);
	void FireDblClick( Tool *tool);
	void FireGotFocus( Tool *tool);
	void FireLostFocus( Tool *tool);
	void FireBandMove( Band *band);
	void FireBandResize( Band *band);
	void FireError( short Number, ReturnString *Description, long Scode, BSTR Source, BSTR HelpFile, long HelpContext, ReturnBool *CancelDisplay);
	void FirePreSysMenu( Band *band);
	void FireDataReady();
	void FireMouseEnter( Tool *tool);
	void FireMouseExit( Tool *tool);
	void FireMenuItemEnter( Tool *tool);
	void FireMenuItemExit( Tool *tool);
	void FirePreCustomizeMenu( ReturnBool *Cancel);
	void FireKeyDown( short keycode, short shift);
	void FireKeyUp( short keycode, short shift);
	void FirePageChange( Page *page);
	void FireBeginCustomize();
	//{END INTERFACEDEFS}
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
	IUnknown *m_pUnkOuter;
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
	class CPropertyNotifySink : public IPropertyNotifySink
	{
		DCChartObj *m_pMainUnknown();
		STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut);
		STDMETHOD_(ULONG, AddRef)(void) {return 1;};
		STDMETHOD_(ULONG, Release)(void) {return 1;};
		STDMETHOD(OnChanged)(DISPID dispID) {return m_pMainUnknown()->OnChanged(dispID);};
		STDMETHOD(OnRequestEdit)(DISPID dispID) {return m_pMainUnknown()->OnChanged(dispID);};
	} m_xPropNotify;
	
	STDMETHOD(OnChanged)(DISPID dispID);
	STDMETHOD(OnRequestEdit)(DISPID dispID);
	
	friend class CPropertyNotifySink;
	
};

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif
