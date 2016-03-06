#ifndef __CCustomizeListbox_H__
#define __CCustomizeListbox_H__

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

struct CategoryMgr;
class CFlickerFree;

class CustomizeListboxErrorObject : public ErrorObject
{
	virtual void AsyncError(short		   nNumber, 
						    IReturnString* pDescription, 
						    long		   nSCode, 
						    BSTR		   bstrSource, 
						    BSTR		   bstrHelpFile, 
						    long		   nHelpContext, 
						    IReturnBool*   pCancelDisplay);
};

class CCustomizeListbox : public ICustomizeListbox,public IOleObject,public IOleInPlaceObject,public IOleInPlaceActiveObject,public IViewObject2,public IPersistPropertyBag,public IPersistStreamInit,public IPersistStorage,public IConnectionPointContainer,public ISpecifyPropertyPages,public IProvideClassInfo,public IOleControl,public ISupportErrorInfo,public IPerPropertyBrowsing,public ICategorizeProperties
{
public:
	CCustomizeListbox(IUnknown *pUnkOuter);
	int m_objectIndex;
	~CCustomizeListbox();

	void Recreate(BOOL bRecreate);

	HWND InternalCreateWindow(HWND hWnd, RECT& rc);

	VARIANT_BOOL AmbientUserMode();
	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CCustomizeListbox *CreateInstance(IUnknown *pUnkOuter);
	#define CCustomizeListbox_SUBCLASS _T("LISTBOX")

	// IUnknown members
	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
	STDMETHOD_(ULONG, AddRef)(void);
	STDMETHOD_(ULONG, Release)(void);
//	DECLARE_STANDARD_UNKNOWN();

	// IDispatch members

	STDMETHOD(GetTypeInfoCount)( UINT *pctinfo);
	STDMETHOD(GetTypeInfo)( UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
	STDMETHOD(GetIDsOfNames)( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid);
	STDMETHOD(Invoke)( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr);
	// DEFS for IDispatch
	
	

	// ICustomizeListbox members

	STDMETHOD(get_Type)(CustomizeListboxTypes *retval);
	STDMETHOD(put_Type)(CustomizeListboxTypes val);
	STDMETHOD(put_ActiveBar)(LPDISPATCH val);
	STDMETHOD(Add)(BSTR strName, BandTypes btType);
	STDMETHOD(Delete)();
	STDMETHOD(Rename)(BSTR strNew);
	STDMETHOD(Reset)();
	STDMETHOD(get_Category)(BSTR *retval);
	STDMETHOD(put_Category)(BSTR val);
	STDMETHOD(get_Font)(LPFONTDISP *retval);
	STDMETHOD(put_Font)(LPFONTDISP val);
	STDMETHOD(putref_Font)(LPFONTDISP *val);
	STDMETHOD(get_ToolDragDrop)(VARIANT_BOOL *retval);
	STDMETHOD(put_ToolDragDrop)(VARIANT_BOOL val);
	// DEFS for ICustomizeListbox
	
	

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
	BOOL OnSetExtent(const SIZEL *);
	HRESULT InPlaceActivate(LONG lVerb);
	HWND CreateInPlaceWindow(int  x,int  y,BOOL fNoRedraw);
	virtual BOOL RegisterClassData(void);
	static LRESULT CALLBACK ReflectWindowProc(HWND hwnd,UINT msg,WPARAM  wParam,LPARAM  lParam);
	static LRESULT CALLBACK ControlWindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
	static BOOL ControlFromHwnd(HWND hWnd, CCustomizeListbox*& pCustomizeListbox);
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
#ifdef CCustomizeListbox_SUBCLASS
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

        CCustomizeListbox *m_pControl;
        CConnectionPoint(CCustomizeListbox *pControl,BYTE type)
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
	
	

	// IPerPropertyBrowsing members

	STDMETHOD(GetDisplayString)( DISPID dispID, BSTR __RPC_FAR *pBstr);
	STDMETHOD(MapPropertyToPage)( DISPID dispID, CLSID __RPC_FAR *pClsid);
	STDMETHOD(GetPredefinedStrings)(  DISPID dispID,  CALPOLESTR __RPC_FAR *pCaStringsOut,  CADWORD __RPC_FAR *pCaCookiesOut);
	STDMETHOD(GetPredefinedValue)( DISPID dispID, DWORD dwCookie, VARIANT __RPC_FAR *pVarOut);
	// DEFS for IPerPropertyBrowsing
	
	

	// ICategorizeProperties members

	STDMETHOD(MapPropertyToCategory)( DISPID dispid,  PROPCAT *ppropcat);
	STDMETHOD(GetCategoryName)( PROPCAT propcat,  LCID lcid, BSTR *pbstrName);
	// DEFS for ICategorizeProperties
	
	
	// Events 
	void FireSelChangeBand(Band *pBand);
	void FireSelChangeCategory(BSTR strCategory);
	void FireSelChangeTool(Tool *pTool);
	void FireError( short Number, ReturnString *Description, long Scode, BSTR Source, BSTR HelpFile, long HelpContext, ReturnBool *CancelDisplay);
	//{END INTERFACEDEFS}
	//{AGGREGATION SUPPORT}
	virtual HRESULT InternalQueryInterface(REFIID riid, void **ppvObjOut);
	inline IUnknown *PrivateUnknown (void) {
	    return &m_UnkPrivate;
	}

	void DrawItem(LPDRAWITEMSTRUCT pDrawItemStruct);
	void Clear();

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
	class CPropertyNotifySink : public IPropertyNotifySink
	{
		CCustomizeListbox *m_pMainUnknown();
		STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut);
		STDMETHOD_(ULONG, AddRef)(void) {return 1;};
		STDMETHOD_(ULONG, Release)(void) {return 1;};
		STDMETHOD(OnChanged)(DISPID dispID) {return m_pMainUnknown()->OnChanged(dispID);};
		STDMETHOD(OnRequestEdit)(DISPID dispID) {return m_pMainUnknown()->OnChanged(dispID);};
	} m_xPropNotify;

	STDMETHOD(OnChanged)(DISPID dispID);
	STDMETHOD(OnRequestEdit)(DISPID dispID);
	
	friend class CPropertyNotifySink;

	class CPrivateUnknownObject : public IUnknown {
	public:
		STDMETHOD(QueryInterface)(REFIID riid, void **ppvObjOut);
		STDMETHOD_(ULONG, AddRef)(void);
		STDMETHOD_(ULONG, Release)(void);
		CPrivateUnknownObject() : m_cRef(1) {}
	private:
		CCustomizeListbox *m_pMainUnknown();
		ULONG m_cRef;
	} m_UnkPrivate;

	friend class CPrivateUnknownObject;
	//{AGGREGATION SUPPORT}

	enum
	{
		eCheckboxWidth = 13,
		eCheckboxHeight = 13

	};

	BOOL CreateCategoryManager();
	void GetNewBandName(BSTR& bstr);
	void FillBands();
	void RecreateControlWindow();
	void DrawCheckItem(LPDRAWITEMSTRUCT pDrawItemStruct);
	void DrawToolItem(LPDRAWITEMSTRUCT pDrawItemStruct);
	void SetItemData(int nIndex, void* pData); 
	DWORD GetItemData(int nIndex);
	void GetItemRect(int nIndex, RECT* prc);
	int AddString(LPCTSTR szString);
	void InvalidateItem(int nIndex);
	int GetCount();
	int GetCurSel(); 
	int GetCheck(int nIndex);
	void SetCheck(int nIndex, BOOL bCheck);
	void OnLButtonDown(UINT nFlags, POINT pt);
	void OnCommand(int nNotifyMessage);
	LRESULT EraseBackground(HDC hDC);
	void OnDeleteItem(LPDELETEITEMSTRUCT pDeleteStruct);

	CustomizeListboxErrorObject m_theErrorObject;
	VARIANT_BOOL		  m_vbToolDragDrop;
	CFlickerFree*		  m_pFF;
	CategoryMgr*		  m_pCategoryMgr;
	IActiveBar2*          m_pActiveBar;
	DWORD				  m_dwToolDropId;
	DWORD				  m_dwToolDropIdEx;
	BOOL				  m_bDontRecreate;
	BOOL				  m_bDropTarget;	
	BOOL				  m_bDirty;
	int					  m_nTextHeight;
	int					  m_nItemHeight;
	int					  m_nDropIndex;

	// Properties
	CustomizeListboxTypes m_clbType;
	CFontHolder			  m_Font;
	BSTR				  m_bstrCategory;

	COLORREF m_crButtonFace;
	COLORREF m_crXPMenuBackground;
	COLORREF m_crXPBandBackground;
	COLORREF m_crXPSelected;
	COLORREF m_crXPSelectedBorder;
};

inline void CCustomizeListbox::Recreate(BOOL bRecreate)
{
	m_bDontRecreate = bRecreate; 
};

inline int CCustomizeListbox::GetCount() 
{
	return (int)SendMessage(m_hWnd,LB_GETCOUNT,0,0);
};

inline int CCustomizeListbox::GetCurSel() 
{
	return (int)SendMessage(m_hWnd,LB_GETCURSEL,0,0);
};

inline void CCustomizeListbox::SetItemData(int nIndex, void* pData) 
{
	LRESULT lResult;
	lResult = SendMessage(m_hWnd, LB_SETITEMDATA, (WPARAM)nIndex, (LPARAM)pData);
};

inline DWORD CCustomizeListbox::GetItemData(int nIndex) 
{
	return (DWORD)SendMessage(m_hWnd, LB_GETITEMDATA, (WPARAM)nIndex, 0);
};

inline void CCustomizeListbox::GetItemRect(int index,RECT *r) 
{
	SendMessage(m_hWnd,LB_GETITEMRECT,(WPARAM)index,(LPARAM)(r));
};

inline int CCustomizeListbox::AddString(LPCTSTR pstr) 
{
	return (int)SendMessage(m_hWnd,LB_ADDSTRING,0,(LPARAM)(pstr));
};

inline void CCustomizeListbox::InvalidateItem(int nIndex)
{
	RECT rc;
	SendMessage(m_hWnd, LB_GETITEMRECT, nIndex, (LPARAM)&rc);
	::InvalidateRect(m_hWnd,&rc,FALSE);	
}

inline BOOL CCustomizeListbox::ControlFromHwnd(HWND hWnd, CCustomizeListbox*& pCustomizeListbox)
{
	return GetGlobals().m_pControls->Lookup((LPVOID)hWnd, (LPVOID&)pCustomizeListbox);
}

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

//
// CCustomizeListboxHost
//

class CCustomizeListboxHost : public IDispatch
{
public:
	STDMETHOD(QueryInterface)(REFIID riid, void** ppvObj);
	STDMETHOD_(ULONG, AddRef)(void);
	STDMETHOD_(ULONG, Release)(void);

	// IDispatch members
	STDMETHOD(GetTypeInfoCount)(UINT* pctinfo);
	STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo);
	STDMETHOD(GetIDsOfNames)(REFIID riid, OLECHAR** rgszNames, UINT cNames, ULONG lcid, long* rgdispid);
	STDMETHOD(Invoke)(long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr);

	// Events
	virtual void FireSelChangeBand(IBand* pBand) {};
	virtual void FireSelChangeCategory(BSTR strCategory) {};
	virtual void FireSelChangeTool(ITool* pTool) {};
};

//
// ToolBarName
//

class ToolBarName : public FDialog
{
public:
	enum Type
	{
		eNew,
		eRename
	};

	ToolBarName(CBar* pBar, Type eType);
	TCHAR* Name();
	void Name(TCHAR* szName);

private:
	void Customize(UINT nId, LocalizationTypes eType);
	virtual BOOL DialogProc(UINT nMsg,WPARAM wParam,LPARAM lParam);
	virtual void OnOK();
	TCHAR m_szBuffer[MAX_PATH];
	CBar* m_pBar;
	Type m_eType;
};

inline TCHAR* ToolBarName::Name()
{
	return m_szBuffer;
}

inline void ToolBarName::Name(TCHAR* szName)
{
	_tcsncpy(m_szBuffer, szName, MAX_PATH);
}

#endif
