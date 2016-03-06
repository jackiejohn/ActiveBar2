#ifndef BROWSER_INCLUDED
#define BROWSER_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Trees.h"

struct IObjectChanged;
class CDesignerPage;
class CEditListBox;
class CDefineColor;
class CFlickerFree;
class CComboList;
class CShortCut;
class CDesigner;
class CBrowser;

//
// COleDispatchDriver
//

class COleDispatchDriver
{
	// Constructors
public:
	COleDispatchDriver(LPDISPATCH pDispatch);

	// Operations
	HRESULT __cdecl SetProperty(DISPID dwDispID, VARTYPE vtProp, ...);
	HRESULT GetProperty(DISPID dwDispID, VARTYPE vtProp, VARIANT* pvProp) const;

	// special operators
	operator LPDISPATCH();

private:
	HRESULT InvokeHelperV(DISPID      dwDispID, 
						  WORD        wFlags, 
						  VARTYPE     vtRet,
						  VARIANT*    pvRet, 
						  const BYTE* pbParamInfo, 
						  va_list     argList);

	// helpers for IDispatch::Invoke
	HRESULT __cdecl InvokeHelper(DISPID      dwDispID, 
								 WORD        wFlags,
								 VARTYPE     vtRet, 
								 VARIANT*    pvRet, 
								 const BYTE* pbParamInfo, 
								 ...);

	const COleDispatchDriver& operator=(const COleDispatchDriver& dispatchSrc);

	// Attributes
	LPDISPATCH m_pDispatch;
	DDString m_strError;
};

inline COleDispatchDriver::operator LPDISPATCH()
{
	return m_pDispatch;

}
//
// EnumItem
//

struct EnumItem
{
	DDString m_strDisplay;
	long    m_nValue;
	void Dump();
};

//
// CEnum
//

class CEnum
{
public:
	CEnum(BSTR bstrEnumName = NULL)
	{
		m_bstrEnumName = SysAllocString(bstrEnumName);
	}

	~CEnum();

	BSTR& EnumName();
	EnumItem* ByValue(long nValue, int& nIndex);
	EnumItem* ByIndex(int nIndex);
	
	int IndexOfValue(long nValue);

	void FillCombo(HWND hWnd);
	EnumItem* Item(int nIndex);
	int Add(EnumItem* pEnumItem);
	int NumOfItems();
	void Dump();
private:
	TypedArray<EnumItem*> m_aEnumItems;
	BSTR				  m_bstrEnumName;

	friend class CBrowserEdit;
};

//
// Item
//

inline EnumItem* CEnum::Item(int nIndex)
{
	return m_aEnumItems.GetAt(nIndex);
}

//
// EnumName
//

inline BSTR& CEnum::EnumName()
{
	return m_bstrEnumName;
}

inline int CEnum::NumOfItems()
{
	return m_aEnumItems.GetSize();
}

//
// CProperty
//

class CProperty
{
public:
	enum EditType
	{
		eColor,
		eFont,
		ePicture,
		eNumber,
		eString,
		eBool,
		eEnum,
		eFlags,
		eEdit,
		eShortCut,
		eComboList,
		eBitmap,
		eVariant
	};

	CProperty(LPDISPATCH pDispatch, MEMBERID nMemId, VARTYPE vt, EditType etType, short nParams, BSTR bstrDocString, DWORD dwContextId, CEnum* pEnum = NULL);
	~CProperty();

	BOOL IsDirty();

	HRESULT GetValue();
	HRESULT SetValue();
	HRESULT SetValueNoCheck();
	LPCTSTR Name();
	void Name(LPCTSTR szName);

	void SetEnum(CEnum* pEnum);
	CEnum* GetEnum();

	VARIANT& GetData();
	
	DWORD ContextId();

	void Get(const BOOL& bGet);
	void Put(const BOOL& bPut);
	void PutRef(const BOOL& bPutRef);
	BOOL Get();
	BOOL Put();
	BOOL PutRef();

	EditType& GetEditType();
	void SetEditType(EditType eType);

	void VarType(VARTYPE vt);
	VARTYPE VarType();
	
	void VarTypeProperty(VARTYPE vt);

	MEMBERID& MemId();

	HFONT GetFont();
	HRESULT SetFont(HFONT hFont);

	short GetPictureType();
	void Clear();

	BSTR DocString();
	LPDISPATCH Driver();

private:
	TCHAR*   m_szName;
	MEMBERID m_nMemId;
 	VARIANT  m_vProperty;
 	VARIANT  m_vPropertyInitial;
	VARTYPE  m_vt;
	EditType m_etType;
	short    m_nParams;
	DWORD    m_dwContextId;
	BSTR	 m_bstrDocString;
	BOOL     m_bGet;
	BOOL     m_bPut;
	BOOL     m_bPutRef;
	CEnum*   m_pEnum;
	COleDispatchDriver m_theDriver;
};

inline DWORD CProperty::ContextId()
{
	return m_dwContextId;
}

inline BSTR CProperty::DocString()
{
	return m_bstrDocString;
}

inline LPDISPATCH CProperty::Driver()
{
	return m_theDriver;
}

inline LPCTSTR CProperty::Name()
{
	return m_szName;
}

inline void CProperty::Name(LPCTSTR szName)
{
	if (NULL == m_szName)
		delete [] m_szName;
	m_szName = new TCHAR[_tcslen(szName) + 1];
	_tcscpy(m_szName, szName);
}

inline MEMBERID& CProperty::MemId()
{
	return m_nMemId;
}

inline VARIANT& CProperty::GetData()
{
	return m_vProperty;
}

inline VARTYPE CProperty::VarType()
{
	return m_vt;
}

inline void CProperty::VarTypeProperty(VARTYPE vt)
{
	m_vProperty.vt = vt;
}

inline void CProperty::Clear()
{
	VariantClear(&m_vProperty);
	VariantInit(&m_vProperty);
	m_vProperty.vt = m_vt;
}

inline CProperty::EditType& CProperty::GetEditType()
{
	return m_etType;
}

inline void CProperty::SetEditType(CProperty::EditType etType)
{
	m_etType = etType;
}

inline void CProperty::Get(const BOOL& bGet)
{
	m_bGet = bGet;
}

inline void CProperty::Put(const BOOL& bPut)
{
	m_bPut = bPut;
}

inline void CProperty::PutRef(const BOOL& bPutRef)
{
	m_bPutRef = bPutRef;
}

inline BOOL CProperty::Get()
{
	return m_bGet;
}

inline BOOL CProperty::Put()
{
	return m_bPut;
}

inline BOOL CProperty::PutRef()
{
	return m_bPutRef;
}

inline void CProperty::VarType(VARTYPE vt)
{
	m_vt = vt;
}

inline void CProperty::SetEnum(CEnum* pEnum)
{
	m_pEnum = pEnum;
}

inline CEnum* CProperty::GetEnum()
{
	return m_pEnum;
}

//
// ParentWndMonitorThread
//

class CPropertyPopup;

struct ParentWndMonitorThread
{
	ParentWndMonitorThread();
	~ParentWndMonitorThread();

	void SetWindow(CPropertyPopup* pPropertyPopup);

	BOOL Run();
	BOOL Pause();
	BOOL WaitUntilDead();
	void Reset();

private:
	static DWORD WINAPI Process(void* pData);

	CPropertyPopup* m_pPropertyPopup;
	HANDLE m_hThread;
	DWORD  m_dwWaitTimer;
	CRect  m_rcParent;
	BOOL   m_bPaused;
};

//
// CPropertyPopup
//

class CPropertyPopup : public FWnd
{
public:
	CPropertyPopup();
	virtual ~CPropertyPopup();

	HWND hWndParent();

	virtual BOOL Create(HWND hWndParent) = 0; 
	virtual void Show(CRect rcPosition,	CProperty* pProperty = NULL) = 0;
	virtual void Hide() = 0;

protected:
	virtual void Fill(void* pData = NULL) {};
	virtual void CleanUp() {};
	
	virtual void RegisterWindow();
	virtual BOOL PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam);

	static LRESULT CALLBACK KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);

	static ParentWndMonitorThread m_theMonitor;
	static CPropertyPopup*		  m_pCurrentPopup;
	static HHOOK				  m_hHook;

	CProperty* m_pProperty;
	HWND	   m_hWndParent;
};

inline HWND CPropertyPopup::hWndParent()
{
	return m_hWndParent;
}

inline BOOL CPropertyPopup::PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam)
{
	return FALSE;
}

//
// CEdit
//

class CEdit : public CPropertyPopup
{
public:
	CEdit();
	~CEdit();

	BOOL Create(HWND hWnd);
	void Show(CRect rcPosition,	CProperty* pProperty);
	void Hide();

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam);
	virtual void Fill(void* pData = NULL);
	void Clear();

	HWND m_hWndParent;
	HWND m_hWndEdit;
};

//
// Clear
//

inline void CEdit::Clear()
{
	::SendMessage(m_hWndEdit, WM_SETTEXT, 0, (LPARAM)_T(""));
}

//
// Hide
//

inline void CEdit::Hide()
{
	CPropertyPopup::Hide();
	ShowWindow(SW_HIDE);
}

//
// Keys
//

struct Keys
{
	Keys();
	~Keys();

	static Keys* GetKeys();

	static DDString m_strShift;
	static DDString m_strCtrl;
	static DDString m_strAlt;
};

//
// CShortCutEdit
//

class CShortCutEdit : public FWnd
{
public:
	CShortCutEdit(CEditListBox& theEditListBox);
	~CShortCutEdit();
	
	BOOL Create(HWND hWndParent); 
	void CleanUp();
	void UpdateString();
	void StartEditing(TCHAR* szText, CRect& rcItem, int& nIndex, HFONT hFont);
	BOOL ProcessKeyboard(WPARAM wParam, LPARAM lParam);
	BOOL CheckForDuplicates(IShortCut* pShortCut);

private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);

	TypedArray<TCHAR*> m_faShortCut;
	CEditListBox&      m_theEditListBox;
	IShortCut*         m_pShortCutStore;
};

//
// CEditListBox
//

LRESULT CALLBACK ListBoxKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);

class CEditListBox : public FWnd
{
public:
	CEditListBox(CShortCut* pShortCut);
	~CEditListBox();

	enum 
	{
		eLastItemWidth = 15,
		eBufferSize = 256,
		eEditControl = 100
	};

	BOOL Create(HWND hWndParent); 

	int GetCount();

	void GetText (int iIndex, DDString& strText);	
	
	const BOOL& Sorted();

	BOOL& EditMode();
	void EditMode(BOOL bEditMode);

	int CurrentItem();
	void CurrentItem(int nCurrentItem);

	void OnDelete();

	void StopEditing();

	void CleanUp();

	void DrawItem(LPDRAWITEMSTRUCT pDrawItemStruct);
	
	void MeasureItem(LPMEASUREITEMSTRUCT pMeasureItemStruct);
	
	void DeleteItem(LPDELETEITEMSTRUCT pDeleteItemStruct);

	void StartEditing();

	void ProcessKeyboard(WPARAM wParam, LPARAM lParam);

	void SaveAndClose();
	void Close();

private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	BOOL PreTranslateMessage(MSG* pMsg);

	LRESULT OnResetContent();
	void DestroyItem (const DWORD& dwItemData);
	LRESULT OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	BOOL OnLButtonDown(UINT nFlags, LPARAM lParam);
	void OnLButtonDblClk(UINT nFlags, POINT pt);	
	void OnShowWindow(BOOL bShow, UINT nStatus);
	void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);

	void ReportError (LRESULT lResult);

	BOOL GetTextMetrics (HDC hDC, TEXTMETRIC& tm);

	//
	// This for searching for the correct postion in the list box
	//
	
	int FindIndex (int nStart, int nEnd, UINT nKey);

	CShortCutEdit* m_pShortCutEdit;
	CShortCut*     m_pShortCut;      
	BOOL		   m_bEmptyEditItem;
	BOOL		   m_bEditMode;
	BOOL		   m_bSorted;
	HWND		   m_hWndParent;
	int			   m_nCount;
	int			   m_nCurrentEditItem;
	friend LRESULT CALLBACK ListBoxKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);
};

inline void CEditListBox::CurrentItem(int nCurrentItem)
{
	m_nCurrentEditItem = nCurrentItem;
}

inline int CEditListBox::CurrentItem()
{
	return m_nCurrentEditItem;
}

inline void CEditListBox::ProcessKeyboard(WPARAM wParam, LPARAM lParam)
{
	m_pShortCutEdit->ProcessKeyboard(wParam, lParam);
}

inline void CEditListBox::CleanUp()
{
	m_pShortCutEdit->CleanUp();
}

inline const BOOL& CEditListBox::Sorted()
{
	return m_bSorted;
}

inline void CEditListBox::EditMode(BOOL bEditMode)
{
	m_bEditMode = bEditMode;
}

inline BOOL& CEditListBox::EditMode()
{
	return m_bEditMode;
}

//
// CShortCut
//

class CShortCut : public CPropertyPopup
{
public:
	CShortCut();
	~CShortCut();

	virtual BOOL Create(HWND hWndParent); 
	virtual void Show(CRect rcPosition,	CProperty* pProperty);
	virtual void Hide();
	void SaveAndClose();
	void Close();

private:
	virtual BOOL PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam);
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual void Fill(void* pData = NULL);

	BOOL SaveShortCuts();
	void Clear();

	CEditListBox* m_pEditListBox;  
};

//
// Clear
//

inline void CShortCut::Clear()
{
	if (m_pEditListBox)
	{
		m_pEditListBox->SendMessage(LB_RESETCONTENT);
		m_pEditListBox->CurrentItem(0);
	}
}

//
// CComboListEdit
//

class CComboListEdit : public FWnd
{
public:
	CComboListEdit();
	~CComboListEdit();
	
	BOOL Create(CComboList* pComboList, HFONT hFont);
private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	CComboList* m_pComboList;
};

//
// CComboList
//

class CComboList : public CPropertyPopup
{
public:
	CComboList();
	~CComboList();

	BOOL Create(HWND hWnd);
	void Show(CRect rcPosition,	CProperty* pProperty);
	void Hide();
	void Save();
private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam);
	virtual void Fill(void* pData = NULL);

	void Clear();
	CComboListEdit m_theEdit;
};

//
// Clear
//

inline void CComboList::Clear()
{
	m_theEdit.SendMessage(WM_SETTEXT, 0, (LPARAM)_T(""));
}

//
// Hide
//

inline void CComboList::Hide()
{
	CPropertyPopup::Hide();
	ShowWindow(SW_HIDE);
}

class CLItemData
{
public:
	CLItemData() {name=NULL;};
	~CLItemData();

	BSTR name;
	BOOL checked;
	LPARAM lParam;
};

//
// DDCheckListBox
//

class DDCheckListBox : public FWnd
{
public:
	enum
	{
		eCheckBoxHeight = 13,
		eCheckBoxWidth = 13
	};

	DDCheckListBox();
	virtual ~DDCheckListBox();

	BOOL Create(DWORD dwStyle, const CRect& rc, HWND hWndParent, UINT nId);

	void AddItem(BSTR bstrName, BOOL bChecked, LPARAM lParam = NULL);
	
	int AddString(LPCTSTR szString);

	int GetCount();
	
	void GetItemRect(int nIndex, CRect& rcItem);

	LPARAM GetlParam(int nIndex);

	int GetCheck(int nIndex);
	void SetCheck(int nIndex, BOOL bValue);
	
	void SetItemData(int nIndex, DWORD dwData);
	DWORD GetItemData(int nIndex);

	BOOL OnChildNotify(UINT nMsg, WPARAM wParam, LPARAM lParam, LRESULT* plResult);
	int GetCurSel();
	void SetCurSel(int nIndex);

	int   m_nItemHeight;
	HFONT m_hFont;

private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual void MeasureItem(LPMEASUREITEMSTRUCT pMeasureItem);
	virtual void DrawItem(LPDRAWITEMSTRUCT pDrawItem);
	virtual int CompareItem(LPCOMPAREITEMSTRUCT pCompareItem);

	void OnLButtonDown(UINT nFlags, POINT pt);
	void OnDeleteItem(LPDELETEITEMSTRUCT pDeleteItem);

	void InvalidateItem(int nIndex);

	// Attribs
	BOOL  m_bDirty;
	int   m_nTextHeight;
	int   m_nTextWidth;
};

inline int DDCheckListBox::AddString (LPCTSTR szString) 
{
	return (int)SendMessage(LB_ADDSTRING, 0, (LPARAM)(szString));
}

inline int DDCheckListBox::GetCount() 
{
	return (int)SendMessage(LB_GETCOUNT);
}

inline DWORD DDCheckListBox::GetItemData(int nIndex) 
{
	return (DWORD)SendMessage(LB_GETITEMDATA, (WPARAM) nIndex, 0);
}

inline void DDCheckListBox::SetItemData(int nIndex, DWORD dwData) 
{
	SendMessage(LB_SETITEMDATA, (WPARAM)nIndex, (LPARAM)dwData);
}

inline void DDCheckListBox::GetItemRect(int nIndex, CRect& rcItem) 
{
	SendMessage(LB_GETITEMRECT, (WPARAM)nIndex, (LPARAM)&rcItem);
}

inline int DDCheckListBox::GetCurSel() 
{
	return (int)SendMessage(LB_GETCURSEL);
}

inline void DDCheckListBox::SetCurSel(int nIndex) 
{
	SendMessage(LB_SETCURSEL, nIndex);
}

//
// CFlags
//

class CFlags : public CPropertyPopup
{
public:
	CFlags();
	~CFlags();

	enum
	{
		eListbox,
		eOk
	};
	BOOL Create(HWND hWnd);
	void Show(CRect rcPosition, CProperty* pProperty = NULL);
	void Hide();

	void Value(DWORD nValue);

private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam);
	virtual void Fill(void* pData = NULL);
	UINT GetFlags();

	DDCheckListBox m_theCheckList;
	DWORD m_nValue;
	HWND  m_hWndOk;
	int   m_nFontHeight;
};

//
// Hide
//

inline void CFlags::Hide()
{
	CPropertyPopup::Hide();
}

//
// Value
//

inline void CFlags::Value(DWORD nValue)
{
	m_nValue = nValue;
}

//
// CTabs
//

class CTabs : public FWnd
{
public:
	CTabs(CFlickerFree* pffDraw);
	~CTabs();
	
	BOOL Create(HWND hWnd, HFONT hFont);

	void Add(const DDString& strCaption);
	void CalcLayout();

	int& TabPosition();
	void TabPosition(int nTag);
private:
	struct TabInfo
	{
		DDString m_strCaption;
		CRect   m_rcTab;
	};

	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void Draw(HDC hDC, const CRect& rcWindow);
	void OnLButtonDown(UINT mFlags, POINT pt);

	TypedArray<TabInfo*> m_aTabs;
	CFlickerFree* m_pffDraw;
	HFONT		  m_hFont;
	HWND		  m_hWndParent;
	int			  m_nCurrentTab;
};

inline int& CTabs::TabPosition()
{
	return m_nCurrentTab;
}

//
// CSystemColors
//

class CSystemColors : public FWnd
{
public:
	CSystemColors(CFlickerFree* pffDraw);
	~CSystemColors();
	
	enum 
	{
		eTimerId = 1001,
		eTimerTime = 300
	};

	BOOL Create(HWND hWnd, HFONT hFont);
	
	BOOL SetSelection(COLORREF crColor);
	BOOL GetSelection(COLORREF& crColor);

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void OnDrawItem(LPDRAWITEMSTRUCT pDrawItem);

	CFlickerFree* m_pffDraw;
	HFONT		  m_hFont;
	HWND		  m_hWndParent;
	int			  m_nItemHeight;
	int			  m_nCurrentIndex;
};

//
// CPaleteColors
//

class CPaleteColors : public FWnd
{
public:
	CPaleteColors(CFlickerFree*  m_pffDraw);
	~CPaleteColors();
	
	enum 
	{
		eHorzCells = 8,
		eVertCells = 8
	};

	BOOL Create(HWND hWnd, HFONT hFont);
	void CalcLayout(const CRect& rcBound, CRect& rcWindow);

	BOOL SetSelection(COLORREF crColor);
	BOOL GetSelection(COLORREF& crColor);

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	int HitTest(const POINT& pt, BOOL Select = TRUE);
	void Draw(HDC hDC, const CRect& rcWindow);
	void UpdatePrev();

	CFlickerFree* m_pffDraw;
	HFONT		  m_hFont;
	HWND		  m_hWndParent;
	int			  m_nCurrentIndex;
	int			  m_nPrevIndex;
	int			  m_nCell;
};

//
// CColor
//

class CColorWnd : public CPropertyPopup
{
public:
	CColorWnd();
	~CColorWnd();
	
	enum 
	{
		ePalette,
		eSystem
	};

	virtual BOOL Create(HWND hWnd);
	virtual void Show(CRect rcPosition,	CProperty* pProperty = NULL);
	virtual void Hide();

	BOOL SetSelection(COLORREF crColor);
	BOOL GetSelection(COLORREF& crColor);
private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam);
	virtual void RegisterWindow();

	CSystemColors* m_pSystemColors;
	CPaleteColors* m_pPaleteColors;
	CFlickerFree*  m_pffDraw;
	CTabs*		   m_pTabs;
	HFONT		   m_hFont;
};

//
// Hide
//

inline void CColorWnd::Hide()
{
	CPropertyPopup::Hide();
}

//
// CListbox
//

class CListbox : public FWnd
{
public:
	CListbox();
	~CListbox();

	BOOL Create(HWND hWndParent); 

private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
};

//
// CBrowserCombo
//

class CBrowserCombo : public CPropertyPopup
{
public:
	enum CONSTANTS 
	{
		eDefaultNumOfLines = 8
	};

	CBrowserCombo();
	~CBrowserCombo();
	
	virtual BOOL Create(HWND hWnd);
	virtual void Show(CRect rcPosition, CProperty* pProperty = NULL);
	virtual void Hide();

	LPTSTR GetString(int& nValue);

	BOOL SetSelection(int nIndex);
	int GetSelection();

	void Clear();
	virtual void Fill(void* pData = NULL);

private:
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam);
	void Update();
	CListbox m_theListBox;
	HFONT	 m_hFont;
};

//
// Clear
//

inline void CBrowserCombo::Clear()
{
	if (IsWindow(m_hWnd))
		m_theListBox.SendMessage(LB_RESETCONTENT);
}

//
// Hide
//

inline void CBrowserCombo::Hide()
{
	CPropertyPopup::Hide();
}

//
// CBrowserEdit
//

class CBrowserEdit : public FWnd
{
public:
	CBrowserEdit();
	~CBrowserEdit();

	BOOL& ChildActive();

	BOOL Create(CBrowser* pBrowser);
	void Show(CRect rcPosition, CProperty* Property, BOOL bShowChild = FALSE, int nVirtKey = -1);
	void Hide(BOOL bHideParent = TRUE);
	void Close(BOOL bHideParent = FALSE);

	CBrowserCombo* BrowserCombo();
	CComboList* ComboList();
	CShortCut* ShortCut();
	CColorWnd* ColorWnd();
	CFlags* Flags();
	CEdit* Edit();
	void SetText(LPTSTR szText);

private:
	BOOL PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam);
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	BOOL UpdateText();
	LPTSTR GetText();
	static LRESULT CALLBACK KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);
	static HHOOK m_hHook;
	static CBrowserEdit* m_pCurrentEdit;
	BOOL OnReturn(UINT nMsg, WPARAM wParam, LPARAM lParam);
	BOOL OnEscape(UINT nMsg, WPARAM wParam, LPARAM lParam);

	CProperty*     m_pProperty;
	CBrowserCombo* m_pBrowserCombo;
	CComboList*	   m_pComboList;
	CShortCut*	   m_pShortCut;
	CColorWnd*	   m_pColor;
	CBrowser*      m_pBrowser;
	CFlags*		   m_pFlags;
	CEdit*		   m_pEdit;
	TCHAR*		   m_szPriorText;
	BOOL		   m_bChildActive;
	BOOL		   m_bProcessing;
};

//
// SetText
//

inline void CBrowserEdit::SetText(LPTSTR szText)
{
	if (m_szPriorText)
	{
		delete [] m_szPriorText;
		m_szPriorText = NULL;
	}
	if (szText)
	{
		m_szPriorText = new TCHAR[_tcslen(szText)+1];
		_tcscpy(m_szPriorText, szText);
	}
	SendMessage(WM_SETTEXT, 0, (LPARAM)szText);
}

//
// GetText
//

inline LPTSTR CBrowserEdit::GetText()
{
	static TCHAR szText[128];
	SendMessage(WM_GETTEXT, 128, (LPARAM)szText);
	return szText;
}

inline BOOL& CBrowserEdit::ChildActive()
{
	return m_bChildActive;
}

//
// Button
//

class Button : public FWnd
{
public:
	Button();
	~Button();

	enum BUTTONTYPE
	{
		eEllipse, 
		eScroll,
		eCancel
	};

	BOOL Create(CBrowser* pBrowser, CRect rc);
	void Type (BUTTONTYPE bnType);

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);

	CBrowser*  m_pBrowser;
	BUTTONTYPE m_bnType;
};

//
// CBrowser
//

class CBrowser : public FWnd
{
public:
	CBrowser(CDesignerPage* pDesigner = NULL);
	~CBrowser();

	enum CONSTANTS
	{
		eError,
		eUpdate,
		eHide,
		ePushButton = 10000
	};

	CEnum* FindEnum(BSTR bstrName);
	CProperty* CurrentProperty();

	CRect& CurrentItemRect();
	CBrowserEdit& Edit();
	void Clear();
	void SetTypeInfo(IDispatch* pDispatch, IObjectChanged* pObjectChanged);
	void SubClassAttach(HWND hWnd);
	void InvalidateItem(int nItem, BOOL bBackground = FALSE);
	void UpdatePrevSelection(int nIndex);
	void SetEditFocus(int nVirtKey);
	void Init();
	void ReleaseObjects();
private:

	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void ShowButton(CRect rc, Button::BUTTONTYPE bnType);
	void OnDrawItem(LPDRAWITEMSTRUCT pDrawItem);
	void OnDeleteItem(LPDELETEITEMSTRUCT pDeleteItem);
	void OnLButtonDblClk(UINT mFlags, POINT pt);
	void OnSelChange();
	LRESULT OnCompareItem(LPCOMPAREITEMSTRUCT pCmpItem);
	LRESULT OnKeyToItem(int nVirtKey, int nCaretPos);

	BOOL OnButtonHit(int nIndex, CProperty* pProperty, CRect& rcItem);

	BOOL ResolvePerPropertyBrowsing(UINT				 nDispId, 
								    BSTR				 bstrFunctionName, 
								    CEnum*&			     pEnum, 
								    USHORT&			     nVariantType, 
								    CProperty::EditType& etProperty, 
								    BOOL&				 bValid);

	BOOL ResolveUserDefinedType(REFCLSID			 rcClassId,
								ITypeInfo*			 pTypeInfo,
								HREFTYPE			 hRefType, 
								VARTYPE&			 vType,
								CProperty::EditType& etProperty,
								CEnum*&				 pEnum);

	BOOL ResolvePointerType(REFCLSID			 rcClassId,
						    ITypeInfo*		     pTypeInfo,
						    HREFTYPE			 hRefType,
						    VARTYPE&			 vType,
						    CProperty::EditType& etProperty);

	HRESULT BuildEnumList(CLSID      rcClassId, 
						  BSTR       bstrName, 
						  ITypeInfo* pRefTypeInfo, 
						  TYPEATTR*  pRefTypeAttr, 
						  CEnum*&    pEnum);

	void CleanUpEnums();

	IDDPerPropertyBrowsing* m_pPerPropertyBrowsing; 
	TypedArray<CEnum*> m_aEnums;
	IObjectChanged* m_pObjectChanged; 
	CFlickerFree*  m_pFF;
	CBrowserEdit   m_theEdit;
	IDispatch*     m_pObject;
	ITypeInfo*     m_pTypeInfo;
	CProperty*     m_pCurrentProperty;
	CProperty*     m_pPrevProperty;
	CDesignerPage* m_pDesigner;
	DDString	   m_strFalse;
	DDString	   m_strTrue;
	DDString	   m_strName;
	DDString	   m_strCaption;
	HBRUSH	       m_hBrushBackground;
	Button         m_thePushButton;
	CRect		   m_rcCurrentItem;
	HPEN		   m_hPen;
	BOOL		   m_bButtonPressed;
	BOOL		   m_bInClick;
	int			   m_nCurrentIndex;
	int			   m_nPrevIndex;
	int			   m_nButtonHeight;
	friend class CBrowserEdit;
	friend class CDesignerTree;
	friend class CCategroy;
};

//
// InvalidateItem
//

inline void CBrowser::InvalidateItem(int nItem, BOOL bBackground)
{
	CRect rcItem;
	LRESULT lResult = SendMessage(LB_GETITEMRECT, nItem, (LPARAM)&rcItem);
	if (LB_ERR != lResult)
		InvalidateRect(&rcItem, bBackground);
}

//
// CurrentProperty
//

inline CProperty* CBrowser::CurrentProperty()
{
	return m_pCurrentProperty;
}

//
// CurrentItemRect
//

inline CRect& CBrowser::CurrentItemRect()
{
	return m_rcCurrentItem;
}

inline CBrowserEdit& CBrowser::Edit()
{
	return m_theEdit;
}

inline void CBrowser::ReleaseObjects()
{
	if (m_pPerPropertyBrowsing)
	{
		m_pPerPropertyBrowsing->Release();
		m_pPerPropertyBrowsing = NULL;
	}
	if (m_pObject)
	{
		m_pObject->Release();
		m_pObject = NULL;
	}
	if (m_pTypeInfo)
	{
		m_pTypeInfo->Release();
		m_pTypeInfo = NULL;
	}
}

#endif