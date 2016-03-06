#ifndef CUSTOM_INCLUDED
#define CUSTOM_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "CustomizeListbox.h"

class CDDPropertySheet;
class ShortCutStore;
class CTool;
class CBar;

//
// CDDProppertyPage
//

class CDDPropertyPage
{
public:
	CDDPropertyPage(const UINT& nId);

	static void Init();
	static void Cleanup();

	const PROPSHEETPAGE& PageInfo();

	virtual BOOL OnPageInit();
	virtual BOOL OnDestroy();
	virtual BOOL OnNotify(const int& nControlId, const LPNMHDR pNMHDR, BOOL& bResult);
	virtual BOOL OnCommand(const WORD& wNotifyCode, const WORD& wID, const HWND hWnd);
	virtual BOOL OnMessage(const UINT& nMsg, const WPARAM& wParam, const LPARAM& lParam);
	
	BOOL ExecuteDlgInit();
	BOOL ExecuteDlgInit(LPVOID pResource);

	void hWnd(const HWND& hWnd) const;
	const HWND& hWnd() const;
	CDDPropertySheet* m_pSheet;

protected:
	PROPSHEETPAGE m_pspInfo;
	mutable HWND  m_hWnd;
	TCHAR m_szTitle[MAX_PATH];

	static BOOL PropertyPageProc(HWND hWndPage, UINT nMsg, UINT wParam, LONG lParam);
	static TypedMap<HWND, CDDPropertyPage*>* m_pmapPages;
};

//
// OnPageInit
//

inline BOOL CDDPropertyPage::OnPageInit()
{ 
	return TRUE; 
};

//
// OnDestroy 
//

inline BOOL CDDPropertyPage::OnDestroy() 
{ 
	return FALSE; 
};

//
// OnNotify 
//

inline BOOL CDDPropertyPage::OnNotify(const int&    nControlId, 
									  const LPNMHDR pNMHDR, 
									  BOOL&         bResult) 
{ 
	switch (pNMHDR->code)
	{
	case PSN_HELP:
		SendMessage(m_hWnd, WM_COMMAND, 0, 0);
		break;
	}
	return FALSE; 
};

//
// OnCommand 
//

inline BOOL CDDPropertyPage::OnCommand(const WORD& wNotifyCode, 
									   const WORD& wID, 
									   const HWND  hWnd) 
{ 
	return FALSE; 
};

//
// OnMessage
//

inline BOOL CDDPropertyPage::OnMessage(const UINT&   nMsg, 
									   const WPARAM& wParam, 
									   const LPARAM& lParam) 
{ 
	return FALSE; 
};

//
// hWnd
//

inline void CDDPropertyPage::hWnd(const HWND& hWnd) const
{
	m_hWnd = hWnd;
}

inline const HWND& CDDPropertyPage::hWnd() const
{
	return m_hWnd;
}

//
// PageInfo
//

inline const PROPSHEETPAGE& CDDPropertyPage::PageInfo()
{
	return m_pspInfo;
}

//
// CDDPropertySheet
//

class CDDPropertySheet
{
public:
	CDDPropertySheet(HWND hWndParent);
	~CDDPropertySheet();

	static void Init();
	static void Cleanup();

	void AddPage(CDDPropertyPage* pPage);
	void hWnd(const HWND& hWnd) const;
	const HWND& hWnd() const;

	virtual HWND ShowSheet();
	virtual BOOL Help(LPHELPINFO pHelpInfo) = 0;
	virtual BOOL PreTranslateMessage(WPARAM wParam, LPARAM lParam);
	virtual void OnMinimize(BOOL bActivate);

protected:
	void SubClassAttach(HWND hWnd);

	TypedArray<CDDPropertyPage*> m_faPages;
	PROPSHEETHEADER m_pshInfo;
	TCHAR			m_szTitle[MAX_PATH];
	mutable HWND	m_hWnd;

	static LRESULT CALLBACK MessageProc(int nCode, WPARAM wParam, LPARAM lParam);
	static BOOL CALLBACK DialogProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);
	static TypedArray<CDDPropertySheet*>* m_paSheets;
	static HHOOK m_hMessageHookProc;
	static BOOL m_bHookSet;
	static WNDPROC CDDPropertySheet::m_wpDialogProc;
};

//
// AddPage
//

inline void CDDPropertySheet::AddPage(CDDPropertyPage* pPage)
{
	HRESULT hResult = m_faPages.Add(pPage);
	assert(SUCCEEDED(hResult));
	pPage->m_pSheet = this;
}

//
// hWnd
//

inline void CDDPropertySheet::hWnd(const HWND& hWnd) const
{
	m_hWnd = hWnd;
}

inline const HWND& CDDPropertySheet::hWnd() const
{
	return m_hWnd;
}

class CCustomizeListboxHost;

//
// CLabel
//

class CLabel : public FWnd
{
public:
	CLabel();
	virtual ~CLabel();
	static void RegisterClass();

private:	
	static LRESULT _stdcall WindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);
	static HFONT m_hFont;
};


//
// CKeyEdit
//

class CKeyEdit : public FWnd
{
public:
	CKeyEdit(CBar* pBar);
	virtual ~CKeyEdit();
	
	enum Constants
	{
		eStringSize = 64
	};

	TCHAR* GetShortCut();
	void GetShortCut(ShortCutStore*& pShortCutStore);
	void Clear();

private:	

	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);

	TCHAR m_szText[eStringSize];
	ShortCutStore* m_pShortCutStore;
};

//
// GetShortCut
//

inline TCHAR* CKeyEdit::GetShortCut()
{
	GetDlgItemText(GetParent(m_hWnd), IDC_ED_SHORTCUT, m_szText, eStringSize);
	return m_szText;
}

inline void CKeyEdit::GetShortCut(ShortCutStore*& pShortCutStore)
{
	pShortCutStore = m_pShortCutStore;
}

//
// KeyBoard
//

class KeyBoard : public CCustomizeListboxHost, public FDialog
{
public:
	KeyBoard(CBar* pBar);
	~KeyBoard();

private:
	virtual void FireSelChangeCategory(BSTR strCategory);
	virtual void FireSelChangeTool(ITool* pTool);

	virtual BOOL DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam);

	void Assign();
	void Remove();
	void ResetAll();

	void Customize(UINT nId, LocalizationTypes eType);

	CCustomizeListbox m_theCustomizeCommands;
	CCustomizeListbox m_theCustomizeTools;

	DWORD m_dwCommandsEventCookie;
	DWORD m_dwToolsEventCookie;

	CKeyEdit m_theKeyEdit;
	CTool*   m_pSelectedTool;
	CBar*    m_pBar;
};

//
// ResetAll
//

inline void KeyBoard::ResetAll()
{
	//
	// Clear the shortcut listbox
	//

	SendMessage(GetDlgItem(m_hWnd, IDC_LB_KEYS), LB_RESETCONTENT, 0, 0);
}

//
// Class CToolBar
//

class CToolBar : public CCustomizeListboxHost, public CDDPropertyPage
{
public:
	CToolBar(CBar* pBar);
	~CToolBar();

	void Bar(CBar* pBar);

	void Clear();
	void Reset();

private:
	virtual BOOL OnPageInit();
	virtual BOOL OnCommand(const WORD& wNotifyCode, const WORD& wID, const HWND hWnd);
	virtual BOOL OnNotify(const int& nControlId, const LPNMHDR pNMHDR, BOOL& bResult);
	
	virtual void FireSelChangeBand(IBand* pBand);

	CCustomizeListbox m_theCustomizeListbox;
	DWORD m_dwEventCookie;
	BSTR  m_bstrCurrentBandName;
	CBar* m_pBar;
};

//
// Bar
//

inline void CToolBar::Bar(CBar* pBar)
{
	m_pBar = pBar;
}

//
// Class CCommands
//

class CCommands : public CCustomizeListboxHost, public CDDPropertyPage
{
public:
	CCommands(CBar* pBar);
	~CCommands();
	
	void Bar(CBar* pBar);

	void Clear();
	void Reset();

private:
	virtual BOOL OnPageInit();
	virtual BOOL OnDestroy();
	virtual BOOL OnCommand(const WORD& wNotifyCode, const WORD& wID, const HWND hWnd);
	virtual BOOL OnNotify(const int& nControlId, const LPNMHDR pNMHDR, BOOL& bResult);

	virtual void FireSelChangeCategory(BSTR strCategory);
	virtual void FireSelChangeTool(ITool* pTool);

	CCustomizeListbox m_theCustomizeCommands;
	CCustomizeListbox m_theCustomizeTools;
	DWORD m_dwCommandsEventCookie;
	DWORD m_dwToolsEventCookie;
	CBar* m_pBar;
};

//
// Bar
//

inline void CCommands::Bar(CBar* pBar)
{
	m_pBar = pBar;
}

//
// Class COptions
//

class COptions : public CDDPropertyPage
{
public:
	COptions(CBar* pBar);
	void Bar(CBar* pBar);

	void Clear();
	void Reset();

private:
	virtual BOOL OnPageInit();
	virtual BOOL OnCommand(const WORD& wNotifyCode, const WORD& wID, const HWND hWnd);
	virtual BOOL OnNotify(const int& nControlId, const LPNMHDR pNMHDR, BOOL& bResult);

	CBar* m_pBar;
};

//
// Bar
//

inline void COptions::Bar(CBar* pBar)
{
	m_pBar = pBar;
}

//
// Customize
//

class CCustomize : public CDDPropertySheet
{
public: 
	CCustomize(HWND hWndParent, CBar* pBar);
	
	enum 
	{
		eKeyboard = 1010,
		eClose = 1011
	};

	void Clear();
	void Reset();

	void EndCustomize();
	virtual BOOL Help(LPHELPINFO pHelpInfo);
	virtual void OnMinimize(BOOL bActivate);
private:
	CToolBar  m_pageToolBar;
	CCommands m_pageCommands;
	COptions  m_pageOptions;
	CBar*     m_pBar;
};

#endif