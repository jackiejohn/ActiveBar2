#ifndef __TEMPPOPUP_H__
#define __TEMPPOPUP_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Flicker.h"
#include <RichEdit.h>

class CBar;
class CTool;
class ActiveCombobox;

//
// ActiveEdit
//

class ActiveEdit : public FWnd
{
public:
	ActiveEdit();
	~ActiveEdit();
	BOOL Create(HWND hWndParent, const CRect& rc, BOOL bXPLook);
	void Exchange(BOOL bFireEvent = TRUE);
	void SetTool(CTool* pTool);
private:
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	CTool* m_pTool;
	DWORD m_dwStyle;
};

inline ActiveEdit::ActiveEdit()
	: m_pTool(NULL)
{
	m_szClassName = _T("EDIT");
	m_dwStyle = ES_LEFT|ES_AUTOHSCROLL|ES_WANTRETURN|WS_CHILDWINDOW;
}

inline void ActiveEdit::SetTool(CTool* pTool)
{
	m_pTool = pTool;
}

//
// ActiveCombobox
//

class ActiveCombobox : public FWnd
{
public:
	ActiveCombobox();
	~ActiveCombobox();
	BOOL Create(HWND hWndParent, const CRect& rc, DWORD dwStyle);
	void SetTool(CTool* pTool);
	DWORD Style();

private:
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	CTool* m_pTool;
	DWORD m_dwStyle;
};

inline ActiveCombobox::ActiveCombobox()
	: m_pTool(NULL)
{
	m_szClassName = _T("COMBOBOX");
	m_dwStyle = 0;
}

inline void ActiveCombobox::SetTool(CTool* pTool)
{
	m_pTool = pTool;
}

inline DWORD ActiveCombobox::Style()
{
	return m_dwStyle;
}

#endif