#ifndef __FWND_H__
#define __FWND_H__

//
//  Copyright (c) 1995-1996, Data Dynamics. All rights reserved.
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//

class SEException;

LRESULT CALLBACK FWndWindowProc(HWND hWnd, UINT nMessage, WPARAM wParam, LPARAM lParam);

class FWnd
{
public:
	FWnd();
	virtual ~FWnd();

	HWND Create(LPCTSTR szWindowName,
				DWORD   dwStyle,
				int	    nLeft,
				int	    nTop,
				int	    nWidth,	
				int	    nHeight,
   				HWND    hWndParent,
				HMENU   hMenu = NULL,
				LPVOID  lParam = NULL);

	HWND MDIChildCreate(LPTSTR szWindowName,
					    DWORD  dwStyle,
					    int	   nLeft,
					    int	   nTop,
					    int	   nWidth,	
					    int	   nHeight,
   					    HWND   hWndParent,
						LPARAM lParam = NULL);

	HWND CreateEx(DWORD   dwExStyle,
				  LPCTSTR szWindowName,
				  DWORD   dwStyle,
				  int	  nLeft,
				  int	  nTop,
				  int	  nWidth,	
				  int	  nHeight,
   				  HWND    hWndParent,
				  HMENU   hMenu = NULL,
				  LPVOID  lParam = NULL);

	BOOL RegisterWindow(LPCTSTR szClassName,
					    UINT    nStyle,
						HBRUSH  hbBackground,
						HCURSOR hCursor,
						HICON   hIcon = NULL,
						LPCTSTR szMenuName = NULL);

	virtual LRESULT WindowProc(UINT nMessage, WPARAM wParam, LPARAM lParam);

	HWND hWnd() const;
	BOOL ShowWindow(int nCmdShow);
	BOOL UpdateWindow();
	BOOL MoveWindow(RECT& rcWindow, BOOL bPaint = TRUE);
	void InvalidateRect(RECT* rcUpdate, BOOL bErase);
	
	void ClientToScreen(POINT& pt) const;
	void ClientToScreen(RECT& rc) const;
	void ScreenToClient(POINT& pPoint) const;
	void ScreenToClient(RECT& rc) const;

	void GetClientRect(RECT& rc) const;
	void GetWindowRect(RECT& rc) const;

	void SetClassName(LPCTSTR szClassName);
	
	void SetFont(HFONT hFont);
	HFONT GetFont();
	
	void SetFocus();
	
	DWORD GetStyle();

	void SubClassAttach(HWND hWnd = NULL);
	void UnsubClass();
	
	BOOL SetWindowText(LPCTSTR szText);
	BOOL SetWindowPos(HWND hWndInsertAfter, int X, int Y, int cx, int cy, UINT uFlags);

	int MessageBox(LPTSTR szMsg, UINT nMsgType = MB_OK);
	int MessageBox(UINT nResourceId, UINT nMsgType = MB_OK);

	LRESULT SendMessage(UINT nMsg, WPARAM wParam = 0, LPARAM lParam = 0);
	LRESULT PostMessage(UINT nMsg, WPARAM wParam = 0, LPARAM lParam = 0);

	BOOL DestroyWindow();

protected:
	DDString m_szMsgTitle;
	LPCTSTR m_szClassName;
	WNDPROC m_wpWindowProc;
	BOOL    m_bSubClassed;
	HWND    m_hWnd;

	friend LRESULT CALLBACK FWndWindowProc(HWND hWnd, UINT nMessage, WPARAM wParam, LPARAM lParam);
};

inline DWORD FWnd::GetStyle()
{
	assert(IsWindow(m_hWnd));
	return GetWindowLong(m_hWnd, GWL_STYLE);
}

inline HWND FWnd::hWnd() const
{ 
	return m_hWnd;
}

inline int FWnd::MessageBox(LPTSTR szMsg, UINT nMsgType)
{
	if (IsWindow(m_hWnd))
		return ::MessageBox (m_hWnd, szMsg, m_szMsgTitle, nMsgType);
	return 0;
}

inline int FWnd::MessageBox(UINT nResourceId, UINT nMsgType)
{
	DDString strMsg;
	strMsg.LoadString(nResourceId);
	if (IsWindow(m_hWnd))
		return ::MessageBox (m_hWnd, strMsg, m_szMsgTitle, nMsgType);
	return 0;
}

inline BOOL FWnd::SetWindowPos(HWND hWndInsertAfter, int X, int Y, int cx, int cy, UINT uFlags)
{
	return ::SetWindowPos(m_hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
}

inline LRESULT FWnd::SendMessage(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	assert (IsWindow(m_hWnd));
	return ::SendMessage(m_hWnd, nMsg, wParam, lParam);
}

inline LRESULT FWnd::PostMessage(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	assert (IsWindow(m_hWnd));
	return ::PostMessage(m_hWnd, nMsg, wParam, lParam);
}

inline void FWnd::ClientToScreen(POINT& pt) const
{ 
	assert(::IsWindow(m_hWnd)); 
	::ClientToScreen(m_hWnd, &pt); 
}

inline void FWnd::ClientToScreen(RECT& rc) const
{ 
	assert(::IsWindow(m_hWnd)); 
	::ClientToScreen(m_hWnd, (LPPOINT)&rc);
	::ClientToScreen(m_hWnd, ((LPPOINT)&rc)+1); 
}

inline void FWnd::ScreenToClient(POINT& pt) const
{ 
	assert(::IsWindow(m_hWnd)); 
	::ScreenToClient(m_hWnd, &pt); 
}

inline void FWnd::ScreenToClient(RECT& rc) const
{ 
	assert(::IsWindow(m_hWnd)); 
	::ScreenToClient(m_hWnd, (LPPOINT)&rc);
	::ScreenToClient(m_hWnd, ((LPPOINT)&rc)+1); 
}

inline BOOL FWnd::ShowWindow(int nCmdShow)
{
	return ::ShowWindow(m_hWnd, nCmdShow);
}

inline BOOL FWnd::UpdateWindow()
{
	return ::UpdateWindow(m_hWnd);
}

inline void FWnd::SetFocus()
{
	::SetFocus(m_hWnd);
}

inline void FWnd::SetFont(HFONT hFont)
{
	::SendMessage(m_hWnd, WM_SETFONT, (WPARAM)hFont, TRUE);
}

inline HFONT FWnd::GetFont()
{
	return (HFONT)::SendMessage(m_hWnd, WM_GETFONT, 0, 0);
}

inline BOOL FWnd::MoveWindow(RECT& r, BOOL bPaint)
{
	return ::MoveWindow(m_hWnd, r.left, r.top, r.right-r.left, r.bottom-r.top, bPaint);
}

inline void FWnd::InvalidateRect(RECT* pRC, BOOL bErase)
{
	::InvalidateRect(m_hWnd, pRC, bErase);
}

inline void FWnd::SetClassName(LPCTSTR szClassName)
{
	m_szClassName = szClassName;
}

inline void FWnd::GetClientRect(RECT& rc) const
{
	::GetClientRect(m_hWnd, &rc);
}

inline void FWnd::GetWindowRect(RECT& rc) const
{
	::GetWindowRect(m_hWnd, &rc);
}

inline BOOL FWnd::DestroyWindow()
{
	assert(IsWindow(m_hWnd));
	return ::DestroyWindow(m_hWnd);
}

inline BOOL FWnd::SetWindowText(LPCTSTR szText)
{
	return ::SetWindowText(m_hWnd, szText);
}

inline void ClientToScreen(HWND hWnd, RECT& rc)
{ 
	assert(::IsWindow(hWnd)); 
	::ClientToScreen(hWnd, (LPPOINT)&rc);
	::ClientToScreen(hWnd, ((LPPOINT)&rc)+1); 
}

inline void ScreenToClient(HWND hWnd, RECT& rc)
{ 
	assert(::IsWindow(hWnd)); 
	::ScreenToClient(hWnd, (LPPOINT)&rc);
	::ScreenToClient(hWnd, ((LPPOINT)&rc)+1); 
}

#endif
