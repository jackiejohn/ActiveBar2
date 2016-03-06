#ifndef __FWND_H__
#define __FWND_H__

//
//  Copyright (c) 1995-1996, Data Dynamics. All rights reserved.
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//

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
				HMENU   hMenu = 0,
				LPVOID  lParam = 0);

	HWND CreateEx(DWORD   dwExStyle,
				  LPCTSTR szWindowName,
				  DWORD   dwStyle,
				  int	  nLeft,
				  int	  nTop,
				  int	  nWidth,	
				  int	  nHeight,
   				  HWND    hWndParent,
				  HMENU   hMenu = 0,
				  LPVOID  lParam = 0);

	BOOL RegisterWindow(LPCTSTR szClassName,
					    UINT    nStyle,
						HBRUSH  hbBackground,
						HCURSOR hCursor,
						HICON   hIcon = 0,
						LPCTSTR szMenuName = 0);

	virtual LRESULT WindowProc(UINT nMessage, WPARAM wParam, LPARAM lParam);

	static int CheckMap();
	static void Init();
	static void CleanUp();
	HWND hWnd() const;
	BOOL ShowWindow(int nCmdShow);
	BOOL UpdateWindow();
	BOOL MoveWindow(const CRect& rcWindow, BOOL bPaint = TRUE);
	BOOL InvalidateRect(RECT* rcUpdate, BOOL bErase);

	void RemoveFromMap();
	BOOL ClientToScreen(POINT& pt) const;
	BOOL ClientToScreen(RECT& rc) const;
	BOOL ScreenToClient(POINT& pPoint) const;
	BOOL ScreenToClient(RECT& rc) const;

	BOOL GetClientRect(RECT& rc) const;
	BOOL GetWindowRect(RECT& rc) const;

	BOOL IsWindow();

	BOOL SetWindowPos(HWND hWndInsertAfter, int X, int Y, int cx, int cy, UINT uFlags);

	void SetClassName(LPCTSTR szClassName);
	
	void SetFont(HFONT hFont);
	
	HWND SetFocus();
	HFONT GetFont();
	
	void SubClassAttach(HWND hWnd = 0);
	void UnsubClass();
	
	int MessageBox(LPTSTR szMsg, UINT nMsgType = MB_OK);
	int MessageBox(UINT nResourceId, UINT nMsgType = MB_OK);

	LRESULT SendMessage(UINT nMsg, WPARAM wParam = 0, LPARAM lParam = 0);
	LRESULT PostMessage(UINT nMsg, WPARAM wParam = 0, LPARAM lParam = 0);
	BOOL DestroyWindow();

protected:

	HWND    m_hWnd;
	LPCTSTR m_szClassName;
	DDString m_szMsgTitle;
	BOOL    m_bSubClassed;
	WNDPROC m_wpWindowProc;

	static LRESULT CALLBACK FWndWindowProc(HWND hWnd, UINT nMessage, WPARAM wParam, LPARAM lParam);
};

inline HWND FWnd::hWnd() const
{ 
	return m_hWnd;
}

inline int FWnd::MessageBox(LPTSTR szMsg, UINT nMsgType)
{
	assert(IsWindow());
	return ::MessageBox (m_hWnd, szMsg, m_szMsgTitle, nMsgType);
}

inline int FWnd::MessageBox(UINT nResourceId, UINT nMsgType)
{
	assert(IsWindow());
	DDString strMsg;
	strMsg.LoadString(nResourceId);
	return ::MessageBox (m_hWnd, strMsg, m_szMsgTitle, nMsgType);
}

inline BOOL FWnd::SetWindowPos(HWND hWndInsertAfter, int X, int Y, int cx, int cy, UINT uFlags)
{
	assert(IsWindow());
	return ::SetWindowPos(m_hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
}

inline LRESULT FWnd::SendMessage(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	assert (IsWindow());
	return ::SendMessage(m_hWnd, nMsg, wParam, lParam);
}

inline LRESULT FWnd::PostMessage(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	assert (IsWindow());
	return ::PostMessage(m_hWnd, nMsg, wParam, lParam);
}
inline BOOL FWnd::ClientToScreen(POINT& pt) const
{ 
	assert(::IsWindow(m_hWnd)); 
	return ::ClientToScreen(m_hWnd, &pt); 
}

inline BOOL FWnd::ClientToScreen(RECT& rc) const
{ 
	assert(::IsWindow(m_hWnd)); 
	::ClientToScreen(m_hWnd, (LPPOINT)&rc);
	return ::ClientToScreen(m_hWnd, ((LPPOINT)&rc)+1); 
}

inline BOOL FWnd::ScreenToClient(POINT& pt) const
{ 
	assert(::IsWindow(m_hWnd)); 
	return ::ScreenToClient(m_hWnd, &pt); 
}

inline BOOL FWnd::ScreenToClient(RECT& rc) const
{ 
	assert(::IsWindow(m_hWnd)); 
	::ScreenToClient(m_hWnd, (LPPOINT)&rc);
	return ::ScreenToClient(m_hWnd, ((LPPOINT)&rc)+1); 
}

inline BOOL FWnd::ShowWindow(int nCmdShow)
{
	assert(IsWindow());
	return ::ShowWindow(m_hWnd, nCmdShow);
}

inline BOOL FWnd::UpdateWindow()
{
	assert(IsWindow());
	return ::UpdateWindow(m_hWnd);
}

inline HWND FWnd::SetFocus()
{
	assert(IsWindow());
	return ::SetFocus(m_hWnd);
}

inline void FWnd::SetFont(HFONT hFont)
{
	assert(IsWindow());
	::SendMessage(m_hWnd, WM_SETFONT, (WPARAM)hFont, TRUE);
}

inline HFONT FWnd::GetFont()
{
	assert(IsWindow());
	return (HFONT)::SendMessage(m_hWnd, WM_GETFONT, 0, 0);
}

inline BOOL FWnd::MoveWindow(const CRect& rc, BOOL bPaint)
{
	assert(IsWindow());
	return ::MoveWindow(m_hWnd, rc.left, rc.top, rc.Width(), rc.Height(), bPaint);
}

inline BOOL FWnd::InvalidateRect(RECT* pRC, BOOL bErase)
{
	assert(IsWindow());
	return ::InvalidateRect(m_hWnd, pRC, bErase);
}

inline void FWnd::SetClassName(LPCTSTR szClassName)
{
	m_szClassName = szClassName;
}

inline BOOL FWnd::GetClientRect(RECT& rc) const
{
	assert(::IsWindow(m_hWnd)); 
	return ::GetClientRect(m_hWnd, &rc);
}

inline BOOL FWnd::GetWindowRect(RECT& rc) const
{
	assert(::IsWindow(m_hWnd)); 
	return ::GetWindowRect(m_hWnd, &rc);
}

inline BOOL FWnd::DestroyWindow()
{
	if (IsWindow())
		return ::DestroyWindow(m_hWnd);
	return FALSE;
}

inline BOOL FWnd::IsWindow()
{
	return ::IsWindow(m_hWnd);
}

#endif
