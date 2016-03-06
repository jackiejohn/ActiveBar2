#ifndef UTILITY_INCLUDED
#define UTILITY_INCLUDED
#include "Map.h"

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

interface IActiveBar2;
interface IBands;
interface IBand;
interface ITool;

enum SpecialToolIds
{
	eSpecialToolId = 0x80000000,
	eToolIdSysCommand = 0x80001000,
	eToolIdMDIButtons = 0x80001001
};


//
// CRect
//

class CRect : public RECT
{
public:
	CRect();

	CRect(const	CRect& rhs);
	CRect(const	RECT& rhs);

	CRect(const int& left, 
		  const int& right, 
		  const int& top, 
		  const int& bottom);

// Attributes
	int operator==(CRect& rhs);
	int operator!=(CRect& rhs);
	CRect& operator=(const CRect& rcBound);
	CRect& operator=(RECT& rhs);

	void SetEmpty();

	void Set(const int& left, 
			 const int& right, 
			 const int& top, 
			 const int& bottom);

// Operations
	int Height() const;
	int Width() const;
	SIZE Size() const;
	
	POINT& TopLeft();
	POINT& BottomRight();

	void Inflate(int x, int y);
	void Offset(int x, int y);
	void Intersect(const CRect& rc, const CRect& rc2);
};

//
// CRect
//

inline CRect::CRect()
{
	::SetRectEmpty(this);
}

inline CRect::CRect(const int& nLeft, 
				    const int& nRight, 
					const int& nTop, 
					const int& nBottom)
{
	::SetRect(this, nLeft, nTop, nRight, nBottom);
}

inline CRect::CRect(const CRect& rhs)
{
	::SetRect(this, rhs.left, rhs.top, rhs.right, rhs.bottom);
}

inline CRect::CRect(const RECT& rhs)
{
	::SetRect(this, rhs.left, rhs.top, rhs.right, rhs.bottom);
}

//
// SetRectEmpty
//

inline void CRect::SetEmpty()
{
	::SetRectEmpty(this);
}

//
// SetRect
//

inline void CRect::Set(const int& nLeft, 
					   const int& nTop, 
					   const int& nRight, 
					   const int& nBottom)
{
	::SetRect(this, nLeft, nTop, nRight, nBottom);
}

//
// operator==
//

inline int CRect::operator==(CRect& rhs)
{
	return memcmp(this, &rhs, sizeof(RECT)) == 0;
}

//
// operator!=
//

inline int CRect::operator!=(CRect& rhs)
{
	return memcmp(this, &rhs, sizeof(RECT)) != 0;
}


//
// operator=
//

inline CRect& CRect::operator=(const CRect& rhs)
{
	if (&rhs == this)
		return *this;
	memcpy(this, &rhs, sizeof(CRect));
	return *this;
}

//
// operator=
//

inline CRect& CRect::operator=(RECT& rhs)
{
	memcpy(this, &rhs, sizeof(CRect));
	return *this;
}

//
// Height
//

inline int CRect::Height() const
{
	return bottom - top;
}

//
// Width
//

inline int CRect::Width() const 
{
	return right - left;
}

//
// Size
//

inline SIZE CRect::Size() const
{
	SIZE size;
	size.cx = Width();
	size.cy = Height();
	return size;
}

//
// TopLeft
//

inline POINT& CRect::TopLeft()
{
	return *((POINT*)this);
}

//
// BottomRight
//

inline POINT& CRect::BottomRight()
{
	return *((POINT*)(this+1));
}

//
// Inflate
//

inline void CRect::Inflate(int x, int y)
{ 
	::InflateRect(this, x, y); 
}

//
// Offset
//

inline void CRect::Offset(int x, int y)
{
	::OffsetRect(this, x, y);
}

//
// Intersect
//

inline void CRect::Intersect(const CRect& rc, const CRect& rc2)
{
	IntersectRect(this, &rc, &rc2);
}

//
// DDString
//

class DDString
{
public:
	DDString(TCHAR* szString);
	
	DDString()
		: m_szBuffer(NULL) 
	{
	}

	~DDString();

	BSTR AllocSysString ( ) const;
	int GetLength();
	BOOL LoadString (UINT nResourceId);
	void Format(LPCTSTR szFormat, ... );
	void Format(UINT nFormatId, ... );
	operator LPCTSTR () const;
	operator LPTSTR () const;
	const DDString& operator=(const TCHAR* sz);
	const DDString& operator=(const DDString& strLeft);

private:
	void FormatV(LPCTSTR lpszFormat, va_list argList);
	TCHAR* m_szBuffer;
};

inline int DDString::GetLength()
{
	if (NULL == m_szBuffer)
		return 0;
	return _tcslen(m_szBuffer);
}

inline DDString::operator LPCTSTR() const
{
	return m_szBuffer;
}

inline DDString::operator LPTSTR() const
{
	return m_szBuffer;
}

inline const DDString& DDString::operator=(const DDString& strLeft)
{
	delete [] m_szBuffer;
	m_szBuffer = new TCHAR[_tcslen(strLeft)+1];
	_tcscpy(m_szBuffer, strLeft); 
	return *this; 
}

inline const DDString& DDString::operator=(const TCHAR* sz)
{
	delete [] m_szBuffer;
	m_szBuffer = new TCHAR[_tcslen(sz)+1];
	_tcscpy(m_szBuffer, sz); 
	return *this; 
}

//
// Report
//

struct Report
{
	Report(HANDLE hFile)
		: m_hFile(hFile)
	{
	}

	static void DoIt(IBand* pBand, IBand* pChildBand, ITool* pTool, void* pData);

	DWORD  m_dwFilePosition;
	TCHAR  m_szLine[256];
	HANDLE m_hFile;
};

//
// Header
//

struct Header
{
	Header(HANDLE hFile)
		: m_hFile(hFile)
	{
	}

	static void DoIt(IBand* pBand, IBand* pChildBand, ITool* pTool, void* pData);

	DWORD  m_dwFilePosition;
	TCHAR  m_szLine[256];
	HANDLE m_hFile;
};

//
// GenSelect
//

struct GenSelect
{
	static void DoIt(IBand* pBand, IBand* pChildBand, ITool* pTool, void* pData);
	DDString m_strSelect;
};

//
// GrabMenu
//

class GrabMenu
{
public:
	GrabMenu(HMENU hMenu, IActiveBar2* pActiveBar);

	BOOL DoIt();

	static HWND FindMenuWindow(HWND hWndParent);

	static BOOL m_bCancelled;
private:
	BOOL CreateSubMenu(IBands* pBands, ITool* pTool, BSTR bstrName, BSTR bstrCaption, HMENU hMenu, int nItem);

	IActiveBar2* m_pActiveBar; 
	HMENU m_hMenu;
};

//
// GetLastErrorMsg
//

inline LPTSTR GetLastErrorMsg()
{
	LPVOID lpMsgBuf;
	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
				  NULL,
				  GetLastError(),
				  MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
				  (LPTSTR)&lpMsgBuf,    
				  0,    
				  NULL);
	return (LPTSTR)lpMsgBuf;
}

//
// FindLastToolId
//

long FindLastToolId(IActiveBar2* pActiveBar);

//
// OleLoadPictureHelper
//

IPicture* OleLoadPictureHelper(LPCTSTR szFileName);

void RemoveAmperstand(TCHAR* szString);

void CleanCaption(BSTR bstrName, WCHAR* wCaption);

typedef void (*PFNMODIFYTOOL) (IBand* pBand, IBand* pChildBand, ITool* pTool, void* pData);
BOOL VisitBarTools(IActiveBar2* pActiveBar, long nToolId, void* pData, PFNMODIFYTOOL pToolFunction);
BOOL VisitBandTools(IActiveBar2* pActiveBar, long nToolId, void* pData, PFNMODIFYTOOL pToolFunction);

BOOL WinHelp(HWND hWndMain, UINT uCommand, DWORD dwData);

HRESULT StWriteBSTR(IStream* pStream, BSTR bstr);

HRESULT StReadBSTR(IStream* pStream, BSTR& bstr);


#endif