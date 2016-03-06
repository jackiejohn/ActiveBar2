#ifndef __APPPREF_H__
#define __APPPREF_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "FRegkey.h"
#include "Utility.h"

class FXRegKey : public FRegKey
{
public:
	BOOL GetValueDWORD(LPCTSTR szName, DWORD& dwValue);
	BOOL SetValueDWORD(LPCTSTR szName, const DWORD& dwValue);
};

class CAppPref 
{
public:
	CAppPref(LPCTSTR szCompanyName, LPCTSTR szAppName);
	~CAppPref();
	
	BOOL GetConfirmSuppress(UINT nVal);
	void SetConfirmSuppress(UINT nVal, DWORD dwResult = 1);

	COLORREF GetColor(long nIndex);
	void SetColor(long nVal, COLORREF crColor);

	BOOL GetLibraryPath(LPTSTR szFile);
	void SetLibraryPath(LPTSTR szFile);
	void SaveWinRect(int nId, HWND hWnd);
	void SaveWinRect(int nId, const CRect& rc);
	void LoadWinRect(int nId, CRect& rc, const CRect& rcDefault);
	BOOL LoadWinRect(int nId, CRect& rc);
	DWORD GetAutoId();
	void SetAutoId(DWORD dwAutoId);
protected:
	FXRegKey m_rkProduct;
};

#endif
