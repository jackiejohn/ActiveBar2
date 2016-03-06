#ifndef __APPPREF_H__
#define __APPPREF_H__

#include "fregkey.h"

class FXRegKey : public FRegKey
{
public:
	BOOL GetValueDWORD(LPCTSTR name,DWORD *lpdw);
	BOOL SetValuePDWORD(LPCTSTR lpszName,LPDWORD pdwValue);
};

class CAppPref 
{
public:
	CAppPref();
	~CAppPref();
	BOOL Init(LPCTSTR companyName,LPCTSTR appName);
	void SaveWinRect(int id,HWND hWnd,const RECT& rect);
	void LoadWinRect(int id,LPRECT rect,UINT *nCmdShow,const RECT& defaultRect);
	BOOL GetConfirmSuppress(UINT val);
	void SetConfirmSuppress(UINT val);
protected:
	FXRegKey prodKey; // product key set up by Init
};


#endif
