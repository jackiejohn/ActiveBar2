#ifndef __FREGKEY_H__
#define __FREGKEY_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class FRegKey
{
public:
	FRegKey() {hKey=NULL;};
	~FRegKey() {CloseKey();};
	operator HKEY() const;
	BOOL Open(HKEY rootKey,LPCTSTR subKey,BOOL createIfNotExists=FALSE);
	inline BOOL Create(HKEY rootKey,LPCTSTR subKey) {return Open(rootKey,subKey,TRUE);};
	BOOL SetValue(LPCTSTR lpszName,LPCTSTR lpszValue);
	BOOL GetValue(LPCTSTR lpszName,LPCTSTR lpszValue);
	BOOL SetValue(LPCTSTR lpszName,DWORD lpszValue);
	BOOL SetKeyValue(LPCTSTR subkey,LPCTSTR value);
	BOOL DeleteSubKey(LPCTSTR subkey);
	void CloseKey();
protected:
	HKEY hKey;
	LONG lastError;
};

inline FRegKey::operator HKEY() const
	{ return hKey; }

#ifndef UNICODE
int StringFromGuidA(REFIID   riid,LPSTR pszBuf);
#endif

extern BOOL RecurseSubKeyDelete(HKEY hkIn,LPCTSTR pszSubKey);

#endif
