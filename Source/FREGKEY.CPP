//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "fregkey.h"
#include <assert.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define keysecurity KEY_ALL_ACCESS

BOOL FRegKey::Open(HKEY rootKey,LPCTSTR subKey,BOOL createIfNotExists)
{
	DWORD temp;
	CloseKey();
	if (createIfNotExists)
	{
		lastError=RegCreateKeyEx(rootKey,subKey,0,REG_NONE,REG_OPTION_NON_VOLATILE,
			keysecurity,NULL,&hKey,&temp);

		if (lastError!=ERROR_SUCCESS)
		{
			lastError=RegCreateKeyEx(rootKey,subKey,0,REG_NONE,REG_OPTION_NON_VOLATILE,
				KEY_WRITE|KEY_READ,NULL,&hKey,&temp);
		}
	}
	else
	{
		lastError=RegOpenKeyEx(rootKey,subKey,0,keysecurity,&hKey);
		if (lastError!=ERROR_SUCCESS)
			lastError=RegOpenKeyEx(rootKey,subKey,0,KEY_WRITE|KEY_READ,&hKey);
	}
	
	if (lastError==ERROR_SUCCESS)
		return TRUE;
	return FALSE;
}

void FRegKey::CloseKey()
{
	if (hKey!=NULL)
		lastError=RegCloseKey(hKey);
	hKey=NULL;
}

BOOL FRegKey::SetValue(LPCTSTR lpszName,LPCTSTR lpszValue)
{
	lastError=RegSetValueEx(hKey,lpszName,NULL,REG_SZ,
		(BYTE * const) lpszValue,(lstrlen(lpszValue)+1)*sizeof(TCHAR));

	assert(lastError==ERROR_SUCCESS);

	return lastError==ERROR_SUCCESS;
}

BOOL FRegKey::SetValue(LPCTSTR lpszName,DWORD dwValue)
{
	lastError=RegSetValueEx(hKey,lpszName,NULL,REG_DWORD,
		(BYTE * const) dwValue,sizeof(DWORD));

	assert(lastError==ERROR_SUCCESS);

	return lastError==ERROR_SUCCESS;
}

BOOL FRegKey::SetKeyValue(LPCTSTR subKey,LPCTSTR value)
{
	FRegKey newKey;
	newKey.Create(hKey,subKey);
	newKey.SetValue(NULL,value);
	return TRUE;
}

#define GUID_STR_LEN    40
#ifndef UNICODE
int StringFromGuidA(REFIID   riid,LPSTR pszBuf)
{
    return wsprintf((char *)pszBuf, "{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}", riid.Data1, 
            riid.Data2, riid.Data3, riid.Data4[0], riid.Data4[1], riid.Data4[2], 
            riid.Data4[3], riid.Data4[4], riid.Data4[5], riid.Data4[6], riid.Data4[7]);
}
#endif


BOOL RecurseSubKeyDelete(HKEY hkIn,LPCTSTR pszSubKey)
{
    HKEY  hk;
    TCHAR szTmp[MAX_PATH];
    DWORD dwTmpSize;
    long  l;
    BOOL  f;
    int   x;

    l = RegOpenKeyEx(hkIn, pszSubKey, 0, keysecurity, &hk);
    if (l != ERROR_SUCCESS) 
		l = RegOpenKeyEx(hkIn, pszSubKey, 0, KEY_READ|KEY_WRITE, &hk);
    if (l != ERROR_SUCCESS) return FALSE;

    // loop through all subkeys, blowing them away.
    //
    f = TRUE;
    x = 0;
    while (f) {
        dwTmpSize = MAX_PATH;
        l = RegEnumKeyEx(hk, 0, szTmp, &dwTmpSize, 0, NULL, NULL, NULL);
        if (l != ERROR_SUCCESS) break;
        f = RecurseSubKeyDelete(hk, szTmp);
        x++;
    }

    RegCloseKey(hk);
    l = RegDeleteKey(hkIn, pszSubKey);

    return (l == ERROR_SUCCESS) ? TRUE : FALSE;
}

BOOL FRegKey::DeleteSubKey(LPCTSTR subKey)
{
	return RecurseSubKeyDelete(hKey,subKey);
}
