//=--------------------------------------------------------------------------=
// Debug.Cpp
//=--------------------------------------------------------------------------=
// Copyright 1995-1996 Microsoft Corporation.  All Rights Reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
//=--------------------------------------------------------------------------=
//
// contains various methods that will only really see any use in DEBUG builds
//
#include "precomp.h"
//#include "ipserver.h"
#include <stdlib.h>
#include "support.h"

#ifdef _DEBUG


//=--------------------------------------------------------------------------=
// Private Constants
//---------------------------------------------------------------------------=
//
static const TCHAR szFormat[]  = _T("%s\nFile %s, Line %d");
static const TCHAR szFormat2[] = _T("%s\n%s\nFile %s, Line %d");

static const TCHAR szTitle[]  = _T( "Assert  (Abort = UAE, Retry = INT 3, Ignore = Continue)");



//=--------------------------------------------------------------------------=
// Local functions
//=--------------------------------------------------------------------------=
int NEAR _IdMsgBox(LPTSTR pszText, LPCTSTR pszTitle, UINT mbFlags);

//=--------------------------------------------------------------------------=
// DisplayAssert
//=--------------------------------------------------------------------------=
// Display an assert message box with the given pszMsg, pszAssert, source
// file name, and line number. The resulting message box has Abort, Retry,
// Ignore buttons with Abort as the default.  Abort does a FatalAppExit;
// Retry does an int 3 then returns; Ignore just returns.
//
VOID DisplayAssert
(
    LPSTR	 pszMsg,
    LPSTR	 pszAssert,
    LPSTR	 pszFile,
    UINT	 line
)
{
    TCHAR	szMsg[250];
    LPTSTR	lpszText;

#ifdef UNICODE
	WCHAR wszFile[MAX_PATH];
	WCHAR wszAssert[MAX_PATH];
	WCHAR wszMsg[MAX_PATH];
	MultiByteToWideChar(CP_ACP, 0, pszMsg, -1, wszMsg, MAX_PATH);
	MultiByteToWideChar(CP_ACP, 0, pszFile, -1, wszFile, MAX_PATH);
	MultiByteToWideChar(CP_ACP, 0, pszAssert, -1, wszAssert, MAX_PATH);

#endif


#ifdef UNICODE
    lpszText = wszMsg;		// Assume no file & line # info
#else
	lpszText = pszMsg;		// Assume no file & line # info
#endif

    // If C file assert, where you've got a file name and a line #
    //
    if (pszFile) {

        // Then format the assert nicely
#ifdef UNICODE        
        wsprintf(szMsg, szFormat, (pszMsg&&*pszMsg) ? wszMsg : wszAssert, wszFile, line);
#else
		wsprintf(szMsg, szFormat, (pszMsg&&*pszMsg) ? pszMsg : pszAssert, pszFile, line);
#endif
        lpszText = szMsg;
    }

    // Put up a dialog box
    //
	switch (_IdMsgBox(lpszText, szTitle, MB_ICONHAND|MB_ABORTRETRYIGNORE|MB_SYSTEMMODAL)) {
        case IDABORT:
            FatalAppExit(0, lpszText);
            return;

        case IDRETRY:
            // call the win32 api to break us.
            //
            DebugBreak();
            return;
    }

    return;
}


//=---------------------------------------------------------------------------=
// Beefed-up version of WinMessageBox.
//=---------------------------------------------------------------------------=
//
int NEAR _IdMsgBox
(
    LPTSTR	pszText,
    LPCTSTR	pszTitle,
    UINT	mbFlags
)
{
    HWND hwndActive;
    int  id;

    hwndActive = GetActiveWindow();

    id = MessageBox(hwndActive, pszText, pszTitle, mbFlags);

    return id;
}


#endif // DEBUG
