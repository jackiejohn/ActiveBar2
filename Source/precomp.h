#include <windows.h>
#ifdef _ABDLL
	#define DD_WNDCLASS(x) "ABDL"##x
#else
	#define DD_WNDCLASS(x) "ABS"##x
#endif

#include <WINDOWSX.H>
#include <comdef.h>
#include <process.h>    
#define OEMRESOURCE
#include <WINUSER.H>
#include <commctrl.h>
#include <olectl.h>
#include <stdio.h>
#include <tchar.h>
#include <stdio.h>
#include <stddef.h>
#include <eh.h>
#include <assert.h>
#include "Debug.h"
#include "EventLog.h"
#include "Utility.h"
#include "Globals.h"
#include "FDialog.h"
#include "Map.h"
#include "FWnd.h"
#include "FontHolder.h"
#include "Interfaces.h"
