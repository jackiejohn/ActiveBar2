#ifdef _ABDLL
	#define DD_WNDCLASS(x) "ABDL"##x
#else
	#define DD_WNDCLASS(x) "ABS"##x
#endif
#include <windows.h>
#include <windowsx.h>
#include <olectl.h>
#include <commctrl.h>
#include <stdio.h>
#include <TCHAR.h>
#include <eh.h>
#include <assert.h>
#include "..\FontHolder.h"
#include "Globals.h"
#include "Utility.h"
#include "Map.h"
#include "FDialog.h"
#include "FWnd.h"
#include "GdiUtil.h"
