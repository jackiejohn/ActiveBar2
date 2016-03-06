//=--------------------------------------------------------------------------=
// Debug.H
//=--------------------------------------------------------------------------=
// Copyright 1995-1996 Microsoft Corporation.  All Rights Reserved.
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
//=--------------------------------------------------------------------------=
//
// contains the various macros and the like which are only useful in DEBUG
// builds
//
#ifndef _DEBUG_H_

//=---------------------------------------------------------------------------=
// all the things required to handle our ASSERT mechanism
//=---------------------------------------------------------------------------=
//


#ifdef _DEBUG

struct DumpContext
{
	DumpContext();
	~DumpContext();

	void Write ();
	void Write (LPCTSTR szText);

	TCHAR  m_szBuffer[512];

private:
	HANDLE m_hDebug;
	DWORD  m_dwResult;
};

inline void DumpContext::Write()
{
	WriteFile(m_hDebug, m_szBuffer, _tcslen(m_szBuffer), &m_dwResult, 0);
}

inline void DumpContext::Write(LPCTSTR szText)
{
	WriteFile(m_hDebug, szText, _tcslen(szText), &m_dwResult, 0);
}

DumpContext& GetDumpContext();
#endif

#ifdef _DEBUG
#define TRACELEVEL 2
#endif

#undef ASSERT
#undef TRACE
#undef TRACE1
#undef TRACE2

#ifdef _DEBUG
	VOID DisplayAssert(LPSTR pszMsg, LPSTR pszAssert, LPSTR pszFile, UINT line);
	#define SZTHISFILE	static char _szThisFile[] = __FILE__;
	#define ASSERT_ONNULL(x) ASSERT(x,"Invalid NULL pointer")
	#define ASSERT(fTest, szMsg)                                \
		if (!(fTest))  {                                        \
			static char szMsgCode[] = szMsg;                    \
			static char szAssert[] = #fTest;                    \
			DisplayAssert(szMsgCode, szAssert, THIS_FILE, __LINE__); \
		}
	#define FAIL(szMsg)                                         \
        { static char szMsgCode[] = szMsg;                    \
        DisplayAssert(szMsgCode, "FAIL", THIS_FILE, __LINE__); }
#else

	#define SZTHISFILE
	#define ASSERT(fTest, err) ;
	#define ASSERT_ONNULL(x) ;
	#define FAIL(err) ;
#endif


//------------------------------
#ifdef _DEBUG
	#define TRACET(x) x
	#ifdef UNICODE
		#define TRACE(level,szMsg) {if (level>=TRACELEVEL) OutputDebugString(L#szMsg);}
		#define TRACE1(level,szMsg,x) {TCHAR temp[255]; if (level>=TRACELEVEL)  {wsprintf(temp,L#szMsg,x); TRACE(level,temp);}}
		#define TRACE2(level,szMsg,x1,x2) {TCHAR temp[255]; if (level>=TRACELEVEL) {wsprintf(temp,L#szMsg,x1,x2); TRACE(level,temp);}}
	#else
		#define TRACE(level,szMsg) {if (level>=TRACELEVEL) OutputDebugString(szMsg);}
		#define TRACE1(level,szMsg,x) {if (level>=TRACELEVEL) {char temp[255]; wsprintf(temp,szMsg,x); TRACE(level,temp);}}
		#define TRACE2(level,szMsg,x1,x2) {if (level>=TRACELEVEL) {char temp[255]; wsprintf(temp,szMsg,x1,x2); TRACE(level,temp);}}
	#endif

	#define CHECK_POINTER(val) if (!(val) || IsBadWritePtr((void *)(val), sizeof(void *))) return E_POINTER
#else
	#define TRACE(level,szMsg)
	#define TRACE1(level,szMsg,x)
	#define TRACE2(level,szMsg,x1,x2)
	#define CHECK_POINTER(val)
#endif


#define _DEBUG_H_

#endif // _DEBUG_H_
