//
//
// File: EventLog.cpp	
//
// Environment: Win32
//
// Author: Chris Longo
//
// Revision History:
//	1.0 Chris Longo Initial Release 1/01/96
//
//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "EventLog.h"

extern HINSTANCE g_hInstance;

//
// SEException
//

SEException::SEException(UINT nSE, _EXCEPTION_POINTERS* pExp) 
	: m_nSE(nSE),
	  m_pExp(pExp)
{
}

LPCTSTR SEException::What()
{ 
	int nIndex = 0; 
	struct SystemExceptions
	{
		DWORD dwId;
		TCHAR* szDesc;
	} exceptions [] = { {EXCEPTION_ACCESS_VIOLATION, _T("Access Violation")},
						{EXCEPTION_ARRAY_BOUNDS_EXCEEDED, _T("Array Bounds Exceeded")},
						{EXCEPTION_BREAKPOINT, _T("Break Point")},
						{EXCEPTION_DATATYPE_MISALIGNMENT, _T("Datatype Misalignment")}, 
						{EXCEPTION_FLT_DENORMAL_OPERAND, _T("Float Denormal Operand")}, 
						{EXCEPTION_FLT_DIVIDE_BY_ZERO, _T("Float Divide by zero")},
						{EXCEPTION_FLT_INEXACT_RESULT, _T("Float Inexact Result")},
						{EXCEPTION_FLT_INVALID_OPERATION, _T("Float Invalid Operation")}, 
						{EXCEPTION_FLT_OVERFLOW, _T("Float Overflow")},
						{EXCEPTION_FLT_STACK_CHECK, _T("Float Stack Check")}, 
						{EXCEPTION_FLT_UNDERFLOW, _T("Float underflow")},
						{EXCEPTION_ILLEGAL_INSTRUCTION, _T("Illegal Instruction")},
						{EXCEPTION_IN_PAGE_ERROR, _T("Page Error")},
						{EXCEPTION_INT_DIVIDE_BY_ZERO, _T("Integer Divide by Zero")},
						{EXCEPTION_INT_OVERFLOW, _T("Integer Overflow")},
						{EXCEPTION_INVALID_DISPOSITION, _T("Invalid Disposition")},
						{EXCEPTION_NONCONTINUABLE_EXCEPTION, _T("Non Continuable Exception")},
						{EXCEPTION_PRIV_INSTRUCTION, _T("Privage Instruction")},
						{EXCEPTION_SINGLE_STEP, _T("Single Step")},
						{EXCEPTION_STACK_OVERFLOW, _T("Stack Overflow")}};
	const int nSize = sizeof(exceptions) / sizeof(SystemExceptions);
	for (nIndex = 0; nIndex < nSize; nIndex++)
	{
		if (exceptions[nIndex].dwId == m_pExp->ExceptionRecord->ExceptionCode)
			break;
	}
	if (nIndex != nSize)
	{
		_sntprintf(m_szBuffer, 
				   eBufferSize,
				   _T("%s, Exception occured at address: %X"), 
				   exceptions[nIndex].szDesc, 
				   m_pExp->ExceptionRecord->ExceptionAddress);
		return m_szBuffer;
	}
	return _T("An undefined exception occurred");
}

//
// ReportException
//

void SEException::ReportException(LPCTSTR szFile, int nLine)
{
	EventLog* pEventLog = GetGlobals().GetEventLog();
	if (pEventLog)
	{
		TCHAR szBuffer[512];
		_sntprintf(szBuffer, 512, _T("Exception: %s, File: %s, Line: %i"), What(), szFile, nLine);
		pEventLog->WriteCustom(szBuffer, EVENTLOG_ERROR_TYPE);
		return;
	}
	assert(FALSE);
}

/*
 * EventLog::EventLog
 *
 * Functional Description:
 *		The constructor opens the Windows NT System Resgistry and trys to create a couple of
 *		registry keys.
 *
 * Formal Parameters:
 *
 *		szSourceName - Application name that is using the event log
 *
 *      szMessageDLL - the file name of the message DLL 
 *
 *		bClearOnExit - clear the registry entries when the class goes out of scope or is deleted.
 *
 * Implicit Parameters:
 *	
 *      m_bClearReg - set this flag to clear the registry entries when the class goes out of scope.
 *
 *      m_szSource - member variable that retains the Application name
 *
 *      m_szRegKey - member variable that keeps the registry key name for an instance of the class
 *
 * Return Value:
 *   	
 *      None
 *  
 * Side Effects:
 *
 *   	None
 */

LPCTSTR EventLog::m_szEventApp = _T("SYSTEM\\CurrentControlSet\\Services\\EventLog\\Application\\");
LPCTSTR EventLog::m_szLogEvent = _T("Software\\Data Dynamics\\ActiveBar\\2.0");

EventLog::EventLog(LPCTSTR szSourceName, 
                   LPCTSTR szMessageDLL,
			       BOOL    bWinNT,
                   BOOL    bClearOnExit)
{
	HKEY  hKey = NULL;
	DWORD dwType = REG_DWORD;
	DWORD dwData = 0;
	DWORD dwSize = sizeof(DWORD);
    m_bClearReg = bClearOnExit;
	m_bWinNT = bWinNT;
    
	long lResult = RegOpenKeyEx(HKEY_CURRENT_USER, 
							    m_szLogEvent,
								0,
							    KEY_QUERY_VALUE, 
							    &hKey);
	if (ERROR_SUCCESS != lResult) 
		return;

	lResult = RegQueryValueEx(hKey, 
							_T("Log Events"), 
							0,
							&dwType,
							(LPBYTE)&dwData, 
							&dwSize);
	RegCloseKey(hKey);

	if (ERROR_SUCCESS != lResult) 
		return;

	if (0 == dwData)
		return;

	if (m_bWinNT)
	{
		dwData = EVENTLOG_ERROR_TYPE   | 
				 EVENTLOG_WARNING_TYPE | 
				 EVENTLOG_INFORMATION_TYPE;

		lstrcpyn (m_szSource, szSourceName, eBufferSize);
		lstrcpyn (m_szRegKey, m_szEventApp, eBufferSize);
		lstrcat (m_szRegKey, szSourceName);

		//
		// Just try creating the registry key, if it exists, it will be opened
		//

		LRESULT lResult = RegCreateKeyEx(HKEY_LOCAL_MACHINE, 
										 m_szRegKey, 
										 0, 
										 _T("\0"),
										 REG_OPTION_NON_VOLATILE, 
										 KEY_ALL_ACCESS | KEY_WRITE, 
										 0, 
										 &hKey, 
										 &dwData);
		if (ERROR_SUCCESS != lResult) 
			throw ERROR_COULD_NOT_OPEN_REGISTRY_KEY;

		dwData = EVENTLOG_ERROR_TYPE   | 
				 EVENTLOG_WARNING_TYPE | 
				 EVENTLOG_INFORMATION_TYPE;

		lResult = RegSetValueEx(hKey, 
								_T("TypesSupported"), 
								0, 
								REG_DWORD, 
								(LPBYTE)&dwData, 
								sizeof(DWORD));
		if (ERROR_SUCCESS != lResult) 
		{
			lResult = RegCloseKey(hKey);
			throw ERROR_COULD_NOT_ADD_REGISTRY_VALUE;
		}
		lResult = RegCloseKey(hKey);
	}
	else
	{
		//
		// Windows 95/98
		//

		TCHAR* szEnd = NULL;
		DWORD dwResult = GetModuleFileName(g_hInstance, m_szSource, _MAX_PATH);
		if (dwResult > 0)
			szEnd = _tcsrchr(m_szSource, '\\');

		if (NULL != szEnd)
		{
			*(++szEnd) = '\0';
			lstrcat(m_szSource, _T("ActiveBarDesignerLog.rtf"));

			m_hFile = CreateFile(m_szSource, 
								 GENERIC_WRITE, 
								 FILE_SHARE_READ, 
								 NULL, 
								 CREATE_ALWAYS, 
								 FILE_ATTRIBUTE_NORMAL, 
								 NULL);
		}
		assert(m_hFile);
	}
} 

/*
 * EventLog::~EventLog
 *
 * Functional Description:
 *		This destructor does nothing if the m_fClearReg is set to FALSE.  
 * If this flag is set to TRUE the registry entries are removed.
 *
 * Formal Parameters:
 *
 *		None
 *
 * Implicit Parameters:
 *	
 *      None
 *
 * Return Value:
 *   	
 *      None
 *  
 * Side Effects:
 *
 *   	None
 */

EventLog::~EventLog()
{
	if (m_bWinNT)
	{
		HKEY hKey;
		long lResult;
		if (m_bClearReg) 
		{
			//         
			// Caller specified for registry entry to be removed
			//

			lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, 
								   m_szEventApp, 
								   0, 
								   KEY_WRITE, 
								   &hKey);
			if (ERROR_SUCCESS == lResult) 
			{
				lResult = RegDeleteKey(hKey, m_szSource); 
				lResult = RegCloseKey(hKey);
			} 
		} 
	}
	else
		CloseHandle(m_hFile);
}

/*
 * EventLog::WriteCustom
 *
 * Functional Description:
 *
 *		This member function will be used to log a custom event to the
 *	event log.
 *
 * Formal Parameters:
 *
 *		szString - a string that will be writen to the event log
 *
 *      wEventType - EVENTLOG_ERROR_TYPE	    - Error event
 *					 EVENTLOG_WARNING_TYPE	    - Warning event
 * 					 EVENTLOG_INFORMATION_TYPE	- Information event
 *					 EVENTLOG_AUDIT_SUCCESS	    - Success Audit event
 *					 EVENTLOG_AUDIT_FAILURE	    - Failure Audit event
 *
 * Implicit Parameters:
 *	
 *      None
 *
 * Return Value:
 *   	
 *      None
 *  
 * Side Effects:
 *
 *   	None
 */

const BOOL EventLog::WriteCustom(LPCTSTR szString, WORD wEventType)
{
	HKEY  hKey = NULL;
	DWORD dwType = REG_DWORD;
	DWORD dwData = 0;
	DWORD dwSize = sizeof(DWORD);
	long lResult = RegOpenKeyEx(HKEY_CURRENT_USER, 
							    m_szLogEvent,
								0,
							    KEY_QUERY_VALUE, 
							    &hKey);
	if (ERROR_SUCCESS != lResult) 
		return FALSE;

	lResult = RegQueryValueEx(hKey, 
							_T("Log Events"), 
							0,
							&dwType,
							(LPBYTE)&dwData, 
							&dwSize);

	RegCloseKey(hKey);
	if (ERROR_SUCCESS != lResult) 
		return FALSE;

	if (0 == dwData)
		return FALSE;

    BOOL bResult = FALSE;
	if (m_bWinNT)
	{
		HANDLE hSource = RegisterEventSource(NULL, m_szSource);
		if (NULL != hSource) 
		{
			BOOL bInternalResult = ReportEvent(hSource, 
											   wEventType, 
											   0, 
											   0, 
											   0, 
											   1, 
											   0, 
											   (LPCTSTR*)&szString, 
											   0);
			if (bInternalResult)
				bResult = TRUE;
			bInternalResult = DeregisterEventSource(hSource);
		}
	}
	else if (m_hFile)
	{

		DWORD dwResult;
		bResult = WriteFile(m_hFile, szString, lstrlen(szString), &dwResult, 0);
		bResult = WriteFile(m_hFile, _T("\n"), lstrlen(_T("\n")), &dwResult, 0);
		if (bResult)
			bResult = FlushFileBuffers(m_hFile);
	}
    return bResult;      
}

