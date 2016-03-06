#ifndef EVENTLOG_H
#define EVENTLOG_H

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

/*
*
* File:	EventLog.h	
*
* Class: SEException 
*
* Facility: Win32
*
* Abstract:  
*	This class converts Windows Structured Exceptions into C++ Exceptions that you 
* can handle in a C++ try catch block.
*
* To Initialize this class
*
*	void InitializeLibrary()
*			m_seFunction = _set_se_translator( trans_func );
*
* To Uninitialize this class
*
*	void UninitializeLibrary()
*			_set_se_translator(m_seFunction);
*
*	You may need to modify the ReportException method because it has a reference to a
* Global Object to retrieve a pointer to the EventLog Class
*
* Sample Usage:
*   try
*   {
*		statements
*	}
*	catch (SEException& e)
*	{
*		e.ReportException(__FILE__, __LINE__);
*	}
*
* Author: Chris Longo
*
* Revision History:
*	1.0 Chris Longo Initial Release 1/01/96
*
*/

//
// SEException 
//

class SEException 
{
public:
	enum Misc
	{
		eBufferSize = 256
	};

    SEException(UINT nSE, _EXCEPTION_POINTERS* pExp);

	virtual ~SEException() {}

    LPCTSTR What();
	BOOL Continue();	
	void ReportException(LPCTSTR szFile, int nLine);

private:    
	_EXCEPTION_POINTERS* m_pExp;
	TCHAR				 m_szBuffer[eBufferSize];
	UINT				 m_nSE;
};

inline BOOL SEException::Continue()
{
	if (0 == m_pExp->ExceptionRecord->ExceptionFlags)
		return TRUE;
	return FALSE;
}

inline void _cdecl trans_func (unsigned int nException, _EXCEPTION_POINTERS* pExp)
{
    throw SEException(nException,  pExp);
}

/*
*
* Copyright	(C) 1996 by DataDynamics
*
* File:	EventLog.h	
*
* Class: EventLog
*
* Facility: Win32
*
* Abstract:  
*		Much of this code was taken from an article in 
* the Windows/DOS Journal written by Paula Tomlinson.  
* I did a little hacking on her code to come up with this 
* implementation of a Windows NT Event Logging Class.   
*
* Environment: Windows NT
*
* Author: Chris Longo
*
* Revision History:
*	1.0 Chris Longo Initial Release 1/01/96
*
*/

class EventLog 
{
public:                                     
	  
	enum ERRORS 
	{
		ERROR_COULD_NOT_ADD_REGISTRY_KEY,
	    ERROR_COULD_NOT_ADD_REGISTRY_VALUE,
		ERROR_COULD_NOT_OPEN_REGISTRY_KEY
	};

	enum Misc
	{
		eBufferSize = 256
	};

    EventLog (LPCTSTR szAppName, 
              LPCTSTR szMessageDLL, 
		      BOOL    bWinNT,
			  BOOL    bClearLogOnExit = FALSE);

    virtual ~EventLog();

    const BOOL WriteCustom(LPCTSTR szString, WORD wEventType);
private:

	// Prevent Copying
    EventLog(const EventLog&);
	
	// Prevent Assignment
    const EventLog& operator=(const EventLog&);

	HANDLE m_hFile;
    TCHAR  m_szSource[eBufferSize];
    TCHAR  m_szRegKey[eBufferSize];
    BOOL   m_bClearReg;
	BOOL   m_bWinNT;
    static LPCTSTR m_szEventApp;
	static LPCTSTR m_szLogEvent;
};

//#ifdef _DEBUG
//#define CATCH catch (SEException& e)
//#define REPORTEXCEPTION(szFile, nLine) e.ReportException(szFile, nLine);
//#define CONTINUE if (!e.Continue()) throw;
//#else
#define CATCH catch (...)
#define REPORTEXCEPTION(szFile, nLine) TCHAR szBuffer[512]; \
				 					_sntprintf(szBuffer, 512, _T("Exception: File: %s, Line: %i"), szFile, nLine); \
									if (GetGlobals().GetEventLog()) \
										GetGlobals().GetEventLog()->WriteCustom(szBuffer, EVENTLOG_ERROR_TYPE);
#define CONTINUE throw;
//#endif

#endif
