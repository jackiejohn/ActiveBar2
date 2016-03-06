//
// ScriptString.h : interface for script string
//
// Copyright 1998-2000 Libronix Corp.
//

#ifndef __SCRIPTSTRING_H__
#define __SCRIPTSTRING_H__

#pragma once

#include <windows.h>
#include <tchar.h>

#include <usp10.h>

#include <comdef.h>
#include <comutil.h>

#pragma warning( disable : 4290 )


BOOL ScriptIsUniscribeEnabled();

void ScriptEnableUniscribe(BOOL bEnable);


BOOL ScriptTextOut(HDC hdc, int nX, int nY, 
	const wchar_t * pszText, int nTextLen);

BOOL ScriptExtTextOut(HDC hdc, int nX, int nY, DWORD dwOptions,
	const RECT * prect, const wchar_t * pszText, int nTextLen, const int * pnSpacing = NULL);

BOOL ScriptDrawText(HDC hdc, const wchar_t * pszText, int nTextLen,
	RECT * prect, UINT nFormat);

BOOL ScriptGetTextExtent(HDC hdc,
	const wchar_t * pszText, int nTextLen, SIZE * psize);


class CScriptString
{
// Construction
public:
	CScriptString();
	CScriptString(HDC hdc, const wchar_t * pszText, int nTextLen = -1) throw(_com_error);

// Attributes
public:
	SIZE GetSize() const throw(_com_error);
	int GetOutChars() const throw(_com_error);

	void Draw(int nX, int nY, 
		DWORD dwOptions = 0, const RECT * prect = NULL) const throw(_com_error);

	int CPtoX(int nCharPos, bool bTrailing) const throw(_com_error);
	int XtoCP(int nX, bool * pbTrailing) const throw(_com_error);

// Operations
public:
	void Analyze(HDC hdc, const wchar_t * pszText, int nTextLen = -1,
		DWORD dwFlags = SSA_FALLBACK | SSA_GLYPHS) throw(_com_error);
	void Free() throw(_com_error);

// Implementation
public:
	~CScriptString();
private:
	CScriptString(const CScriptString & str);
	CScriptString & operator =(const CScriptString & str);

// Implementation Data
protected:
	SCRIPT_STRING_ANALYSIS m_ssa;
};


#endif
