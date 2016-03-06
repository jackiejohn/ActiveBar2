//
// ScriptString.cpp : implementation of script string
//
// Copyright 1998-2000 Libronix Corp.
//

#include "PreComp.h"
#include "ScriptString.h"

#pragma warning( disable : 4127 )


class CScriptStringLibrary
{
// Construction
public:
	CScriptStringLibrary();

// Attributes
public:
	typedef HRESULT (WINAPI * ScriptStringAnalysePtr)(HDC hdc, const void *pString,
		int cString, int cGlyphs, int iCharset, DWORD dwFlags, int iReqWidth,
		SCRIPT_CONTROL * psControl, SCRIPT_STATE * psState,
		const int * piDx, SCRIPT_TABDEF * pTabdef, const BYTE * pbInClass,
		SCRIPT_STRING_ANALYSIS * pssa);
	typedef const SIZE * (WINAPI * ScriptString_pSizePtr)(
		SCRIPT_STRING_ANALYSIS ssa); 
	typedef const int * (WINAPI * ScriptString_pcOutCharsPtr)(
		SCRIPT_STRING_ANALYSIS ssa); 
	typedef HRESULT (WINAPI * ScriptStringCPtoXPtr)(
		SCRIPT_STRING_ANALYSIS ssa, int icp, BOOL fTrailing, int * pX);
	typedef HRESULT (WINAPI * ScriptStringXtoCPPtr)(
		SCRIPT_STRING_ANALYSIS ssa, int iX, int * piCh, int * piTrailing);
	typedef HRESULT (WINAPI * ScriptStringOutPtr)(
		SCRIPT_STRING_ANALYSIS ssa, int iX, int iY, UINT uOptions, 
		const RECT * prc, int iMinSel, int iMaxSel, BOOL fDisabled);
	typedef HRESULT (WINAPI * ScriptStringFreePtr)(
		SCRIPT_STRING_ANALYSIS * pssa);

	ScriptStringAnalysePtr m_pfnScriptStringAnalyse;
	ScriptString_pSizePtr m_pfnScriptString_pSize;
	ScriptString_pcOutCharsPtr m_pfnScriptString_pcOutChars;
	ScriptStringCPtoXPtr m_pfnScriptStringCPtoX;
	ScriptStringXtoCPPtr m_pfnScriptStringXtoCP;
	ScriptStringOutPtr m_pfnScriptStringOut;
	ScriptStringFreePtr m_pfnScriptStringFree;

// Operations
public:
	bool IsUniscribeEnabled();
	void EnableUniscribe(bool bEnable);

// Implementation
public:
	~CScriptStringLibrary();
protected:
	HINSTANCE m_hUniscribeLib;
	bool m_bUniscribeEnabled;
};


CScriptStringLibrary g_libUniscribe;


BOOL ScriptIsUniscribeEnabled()
{
	// check for Uniscribe
	return g_libUniscribe.IsUniscribeEnabled() ? TRUE : FALSE;
}


void ScriptEnableUniscribe(BOOL bEnable)
{
	// enable Uniscribe
	g_libUniscribe.EnableUniscribe(bEnable != FALSE);
}


struct SCRIPT_STRING_ANALYSIS_STUB
{
	HDC hdc;
	LPWSTR pszText;
	int nTextLen;
	SIZE sizeText;
	DWORD dwFlags;

	SCRIPT_STRING_ANALYSIS_STUB()
		{ hdc = NULL; pszText = NULL; nTextLen = 0; sizeText.cx = 0; sizeText.cy = 0; dwFlags = 0; }
	~SCRIPT_STRING_ANALYSIS_STUB()
		{ hdc = NULL; delete [] pszText; pszText = NULL; }
};


HRESULT WINAPI ScriptStringAnalyseStub(HDC hdc, const void *pString,
	int cString, int cGlyphs, int iCharset, DWORD dwFlags, int iReqWidth,
	SCRIPT_CONTROL * psControl, SCRIPT_STATE * psState,
	const int * piDx, SCRIPT_TABDEF * pTabdef, const BYTE * pbInClass,
	SCRIPT_STRING_ANALYSIS * pssa)
{
	// analysis
	SCRIPT_STRING_ANALYSIS_STUB * pAnalysis = NULL;
	
	try
	{
		// check arguments
		_com_util::CheckError(cGlyphs == 0 ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(iCharset == -1 ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(iReqWidth == 0 ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(psControl == NULL ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(psState == NULL ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(piDx == NULL ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(pTabdef == NULL ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(pbInClass == NULL ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(pssa != NULL ? S_OK : E_INVALIDARG);

		// create analysis
		pAnalysis = new SCRIPT_STRING_ANALYSIS_STUB;

		// store device context, flags
		pAnalysis->hdc = hdc;
		pAnalysis->dwFlags = dwFlags;

		// store string
		if (pString != NULL && cString > 0)
		{
			// store copy of string
			pAnalysis->pszText = new wchar_t [cString + 1];
			wcsncpy(pAnalysis->pszText, 
				reinterpret_cast<LPCWSTR>(pString), static_cast<size_t>(cString));
			pAnalysis->pszText[cString] = L'\0';
			pAnalysis->nTextLen = cString;
		}
		else
		{
			// store no string
			pAnalysis->pszText = NULL;
			pAnalysis->nTextLen = 0;
		}

		// check if processing ampersands (if so, we must use DrawText)
		if (pAnalysis->dwFlags & SSA_HOTKEY)
		{
			// determine rectangle
			RECT rectDrawText;
			rectDrawText.left = 0;
			rectDrawText.top = 0;
			rectDrawText.right = 0;
			rectDrawText.bottom = 0;

			// determine format
			UINT nFormat = DT_LEFT | DT_TOP | DT_SINGLELINE | DT_CALCRECT | DT_NOCLIP;

			// measure string (Win9x doesn't support DrawTextW)
			if (::GetVersion() < 0x80000000)
			{
				// Windows NT/2000
				if (::DrawTextW(pAnalysis->hdc, pAnalysis->pszText, pAnalysis->nTextLen, &rectDrawText, nFormat) == 0)
					_com_util::CheckError(HRESULT_FROM_WIN32(::GetLastError()));
			}
			else
			{
				// Windows 9x
				size_t nAnsiLength = pAnalysis->nTextLen * 2;
				char * pszAnsi = reinterpret_cast<char *>(malloc(nAnsiLength * sizeof(char)));
				if (pszAnsi == NULL)
					_com_util::CheckError(E_OUTOFMEMORY);
				pszAnsi[0] = '\0';
				nAnsiLength = ::WideCharToMultiByte(CP_ACP, 0, pAnalysis->pszText, pAnalysis->nTextLen, pszAnsi, nAnsiLength, NULL, NULL);
				if (::DrawText(pAnalysis->hdc, pszAnsi, nAnsiLength, &rectDrawText, nFormat) == 0)
				{
					HRESULT hrError = HRESULT_FROM_WIN32(::GetLastError());
					free(pszAnsi); pszAnsi = NULL;
					_com_util::CheckError(hrError);
				}
				free(pszAnsi);
			}

			// set size
			pAnalysis->sizeText.cx = rectDrawText.right - rectDrawText.left;
			pAnalysis->sizeText.cy = rectDrawText.bottom - rectDrawText.top;
		}
		else
		{
			// get text extent (GetTextExtentPoint32W works even on Win95)
			if (!::GetTextExtentPoint32W(pAnalysis->hdc, pAnalysis->pszText, pAnalysis->nTextLen, &pAnalysis->sizeText))
				_com_util::CheckError(HRESULT_FROM_WIN32(::GetLastError()));
		}

		// return analysis
		*pssa = pAnalysis;
		return S_OK;
	}
	catch (_com_error & e)
	{
		// return error
		delete pAnalysis;
		return e.Error();
	}
}


const SIZE * WINAPI ScriptString_pSizeStub(
	SCRIPT_STRING_ANALYSIS ssa)
{
	try
	{
		// check arguments
		_com_util::CheckError(ssa != NULL ? S_OK : E_INVALIDARG);

		// cast analysis
		SCRIPT_STRING_ANALYSIS_STUB * pAnalysis =
			reinterpret_cast<SCRIPT_STRING_ANALYSIS_STUB *>(ssa);

		// return pointer to size
		return &pAnalysis->sizeText;
	}
	catch (_com_error &)
	{
		// failure
		return NULL;
	}
}


const int * WINAPI ScriptString_pcOutCharsStub(
	SCRIPT_STRING_ANALYSIS ssa)
{
	try
	{
		// check arguments
		_com_util::CheckError(ssa != NULL ? S_OK : E_INVALIDARG);

		// SSA_CLIP not yet supported by analyze stub
		return NULL;
	}
	catch (_com_error &)
	{
		// failure
		return NULL;
	}
}


HRESULT WINAPI ScriptStringCPtoXStub(
	SCRIPT_STRING_ANALYSIS ssa, int icp, BOOL fTrailing, int * pX)
{
	try
	{
		// check arguments
		_com_util::CheckError(ssa != NULL ? S_OK : E_INVALIDARG);
		(void)icp; (void)fTrailing; (void)pX;

		// not implemented yet
		return E_NOTIMPL;
	}
	catch (_com_error & e)
	{
		// return error
		return e.Error();
	}
}


HRESULT WINAPI ScriptStringXtoCPStub(
	SCRIPT_STRING_ANALYSIS ssa, int iX, int * piCh, int * piTrailing)
{
	try
	{
		// check arguments
		_com_util::CheckError(ssa != NULL ? S_OK : E_INVALIDARG);
		(void)iX; (void)piCh; (void)piTrailing;

		// not implemented yet
		return E_NOTIMPL;
	}
	catch (_com_error & e)
	{
		// return error
		return e.Error();
	}
}


HRESULT WINAPI ScriptStringOutStub(
	SCRIPT_STRING_ANALYSIS ssa, int iX, int iY, UINT uOptions, 
	const RECT * prc, int iMinSel, int iMaxSel, BOOL fDisabled)
{
	try
	{
		// check arguments
		_com_util::CheckError(ssa != NULL ? S_OK : E_INVALIDARG);
		_com_util::CheckError(iMinSel == 0 ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(iMaxSel == 0 ? S_OK : E_UNEXPECTED);
		_com_util::CheckError(!fDisabled ? S_OK : E_UNEXPECTED);

		// cast analysis
		SCRIPT_STRING_ANALYSIS_STUB * pAnalysis =
			reinterpret_cast<SCRIPT_STRING_ANALYSIS_STUB *>(ssa);

		// check if processing ampersands (if so, we must use DrawText)
		if (pAnalysis->dwFlags & SSA_HOTKEY)
		{
			// determine rectangle
			RECT rectDrawText;
			rectDrawText.left = iX;
			rectDrawText.top = iY;
			rectDrawText.right = iX;
			rectDrawText.bottom = iY;

			// determine format
			UINT nFormat = DT_LEFT | DT_TOP | DT_SINGLELINE;
			if (uOptions & ETO_CLIPPED)
			{
				// clippped (it's not perfect, but it's the best we can do with DrawText)
				if (prc != NULL && prc->right > rectDrawText.right)
					rectDrawText.right = prc->right;
				if (prc != NULL && prc->bottom > rectDrawText.bottom)
					rectDrawText.bottom = prc->bottom;
			}
			else
			{
				// not clipped
				nFormat |= DT_NOCLIP;
			}

			// check if opaque
			if (uOptions & ETO_OPAQUE)
				::ExtTextOut(pAnalysis->hdc, iX, iY, uOptions, prc, _T(""), 0, NULL);

			// write string (Win9x doesn't support DrawTextW)
			if (::GetVersion() < 0x80000000)
			{
				// Windows NT/2000
				if (::DrawTextW(pAnalysis->hdc, pAnalysis->pszText, pAnalysis->nTextLen, &rectDrawText, nFormat) == 0)
					_com_util::CheckError(HRESULT_FROM_WIN32(::GetLastError()));
				return S_OK;
			}
			else
			{
				// Windows 9x
				size_t nAnsiLength = pAnalysis->nTextLen * 2;
				char * pszAnsi = reinterpret_cast<char *>(malloc(nAnsiLength * sizeof(char)));
				if (pszAnsi == NULL)
					_com_util::CheckError(E_OUTOFMEMORY);
				pszAnsi[0] = '\0';
				nAnsiLength = ::WideCharToMultiByte(CP_ACP, 0, pAnalysis->pszText, pAnalysis->nTextLen, pszAnsi, nAnsiLength, NULL, NULL);
				if (::DrawText(pAnalysis->hdc, pszAnsi, nAnsiLength, &rectDrawText, nFormat) == 0)
				{
					HRESULT hrError = HRESULT_FROM_WIN32(::GetLastError());
					free(pszAnsi); pszAnsi = NULL;
					_com_util::CheckError(hrError);
				}
				free(pszAnsi);
				return S_OK;
			}
		}
		else
		{
			// write string (ExtTextOutW works even on Win95)
			if (!::ExtTextOutW(pAnalysis->hdc, iX, iY, uOptions, prc, pAnalysis->pszText, static_cast<size_t>(pAnalysis->nTextLen), NULL))
				_com_util::CheckError(HRESULT_FROM_WIN32(::GetLastError()));
			return S_OK;
		}
	}
	catch (_com_error & e)
	{
		// return error
		return e.Error();
	}
}


HRESULT WINAPI ScriptStringFreeStub(
	SCRIPT_STRING_ANALYSIS * pssa)
{
	// check for structure
	if (pssa != NULL && *pssa != NULL)
	{
		// free analysis
		SCRIPT_STRING_ANALYSIS_STUB * pAnalysis =
			reinterpret_cast<SCRIPT_STRING_ANALYSIS_STUB *>(*pssa);
		delete pAnalysis;
		*pssa = NULL;
		return S_OK;
	}
	else
	{
		// nothing to free
		return S_FALSE;
	}
}


/***
@api ScriptTextOut
	Writes a Unicode string to the specified device, at the specified
	location, using the currently selected font, color, etc.

@param HDC hdc
	Handle to the device context.

@param int nX
	Specifies the logical x-coordinate used to align the string.

@param int nY
	Specifies the logical y-coordinate used to align the string.

@param const wchar_t * pszText
	Points to the string to be drawn.

@param int nTextLen
	Specifies the number of characters in the string pointed to by
	_pszText_. Unlike *TextOut*, this function allows -1 to be passed
	as the length of a null-terminated string.

@returns BOOL
	TRUE on success.
***/
BOOL ScriptTextOut(HDC hdc, int nX, int nY, 
	const wchar_t * pszText, int nTextLen)
{
	try
	{
		// check for empty string
		if (nTextLen == 0 || (nTextLen < 0 && *pszText == L'\0'))
		{
			// use Windows API for empty strings
			::TextOut(hdc, nX, nY, _T(""), 0);
		}
		else
		{
			// draw string
			CScriptString str;
			str.Analyze(hdc, pszText, nTextLen);
			str.Draw(nX, nY);
		}
		
		// success
		return TRUE;
	}
	catch (_com_error & e)
	{
		// return error
		::SetLastError(e.Error());
		return FALSE;
	}
}


/***
@api ScriptExtTextOut
	Writes a Unicode string to the specified device, at the specified
	location, using the currently selected font, color, etc. An optional
	rectangle allows clipping and/or opaquing.

@param HDC hdc
	Handle to the device context.

@param int nX
	Specifies the logical x-coordinate used to align the string.

@param int nY
	Specifies the logical y-coordinate used to align the string.

@param DWORD dwOptions
	See the documentation for *ExtTextOut*. The most common options
	are ETO_CLIPPED for clipping and ETO_OPAQUE for opaquing.

@param const RECT * prect
	Pointer to an optional rectangle used for clipping and/or opaquing.

@param const wchar_t * pszText
	Points to the string to be drawn.

@param int nTextLen
	Specifies the number of characters in the string pointed to by
	_pszText_. Unlike *ExtTextOut*, this function allows -1 to be passed
	as the length of a null-terminated string.

@param const int * pnSpacing
	Not supported at this time; should be NULL.

@returns BOOL
	TRUE on success.
***/
BOOL ScriptExtTextOut(HDC hdc, int nX, int nY, DWORD dwOptions,
	const RECT * prect, const wchar_t * pszText, int nTextLen, const int * pnSpacing)
{
	try
	{
		// check for empty string
		if (nTextLen == 0 || (nTextLen < 0 && *pszText == L'\0'))
		{
			// use Windows API for empty strings
			::ExtTextOut(hdc, nX, nY, dwOptions, prect, _T(""), 0, NULL);
		}
		else
		{
			// draw string
			CScriptString str;
			str.Analyze(hdc, pszText, nTextLen);
			str.Draw(nX, nY, dwOptions, prect);
		}
		
		// success
		return TRUE;
	}
	catch (_com_error & e)
	{
		// return error
		::SetLastError(e.Error());
		return FALSE;
	}
}


/***
@api ScriptDrawText
	Writes a Unicode string to the specified device, at the specified
	location, using the currently selected font, color, etc., in the
	specified format.

@param HDC hdc
	Handle to the device context.

@param const wchar_t * pszText
	Points to the string to be drawn.

@param int nTextLen
	Specifies the number of characters in the string pointed to by
	_pszText_. Unlike *DrawText*, this function allows -1 to be passed
	as the length of a null-terminated string.

@param RECT * prect
	Pointer to the rectangle used for formatting. If DT_CALCRECT is specified,
	this rectangle is modified (see *DrawText*).

@param UINT nFormat
	See the documentation for *DrawText*. DT_SINGLELINE must be specified, since
	line breaks are not supported at this time. The following values are also supported:
	  DT_BOTTOM, DT_CALCRECT, DT_CENTER, DT_LEFT, DT_NOCLIP, DT_NOPREFIX,
	  DT_RIGHT, DT_TOP, and DT_VCENTER.

@returns int
	The height of the text in logical units, or 0 on failure.
***/
int ScriptDrawText(HDC hdc, const wchar_t * pszText, int nTextLen,
	RECT * prect, UINT nFormat)
{
	try
	{
		// check for DT_SINGLELINE
		_com_util::CheckError((nFormat & DT_SINGLELINE) != 0 ? S_OK : E_INVALIDARG);

		// check for empty string
		if (nTextLen == 0 || (nTextLen < 0 && *pszText == L'\0'))
		{
			// use Windows API for empty strings
			return ::DrawText(hdc, _T(""), nTextLen, prect, nFormat);
		}
		else
		{
			// determine analyze flags
			DWORD dwFlags = SSA_FALLBACK | SSA_GLYPHS;
			if ((nFormat & DT_NOPREFIX) == 0)
				dwFlags |= SSA_HOTKEY;

			// measure string
			CScriptString str;
			str.Analyze(hdc, pszText, nTextLen, dwFlags);
			SIZE sizeText = str.GetSize();

			// determine horizontal position
			int nX = prect->left;
			if (nFormat & DT_RIGHT)
				nX = prect->right - sizeText.cx;
			else if (nFormat & DT_CENTER)
				nX = (prect->left + prect->right - sizeText.cx) / 2;

			// determine vertical position
			int nY = prect->top;
			if (nFormat & DT_BOTTOM)
				nY = prect->bottom - sizeText.cy;
			else if (nFormat & DT_VCENTER)
				nY = (prect->top + prect->bottom - sizeText.cy) / 2;

			// determine options
			DWORD dwOptions = 0;
			if ((nFormat & DT_NOCLIP) == 0)
				dwOptions |= ETO_CLIPPED;

			// check if only calculating rectangle
			if (nFormat & DT_CALCRECT)
			{
				// reproduce odd behavior of DrawText
				prect->right = prect->left + sizeText.cx;
				prect->bottom = nY + sizeText.cy;
			}
			else
			{
				// draw string
				str.Draw(nX, nY, dwOptions, dwOptions != 0 ? prect : NULL);
			}
		
			// success (reproduce odd behavior of DrawText)
			return nY + sizeText.cy - prect->top;
		}
	}
	catch (_com_error & e)
	{
		// return error
		::SetLastError(e.Error());
		return 0;
	}
}


/***
@api ScriptGetTextExtent
	Computes the width and height of the specified Unicode string of text.

@param HDC hdc
	Handle to the device context.

@param const wchar_t * pszText
	Points to the string to be measured.

@param int nTextLen
	Specifies the number of characters in the string pointed to by
	_pszText_. Unlike *GetTextExtentPoint32*, this function allows -1
	to be passed as the length of a null-terminated string.

@param SIZE * psize
	Pointer to the SIZE structure in which the dimensions of the string
	are to be returned.

@returns BOOL
	TRUE on success.
***/
BOOL ScriptGetTextExtent(HDC hdc,
	const wchar_t * pszText, int nTextLen, SIZE * psize)
{
	try
	{
		// draw string
		CScriptString str;
		str.Analyze(hdc, pszText, nTextLen);
		if (psize != NULL)
			*psize = str.GetSize();
		
		// success
		return TRUE;
	}
	catch (_com_error & e)
	{
		// return error
		::SetLastError(e.Error());
		return FALSE;
	}
}


/***
@class CScriptString
	Encapsulates a Unicode string analyzed for measuring, rendering, etc.

@remarks
	Use this class instead of the global functions when you want to
	efficiently measure and render the same Unicode string.
***/


/***
@mgroup Construction

@method CScriptString
	Constructs a *CScriptString* object.

@remarks
	The *Analyze* method must be called before the object can be used.
***/
CScriptString::CScriptString()
{
	// init members
	m_ssa = NULL;
}


/***
@method CScriptString
	Constructs a *CScriptString* object, analyzing the specified Unicode string.

@param HDC hdc
	Handle to the device context.

@param const wchar_t * pszText
	Points to the string to be analyzed.

@param int nTextLen
	Specifies the number of characters in the string pointed to by
	_pszText_. This function allows -1 to be passed as the length
	of a null-terminated string.

@throws _com_error
	Throws the same exceptions as *Analyze*.

@remarks
	This constructor calls the *Analyze* method with the specified
	arguments.
***/
CScriptString::CScriptString(HDC hdc, const wchar_t * pszText, int nTextLen) throw(_com_error)
{
	// init members
	m_ssa = NULL;

	// analyze string
	Analyze(hdc, pszText, nTextLen);
}


/***
@mgroup Attributes

@method GetSize [const]
	Returns the size of the analyzed text.

@returns SIZE
	The dimensions of the analyzed text.

@throws _com_error
	/ E_UNEXPECTED / No string has been analyzed. \
	/ E_FAIL / Unable to calculate size.
***/
SIZE CScriptString::GetSize() const throw(_com_error)
{
	// check for analysis
	if (m_ssa == NULL)
		_com_issue_error(E_UNEXPECTED);

	// get size
	const SIZE * psize = g_libUniscribe.m_pfnScriptString_pSize(m_ssa);
	if (psize == NULL)
		_com_issue_error(E_FAIL);
	return *psize;
}


/***
@method GetOutChars [const]
	Returns the number of analyzed characters.

@returns int
	The number of analyzed characters (reflects cropping).

@throws _com_error
	/ E_UNEXPECTED / No string has been analyzed. \
	/ E_FAIL / Unable to calculate number of characters.
***/
int CScriptString::GetOutChars() const throw(_com_error)
{
	// check for analysis
	if (m_ssa == NULL)
		_com_issue_error(E_UNEXPECTED);

	// get out characters
	const int * pnOutChars = g_libUniscribe.m_pfnScriptString_pcOutChars(m_ssa);
	if (pnOutChars == NULL)
		_com_issue_error(E_FAIL);
	return *pnOutChars;
}


/***
@method CPtoX [const]
	Returns the x-coordinate of the specified character position.

@param int nCharPos
	The character position to test.

@param bool bTrailing
	True to calculate the x-coordinate for the trailing case. (Huh?)

@returns int
	The x-coordinate of the specified character position.

@throws _com_error
	/ E_UNEXPECTED / No string has been analyzed. \
	/ other / Error returned by the Uniscribe library.
***/
int CScriptString::CPtoX(int nCharPos, bool bTrailing) const throw(_com_error)
{
	// check for analysis
	if (m_ssa == NULL)
		_com_issue_error(E_UNEXPECTED);

	// call function
	int nResult = 0;
	_com_util::CheckError(g_libUniscribe.m_pfnScriptStringCPtoX(m_ssa, nCharPos, bTrailing ? TRUE : FALSE, &nResult));
	return nResult;
}


/***
@method XtoCP [const]
	Returns the character position of the specified x-coordinate.

@param int nX
	The x-coordinate to test.

@param bool * pbTrailing
	If non-null, the pointed to boolean is set to true if the character position
	is trailing. (Huh?)

@returns int
	The character position of the specified x-coordinate.

@throws _com_error
	/ E_UNEXPECTED / No string has been analyzed. \
	/ other / Error returned by the Uniscribe library.
***/
int CScriptString::XtoCP(int nX, bool * pbTrailing) const throw(_com_error)
{
	// check for analysis
	if (m_ssa == NULL)
		_com_issue_error(E_UNEXPECTED);

	// call function
	int nResult = 0;
	int nTrailing = 0;
	_com_util::CheckError(
		g_libUniscribe.m_pfnScriptStringXtoCP(m_ssa, nX, &nResult, &nTrailing));
	if (pbTrailing != NULL)
		*pbTrailing = nTrailing != 0;
	return nResult;
}


/***
@method Draw [const]
	Writes the Unicode string to the previously specified device, at
	the specified location, using the currently selected font, color, etc.

@param int nX
	Specifies the logical x-coordinate used to align the string.

@param int nY
	Specifies the logical y-coordinate used to align the string.

@param DWORD dwOptions
	See the documentation for *ExtTextOut*. The most common options
	are ETO_CLIPPED for clipping and ETO_OPAQUE for opaquing.

@param const RECT * prect
	Pointer to an optional rectangle used for clipping and/or opaquing.

@throws _com_error
	/ E_UNEXPECTED / No string has been analyzed. \
	/ other / Error returned by the Uniscribe library.
***/
void CScriptString::Draw(int nX, int nY, DWORD dwOptions, const RECT * prect) const throw(_com_error)
{
	// check for analysis
	if (m_ssa == NULL)
		_com_issue_error(E_UNEXPECTED);

	// draw string
	_com_util::CheckError(g_libUniscribe.m_pfnScriptStringOut(
		m_ssa, nX, nY, dwOptions, prect, 0, 0, FALSE));
}


/***
@mgroup Operations

@method Analyze
	Analyzes the specified Unicode string in preparation for its measurement
	and/or rendering.

@param HDC hdc
	Handle to the device context.

@param const wchar_t * pszText
	Points to the string to be analyzed. The empty string is an invalid argument.

@param int nTextLen
	Specifies the number of characters in the string pointed to by
	_pszText_. This function allows -1 to be passed as the length
	of a null-terminated string.

@param DWORD dwFlags
	Specifies the flags passed to ScriptStringAnalyze. Uses SSA_FALLBACK | SSA_GLYPHS
	by default. Add SSA_HOTKEY to process ampersand prefixes.

@throws _com_error
	Throws any errors returned by the Uniscribe library.
***/
void CScriptString::Analyze(HDC hdc, const wchar_t * pszText, int nTextLen, DWORD dwFlags) throw(_com_error)
{
	try
	{
		// check for analysis
		if (m_ssa != NULL)
		{
			// free analysis
			_com_util::CheckError(g_libUniscribe.m_pfnScriptStringFree(&m_ssa));
		}

		// check string length
		if (pszText == NULL)
			nTextLen = 0;
		if (nTextLen < 0)
			nTextLen = static_cast<int>(wcslen(pszText));
		
		// analyze string
		_com_util::CheckError(g_libUniscribe.m_pfnScriptStringAnalyse(
			hdc, pszText, nTextLen, 0, -1, dwFlags,
			0, NULL, NULL, NULL, NULL, NULL, &m_ssa));
	}
	catch (_com_error &)
	{
		// clear members
		m_ssa = NULL;
		throw;
	}
}


/***
@method Free
	Prematurely frees the analysis of the previously analyzed string.

@throws _com_error
	Throws any errors returned by the Uniscribe library.

@remarks
	Since the destructor automatically frees the analysis of any previously
	analyzed string, this method is generally unnecessary.
***/
void CScriptString::Free() throw(_com_error)
{
	try
	{
		// check for analysis
		if (m_ssa != NULL)
		{
			// free analysis
			_com_util::CheckError(g_libUniscribe.m_pfnScriptStringFree(&m_ssa));
		}
	}
	catch (_com_error &)
	{
		// clear members
		m_ssa = NULL;
		throw;
	}
}


/***
@mgroup Implementation

@method ~CScriptString
	Destroys the *CScriptString* object.

@remarks
	The destructor automatically frees the analysis of any previously
	analyzed strings.
***/
CScriptString::~CScriptString()
{
	// check for analysis
	if (m_ssa != NULL)
	{
		// free analysis
		g_libUniscribe.m_pfnScriptStringFree(&m_ssa);
		m_ssa = NULL;
	}
}


CScriptStringLibrary::CScriptStringLibrary()
{
	// use Uniscribe by default (if available)
	EnableUniscribe(true);
}


bool CScriptStringLibrary::IsUniscribeEnabled()
{
	// return boolean
	return m_bUniscribeEnabled;
}


void CScriptStringLibrary::EnableUniscribe(bool bEnable)
{
	// set success
	bool bSuccess = false;
	m_bUniscribeEnabled = false;

	do
	{
		// don't load Uniscribe if not enabled
		if (!bEnable) break;

		// load Uniscribe library
		m_hUniscribeLib = ::LoadLibrary(_T("usp10.dll"));
		if (m_hUniscribeLib == NULL) break;

		// load functions
		m_pfnScriptStringAnalyse = reinterpret_cast<ScriptStringAnalysePtr>(
			::GetProcAddress(m_hUniscribeLib, "ScriptStringAnalyse"));
		if (m_pfnScriptStringAnalyse == NULL) break;
		m_pfnScriptString_pSize = reinterpret_cast<ScriptString_pSizePtr>(
			::GetProcAddress(m_hUniscribeLib, "ScriptString_pSize"));
		if (m_pfnScriptString_pSize == NULL) break;
		m_pfnScriptString_pcOutChars = reinterpret_cast<ScriptString_pcOutCharsPtr>(
			::GetProcAddress(m_hUniscribeLib, "ScriptString_pcOutChars"));
		if (m_pfnScriptString_pcOutChars == NULL) break;
		m_pfnScriptStringCPtoX = reinterpret_cast<ScriptStringCPtoXPtr>(
			::GetProcAddress(m_hUniscribeLib, "ScriptStringCPtoX"));
		if (m_pfnScriptStringCPtoX == NULL) break;
		m_pfnScriptStringXtoCP = reinterpret_cast<ScriptStringXtoCPPtr>(
			::GetProcAddress(m_hUniscribeLib, "ScriptStringXtoCP"));
		if (m_pfnScriptStringXtoCP == NULL) break;
		m_pfnScriptStringOut = reinterpret_cast<ScriptStringOutPtr>(
			::GetProcAddress(m_hUniscribeLib, "ScriptStringOut"));
		if (m_pfnScriptStringOut == NULL) break;
		m_pfnScriptStringFree = reinterpret_cast<ScriptStringFreePtr>(
			::GetProcAddress(m_hUniscribeLib, "ScriptStringFree"));
		if (m_pfnScriptStringFree == NULL) break;

		// success
		bSuccess = true;
		m_bUniscribeEnabled = true;
	}
	while (0);

	// check success
	if (!bSuccess)
	{
		// use stubs
		m_pfnScriptStringAnalyse = ScriptStringAnalyseStub;
		m_pfnScriptString_pSize = ScriptString_pSizeStub;
		m_pfnScriptString_pcOutChars = ScriptString_pcOutCharsStub;
		m_pfnScriptStringCPtoX = ScriptStringCPtoXStub;
		m_pfnScriptStringXtoCP = ScriptStringXtoCPStub;
		m_pfnScriptStringOut = ScriptStringOutStub;
		m_pfnScriptStringFree = ScriptStringFreeStub;
	}
}


CScriptStringLibrary::~CScriptStringLibrary()
{
	// unload Uniscribe library
	if (m_hUniscribeLib != NULL)
		::FreeLibrary(m_hUniscribeLib);
	m_hUniscribeLib = NULL;
	m_pfnScriptStringAnalyse = NULL;
	m_pfnScriptString_pSize = NULL;
	m_pfnScriptString_pcOutChars = NULL;
	m_pfnScriptStringCPtoX = NULL;
	m_pfnScriptStringXtoCP = NULL;
	m_pfnScriptStringOut = NULL;
	m_pfnScriptStringFree = NULL;
}
