//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Utility.h"
#include "Resource.h"
#include "AppPref.h"

//
// FXRegKey
//

BOOL FXRegKey::GetValueDWORD(LPCTSTR szName, DWORD& dwValue)
{
	DWORD dwType = REG_DWORD;
	DWORD dwSize = sizeof(DWORD);
	lastError = RegQueryValueEx(hKey, szName, NULL, &dwType, (LPBYTE)&dwValue, &dwSize);
	if (dwType != REG_DWORD)
		return FALSE;
	return ERROR_SUCCESS == lastError;
}

BOOL FXRegKey::SetValueDWORD(LPCTSTR szName, const DWORD& dwValue)
{
	lastError = RegSetValueEx(hKey, szName, NULL, REG_DWORD, (BYTE*const)&dwValue, sizeof(DWORD));
	return ERROR_SUCCESS == lastError;
}

//
// CAppPref
//

CAppPref::CAppPref(LPCTSTR szCompanyName, LPCTSTR szAppName)
{
	FXRegKey rkSoftware;
	rkSoftware.Open(HKEY_CURRENT_USER, _T("Software"), TRUE);
	FXRegKey rkCompany;
	rkCompany.Open(rkSoftware, szCompanyName, TRUE);
	m_rkProduct.Open(rkCompany, szAppName, TRUE);
}

CAppPref::~CAppPref()
{
}

BOOL CAppPref::GetConfirmSuppress(UINT nValue)
{
	FXRegKey rkWindow;

	if (!rkWindow.Open(m_rkProduct, _T("Confirm"), TRUE))
		return FALSE;
	
	TCHAR szStr[20];
	wsprintf(szStr, _T("%d"), nValue);
	
	DWORD dwResult;
	if (!rkWindow.GetValueDWORD(szStr, dwResult))
		return FALSE;

	if (1 == dwResult)
		return TRUE;
	return FALSE;
}

void CAppPref::SetConfirmSuppress(UINT nValue, DWORD dwResult)
{
	FXRegKey rkWindow;
	if (!rkWindow.Open(m_rkProduct, _T("Confirm"), TRUE))
		return;

	TCHAR szStr[20];
	wsprintf(szStr,_T("%d"), nValue);

	rkWindow.SetValueDWORD(szStr, dwResult);
}

COLORREF CAppPref::GetColor(long nIndex)
{
	FXRegKey rkWindow;

	if (!rkWindow.Open(m_rkProduct, _T("Color"), TRUE))
		return RGB(255, 255, 255);
	
	TCHAR szStr[20];
	wsprintf(szStr, _T("%d"), nIndex);
	
	DWORD dwResult;
	if (!rkWindow.GetValueDWORD(szStr, dwResult))
		return RGB(255, 255, 255);

	return dwResult;
}

void CAppPref::SetColor(long nIndex, COLORREF crColor)
{
	FXRegKey rkWindow;
	if (!rkWindow.Open(m_rkProduct, _T("Color"), TRUE))
		return;

	TCHAR szStr[20];
	wsprintf(szStr,_T("%d"), nIndex);

	rkWindow.SetValueDWORD(szStr, crColor);
}

BOOL CAppPref::GetLibraryPath(LPTSTR szFile)
{
	FXRegKey rkWindow;

	if (!rkWindow.Open(m_rkProduct, _T("Library"), TRUE))
		return FALSE;
	
	if (!rkWindow.GetValue(_T("File"), szFile))
		return FALSE;

	return TRUE;
}

void CAppPref::SetLibraryPath(LPTSTR szFile)
{
	FXRegKey rkWindow;
	if (!rkWindow.Open(m_rkProduct, _T("Library"), TRUE))
		return;

	rkWindow.SetValue(_T("File"), szFile);
}

void CAppPref::SaveWinRect(int nId, HWND hWnd)
{
	CRect rc;
	GetWindowRect(hWnd, &rc);
	SaveWinRect(nId, rc);
}

void CAppPref::SaveWinRect(int nId, const CRect& rc)
{
	FXRegKey rkWindow;
	if (!rkWindow.Open(m_rkProduct, _T("Windows"), TRUE))
		return;

	TCHAR szId[20];
	wsprintf(szId,"%dLeft",nId);
	rkWindow.SetValueDWORD(szId,(DWORD)rc.left);
	wsprintf(szId,"%dTop",nId);
	rkWindow.SetValueDWORD(szId,(DWORD)rc.top);
	wsprintf(szId,"%dRight",nId);
	rkWindow.SetValueDWORD(szId,(DWORD)rc.right);
	wsprintf(szId,"%dBottom",nId);
	rkWindow.SetValueDWORD(szId,(DWORD)rc.bottom);
}

void CAppPref::LoadWinRect(int nId, CRect& rc, const CRect& rcDefault)
{
	FXRegKey rkWindow;
	if (!rkWindow.Open(m_rkProduct, _T("Windows"), TRUE))
		return;

	TCHAR szId[20];
	wsprintf(szId, _T("%dLeft"), nId);
	if (!rkWindow.GetValueDWORD(szId, (DWORD&)rc.left))
		goto LoadDefault;

	wsprintf(szId, _T("%dTop"), nId);
	rkWindow.GetValueDWORD(szId, (DWORD&)rc.top);
	
	wsprintf(szId, _T("%dRight") ,nId);
	rkWindow.GetValueDWORD(szId, (DWORD&)rc.right);
	
	wsprintf(szId,_T("%dBottom"), nId);
	rkWindow.GetValueDWORD(szId, (DWORD&)rc.bottom);
	return;

LoadDefault:
	rc = rcDefault;
}

BOOL CAppPref::LoadWinRect(int nId, CRect& rc)
{
	FXRegKey rkWindow;
	if (!rkWindow.Open(m_rkProduct, _T("Windows"), TRUE))
		return FALSE;

	TCHAR szId[20];
	wsprintf(szId, _T("%dLeft"), nId);
	if (!rkWindow.GetValueDWORD(szId, (DWORD&)rc.left))
		return FALSE;

	wsprintf(szId, _T("%dTop"), nId);
	rkWindow.GetValueDWORD(szId, (DWORD&)rc.top);
	
	wsprintf(szId, _T("%dRight") ,nId);
	rkWindow.GetValueDWORD(szId, (DWORD&)rc.right);
	
	wsprintf(szId,_T("%dBottom"), nId);
	rkWindow.GetValueDWORD(szId, (DWORD&)rc.bottom);
	return TRUE;
}

DWORD CAppPref::GetAutoId()
{
	FXRegKey rkWindow;
	if (!rkWindow.Open(m_rkProduct, _T("Library"), TRUE))
		return 0;

	DWORD dwAutoId;
	rkWindow.GetValueDWORD(_T("Auto Id"), dwAutoId);
	return dwAutoId;
}

void CAppPref::SetAutoId(DWORD dwAutoId)
{
	FXRegKey rkWindow;
	if (!rkWindow.Open(m_rkProduct, _T("Library"), TRUE))
		return;

	rkWindow.SetValueDWORD(_T("Auto Id"), dwAutoId);
}

