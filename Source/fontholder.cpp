//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <olectl.h>
#include "fontholder.h"
#include <assert.h>
#include "debug.h"
#include "support.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


BOOL ConnectionAdvise(LPUNKNOWN pUnkSrc, REFIID iid,LPUNKNOWN pUnkSink, BOOL bRefCount, DWORD* pdwCookie)
{
	ASSERT_ONNULL(pUnkSrc);
	ASSERT_ONNULL(pUnkSink);
	ASSERT_ONNULL(pdwCookie);

	BOOL bSuccess = FALSE;
	LPCONNECTIONPOINTCONTAINER pCPC;
	if (SUCCEEDED(pUnkSrc->QueryInterface(IID_IConnectionPointContainer,(LPVOID*)&pCPC)))
	{
		ASSERT_ONNULL(pCPC);
		LPCONNECTIONPOINT pCP;
		if (SUCCEEDED(pCPC->FindConnectionPoint(iid, &pCP)))
		{
			ASSERT_ONNULL(pCP);
			if (SUCCEEDED(pCP->Advise(pUnkSink, pdwCookie)))
				bSuccess = TRUE;
			pCP->Release();
			if (bSuccess && !bRefCount)
				pUnkSink->Release();
		}
		pCPC->Release();
	}
	return bSuccess;
}

BOOL ConnectionUnadvise(LPUNKNOWN pUnkSrc, REFIID iid,LPUNKNOWN pUnkSink, BOOL bRefCount, DWORD dwCookie)
{
	ASSERT_ONNULL(pUnkSrc);
	ASSERT_ONNULL(pUnkSink);
	if (!bRefCount)
		pUnkSink->AddRef();
	BOOL bSuccess = FALSE;
	LPCONNECTIONPOINTCONTAINER pCPC;
	if (SUCCEEDED(pUnkSrc->QueryInterface(IID_IConnectionPointContainer,(LPVOID*)&pCPC)))
	{
		ASSERT_ONNULL(pCPC);
		LPCONNECTIONPOINT pCP;
		if (SUCCEEDED(pCPC->FindConnectionPoint(iid, &pCP)))
		{
			ASSERT_ONNULL(pCP);
			if (SUCCEEDED(pCP->Unadvise(dwCookie)))
				bSuccess = TRUE;
			pCP->Release();
		}
		pCPC->Release();
	}
	// If we failed, undo the earlier AddRef.
	if (!bRefCount && !bSuccess)
		pUnkSink->Release();
	return bSuccess;
}



CFontHolder::CFontHolder() :
	m_pFont(NULL),
	m_dwConnectCookie(0)

{
	m_pNotify=NULL;
}

void CFontHolder::SetPropertyNotifySink(LPPROPERTYNOTIFYSINK pNotify)
{
	m_pNotify=pNotify;
	//if (m_pNotify)
	//	m_pNotify->AddRef();
}

CFontHolder::~CFontHolder()
{
	//if (m_pNotify) 
	//	m_pNotify->Release();
	ReleaseFont();
}

void CFontHolder::ReleaseFont()
{
	if ((m_pFont != NULL) && (m_pNotify != NULL))
	{
		ConnectionUnadvise(m_pFont, IID_IPropertyNotifySink, m_pNotify,
			FALSE, m_dwConnectCookie);
	}

	if (m_pFont)
	{
		int testVal = m_pFont->Release();
		m_pFont=NULL;
	}
}

#ifdef JAPBUILD

static const FONTDESC _fdDefault =
{ sizeof(FONTDESC), (LPOLESTR)"‚l‚r ‚oƒSƒVƒbƒN\0\0", FONTSIZE(9), FW_NORMAL,
DEFAULT_CHARSET, FALSE, FALSE, FALSE };

#else

static const FONTDESC _fdDefault =
	{ sizeof(FONTDESC), OLESTR("MS Sans Serif"), FONTSIZE(8), FW_NORMAL,
	  DEFAULT_CHARSET, FALSE, FALSE, FALSE };

#endif


void CFontHolder::InitializeFont(const FONTDESC* pFontDesc,LPDISPATCH pFontDispAmbient)
{
#ifdef _DEBUG
	if (pFontDesc != NULL)
		ASSERT(pFontDesc->cbSizeofstruct == sizeof(FONTDESC),"INVALID FONTDESC");
#endif

	// Release any previous font, in preparation for creating a new one.
	ReleaseFont();

	LPFONT pFontAmbient;
	LPFONT pFontNew = NULL;

	if ((pFontDispAmbient != NULL) &&
		SUCCEEDED(pFontDispAmbient->QueryInterface(IID_IFont,(LPVOID*)&pFontAmbient)))
	{
		ASSERT_ONNULL(pFontAmbient);
		// Make a clone of the ambient font.
		pFontAmbient->Clone(&pFontNew);
		pFontAmbient->Release();
	}
	else
	{
		// Create the font.
		if (pFontDesc == NULL)
			pFontDesc = &_fdDefault;
		if (FAILED(OleCreateFontIndirect((LPFONTDESC)pFontDesc, IID_IFont,
				(LPVOID *)&pFontNew)))
			pFontNew = NULL;
	}
	// Setup advisory connection and find dispatch interface.
	if (pFontNew != NULL)
	{
		SetFont(pFontNew);
		pFontNew->Release();
	}
}

BOOL IsSameFont(CFontHolder& font, const FONTDESC* pFontDesc,LPFONTDISP pFontDispAmbient)
{
	if (font.m_pFont == NULL)
		return FALSE;
	BOOL bSame = FALSE;
	if (pFontDispAmbient != NULL)
	{
		LPFONT pFontAmbient;
		if (SUCCEEDED(pFontDispAmbient->QueryInterface(IID_IFont, (LPVOID*)&pFontAmbient)))
		{
			ASSERT_ONNULL(pFontAmbient);
			bSame = pFontAmbient->IsEqual(font.m_pFont) == S_OK;
			pFontAmbient->Release();
		}
	}
	else
	{
		if (pFontDesc == NULL)
			pFontDesc = &_fdDefault;
		bSame = TRUE;
		BOOL bFlag;
		font.m_pFont->get_Italic(&bFlag);
		bSame = (bFlag == pFontDesc->fItalic);
		if (bSame)
		{
			font.m_pFont->get_Underline(&bFlag);
			bSame = (bFlag == pFontDesc->fUnderline);
			if (bSame)
			{
				font.m_pFont->get_Strikethrough(&bFlag);
				bSame = (bFlag == pFontDesc->fStrikethrough);
			
				if (bSame)
				{
					short sCharset;
					font.m_pFont->get_Charset(&sCharset);
					bSame = (sCharset == pFontDesc->sCharset);
					if (bSame)
					{
						short sWeight;
						font.m_pFont->get_Weight(&sWeight);
						bSame = (sWeight == pFontDesc->sWeight);
						if (bSame)
						{
							CURRENCY cy;
							font.m_pFont->get_Size(&cy);
							bSame = (memcmp(&cy, &pFontDesc->cySize, sizeof(CURRENCY)) == 0);
							if (bSame)
							{
								BSTR bstrName;
								font.m_pFont->get_Name(&bstrName);
								bSame = lstrcmpW(bstrName,pFontDesc->lpstrName)==0;
								SysFreeString(bstrName);
							}
						}
					}
				}
			}
		}
	}
	return bSame;
}

HFONT CFontHolder::GetFontHandle()
{
	// Assume a screen DC for logical/himetric ratio.
	CacheScreenMetrics();
	return GetFontHandle(s_iYppli, HIMETRIC_PER_INCH);
}

HFONT CFontHolder::GetFontHandle(long cyLogical, long cyHimetric)
{
	HFONT hFont = NULL;

	if ((m_pFont != NULL) &&
		SUCCEEDED(m_pFont->SetRatio(cyLogical, cyHimetric)) &&
		SUCCEEDED(m_pFont->get_hFont(&hFont)))
	{
		ASSERT_ONNULL(hFont);
	}

	return hFont;
}

HFONT CFontHolder::Select(HDC hdc, long cyLogical, long cyHimetric)
{
	HFONT hFont = NULL;
	if (m_pFont != NULL)
		hFont = GetFontHandle(cyLogical, cyHimetric);
	if (hFont != NULL)
		return (HFONT)::SelectObject(hdc, hFont);
	return NULL;
}

void CFontHolder::QueryTextMetrics(LPTEXTMETRIC lptm)
{
	ASSERT_ONNULL(lptm);
	if (m_pFont != NULL)
	{
#if defined(_UNICODE) || defined(OLE2ANSI)
		m_pFont->QueryTextMetrics(lptm);
#else
		TEXTMETRICW tmw;
		m_pFont->QueryTextMetrics(&tmw);
		// convert metric to ansi
		memcpy(lptm, &tmw, sizeof(LONG) * 11);
		memcpy(&lptm->tmItalic, &tmw.tmItalic, sizeof(BYTE) * 5);
		WideCharToMultiByte(CP_ACP, 0, &tmw.tmFirstChar, 1, (LPSTR)&lptm->tmFirstChar, 1, NULL, NULL);
		WideCharToMultiByte(CP_ACP, 0, &tmw.tmLastChar, 1, (LPSTR)&lptm->tmLastChar, 1, NULL, NULL);
		WideCharToMultiByte(CP_ACP, 0, &tmw.tmDefaultChar, 1, (LPSTR)&lptm->tmDefaultChar, 1, NULL, NULL);
		WideCharToMultiByte(CP_ACP, 0, &tmw.tmBreakChar, 1, (LPSTR)&lptm->tmBreakChar, 1, NULL, NULL);
	
#endif
	}
	else
	{
		memset(lptm, 0, sizeof(TEXTMETRIC));
	}
}

LPFONTDISP CFontHolder::GetFontDispatch()
{
	LPFONTDISP pFontDisp = NULL;
	if ((m_pFont != NULL) &&
		SUCCEEDED(m_pFont->QueryInterface(IID_IFontDisp, (LPVOID*)&pFontDisp)))
	{
		ASSERT_ONNULL(pFontDisp);
	}
	return pFontDisp;
}

void CFontHolder::SetFont(LPFONT pFontNew)
{
	if (m_pFont != NULL)
		ReleaseFont();
	m_pFont = pFontNew;
	m_pFont->AddRef();
	if (m_pNotify != NULL)
	{
		ConnectionAdvise(m_pFont, IID_IPropertyNotifySink, m_pNotify,FALSE, &m_dwConnectCookie);
	}
}

void CFontHolder::SetFontDispatch(LPFONTDISP pNewFontDisp)
{
	if (pNewFontDisp==NULL)
		SetFont(NULL);
	IFont *m_pFont;
	if (SUCCEEDED(pNewFontDisp->QueryInterface(IID_IFont,(LPVOID *)&m_pFont)))
	{
		SetFont(m_pFont);
		m_pFont->Release();
		m_pFont=NULL;
	}

}

///////////////////// PERSISTANCE


HRESULT PersistBagFont(IPropertyBag *pPropBag,LPCOLESTR propName,IErrorLog *pErrorLog,CFontHolder *fontHolder,const FONTDESC *fontDesc,VARIANT_BOOL vbSave)
{
	HRESULT hr;
	VARIANT v;
	VariantInit(&v);
	
	if (VARIANT_TRUE == vbSave)
	{
		if (fontDesc!=NULL)
		{
			// compare fontdesc to current font if the same dont save just write out a NULL
		}
		VARIANT v;
		v.vt=VT_UNKNOWN;
		v.punkVal=fontHolder->m_pFont;
		v.punkVal->AddRef();
		hr=pPropBag->Write(propName,&v);
		VariantClear(&v);
	}
	else
	{
		VARIANT v;
		v.vt=VT_UNKNOWN;
		v.punkVal=NULL;
		hr=pPropBag->Read(propName,&v,pErrorLog);

		if (SUCCEEDED(hr))
		{
			IFont* m_pFont;
			hr=v.punkVal->QueryInterface(IID_IFont,(LPVOID *)&m_pFont);
			if (SUCCEEDED(hr))
			{
				fontHolder->SetFont(m_pFont);
				m_pFont->Release();
			}
		}
		VariantClear(&v);
	}
	return hr;
}
typedef struct tagddFontInfo
{
	WCHAR faceName[LF_FACESIZE];
	CY size;
	BOOL bBold,bItalic,bUnderline,bStrike;
	SHORT weight;
	short charset;
} ddFontInfo;

long CFontHolder::GetSize()
{
	if (m_pFont)
		return sizeof(LONG) + sizeof(ddFontInfo);
	else
		return sizeof(LONG);
}

HRESULT PersistFont(IStream* pStream, CFontHolder* pFontHolder, const FONTDESC* pFontDesc, VARIANT_BOOL vbSave)
{
	try
	{
		LONG sig=0;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			if (NULL == pFontHolder->m_pFont)
			{
				hResult = pStream->Write(&sig,sizeof(LONG),NULL);
				if (FAILED(hResult))
					return hResult;
				return NOERROR;
			}
			else
			{
				sig=2;
				hResult = pStream->Write(&sig,sizeof(LONG),NULL);
				if (FAILED(hResult))
					return hResult;

				IFont* pFont = pFontHolder->m_pFont;
				assert(pFont);

				ddFontInfo fontInfo;
				memset(&fontInfo,0,sizeof(fontInfo));

				BSTR bstrName;
				hResult = pFont->get_Name(&bstrName);
				if (FAILED(hResult))
					return hResult;

				wcscpy(fontInfo.faceName,bstrName);
				SysFreeString(bstrName);
				
				hResult = pFont->get_Size(&fontInfo.size);
				if (FAILED(hResult))
					return hResult;
				
				hResult = pFont->get_Bold(&fontInfo.bBold);
				if (FAILED(hResult))
					return hResult;
				
				hResult = pFont->get_Italic(&fontInfo.bItalic);
				if (FAILED(hResult))
					return hResult;
				
				hResult = pFont->get_Underline(&fontInfo.bUnderline);
				if (FAILED(hResult))
					return hResult;
				
				hResult = pFont->get_Strikethrough(&fontInfo.bStrike);
				if (FAILED(hResult))
					return hResult;
				
				hResult = pFont->get_Weight(&fontInfo.weight);
				if (FAILED(hResult))
					return hResult;
				
				hResult = pFont->get_Charset(&fontInfo.charset);
				if (FAILED(hResult))
					return hResult;
				
				hResult = pStream->Write(&fontInfo,sizeof(ddFontInfo),NULL);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
			hResult = pStream->Read(&sig,sizeof(LONG),NULL);
			if (FAILED(hResult))
				return hResult;

			pFontHolder->InitializeFont(pFontDesc);
			switch(sig)
			{
				case 1:
				{
					IPersistStream* pFontPersist;
					hResult = pFontHolder->m_pFont->QueryInterface(IID_IPersistStream,(LPVOID*)&pFontPersist);
					if (FAILED(hResult))
						return hResult;

					hResult = pFontPersist->Load(pStream);
					pFontPersist->Release();
					if (FAILED(hResult))
						return hResult;
				}
				break;
			
				case 2:
				{
					IFont* pFont = pFontHolder->m_pFont;
					assert(pFont);

					ddFontInfo fontInfo;
					
					hResult = pStream->Read(&fontInfo,sizeof(ddFontInfo),NULL);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pFont->put_Name(fontInfo.faceName);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pFont->put_Size(fontInfo.size);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pFont->put_Bold(fontInfo.bBold);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pFont->put_Italic(fontInfo.bItalic);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pFont->put_Underline(fontInfo.bUnderline);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pFont->put_Strikethrough(fontInfo.bStrike);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pFont->put_Weight(fontInfo.weight);
					if (FAILED(hResult))
						return hResult;

					hResult = pFont->put_Charset(fontInfo.charset);
					if (FAILED(hResult))
						return hResult;
				}
				break;

			default:
				break;
			}
		}
		return NOERROR;
	}
	catch (...)
	{
		assert(FALSE);
		return E_FAIL;
	}
}