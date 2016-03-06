//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <assert.h>
#include <olectl.h>
#include "fontholder.h"

extern HINSTANCE g_hInstance;
#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CPictureHolder::CPictureHolder()
	: m_pPict(NULL)
{
}

CPictureHolder::~CPictureHolder()
{
	if (m_pPict)
		m_pPict->Release();
}

BOOL CPictureHolder::CreateEmpty()
{
	if (m_pPict)
		m_pPict->Release();
	PICTDESC pdesc;
	pdesc.cbSizeofstruct = sizeof(pdesc);
	pdesc.picType = PICTYPE_NONE;
	return SUCCEEDED(OleCreatePictureIndirect(&pdesc, 
											  IID_IPicture, 
											  FALSE,
											  (LPVOID*)&m_pPict));

}

BOOL CPictureHolder::CreateFromBitmap(UINT idResource)
{
	HBITMAP hBmp=LoadBitmap(g_hInstance,MAKEINTRESOURCE(idResource));
	return CreateFromBitmap(hBmp, 0, TRUE);
}

BOOL CPictureHolder::CreateFromBitmap(HBITMAP hBitmap, HPALETTE hPalette, BOOL bTransferOwnership)
{
	if (m_pPict)
		m_pPict->Release();
	PICTDESC pdesc;
	pdesc.cbSizeofstruct = sizeof(pdesc);
	pdesc.picType = PICTYPE_BITMAP;
	pdesc.bmp.hbitmap = hBitmap;
	pdesc.bmp.hpal = hPalette;
	return SUCCEEDED(OleCreatePictureIndirect(&pdesc, 
											  IID_IPicture,
											  bTransferOwnership, 
											  (LPVOID*)&m_pPict));
}


BOOL CPictureHolder::CreateFromMetafile(HMETAFILE hmf, int xExt, int yExt, BOOL bTransferOwnership)
{
	if (m_pPict)
		m_pPict->Release();
	PICTDESC pdesc;
	pdesc.cbSizeofstruct = sizeof(pdesc);
	pdesc.picType = PICTYPE_METAFILE;
	pdesc.wmf.hmeta = hmf;
	pdesc.wmf.xExt = xExt;
	pdesc.wmf.yExt = yExt;
	return SUCCEEDED(OleCreatePictureIndirect(&pdesc, 
											  IID_IPicture,
											  bTransferOwnership, 
											  (LPVOID*)&m_pPict));
}

BOOL CPictureHolder::CreateFromIcon(UINT idResource)
{
	HICON hIcon = LoadIcon(g_hInstance,MAKEINTRESOURCE(idResource));
	return CreateFromIcon(hIcon, TRUE);
}

BOOL CPictureHolder::CreateFromIcon(HICON hIcon, BOOL bTransferOwnership)
{
	if (m_pPict)
		m_pPict->Release();
	PICTDESC pdesc;
	pdesc.cbSizeofstruct = sizeof(pdesc);
	pdesc.picType = PICTYPE_ICON;
	pdesc.icon.hicon = hIcon;
	return SUCCEEDED(OleCreatePictureIndirect(&pdesc, IID_IPicture,
		bTransferOwnership, (LPVOID*)&m_pPict));
}

short CPictureHolder::GetType()
{
	short sPicType = (short)PICTYPE_UNINITIALIZED;

	if (m_pPict)
		m_pPict->get_Type(&sPicType);
	return sPicType;
}

LPPICTUREDISP CPictureHolder::GetPictureDispatch()
{
	LPPICTUREDISP pPictDisp = NULL;
	if ((m_pPict) && SUCCEEDED(m_pPict->QueryInterface(IID_IPictureDisp, (LPVOID*)&pPictDisp)))
	{
		assert(pPictDisp);
	}
	return pPictDisp;
}

void CPictureHolder::SetPictureDispatch(LPPICTUREDISP pDisp)
{
	if (m_pPict)
		m_pPict->Release();

	LPPICTURE pPict = NULL;
	if (pDisp && SUCCEEDED(pDisp->QueryInterface(IID_IPicture, (LPVOID*)&pPict)))
	{
		assert(pPict);
		m_pPict = pPict;
	}
	else
		m_pPict = NULL;
}

void CPictureHolder::Render(HDC hdc, const RECT& rcRender, const RECT& rcWBounds)
{
	if (m_pPict)
	{
		long hmWidth;
		long hmHeight;

		m_pPict->get_Width(&hmWidth);
		m_pPict->get_Height(&hmHeight);

		m_pPict->Render(hdc, 
						rcRender.left, 
						rcRender.top,
						rcRender.right-rcRender.left, 
						rcRender.bottom-rcRender.top, 
						0, 
						hmHeight-1,
					    hmWidth, 
						-hmHeight,
						&rcWBounds);
	}
}

long CPictureHolder::GetSize()
{
	IPersistStream* pPictPersist;
	if (NULL == m_pPict || 
		PICTYPE_UNINITIALIZED == GetType() ||
		FAILED(m_pPict->QueryInterface(IID_IPersistStream, (LPVOID*)&pPictPersist)))
	{
		return sizeof(LONG);
	}
	else
	{
		IStream* pStream;
		HRESULT hResult = CreateStreamOnHGlobal(NULL, TRUE, &pStream);
		if (FAILED(hResult))
		{
			pPictPersist->Release();
			return sizeof(LONG);
		}

		hResult = pPictPersist->Save(pStream, FALSE);
		if (FAILED(hResult))
		{
			pStream->Release();
			pPictPersist->Release();
			return sizeof(LONG);
		}

		LARGE_INTEGER liMove;
		liMove.LowPart = liMove.HighPart = 0;
		ULARGE_INTEGER liEnd;
		hResult = pStream->Seek(liMove, STREAM_SEEK_END, &liEnd);
		pStream->Release();
		pPictPersist->Release();
		if (FAILED(hResult))
			return sizeof(LONG);
		return liEnd.LowPart + sizeof(LONG);
	}
}

HRESULT PersistPicture(IStream* pStream, CPictureHolder* pPictHolder, VARIANT_BOOL vbSave)
{
	LONG sig=0;
	HRESULT hResult;
	IPersistStream* pPersist;
	if (VARIANT_TRUE == vbSave)
	{
		if (NULL == pPictHolder->m_pPict)
		{
			hResult = pStream->Write(&sig, sizeof(LONG), NULL);
			return hResult;
		}
		
		if (NULL == pPictHolder->m_pPict)
		{
			hResult = pStream->Write(&sig, sizeof(LONG), NULL);
			return hResult;
		}

		hResult = pPictHolder->m_pPict->QueryInterface(IID_IPersistStream,(LPVOID*)&pPersist);
		if (FAILED(hResult))
		{
			hResult = pStream->Write(&sig, sizeof(LONG), NULL);
			return hResult;
		}
		else
		{
			sig = 1;
			short picType;
			hResult = pPictHolder->m_pPict->get_Type(&picType);
			if (picType != PICTYPE_METAFILE)
			{
				hResult = pStream->Write(&sig, sizeof(LONG), NULL);
				hResult = pPersist->Save(pStream,FALSE);
				pPersist->Release();
			}
			else
			{ 
				// Metafiles are causing leak in Picture object so save ourself
				pPersist->Release();

				sig=2;
				HMETAFILE hwmf;
				hResult = pStream->Write(&sig,sizeof(LONG),NULL);

				OLE_XSIZE_HIMETRIC w;
				OLE_YSIZE_HIMETRIC h;
				hResult = pPictHolder->m_pPict->get_Width(&w);
				hResult = pPictHolder->m_pPict->get_Height(&h);

				hResult = pPictHolder->m_pPict->get_Handle((OLE_HANDLE *)&hwmf);

				UINT metaFileSize = GetMetaFileBitsEx(hwmf,0,NULL);
				if (metaFileSize == 0)
					return E_FAIL;

				char* buffer = (char*)malloc(metaFileSize);
				if (NULL == buffer)
					return E_OUTOFMEMORY;

				if (GetMetaFileBitsEx(hwmf,metaFileSize,buffer)==0)
				{
					free(buffer);
					return E_FAIL;
				}

				hResult = pStream->Write(&metaFileSize,sizeof(UINT),NULL);
				if (FAILED(hResult))
					return hResult;

				hResult = pStream->Write(buffer,metaFileSize,NULL);
				if (FAILED(hResult))
					return hResult;

				hResult = pStream->Write(&w,sizeof(OLE_XSIZE_HIMETRIC),NULL);
				hResult = pStream->Write(&h,sizeof(OLE_YSIZE_HIMETRIC),NULL);

				free(buffer);
				return NOERROR;
			}
		}
	}
	else
	{
		if (NULL != pPictHolder->m_pPict)
		{
			pPictHolder->m_pPict->Release();
			pPictHolder->m_pPict = NULL;
		}

		hResult = pStream->Read(&sig, sizeof(LONG), NULL);
		if (sig!=0)
		{
			if (sig==1)
			{
				hResult = OleLoadPicture(pStream,0,FALSE,IID_IPicture,(LPVOID*)&(pPictHolder->m_pPict));
				if (FAILED(hResult))
				{
					return hResult;
				}
			}
			if (sig==2) // Windows metafile
			{
				UINT metaFileSize;
				HMETAFILE hwmf;
				OLE_XSIZE_HIMETRIC w;
				OLE_YSIZE_HIMETRIC h;

				hResult = pStream->Read(&metaFileSize,sizeof(UINT),NULL);
				LPBYTE buffer=(LPBYTE)malloc(metaFileSize);
				if (NULL == buffer)
					return E_OUTOFMEMORY;
				
				hResult = pStream->Read(buffer,metaFileSize,NULL);
				if (FAILED(hResult))
					return hResult;

				hResult = pStream->Read(&w,sizeof(OLE_XSIZE_HIMETRIC),NULL);
				hResult = pStream->Read(&h,sizeof(OLE_YSIZE_HIMETRIC),NULL);

				hwmf = SetMetaFileBitsEx(metaFileSize,buffer);

				hResult = pPictHolder->CreateFromMetafile(hwmf,w,h,TRUE);

				free(buffer);

				return hResult;
			}
		}
	}
	return NOERROR;
}
/*
HRESULT PersistPict(IStream *pStream, CPictureHolder *pictHolder,VARIANT_BOOL vbSave)
{
	LONG sig=0;
	HRESULT hr;
	IPersistStream *pPictPersist;
	if (VARIANT_TRUE == vbSave)
	{
		if (0 == pictHolder->m_pPict || 
			pictHolder->GetType()==PICTYPE_UNINITIALIZED ||
			FAILED(pictHolder->m_pPict->QueryInterface(IID_IPersistStream,(LPVOID*)&pPictPersist)))
		{
			pStream->Write(&sig,sizeof(LONG),NULL);
			return NOERROR;
		}
		else
		{
			sig=1;
			pStream->Write(&sig,sizeof(LONG),NULL);
			hr=pPictPersist->Save(pStream,FALSE);
			pPictPersist->Release();
		}
	}
	else
	{
		pStream->Read(&sig,sizeof(LONG),NULL);
		if (pictHolder->m_pPict) // clear
		{
			pictHolder->m_pPict->Release();
			pictHolder->m_pPict=0;
		}
		if (sig==1)
		{
			pictHolder->CreateEmpty();
			if (pictHolder->m_pPict != 0)
			{
				hr=pictHolder->m_pPict->QueryInterface(IID_IPersistStream,(LPVOID *)&pPictPersist);
				if (SUCCEEDED(hr))
				{
					pPictPersist->Load(pStream);
					pPictPersist->Release();
				}
			}
		}
	}
	return NOERROR;
}

HRESULT PersistPict(IStream* pStream, CPictureHolder* pPictHolder, VARIANT_BOOL vbSave)
{
	try
	{
		LONG sig = 0;
		HRESULT hResult;
		IPersistStream* pPictPersist;
		if (VARIANT_TRUE == vbSave)
		{
			if (NULL == pPictHolder->m_pPict || 
				PICTYPE_UNINITIALIZED == pPictHolder->GetType() ||
				FAILED(pPictHolder->m_pPict->QueryInterface(IID_IPersistStream, (LPVOID*)&pPictPersist)))
			{
				hResult = pStream->Write(&sig, sizeof(LONG), NULL);
				if (FAILED(hResult))
					return hResult;
				
				return NOERROR;
			}
			else
			{
				sig = 1;
				hResult = pStream->Write(&sig, sizeof(LONG), NULL);
				if (FAILED(hResult))
					return hResult;

				hResult = pPictPersist->Save(pStream, FALSE);
				if (FAILED(hResult))
					return hResult;
				
				pPictPersist->Release();
			}
		}
		else
		{
			hResult = pStream->Read(&sig,sizeof(LONG),NULL);
			if (FAILED(hResult))
				return hResult;

			if (pPictHolder->m_pPict) // clear
			{
				pPictHolder->m_pPict->Release();
				pPictHolder->m_pPict = NULL;
			}
			
			if (1 == sig)
			{
				pPictHolder->CreateEmpty();
				if (pPictHolder->m_pPict)
				{
					hResult = pPictHolder->m_pPict->QueryInterface(IID_IPersistStream, (LPVOID*)&pPictPersist);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pPictPersist->Load(pStream);
					pPictPersist->Release();
					
					if (FAILED(hResult))
						return hResult;
				}
			}
		}
		return NOERROR;
	}
	catch (...)
	{
		return E_FAIL;
	}
}

*/


