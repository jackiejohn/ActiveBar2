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

	if (NULL != m_pPict)
		m_pPict->get_Type(&sPicType);
	return sPicType;
}

LPPICTUREDISP CPictureHolder::GetPictureDispatch()
{
	LPPICTUREDISP pPictDisp = NULL;
	if ((m_pPict != NULL) &&
		SUCCEEDED(m_pPict->QueryInterface(IID_IPictureDisp, (LPVOID*)&pPictDisp)))
	{
		assert(pPictDisp != NULL);
	}
	return pPictDisp;
}

void CPictureHolder::SetPictureDispatch(LPPICTUREDISP pDisp)
{
	if (m_pPict != 0)
		m_pPict->Release();

	LPPICTURE pPict = 0;
	if ((pDisp != 0) &&
		SUCCEEDED(pDisp->QueryInterface(IID_IPicture, (LPVOID*)&pPict)))
	{
		assert(pPict != 0);
		m_pPict = pPict;
	}
	else
		m_pPict = 0;
}

void CPictureHolder::Render(HDC hdc, const RECT& rcRender, const RECT& rcWBounds)
{
	if (m_pPict != 0)
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

void PersistPict(IStream *pStream, CPictureHolder *pictHolder,VARIANT_BOOL vbSave)
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
			return;
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
}




