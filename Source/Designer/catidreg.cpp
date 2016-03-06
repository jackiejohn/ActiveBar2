#include "precomp.h"
#include "catidreg.h"

const CLSID CLSID_StdComponentCategoriesMgr={0x2E005,0x0,0x0,{0xC0,0x0,0x0,0x0,0x0,0x0,0x0,0x46}};
const IID IID_ICatRegister={0x2E012,0x0,0x0,{0xC0,0x0,0x0,0x0,0x0,0x0,0x0,0x46}};

//{BEGIN CATDEFS}
//{END CATDEFS}

HRESULT RegisterCLSIDInCategory(REFCLSID clsid, CATID catid)
	{
// Register your component categories information.
    ICatRegister* pcr = NULL ;
    HRESULT hr = S_OK ;
    hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, 
			NULL, CLSCTX_INPROC_SERVER, IID_ICatRegister, (void**)&pcr);
    if (SUCCEEDED(hr))
    {
       // Register this category as being "implemented" by
       // the class.
       CATID rgcatid[1] ;
       rgcatid[0] = catid;
       hr = pcr->RegisterClassImplCategories(clsid, 1, rgcatid);
    }

    if (pcr != NULL)
        pcr->Release();
  
	return hr;
	}

// Helper function to unregister a CLSID as belonging to a component category
HRESULT UnRegisterCLSIDInCategory(REFCLSID clsid, CATID catid)
{
    ICatRegister* pcr = NULL ;
    HRESULT hr = S_OK ;
    hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, 
			NULL, CLSCTX_INPROC_SERVER, IID_ICatRegister, (void**)&pcr);
    if (SUCCEEDED(hr))
    {
       // Unregister this category as being "implemented" by
       // the class.
       CATID rgcatid[1] ;
       rgcatid[0] = catid;
       hr = pcr->UnRegisterClassImplCategories(clsid, 1, rgcatid);
    }

    if (pcr != NULL)
        pcr->Release();
  
	return hr;
}

