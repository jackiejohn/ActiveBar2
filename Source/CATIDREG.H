#ifndef __CATIDREG_H__
#define __CATIDREG_H__

#include <comcat.h>

//{BEGIN CATDEFS}
//{END CATDEFS}

extern HRESULT RegisterCLSIDInCategory(REFCLSID clsid, CATID catid);
extern HRESULT UnRegisterCLSIDInCategory(REFCLSID clsid, CATID catid);

#endif
