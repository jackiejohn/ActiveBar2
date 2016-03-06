#ifndef DRAGDROPMANAGER_INCLUDED
#define DRAGDROPMANAGER_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "..\Interfaces.h"
#include "..\PrivateInterfaces.h"

//
// Drag Drop Manager
//

class CDragDropMgr : IDragDropManager
{
public:
	STDMETHOD(RegisterDragDrop)(OLE_HANDLE hWnd, LPUNKNOWN pDropTarget);
	STDMETHOD(RevokeDragDrop)(OLE_HANDLE hWnd);

	STDMETHOD(DoDragDrop)(LPUNKNOWN pDataObject, 
					      LPUNKNOWN pDropSource, 
					      DWORD	    dwOKEffect, 
					      DWORD*	pdwEffect);
private:
	TypedMap<OLE_HANDLE, LPUNKNOWN> m_mapTargets;
};

#endif