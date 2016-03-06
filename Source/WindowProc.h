#ifndef WINDOWPROC_INCLUDED
#define WINDOWPROC_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

LRESULT CALLBACK FrameWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK MDIClientWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK IEWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK IEClientWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK FormWindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK KeyBoardHook(int nCode, WPARAM wParam, LPARAM lParam);
#endif