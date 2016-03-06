# Microsoft Developer Studio Project File - Name="Designer" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=Designer - Win32 Release
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "Designer.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Designer.mak" CFG="Designer - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Designer - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Designer - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "Designer - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir ".\Release"
# PROP BASE Intermediate_Dir ".\Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ".\Release"
# PROP Intermediate_Dir ".\Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /YX /c
# ADD CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "DESIGNER" /D "_DDINC_MODELESS_CODE" /Yu"precomp.h" /FD /c
# SUBTRACT CPP /Fr
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 libcmt.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib /nologo /base:"0x120000000" /entry:"_DllMainCRTStartup@12" /subsystem:windows /dll /machine:I386 /nodefaultlib
# SUBTRACT LINK32 /pdb:none /debug
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=copy .\release\designer.dll ..\..\setup\distrib
# End Special Build Tool

!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir ".\Debug"
# PROP BASE Intermediate_Dir ".\Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir ".\Debug"
# PROP Intermediate_Dir ".\Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /YX /c
# ADD CPP /nologo /MTd /W3 /GX /ZI /Od /D "DESIGNER" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_DDINC_MODELESS_CODE" /FR /Yu"precomp.h" /FD /c
# SUBTRACT CPP /X
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib /nologo /subsystem:windows /dll /profile /debug /machine:I386

!ENDIF 

# Begin Target

# Name "Designer - Win32 Release"
# Name "Designer - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;hpj;bat;for;f90"
# Begin Source File

SOURCE=.\AB10Format.cpp
# End Source File
# Begin Source File

SOURCE=.\APPPREF.CPP
# End Source File
# Begin Source File

SOURCE=.\Browser.cpp
# End Source File
# Begin Source File

SOURCE=.\CategoryMgr.cpp
# End Source File
# Begin Source File

SOURCE=.\catidreg.cpp
# End Source File
# Begin Source File

SOURCE=.\debug.cpp
# End Source File
# Begin Source File

SOURCE=.\Designer.cpp
# End Source File
# Begin Source File

SOURCE=.\DesignerPage.cpp
# End Source File
# Begin Source File

SOURCE=.\Dialogs.cpp
# End Source File
# Begin Source File

SOURCE=..\Dib.cpp
# End Source File
# Begin Source File

SOURCE=.\dibstuff.cpp
# End Source File
# Begin Source File

SOURCE=.\DragDrop.cpp
# End Source File
# Begin Source File

SOURCE=.\DragDropMgr.cpp
# End Source File
# Begin Source File

SOURCE=.\enumx.cpp
# End Source File
# Begin Source File

SOURCE=..\EventLog.cpp
# End Source File
# Begin Source File

SOURCE=.\FDialog.cpp
# End Source File
# Begin Source File

SOURCE=..\fontholder.cpp
# End Source File
# Begin Source File

SOURCE=.\fregkey.cpp
# End Source File
# Begin Source File

SOURCE=.\FWND.CPP
# End Source File
# Begin Source File

SOURCE=.\GdiUtil.cpp
# End Source File
# Begin Source File

SOURCE=.\Globals.cpp
# End Source File
# Begin Source File

SOURCE=.\guids.cpp
# End Source File
# Begin Source File

SOURCE=.\Iconedit.cpp
# End Source File
# Begin Source File

SOURCE=.\ImageMgr.cpp
# End Source File
# Begin Source File

SOURCE=.\ipserver.cpp
# End Source File
# Begin Source File

SOURCE=.\Map.cpp
# End Source File
# Begin Source File

SOURCE=..\pictureholder.cpp
# End Source File
# Begin Source File

SOURCE=.\precomp.cpp
# ADD CPP /Yc"precomp.h"
# End Source File
# Begin Source File

SOURCE=.\prj.def
# End Source File
# Begin Source File

SOURCE=.\prj.odl
# ADD MTL /tlb ".\prj.tlb" /h ".\DesignerInterfaces.h"
# End Source File
# Begin Source File

SOURCE=.\prj.rc
# End Source File
# Begin Source File

SOURCE=.\proppage.cpp
# End Source File
# Begin Source File

SOURCE=..\StaticLink.cpp
# End Source File
# Begin Source File

SOURCE=.\Stream.cpp
# End Source File
# Begin Source File

SOURCE=.\support.cpp
# End Source File
# Begin Source File

SOURCE=.\Trees.cpp
# End Source File
# Begin Source File

SOURCE=.\Utility.cpp
# End Source File
# Begin Source File

SOURCE=.\xevents.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl;fi;fd"
# Begin Source File

SOURCE=.\AB10Format.h
# End Source File
# Begin Source File

SOURCE=.\APPPREF.H
# End Source File
# Begin Source File

SOURCE=.\Browser.h
# End Source File
# Begin Source File

SOURCE=.\CategoryMgr.h
# End Source File
# Begin Source File

SOURCE=.\catidreg.h
# End Source File
# Begin Source File

SOURCE=.\ChildWnds.h
# End Source File
# Begin Source File

SOURCE=.\debug.h
# End Source File
# Begin Source File

SOURCE=.\Designer.h
# End Source File
# Begin Source File

SOURCE=.\DesignerInterfaces.h
# End Source File
# Begin Source File

SOURCE=.\DesignerPage.h
# End Source File
# Begin Source File

SOURCE=.\Dialogs.h
# End Source File
# Begin Source File

SOURCE=.\dispids.h
# End Source File
# Begin Source File

SOURCE=.\Dispids2.h
# End Source File
# Begin Source File

SOURCE=.\DragDrop.h
# End Source File
# Begin Source File

SOURCE=.\DragDropMgr.h
# End Source File
# Begin Source File

SOURCE=.\enumx.h
# End Source File
# Begin Source File

SOURCE=..\EventLog.h
# End Source File
# Begin Source File

SOURCE=.\Fdialog.h
# End Source File
# Begin Source File

SOURCE=..\fontholder.h
# End Source File
# Begin Source File

SOURCE=.\fregkey.h
# End Source File
# Begin Source File

SOURCE=.\Fwnd.h
# End Source File
# Begin Source File

SOURCE=.\GdiUtil.h
# End Source File
# Begin Source File

SOURCE=.\Globals.h
# End Source File
# Begin Source File

SOURCE=.\Iconedit.h
# End Source File
# Begin Source File

SOURCE=.\ImageMgr.h
# End Source File
# Begin Source File

SOURCE=..\interfaces.h
# End Source File
# Begin Source File

SOURCE=.\ipserver.h
# End Source File
# Begin Source File

SOURCE=.\mainwin.h
# End Source File
# Begin Source File

SOURCE=.\Map.h
# End Source File
# Begin Source File

SOURCE=.\precomp.h
# End Source File
# Begin Source File

SOURCE=.\proppage.h
# End Source File
# Begin Source File

SOURCE=.\resource.h
# End Source File
# Begin Source File

SOURCE=.\Stream.h
# End Source File
# Begin Source File

SOURCE=.\support.h
# End Source File
# Begin Source File

SOURCE=.\Trees.h
# End Source File
# Begin Source File

SOURCE=.\Utility.h
# End Source File
# Begin Source File

SOURCE=.\xevents.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;cnt;rtf;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\RES\bitmap1.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\bmp00001.bmp
# End Source File
# Begin Source File

SOURCE=.\RES\bmp00002.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\category.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\cur00001.cur
# End Source File
# Begin Source File

SOURCE=.\RES\cur00002.cur
# End Source File
# Begin Source File

SOURCE=.\RES\cur00003.cur
# End Source File
# Begin Source File

SOURCE=.\RES\cur00004.cur
# End Source File
# Begin Source File

SOURCE=.\res\cursor1.cur
# End Source File
# Begin Source File

SOURCE=.\RES\CYCLE.BMP
# End Source File
# Begin Source File

SOURCE=.\res\filedeme.bmp
# End Source File
# Begin Source File

SOURCE=.\RES\HATCHBMP.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\heirarch.bmp
# End Source File
# Begin Source File

SOURCE=.\RES\icon1.ico
# End Source File
# Begin Source File

SOURCE=.\res\iconedit.bmp
# End Source File
# Begin Source File

SOURCE=.\res\idcc_buc.cur
# End Source File
# Begin Source File

SOURCE=.\res\idcc_dra.cur
# End Source File
# Begin Source File

SOURCE=.\res\idcc_hai.cur
# End Source File
# Begin Source File

SOURCE=.\res\idcc_mar.cur
# End Source File
# Begin Source File

SOURCE=.\res\idcc_pen.cur
# End Source File
# Begin Source File

SOURCE=.\res\idcc_pip.cur
# End Source File
# Begin Source File

SOURCE=.\res\idcc_siz.cur
# End Source File
# Begin Source File

SOURCE=.\res\idcc_spr.cur
# End Source File
# Begin Source File

SOURCE=.\RES\lumi.cur
# End Source File
# Begin Source File

SOURCE=.\RES\mainwin.ico
# End Source File
# Begin Source File

SOURCE=.\RES\menugrab.cur
# End Source File
# Begin Source File

SOURCE=.\res\mngcate.bmp
# End Source File
# Begin Source File

SOURCE=.\RES\Spec.dib
# End Source File
# Begin Source File

SOURCE=.\RES\splith.cur
# End Source File
# Begin Source File

SOURCE=.\res\stateand.bmp
# End Source File
# Begin Source File

SOURCE=.\res\toolbar1.bmp
# End Source File
# End Group
# Begin Source File

SOURCE=.\prj.tlb
# End Source File
# End Target
# End Project
