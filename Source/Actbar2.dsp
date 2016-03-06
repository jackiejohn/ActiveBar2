# Microsoft Developer Studio Project File - Name="Actbar" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=Actbar - Win32 AB DLL Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "Actbar2.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Actbar2.mak" CFG="Actbar - Win32 AB DLL Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Actbar - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Actbar - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Actbar - Win32 AB DLL Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Actbar - Win32 AB DLL Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Actbar - Win32 Japanese AB DLL Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/ActBar2.0/Source", ETEAAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "Actbar - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir ".\Release"
# PROP BASE Intermediate_Dir ".\Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ".\Release"
# PROP Intermediate_Dir ".\Release"
# PROP Ignore_Export_Lib 1
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /YX /c
# ADD CPP /nologo /MT /W3 /GX /O2 /Op /Oy- /Ob2 /I ".\\" /I "\ddlib" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "STRICT" /D "_MT" /D "_SCRIPTSTRING" /Fr /Yu"precomp.h" /FD /c
# SUBTRACT CPP /X
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 libcmt.lib kernel32.lib user32.lib gdi32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib comdlg32.lib comsupp.lib /nologo /base:"0x35000000" /entry:"_DllMainCRTStartup@12" /subsystem:windows /dll /machine:I386 /nodefaultlib /out:".\Release/Actbar2.ocx" /opt:nowin98
# SUBTRACT LINK32 /pdb:none /map /debug
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=copy .\release\actbar2.ocx ..\setup\distrib	cd ..\setup\distrib	prepall	cd ..\..\source
# End Special Build Tool

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

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
# ADD CPP /nologo /MTd /W3 /GR /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_SCRIPTSTRING" /FR /Yu"precomp.h" /D /D /c
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 libcmtd.lib kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib /nologo /base:"0x35000000" /subsystem:windows /dll /profile /debug /machine:I386 /out:".\Debug/Actbar2.ocx"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Actbar___Win32_AB_DLL_Release"
# PROP BASE Intermediate_Dir "Actbar___Win32_AB_DLL_Release"
# PROP BASE Ignore_Export_Lib 1
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "AB_DLL_Release"
# PROP Intermediate_Dir "AB_DLL_Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /I ".\\" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "STRICT" /D "_MT" /Fr /Yu"precomp.h" /FD /c
# SUBTRACT BASE CPP /X
# ADD CPP /nologo /MT /W3 /GX /Zi /O2 /I ".\\" /I "\ddlib" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "STRICT" /D "_MT" /D "_DDINC_MODELESS_CODE" /D "_ABDLL" /D "AR20" /D "___ABDLL" /Fr /Yu"precomp.h" /FD /c
# SUBTRACT CPP /X
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib LIBCMT.LIB /nologo /base:"0x35000000" /entry:"_DllMainCRTStartup@12" /subsystem:windows /dll /machine:I386 /nodefaultlib /out:".\Release/Actbar2.ocx" /opt:nowin98
# SUBTRACT BASE LINK32 /pdb:none /map /debug
# ADD LINK32 libcmt.lib kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib /nologo /base:"0x35000000" /entry:"_DllMainCRTStartup@12" /subsystem:windows /dll /machine:I386 /nodefaultlib /out:"AB_DLL_Release\AB2DLL.dll" /opt:nowin98
# SUBTRACT LINK32 /pdb:none /map /debug
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=copy .\release\actbar2.ocx ..\setup\distrib	cd ..\setup\distrib	prepall	cd ..\..\source
# End Special Build Tool

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Actbar___Win32_AB_DLL_Debug"
# PROP BASE Intermediate_Dir "Actbar___Win32_AB_DLL_Debug"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Actbar___Win32_AB_DLL_Debug"
# PROP Intermediate_Dir "Actbar___Win32_AB_DLL_Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /GR /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_DDINC_MODELESS_CODE" /FR /Yu"precomp.h" /D /D /c
# ADD CPP /nologo /MTd /W3 /GR /GX /ZI /Od /I ".\\" /I "\ddlib" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_DDINC_MODELESS_CODE" /D "_ABDLL" /D "AR20" /FR /Yu"precomp.h" /D /D /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 libcmtd.lib kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib /nologo /base:"0x11000000" /subsystem:windows /dll /profile /debug /machine:I386 /nodefaultlib /out:".\Debug/Actbar2.ocx"
# ADD LINK32 libcmtd.lib kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib /nologo /base:"0x11000000" /subsystem:windows /dll /profile /debug /machine:I386 /nodefaultlib /out:".\Actbar___Win32_AB_DLL_Debug/AB2DLL.dll"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Actbar___Win32_Japanese_AB_DLL_Release"
# PROP BASE Intermediate_Dir "Actbar___Win32_Japanese_AB_DLL_Release"
# PROP BASE Ignore_Export_Lib 1
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Japanese_AB_DLL_Release"
# PROP Intermediate_Dir "Japanese_AB_DLL_Release"
# PROP Ignore_Export_Lib 1
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /Zi /O2 /I ".\\" /I "\ddlib" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "STRICT" /D "_MT" /D "_DDINC_MODELESS_CODE" /D "_ABDLL" /D "AR20" /Yu"precomp.h" /FD /c
# SUBTRACT BASE CPP /X /Fr
# ADD CPP /nologo /MT /W3 /GX /Zi /O2 /I ".\\" /I "\ddlib" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "STRICT" /D "_MT" /D "_DDINC_MODELESS_CODE" /D "_ABDLL" /D "AR20" /D "JAPBUILD" /Yu"precomp.h" /FD /c
# SUBTRACT CPP /X /Fr
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 libcmt.lib kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib /nologo /base:"0x35000000" /entry:"_DllMainCRTStartup@12" /subsystem:windows /dll /machine:I386 /nodefaultlib /out:"AB_DLL_Release\AB2DLL.dll" /opt:nowin98
# SUBTRACT BASE LINK32 /pdb:none /map /debug
# ADD LINK32 libcmt.lib kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib /nologo /base:"0x35000000" /entry:"_DllMainCRTStartup@12" /subsystem:windows /dll /machine:I386 /nodefaultlib /out:"Japanese_AB_DLL_Release\AB2DLL.dll" /opt:nowin98
# SUBTRACT LINK32 /pdb:none /map /debug
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=copy .\release\actbar2.ocx ..\setup\distrib	cd ..\setup\distrib	prepall	cd ..\..\source
# End Special Build Tool

!ENDIF 

# Begin Target

# Name "Actbar - Win32 Release"
# Name "Actbar - Win32 Debug"
# Name "Actbar - Win32 AB DLL Release"
# Name "Actbar - Win32 AB DLL Debug"
# Name "Actbar - Win32 Japanese AB DLL Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;hpj;bat;for;f90"
# Begin Source File

SOURCE=.\ABDLLPRJ.DEF

!IF  "$(CFG)" == "Actbar - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=".\ActiveBar 2 work items.xls"
# End Source File
# Begin Source File

SOURCE=.\band.cpp
DEP_CPP_BAND_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDrop.h"\
	".\Designer\DragDropMgr.h"\
	".\Dispids.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Bands.cpp
DEP_CPP_BANDS=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Bar.cpp
DEP_CPP_BAR_C=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\CBList.h"\
	".\ChildBands.h"\
	".\Custom.h"\
	".\CustomizeListbox.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DesignerInterfaces.h"\
	".\Designer\DragDropMgr.h"\
	".\Dib.h"\
	".\Dispids.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\GUIDS.H"\
	".\hlp\ActiveBar2Ref.hh"\
	".\ImageMgr.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\ShortCuts.h"\
	".\StaticLink.h"\
	".\Streams.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Tooltip.h"\
	".\Tppopup.h"\
	".\Utility.h"\
	".\WindowProc.h"\
	".\XEVENTS.H"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Btabs.cpp
DEP_CPP_BTABS=\
	".\Bar.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\ScriptString.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	{$(INCLUDE)}"usp10.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\CategoryMgr.cpp
DEP_CPP_CATEG=\
	".\CategoryMgr.h"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\catidreg.cpp
DEP_CPP_CATID=\
	".\CATIDREG.H"\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\CBList.cpp
DEP_CPP_CBLIS=\
	".\Bar.h"\
	".\CBList.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ChildBands.cpp
DEP_CPP_CHILD=\
	".\Band.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Dib.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Custom.cpp
DEP_CPP_CUSTO=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\Custom.h"\
	".\CustomizeListbox.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\ShortCuts.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\CustomizeListbox.cpp
DEP_CPP_CUSTOM=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\CategoryMgr.h"\
	".\CustomizeListbox.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDrop.h"\
	".\Designer\DragDropMgr.h"\
	".\Dispids.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	".\XEVENTS.H"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\customproxy.cpp
DEP_CPP_CUSTOMP=\
	".\customproxy.h"\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\GUIDS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\Utility.h"\
	".\XEVENTS.H"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DDASYNC.CPP
DEP_CPP_DDASY=\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Streams.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\debug.cpp
DEP_CPP_DEBUG=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Dib.cpp
DEP_CPP_DIB_C=\
	".\DEBUG.H"\
	".\Dib.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Dock.cpp
DEP_CPP_DOCK_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DesignerInterfaces.h"\
	".\Designer\DragDrop.h"\
	".\Designer\DragDropMgr.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FREGKEY.H"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Designer\DragDrop.cpp
DEP_CPP_DRAGD=\
	".\Designer\AppPref.h"\
	".\Designer\Debug.h"\
	".\Designer\DragDrop.h"\
	".\Designer\FDialog.h"\
	".\Designer\FRegkey.h"\
	".\Designer\FWnd.h"\
	".\Designer\GdiUtil.h"\
	".\Designer\Globals.h"\
	".\Designer\Map.h"\
	".\Designer\PreComp.h"\
	".\Designer\Utility.h"\
	".\EventLog.h"\
	".\fontholder.h"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\PrivateInterfaces.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Designer\DragDropMgr.cpp
DEP_CPP_DRAGDR=\
	".\Designer\AppPref.h"\
	".\Designer\Debug.h"\
	".\Designer\DragDropMgr.h"\
	".\Designer\FDialog.h"\
	".\Designer\FRegkey.h"\
	".\Designer\FWnd.h"\
	".\Designer\GdiUtil.h"\
	".\Designer\Globals.h"\
	".\Designer\Map.h"\
	".\Designer\PreComp.h"\
	".\Designer\Utility.h"\
	".\EventLog.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\PrivateInterfaces.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DropSource.cpp
DEP_CPP_DROPS=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDrop.h"\
	".\Designer\DragDropMgr.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\enumx.cpp
DEP_CPP_ENUMX=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\EventLog.cpp
DEP_CPP_EVENT=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\FDIALOG.CPP
DEP_CPP_FDIAL=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Flicker.cpp
DEP_CPP_FLICK=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\fontholder.cpp
DEP_CPP_FONTH=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\fregkey.cpp
DEP_CPP_FREGK=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FREGKEY.H"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\FWND.CPP
DEP_CPP_FWND_=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Globals.cpp
DEP_CPP_GLOBA=\
	".\Bar.h"\
	".\Btabs.h"\
	".\Custom.h"\
	".\CustomizeListbox.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDropMgr.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\ShortCuts.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	".\WindowProc.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\guids.cpp
DEP_CPP_GUIDS=\
	".\DEBUG.H"\
	".\Designer\DesignerInterfaces.h"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\GUIDS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Ibar.cpp
DEP_CPP_IBAR_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Custom.h"\
	".\CustomizeListbox.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDropMgr.h"\
	".\Dispids.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\ImageMgr.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\StaticLink.h"\
	".\Streams.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ImageMgr.cpp
DEP_CPP_IMAGE=\
	".\Bar.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Dib.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\ImageMgr.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ipserver.cpp
DEP_CPP_IPSER=\
	".\CATIDREG.H"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FREGKEY.H"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\localizer.cpp
DEP_CPP_LOCAL=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\localizer.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Map.cpp
DEP_CPP_MAP_C=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Miniwin.cpp
DEP_CPP_MINIW=\
	".\Band.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DesignerInterfaces.h"\
	".\Designer\DragDrop.h"\
	".\Designer\DragDropMgr.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\pictureholder.cpp
DEP_CPP_PICTU=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Popupwin.cpp
DEP_CPP_POPUP=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DesignerInterfaces.h"\
	".\Designer\DragDrop.h"\
	".\Designer\DragDropMgr.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Tppopup.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\precomp.cpp
DEP_CPP_PRECO=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

# ADD CPP /Yc"precomp.h"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

# ADD CPP /Yc"precomp.h"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# ADD BASE CPP /Yc"precomp.h"
# ADD CPP /Yc"precomp.h"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# ADD BASE CPP /Yc"precomp.h"
# ADD CPP /Yc"precomp.h"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# ADD BASE CPP /Yc"precomp.h"
# ADD CPP /Yc"precomp.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Private.cpp
DEP_CPP_PRIVA=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DesignerInterfaces.h"\
	".\Designer\DragDropMgr.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\ImageMgr.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	".\WindowProc.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Private.odl
# ADD MTL /h ".\PrivateInterfaces.h"
# End Source File
# Begin Source File

SOURCE=.\PRJ.DEF

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\prj.odl
# ADD MTL /tlb ".\prj.tlb" /h ".\interfaces.h"
# End Source File
# Begin Source File

SOURCE=.\prj.rc

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\prj.tlb
# End Source File
# Begin Source File

SOURCE=.\resJap\prjJap.rc

!IF  "$(CFG)" == "Actbar - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\proppage.cpp
DEP_CPP_PROPP=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\proppage.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Crt\QSORT.C
DEP_CPP_QSORT=\
	{$(INCLUDE)}"cruntime.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Crt\rand.c
DEP_CPP_RAND_=\
	{$(INCLUDE)}"cruntime.h"\
	{$(INCLUDE)}"mtdll.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\returnbool.cpp
DEP_CPP_RETUR=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\returnstring.cpp
DEP_CPP_RETURN=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ScriptString.cpp
DEP_CPP_SCRIP=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\ScriptString.h"\
	".\Utility.h"\
	{$(INCLUDE)}"usp10.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ShortCuts.cpp
DEP_CPP_SHORT=\
	".\Band.h"\
	".\Bar.h"\
	".\Custom.h"\
	".\CustomizeListbox.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\ShortCuts.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\StaticLink.cpp
DEP_CPP_STATI=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\StaticLink.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Streams.cpp
DEP_CPP_STREA=\
	".\DEBUG.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\Streams.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\support.cpp
DEP_CPP_SUPPO=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\tool.cpp
DEP_CPP_TOOL_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\CBList.h"\
	".\ChildBands.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDropMgr.h"\
	".\Dib.h"\
	".\Dispids.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\ImageMgr.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\ScriptString.h"\
	".\ShortCuts.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Tppopup.h"\
	".\Utility.h"\
	{$(INCLUDE)}"usp10.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Tools.cpp
DEP_CPP_TOOLS=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\Btabs.h"\
	".\ChildBands.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\ImageMgr.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\localizer.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Tooltip.cpp
DEP_CPP_TOOLT=\
	".\Bar.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tooltip.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Tppopup.cpp
DEP_CPP_TPPOP=\
	".\Band.h"\
	".\Bar.h"\
	".\CBList.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDropMgr.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Tppopup.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Utility.cpp
DEP_CPP_UTILI=\
	".\Bar.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\ScriptString.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	{$(INCLUDE)}"usp10.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Crt\WCSCAT.C
DEP_CPP_WCSCA=\
	{$(INCLUDE)}"cruntime.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Crt\WCSCMP.C
DEP_CPP_WCSCM=\
	{$(INCLUDE)}"cruntime.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Crt\WCSLEN.C
DEP_CPP_WCSLE=\
	{$(INCLUDE)}"cruntime.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

# SUBTRACT BASE CPP /YX /Yc /Yu
# SUBTRACT CPP /YX /Yc /Yu

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\WhatsThisHelp.cpp
DEP_CPP_WHATS=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDropMgr.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\WindowProc.cpp
DEP_CPP_WINDO=\
	".\Band.h"\
	".\Bar.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\Designer\DragDropMgr.h"\
	".\Dock.h"\
	".\DropSource.h"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\PrivateInterfaces.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\ShortCuts.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Utility.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\xevents.cpp
DEP_CPP_XEVEN=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\EventLog.h"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Utility.h"\
	".\XEVENTS.H"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Release"

!ELSEIF  "$(CFG)" == "Actbar - Win32 AB DLL Debug"

!ELSEIF  "$(CFG)" == "Actbar - Win32 Japanese AB DLL Release"

!ENDIF 

# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl;fi;fd"
# Begin Source File

SOURCE=.\Band.h
# End Source File
# Begin Source File

SOURCE=.\Bands.h
# End Source File
# Begin Source File

SOURCE=.\Bar.h
# End Source File
# Begin Source File

SOURCE=.\Btabs.h
# End Source File
# Begin Source File

SOURCE=.\Categories.h
# End Source File
# Begin Source File

SOURCE=.\CategoryMgr.h
# End Source File
# Begin Source File

SOURCE=.\CATIDREG.H
# End Source File
# Begin Source File

SOURCE=.\CBList.h
# End Source File
# Begin Source File

SOURCE=.\ChildBands.h
# End Source File
# Begin Source File

SOURCE="C:\Program Files\Microsoft Visual Studio\VC98\CRT\SRC\CRUNTIME.H"
# End Source File
# Begin Source File

SOURCE=.\Custom.h
# End Source File
# Begin Source File

SOURCE=.\CustomizeListbox.h
# End Source File
# Begin Source File

SOURCE=.\customproxy.h
# End Source File
# Begin Source File

SOURCE=.\customtool.h
# End Source File
# Begin Source File

SOURCE=.\DDASYNC.H
# End Source File
# Begin Source File

SOURCE=.\DEBUG.H
# End Source File
# Begin Source File

SOURCE=.\Designer\DesignerInterfaces.h
# End Source File
# Begin Source File

SOURCE=.\Dib.h
# End Source File
# Begin Source File

SOURCE=.\Dispids.h
# End Source File
# Begin Source File

SOURCE=.\Dock.h
# End Source File
# Begin Source File

SOURCE=.\Designer\DragDrop.h
# End Source File
# Begin Source File

SOURCE=.\Designer\DragDropMgr.h
# End Source File
# Begin Source File

SOURCE=.\DropSource.h
# End Source File
# Begin Source File

SOURCE=.\ENUMX.H
# End Source File
# Begin Source File

SOURCE=.\ERRORS.H
# End Source File
# Begin Source File

SOURCE=.\EventLog.h
# End Source File
# Begin Source File

SOURCE=.\FDIALOG.H
# End Source File
# Begin Source File

SOURCE=.\Flicker.h
# End Source File
# Begin Source File

SOURCE=.\fontholder.h
# End Source File
# Begin Source File

SOURCE=.\FREGKEY.H
# End Source File
# Begin Source File

SOURCE=.\FWND.H
# End Source File
# Begin Source File

SOURCE=.\GLOBALS.H
# End Source File
# Begin Source File

SOURCE=.\GUIDS.H
# End Source File
# Begin Source File

SOURCE=.\ImageMgr.h
# End Source File
# Begin Source File

SOURCE=.\interfaces.h
# End Source File
# Begin Source File

SOURCE=.\IPSERVER.H
# End Source File
# Begin Source File

SOURCE=.\localizer.h
# End Source File
# Begin Source File

SOURCE=.\MAP.H
# End Source File
# Begin Source File

SOURCE=.\MINIWIN.H
# End Source File
# Begin Source File

SOURCE=.\MTDLL.H
# End Source File
# Begin Source File

SOURCE=.\Popupwin.h
# End Source File
# Begin Source File

SOURCE=.\precomp.h
# End Source File
# Begin Source File

SOURCE=.\PrivateInterfaces.h
# End Source File
# Begin Source File

SOURCE=.\resource.h
# End Source File
# Begin Source File

SOURCE=.\returnbool.h
# End Source File
# Begin Source File

SOURCE=.\returnstring.h
# End Source File
# Begin Source File

SOURCE=.\ScriptString.h
# End Source File
# Begin Source File

SOURCE=.\ShortCuts.h
# End Source File
# Begin Source File

SOURCE=.\StaticLink.h
# End Source File
# Begin Source File

SOURCE=.\Streams.h
# End Source File
# Begin Source File

SOURCE=.\SUPPORT.H
# End Source File
# Begin Source File

SOURCE=.\Tool.h
# End Source File
# Begin Source File

SOURCE=.\Tools.h
# End Source File
# Begin Source File

SOURCE=.\Tooltip.h
# End Source File
# Begin Source File

SOURCE=.\Tppopup.h
# End Source File
# Begin Source File

SOURCE=.\Utility.h
# End Source File
# Begin Source File

SOURCE=.\WindowProc.h
# End Source File
# Begin Source File

SOURCE=.\XEVENTS.H
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;cnt;rtf;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\Res\BITMAP1.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\bitmap3.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\bmp00001.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\BMP00002.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\BMP00003.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\bmp00004.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\bmp00005.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\bmp00006.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\bmp00007.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\cascade.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\CHECKBMP.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\cur00001.cur
# End Source File
# Begin Source File

SOURCE=.\Res\cur00002.cur
# End Source File
# Begin Source File

SOURCE=.\Res\cursor1.cur
# End Source File
# Begin Source File

SOURCE=.\Res\expandve.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\GRAB1.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\GRAB2.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\GRAB3.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\GRAB4.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\H_POINT.CUR
# End Source File
# Begin Source File

SOURCE=.\Res\hsplitba.cur
# End Source File
# Begin Source File

SOURCE=.\Res\icon1.ico
# End Source File
# Begin Source File

SOURCE=.\Res\ICON2.ICO
# End Source File
# Begin Source File

SOURCE=.\Res\IDCC_DRA.CUR
# End Source File
# Begin Source File

SOURCE=.\Res\menuexpa.bmp
# End Source File
# Begin Source File

SOURCE=.\Res\MENUSCRO.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\OVERFLOW.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\splith.cur
# End Source File
# Begin Source File

SOURCE=.\Res\splitv.cur
# End Source File
# Begin Source File

SOURCE=.\Res\SUBMENUB.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\SYSCLOSE.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\SYSMAX.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\SYSMINIM.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\SYSRESTO.BMP
# End Source File
# Begin Source File

SOURCE=.\Res\titehorz.bmp
# End Source File
# End Group
# End Target
# End Project
