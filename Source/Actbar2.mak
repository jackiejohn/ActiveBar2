# Microsoft Developer Studio Generated NMAKE File, Based on Actbar2.dsp
!IF "$(CFG)" == ""
CFG=Actbar - Win32 Release
!MESSAGE No configuration specified. Defaulting to Actbar - Win32 Release.
!ENDIF 

!IF "$(CFG)" != "Actbar - Win32 Release" && "$(CFG)" != "Actbar - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Actbar2.mak" CFG="Actbar - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Actbar - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Actbar - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 

!IF  "$(CFG)" == "Actbar - Win32 Release"

OUTDIR=.\Release
INTDIR=.\Release
# Begin Custom Macros
OutDir=.\.\Release
# End Custom Macros

!IF "$(RECURSE)" == "0" 

ALL : "$(OUTDIR)\Actbar.ocx"

!ELSE 

ALL : "$(OUTDIR)\Actbar.ocx"

!ENDIF 

CLEAN :
	-@erase "$(INTDIR)\Actbar2.pch"
	-@erase "$(INTDIR)\band.obj"
	-@erase "$(INTDIR)\Bands.obj"
	-@erase "$(INTDIR)\Bar.obj"
	-@erase "$(INTDIR)\bitmapmng.obj"
	-@erase "$(INTDIR)\Btabs.obj"
	-@erase "$(INTDIR)\CategoryMgr.obj"
	-@erase "$(INTDIR)\catidreg.obj"
	-@erase "$(INTDIR)\CBList.obj"
	-@erase "$(INTDIR)\CDIB.OBJ"
	-@erase "$(INTDIR)\Custom.obj"
	-@erase "$(INTDIR)\CustomizeListbox.obj"
	-@erase "$(INTDIR)\customproxy.obj"
	-@erase "$(INTDIR)\DDASYNC.OBJ"
	-@erase "$(INTDIR)\debug.obj"
	-@erase "$(INTDIR)\DIBSTUFF.OBJ"
	-@erase "$(INTDIR)\Dock.obj"
	-@erase "$(INTDIR)\Draw.obj"
	-@erase "$(INTDIR)\enumx.obj"
	-@erase "$(INTDIR)\FDIALOG.OBJ"
	-@erase "$(INTDIR)\filestream.obj"
	-@erase "$(INTDIR)\Flicker.obj"
	-@erase "$(INTDIR)\fontholder.obj"
	-@erase "$(INTDIR)\fregkey.obj"
	-@erase "$(INTDIR)\FWND.OBJ"
	-@erase "$(INTDIR)\guids.obj"
	-@erase "$(INTDIR)\Ibar.obj"
	-@erase "$(INTDIR)\ipserver.obj"
	-@erase "$(INTDIR)\Map.obj"
	-@erase "$(INTDIR)\memstream.obj"
	-@erase "$(INTDIR)\Miniwin.obj"
	-@erase "$(INTDIR)\opendesign.obj"
	-@erase "$(INTDIR)\page.obj"
	-@erase "$(INTDIR)\Pages.obj"
	-@erase "$(INTDIR)\pictureholder.obj"
	-@erase "$(INTDIR)\Popupwin.obj"
	-@erase "$(INTDIR)\precomp.obj"
	-@erase "$(INTDIR)\Private.obj"
	-@erase "$(INTDIR)\prj.res"
	-@erase "$(INTDIR)\proppage.obj"
	-@erase "$(INTDIR)\QSORT.OBJ"
	-@erase "$(INTDIR)\rand.obj"
	-@erase "$(INTDIR)\returnbool.obj"
	-@erase "$(INTDIR)\returnstring.obj"
	-@erase "$(INTDIR)\StaticLink.obj"
	-@erase "$(INTDIR)\support.obj"
	-@erase "$(INTDIR)\tool.obj"
	-@erase "$(INTDIR)\Tool2.obj"
	-@erase "$(INTDIR)\ToolDrop.obj"
	-@erase "$(INTDIR)\Tools.obj"
	-@erase "$(INTDIR)\Tooltip.obj"
	-@erase "$(INTDIR)\Tppopup.obj"
	-@erase "$(INTDIR)\vc50.idb"
	-@erase "$(INTDIR)\WCSCAT.OBJ"
	-@erase "$(INTDIR)\WCSCMP.OBJ"
	-@erase "$(INTDIR)\WCSLEN.OBJ"
	-@erase "$(INTDIR)\xevents.obj"
	-@erase "$(OUTDIR)\Actbar.exp"
	-@erase "$(OUTDIR)\Actbar.lib"
	-@erase "$(OUTDIR)\Actbar.ocx"
	-@erase ".\interfaces.h"
	-@erase ".\prj.tlb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MT /W3 /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "ACTBARMAIN" /Fp"$(INTDIR)\Actbar2.pch" /Yu"precomp.h" /Fo"$(INTDIR)\\"\
 /Fd"$(INTDIR)\\" /FD /c 
CPP_OBJS=.\Release/
CPP_SBRS=.

.c{$(CPP_OBJS)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(CPP_OBJS)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(CPP_OBJS)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(CPP_SBRS)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(CPP_SBRS)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(CPP_SBRS)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/nologo /D "NDEBUG" /mktyplib203 /win32 
RSC=rc.exe
RSC_PROJ=/l 0x409 /fo"$(INTDIR)\prj.res" /d "NDEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\Actbar2.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib\
 shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib version.lib /nologo\
 /entry:"DllMain" /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)\Actbar.pdb" /machine:I386 /nodefaultlib /def:".\prj.def"\
 /out:"$(OUTDIR)\Actbar.ocx" /implib:"$(OUTDIR)\Actbar.lib" 
DEF_FILE= \
	".\prj.def"
LINK32_OBJS= \
	"$(INTDIR)\band.obj" \
	"$(INTDIR)\Bands.obj" \
	"$(INTDIR)\Bar.obj" \
	"$(INTDIR)\bitmapmng.obj" \
	"$(INTDIR)\Btabs.obj" \
	"$(INTDIR)\CategoryMgr.obj" \
	"$(INTDIR)\catidreg.obj" \
	"$(INTDIR)\CBList.obj" \
	"$(INTDIR)\CDIB.OBJ" \
	"$(INTDIR)\Custom.obj" \
	"$(INTDIR)\CustomizeListbox.obj" \
	"$(INTDIR)\customproxy.obj" \
	"$(INTDIR)\DDASYNC.OBJ" \
	"$(INTDIR)\debug.obj" \
	"$(INTDIR)\DIBSTUFF.OBJ" \
	"$(INTDIR)\Dock.obj" \
	"$(INTDIR)\Draw.obj" \
	"$(INTDIR)\enumx.obj" \
	"$(INTDIR)\FDIALOG.OBJ" \
	"$(INTDIR)\filestream.obj" \
	"$(INTDIR)\Flicker.obj" \
	"$(INTDIR)\fontholder.obj" \
	"$(INTDIR)\fregkey.obj" \
	"$(INTDIR)\FWND.OBJ" \
	"$(INTDIR)\guids.obj" \
	"$(INTDIR)\Ibar.obj" \
	"$(INTDIR)\ipserver.obj" \
	"$(INTDIR)\Map.obj" \
	"$(INTDIR)\memstream.obj" \
	"$(INTDIR)\Miniwin.obj" \
	"$(INTDIR)\opendesign.obj" \
	"$(INTDIR)\page.obj" \
	"$(INTDIR)\Pages.obj" \
	"$(INTDIR)\pictureholder.obj" \
	"$(INTDIR)\Popupwin.obj" \
	"$(INTDIR)\precomp.obj" \
	"$(INTDIR)\Private.obj" \
	"$(INTDIR)\prj.res" \
	"$(INTDIR)\proppage.obj" \
	"$(INTDIR)\QSORT.OBJ" \
	"$(INTDIR)\rand.obj" \
	"$(INTDIR)\returnbool.obj" \
	"$(INTDIR)\returnstring.obj" \
	"$(INTDIR)\StaticLink.obj" \
	"$(INTDIR)\support.obj" \
	"$(INTDIR)\tool.obj" \
	"$(INTDIR)\Tool2.obj" \
	"$(INTDIR)\ToolDrop.obj" \
	"$(INTDIR)\Tools.obj" \
	"$(INTDIR)\Tooltip.obj" \
	"$(INTDIR)\Tppopup.obj" \
	"$(INTDIR)\WCSCAT.OBJ" \
	"$(INTDIR)\WCSCMP.OBJ" \
	"$(INTDIR)\WCSLEN.OBJ" \
	"$(INTDIR)\xevents.obj" \
	".\Crt\DLLSUPP.OBJ" \
	".\Crt\MEMMOVE.OBJ"

"$(OUTDIR)\Actbar.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

OUTDIR=.\Debug
INTDIR=.\Debug
# Begin Custom Macros
OutDir=.\.\Debug
# End Custom Macros

!IF "$(RECURSE)" == "0" 

ALL : "$(OUTDIR)\Actbar.ocx" "$(OUTDIR)\Actbar2.bsc"

!ELSE 

ALL : "$(OUTDIR)\Actbar.ocx" "$(OUTDIR)\Actbar2.bsc"

!ENDIF 

CLEAN :
	-@erase "$(INTDIR)\Actbar2.pch"
	-@erase "$(INTDIR)\band.obj"
	-@erase "$(INTDIR)\band.sbr"
	-@erase "$(INTDIR)\Bands.obj"
	-@erase "$(INTDIR)\Bands.sbr"
	-@erase "$(INTDIR)\Bar.obj"
	-@erase "$(INTDIR)\Bar.sbr"
	-@erase "$(INTDIR)\bitmapmng.obj"
	-@erase "$(INTDIR)\bitmapmng.sbr"
	-@erase "$(INTDIR)\Btabs.obj"
	-@erase "$(INTDIR)\Btabs.sbr"
	-@erase "$(INTDIR)\CategoryMgr.obj"
	-@erase "$(INTDIR)\CategoryMgr.sbr"
	-@erase "$(INTDIR)\catidreg.obj"
	-@erase "$(INTDIR)\catidreg.sbr"
	-@erase "$(INTDIR)\CBList.obj"
	-@erase "$(INTDIR)\CBList.sbr"
	-@erase "$(INTDIR)\CDIB.OBJ"
	-@erase "$(INTDIR)\CDIB.SBR"
	-@erase "$(INTDIR)\Custom.obj"
	-@erase "$(INTDIR)\Custom.sbr"
	-@erase "$(INTDIR)\CustomizeListbox.obj"
	-@erase "$(INTDIR)\CustomizeListbox.sbr"
	-@erase "$(INTDIR)\customproxy.obj"
	-@erase "$(INTDIR)\customproxy.sbr"
	-@erase "$(INTDIR)\DDASYNC.OBJ"
	-@erase "$(INTDIR)\DDASYNC.SBR"
	-@erase "$(INTDIR)\debug.obj"
	-@erase "$(INTDIR)\debug.sbr"
	-@erase "$(INTDIR)\DIBSTUFF.OBJ"
	-@erase "$(INTDIR)\DIBSTUFF.SBR"
	-@erase "$(INTDIR)\Dock.obj"
	-@erase "$(INTDIR)\Dock.sbr"
	-@erase "$(INTDIR)\Draw.obj"
	-@erase "$(INTDIR)\Draw.sbr"
	-@erase "$(INTDIR)\enumx.obj"
	-@erase "$(INTDIR)\enumx.sbr"
	-@erase "$(INTDIR)\FDIALOG.OBJ"
	-@erase "$(INTDIR)\FDIALOG.SBR"
	-@erase "$(INTDIR)\filestream.obj"
	-@erase "$(INTDIR)\filestream.sbr"
	-@erase "$(INTDIR)\Flicker.obj"
	-@erase "$(INTDIR)\Flicker.sbr"
	-@erase "$(INTDIR)\fontholder.obj"
	-@erase "$(INTDIR)\fontholder.sbr"
	-@erase "$(INTDIR)\fregkey.obj"
	-@erase "$(INTDIR)\fregkey.sbr"
	-@erase "$(INTDIR)\FWND.OBJ"
	-@erase "$(INTDIR)\FWND.SBR"
	-@erase "$(INTDIR)\guids.obj"
	-@erase "$(INTDIR)\guids.sbr"
	-@erase "$(INTDIR)\Ibar.obj"
	-@erase "$(INTDIR)\Ibar.sbr"
	-@erase "$(INTDIR)\ipserver.obj"
	-@erase "$(INTDIR)\ipserver.sbr"
	-@erase "$(INTDIR)\Map.obj"
	-@erase "$(INTDIR)\Map.sbr"
	-@erase "$(INTDIR)\memstream.obj"
	-@erase "$(INTDIR)\memstream.sbr"
	-@erase "$(INTDIR)\Miniwin.obj"
	-@erase "$(INTDIR)\Miniwin.sbr"
	-@erase "$(INTDIR)\opendesign.obj"
	-@erase "$(INTDIR)\opendesign.sbr"
	-@erase "$(INTDIR)\page.obj"
	-@erase "$(INTDIR)\page.sbr"
	-@erase "$(INTDIR)\Pages.obj"
	-@erase "$(INTDIR)\Pages.sbr"
	-@erase "$(INTDIR)\pictureholder.obj"
	-@erase "$(INTDIR)\pictureholder.sbr"
	-@erase "$(INTDIR)\Popupwin.obj"
	-@erase "$(INTDIR)\Popupwin.sbr"
	-@erase "$(INTDIR)\precomp.obj"
	-@erase "$(INTDIR)\precomp.sbr"
	-@erase "$(INTDIR)\Private.obj"
	-@erase "$(INTDIR)\Private.sbr"
	-@erase "$(INTDIR)\prj.res"
	-@erase "$(INTDIR)\proppage.obj"
	-@erase "$(INTDIR)\proppage.sbr"
	-@erase "$(INTDIR)\QSORT.OBJ"
	-@erase "$(INTDIR)\QSORT.SBR"
	-@erase "$(INTDIR)\rand.obj"
	-@erase "$(INTDIR)\rand.sbr"
	-@erase "$(INTDIR)\returnbool.obj"
	-@erase "$(INTDIR)\returnbool.sbr"
	-@erase "$(INTDIR)\returnstring.obj"
	-@erase "$(INTDIR)\returnstring.sbr"
	-@erase "$(INTDIR)\StaticLink.obj"
	-@erase "$(INTDIR)\StaticLink.sbr"
	-@erase "$(INTDIR)\support.obj"
	-@erase "$(INTDIR)\support.sbr"
	-@erase "$(INTDIR)\tool.obj"
	-@erase "$(INTDIR)\tool.sbr"
	-@erase "$(INTDIR)\Tool2.obj"
	-@erase "$(INTDIR)\Tool2.sbr"
	-@erase "$(INTDIR)\ToolDrop.obj"
	-@erase "$(INTDIR)\ToolDrop.sbr"
	-@erase "$(INTDIR)\Tools.obj"
	-@erase "$(INTDIR)\Tools.sbr"
	-@erase "$(INTDIR)\Tooltip.obj"
	-@erase "$(INTDIR)\Tooltip.sbr"
	-@erase "$(INTDIR)\Tppopup.obj"
	-@erase "$(INTDIR)\Tppopup.sbr"
	-@erase "$(INTDIR)\vc50.idb"
	-@erase "$(INTDIR)\vc50.pdb"
	-@erase "$(INTDIR)\WCSCAT.OBJ"
	-@erase "$(INTDIR)\WCSCAT.SBR"
	-@erase "$(INTDIR)\WCSCMP.OBJ"
	-@erase "$(INTDIR)\WCSCMP.SBR"
	-@erase "$(INTDIR)\WCSLEN.OBJ"
	-@erase "$(INTDIR)\WCSLEN.SBR"
	-@erase "$(INTDIR)\xevents.obj"
	-@erase "$(INTDIR)\xevents.sbr"
	-@erase "$(OUTDIR)\Actbar.exp"
	-@erase "$(OUTDIR)\Actbar.ilk"
	-@erase "$(OUTDIR)\Actbar.lib"
	-@erase "$(OUTDIR)\Actbar.ocx"
	-@erase "$(OUTDIR)\Actbar.pdb"
	-@erase "$(OUTDIR)\Actbar2.bsc"
	-@erase ".\interfaces.h"
	-@erase ".\prj.tlb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS"\
 /FR"$(INTDIR)\\" /Fp"$(INTDIR)\Actbar2.pch" /Yu"precomp.h" /Fo"$(INTDIR)\\"\
 /Fd"$(INTDIR)\\" /FD /c 
CPP_OBJS=.\Debug/
CPP_SBRS=.\Debug/

.c{$(CPP_OBJS)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(CPP_OBJS)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(CPP_OBJS)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(CPP_SBRS)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(CPP_SBRS)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(CPP_SBRS)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/nologo /D "_DEBUG" /mktyplib203 /win32 
RSC=rc.exe
RSC_PROJ=/l 0x409 /fo"$(INTDIR)\prj.res" /d "_DEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\Actbar2.bsc" 
BSC32_SBRS= \
	"$(INTDIR)\band.sbr" \
	"$(INTDIR)\Bands.sbr" \
	"$(INTDIR)\Bar.sbr" \
	"$(INTDIR)\bitmapmng.sbr" \
	"$(INTDIR)\Btabs.sbr" \
	"$(INTDIR)\CategoryMgr.sbr" \
	"$(INTDIR)\catidreg.sbr" \
	"$(INTDIR)\CBList.sbr" \
	"$(INTDIR)\CDIB.SBR" \
	"$(INTDIR)\Custom.sbr" \
	"$(INTDIR)\CustomizeListbox.sbr" \
	"$(INTDIR)\customproxy.sbr" \
	"$(INTDIR)\DDASYNC.SBR" \
	"$(INTDIR)\debug.sbr" \
	"$(INTDIR)\DIBSTUFF.SBR" \
	"$(INTDIR)\Dock.sbr" \
	"$(INTDIR)\Draw.sbr" \
	"$(INTDIR)\enumx.sbr" \
	"$(INTDIR)\FDIALOG.SBR" \
	"$(INTDIR)\filestream.sbr" \
	"$(INTDIR)\Flicker.sbr" \
	"$(INTDIR)\fontholder.sbr" \
	"$(INTDIR)\fregkey.sbr" \
	"$(INTDIR)\FWND.SBR" \
	"$(INTDIR)\guids.sbr" \
	"$(INTDIR)\Ibar.sbr" \
	"$(INTDIR)\ipserver.sbr" \
	"$(INTDIR)\Map.sbr" \
	"$(INTDIR)\memstream.sbr" \
	"$(INTDIR)\Miniwin.sbr" \
	"$(INTDIR)\opendesign.sbr" \
	"$(INTDIR)\page.sbr" \
	"$(INTDIR)\Pages.sbr" \
	"$(INTDIR)\pictureholder.sbr" \
	"$(INTDIR)\Popupwin.sbr" \
	"$(INTDIR)\precomp.sbr" \
	"$(INTDIR)\Private.sbr" \
	"$(INTDIR)\proppage.sbr" \
	"$(INTDIR)\QSORT.SBR" \
	"$(INTDIR)\rand.sbr" \
	"$(INTDIR)\returnbool.sbr" \
	"$(INTDIR)\returnstring.sbr" \
	"$(INTDIR)\StaticLink.sbr" \
	"$(INTDIR)\support.sbr" \
	"$(INTDIR)\tool.sbr" \
	"$(INTDIR)\Tool2.sbr" \
	"$(INTDIR)\ToolDrop.sbr" \
	"$(INTDIR)\Tools.sbr" \
	"$(INTDIR)\Tooltip.sbr" \
	"$(INTDIR)\Tppopup.sbr" \
	"$(INTDIR)\WCSCAT.SBR" \
	"$(INTDIR)\WCSCMP.SBR" \
	"$(INTDIR)\WCSLEN.SBR" \
	"$(INTDIR)\xevents.sbr"

"$(OUTDIR)\Actbar2.bsc" : "$(OUTDIR)" $(BSC32_SBRS)
    $(BSC32) @<<
  $(BSC32_FLAGS) $(BSC32_SBRS)
<<

LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib\
 shell32.lib ole32.lib oleaut32.lib uuid.lib comctl32.lib version.lib /nologo\
 /subsystem:windows /dll /incremental:yes /pdb:"$(OUTDIR)\Actbar.pdb" /debug\
 /machine:I386 /def:".\prj.def" /out:"$(OUTDIR)\Actbar.ocx"\
 /implib:"$(OUTDIR)\Actbar.lib" 
DEF_FILE= \
	".\prj.def"
LINK32_OBJS= \
	"$(INTDIR)\band.obj" \
	"$(INTDIR)\Bands.obj" \
	"$(INTDIR)\Bar.obj" \
	"$(INTDIR)\bitmapmng.obj" \
	"$(INTDIR)\Btabs.obj" \
	"$(INTDIR)\CategoryMgr.obj" \
	"$(INTDIR)\catidreg.obj" \
	"$(INTDIR)\CBList.obj" \
	"$(INTDIR)\CDIB.OBJ" \
	"$(INTDIR)\Custom.obj" \
	"$(INTDIR)\CustomizeListbox.obj" \
	"$(INTDIR)\customproxy.obj" \
	"$(INTDIR)\DDASYNC.OBJ" \
	"$(INTDIR)\debug.obj" \
	"$(INTDIR)\DIBSTUFF.OBJ" \
	"$(INTDIR)\Dock.obj" \
	"$(INTDIR)\Draw.obj" \
	"$(INTDIR)\enumx.obj" \
	"$(INTDIR)\FDIALOG.OBJ" \
	"$(INTDIR)\filestream.obj" \
	"$(INTDIR)\Flicker.obj" \
	"$(INTDIR)\fontholder.obj" \
	"$(INTDIR)\fregkey.obj" \
	"$(INTDIR)\FWND.OBJ" \
	"$(INTDIR)\guids.obj" \
	"$(INTDIR)\Ibar.obj" \
	"$(INTDIR)\ipserver.obj" \
	"$(INTDIR)\Map.obj" \
	"$(INTDIR)\memstream.obj" \
	"$(INTDIR)\Miniwin.obj" \
	"$(INTDIR)\opendesign.obj" \
	"$(INTDIR)\page.obj" \
	"$(INTDIR)\Pages.obj" \
	"$(INTDIR)\pictureholder.obj" \
	"$(INTDIR)\Popupwin.obj" \
	"$(INTDIR)\precomp.obj" \
	"$(INTDIR)\Private.obj" \
	"$(INTDIR)\prj.res" \
	"$(INTDIR)\proppage.obj" \
	"$(INTDIR)\QSORT.OBJ" \
	"$(INTDIR)\rand.obj" \
	"$(INTDIR)\returnbool.obj" \
	"$(INTDIR)\returnstring.obj" \
	"$(INTDIR)\StaticLink.obj" \
	"$(INTDIR)\support.obj" \
	"$(INTDIR)\tool.obj" \
	"$(INTDIR)\Tool2.obj" \
	"$(INTDIR)\ToolDrop.obj" \
	"$(INTDIR)\Tools.obj" \
	"$(INTDIR)\Tooltip.obj" \
	"$(INTDIR)\Tppopup.obj" \
	"$(INTDIR)\WCSCAT.OBJ" \
	"$(INTDIR)\WCSCMP.OBJ" \
	"$(INTDIR)\WCSLEN.OBJ" \
	"$(INTDIR)\xevents.obj" \
	".\Crt\DLLSUPP.OBJ" \
	".\Crt\MEMMOVE.OBJ"

"$(OUTDIR)\Actbar.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ENDIF 


!IF "$(CFG)" == "Actbar - Win32 Release" || "$(CFG)" == "Actbar - Win32 Debug"
SOURCE=.\band.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_BAND_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Page.h"\
	".\Pages.h"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\band.obj" : $(SOURCE) $(DEP_CPP_BAND_) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_BAND_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Page.h"\
	".\Pages.h"\
	".\Popupwin.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\band.obj"	"$(INTDIR)\band.sbr" : $(SOURCE) $(DEP_CPP_BAND_)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Bands.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_BANDS=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Bands.obj" : $(SOURCE) $(DEP_CPP_BANDS) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_BANDS=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Bands.obj"	"$(INTDIR)\Bands.sbr" : $(SOURCE) $(DEP_CPP_BANDS)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Bar.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_BAR_C=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\fdialog.h"\
	".\DESIGNER\interfaces2.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dispids.h"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\Hlp\ActiveBar.hh"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\memstream.h"\
	".\MINIWIN.H"\
	".\Page.h"\
	".\Pages.h"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Tooltip.h"\
	".\XEVENTS.H"\
	

"$(INTDIR)\Bar.obj" : $(SOURCE) $(DEP_CPP_BAR_C) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_BAR_C=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\fdialog.h"\
	".\DESIGNER\interfaces2.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dispids.h"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\Hlp\ActiveBar.hh"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\memstream.h"\
	".\MINIWIN.H"\
	".\Page.h"\
	".\Pages.h"\
	".\Popupwin.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Tooltip.h"\
	".\XEVENTS.H"\
	

"$(INTDIR)\Bar.obj"	"$(INTDIR)\Bar.sbr" : $(SOURCE) $(DEP_CPP_BAR_C)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\bitmapmng.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_BITMA=\
	".\bitmapmng.h"\
	".\DEBUG.H"\
	".\DIBSTUFF.H"\
	".\MAP.H"\
	".\precomp.h"\
	

"$(INTDIR)\bitmapmng.obj" : $(SOURCE) $(DEP_CPP_BITMA) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_BITMA=\
	".\bitmapmng.h"\
	".\DEBUG.H"\
	".\DIBSTUFF.H"\
	".\MAP.H"\
	

"$(INTDIR)\bitmapmng.obj"	"$(INTDIR)\bitmapmng.sbr" : $(SOURCE)\
 $(DEP_CPP_BITMA) "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\Btabs.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_BTABS=\
	".\Btabs.h"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\Btabs.obj" : $(SOURCE) $(DEP_CPP_BTABS) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_BTABS=\
	".\Btabs.h"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\Btabs.obj"	"$(INTDIR)\Btabs.sbr" : $(SOURCE) $(DEP_CPP_BTABS)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\DESIGNER\CategoryMgr.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_CATEG=\
	".\DESIGNER\CategoryMgr.h"\
	".\designer\enumx.h"\
	".\designer\globals.h"\
	".\DESIGNER\MSGHELP.H"\
	".\DESIGNER\precomp.h"\
	".\designer\support.h"\
	".\interfaces.h"\
	".\MAP.H"\
	

"$(INTDIR)\CategoryMgr.obj" : $(SOURCE) $(DEP_CPP_CATEG) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_CATEG=\
	".\DESIGNER\CategoryMgr.h"\
	".\designer\enumx.h"\
	".\designer\globals.h"\
	".\DESIGNER\MSGHELP.H"\
	".\designer\support.h"\
	".\interfaces.h"\
	".\MAP.H"\
	

"$(INTDIR)\CategoryMgr.obj"	"$(INTDIR)\CategoryMgr.sbr" : $(SOURCE)\
 $(DEP_CPP_CATEG) "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\catidreg.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_CATID=\
	".\CATIDREG.H"\
	".\precomp.h"\
	

"$(INTDIR)\catidreg.obj" : $(SOURCE) $(DEP_CPP_CATID) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_CATID=\
	".\CATIDREG.H"\
	

"$(INTDIR)\catidreg.obj"	"$(INTDIR)\catidreg.sbr" : $(SOURCE) $(DEP_CPP_CATID)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\CBList.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_CBLIS=\
	".\CBList.h"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\CBList.obj" : $(SOURCE) $(DEP_CPP_CBLIS) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_CBLIS=\
	".\CBList.h"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\CBList.obj"	"$(INTDIR)\CBList.sbr" : $(SOURCE) $(DEP_CPP_CBLIS)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\CDIB.CPP

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_CDIB_=\
	".\CDIB.H"\
	".\precomp.h"\
	

"$(INTDIR)\CDIB.OBJ" : $(SOURCE) $(DEP_CPP_CDIB_) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_CDIB_=\
	".\CDIB.H"\
	

"$(INTDIR)\CDIB.OBJ"	"$(INTDIR)\CDIB.SBR" : $(SOURCE) $(DEP_CPP_CDIB_)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\Custom.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_CUSTO=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\fdialog.h"\
	".\DESIGNER\MSGHELP.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\designer\toollist.h"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Page.h"\
	".\Pages.h"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Custom.obj" : $(SOURCE) $(DEP_CPP_CUSTO) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_CUSTO=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\fdialog.h"\
	".\DESIGNER\MSGHELP.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\designer\toollist.h"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Page.h"\
	".\Pages.h"\
	".\Popupwin.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Custom.obj"	"$(INTDIR)\Custom.sbr" : $(SOURCE) $(DEP_CPP_CUSTO)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\CustomizeListbox.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_CUSTOM=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\CustomizeListbox.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\CategoryMgr.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dispids.h"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\XEVENTS.H"\
	

"$(INTDIR)\CustomizeListbox.obj" : $(SOURCE) $(DEP_CPP_CUSTOM) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_CUSTOM=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\CustomizeListbox.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\CategoryMgr.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dispids.h"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\XEVENTS.H"\
	

"$(INTDIR)\CustomizeListbox.obj"	"$(INTDIR)\CustomizeListbox.sbr" : $(SOURCE)\
 $(DEP_CPP_CUSTOM) "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\customproxy.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_CUSTOMP=\
	".\customproxy.h"\
	".\GUIDS.H"\
	".\interfaces.h"\
	".\precomp.h"\
	".\XEVENTS.H"\
	

"$(INTDIR)\customproxy.obj" : $(SOURCE) $(DEP_CPP_CUSTOMP) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_CUSTOMP=\
	".\customproxy.h"\
	".\GUIDS.H"\
	".\interfaces.h"\
	".\XEVENTS.H"\
	

"$(INTDIR)\customproxy.obj"	"$(INTDIR)\customproxy.sbr" : $(SOURCE)\
 $(DEP_CPP_CUSTOMP) "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\DDASYNC.CPP

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_DDASY=\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\precomp.h"\
	

"$(INTDIR)\DDASYNC.OBJ" : $(SOURCE) $(DEP_CPP_DDASY) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_DDASY=\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	

"$(INTDIR)\DDASYNC.OBJ"	"$(INTDIR)\DDASYNC.SBR" : $(SOURCE) $(DEP_CPP_DDASY)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\debug.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_DEBUG=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\debug.obj" : $(SOURCE) $(DEP_CPP_DEBUG) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_DEBUG=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\debug.obj"	"$(INTDIR)\debug.sbr" : $(SOURCE) $(DEP_CPP_DEBUG)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\DIBSTUFF.CPP

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_DIBST=\
	".\DIBSTUFF.H"\
	".\precomp.h"\
	

"$(INTDIR)\DIBSTUFF.OBJ" : $(SOURCE) $(DEP_CPP_DIBST) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_DIBST=\
	".\DIBSTUFF.H"\
	

"$(INTDIR)\DIBSTUFF.OBJ"	"$(INTDIR)\DIBSTUFF.SBR" : $(SOURCE) $(DEP_CPP_DIBST)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\Dock.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_DOCK_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\fdialog.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Dock.obj" : $(SOURCE) $(DEP_CPP_DOCK_) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_DOCK_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\fdialog.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Dock.obj"	"$(INTDIR)\Dock.sbr" : $(SOURCE) $(DEP_CPP_DOCK_)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Draw.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_DRAW_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\XEVENTS.H"\
	

"$(INTDIR)\Draw.obj" : $(SOURCE) $(DEP_CPP_DRAW_) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_DRAW_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dock.h"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\XEVENTS.H"\
	

"$(INTDIR)\Draw.obj"	"$(INTDIR)\Draw.sbr" : $(SOURCE) $(DEP_CPP_DRAW_)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\enumx.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_ENUMX=\
	".\ENUMX.H"\
	".\precomp.h"\
	

"$(INTDIR)\enumx.obj" : $(SOURCE) $(DEP_CPP_ENUMX) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_ENUMX=\
	".\ENUMX.H"\
	

"$(INTDIR)\enumx.obj"	"$(INTDIR)\enumx.sbr" : $(SOURCE) $(DEP_CPP_ENUMX)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\FDIALOG.CPP

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_FDIAL=\
	".\DEBUG.H"\
	".\FDIALOG.H"\
	".\MAP.H"\
	".\precomp.h"\
	

"$(INTDIR)\FDIALOG.OBJ" : $(SOURCE) $(DEP_CPP_FDIAL) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_FDIAL=\
	".\DEBUG.H"\
	".\FDIALOG.H"\
	".\MAP.H"\
	

"$(INTDIR)\FDIALOG.OBJ"	"$(INTDIR)\FDIALOG.SBR" : $(SOURCE) $(DEP_CPP_FDIAL)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\filestream.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_FILES=\
	".\filestream.h"\
	".\precomp.h"\
	

"$(INTDIR)\filestream.obj" : $(SOURCE) $(DEP_CPP_FILES) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_FILES=\
	".\filestream.h"\
	

"$(INTDIR)\filestream.obj"	"$(INTDIR)\filestream.sbr" : $(SOURCE)\
 $(DEP_CPP_FILES) "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\Flicker.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_FLICK=\
	".\Flicker.h"\
	".\precomp.h"\
	

"$(INTDIR)\Flicker.obj" : $(SOURCE) $(DEP_CPP_FLICK) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_FLICK=\
	".\Flicker.h"\
	

"$(INTDIR)\Flicker.obj"	"$(INTDIR)\Flicker.sbr" : $(SOURCE) $(DEP_CPP_FLICK)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\fontholder.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_FONTH=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\fontholder.obj" : $(SOURCE) $(DEP_CPP_FONTH) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_FONTH=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\fontholder.obj"	"$(INTDIR)\fontholder.sbr" : $(SOURCE)\
 $(DEP_CPP_FONTH) "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\fregkey.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_FREGK=\
	".\FREGKEY.H"\
	".\precomp.h"\
	

"$(INTDIR)\fregkey.obj" : $(SOURCE) $(DEP_CPP_FREGK) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_FREGK=\
	".\FREGKEY.H"\
	

"$(INTDIR)\fregkey.obj"	"$(INTDIR)\fregkey.sbr" : $(SOURCE) $(DEP_CPP_FREGK)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\FWND.CPP

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_FWND_=\
	".\DEBUG.H"\
	".\FWND.H"\
	".\MAP.H"\
	".\precomp.h"\
	

"$(INTDIR)\FWND.OBJ" : $(SOURCE) $(DEP_CPP_FWND_) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_FWND_=\
	".\DEBUG.H"\
	".\FWND.H"\
	".\MAP.H"\
	

"$(INTDIR)\FWND.OBJ"	"$(INTDIR)\FWND.SBR" : $(SOURCE) $(DEP_CPP_FWND_)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\guids.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_GUIDS=\
	".\GUIDS.H"\
	".\interfaces.h"\
	".\precomp.h"\
	

"$(INTDIR)\guids.obj" : $(SOURCE) $(DEP_CPP_GUIDS) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_GUIDS=\
	".\GUIDS.H"\
	".\interfaces.h"\
	

"$(INTDIR)\guids.obj"	"$(INTDIR)\guids.sbr" : $(SOURCE) $(DEP_CPP_GUIDS)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Ibar.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_IBAR_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\fdialog.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dispids.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\filestream.h"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\StaticLink.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Ibar.obj" : $(SOURCE) $(DEP_CPP_IBAR_) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_IBAR_=\
	".\Band.h"\
	".\Bands.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\fdialog.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\Dispids.h"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\filestream.h"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\StaticLink.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Ibar.obj"	"$(INTDIR)\Ibar.sbr" : $(SOURCE) $(DEP_CPP_IBAR_)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\ipserver.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_IPSER=\
	".\CATIDREG.H"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FREGKEY.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\ipserver.obj" : $(SOURCE) $(DEP_CPP_IPSER) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_IPSER=\
	".\CATIDREG.H"\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\FREGKEY.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\ipserver.obj"	"$(INTDIR)\ipserver.sbr" : $(SOURCE) $(DEP_CPP_IPSER)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Map.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_MAP_C=\
	".\MAP.H"\
	".\precomp.h"\
	

"$(INTDIR)\Map.obj" : $(SOURCE) $(DEP_CPP_MAP_C) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_MAP_C=\
	".\MAP.H"\
	

"$(INTDIR)\Map.obj"	"$(INTDIR)\Map.sbr" : $(SOURCE) $(DEP_CPP_MAP_C)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\memstream.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_MEMST=\
	".\DEBUG.H"\
	".\memstream.h"\
	".\precomp.h"\
	

"$(INTDIR)\memstream.obj" : $(SOURCE) $(DEP_CPP_MEMST) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_MEMST=\
	".\DEBUG.H"\
	".\memstream.h"\
	

"$(INTDIR)\memstream.obj"	"$(INTDIR)\memstream.sbr" : $(SOURCE)\
 $(DEP_CPP_MEMST) "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\Miniwin.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_MINIW=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Pages.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Miniwin.obj" : $(SOURCE) $(DEP_CPP_MINIW) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_MINIW=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Pages.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Miniwin.obj"	"$(INTDIR)\Miniwin.sbr" : $(SOURCE) $(DEP_CPP_MINIW)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\opendesign.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_OPEND=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\interfaces2.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\GUIDS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\opendesign.obj" : $(SOURCE) $(DEP_CPP_OPEND) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_OPEND=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\interfaces2.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\GUIDS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\opendesign.obj"	"$(INTDIR)\opendesign.sbr" : $(SOURCE)\
 $(DEP_CPP_OPEND) "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\page.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_PAGE_=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\Page.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\page.obj" : $(SOURCE) $(DEP_CPP_PAGE_) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_PAGE_=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\Page.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\page.obj"	"$(INTDIR)\page.sbr" : $(SOURCE) $(DEP_CPP_PAGE_)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Pages.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_PAGES=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\Page.h"\
	".\Pages.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Pages.obj" : $(SOURCE) $(DEP_CPP_PAGES) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_PAGES=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\Page.h"\
	".\Pages.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Pages.obj"	"$(INTDIR)\Pages.sbr" : $(SOURCE) $(DEP_CPP_PAGES)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\pictureholder.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_PICTU=\
	".\fontholder.h"\
	".\precomp.h"\
	

"$(INTDIR)\pictureholder.obj" : $(SOURCE) $(DEP_CPP_PICTU) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_PICTU=\
	".\fontholder.h"\
	

"$(INTDIR)\pictureholder.obj"	"$(INTDIR)\pictureholder.sbr" : $(SOURCE)\
 $(DEP_CPP_PICTU) "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\Popupwin.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_POPUP=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Popupwin.obj" : $(SOURCE) $(DEP_CPP_POPUP) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_POPUP=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\MINIWIN.H"\
	".\Popupwin.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Popupwin.obj"	"$(INTDIR)\Popupwin.sbr" : $(SOURCE) $(DEP_CPP_POPUP)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\precomp.cpp
DEP_CPP_PRECO=\
	".\precomp.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

CPP_SWITCHES=/nologo /MT /W3 /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "ACTBARMAIN" /Fp"$(INTDIR)\Actbar2.pch" /Yc"precomp.h" /Fo"$(INTDIR)\\"\
 /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\precomp.obj"	"$(INTDIR)\Actbar2.pch" : $(SOURCE) $(DEP_CPP_PRECO)\
 "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

CPP_SWITCHES=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D\
 "_WINDOWS" /FR"$(INTDIR)\\" /Fp"$(INTDIR)\Actbar2.pch" /Yc"precomp.h"\
 /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\precomp.obj"	"$(INTDIR)\precomp.sbr"	"$(INTDIR)\Actbar2.pch" : \
$(SOURCE) $(DEP_CPP_PRECO) "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\Private.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_PRIVA=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\interfaces2.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Private.obj" : $(SOURCE) $(DEP_CPP_PRIVA) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_PRIVA=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\interfaces2.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Private.obj"	"$(INTDIR)\Private.sbr" : $(SOURCE) $(DEP_CPP_PRIVA)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\prj.odl
DEP_MTL_PRJ_O=\
	".\customtool.h"\
	".\customtool.idl"\
	".\Dispids.h"\
	".\Hlp\ActiveBar.hh"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

MTL_SWITCHES=/nologo /D "NDEBUG" /tlb ".\prj.tlb" /h ".\interfaces.h"\
 /mktyplib203 /win32 

".\prj.tlb"	".\interfaces.h" : $(SOURCE) $(DEP_MTL_PRJ_O) "$(OUTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

MTL_SWITCHES=/nologo /D "_DEBUG" /tlb ".\prj.tlb" /h ".\interfaces.h"\
 /mktyplib203 /win32 

".\prj.tlb"	".\interfaces.h" : $(SOURCE) $(DEP_MTL_PRJ_O) "$(OUTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\prj.rc
DEP_RSC_PRJ_R=\
	".\prj.tlb"\
	".\Res\BITMAP1.BMP"\
	".\Res\bmp00001.bmp"\
	".\Res\BMP00002.BMP"\
	".\Res\BMP00003.BMP"\
	".\Res\bmp00004.bmp"\
	".\Res\bmp00005.bmp"\
	".\Res\CHECKBMP.BMP"\
	".\Res\GRAB1.BMP"\
	".\Res\GRAB2.BMP"\
	".\Res\GRAB3.BMP"\
	".\Res\GRAB4.BMP"\
	".\Res\ICON1.ICO"\
	".\Res\ICON2.ICO"\
	".\Res\IDCC_DRA.CUR"\
	".\Res\MENUSCRO.BMP"\
	".\Res\OVERFLOW.BMP"\
	".\Res\SUBMENUB.BMP"\
	".\Res\SYSCLOSE.BMP"\
	".\Res\SYSMAX.BMP"\
	".\Res\SYSMINIM.BMP"\
	".\Res\SYSRESTO.BMP"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"


"$(INTDIR)\prj.res" : $(SOURCE) $(DEP_RSC_PRJ_R) "$(INTDIR)" ".\prj.tlb"
	$(RSC) /l 0x409 /fo"$(INTDIR)\prj.res" /i ".\Release" /d "NDEBUG" $(SOURCE)


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"


"$(INTDIR)\prj.res" : $(SOURCE) $(DEP_RSC_PRJ_R) "$(INTDIR)" ".\prj.tlb"
	$(RSC) /l 0x409 /fo"$(INTDIR)\prj.res" /i ".\Debug" /d "_DEBUG" $(SOURCE)


!ENDIF 

SOURCE=.\proppage.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_PROPP=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\PROPPAGE.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\proppage.obj" : $(SOURCE) $(DEP_CPP_PROPP) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_PROPP=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\PROPPAGE.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\proppage.obj"	"$(INTDIR)\proppage.sbr" : $(SOURCE) $(DEP_CPP_PROPP)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Crt\QSORT.C
DEP_CPP_QSORT=\
	{$(INCLUDE)}"cruntime.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

CPP_SWITCHES=/nologo /MT /W3 /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "ACTBARMAIN" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\QSORT.OBJ" : $(SOURCE) $(DEP_CPP_QSORT) "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

CPP_SWITCHES=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D\
 "_WINDOWS" /FR"$(INTDIR)\\" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\QSORT.OBJ"	"$(INTDIR)\QSORT.SBR" : $(SOURCE) $(DEP_CPP_QSORT)\
 "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\Crt\rand.c

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_RAND_=\
	{$(INCLUDE)}"cruntime.h"\
	{$(INCLUDE)}"mtdll.h"\
	
CPP_SWITCHES=/nologo /MT /W3 /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "ACTBARMAIN" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\rand.obj" : $(SOURCE) $(DEP_CPP_RAND_) "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_RAND_=\
	{$(INCLUDE)}"cruntime.h"\
	{$(INCLUDE)}"mtdll.h"\
	
CPP_SWITCHES=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D\
 "_WINDOWS" /FR"$(INTDIR)\\" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\rand.obj"	"$(INTDIR)\rand.sbr" : $(SOURCE) $(DEP_CPP_RAND_)\
 "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\returnbool.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_RETUR=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\returnbool.obj" : $(SOURCE) $(DEP_CPP_RETUR) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_RETUR=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\returnbool.obj"	"$(INTDIR)\returnbool.sbr" : $(SOURCE)\
 $(DEP_CPP_RETUR) "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\returnstring.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_RETURN=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\returnstring.obj" : $(SOURCE) $(DEP_CPP_RETURN) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_RETURN=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\returnstring.obj"	"$(INTDIR)\returnstring.sbr" : $(SOURCE)\
 $(DEP_CPP_RETURN) "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\StaticLink.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_STATI=\
	".\FWND.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\StaticLink.h"\
	

"$(INTDIR)\StaticLink.obj" : $(SOURCE) $(DEP_CPP_STATI) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_STATI=\
	".\FWND.H"\
	".\MAP.H"\
	".\StaticLink.h"\
	

"$(INTDIR)\StaticLink.obj"	"$(INTDIR)\StaticLink.sbr" : $(SOURCE)\
 $(DEP_CPP_STATI) "$(INTDIR)" "$(INTDIR)\Actbar2.pch"


!ENDIF 

SOURCE=.\support.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_SUPPO=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\support.obj" : $(SOURCE) $(DEP_CPP_SUPPO) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_SUPPO=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\GLOBALS.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\support.obj"	"$(INTDIR)\support.sbr" : $(SOURCE) $(DEP_CPP_SUPPO)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\tool.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_TOOL_=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\CBList.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\DIBSTUFF.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\Pages.h"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\tool.obj" : $(SOURCE) $(DEP_CPP_TOOL_) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_TOOL_=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\Btabs.h"\
	".\CBList.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\DIBSTUFF.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\Pages.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\tool.obj"	"$(INTDIR)\tool.sbr" : $(SOURCE) $(DEP_CPP_TOOL_)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Tool2.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_TOOL2=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\CBList.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Tppopup.h"\
	

"$(INTDIR)\Tool2.obj" : $(SOURCE) $(DEP_CPP_TOOL2) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_TOOL2=\
	".\Band.h"\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\CBList.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	".\Tppopup.h"\
	

"$(INTDIR)\Tool2.obj"	"$(INTDIR)\Tool2.sbr" : $(SOURCE) $(DEP_CPP_TOOL2)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\DESIGNER\ToolDrop.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_TOOLD=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\debug.h"\
	".\DESIGNER\precomp.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\ToolDrop.obj" : $(SOURCE) $(DEP_CPP_TOOLD) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_TOOLD=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\designer\debug.h"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	

"$(INTDIR)\ToolDrop.obj"	"$(INTDIR)\ToolDrop.sbr" : $(SOURCE) $(DEP_CPP_TOOLD)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\Tools.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_TOOLS=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Tools.obj" : $(SOURCE) $(DEP_CPP_TOOLS) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_TOOLS=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\customproxy.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\ERRORS.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tools.h"\
	

"$(INTDIR)\Tools.obj"	"$(INTDIR)\Tools.sbr" : $(SOURCE) $(DEP_CPP_TOOLS)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Tooltip.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_TOOLT=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tooltip.h"\
	

"$(INTDIR)\Tooltip.obj" : $(SOURCE) $(DEP_CPP_TOOLT) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_TOOLT=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tooltip.h"\
	

"$(INTDIR)\Tooltip.obj"	"$(INTDIR)\Tooltip.sbr" : $(SOURCE) $(DEP_CPP_TOOLT)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Tppopup.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_TPPOP=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\CBList.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tppopup.h"\
	

"$(INTDIR)\Tppopup.obj" : $(SOURCE) $(DEP_CPP_TPPOP) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_TPPOP=\
	".\Bar.h"\
	".\bitmapmng.h"\
	".\CBList.h"\
	".\DDASYNC.H"\
	".\DEBUG.H"\
	".\DESIGNER\TOOLDROP.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\Flicker.h"\
	".\fontholder.h"\
	".\FWND.H"\
	".\interfaces.h"\
	".\IPSERVER.H"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\Tool.h"\
	".\Tppopup.h"\
	

"$(INTDIR)\Tppopup.obj"	"$(INTDIR)\Tppopup.sbr" : $(SOURCE) $(DEP_CPP_TPPOP)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 

SOURCE=.\Crt\WCSCAT.C
DEP_CPP_WCSCA=\
	{$(INCLUDE)}"cruntime.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

CPP_SWITCHES=/nologo /MT /W3 /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "ACTBARMAIN" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\WCSCAT.OBJ" : $(SOURCE) $(DEP_CPP_WCSCA) "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

CPP_SWITCHES=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D\
 "_WINDOWS" /FR"$(INTDIR)\\" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\WCSCAT.OBJ"	"$(INTDIR)\WCSCAT.SBR" : $(SOURCE) $(DEP_CPP_WCSCA)\
 "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\Crt\WCSCMP.C
DEP_CPP_WCSCM=\
	{$(INCLUDE)}"cruntime.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

CPP_SWITCHES=/nologo /MT /W3 /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "ACTBARMAIN" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\WCSCMP.OBJ" : $(SOURCE) $(DEP_CPP_WCSCM) "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

CPP_SWITCHES=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D\
 "_WINDOWS" /FR"$(INTDIR)\\" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\WCSCMP.OBJ"	"$(INTDIR)\WCSCMP.SBR" : $(SOURCE) $(DEP_CPP_WCSCM)\
 "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\Crt\WCSLEN.C
DEP_CPP_WCSLE=\
	{$(INCLUDE)}"cruntime.h"\
	

!IF  "$(CFG)" == "Actbar - Win32 Release"

CPP_SWITCHES=/nologo /MT /W3 /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D\
 "ACTBARMAIN" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\WCSLEN.OBJ" : $(SOURCE) $(DEP_CPP_WCSLE) "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

CPP_SWITCHES=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D\
 "_WINDOWS" /FR"$(INTDIR)\\" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\WCSLEN.OBJ"	"$(INTDIR)\WCSLEN.SBR" : $(SOURCE) $(DEP_CPP_WCSLE)\
 "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\xevents.cpp

!IF  "$(CFG)" == "Actbar - Win32 Release"

DEP_CPP_XEVEN=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\MAP.H"\
	".\precomp.h"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\XEVENTS.H"\
	

"$(INTDIR)\xevents.obj" : $(SOURCE) $(DEP_CPP_XEVEN) "$(INTDIR)"\
 "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ELSEIF  "$(CFG)" == "Actbar - Win32 Debug"

DEP_CPP_XEVEN=\
	".\DEBUG.H"\
	".\ENUMX.H"\
	".\FDIALOG.H"\
	".\fontholder.h"\
	".\interfaces.h"\
	".\MAP.H"\
	".\returnbool.h"\
	".\returnstring.h"\
	".\SUPPORT.H"\
	".\XEVENTS.H"\
	

"$(INTDIR)\xevents.obj"	"$(INTDIR)\xevents.sbr" : $(SOURCE) $(DEP_CPP_XEVEN)\
 "$(INTDIR)" "$(INTDIR)\Actbar2.pch" ".\interfaces.h"


!ENDIF 


!ENDIF 

