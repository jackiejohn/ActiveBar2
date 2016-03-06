# Microsoft Developer Studio Generated NMAKE File, Based on Designer.dsp
!IF "$(CFG)" == ""
CFG=Designer - Win32 Release
!MESSAGE No configuration specified. Defaulting to Designer - Win32 Release.
!ENDIF 

!IF "$(CFG)" != "Designer - Win32 Release" && "$(CFG)" !=\
 "Designer - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Designer.mak" CFG="Designer - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Designer - Win32 Release" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Designer - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 

!IF  "$(CFG)" == "Designer - Win32 Release"

OUTDIR=.\Release
INTDIR=.\Release
# Begin Custom Macros
OutDir=.\.\Release
# End Custom Macros

!IF "$(RECURSE)" == "0" 

ALL : "$(OUTDIR)\Designer.dll"

!ELSE 

ALL : "$(OUTDIR)\Designer.dll"

!ENDIF 

CLEAN :
	-@erase "$(INTDIR)\APPPREF.OBJ"
	-@erase "$(INTDIR)\Browser.obj"
	-@erase "$(INTDIR)\CategoryMgr.obj"
	-@erase "$(INTDIR)\catidreg.obj"
	-@erase "$(INTDIR)\debug.obj"
	-@erase "$(INTDIR)\Designer.obj"
	-@erase "$(INTDIR)\Designer.pch"
	-@erase "$(INTDIR)\DesignerImpl.obj"
	-@erase "$(INTDIR)\Dialogs.obj"
	-@erase "$(INTDIR)\DragDrop.obj"
	-@erase "$(INTDIR)\enumx.obj"
	-@erase "$(INTDIR)\FDialog.obj"
	-@erase "$(INTDIR)\fregkey.obj"
	-@erase "$(INTDIR)\FWND.OBJ"
	-@erase "$(INTDIR)\GdiUtil.obj"
	-@erase "$(INTDIR)\Globals.obj"
	-@erase "$(INTDIR)\guids.obj"
	-@erase "$(INTDIR)\Iconedit.obj"
	-@erase "$(INTDIR)\ipserver.obj"
	-@erase "$(INTDIR)\Map.obj"
	-@erase "$(INTDIR)\precomp.obj"
	-@erase "$(INTDIR)\prj.res"
	-@erase "$(INTDIR)\proppage.obj"
	-@erase "$(INTDIR)\StringUtil.obj"
	-@erase "$(INTDIR)\support.obj"
	-@erase "$(INTDIR)\Utility.obj"
	-@erase "$(INTDIR)\vc50.idb"
	-@erase "$(INTDIR)\xevents.obj"
	-@erase "$(OUTDIR)\Designer.dll"
	-@erase "$(OUTDIR)\Designer.exp"
	-@erase "$(OUTDIR)\Designer.lib"
	-@erase ".\DesignerInterfaces.h"
	-@erase ".\prj.tlb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS"\
 /Fp"$(INTDIR)\Designer.pch" /Yu"precomp.h" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\"\
 /FD /c 
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
BSC32_FLAGS=/nologo /o"$(OUTDIR)\Designer.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib\
 advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib\
 odbccp32.lib /nologo /entry:"DllMain" /subsystem:windows /dll /incremental:no\
 /pdb:"$(OUTDIR)\Designer.pdb" /machine:I386 /nodefaultlib /def:".\prj.def"\
 /out:"$(OUTDIR)\Designer.dll" /implib:"$(OUTDIR)\Designer.lib" 
DEF_FILE= \
	".\prj.def"
LINK32_OBJS= \
	"$(INTDIR)\APPPREF.OBJ" \
	"$(INTDIR)\Browser.obj" \
	"$(INTDIR)\CategoryMgr.obj" \
	"$(INTDIR)\catidreg.obj" \
	"$(INTDIR)\debug.obj" \
	"$(INTDIR)\Designer.obj" \
	"$(INTDIR)\DesignerImpl.obj" \
	"$(INTDIR)\Dialogs.obj" \
	"$(INTDIR)\DragDrop.obj" \
	"$(INTDIR)\enumx.obj" \
	"$(INTDIR)\FDialog.obj" \
	"$(INTDIR)\fregkey.obj" \
	"$(INTDIR)\FWND.OBJ" \
	"$(INTDIR)\GdiUtil.obj" \
	"$(INTDIR)\Globals.obj" \
	"$(INTDIR)\guids.obj" \
	"$(INTDIR)\Iconedit.obj" \
	"$(INTDIR)\ipserver.obj" \
	"$(INTDIR)\Map.obj" \
	"$(INTDIR)\precomp.obj" \
	"$(INTDIR)\prj.res" \
	"$(INTDIR)\proppage.obj" \
	"$(INTDIR)\StringUtil.obj" \
	"$(INTDIR)\support.obj" \
	"$(INTDIR)\Utility.obj" \
	"$(INTDIR)\xevents.obj"

"$(OUTDIR)\Designer.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

OUTDIR=.\Debug
INTDIR=.\Debug
# Begin Custom Macros
OutDir=.\.\Debug
# End Custom Macros

!IF "$(RECURSE)" == "0" 

ALL : "$(OUTDIR)\Designer.dll" "$(OUTDIR)\Designer.bsc"

!ELSE 

ALL : "$(OUTDIR)\Designer.dll" "$(OUTDIR)\Designer.bsc"

!ENDIF 

CLEAN :
	-@erase "$(INTDIR)\APPPREF.OBJ"
	-@erase "$(INTDIR)\APPPREF.SBR"
	-@erase "$(INTDIR)\Browser.obj"
	-@erase "$(INTDIR)\Browser.sbr"
	-@erase "$(INTDIR)\CategoryMgr.obj"
	-@erase "$(INTDIR)\CategoryMgr.sbr"
	-@erase "$(INTDIR)\catidreg.obj"
	-@erase "$(INTDIR)\catidreg.sbr"
	-@erase "$(INTDIR)\debug.obj"
	-@erase "$(INTDIR)\debug.sbr"
	-@erase "$(INTDIR)\Designer.obj"
	-@erase "$(INTDIR)\Designer.pch"
	-@erase "$(INTDIR)\Designer.sbr"
	-@erase "$(INTDIR)\DesignerImpl.obj"
	-@erase "$(INTDIR)\DesignerImpl.sbr"
	-@erase "$(INTDIR)\Dialogs.obj"
	-@erase "$(INTDIR)\Dialogs.sbr"
	-@erase "$(INTDIR)\DragDrop.obj"
	-@erase "$(INTDIR)\DragDrop.sbr"
	-@erase "$(INTDIR)\enumx.obj"
	-@erase "$(INTDIR)\enumx.sbr"
	-@erase "$(INTDIR)\FDialog.obj"
	-@erase "$(INTDIR)\FDialog.sbr"
	-@erase "$(INTDIR)\fregkey.obj"
	-@erase "$(INTDIR)\fregkey.sbr"
	-@erase "$(INTDIR)\FWND.OBJ"
	-@erase "$(INTDIR)\FWND.SBR"
	-@erase "$(INTDIR)\GdiUtil.obj"
	-@erase "$(INTDIR)\GdiUtil.sbr"
	-@erase "$(INTDIR)\Globals.obj"
	-@erase "$(INTDIR)\Globals.sbr"
	-@erase "$(INTDIR)\guids.obj"
	-@erase "$(INTDIR)\guids.sbr"
	-@erase "$(INTDIR)\Iconedit.obj"
	-@erase "$(INTDIR)\Iconedit.sbr"
	-@erase "$(INTDIR)\ipserver.obj"
	-@erase "$(INTDIR)\ipserver.sbr"
	-@erase "$(INTDIR)\Map.obj"
	-@erase "$(INTDIR)\Map.sbr"
	-@erase "$(INTDIR)\precomp.obj"
	-@erase "$(INTDIR)\precomp.sbr"
	-@erase "$(INTDIR)\prj.res"
	-@erase "$(INTDIR)\proppage.obj"
	-@erase "$(INTDIR)\proppage.sbr"
	-@erase "$(INTDIR)\StringUtil.obj"
	-@erase "$(INTDIR)\StringUtil.sbr"
	-@erase "$(INTDIR)\support.obj"
	-@erase "$(INTDIR)\support.sbr"
	-@erase "$(INTDIR)\Utility.obj"
	-@erase "$(INTDIR)\Utility.sbr"
	-@erase "$(INTDIR)\vc50.idb"
	-@erase "$(INTDIR)\vc50.pdb"
	-@erase "$(INTDIR)\xevents.obj"
	-@erase "$(INTDIR)\xevents.sbr"
	-@erase "$(OUTDIR)\Designer.bsc"
	-@erase "$(OUTDIR)\Designer.dll"
	-@erase "$(OUTDIR)\Designer.exp"
	-@erase "$(OUTDIR)\Designer.ilk"
	-@erase "$(OUTDIR)\Designer.lib"
	-@erase "$(OUTDIR)\Designer.pdb"
	-@erase ".\DesignerInterfaces.h"
	-@erase ".\prj.tlb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS"\
 /FR"$(INTDIR)\\" /Fp"$(INTDIR)\Designer.pch" /Yu"precomp.h" /Fo"$(INTDIR)\\"\
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
BSC32_FLAGS=/nologo /o"$(OUTDIR)\Designer.bsc" 
BSC32_SBRS= \
	"$(INTDIR)\APPPREF.SBR" \
	"$(INTDIR)\Browser.sbr" \
	"$(INTDIR)\CategoryMgr.sbr" \
	"$(INTDIR)\catidreg.sbr" \
	"$(INTDIR)\debug.sbr" \
	"$(INTDIR)\Designer.sbr" \
	"$(INTDIR)\DesignerImpl.sbr" \
	"$(INTDIR)\Dialogs.sbr" \
	"$(INTDIR)\DragDrop.sbr" \
	"$(INTDIR)\enumx.sbr" \
	"$(INTDIR)\FDialog.sbr" \
	"$(INTDIR)\fregkey.sbr" \
	"$(INTDIR)\FWND.SBR" \
	"$(INTDIR)\GdiUtil.sbr" \
	"$(INTDIR)\Globals.sbr" \
	"$(INTDIR)\guids.sbr" \
	"$(INTDIR)\Iconedit.sbr" \
	"$(INTDIR)\ipserver.sbr" \
	"$(INTDIR)\Map.sbr" \
	"$(INTDIR)\precomp.sbr" \
	"$(INTDIR)\proppage.sbr" \
	"$(INTDIR)\StringUtil.sbr" \
	"$(INTDIR)\support.sbr" \
	"$(INTDIR)\Utility.sbr" \
	"$(INTDIR)\xevents.sbr"

"$(OUTDIR)\Designer.bsc" : "$(OUTDIR)" $(BSC32_SBRS)
    $(BSC32) @<<
  $(BSC32_FLAGS) $(BSC32_SBRS)
<<

LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib\
 advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib\
 odbccp32.lib comctl32.lib /nologo /subsystem:windows /dll /incremental:yes\
 /pdb:"$(OUTDIR)\Designer.pdb" /debug /machine:I386 /def:".\prj.def"\
 /out:"$(OUTDIR)\Designer.dll" /implib:"$(OUTDIR)\Designer.lib" 
DEF_FILE= \
	".\prj.def"
LINK32_OBJS= \
	"$(INTDIR)\APPPREF.OBJ" \
	"$(INTDIR)\Browser.obj" \
	"$(INTDIR)\CategoryMgr.obj" \
	"$(INTDIR)\catidreg.obj" \
	"$(INTDIR)\debug.obj" \
	"$(INTDIR)\Designer.obj" \
	"$(INTDIR)\DesignerImpl.obj" \
	"$(INTDIR)\Dialogs.obj" \
	"$(INTDIR)\DragDrop.obj" \
	"$(INTDIR)\enumx.obj" \
	"$(INTDIR)\FDialog.obj" \
	"$(INTDIR)\fregkey.obj" \
	"$(INTDIR)\FWND.OBJ" \
	"$(INTDIR)\GdiUtil.obj" \
	"$(INTDIR)\Globals.obj" \
	"$(INTDIR)\guids.obj" \
	"$(INTDIR)\Iconedit.obj" \
	"$(INTDIR)\ipserver.obj" \
	"$(INTDIR)\Map.obj" \
	"$(INTDIR)\precomp.obj" \
	"$(INTDIR)\prj.res" \
	"$(INTDIR)\proppage.obj" \
	"$(INTDIR)\StringUtil.obj" \
	"$(INTDIR)\support.obj" \
	"$(INTDIR)\Utility.obj" \
	"$(INTDIR)\xevents.obj"

"$(OUTDIR)\Designer.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ENDIF 


!IF "$(CFG)" == "Designer - Win32 Release" || "$(CFG)" ==\
 "Designer - Win32 Debug"
SOURCE=.\APPPREF.CPP

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_APPPR=\
	".\APPPREF.H"\
	".\fregkey.h"\
	".\precomp.h"\
	

"$(INTDIR)\APPPREF.OBJ" : $(SOURCE) $(DEP_CPP_APPPR) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_APPPR=\
	".\APPPREF.H"\
	".\fregkey.h"\
	

"$(INTDIR)\APPPREF.OBJ"	"$(INTDIR)\APPPREF.SBR" : $(SOURCE) $(DEP_CPP_APPPR)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\Browser.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_BROWS=\
	".\Browser.h"\
	".\debug.h"\
	".\enumx.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Map.h"\
	".\precomp.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\Browser.obj" : $(SOURCE) $(DEP_CPP_BROWS) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_BROWS=\
	".\Browser.h"\
	".\debug.h"\
	".\enumx.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Map.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\Browser.obj"	"$(INTDIR)\Browser.sbr" : $(SOURCE) $(DEP_CPP_BROWS)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\CategoryMgr.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_CATEG=\
	".\APPPREF.H"\
	".\Browser.h"\
	".\CategoryMgr.h"\
	".\Dialogs.h"\
	".\DragDrop.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\Globals.h"\
	".\interfaces.h"\
	".\Map.h"\
	".\precomp.h"\
	".\StringUtil.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\CategoryMgr.obj" : $(SOURCE) $(DEP_CPP_CATEG) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_CATEG=\
	".\APPPREF.H"\
	".\Browser.h"\
	".\CategoryMgr.h"\
	".\Dialogs.h"\
	".\DragDrop.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\Globals.h"\
	".\interfaces.h"\
	".\Map.h"\
	".\StringUtil.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\CategoryMgr.obj"	"$(INTDIR)\CategoryMgr.sbr" : $(SOURCE)\
 $(DEP_CPP_CATEG) "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\catidreg.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_CATID=\
	".\catidreg.h"\
	".\precomp.h"\
	

"$(INTDIR)\catidreg.obj" : $(SOURCE) $(DEP_CPP_CATID) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_CATID=\
	".\catidreg.h"\
	

"$(INTDIR)\catidreg.obj"	"$(INTDIR)\catidreg.sbr" : $(SOURCE) $(DEP_CPP_CATID)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\debug.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_DEBUG=\
	".\debug.h"\
	".\enumx.h"\
	".\ipserver.h"\
	".\precomp.h"\
	".\support.h"\
	

"$(INTDIR)\debug.obj" : $(SOURCE) $(DEP_CPP_DEBUG) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_DEBUG=\
	".\debug.h"\
	".\enumx.h"\
	".\ipserver.h"\
	".\support.h"\
	

"$(INTDIR)\debug.obj"	"$(INTDIR)\debug.sbr" : $(SOURCE) $(DEP_CPP_DEBUG)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\Designer.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_DESIG=\
	".\Browser.h"\
	".\CategoryMgr.h"\
	".\debug.h"\
	".\Designer.h"\
	".\DesignerImpl.h"\
	".\DesignerInterfaces.h"\
	".\Dialogs.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\fontholder.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\interfaces.h"\
	".\ipserver.h"\
	".\Map.h"\
	".\precomp.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\Designer.obj" : $(SOURCE) $(DEP_CPP_DESIG) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch" ".\DesignerInterfaces.h"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_DESIG=\
	".\Browser.h"\
	".\CategoryMgr.h"\
	".\debug.h"\
	".\Designer.h"\
	".\DesignerImpl.h"\
	".\DesignerInterfaces.h"\
	".\Dialogs.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\fontholder.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\interfaces.h"\
	".\ipserver.h"\
	".\Map.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\Designer.obj"	"$(INTDIR)\Designer.sbr" : $(SOURCE) $(DEP_CPP_DESIG)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch" ".\DesignerInterfaces.h"


!ENDIF 

SOURCE=.\DesignerImpl.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_DESIGN=\
	".\Browser.h"\
	".\CategoryMgr.h"\
	".\DesignerImpl.h"\
	".\Dialogs.h"\
	".\DragDrop.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\Fwnd.h"\
	".\interfaces.h"\
	".\Map.h"\
	".\precomp.h"\
	".\streams.h"\
	".\StringUtil.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\DesignerImpl.obj" : $(SOURCE) $(DEP_CPP_DESIGN) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_DESIGN=\
	".\Browser.h"\
	".\CategoryMgr.h"\
	".\DesignerImpl.h"\
	".\Dialogs.h"\
	".\DragDrop.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\Fwnd.h"\
	".\interfaces.h"\
	".\Map.h"\
	".\streams.h"\
	".\StringUtil.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\DesignerImpl.obj"	"$(INTDIR)\DesignerImpl.sbr" : $(SOURCE)\
 $(DEP_CPP_DESIGN) "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\Dialogs.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_DIALO=\
	".\APPPREF.H"\
	".\Dialogs.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Iconedit.h"\
	".\interfaces.h"\
	".\Map.h"\
	".\precomp.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\Dialogs.obj" : $(SOURCE) $(DEP_CPP_DIALO) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_DIALO=\
	".\APPPREF.H"\
	".\Dialogs.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Iconedit.h"\
	".\interfaces.h"\
	".\Map.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\Dialogs.obj"	"$(INTDIR)\Dialogs.sbr" : $(SOURCE) $(DEP_CPP_DIALO)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\DragDrop.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_DRAGD=\
	".\CategoryMgr.h"\
	".\DesignerImpl.h"\
	".\DragDrop.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\Fwnd.h"\
	".\interfaces.h"\
	".\Map.h"\
	".\precomp.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\DragDrop.obj" : $(SOURCE) $(DEP_CPP_DRAGD) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_DRAGD=\
	".\CategoryMgr.h"\
	".\DesignerImpl.h"\
	".\DragDrop.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\Fwnd.h"\
	".\interfaces.h"\
	".\Map.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\DragDrop.obj"	"$(INTDIR)\DragDrop.sbr" : $(SOURCE) $(DEP_CPP_DRAGD)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\enumx.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_ENUMX=\
	".\enumx.h"\
	".\precomp.h"\
	

"$(INTDIR)\enumx.obj" : $(SOURCE) $(DEP_CPP_ENUMX) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_ENUMX=\
	".\enumx.h"\
	

"$(INTDIR)\enumx.obj"	"$(INTDIR)\enumx.sbr" : $(SOURCE) $(DEP_CPP_ENUMX)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\FDialog.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_FDIAL=\
	".\APPPREF.H"\
	".\debug.h"\
	".\Fdialog.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Map.h"\
	".\precomp.h"\
	".\Utility.h"\
	

"$(INTDIR)\FDialog.obj" : $(SOURCE) $(DEP_CPP_FDIAL) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_FDIAL=\
	".\APPPREF.H"\
	".\debug.h"\
	".\Fdialog.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Map.h"\
	".\Utility.h"\
	

"$(INTDIR)\FDialog.obj"	"$(INTDIR)\FDialog.sbr" : $(SOURCE) $(DEP_CPP_FDIAL)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\fregkey.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_FREGK=\
	".\fregkey.h"\
	".\precomp.h"\
	

"$(INTDIR)\fregkey.obj" : $(SOURCE) $(DEP_CPP_FREGK) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_FREGK=\
	".\fregkey.h"\
	

"$(INTDIR)\fregkey.obj"	"$(INTDIR)\fregkey.sbr" : $(SOURCE) $(DEP_CPP_FREGK)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\FWND.CPP

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_FWND_=\
	".\debug.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Map.h"\
	".\precomp.h"\
	".\Utility.h"\
	

"$(INTDIR)\FWND.OBJ" : $(SOURCE) $(DEP_CPP_FWND_) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_FWND_=\
	".\debug.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Map.h"\
	".\Utility.h"\
	

"$(INTDIR)\FWND.OBJ"	"$(INTDIR)\FWND.SBR" : $(SOURCE) $(DEP_CPP_FWND_)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\GdiUtil.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_GDIUT=\
	".\APPPREF.H"\
	".\enumx.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Map.h"\
	".\precomp.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\GdiUtil.obj" : $(SOURCE) $(DEP_CPP_GDIUT) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_GDIUT=\
	".\APPPREF.H"\
	".\enumx.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Map.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\GdiUtil.obj"	"$(INTDIR)\GdiUtil.sbr" : $(SOURCE) $(DEP_CPP_GDIUT)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\Globals.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_GLOBA=\
	".\APPPREF.H"\
	".\debug.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Iconedit.h"\
	".\Map.h"\
	".\precomp.h"\
	".\Utility.h"\
	

"$(INTDIR)\Globals.obj" : $(SOURCE) $(DEP_CPP_GLOBA) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_GLOBA=\
	".\APPPREF.H"\
	".\debug.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Iconedit.h"\
	".\Map.h"\
	".\Utility.h"\
	

"$(INTDIR)\Globals.obj"	"$(INTDIR)\Globals.sbr" : $(SOURCE) $(DEP_CPP_GLOBA)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\guids.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_GUIDS=\
	".\DesignerInterfaces.h"\
	".\interfaces.h"\
	".\precomp.h"\
	

"$(INTDIR)\guids.obj" : $(SOURCE) $(DEP_CPP_GUIDS) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch" ".\DesignerInterfaces.h"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_GUIDS=\
	".\DesignerInterfaces.h"\
	".\interfaces.h"\
	

"$(INTDIR)\guids.obj"	"$(INTDIR)\guids.sbr" : $(SOURCE) $(DEP_CPP_GUIDS)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch" ".\DesignerInterfaces.h"


!ENDIF 

SOURCE=.\Iconedit.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_ICONE=\
	".\APPPREF.H"\
	".\debug.h"\
	".\DesignerInterfaces.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Iconedit.h"\
	".\interfaces.h"\
	".\ipserver.h"\
	".\Map.h"\
	".\precomp.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\Iconedit.obj" : $(SOURCE) $(DEP_CPP_ICONE) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch" ".\DesignerInterfaces.h"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_ICONE=\
	".\APPPREF.H"\
	".\debug.h"\
	".\DesignerInterfaces.h"\
	".\enumx.h"\
	".\Fdialog.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\Globals.h"\
	".\Iconedit.h"\
	".\interfaces.h"\
	".\ipserver.h"\
	".\Map.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\Iconedit.obj"	"$(INTDIR)\Iconedit.sbr" : $(SOURCE) $(DEP_CPP_ICONE)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch" ".\DesignerInterfaces.h"


!ENDIF 

SOURCE=.\ipserver.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_IPSER=\
	".\catidreg.h"\
	".\debug.h"\
	".\DesignerInterfaces.h"\
	".\enumx.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\ipserver.h"\
	".\Map.h"\
	".\precomp.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\ipserver.obj" : $(SOURCE) $(DEP_CPP_IPSER) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch" ".\DesignerInterfaces.h"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_IPSER=\
	".\catidreg.h"\
	".\debug.h"\
	".\DesignerInterfaces.h"\
	".\enumx.h"\
	".\fregkey.h"\
	".\Fwnd.h"\
	".\GdiUtil.h"\
	".\ipserver.h"\
	".\Map.h"\
	".\support.h"\
	".\Utility.h"\
	

"$(INTDIR)\ipserver.obj"	"$(INTDIR)\ipserver.sbr" : $(SOURCE) $(DEP_CPP_IPSER)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch" ".\DesignerInterfaces.h"


!ENDIF 

SOURCE=.\Map.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_MAP_C=\
	".\Map.h"\
	".\precomp.h"\
	

"$(INTDIR)\Map.obj" : $(SOURCE) $(DEP_CPP_MAP_C) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_MAP_C=\
	".\Map.h"\
	

"$(INTDIR)\Map.obj"	"$(INTDIR)\Map.sbr" : $(SOURCE) $(DEP_CPP_MAP_C)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\precomp.cpp
DEP_CPP_PRECO=\
	".\precomp.h"\
	

!IF  "$(CFG)" == "Designer - Win32 Release"

CPP_SWITCHES=/nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS"\
 /Fp"$(INTDIR)\Designer.pch" /Yc"precomp.h" /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\"\
 /FD /c 

"$(INTDIR)\precomp.obj"	"$(INTDIR)\Designer.pch" : $(SOURCE) $(DEP_CPP_PRECO)\
 "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

CPP_SWITCHES=/nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D\
 "_WINDOWS" /FR"$(INTDIR)\\" /Fp"$(INTDIR)\Designer.pch" /Yc"precomp.h"\
 /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

"$(INTDIR)\precomp.obj"	"$(INTDIR)\precomp.sbr"	"$(INTDIR)\Designer.pch" : \
$(SOURCE) $(DEP_CPP_PRECO) "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\prj.odl
DEP_MTL_PRJ_O=\
	".\dispids.h"\
	

!IF  "$(CFG)" == "Designer - Win32 Release"

MTL_SWITCHES=/nologo /D "NDEBUG" /tlb ".\prj.tlb" /h ".\DesignerInterfaces.h"\
 /mktyplib203 /win32 

".\prj.tlb"	".\DesignerInterfaces.h" : $(SOURCE) $(DEP_MTL_PRJ_O) "$(OUTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

MTL_SWITCHES=/nologo /D "_DEBUG" /tlb ".\prj.tlb" /h ".\DesignerInterfaces.h"\
 /mktyplib203 /win32 

".\prj.tlb"	".\DesignerInterfaces.h" : $(SOURCE) $(DEP_MTL_PRJ_O) "$(OUTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\prj.rc
DEP_RSC_PRJ_R=\
	".\prj.tlb"\
	".\Res\bmp00001.bmp"\
	".\Res\category.bmp"\
	".\Res\cur00001.cur"\
	".\RES\cur00002.cur"\
	".\res\cursor1.cur"\
	".\RES\CYCLE.BMP"\
	".\res\filedeme.bmp"\
	".\RES\HATCHBMP.BMP"\
	".\Res\heirarch.bmp"\
	".\res\iconedit.bmp"\
	".\res\idcc_buc.cur"\
	".\res\idcc_dra.cur"\
	".\res\idcc_hai.cur"\
	".\res\idcc_mar.cur"\
	".\res\idcc_pen.cur"\
	".\res\idcc_pip.cur"\
	".\res\idcc_siz.cur"\
	".\res\idcc_spr.cur"\
	".\res\mngcate.bmp"\
	".\res\stateand.bmp"\
	".\res\toolbar1.bmp"\
	".\resource.hm"\
	

!IF  "$(CFG)" == "Designer - Win32 Release"


"$(INTDIR)\prj.res" : $(SOURCE) $(DEP_RSC_PRJ_R) "$(INTDIR)" ".\prj.tlb"
	$(RSC) /l 0x409 /fo"$(INTDIR)\prj.res" /i ".\Release" /d "NDEBUG" $(SOURCE)


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"


"$(INTDIR)\prj.res" : $(SOURCE) $(DEP_RSC_PRJ_R) "$(INTDIR)" ".\prj.tlb"
	$(RSC) /l 0x409 /fo"$(INTDIR)\prj.res" /i ".\Debug" /d "_DEBUG" $(SOURCE)


!ENDIF 

SOURCE=.\proppage.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_PROPP=\
	".\debug.h"\
	".\enumx.h"\
	".\ipserver.h"\
	".\precomp.h"\
	".\proppage.h"\
	".\support.h"\
	

"$(INTDIR)\proppage.obj" : $(SOURCE) $(DEP_CPP_PROPP) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_PROPP=\
	".\debug.h"\
	".\enumx.h"\
	".\ipserver.h"\
	".\proppage.h"\
	".\support.h"\
	

"$(INTDIR)\proppage.obj"	"$(INTDIR)\proppage.sbr" : $(SOURCE) $(DEP_CPP_PROPP)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\StringUtil.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_STRIN=\
	".\precomp.h"\
	".\StringUtil.h"\
	

"$(INTDIR)\StringUtil.obj" : $(SOURCE) $(DEP_CPP_STRIN) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_STRIN=\
	".\StringUtil.h"\
	

"$(INTDIR)\StringUtil.obj"	"$(INTDIR)\StringUtil.sbr" : $(SOURCE)\
 $(DEP_CPP_STRIN) "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\support.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_SUPPO=\
	".\debug.h"\
	".\enumx.h"\
	".\ipserver.h"\
	".\precomp.h"\
	".\support.h"\
	

"$(INTDIR)\support.obj" : $(SOURCE) $(DEP_CPP_SUPPO) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_SUPPO=\
	".\debug.h"\
	".\enumx.h"\
	".\ipserver.h"\
	".\support.h"\
	

"$(INTDIR)\support.obj"	"$(INTDIR)\support.sbr" : $(SOURCE) $(DEP_CPP_SUPPO)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\Utility.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_UTILI=\
	".\precomp.h"\
	".\Utility.h"\
	

"$(INTDIR)\Utility.obj" : $(SOURCE) $(DEP_CPP_UTILI) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_UTILI=\
	".\Utility.h"\
	

"$(INTDIR)\Utility.obj"	"$(INTDIR)\Utility.sbr" : $(SOURCE) $(DEP_CPP_UTILI)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 

SOURCE=.\xevents.cpp

!IF  "$(CFG)" == "Designer - Win32 Release"

DEP_CPP_XEVEN=\
	".\debug.h"\
	".\enumx.h"\
	".\precomp.h"\
	".\support.h"\
	".\xevents.h"\
	

"$(INTDIR)\xevents.obj" : $(SOURCE) $(DEP_CPP_XEVEN) "$(INTDIR)"\
 "$(INTDIR)\Designer.pch"


!ELSEIF  "$(CFG)" == "Designer - Win32 Debug"

DEP_CPP_XEVEN=\
	".\debug.h"\
	".\enumx.h"\
	".\support.h"\
	".\xevents.h"\
	

"$(INTDIR)\xevents.obj"	"$(INTDIR)\xevents.sbr" : $(SOURCE) $(DEP_CPP_XEVEN)\
 "$(INTDIR)" "$(INTDIR)\Designer.pch"


!ENDIF 


!ENDIF 

