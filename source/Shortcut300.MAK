# Microsoft Developer Studio Generated NMAKE File, Format Version 4.10
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

!IF "$(CFG)" == ""
CFG=SHORTCUT - Win32 Debug
!MESSAGE No configuration specified.  Defaulting to SHORTCUT - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "SHORTCUT - Win32 Release" && "$(CFG)" !=\
 "SHORTCUT - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "Shortcut.mak" CFG="SHORTCUT - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "SHORTCUT - Win32 Release" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE "SHORTCUT - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 
################################################################################
# Begin Project
# PROP Target_Last_Scanned "SHORTCUT - Win32 Debug"
RSC=rc.exe
MTL=mktyplib.exe
CPP=cl.exe

!IF  "$(CFG)" == "SHORTCUT - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "WinRel"
# PROP BASE Intermediate_Dir "WinRel"
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "WinRel"
# PROP Intermediate_Dir "WinRel"
OUTDIR=.\WinRel
INTDIR=.\WinRel

ALL : "$(OUTDIR)\Shortcut.pll"

CLEAN : 
	-@erase "$(INTDIR)\SHORTCUT.OBJ"
	-@erase "$(OUTDIR)\Shortcut.exp"
	-@erase "$(OUTDIR)\Shortcut.lib"
	-@erase "$(OUTDIR)\Shortcut.pll"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /FR /YX /c
# ADD CPP /nologo /MT /W3 /GX /O2 /I "../../../" /I "../../../inc" /D "NDEBUG" /D "WIN32" /D "_WINDOWS" /D "MSWIN32" /D "PERL_OBJECT" /D "EMBED" /YX /c
# SUBTRACT CPP /Fr
CPP_PROJ=/nologo /MT /W3 /GX /O2 /I "../../../" /I "../../../inc" /D "NDEBUG"\
 /D "WIN32" /D "_WINDOWS" /D "MSWIN32" /D "PERL_OBJECT" /D "EMBED"\
 /Fp"$(INTDIR)/Shortcut.pch" /YX /Fo"$(INTDIR)/" /c 
CPP_OBJS=.\WinRel/
CPP_SBRS=.\.
# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /win32
MTL_PROJ=/nologo /D "NDEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/Shortcut.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 ..\..\..\Release\perl300.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /base:0x10120000 /subsystem:windows /dll /machine:I386 /out:"WinRel\Shortcut.pll"
LINK32_FLAGS=..\..\..\Release\perl300.lib kernel32.lib user32.lib gdi32.lib\
 winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib\
 uuid.lib odbc32.lib odbccp32.lib /nologo /base:0x10120000 /subsystem:windows\
 /dll /incremental:no /pdb:"$(OUTDIR)/Shortcut.pdb" /machine:I386\
 /def:".\SHORTCUT.DEF" /out:"$(OUTDIR)/Shortcut.pll"\
 /implib:"$(OUTDIR)/Shortcut.lib" 
DEF_FILE= \
	".\SHORTCUT.DEF"
LINK32_OBJS= \
	"$(INTDIR)\SHORTCUT.OBJ"

"$(OUTDIR)\Shortcut.pll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "SHORTCUT - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "WinDebug"
# PROP BASE Intermediate_Dir "WinDebug"
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "WinDebug"
# PROP Intermediate_Dir "WinDebug"
OUTDIR=.\WinDebug
INTDIR=.\WinDebug

ALL : "..\..\..\..\lib\auto\win32\shortcut\Shortcut.pll"

CLEAN : 
	-@erase "$(INTDIR)\SHORTCUT.OBJ"
	-@erase "$(INTDIR)\vc40.idb"
	-@erase "$(INTDIR)\vc40.pdb"
	-@erase "$(OUTDIR)\Shortcut.exp"
	-@erase "$(OUTDIR)\Shortcut.lib"
	-@erase "$(OUTDIR)\Shortcut.pdb"
	-@erase "..\..\..\..\lib\auto\win32\shortcut\Shortcut.ilk"
	-@erase "..\..\..\..\lib\auto\win32\shortcut\Shortcut.pll"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

# ADD BASE CPP /nologo /MT /W3 /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /FR /YX /c
# ADD CPP /nologo /MTd /w /W0 /Gm /GX /Zi /Od /I "../../../" /I "../../../inc" /D "_DEBUG" /D "WIN32" /D "_WINDOWS" /D "MSWIN32" /D "PERL_OBJECT" /D "EMBED" /YX /c
# SUBTRACT CPP /Fr
CPP_PROJ=/nologo /MTd /w /W0 /Gm /GX /Zi /Od /I "../../../" /I "../../../inc"\
 /D "_DEBUG" /D "WIN32" /D "_WINDOWS" /D "MSWIN32" /D "PERL_OBJECT" /D "EMBED"\
 /Fp"$(INTDIR)/Shortcut.pch" /YX /Fo"$(INTDIR)/" /Fd"$(INTDIR)/" /c 
CPP_OBJS=.\WinDebug/
CPP_SBRS=.\.
# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o"$(OUTDIR)/Shortcut.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 ..\..\..\Release\perl300.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /base:0x10120000 /subsystem:windows /dll /debug /machine:I386 /out:"f:\pg\perl\lib\auto\win32\shortcut\Shortcut.pll"
LINK32_FLAGS=..\..\..\Release\perl300.lib kernel32.lib user32.lib gdi32.lib\
 winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib\
 uuid.lib odbc32.lib odbccp32.lib /nologo /base:0x10120000 /subsystem:windows\
 /dll /incremental:yes /pdb:"$(OUTDIR)/Shortcut.pdb" /debug /machine:I386\
 /def:".\SHORTCUT.DEF" /out:"f:\pg\perl\lib\auto\win32\shortcut\Shortcut.pll"\
 /implib:"$(OUTDIR)/Shortcut.lib" 
DEF_FILE= \
	".\SHORTCUT.DEF"
LINK32_OBJS= \
	"$(INTDIR)\SHORTCUT.OBJ"

"..\..\..\..\lib\auto\win32\shortcut\Shortcut.pll" : "$(OUTDIR)" $(DEF_FILE)\
 $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ENDIF 

.c{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

.cpp{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

.cxx{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

.c{$(CPP_SBRS)}.sbr:
   $(CPP) $(CPP_PROJ) $<  

.cpp{$(CPP_SBRS)}.sbr:
   $(CPP) $(CPP_PROJ) $<  

.cxx{$(CPP_SBRS)}.sbr:
   $(CPP) $(CPP_PROJ) $<  

################################################################################
# Begin Target

# Name "SHORTCUT - Win32 Release"
# Name "SHORTCUT - Win32 Debug"

!IF  "$(CFG)" == "SHORTCUT - Win32 Release"

!ELSEIF  "$(CFG)" == "SHORTCUT - Win32 Debug"

!ENDIF 

################################################################################
# Begin Source File

SOURCE=.\SHORTCUT.DEF

!IF  "$(CFG)" == "SHORTCUT - Win32 Release"

!ELSEIF  "$(CFG)" == "SHORTCUT - Win32 Debug"

!ENDIF 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\SHORTCUT.CPP

!IF  "$(CFG)" == "SHORTCUT - Win32 Release"

DEP_CPP_SHORT=\
	"..\..\..\av.h"\
	"..\..\..\build.h"\
	"..\..\..\config.h"\
	"..\..\..\cop.h"\
	"..\..\..\cv.h"\
	"..\..\..\dosish.h"\
	"..\..\..\embed.h"\
	"..\..\..\form.h"\
	"..\..\..\gv.h"\
	"..\..\..\handy.h"\
	"..\..\..\hv.h"\
	"..\..\..\IPerlSup.h"\
	"..\..\..\mg.h"\
	"..\..\..\nostdio.h"\
	"..\..\..\nt.h"\
	"..\..\..\NTXSub.h"\
	"..\..\..\objpp.h"\
	"..\..\..\op.h"\
	"..\..\..\opcode.h"\
	"..\..\..\perlio.h"\
	"..\..\..\perllib.h"\
	"..\..\..\perlsdio.h"\
	"..\..\..\perlsfio.h"\
	"..\..\..\perly.h"\
	"..\..\..\pp.h"\
	"..\..\..\proto.h"\
	"..\..\..\regexp.h"\
	"..\..\..\scope.h"\
	"..\..\..\sv.h"\
	"..\..\..\unixish.h"\
	"..\..\..\util.h"\
	".\../../../EXTERN.h"\
	".\../../../inc\dirent.h"\
	".\../../../perl.h"\
	".\../../../XSUB.h"\
	{$(INCLUDE)}"\sys\STAT.H"\
	{$(INCLUDE)}"\sys\TYPES.H"\
	
NODEP_CPP_SHORT=\
	"..\..\..\os2ish.h"\
	"..\..\..\plan9\plan9ish.h"\
	"..\..\..\vmsish.h"\
	

"$(INTDIR)\SHORTCUT.OBJ" : $(SOURCE) $(DEP_CPP_SHORT) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "SHORTCUT - Win32 Debug"

DEP_CPP_SHORT=\
	"..\..\..\av.h"\
	"..\..\..\build.h"\
	"..\..\..\config.h"\
	"..\..\..\cop.h"\
	"..\..\..\cv.h"\
	"..\..\..\dosish.h"\
	"..\..\..\embed.h"\
	"..\..\..\form.h"\
	"..\..\..\gv.h"\
	"..\..\..\handy.h"\
	"..\..\..\hv.h"\
	"..\..\..\IPerlSup.h"\
	"..\..\..\mg.h"\
	"..\..\..\nostdio.h"\
	"..\..\..\nt.h"\
	"..\..\..\NTXSub.h"\
	"..\..\..\op.h"\
	"..\..\..\opcode.h"\
	"..\..\..\perlio.h"\
	"..\..\..\perllib.h"\
	"..\..\..\perlsdio.h"\
	"..\..\..\perlsfio.h"\
	"..\..\..\perly.h"\
	"..\..\..\proto.h"\
	"..\..\..\regexp.h"\
	"..\..\..\scope.h"\
	"..\..\..\sv.h"\
	"..\..\..\unixish.h"\
	"..\..\..\util.h"\
	".\../../../EXTERN.h"\
	".\../../../inc\dirent.h"\
	".\../../../perl.h"\
	".\../../../XSUB.h"\
	{$(INCLUDE)}"\sys\STAT.H"\
	{$(INCLUDE)}"\sys\TYPES.H"\
	
NODEP_CPP_SHORT=\
	"..\..\..\os2ish.h"\
	"..\..\..\plan9\plan9ish.h"\
	"..\..\..\vmsish.h"\
	

"$(INTDIR)\SHORTCUT.OBJ" : $(SOURCE) $(DEP_CPP_SHORT) "$(INTDIR)"


!ENDIF 

# End Source File
# End Target
# End Project
################################################################################
