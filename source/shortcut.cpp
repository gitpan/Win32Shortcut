/*
 * Shortcut.CPP
 * 15 Jan 97 by Aldo Calpini <dada@divinf.it>
 *
 * XS interface to the Win32 IShellLink Interface
 * based on Registry.CPP written by Jesse Dougherty
 *
 * Version: 0.03 07 Apr 97
 *
 */

#define  WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <shlobj.h>
#include <objbase.h>

#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"
// Section for the constant definitions.
#define CROAK croak
#define MAX_LENGTH 2048
#define TMPBUFSZ 1024

#define SETIV(index,value) sv_setiv(ST(index), value)
#define SETPV(index,string) sv_setpv(ST(index), string)
#define SETPVN(index, buffer, length) sv_setpvn(ST(index), (char*)buffer, length)

#define NEW(x,v,n,t)  (v = (t*)safemalloc((MEM_SIZE)((n) * sizeof(t))))
#define PERLSvIV(sv) (SvIOK(sv) ? SvIVX(sv) : sv_2iv(sv))
#define PERLSvPV(sv, lp) (SvPOK(sv) ? ((lp = SvCUR(sv)), SvPVX(sv)) : sv_2pv(sv, &lp))

#define PERLPUSHMARK(p) if (++markstack_ptr == markstack_max)	\
			markstack_grow();			\
		    *markstack_ptr = (p) - stack_base

#define PERLXPUSHs(s)	do {\
 		if (stack_max - sp < 1) {\
			    sp = stack_grow(sp, sp, 1);\
			}\
  (*++sp = (s)); } while (0)



DWORD
constant(CPerl* pPerl,char *name, int arg)
{
    errno = 0;
    switch (*name) {
    case 'A':
		break;
    case 'B':
		break;
	case 'C':
		break;
    case 'D':
		break;
    case 'E':
		break;
    case 'F':
		break;
    case 'G':
		break;
    case 'H':
		break;
    case 'I':
		break;
    case 'J':
		break;
    case 'K':
		break;
    case 'L':
		break;
    case 'M':
		break;
    case 'N':
		break;
    case 'O':
		break;
    case 'P':
		break;
    case 'Q':
		break;
    case 'R':
		break;
    case 'S':
		if(strncmp(name, "SLGP_", 5) == 0)
			switch(name[5]) {	
			case 'S':
				if (strEQ(name, "SLGP_SHORTPATH"))
					#ifdef SLGP_SHORTPATH
						return SLGP_SHORTPATH;
					#else
						goto not_there;
					#endif
				break;
			case 'U':
				if (strEQ(name, "SLGP_UNCPRIORITY"))
					#ifdef SLGP_UNCPRIORITY
						return SLGP_UNCPRIORITY;
					#else
						goto not_there;
					#endif
				break;
			}
		if(strncmp(name, "SW_", 3) == 0)
			switch(name[3]) {	
			case 'H':
				if (strEQ(name, "SW_HIDE"))
					#ifdef SW_HIDE
						return SW_HIDE;
					#else
						goto not_there;
					#endif
				break;
			case 'M':
				if (strEQ(name, "SW_MINIMIZE"))
					#ifdef SW_MINIMIZE
						return SW_MINIMIZE;
					#else
						goto not_there;
					#endif
				break;
			case 'R':
				if (strEQ(name, "SW_RESTORE"))
					#ifdef SW_RESTORE
						return SW_RESTORE;
					#else
						goto not_there;
					#endif
				break;
			case 'S':
				if (strEQ(name, "SW_SHOW"))
					#ifdef SW_SHOW
						return SW_SHOW;
					#else
						goto not_there;
					#endif
				if (strEQ(name, "SW_SHOWMAXIMIZED"))
					#ifdef SW_SHOWMAXIMIZED
						return SW_SHOWMAXIMIZED;
					#else
						goto not_there;
					#endif
				if (strEQ(name, "SW_SHOWMINIMIZED"))
					#ifdef SW_SHOWMINIMIZED
						return SW_SHOWMINIMIZED;
					#else
						goto not_there;
					#endif
				if (strEQ(name, "SW_SHOWMINNOACTIVE"))
					#ifdef SW_SHOWMINNOACTIVE
						return SW_SHOWMINNOACTIVE;
					#else
						goto not_there;
					#endif
				if (strEQ(name, "SW_SHOWNA"))
					#ifdef SW_SHOWNA
						return SW_SHOWNA;
					#else
						goto not_there;
					#endif
				if (strEQ(name, "SW_SHOWNOACTIVE"))
					#ifdef SW_SHOWNOACTIVE
						return SW_SHOWNOACTIVE;
					#else
						goto not_there;
					#endif
				if (strEQ(name, "SW_SHOWNORMAL"))
					#ifdef SW_SHOWNORMAL
						return SW_SHOWNORMAL;
					#else
						goto not_there;
					#endif
				break;
			}
		break;
    case 'T':
		break;
    case 'U':
		break;
    case 'V':
		break;
    case 'W':
		break;
    case 'X':
		break;
    case 'Y':
		break;
    case 'Z':
		break;
    }
    errno = EINVAL;
    return 0;

not_there:
    errno = ENOENT;
    return 0;
}


XS(XS_NT__Shortcut_constant)
{
    dXSARGS;

    if (items != 2) {
		croak("Usage: Win32::Shortcut::constant(name,arg)");
    }
    {
		char *	name = (char *)SvPV(ST(0),na);
		int	arg = (int)SvIV(ST(1));
		DWORD RETVAL;

		RETVAL = constant(pPerl,name, arg);
		ST(0) = sv_newmortal();
		sv_setiv(ST(0), (IV)RETVAL);
    }
    XSRETURN(1);
}



XS(XS_NT__Shortcut_Instance) {

	dXSARGS;

	HRESULT hres;
	IShellLink* ilink;
        
	hres=CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
                              IID_IShellLink, (void **) &ilink);
	if(SUCCEEDED(hres)) {
		
		IPersistFile* ifile;

		hres=ilink->QueryInterface(IID_IPersistFile, (void **) &ifile);
		if(SUCCEEDED(hres)) {
			ST(0)=sv_2mortal(newSViv((DWORD) ilink));
			ST(1)=sv_2mortal(newSViv((DWORD) ifile));
			XSRETURN(2);
		}
		XSRETURN_NO;
	}
	XSRETURN_NO;
}


XS(XS_NT__Shortcut_Release) {

	dXSARGS;

	if(items != 2) {
		CROAK("usage: _Release($ilink, $ifile);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));

	ifile->Release();
	ilink->Release();

	XSRETURN_YES;
}


XS(XS_NT__Shortcut_SetDescription) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _SetDescription($ilink, $ifile, $description);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));

	hres=ilink->SetDescription((LPCSTR) PERLSvPV(ST(2),na));
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_GetDescription) {

	dXSARGS;

	if(items != 2) {
		CROAK("usage: _GetDescription($ilink, $ifile);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	char *description = (char *)safemalloc(1024);

	hres=ilink->GetDescription((LPSTR) description, 1024);
	if(SUCCEEDED(hres)) {
		ST(0)=sv_2mortal(newSVpv((LPSTR) description,0));
		safefree((char *)description);
		XSRETURN(1);
	} else {
		safefree((char *)description);
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_SetPath) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _SetPath($ilink, $ifile, $path);\n");
	}


	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));

	hres=ilink->SetPath((LPCSTR) PERLSvPV(ST(2),na));
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_GetPath) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _GetPath($ilink, $ifile, $flags);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	char *path = (char *)safemalloc(MAX_PATH);
	WIN32_FIND_DATA file;

	hres=ilink->GetPath((LPSTR) path, MAX_PATH, &file, (DWORD) PERLSvIV(ST(2)));
	if(SUCCEEDED(hres)) {
		ST(0)=sv_2mortal(newSVpv((LPSTR) path,0));
		safefree((char *)path);
		XSRETURN(1);
	} else {
		safefree((char *)path);
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_SetArguments) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _SetArguments($ilink, $ifile, $arguments);\n");
	}


	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));

	hres=ilink->SetArguments((LPCSTR) PERLSvPV(ST(2),na));
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}

XS(XS_NT__Shortcut_GetArguments) {

	dXSARGS;

	if(items != 2) {
		CROAK("usage: _GetArguments($ilink, $ifile);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	char *arguments = (char *)safemalloc(1024);

	hres=ilink->GetArguments((LPSTR) arguments, 1024);
	if(SUCCEEDED(hres)) {
		ST(0)=sv_2mortal(newSVpv((LPSTR) arguments,0));
		safefree((char *)arguments);
		XSRETURN(1);
	} else {
		safefree((char *)arguments);
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_SetWorkingDirectory) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _SetWorkingDirectory($ilink, $ifile, $dir);\n");
	}


	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));

	hres=ilink->SetWorkingDirectory((LPCSTR) PERLSvPV(ST(2),na));
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_GetWorkingDirectory) {

	dXSARGS;

	if(items != 2) {
		CROAK("usage: _GetWorkingDirectory($ilink, $ifile);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	char *dir = (char *)safemalloc(MAX_PATH);

	hres=ilink->GetWorkingDirectory((LPSTR) dir, MAX_PATH);
	if(SUCCEEDED(hres)) {
		ST(0)=sv_2mortal(newSVpv((LPSTR) dir,0));
		safefree((char *)dir);
		XSRETURN(1);
	} else {
		safefree((char *)dir);
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_SetShowCmd) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _SetShowCmd($ilink, $ifile, $flag);\n");
	}


	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));

	hres=ilink->SetShowCmd((int) PERLSvIV(ST(2)));
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_GetShowCmd) {

	dXSARGS;

	if(items != 2) {
		CROAK("usage: _GetShowCmd($ilink, $ifile);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	int show;

	hres=ilink->GetShowCmd(&show);
	if(SUCCEEDED(hres)) {
		ST(0)=sv_2mortal(newSViv(show));
		XSRETURN(1);
	} else {
		XSRETURN_NO;
	}
}

XS(XS_NT__Shortcut_SetHotkey) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _SetHotkey($ilink, $ifile, $hotkey);\n");
	}


	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));

	hres=ilink->SetHotkey((unsigned short) PERLSvIV(ST(2)));
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_GetHotkey) {

	dXSARGS;

	if(items != 2) {
		CROAK("usage: _GetHotkey($ilink, $ifile);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	unsigned short hotkey;

	hres=ilink->GetHotkey(&hotkey);
	if(SUCCEEDED(hres)) {
		ST(0)=sv_2mortal(newSViv(hotkey));
		XSRETURN(1);
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_SetIconLocation) {

	dXSARGS;

	if(items != 4) {
		CROAK("usage: _SetIconLocation($ilink, $ifile, $location, $number);\n");
	}


	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));

	hres=ilink->SetIconLocation((char *) PERLSvPV(ST(2),na),
								(int) PERLSvIV(ST(3)));
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_GetIconLocation) {

	dXSARGS;

	if(items != 2) {
		CROAK("usage: _GetIconLocation($ilink, $ifile);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	int number;
	char *location = (char *)safemalloc(MAX_PATH);

	hres=ilink->GetIconLocation((LPSTR)location, MAX_PATH, &number);
	if(SUCCEEDED(hres)) {
		// [dada] does actually returns nothing?
		// printf("_GetIconLocation: location=\"%s\",%d\n",location,number);
		ST(0)=sv_2mortal(newSVpv((LPSTR)location,0));
		ST(1)=sv_2mortal(newSViv(number));
		safefree((char *)location);
		XSRETURN(2);
	} else {
		safefree((char *)location);
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_Resolve) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _Resolve($ilink, $ifile, $flags);\n");
	}

	HRESULT hres;
	IShellLink* ilink = (IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile = (IPersistFile*) PERLSvIV(ST(1));

	// [dada] hwnd=NULL, not sure about it...
	hres = ilink->Resolve(NULL, PERLSvIV(ST(2)));

	if(SUCCEEDED(hres)) {
		char *path = (char *)safemalloc(MAX_PATH);
		WIN32_FIND_DATA file;

		hres = ilink->GetPath((LPSTR)path, MAX_PATH, &file, 0);
		if(SUCCEEDED(hres)) {
			ST(0) = sv_2mortal(newSVpv((LPSTR)path, 0));
			safefree((char *)path);
			XSRETURN(1);
		} else {
			safefree((char *)path);
			XSRETURN_NO;
		}
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_Save) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: _Save($ilink, $ifile, $filename);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	unsigned short filename[MAX_PATH];

	MultiByteToWideChar(CP_ACP, 0, 
						(LPSTR) PERLSvPV(ST(2),na),
						-1, (unsigned short *) filename, MAX_PATH);
		
	hres=ifile->Save((unsigned short *) filename, TRUE);
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}

XS(XS_NT__Shortcut_Load) {

	dXSARGS;

	if(items != 3) {
		CROAK("usage: Load($ilink, $ifile, $filename);\n");
	}

	HRESULT hres;
	IShellLink* ilink=(IShellLink*) PERLSvIV(ST(0));
	IPersistFile* ifile=(IPersistFile*) PERLSvIV(ST(1));
	unsigned short filename[MAX_PATH];

	MultiByteToWideChar(CP_ACP, 0, 
						(LPSTR) PERLSvPV(ST(2),na),
						-1, (unsigned short *) filename, MAX_PATH);
	
	hres = ifile->Load((unsigned short *) filename, STGM_READ);
	if(SUCCEEDED(hres)) {
		XSRETURN_YES;
	} else {
		XSRETURN_NO;
	}
}


XS(XS_NT__Shortcut_Exit) {

	dXSARGS;

	CoUninitialize();

	XSRETURN(1);

}
 

XS(boot_Win32__Shortcut)
{
	dXSARGS;
	char* file = __FILE__;

	newXS("Win32::Shortcut::constant", XS_NT__Shortcut_constant, file);

	newXS("Win32::Shortcut::_Instance", XS_NT__Shortcut_Instance, file);
	newXS("Win32::Shortcut::_Release", XS_NT__Shortcut_Release, file);
	newXS("Win32::Shortcut::_SetPath", XS_NT__Shortcut_SetPath, file);
	newXS("Win32::Shortcut::_GetPath", XS_NT__Shortcut_GetPath, file);
	newXS("Win32::Shortcut::_SetDescription", XS_NT__Shortcut_SetDescription, file);
	newXS("Win32::Shortcut::_GetDescription", XS_NT__Shortcut_GetDescription, file);
	newXS("Win32::Shortcut::_SetWorkingDirectory", XS_NT__Shortcut_SetWorkingDirectory, file);
	newXS("Win32::Shortcut::_GetWorkingDirectory", XS_NT__Shortcut_GetWorkingDirectory, file);
	newXS("Win32::Shortcut::_SetArguments", XS_NT__Shortcut_SetArguments, file);
	newXS("Win32::Shortcut::_GetArguments", XS_NT__Shortcut_GetArguments, file);
	newXS("Win32::Shortcut::_SetShowCmd", XS_NT__Shortcut_SetShowCmd, file);
	newXS("Win32::Shortcut::_GetShowCmd", XS_NT__Shortcut_GetShowCmd, file);
	newXS("Win32::Shortcut::_SetHotkey", XS_NT__Shortcut_SetHotkey, file);
	newXS("Win32::Shortcut::_GetHotkey", XS_NT__Shortcut_GetHotkey, file);
	newXS("Win32::Shortcut::_SetIconLocation", XS_NT__Shortcut_SetIconLocation, file);
	newXS("Win32::Shortcut::_GetIconLocation", XS_NT__Shortcut_GetIconLocation, file);
	newXS("Win32::Shortcut::_Resolve", XS_NT__Shortcut_Resolve, file);
	newXS("Win32::Shortcut::_Save", XS_NT__Shortcut_Save, file);
	newXS("Win32::Shortcut::_Load", XS_NT__Shortcut_Load, file);
	newXS("Win32::Shortcut::_Exit", XS_NT__Shortcut_Exit, file);

	CoInitialize(NULL);

	ST(0) = &sv_yes;
	XSRETURN(1);
}
