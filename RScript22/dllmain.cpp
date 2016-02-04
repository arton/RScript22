/*
 * Copyright(c) 2014 arton
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 */

#include "stdafx.h"
#include "resource.h"
#include "RScript22_i.h"
#include "dllmain.h"
#include "xdlldata.h"

#include "RubyScript.h"
#include "Rubyize.h"

#if defined(RUBY_2_2)
#define RSCRIPT_VERSION L"2.2"
#define CLSID_RUBYSCRIPT  L"{456A3763-90A4-4F2A-BFF1-4B773C1056EC}"
#define CLSID_RUBYIZE  L"{0BCFF05A-C2BF-4CB2-A778-3428A8E85A21}"
#else
#error "no ruby version defined"
#endif

CRScript22Module _AtlModule;
GIT _init_git;

OBJECT_ENTRY_AUTO(CLSID_Rubyize, CRubyize)
OBJECT_ENTRY_AUTO(CLSID_RubyScript, CRubyScript)

_ATL_REGMAP_ENTRY CRubyize::RegEntries[] = {
	{ L"RSCRIPT_VERSION", RSCRIPT_VERSION },
	{ L"CLSID", CLSID_RUBYIZE },
	{ NULL, NULL }
};
_ATL_REGMAP_ENTRY CRubyScript::RegEntries[] = {
	{ L"RSCRIPT_VERSION", RSCRIPT_VERSION },
	{ L"CLSID", CLSID_RUBYSCRIPT },
        { L"SCPATH", NULL },
	{ NULL, NULL }
};

HRESULT WINAPI CRubyScript::UpdateRegistry(BOOL bRegister)
{
    UINT (WINAPI *getsysdirfun)(LPWSTR, UINT) = GetSystemWow64Directory;
    UINT len = GetSystemWow64Directory(NULL, 0);
    if (!len)
    {
        getsysdirfun = GetSystemDirectory;
        len = GetSystemDirectory(NULL, 0);
    }
    RegEntries[2].szData = new WCHAR[len];
    getsysdirfun(const_cast<LPWSTR>(RegEntries[2].szData), len);
    return _AtlModule.CAtlModule::UpdateRegistryFromResourceS(IDR_RUBYSCRIPT, bRegister, RegEntries);
}
HRESULT WINAPI CRubyize::UpdateRegistry(BOOL bRegister)
{
    return _AtlModule.CAtlModule::UpdateRegistryFromResourceS(IDR_RUBYIZE, bRegister, RegEntries);
}


// DLL エントリ ポイント
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(hInstance, dwReason, lpReserved))
		return FALSE;
#endif
	hInstance;
	return _AtlModule.DllMain(dwReason, lpReserved); 
}
