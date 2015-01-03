#ifdef _MERGE_PROXYSTUB // proxy stub DLL の結合

#define REGISTER_PROXY_DLL //DllRegisterServer、他

#define _WIN32_WINNT 0x0500	//WinNT 4.0 および DCOM を持つ Win95 用
#define USE_STUBLESS_PROXY	//MIDL のオプションで /Oicf を指定した場合のみ定義

#pragma comment(lib, "rpcns4.lib")
#pragma comment(lib, "rpcrt4.lib")

#define ENTRY_PREFIX	Prx

#include "dlldata.c"
#include "RScript22_p.c"

#endif //_MERGE_PROXYSTUB
