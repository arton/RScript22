#ifdef _MERGE_PROXYSTUB // proxy stub DLL �̌���

#define REGISTER_PROXY_DLL //DllRegisterServer�A��

#define _WIN32_WINNT 0x0500	//WinNT 4.0 ����� DCOM ������ Win95 �p
#define USE_STUBLESS_PROXY	//MIDL �̃I�v�V������ /Oicf ���w�肵���ꍇ�̂ݒ�`

#pragma comment(lib, "rpcns4.lib")
#pragma comment(lib, "rpcrt4.lib")

#define ENTRY_PREFIX	Prx

#include "dlldata.c"
#include "RScript22_p.c"

#endif //_MERGE_PROXYSTUB
