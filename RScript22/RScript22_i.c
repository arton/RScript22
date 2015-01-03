

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Sat Jan 03 17:22:02 2015
 */
/* Compiler settings for RScript22.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, LIBID_RScript22Lib,0x10D4CB67,0xC6A3,0x46CF,0xBE,0x81,0xB1,0x58,0x99,0xAB,0xA6,0x02);


MIDL_DEFINE_GUID(CLSID, CLSID_RubyScript,0x456A3763,0x90A4,0x4F2A,0xBF,0xF1,0x4B,0x77,0x3C,0x10,0x56,0xEC);


MIDL_DEFINE_GUID(IID, IID_IRubyize,0x0A4CBEBD,0xC46B,0x4A7C,0xA1,0xE2,0xAD,0x47,0x4C,0x33,0x0C,0x7A);


MIDL_DEFINE_GUID(CLSID, CLSID_Rubyize,0x0BCFF05A,0xC2BF,0x4CB2,0xA7,0x78,0x34,0x28,0xA8,0xE8,0x5A,0x21);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



