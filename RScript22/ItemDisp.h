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

#if !defined(ITEM_DISP_H)
#define ITEM_DISP_H
#include "giplip.h"

class CItemDisp
{
public:
    CItemDisp(DWORD dwFlag = 0) : m_dwFlag(dwFlag), m_pDisp(NULL), m_pDispEx(NULL)
    {
    }
    ~CItemDisp()
    {
        Empty();
    }
    inline DWORD GetFlag() const { return m_dwFlag; }
    inline void SetFlag(DWORD dw) { m_dwFlag = dw; }
    inline bool IsSource() const { return (m_dwFlag | SCRIPTITEM_ISSOURCE) ? true : false; }
    IDispatch* GetDispatch()
    {
        if (m_pGIPDispEx.IsOK())
	{
            IDispatchEx* p;
            HRESULT hr = m_pGIPDispEx.Localize(&p);
            return p;
	}
	if (m_pGIPDisp.IsOK())
	{
            IDispatch* p;
            HRESULT hr = m_pGIPDisp.Localize(&p);
            return p;
	}
	if (m_pDispEx)
	{
            m_pDispEx->AddRef();
            return m_pDispEx;
	}
	if (m_pDisp)
	{
            m_pDisp->AddRef();
            return m_pDisp;
	}
	return NULL;
    }
    inline bool IsOK()
    {
        return (m_pGIPDisp.IsOK() || m_pGIPDispEx.IsOK() || m_pDispEx || m_pDisp);
    }
    IDispatch* GetDispatch(IActiveScriptSite* pSite, LPCOLESTR pstrName, bool fSameApt)
    {
        IDispatchEx* pDispEx = m_pDispEx;
        IDispatch* pDisp = m_pDisp;
        ATLTRACE(_("CItemDisp::GetDispatch in Thread:%08X name=%ls sameptr=%d\n"), GetCurrentThreadId(), pstrName, fSameApt);
        if (fSameApt)
        {
            if (pDisp == NULL && pDispEx == NULL)
            {
                ATLTRACE(_("CItemDisp::GetDispatch in Thread:%08X\n"), GetCurrentThreadId());
                IUnknown* pUnk = NULL;
                ITypeInfo* pTypeInfo = NULL;
                HRESULT hr = pSite->GetItemInfo(pstrName, SCRIPTINFO_IUNKNOWN, &pUnk, &pTypeInfo);
                if (hr == S_OK)
                {
                    if (pTypeInfo) pTypeInfo->Release();
                    if (pUnk->QueryInterface(IID_IDispatchEx, (void**)&pDispEx) == S_OK)
                    {
                        HRESULT hr = m_pGIPDispEx.Globalize(pDispEx);
                        ATLTRACE(_T("Globalize Item = %08X\n"), hr);
                        m_pDispEx = pDispEx;
                    }
                    else
                    {
                        pDispEx = NULL;
                        if (pUnk->QueryInterface(IID_IDispatch, (void**)&pDisp) == S_OK)
                        {
                            HRESULT hr = m_pGIPDisp.Globalize(pDisp);
                            ATLTRACE(_T("Globalize Item = %08X\n"), hr);
                            m_pDisp = pDisp;
                        }
                    }
                    pUnk->Release();
                }
            }
            if (pDispEx)
            {
                pDispEx->AddRef();
            }
            else if (pDisp)
            {
                pDisp->AddRef();
            }
        }
        else
        {
            if (m_pGIPDisp.IsOK() == false && m_pGIPDispEx.IsOK() == false)
            {
                ATLTRACE(_("CItemDisp::GetDispatch in Thread:%08X\n"), GetCurrentThreadId());
                IUnknown* pUnk = NULL;
                ITypeInfo* pTypeInfo = NULL;
                HRESULT hr = pSite->GetItemInfo(pstrName, SCRIPTINFO_IUNKNOWN, &pUnk, &pTypeInfo);
                if (hr == S_OK)
                {
                    if (pTypeInfo) pTypeInfo->Release();
                    if (pUnk->QueryInterface(IID_IDispatchEx, (void**)&pDispEx) != S_OK)
                    {
                        pDispEx = NULL;
                        if (pUnk->QueryInterface(IID_IDispatch, (void**)&pDisp) != S_OK)
                        {
                            pDisp = NULL;
                        }
                    }
                    pUnk->Release();
                }
            }
            if (m_pGIPDispEx.IsOK())
            {
                IDispatchEx* p;
                HRESULT hr = m_pGIPDispEx.Localize(&p);
                ATLTRACE(_("Localize DispEx = %08X\n"), hr);
                return p;
            }
            if (m_pGIPDisp.IsOK())
            {
                IDispatch* p;
                HRESULT hr = m_pGIPDisp.Localize(&p);
                ATLTRACE(_("Localize Disp = %08X\n"), hr);
                return p;
            }
        }
        return (pDispEx) ? pDispEx : pDisp;
    }

    void Empty()
    {
        if (m_pDispEx)
        {
            m_pDispEx->Release();
            m_pDispEx = NULL;
        }
        if (m_pDisp) 
        {
            m_pDisp->Release();
            m_pDisp = NULL;
        }
        if (m_pGIPDispEx.IsOK())
        {
            m_pGIPDispEx.Unglobalize();
        }
        if (m_pGIPDisp.IsOK())
        {
            m_pGIPDisp.Unglobalize();
        }
    }
private:
    DWORD m_dwFlag;
    IDispatch* m_pDisp;
    IDispatchEx* m_pDispEx;
    GIP(IDispatch) m_pGIPDisp;
    GIP(IDispatchEx) m_pGIPDispEx;
};

#endif
