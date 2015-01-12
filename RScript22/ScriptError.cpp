/*
 * Copyright(c) 2015 arton
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
#include "ScriptError.h"

CScriptError::CScriptError(LPCSTR script)
        : m_lCount(1), m_nLine(-1)
{
    VALUE v = rb_errinfo();
    VALUE msg = rb_funcall(v, rb_intern("message"), 0);
    m_strMessage = StringValueCStr(msg);
    VALUE bt = rb_funcall(v, rb_intern("backtrace"), 0);
    if (!NIL_P(bt))
    {
        VALUE top = rb_ary_entry(bt, 0);
        if (!NIL_P(top))
        {
            m_nLine = GetLineNumber(StringValueCStr(top));
        }
    }
    if (script && m_nLine >= 0)
    {
        SetupSourceLine(script, m_nLine);
    }
}

HRESULT STDMETHODCALLTYPE CScriptError::GetExceptionInfo(
    EXCEPINFO *pExcepInfo  // structure for exception information
)
{
    if (!pExcepInfo) return E_POINTER;
    USES_CONVERSION;
    memset(pExcepInfo, 0, sizeof(EXCEPINFO));
    pExcepInfo->wCode = 0x0200;// + istat;
    pExcepInfo->bstrSource = SysAllocString(L"RScript");
    pExcepInfo->bstrDescription = SysAllocString(A2W(m_strMessage.c_str()));
    pExcepInfo->scode = DISP_E_EXCEPTION;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE CScriptError::GetSourcePosition(
    DWORD *pdwSourceContext,  // context cookie
    ULONG *pulLineNumber,     // line number of error
    LONG *pichCharPosition    // character position of error
)
{
    if (pdwSourceContext) *pdwSourceContext = 0;
    if (pulLineNumber) *pulLineNumber = m_nLine;
    if (pichCharPosition) *pichCharPosition = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE CScriptError::GetSourceLineText(
    BSTR *pbstrSourceLine  // address of buffer for source line
)
{
    if (!pbstrSourceLine) return E_POINTER;
    if (m_strSource.length() > 0)
    {
        USES_CONVERSION;
        *pbstrSourceLine = SysAllocString(A2W(m_strSource.c_str()));
    }
    return S_OK;
}

int CScriptError::GetLineNumber(LPCSTR btline)
{
    if (btline)
    {
        for (LPCSTR p = btline; *p; p++)
        {
            if (*p == ':' && *(p + 1) != '\\') // skip drive letter
            {
                int n = 0;
                for (LPCSTR start = p + 1; *start && isdigit(*start); start++)
                {
                    n *= 10;
                    n += *start - '0';
                }
                return n;
            }
        }
    }
    return 0;
}

void CScriptError::SetupSourceLine(LPCSTR p, int n)
{
    LPCSTR p2(NULL);
    for (int i = 0; i < n; i++)
    {
        p2 = strchr(p, '\n');
        if (!p2 || !*(p2 + 1)) break;
        p = p2 + 1;
    }
    size_t len;
    p2 = strchr(p, '\n');
    if (!p2)
    {
        len = strlen(p);
    }
    else if (p2 > p && *(p2 - 1) == '\r')
    {
        len = p2 - p - 1;
    }
    else
    {
        len = p2 - p;
    }
    m_strSource.append(p, len);
}

