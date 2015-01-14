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

class CEnumVariant : public IEnumVARIANT
{
public:
    // param v: Enumerator
    CEnumVariant(VALUE v0) :
        m_object(v0),
        m_lCount(1),
        m_endOfEnum(false),
        m_nIterate(0)
    {
        m_enum = rb_funcall(m_object, rb_intern("to_enum"), 0);
        rb_gc_register_address(&m_enum); // refer to the orijinal
    }

    CEnumVariant(CEnumVariant& o)
        : m_object(o.m_object),
          m_lCount(1),
          m_endOfEnum(o.m_endOfEnum),
          m_nIterate(0)
    {
        m_enum = rb_funcall(m_object, rb_intern("to_enum"), 0);
        rb_gc_register_address(&m_enum); // refer to the orijinal
        Skip(o.m_nIterate);
    }

    ~CEnumVariant()
    {
        rb_gc_unregister_address(&m_object);
    }

    HRESULT  STDMETHODCALLTYPE QueryInterface(
		const IID & riid,  
		void **ppvObj);

    ULONG  STDMETHODCALLTYPE AddRef();

    ULONG  STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE Next( 
        /* [in] */ ULONG celt,
        /* [length_is][size_is][out] */ VARIANT *rgVar,
        /* [out] */ ULONG *pCeltFetched);
        
    HRESULT STDMETHODCALLTYPE Skip( 
        /* [in] */ ULONG celt);
        
    HRESULT STDMETHODCALLTYPE Reset( void);
        
    HRESULT STDMETHODCALLTYPE Clone( 
        /* [out] */ __RPC__deref_out_opt IEnumVARIANT **ppEnum);

private:
    ULONG m_lCount;
    VALUE m_object;
    VALUE m_enum;
    int m_nIterate;
    bool m_endOfEnum;
};
