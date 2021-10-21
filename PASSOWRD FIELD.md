BOOL  CIeLoginHelper::IsLoginPage(CComPtr<IHTMLElementCollection>&spElemColl)
{
    if(spElemColl == NULL)
        return m_bFoungLoginPage;
    _variant_t varIdx(0L, VT_I4);
    long lCount = 0;
    HRESULT hr  = S_OK;
    hr = spElemColl->get_length (&lCount);
    if (SUCCEEDED(hr))
    {
        for(long lIndex = 0; lIndex <lCount; lIndex++ ) 
        { 
            varIdx=lIndex;
                    CComPtr<IDispatch>spElemDisp;
            hr = spElemColl->item(varIdx, varIdx, &spElemDisp);
            if (SUCCEEDED(hr))
            {
                CComPtr<IHTMLInputElement> spElem;
                hr = spElemDisp->QueryInterface(IID_IHTMLInputElement, (void**)&spElem);
                if (SUCCEEDED(hr))
                {
                    _bstr_t bsType;
                    hr = spElem->get_type(&bsType.GetBSTR());
                    if(SUCCEEDED(hr) && bsType.operator==(L"password"))
                    {
                        m_bFoungLoginPage = true;
                    }
                }
            }
            if(m_bFoungLoginPage)
                return m_bFoungLoginPage;
        }
    }
    return m_bFoungLoginPage;
