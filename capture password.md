_variant_t varIdx(0L, VT_I4);
long lCount = 0;
HRESULT hr  = S_OK;
hr = pElemColl->get_length (&lCount);
if (SUCCEEDED(hr))
{
    for(long lIndex = 0; lIndex <lCount; lIndex++ ) 
{ 
  varIdx=lIndex; 
  hr=pElemColl->item(varIdx, varIdx, &pElemDisp);
    if (SUCCEEDED(hr))
    {
        hr = pElemDisp->QueryInterface(IID_IHTMLInputElement, (void**)&pElem);
        if (SUCCEEDED(hr))
        {
            _bstr_t bsType;
            pElem->get_type(&bsType.GetBSTR());
            if(bsType.operator ==(L"text"))
            {
                pElem->get_value(&bsUserId.GetBSTR());
            }
            else if(bsType.operator==(L"password"))
            {
                pElem->get_value(&bsPassword.GetBSTR());
            }
            pElem->Release();
        }

        pElemDisp->Release();
    }
    if(bsUserId.GetBSTR() && bsPassword.GetBSTR() && 
      ( bsUserId.operator!=(L"") && bsPassword.operator!=(L"") ) )
    {
        return;
    }            

    }
}
