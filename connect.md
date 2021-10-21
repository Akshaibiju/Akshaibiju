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
            hr = pElemDisp->QueryInterface(IID_IHTMLFormElement, (void**)&pElem);

            if (SUCCEEDED(hr))
            {
                // Obtained a form object.
                IConnectionPointContainer* pConPtContainer = NULL;
                IConnectionPoint* pConPt = NULL;    
                // Check that this is a connectable object.
                hr = pElem->QueryInterface(IID_IConnectionPointContainer,
                    (void**)&pConPtContainer);
                if (SUCCEEDED(hr))
                {
                    // Find the connection point.
                    hr = pConPtContainer->FindConnectionPoint(
                        DIID_HTMLFormElementEvents2, &pConPt);

                    if (SUCCEEDED(hr))
                    {
                        // Advise the connection point.
                        // pUnk is the IUnknown interface pointer for your event sink
                        hr = pConPt->Advise((IDispatch*)this, &m_dwBrowserCookie);
                        pConPt->Release();
                    }
                }
                pElem->Release();
            }
            pElemDisp->Release();
        }
    }
}
