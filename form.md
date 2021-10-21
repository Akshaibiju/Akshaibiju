##CONNECTING TO <FORM> EVENTS##

//
CComPtr<IDISPATCH> spDisp; 
HRESULT hr = m_pWebBrowser2->get_Document(&spDisp);
if (SUCCEEDED(hr) && spDisp)
{ 
    // If this is not an HTML document (e.g., it's a Word doc or a PDF), don't sink.
    CComQIPtr<IHTMLDOCUMENT2 &IID_IHTMLDocument2> spHTML(spDisp);
     if (spHTML)
     { 
         /*there can be frames in HTML page enumerate each of frameset or iframe
              and find out if any of them contain a login page*/
           EnumFrames(spHTML);  
    }
}

void CIeLoginHelper::EnumFrames(CComPtr<IHTMLDocument2>& spDocument) 
{    
    
    CComPtr<IIHTMLDocument2> spDocument2;
    CComPtr<IIOleContainer> pContainer;
    // Get the container
    HRESULT hr = spDocument->QueryInterface(IID_IOleContainer,
                (void**)&pContainer);
    
    CComPtr<IIEnumUnknown>pEnumerator;
    // Get an enumerator for the frames
    hr = pContainer->EnumObjects(OLECONTF_EMBEDDINGS, &pEnumerator);
    IUnknown* pUnk;
    ULONG uFetched;

    // Enumerate and refresh all the frames
    BOOL bFrameFound = FALSE;
    for (UINT nIndex = 0; 
            S_OK == pEnumerator->Next(1, &pUnk, &uFetched);
                nIndex++)
    {
        CComPtr<IIWebBrowser2> pBrowser;
        hr = pUnk->QueryInterface(IID_IWebBrowser2, 
                (void**)&pBrowser);
        pUnk->Release();
        if (SUCCEEDED(hr))
        {
            CComPtr<IIDispatch> spDisp;
            pBrowser->get_Document(&spDisp);
            CComQIPtr<IHTMLDocument2, &
                         IID_IHTMLDocument2> spDocument2(spDisp);
            //Now recursivley browse through all of
                        //IHTMLWindow2 in a doc                    
            RecurseWindows(spDocument2);
            bFrameFound = TRUE;

        }
    }
    if(!bFrameFound || !m_bFoungLoginPage)
    {

        CComPtr<IIHTMLElementCollection> spFrmCol;
        CComPtr<IIHTMLElementCollection> spElmntCol;
        /*multipe <FORM> object can be in a page,
                 connect to each one them
        You never know which one contains uid and pwd fields
        */
        hr = spDocument->get_forms(&spFrmCol);
        // get element collection from page to check 
                if a page is a lgoin page
        hr = spDocument->get_all(&spElmntCol);
        if(IsLoginPage(spElmntCol))
                   EnableEvents(spFrmCol);    
