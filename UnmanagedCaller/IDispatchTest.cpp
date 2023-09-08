//#include <windows.h>
//#include <comdef.h>
//#include <iostream>
//
//int main()
//{
//    CoInitialize(NULL);
//    CLSID clsid;
//    CLSIDFromProgID(L"Excel.Application", &clsid);
//    IDispatch* pXlApp;
//    HRESULT hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pXlApp);
//    if (FAILED(hr)) {
//        std::cout << "Failed to create instance of Excel" << std::endl;
//        return -1;
//    }
//    LPCOLESTR szVisible = L"Visible";
//    DISPID dispid;
//    hr = pXlApp->GetIDsOfNames(IID_NULL, (LPOLESTR*)& szVisible, 1, LOCALE_USER_DEFAULT, &dispid);
//    if (SUCCEEDED(hr)) {
//        VARIANT x;
//        x.vt = VT_BOOL;
//        x.boolVal = VARIANT_TRUE;
//
//        // Allocate memory for arguments...
//        VARIANT* pArgs = new VARIANT[1];
//        // Extract arguments...
//        pArgs[0] = x;
//        DISPPARAMS dp = { NULL, NULL, 0, 0 };
//        EXCEPINFO excep;
//        VARIANT pRes;
//        UINT par;
//        DISPID dispidNamed = DISPID_PROPERTYPUT;
//        // Build DISPPARAMS
//        dp.cArgs = 1;
//        dp.rgvarg = &x;
//        dp.cNamedArgs = 1;
//        dp.rgdispidNamedArgs = &dispidNamed;
//
//        hr = pXlApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, &pRes, &excep, &par);
//        if (FAILED(hr)) {
//            std::cout << "Failed to set Visible property" << std::endl;
//            return -1;
//        }
//    }
//    else {
//        std::cout << "Failed to get dispid for Visible property" << std::endl;
//        return -1;
//    }
//    pXlApp->Release();
//    CoUninitialize();
//}
