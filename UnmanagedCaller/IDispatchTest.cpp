//#include <windows.h>
//#include <comdef.h>
//#include <iostream>
//#include <atlcomcli.h>
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
//
//        dispid;
//        LPCOLESTR name = L"Workbooks";
//        hr = pXlApp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
//        if (FAILED(hr)) {
//            printf("Failed to get Workbooks ID\n");
//            return 1;
//        }
//
//        DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
//        CComVariant result;
//        hr = pXlApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
//        if (FAILED(hr)) {
//            printf("Failed to get Workbooks object\n");
//            return 1;
//        }
//
//        CComPtr<IDispatch> pWorkbooks = result.pdispVal;
//
//        name = L"Add";
//        hr = pWorkbooks->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
//        if (FAILED(hr)) {
//            printf("Failed to get Add method ID\n");
//            return 1;
//        }
//
//        hr = pWorkbooks->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &noArgs, &result, NULL, NULL);
//        if (FAILED(hr)) {
//            printf("Failed to add workbook\n");
//            return 1;
//        }
//
//        CComPtr<IDispatch> pWorkbook = result.pdispVal;
//
//        name = L"Sheets";
//        hr = pWorkbook->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
//        if (FAILED(hr)) {
//            printf("Failed to get Sheets method ID\n");
//            return 1;
//        }
//
//        hr = pWorkbook->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
//        if (FAILED(hr)) {
//            printf("Failed to get Workbooks object\n");
//            return 1;
//        }
//
//        CComPtr<IDispatch> pSheets = result.pdispVal;
//
//        name = L"Item";
//        hr = pSheets->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
//        if (FAILED(hr)) {
//            printf("Failed to get Item method ID\n");
//            return 1;
//        }
//
//        CComVariant arg(1);  // Index of the first sheet
//        DISPPARAMS args = { &arg, NULL, 1, 0 };
//        hr = pSheets->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &args, &result, NULL, NULL);
//        if (FAILED(hr)) {
//            printf("Failed to get first sheet\n");
//            return 1;
//        }
//
//        CComPtr<IDispatch> pSheet = result.pdispVal;
//
//        name = L"Name";
//        hr = pSheet->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
//        if (FAILED(hr)) {
//            printf("Failed to get Name property ID\n");
//            return 1;
//        }
//
//        CComVariant newName("TestTest");
//        args = { &newName, NULL, 1, 0 };
//        hr = pSheet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &args, &result, NULL, NULL);
//        if (FAILED(hr)) {
//            printf("Failed to set sheet name\n");
//            return 1;
//        }
//    }
//    else {
//        std::cout << "Failed to get dispid for Visible property" << std::endl;
//        return -1;
//    }
//    pXlApp->Release();
//    CoUninitialize();
//}
