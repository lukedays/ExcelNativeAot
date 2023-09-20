//#include <windows.h>
//#include <oleacc.h>
//#include <olectl.h>
//#include <atlcomcli.h>
//#include <comdef.h>
////#include "C:/Lucas/dev/ExcelNativeAot/UnmanagedCaller/x64/Debug/MSO.tlh"
////#include "C:/Lucas/dev/ExcelNativeAot/UnmanagedCaller/x64/Debug/EXCEL.tlh"
////#import "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\MSO.DLL" no_implementation rename("RGB", "ExclRGB") rename("DocumentProperties", "ExclDocumentProperties") rename("SearchPath", "ExclSearchPath")
////#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB" no_implementation
////#import "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" \
////    rename("DialogBox", "ExcelDialogBox") \
////    rename("RGB", "ExcelRGB") \
////    rename("CopyFile", "ExcelCopyFile") \
////    rename("ReplaceText", "ExcelReplaceText") \
////    exclude("IFont", "IPicture")
//
//#pragma comment (lib, "oleacc.lib")
//
//BOOL CALLBACK EnumChildProc(HWND hwnd, LPARAM lParam)
//{
//    WCHAR szClassName[64];
//    if (GetClassNameW(hwnd, szClassName, 64))
//    {
//        if (_wcsicmp(szClassName, L"EXCEL7") == 0)
//        {
//            // Get AccessibleObject
//            IDispatch* pDisp = NULL;
//            HRESULT hr = AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, IID_IDispatch, (void**)&pDisp);
//            if (FAILED(hr) || pDisp == NULL) {
//                printf("Failed to get IUnknown interface pointer.\n");
//                return 1;
//            }
//
//            if (hr == S_OK)
//            {
//                //IDispatch* pApp = NULL;
//                //pApp = pDisp->GetApplication();
//
//                DISPID dispid;
//                LPCOLESTR name = L"Application";
//                hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
//                if (FAILED(hr)) {
//                    printf("Failed to get Version\n");
//                    return true;
//                }
//
//                DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
//                CComVariant result;
//                hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
//                if (FAILED(hr)) {
//                    printf("Failed to get Version object\n");
//                    return true;
//                }
//                pDisp->Release();
//
//                
//            }
//            return false; // Stops enumerating through children
//        }
//    }
//    return true;
//}
//
//int main(int argc, CHAR* argv[])
//{
//    // Initialize COM
//    CoInitialize(NULL);
//
//    // The main window in Microsoft Excel has a class name of XLMAIN
//    HWND excelWindow = FindWindow(L"XLMAIN", NULL);
//
//    // Use the EnumChildWindows function to iterate through all child windows until we find EXCEL7
//    EnumChildWindows(excelWindow, (WNDENUMPROC)EnumChildProc, (LPARAM)1);
//
//    /*IDispatch* pDisp = NULL;
//    HRESULT hr = AccessibleObjectFromWindow(excelWindow, OBJID_NATIVEOM, IID_IDispatch, (void**)&pDisp);
//    if (hr == S_OK)
//    {
//        DISPID dispid;
//        LPCOLESTR name = L"Version";
//        hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
//        if (FAILED(hr)) {
//            printf("Failed to get Version\n");
//            return 1;
//        }
//
//        DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
//        CComVariant result;
//        hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
//        if (FAILED(hr)) {
//            printf("Failed to get Version object\n");
//            return 1;
//        }
//        pDisp->Release();
//    }*/
//
//    return 0;
//}
////
////#include <windows.h>
////#include <comdef.h>
////#include <iostream>
////#include <atlcomcli.h>
////
////int main()
////{
////    CoInitialize(NULL);
////    CLSID clsid;
////    CLSIDFromProgID(L"Excel.Application", &clsid);
////    IDispatch* pXlApp;
////    HRESULT hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pXlApp);
////    if (FAILED(hr)) {
////        std::cout << "Failed to create instance of Excel" << std::endl;
////        return -1;
////    }
////    LPCOLESTR szVisible = L"Visible";
////    DISPID dispid;
////    hr = pXlApp->GetIDsOfNames(IID_NULL, (LPOLESTR*)& szVisible, 1, LOCALE_USER_DEFAULT, &dispid);
////    if (SUCCEEDED(hr)) {
////        VARIANT x;
////        x.vt = VT_BOOL;
////        x.boolVal = VARIANT_TRUE;
////
////        // Allocate memory for arguments...
////        VARIANT* pArgs = new VARIANT[1];
////        // Extract arguments...
////        pArgs[0] = x;
////        DISPPARAMS dp = { NULL, NULL, 0, 0 };
////        EXCEPINFO excep;
////        VARIANT pRes;
////        UINT par;
////        DISPID dispidNamed = DISPID_PROPERTYPUT;
////        // Build DISPPARAMS
////        dp.cArgs = 1;
////        dp.rgvarg = &x;
////        dp.cNamedArgs = 1;
////        dp.rgdispidNamedArgs = &dispidNamed;
////
////        hr = pXlApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, &pRes, &excep, &par);
////        if (FAILED(hr)) {
////            std::cout << "Failed to set Visible property" << std::endl;
////            return -1;
////        }
////
////        dispid;
////        LPCOLESTR name = L"Workbooks";
////        hr = pXlApp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
////        if (FAILED(hr)) {
////            printf("Failed to get Workbooks ID\n");
////            return 1;
////        }
////
////        DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
////        CComVariant result;
////        hr = pXlApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
////        if (FAILED(hr)) {
////            printf("Failed to get Workbooks object\n");
////            return 1;
////        }
////
////        CComPtr<IDispatch> pWorkbooks = result.pdispVal;
////
////        name = L"Add";
////        hr = pWorkbooks->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
////        if (FAILED(hr)) {
////            printf("Failed to get Add method ID\n");
////            return 1;
////        }
////
////        hr = pWorkbooks->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &noArgs, &result, NULL, NULL);
////        if (FAILED(hr)) {
////            printf("Failed to add workbook\n");
////            return 1;
////        }
////
////        CComPtr<IDispatch> pWorkbook = result.pdispVal;
////
////        name = L"Sheets";
////        hr = pWorkbook->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
////        if (FAILED(hr)) {
////            printf("Failed to get Sheets method ID\n");
////            return 1;
////        }
////
////        hr = pWorkbook->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
////        if (FAILED(hr)) {
////            printf("Failed to get Workbooks object\n");
////            return 1;
////        }
////
////        CComPtr<IDispatch> pSheets = result.pdispVal;
////
////        name = L"Item";
////        hr = pSheets->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
////        if (FAILED(hr)) {
////            printf("Failed to get Item method ID\n");
////            return 1;
////        }
////
////        CComVariant arg(1);  // Index of the first sheet
////        DISPPARAMS args = { &arg, NULL, 1, 0 };
////        hr = pSheets->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &args, &result, NULL, NULL);
////        if (FAILED(hr)) {
////            printf("Failed to get first sheet\n");
////            return 1;
////        }
////
////        CComPtr<IDispatch> pSheet = result.pdispVal;
////
////        name = L"Name";
////        hr = pSheet->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
////        if (FAILED(hr)) {
////            printf("Failed to get Name property ID\n");
////            return 1;
////        }
////
////        CComVariant newName("TestTest");
////        args = { &newName, NULL, 1, 0 };
////        hr = pSheet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &args, &result, NULL, NULL);
////        if (FAILED(hr)) {
////            printf("Failed to set sheet name\n");
////            return 1;
////        }
////    }
////    else {
////        std::cout << "Failed to get dispid for Visible property" << std::endl;
////        return -1;
////    }
////    pXlApp->Release();
////    CoUninitialize();
////}
