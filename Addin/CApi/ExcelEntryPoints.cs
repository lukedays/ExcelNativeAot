namespace Addin.CApi;

using System.Runtime.InteropServices;
using static Addin.CApi.ExcelConstants;

public static partial class ExcelEntryPoints
{
    [LibraryImport("kernel32.dll")]
    public static partial nint GetModuleHandleW(
        [MarshalAs(UnmanagedType.LPWStr)] string lpModuleName
    );

    [LibraryImport("kernel32.dll")]
    public static partial nint GetProcAddress(
        nint hModule,
        [MarshalAs(UnmanagedType.LPStr)] string procName
    );

    public delegate int EXCEL12PROC(int xlfn, int coper, nint[] rgpxloper12, nint xloper12Res);
    public static nint hmodule;
    public static EXCEL12PROC pexcel12;

    public static int Excel12v(int xlfn, nint operRes, int count, nint[] opers)
    {
        FetchExcel12EntryPt();

        return pexcel12 == null ? xlretFailed : pexcel12(xlfn, count, opers, operRes);
    }

    public static void FetchExcel12EntryPt()
    {
        if (pexcel12 != null)
            return;

        hmodule = GetModuleHandleW(null);
        if (hmodule == nint.Zero)
            return;

        pexcel12 = Marshal.GetDelegateForFunctionPointer<EXCEL12PROC>(
            GetProcAddress(hmodule, "MdCallBack12")
        );
    }
}
