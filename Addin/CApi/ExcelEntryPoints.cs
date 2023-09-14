namespace Addin.CApi;

using System.Runtime.InteropServices;
using static Addin.CApi.ExcelConstants;

public static class ExcelEntryPoints
{
    [DllImport("kernel32.dll")]
    public static extern nint GetModuleHandle(string lpModuleName);

    [DllImport("kernel32.dll")]
    public static extern nint GetProcAddress(nint hModule, string procName);

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

        hmodule = GetModuleHandle(null);
        if (hmodule == nint.Zero)
            return;

        pexcel12 = Marshal.GetDelegateForFunctionPointer<EXCEL12PROC>(
            GetProcAddress(hmodule, "MdCallBack12")
        );
    }
}
