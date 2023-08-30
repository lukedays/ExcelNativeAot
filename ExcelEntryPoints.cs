namespace ExcelNativeAot;

using System.Runtime.InteropServices;
using static ExcelNativeAot.ExcelConstants;

public static class ExcelEntryPoints
{
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("kernel32.dll")]
    public static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

    public delegate int EXCEL12PROC(int xlfn, int coper, IntPtr[] rgpxloper12, IntPtr xloper12Res);
    public static IntPtr hmodule;
    public static EXCEL12PROC pexcel12;

    public static int Excel12v(int xlfn, IntPtr operRes, int count, IntPtr[] opers)
    {
        FetchExcel12EntryPt();

        return pexcel12 == null ? xlretFailed : pexcel12(xlfn, count, opers, operRes);
    }

    public static void FetchExcel12EntryPt()
    {
        if (pexcel12 != null)
            return;

        hmodule = GetModuleHandle(null);
        if (hmodule == IntPtr.Zero)
            return;

        pexcel12 = Marshal.GetDelegateForFunctionPointer<EXCEL12PROC>(
            GetProcAddress(hmodule, "MdCallBack12")
        );
    }
}
