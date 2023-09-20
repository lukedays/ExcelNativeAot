namespace Addin;

using Addin.ComApi;
using Addin.Types.Managed;
using System.Runtime.InteropServices;
using static Addin.CApi.ExcelEntryPoints;
using static Addin.Types.Unmanaged.ExcelConstants;

public static class UserFunctions
{
    [UnmanagedCallersOnly(EntryPoint = nameof(TestVersion))]
    public static unsafe nint TestVersion()
    {
        var app = InstanceFinder.GetCurrentExcelInstance();

        var version = app.GetProperty("Version") as string; // This works

        //var version = (app as dynamic).Version as string; // Doesn't work

        return new XlOper(version).ToPtr();
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(TestAddDouble))]
    public static double TestAddDouble(double x, double y)
    {
        return x + y;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(TestConcatString))]
    public static nint TestConcatString(nint ptr1, nint ptr2)
    {
        var str1 = new XlOper(ptr1).ToString() ?? "";
        var str2 = new XlOper(ptr2).ToString() ?? "";
        return new XlOper(str1 + str2).ToPtr();
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(xlAutoOpen))]
    public static int xlAutoOpen()
    {
        var dllPtr = new XlOper().ToPtr();

        // Get DLL name
        Excel12v(xlGetName, dllPtr, 0, []);

        // Register test functions
        Excel12v(
            xlfRegister,
            0,
            4,
            [
                dllPtr,
                new XlOper(nameof(TestAddDouble)).ToPtr(),
                new XlOper("BBB").ToPtr(),
                new XlOper(nameof(TestAddDouble)).ToPtr(),
            ]
        );

        Excel12v(
            xlfRegister,
            0,
            4,
            [
                dllPtr,
                new XlOper(nameof(TestConcatString)).ToPtr(),
                new XlOper("QQQ").ToPtr(),
                new XlOper(nameof(TestConcatString)).ToPtr(),
            ]
        );

        Excel12v(
            xlfRegister,
            0,
            4,
            [dllPtr, new XlOper(nameof(TestVersion)).ToPtr(), new XlOper("Q").ToPtr(), new XlOper(nameof(TestVersion)).ToPtr()]
        );

        // Free the handler
        Excel12v(xlFree, 0, 1, [dllPtr]);

        return 1;
    }
}
