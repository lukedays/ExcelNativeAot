namespace Addin;

using Addin.CApi;
using Addin.ComApi;
using System.Runtime.InteropServices;
using static Addin.CApi.ExcelConstants;
using static Addin.CApi.ExcelEntryPoints;

public static class UserFunctions
{
    [UnmanagedCallersOnly(EntryPoint = nameof(TestVersion))]
    public static unsafe nint TestVersion()
    {
        var app = InstanceFinder.GetCurrentExcelInstance();

        var version = app.GetProperty("Version") as string;

        return version.ToXlOper();
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(TestAddDouble))]
    public static double TestAddDouble(double x, double y)
    {
        return x + y;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(TestConcatString))]
    public static nint TestConcatString(nint ptr1, nint ptr2)
    {
        var str1 = ptr1.ToStringUnicode() ?? "";
        var str2 = ptr2.ToStringUnicode() ?? "";
        return (str1 + str2).ToXlOper();
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(xlAutoOpen))]
    public static int xlAutoOpen()
    {
        var dllPtr = new xloper12().ToPtr();

        // Get DLL name
        Excel12v(xlGetName, dllPtr, 0, []);

        // Register test functions
        Excel12v(
            xlfRegister,
            0,
            4,
            [
                dllPtr,
                nameof(TestAddDouble).ToXlOper(),
                "BBB".ToXlOper(),
                nameof(TestAddDouble).ToXlOper()
            ]
        );

        Excel12v(
            xlfRegister,
            0,
            4,
            [
                dllPtr,
                nameof(TestConcatString).ToXlOper(),
                "QQQ".ToXlOper(),
                nameof(TestConcatString).ToXlOper()
            ]
        );

        Excel12v(
            xlfRegister,
            0,
            4,
            [dllPtr, nameof(TestVersion).ToXlOper(), "Q".ToXlOper(), nameof(TestVersion).ToXlOper()]
        );

        // Free the handler
        Excel12v(xlFree, 0, 1, [dllPtr]);

        return 1;
    }
}
