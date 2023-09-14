namespace Addin;

using Addin.CApi;
using Addin.ComApi;
using System.Runtime.InteropServices;
using static Addin.CApi.ExcelConstants;
using static Addin.CApi.ExcelEntryPoints;

public static class UserFunctions
{
    //[UnmanagedCallersOnly(EntryPoint = nameof(ComTest))] // TODO: call from unmanaged code
    public static unsafe void ComTest()
    {
        dynamic app = new ExcelApplication();

        Console.WriteLine($"Version: {app.Version}");

        Console.WriteLine($"Visible: {app.Visible}");

        app.Visible = true;

        Console.WriteLine($"Visible: {app.Visible}");

        var wb = app.Workbooks.Add();

        Console.WriteLine($"Name: {wb.Sheets[1].Name}");

        wb.Sheets[1].Name = "FirstSheet";

        Console.WriteLine($"Name: {wb.Sheets[1].Name}");
    }

    public static double ManagedAdd(double x, double y)
    {
        return x + y;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(TestAddDouble))]
    public static double TestAddDouble(double x, double y)
    {
        return ManagedAdd(x, y);
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
        Excel12v(xlGetName, dllPtr, 0, Array.Empty<nint>());

        // Register test functions
        Excel12v(
            xlfRegister,
            0,
            4,
            new[]
            {
                dllPtr,
                nameof(TestAddDouble).ToXlOper(),
                "BBB".ToXlOper(),
                nameof(TestAddDouble).ToXlOper()
            }
        );

        Excel12v(
            xlfRegister,
            0,
            4,
            new[]
            {
                dllPtr,
                nameof(TestConcatString).ToXlOper(),
                "QQQ".ToXlOper(),
                nameof(TestConcatString).ToXlOper()
            }
        );

        // Free the handler
        Excel12v(xlFree, 0, 1, new[] { dllPtr });

        return 1;
    }
}
