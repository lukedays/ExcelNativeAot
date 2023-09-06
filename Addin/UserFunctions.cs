﻿namespace Addin;

using System.Runtime.InteropServices;
using static Addin.ExcelConstants;
using static Addin.ExcelEntryPoints;

public static class UserFunctions
{
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
    public static IntPtr TestConcatString(IntPtr ptr1, IntPtr ptr2)
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
        Excel12v(xlGetName, dllPtr, 0, Array.Empty<IntPtr>());

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
