using Addin.CApi;
using Addin.Types.Unmanaged;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using static Addin.CApi.ExcelEntryPoints;
using static Addin.Types.Unmanaged.ExcelConstants;

namespace Addin.ComApi;

internal static partial class InstanceFinder
{
    public static unsafe ExcelApplication? GetCurrentExcelInstance()
    {
        // Initialize COM
        CoInitialize(nint.Zero);

        // Get the pointer to the current Excel window handler
        var hwndPtr = new xloper12().ToPtr();

        Excel12v(xlGetHwnd, hwndPtr, 0, []);

        var hwnd = Marshal.ReadIntPtr(hwndPtr);

        // Search the accessible child window (it has class name "EXCEL7")
        var callback = new EnumChildCallback(EnumChildProc);
        var excelWindowPtr = nint.Zero;
        var res = EnumChildWindows(hwnd, callback, ref excelWindowPtr);

        // Convert to a managed IDispatch
        var excelWindow = ComInterfaceMarshaller<IDispatch>.ConvertToManaged((void*)excelWindowPtr);
        var excelWindowWrapper = new ExcelApplication(excelWindow);

        return excelWindowWrapper.GetProperty("Application") as ExcelApplication;
    }

    [LibraryImport("oleacc.dll")]
    public static partial int AccessibleObjectFromWindow(
        nint hwnd,
        uint dwId,
        Guid riid,
        out nint ppvObject
    );

    [LibraryImport("ole32.dll")]
    public static partial int CoInitialize(nint pvReserved);

    public static bool EnumChildProc(nint hwndChild, ref nint lParam)
    {
        if (GetClassName(hwndChild) == "EXCEL7")
        {
            var guid = typeof(IDispatch).GUID;
            var hr = AccessibleObjectFromWindow(
                hwndChild,
                ComConstants.OBJID_NATIVEOM,
                guid,
                out nint excelWindowPtr
            );

            Marshal.ThrowExceptionForHR(hr);

            lParam = excelWindowPtr;

            // Found handler, stop iterating
            return false;
        }
        // Continue iterating through child windows
        return true;
    }

    public delegate bool EnumChildCallback(nint hwnd, ref nint lParam);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool EnumChildWindows(
        nint hWndParent,
        EnumChildCallback lpEnumFunc,
        ref nint lParam
    );

    [LibraryImport("user32.dll")]
    public static partial int GetClassNameW(
        nint hWnd,
        [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U2)] char[] lpClassName,
        int nMaxCount
    );

    private static string GetClassName(nint hwndChild)
    {
        var buffer = new char[256];
        GetClassNameW(hwndChild, buffer, buffer.Length);
        return new string(buffer).TrimEnd('\0'); // Important for string comparison
    }
}