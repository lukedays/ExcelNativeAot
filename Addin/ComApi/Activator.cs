using System.Runtime.InteropServices;
using static Addin.ComApi.Constants;

namespace Addin.ComApi;

public partial class Activator
{
    [LibraryImport("ole32.dll")]
    public static partial int CoCreateInstance(
        ref Guid rclsid,
        nint pUnkOuter,
        CLSCTX dwClsContext,
        ref Guid riid,
        out IDispatch ppv
    );

    public static IDispatch ActivateClass(Guid clsid, CLSCTX server)
    {
        var guid = typeof(IDispatch).GUID;

        int hr = CoCreateInstance(ref clsid, nint.Zero, server, ref guid, out IDispatch obj);
        if (hr < 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }
        return obj;
    }
}
