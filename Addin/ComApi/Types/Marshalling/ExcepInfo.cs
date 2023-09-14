using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Unmanaged = Addin.ComApi.Types.Unmanaged;

namespace Addin.ComApi.Marshalling;

[CustomMarshaller(typeof(Types.Managed.ExcepInfo), MarshalMode.Default, typeof(ExcepInfo))]
public static partial class ExcepInfo
{
    public static Unmanaged.ExcepInfo ConvertToUnmanaged(Types.Managed.ExcepInfo managed)
    {
        return new Unmanaged.ExcepInfo
        {
            bstrDescription = Marshal.StringToBSTR(managed.bstrDescription),
            bstrHelpFile = Marshal.StringToBSTR(managed.bstrHelpFile),
            bstrSource = Marshal.StringToBSTR(managed.bstrSource),
            dwHelpContext = managed.dwHelpContext,
            pfnDeferredFillIn = managed.pfnDeferredFillIn,
            pvReserved = managed.pvReserved,
            scode = managed.scode,
            wCode = managed.wCode,
            wReserved = managed.wReserved,
        };
    }

    public static Types.Managed.ExcepInfo ConvertToManaged(Unmanaged.ExcepInfo unmanaged)
    {
        return new Types.Managed.ExcepInfo
        {
            bstrDescription =
                unmanaged.bstrDescription != 0
                    ? Marshal.PtrToStringBSTR(unmanaged.bstrDescription)
                    : "",
            bstrHelpFile =
                unmanaged.bstrHelpFile != 0 ? Marshal.PtrToStringBSTR(unmanaged.bstrHelpFile) : "",
            bstrSource =
                unmanaged.bstrSource != 0 ? Marshal.PtrToStringBSTR(unmanaged.bstrSource) : "",
            dwHelpContext = unmanaged.dwHelpContext,
            pfnDeferredFillIn = unmanaged.pfnDeferredFillIn,
            pvReserved = unmanaged.pvReserved,
            scode = unmanaged.scode,
            wCode = unmanaged.wCode,
            wReserved = unmanaged.wReserved,
        };
    }
}
