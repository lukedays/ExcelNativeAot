using System.Runtime.InteropServices.Marshalling;
using Unmanaged = Addin.ComApi.Types.Unmanaged;

namespace Addin.ComApi.Marshalling;

[CustomMarshaller(typeof(Types.Managed.DispParams), MarshalMode.Default, typeof(DispParams))]
public static partial class DispParams
{
    public static unsafe Unmanaged.DispParams ConvertToUnmanaged(Types.Managed.DispParams managed)
    {
        return new Unmanaged.DispParams
        {
            cArgs = managed.cArgs,
            cNamedArgs = managed.cNamedArgs,
            rgdispidNamedArgs = &managed.rgdispidNamedArgs,
            rgvarg =
                managed.rgvarg != null
                    ? Helpers.ArrayToPtr(
                        managed.rgvarg.Select(Variant.ConvertToUnmanaged).ToArray()
                    )
                    : nint.Zero
        };
    }

    public static unsafe Types.Managed.DispParams ConvertToManaged(Unmanaged.DispParams unmanaged)
    {
        return new Types.Managed.DispParams
        {
            cArgs = unmanaged.cArgs,
            cNamedArgs = unmanaged.cNamedArgs,
            rgdispidNamedArgs = *unmanaged.rgdispidNamedArgs,
            rgvarg = Helpers
                .PtrToArray<Unmanaged.Variant>(unmanaged.rgvarg, unmanaged.cArgs)
                .Select(Variant.ConvertToManaged)
                .ToArray(),
        };
    }
}
