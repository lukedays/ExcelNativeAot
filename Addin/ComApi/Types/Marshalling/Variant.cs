using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using static Addin.ComApi.Constants;
using Unmanaged = Addin.ComApi.Types.Unmanaged;

namespace Addin.ComApi.Marshalling;

[CustomMarshaller(typeof(Types.Managed.Variant), MarshalMode.Default, typeof(Variant))]
public static class Variant
{
    public static Unmanaged.Variant ConvertToUnmanaged(Types.Managed.Variant managed)
    {
        return managed.Value switch
        {
            bool boolVal
                => new Unmanaged.Variant
                {
                    vt = (ushort)VARTYPE.VT_BOOL,
                    boolVal = (short)(
                        boolVal ? VARIANT_BOOL.VARIANT_TRUE : VARIANT_BOOL.VARIANT_FALSE
                    ),
                },
            int lVal => new Unmanaged.Variant { vt = (ushort)VARTYPE.VT_I4, lVal = lVal, },
            string bstrVal
                => new Unmanaged.Variant
                {
                    vt = (ushort)VARTYPE.VT_BSTR,
                    bstrVal = Marshal.StringToBSTR(bstrVal),
                },
            null => new Unmanaged.Variant { },
            _ => throw new NotImplementedException(),
        };
    }

    public static unsafe Types.Managed.Variant ConvertToManaged(Unmanaged.Variant unmanaged)
    {
        var vt = (VARTYPE)unmanaged.vt;
        return vt switch
        {
            VARTYPE.VT_BOOL
                => new Types.Managed.Variant
                {
                    Value = unmanaged.boolVal == (short)VARIANT_BOOL.VARIANT_TRUE,
                },
            VARTYPE.VT_I4 => new Types.Managed.Variant { Value = unmanaged.lVal, },
            VARTYPE.VT_BSTR
                => new Types.Managed.Variant
                {
                    Value = Marshal.PtrToStringBSTR(unmanaged.bstrVal),
                },
            VARTYPE.VT_DISPATCH
                => new Types.Managed.Variant
                {
                    Value = ComInterfaceMarshaller<IDispatch>.ConvertToManaged(
                        (void*)unmanaged.pdispVal
                    ),
                },
            VARTYPE.VT_EMPTY => new Types.Managed.Variant { },
            _ => throw new NotImplementedException(),
        };
    }
}
