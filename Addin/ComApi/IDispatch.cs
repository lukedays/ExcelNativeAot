using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;
using Managed = Addin.ComApi.Types.Managed;

namespace Addin.ComApi;

[GeneratedComInterface]
[Guid("00020400-0000-0000-C000-000000000046")] // The IID for IDispatch
public partial interface IDispatch
{
    [PreserveSig]
    int GetTypeInfoCount(out uint pctinfo);

    [PreserveSig]
    int GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo);

    [PreserveSig]
    int GetIDsOfNames(
        ref Guid riid,
        [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames,
        uint cNames,
        uint lcid,
        [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId
    );

    [PreserveSig]
    int Invoke(
        [MarshalAs(UnmanagedType.I4)] int dispIdMember,
        Guid riid,
        [MarshalAs(UnmanagedType.U4)] uint lcid,
        INVOKEKIND wFlags,
        [MarshalUsing(typeof(StructMarshalling))] ref Managed.DISPPARAMS pDispParams,
        [MarshalUsing(typeof(StructMarshalling))] ref Managed.VARIANT pVarResult,
        [MarshalUsing(typeof(StructMarshalling))] ref Managed.EXCEPINFO pExcepInfo,
        ref uint puArgErr
    );
}
