using Addin.ComApi.Types.Managed;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Unmanaged = Addin.ComApi.Types.Unmanaged;

namespace Addin.ComApi;

[CustomMarshaller(typeof(VARIANT), MarshalMode.Default, typeof(VARIANT_))]
[CustomMarshaller(typeof(DISPPARAMS), MarshalMode.Default, typeof(DISPPARAMS_))]
[CustomMarshaller(typeof(EXCEPINFO), MarshalMode.Default, typeof(EXCEPINFO_))]
public static partial class StructMarshalling
{
    public static class VARIANT_
    {
        public static Unmanaged.VARIANT ConvertToUnmanaged(VARIANT managed)
        {
            return managed.vt switch
            {
                VARTYPE.VT_BOOL
                    => new Unmanaged.VARIANT
                    {
                        vt = (ushort)managed.vt,
                        boolVal = (short)(
                            managed.boolVal ? VARIANT_BOOL.VARIANT_TRUE : VARIANT_BOOL.VARIANT_FALSE
                        ),
                    },
                VARTYPE.VT_EMPTY => new Unmanaged.VARIANT { },
                _ => throw new NotImplementedException(),
            };
            //var a = new Unmanaged.VARIANT
            //{
            //    decVal = managed.decVal,
            //    vt = (ushort)managed.vt,
            //    wReserved1 = managed.wReserved1,
            //    wReserved2 = managed.wReserved2,
            //    wReserved3 = managed.wReserved3,
            //    llVal = managed.llVal,
            //    lVal = managed.lVal,
            //    bVal = managed.bVal,
            //    iVal = managed.iVal,
            //    fltVal = managed.fltVal,
            //    dblVal = managed.dblVal,
            //    boolVal = (short)managed.boolVal,
            //    __OBSOLETE__VARIANT_BOOL = (short)managed.__OBSOLETE__VARIANT_BOOL,
            //    scode = managed.scode,
            //    cyVal = managed.cyVal,
            //    date = managed.date,
            //    bstrVal = Marshal.StringToBSTR(managed.bstrVal),
            //    punkVal = managed.punkVal,
            //    pdispVal = managed.pdispVal,
            //    parray = managed.parray,
            //    pbVal = managed.pbVal,
            //    piVal = managed.piVal,
            //    plVal = managed.plVal,
            //    pllVal = managed.pllVal,
            //    pfltVal = managed.pfltVal,
            //    pdblVal = managed.pdblVal,
            //    pboolVal = managed.pboolVal,
            //    __OBSOLETE__VARIANT_PBOOL = managed.__OBSOLETE__VARIANT_PBOOL,
            //    pscode = managed.pscode,
            //    pcyVal = managed.pcyVal,
            //    pdate = managed.pdate,
            //    pbstrVal = managed.pbstrVal,
            //    ppunkVal = managed.ppunkVal,
            //    ppdispVal = managed.ppdispVal,
            //    pparray = managed.pparray,
            //    pvarVal = managed.pvarVal,
            //    byref = managed.byref,
            //    cVal = managed.cVal,
            //    uiVal = managed.uiVal,
            //    ulVal = managed.ulVal,
            //    ullVal = managed.ullVal,
            //    intVal = managed.intVal,
            //    uintVal = managed.uintVal,
            //    pdecVal = managed.pdecVal,
            //    pcVal = managed.pcVal,
            //    puiVal = managed.puiVal,
            //    pulVal = managed.pulVal,
            //    pullVal = managed.pullVal,
            //    pintVal = managed.pintVal,
            //    puintVal = managed.puintVal,
            //    //__tagBRECORD = managed.__tagBRECORD,
            //};
        }

        public static VARIANT ConvertToManaged(Unmanaged.VARIANT unmanaged)
        {
            return new VARIANT
            {
                decVal = unmanaged.decVal,
                vt = (VARTYPE)unmanaged.vt,
                wReserved1 = unmanaged.wReserved1,
                wReserved2 = unmanaged.wReserved2,
                wReserved3 = unmanaged.wReserved3,
                llVal = unmanaged.llVal,
                lVal = unmanaged.lVal,
                bVal = unmanaged.bVal,
                iVal = unmanaged.iVal,
                fltVal = unmanaged.fltVal,
                dblVal = unmanaged.dblVal,
                boolVal = unmanaged.boolVal == (short)VARIANT_BOOL.VARIANT_TRUE,
                __OBSOLETE__VARIANT_BOOL =
                    unmanaged.__OBSOLETE__VARIANT_BOOL == (short)VARIANT_BOOL.VARIANT_TRUE,
                scode = unmanaged.scode,
                cyVal = unmanaged.cyVal,
                date = unmanaged.date,
                bstrVal =
                    unmanaged.vt == (ushort)VARTYPE.VT_BSTR
                        ? Marshal.PtrToStringBSTR(unmanaged.bstrVal)
                        : "",
                punkVal = unmanaged.punkVal,
                pdispVal = unmanaged.pdispVal,
                parray = unmanaged.parray,
                pbVal = unmanaged.pbVal,
                piVal = unmanaged.piVal,
                plVal = unmanaged.plVal,
                pllVal = unmanaged.pllVal,
                pfltVal = unmanaged.pfltVal,
                pdblVal = unmanaged.pdblVal,
                pboolVal = unmanaged.pboolVal,
                __OBSOLETE__VARIANT_PBOOL = unmanaged.__OBSOLETE__VARIANT_PBOOL,
                pscode = unmanaged.pscode,
                pcyVal = unmanaged.pcyVal,
                pdate = unmanaged.pdate,
                pbstrVal = unmanaged.pbstrVal,
                ppunkVal = unmanaged.ppunkVal,
                ppdispVal = unmanaged.ppdispVal,
                pparray = unmanaged.pparray,
                pvarVal = unmanaged.pvarVal,
                byref = unmanaged.byref,
                cVal = unmanaged.cVal,
                uiVal = unmanaged.uiVal,
                ulVal = unmanaged.ulVal,
                ullVal = unmanaged.ullVal,
                intVal = unmanaged.intVal,
                uintVal = unmanaged.uintVal,
                pdecVal = unmanaged.pdecVal,
                pcVal = unmanaged.pcVal,
                puiVal = unmanaged.puiVal,
                pulVal = unmanaged.pulVal,
                pullVal = unmanaged.pullVal,
                pintVal = unmanaged.pintVal,
                puintVal = unmanaged.puintVal,
                __tagBRECORD = unmanaged.__tagBRECORD,
            };
        }
    }

    public static partial class DISPPARAMS_
    {
        public static unsafe Unmanaged.DISPPARAMS ConvertToUnmanaged(DISPPARAMS managed)
        {
            return new Unmanaged.DISPPARAMS
            {
                cArgs = managed.cArgs,
                cNamedArgs = managed.cNamedArgs,
                rgdispidNamedArgs = &managed.rgdispidNamedArgs,
                rgvarg =
                    managed.rgvarg != null
                        ? Helpers.ArrayToPtr(
                            managed.rgvarg.Select(VARIANT_.ConvertToUnmanaged).ToArray()
                        )
                        : nint.Zero
            };
        }

        public static unsafe DISPPARAMS ConvertToManaged(Unmanaged.DISPPARAMS unmanaged)
        {
            return new DISPPARAMS
            {
                cArgs = unmanaged.cArgs,
                cNamedArgs = unmanaged.cNamedArgs,
                rgdispidNamedArgs = *unmanaged.rgdispidNamedArgs,
                rgvarg = Helpers
                    .PtrToArray<Unmanaged.VARIANT>(unmanaged.rgvarg, unmanaged.cArgs)
                    .Select(VARIANT_.ConvertToManaged)
                    .ToArray(),
            };
        }
    }

    public static partial class EXCEPINFO_
    {
        public static Unmanaged.EXCEPINFO ConvertToUnmanaged(EXCEPINFO managed)
        {
            return new Unmanaged.EXCEPINFO
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

        public static EXCEPINFO ConvertToManaged(Unmanaged.EXCEPINFO unmanaged)
        {
            return new EXCEPINFO
            {
                bstrDescription =
                    unmanaged.bstrDescription != 0
                        ? Marshal.PtrToStringBSTR(unmanaged.bstrDescription)
                        : "",
                bstrHelpFile =
                    unmanaged.bstrHelpFile != 0
                        ? Marshal.PtrToStringBSTR(unmanaged.bstrHelpFile)
                        : "",
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
}
