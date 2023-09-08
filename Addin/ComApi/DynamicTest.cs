using System.Dynamic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using static Addin.ComApi.Constants;
using Managed = Addin.ComApi.Types.Managed;

namespace Addin.ComApi;

public class DynamicTest : DynamicObject
{
    IDispatch? excel;
    Guid emptyGuid = Guid.Empty;

    public DynamicTest()
    {
        // The CLSID for Excel.Application (COMView.exe->CLSID table)
        var clsid = new Guid("{00024500-0000-0000-C000-000000000046}");

        // COMView.exe -> CLSID table -> Type column
        var server = CLSCTX.CLSCTX_LOCAL_SERVER;

        excel = Activator.ActivateClass(clsid, server);
    }

    public override bool TryGetMember(GetMemberBinder binder, out object result)
    {
        var propName = binder.Name;
        var dispIds = GetDispIDs(propName);

        Managed.DISPPARAMS dispParams = new();
        Managed.EXCEPINFO excep = new();
        Managed.VARIANT pVarResult = new();
        uint puArg = 0;

        var hr = excel.Invoke(
            dispIds[0],
            emptyGuid,
            LOCALE_USER_DEFAULT,
            INVOKEKIND.INVOKE_PROPERTYGET,
            ref dispParams,
            ref pVarResult,
            ref excep,
            ref puArg
        );

        Marshal.ThrowExceptionForHR(hr);

        result = pVarResult.boolVal;

        return true;
    }

    private int[] GetDispIDs(string propName)
    {
        var names = new string[] { propName };

        var dispIds = new int[names.Length];
        var hr = excel.GetIDsOfNames(
            ref emptyGuid,
            names,
            (uint)names.Length,
            LOCALE_USER_DEFAULT,
            dispIds
        );

        Marshal.ThrowExceptionForHR(hr);

#if DEBUG
        for (int i = 0; i < names.Length; i++)
            Console.WriteLine($"{names[i]}: {dispIds[i]}");
#endif
        return dispIds;
    }

    public override bool TrySetMember(SetMemberBinder binder, object value)
    {
        var propName = binder.Name;
        var dispIds = GetDispIDs(propName);

        var dispParams = new Managed.DISPPARAMS
        {
            rgvarg = new Managed.VARIANT[] { Helpers.BoolToVariant((bool)value) },
            rgdispidNamedArgs = DISPID_PROPERTYPUT,
            cArgs = 1,
            cNamedArgs = 1
        };

        Managed.EXCEPINFO excep = new();
        Managed.VARIANT pVarResult = new();
        uint puArg = 0;

        var hr = excel.Invoke(
            dispIds[0],
            emptyGuid,
            LOCALE_USER_DEFAULT,
            INVOKEKIND.INVOKE_PROPERTYPUT,
            ref dispParams,
            ref pVarResult,
            ref excep,
            ref puArg
        );

        Marshal.ThrowExceptionForHR(hr);

        return true;
    }
}
