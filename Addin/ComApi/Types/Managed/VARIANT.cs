using System.Runtime.InteropServices;

namespace Addin.ComApi.Types.Managed;

public struct VARIANT
{
    public DECIMAL decVal;
    public VARTYPE vt;
    public ushort wReserved1;
    public ushort wReserved2;
    public ushort wReserved3;
    public long llVal;
    public int lVal;
    public byte bVal;
    public short iVal;
    public float fltVal;
    public double dblVal;
    public bool boolVal;
    public bool __OBSOLETE__VARIANT_BOOL;
    public int scode;
    public CY cyVal;
    public double date;
    public string bstrVal;
    public nint punkVal;
    public nint pdispVal;
    public nint parray;
    public nint pbVal;
    public nint piVal;
    public nint plVal;
    public nint pllVal;
    public nint pfltVal;
    public nint pdblVal;
    public nint pboolVal;
    public nint __OBSOLETE__VARIANT_PBOOL;
    public nint pscode;
    public nint pcyVal;
    public nint pdate;
    public nint pbstrVal;
    public nint ppunkVal;
    public nint ppdispVal;
    public nint pparray;
    public nint pvarVal;
    public nint byref;
    public sbyte cVal;
    public ushort uiVal;
    public uint ulVal;
    public ulong ullVal;
    public int intVal;
    public uint uintVal;
    public nint pdecVal;
    public nint pcVal;
    public nint puiVal;
    public nint pulVal;
    public nint pullVal;
    public nint pintVal;
    public nint puintVal;
    public __tagBRECORD __tagBRECORD;
}

[StructLayout(LayoutKind.Explicit)]
public struct CY
{
    [FieldOffset(0)]
    public uint Lo;

    [FieldOffset(4)]
    public int Hi;

    [FieldOffset(0)]
    public long int64;
}

[StructLayout(LayoutKind.Sequential)]
public struct __tagBRECORD
{
    public nint pvRecord;
    public nint pRecInfo;
}

[StructLayout(LayoutKind.Explicit)]
public struct DECIMAL
{
    [FieldOffset(0)]
    public ushort wReserved;

    [FieldOffset(2)]
    public ushort signscale;

    [FieldOffset(2)]
    public byte scale;

    [FieldOffset(3)]
    public byte sign;

    [FieldOffset(4)]
    public uint Hi32;

    [FieldOffset(8)]
    public ulong Lo64;

    [FieldOffset(8)]
    public uint Lo32;

    [FieldOffset(12)]
    public uint Mid32;
}

public enum VARIANT_BOOL : short
{
    VARIANT_FALSE = 0,
    VARIANT_TRUE = 5
}

public enum VARTYPE : ushort
{
    VT_EMPTY = 0,
    VT_NULL = 1,
    VT_I2 = 2,
    VT_I4 = 3,
    VT_R4 = 4,
    VT_R8 = 5,
    VT_CY = 6,
    VT_DATE = 7,
    VT_BSTR = 8,
    VT_DISPATCH = 9,
    VT_ERROR = 10,
    VT_BOOL = 11,
    VT_VARIANT = 12,
    VT_UNKNOWN = 13,
    VT_DECIMAL = 14,
    VT_I1 = 16,
    VT_UI1 = 17,
    VT_UI2 = 18,
    VT_UI4 = 19,
    VT_I8 = 20,
    VT_UI8 = 21,
    VT_INT = 22,
    VT_UINT = 23,
    VT_VOID = 24,
    VT_HRESULT = 25,
    VT_PTR = 26,
    VT_SAFEARRAY = 27,
    VT_CARRAY = 28,
    VT_USERDEFINED = 29,
    VT_LPSTR = 30,
    VT_LPWSTR = 31,
    VT_RECORD = 36,
    VT_INT_PTR = 37,
    VT_UINT_PTR = 38,
    VT_FILETIME = 64,
    VT_BLOB = 65,
    VT_STREAM = 66,
    VT_STORAGE = 67,
    VT_STREAMED_OBJECT = 68,
    VT_STORED_OBJECT = 69,
    VT_BLOB_OBJECT = 70,
    VT_CF = 71,
    VT_CLSID = 72,
    VT_VERSIONED_STREAM = 73,
    VT_BSTR_BLOB = 0xfff,
    VT_VECTOR = 0x1000,
    VT_ARRAY = 0x2000,
    VT_BYREF = 0x4000,
    VT_RESERVED = 0x8000,
    VT_X = 0xffff,
    VT_Y = 0xfff,
    VT_TYPEMASK = 0xfff
};
