﻿namespace Addin.CApi;

using Addin.Types.Unmanaged;
using System.Runtime.InteropServices;
using static Addin.Types.Unmanaged.ExcelConstants;

public static class Marshalling
{
    // String marshalling
    public static nint ToXlOper(this string str)
    {
        // Calculate sizes
        var strLen = str.Length + 1;
        var charLen = Marshal.SizeOf(str[0]);
        var byteCount = charLen * strLen;

        // Create byte array, with the length stored in the first char and the string contents on the rest
        var bytes = new char[strLen];
        bytes[0] = (char)(strLen - 1);
        Buffer.BlockCopy(str.ToCharArray(), 0, bytes, charLen, byteCount - charLen);

        // Convert to unmanaged bytes
        var strPtr = Marshal.AllocHGlobal(byteCount);
        Marshal.Copy(bytes, 0, strPtr, strLen);

        // Add to XlOper structure
        xloper12 lpx = new() { xltype = xltypeStr };
        lpx.val.str = strPtr;
        return lpx.ToPtr();
    }

    public static string? ToStringUnicode(this nint ptr)
    {
        var xlo = Marshal.PtrToStructure<xloper12>(ptr);
        if (xlo.xltype != xltypeStr)
            return null;
        var str = Marshal.PtrToStringUni(xlo.val.str);
        var bytes = str?.ToCharArray().Skip(1).ToArray();
        return new string(bytes);
    }

    // Array marshalling
    public static nint ToPtr(this nint[] array)
    {
        var size = Marshal.SizeOf(array[0]) * array.Length;

        var ptr = Marshal.AllocHGlobal(size);

        Marshal.Copy(array, 0, ptr, array.Length);

        return ptr;
    }

    // XlOper marshalling
    public static nint ToPtr(this xloper12 xlo)
    {
        var ptr = Marshal.AllocHGlobal(Marshal.SizeOf(xlo));

        Marshal.StructureToPtr(xlo, ptr, false);

        return ptr;
    }
}
