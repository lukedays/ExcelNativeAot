using Addin.ComApi.Types.Managed;
using System.Runtime.InteropServices;

namespace Addin.ComApi
{
    internal static class Helpers
    {
        public static VARIANT IntToVariant(int val)
        {
            return new VARIANT { vt = VARTYPE.VT_I4, lVal = val };
        }

        public static VARIANT BoolToVariant(bool val)
        {
            return new VARIANT { vt = VARTYPE.VT_BOOL, boolVal = val };
        }

        public static nint StructureToPtr<T>(T str)
        {
            var ptr = Marshal.AllocHGlobal(Marshal.SizeOf(str));
            Marshal.StructureToPtr(str, ptr, false);
            return ptr;
        }

        public static nint ArrayToPtr<T>(T[] str)
        {
            var size = Marshal.SizeOf<T>();
            var len = str.Length;
            var ptrs = new nint[str.Length];
            var basePtr = Marshal.AllocHGlobal(size * len);
            for (int i = 0; i < len; ++i)
            {
                ptrs[i] = nint.Add(basePtr, i * size);
                Marshal.StructureToPtr(str[i], ptrs[i], false);
            }
            return basePtr;
        }

        public static T[] PtrToArray<T>(nint str, int len)
        {
            var size = Marshal.SizeOf<T>();
            var ret = new T[len];
            for (int i = 0; i < len; ++i)
            {
                ret[i] = Marshal.PtrToStructure<T>(nint.Add(str, i * size));
            }
            return ret;
        }
    }
}
