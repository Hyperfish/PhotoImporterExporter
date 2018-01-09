using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Hyperfish.ImportExport.Helpers
{
    public static class StringExtensions
    {
        public static SecureString Secure(this String str)
        {
            var secureString = new SecureString();

            foreach (char c in str.ToCharArray())
            {
                secureString.AppendChar(c);
            }

            return secureString;
        }

        public static String UnSecure(this SecureString value)
        {
            IntPtr valuePtr = IntPtr.Zero;
            try
            {
                valuePtr = Marshal.SecureStringToGlobalAllocUnicode(value);
                return Marshal.PtrToStringUni(valuePtr);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(valuePtr);
            }
        }
    }
}