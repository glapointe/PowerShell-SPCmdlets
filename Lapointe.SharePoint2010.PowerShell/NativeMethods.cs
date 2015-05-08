using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Lapointe.SharePoint.PowerShell
{
	internal class NativeMethods
	{
        public const int LOGON32_LOGON_INTERACTIVE = 2;
        public const int LOGON32_LOGON_SERVICE = 3;
        public const int LOGON32_PROVIDER_DEFAULT = 0;

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern bool LogonUser(
          String lpszUserName,
          String lpszDomain,
          String lpszPassword,
          int dwLogonType,
          int dwLogonProvider,
          ref IntPtr phToken
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public extern static bool CloseHandle(IntPtr handle);

	}
}
