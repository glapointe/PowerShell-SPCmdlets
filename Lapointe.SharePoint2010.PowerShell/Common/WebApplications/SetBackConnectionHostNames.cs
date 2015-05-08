using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.Win32;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation.Internal;

namespace Lapointe.SharePoint.PowerShell.Common.WebApplications
{
    internal class SetBackConnectionHostNames
    {
        public static List<string> GetUrls()
        {
            List<string> urls = new List<string>();
            foreach (SPService svc in SPFarm.Local.Services)
            {
                if (!(svc is SPWebService))
                    continue;

                foreach (SPWebApplication webApp in ((SPWebService)svc).WebApplications)
                {

                    GetUrls(urls, webApp);
                }
            }
            return urls;
        }

        internal static void GetUrls(List<string> urls, SPWebApplication webApp)
        {
            foreach (SPAlternateUrl url in webApp.AlternateUrls)
            {
                GetUrls(urls, url);
            }
        }

        internal static void GetUrls(List<string> urls, SPAlternateUrl url)
        {
            string host = url.Uri.Host.ToLower();
            if (!urls.Contains(host) && // Don't add if we already have it
                !url.Uri.IsLoopback && // Quick check to short circuit the more elaborate checks
                host != Environment.MachineName.ToLower() && // Quick check to short circuit the more elaborate checks
                IsLocalIpAddress(host) && // If the host name points locally then we need to add it
                !IsSharePointServer(host)) // Don't add if it matches an SP server name (handles central admin)
            {
                urls.Add(host);
            }
        }

        internal static bool IsSharePointServer(string host)
        {
            foreach (SPServer server in SPFarm.Local.Servers)
            {
                if (server.Address.ToLower() == host)
                    return true;
            }
            return false;
        }

        internal static bool IsLocalIpAddress(string host)
        {
            try
            {
                IPAddress[] hostIPs = Dns.GetHostAddresses(host);
                IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());

                // test if any host IP equals to any local IP or to localhost
                foreach (IPAddress hostIP in hostIPs)
                {
                    // is localhost
                    if (IPAddress.IsLoopback(hostIP)) return true;
                    // is local address
                    foreach (IPAddress localIP in localIPs)
                    {
                        if (hostIP.Equals(localIP)) return true;
                    }
                }
            }
            catch { }
            return false;
        }

        public static void SetBackConnectionRegKey(List<string> urls)
        {
            const string KEY_NAME = "SYSTEM\\CurrentControlSet\\Control\\Lsa\\MSV1_0";
            const string KEY_VAL_NAME = "BackConnectionHostNames";

            RegistryKey reg = Registry.LocalMachine.OpenSubKey(KEY_NAME, true);
            if (reg != null)
            {
                string[] existing = (string[])reg.GetValue(KEY_VAL_NAME);
                if (existing != null)
                {
                    foreach (string val in existing)
                    {
                        if (!urls.Contains(val.ToLower()))
                            urls.Add(val.ToLower());
                    }
                }
                string[] multiVal = new string[urls.Count];
                urls.CopyTo(multiVal);
                foreach (string url in urls)
                {
                    Logger.Write("Setting {0}", url);
                }
                reg.SetValue(KEY_VAL_NAME, multiVal, RegistryValueKind.MultiString);
            }
            else
            {
                throw new SPException("Unable to open registry key.");
            }
        }
    }
}
