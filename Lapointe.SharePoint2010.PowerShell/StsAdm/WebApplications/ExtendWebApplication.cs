using System;
using System.DirectoryServices;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebApplications
{
    public class ExtendWebApplication : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendWebApplication"/> class.
        /// </summary>
        public ExtendWebApplication()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator()));
            parameters.Add(new SPParam("vsname", "vsname", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("allowanonymous", "anon"));
            parameters.Add(new SPParam("exclusivelyusentlm", "ntlm"));
            parameters.Add(new SPParam("usessl", "ssl"));
            parameters.Add(new SPParam("hostheader", "hostheader", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("port", "p", false, "80", new SPIntRangeValidator(0, int.MaxValue)));
            parameters.Add(new SPParam("path", "path", true, null, new SPNonEmptyValidator()));
            SPEnumValidator zoneValidator = new SPEnumValidator(typeof (SPUrlZone));
            parameters.Add(new SPParam("zone", "zone", false, SPUrlZone.Custom.ToString(), zoneValidator));
            parameters.Add(new SPParam("loadbalancedurl", "lburl", true, null, new SPUrlValidator()));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nExtends a web application onto another IIS web site.  This allows you to serve the same content on another port or to a different audience\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <url of the web application to extend>");
            sb.Append("\r\n\t-vsname <web application name>");
            sb.Append("\r\n\t-path <path>");
            sb.Append("\r\n\t-loadbalancedurl <the load balanced URL is the domain name for all sites users will access in this SharePoint Web application>");
            sb.AppendFormat("\r\n\t[-zone <{0} (defaults to Custom)>]", zoneValidator.DisplayValue);
            sb.Append("\r\n\t[-port <port number (default is 80)>]");
            sb.Append("\r\n\t[-hostheader <host header>]");
            sb.Append("\r\n\t[-exclusivelyusentlm]");
            sb.Append("\r\n\t[-allowanonymous]");
            sb.Append("\r\n\t[-usessl]");

            Init(parameters, sb.ToString());

        }

        /// <summary>
        /// Gets the help message.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <returns></returns>
        public override string GetHelpMessage(string command)
        {
            return HelpMessage;
        }

        /// <summary>
        /// Runs the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public override int Execute(string command, System.Collections.Specialized.StringDictionary keyValues, out string output)
        {
            output = string.Empty;

            SPWebApplication webApplication = SPWebApplication.Lookup(new Uri(Params["url"].Value.TrimEnd('/')));
            string description = Params["vsname"].Value;
            bool useSsl = Params["usessl"].UserTypedIn;
            string hostHeader = Params["hostheader"].Value;
            int port = int.Parse(Params["port"].Value);
            bool allowAnonymous = Params["allowanonymous"].UserTypedIn;
            bool useNtlm = Params["exclusivelyusentlm"].UserTypedIn;
            string path = Params["path"].Value;
            SPUrlZone zone = (SPUrlZone) Enum.Parse(typeof (SPUrlZone), Params["zone"].Value, true);
            string loadBalancedUrl = Params["loadbalancedurl"].Value;

            ExtendWebApp(webApplication, description, hostHeader, port, loadBalancedUrl, path, allowAnonymous, useNtlm, useSsl, zone);

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Extends the web app.
        /// </summary>
        /// <param name="webApplication">The web application.</param>
        /// <param name="description">The description.</param>
        /// <param name="hostHeader">The host header.</param>
        /// <param name="port">The port.</param>
        /// <param name="loadBalancedUrl">The load balanced URL.</param>
        /// <param name="path">The path.</param>
        /// <param name="allowAnonymous">if set to <c>true</c> [allow anonymous].</param>
        /// <param name="useNtlm">if set to <c>true</c> [use NTLM].</param>
        /// <param name="useSsl">if set to <c>true</c> [use SSL].</param>
        /// <param name="zone">The zone.</param>
        public static void ExtendWebApp(SPWebApplication webApplication, string description, string hostHeader, int port, string loadBalancedUrl, string path, bool allowAnonymous, bool useNtlm, bool useSsl, SPUrlZone zone)
        {
            SPServerBinding serverBinding = null;
            SPSecureBinding secureBinding = null;
            if (!useSsl)
            {
                serverBinding = new SPServerBinding();
                serverBinding.Port = port;
                serverBinding.HostHeader = hostHeader;
            }
            else
            {
                secureBinding = new SPSecureBinding();
                secureBinding.Port = port;
            }

            SPIisSettings settings = new SPIisSettings(description, allowAnonymous, useNtlm, serverBinding, secureBinding, new DirectoryInfo(path.Trim()));
            settings.PreferredInstanceId = GetPreferredInstanceId(description);

            webApplication.IisSettings.Add(zone, settings);
            webApplication.AlternateUrls.SetResponseUrl(new SPAlternateUrl(new Uri(loadBalancedUrl), zone));
            webApplication.AlternateUrls.Update();
            webApplication.Update();
            webApplication.ProvisionGlobally();
        }

        /// <summary>
        /// Gets the preferred instance id.
        /// </summary>
        /// <param name="iisServerComment">The IIS server comment.</param>
        /// <returns></returns>
        private static int GetPreferredInstanceId(string iisServerComment)
        {
            try
            {
                int num;
                if (!LookupByServerComment(iisServerComment, out num))
                {
                    return GetUnusedInstanceId(0);
                }
                return num;
            }
            catch
            {
                return GetUnusedInstanceId(0);
            }
        }

        /// <summary>
        /// Lookups the by server comment.
        /// </summary>
        /// <param name="serverComment">The server comment.</param>
        /// <param name="instanceId">The instance id.</param>
        /// <returns></returns>
        private static bool LookupByServerComment(string serverComment, out int instanceId)
        {
            instanceId = -1;
            using (DirectoryEntry entry = new DirectoryEntry("IIS://localhost/w3svc"))
            {
                foreach (DirectoryEntry entry2 in entry.Children)
                {
                    if (entry2.SchemaClassName != "IIsWebServer")
                    {
                        continue;
                    }
                    string str = (string) entry2.Properties["ServerComment"].Value;
                    if (!Utilities.StsCompareStrings(str, serverComment))
                        continue;

                    instanceId = int.Parse(entry2.Name, NumberFormatInfo.InvariantInfo);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Gets the unused instance id.
        /// </summary>
        /// <param name="preferredInstanceId">The preferred instance id.</param>
        /// <returns></returns>
        private static int GetUnusedInstanceId(int preferredInstanceId)
        {
            Random random = new Random();
            int num = 0;
            int num2 = preferredInstanceId;
            if (num2 < 1)
            {
                num2 = random.Next(1, 0x7fffffff);
            }

            while (true)
            {
                if (++num >= 0x19)
                {
                    throw new InvalidOperationException(SPResource.GetString("CannotFindUnusedInstanceId", new object[0]));
                }
                if (DirectoryEntry.Exists("IIS://localhost/w3svc/" + num2))
                {
                    num2 = random.Next(1, 0x7fffffff);
                }
                else
                    break;
            }
            return num2;
        }
    }
}
