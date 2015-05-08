using System;
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
    public class UnExtendWebApp : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UnExtendWebApp"/> class.
        /// </summary>
        public UnExtendWebApp()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator()));
            parameters.Add(new SPParam("deleteiiswebsite", "iis"));
            SPEnumValidator urlZoneValidator = new SPEnumValidator(typeof (SPUrlZone));
            parameters.Add(new SPParam("zone", "z", true, null, urlZoneValidator));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nDeletes a web application.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <url>");
            sb.AppendFormat("\r\n\t-zone <{0}>", urlZoneValidator.DisplayValue);
            sb.Append("\r\n\t[-deleteiiswebsite]");

            Init(parameters, sb.ToString());
        }

        #region ISPStsadmCommand Members

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

            

            string url = Params["url"].Value;
            bool deleteIis = Params["deleteiiswebsite"].UserTypedIn;

            SPWebApplication webApp = SPWebApplication.Lookup(new Uri(url));
            SPUrlZone zone = (SPUrlZone) Enum.Parse(typeof (SPUrlZone), Params["zone"].Value, true);


            UnExtend(webApp, zone, deleteIis);


            return (int)ErrorCodes.NoError;
        }
        #endregion

        /// <summary>
        /// Uns the extend.
        /// </summary>
        /// <param name="webApp">The web app.</param>
        /// <param name="zone">The zone.</param>
        /// <param name="deleteIis">if set to <c>true</c> [delete IIS].</param>
        public static void UnExtend(SPWebApplication webApp, SPUrlZone zone, bool deleteIis)
        {

            webApp.UnprovisionGlobally(deleteIis);

            webApp.IisSettings.Remove(zone);
            if (zone != SPUrlZone.Default)
            {
                webApp.AlternateUrls.UnsetResponseUrl(zone);
                webApp.AlternateUrls.Update();
            }
            webApp.Update();
        }
    }
}
