using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    public class EnumInstalledSiteTemplates : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EnumInstalledSiteTemplates"/> class.
        /// Use Get-SPWebTemplate instead.
        /// </summary>
        public EnumInstalledSiteTemplates()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site collection"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nReturns the list of site templates installed for the given site collection.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <site collection url>");

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
        public override int Execute(string command, StringDictionary keyValues, out string output)
        {
            output = string.Empty;

            

            string url = Params["url"].Value.TrimEnd('/');

            using (SPSite site = new SPSite(url))
            {
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {

                    foreach (SPLanguage lang in web.RegionalSettings.InstalledLanguages)
                    {
                        foreach (SPWebTemplate template in site.GetWebTemplates((uint)lang.LCID))
                        {
                            output += template.Name + " = " + template.Title + " (" + lang.LCID + ")\r\n";
                        }
                        foreach (SPWebTemplate template in site.GetCustomWebTemplates((uint)lang.LCID))
                        {
                            output += template.Name + " = " + template.Title + " (Custom)(" + lang.LCID + ")\r\n";
                        }
                    }
                }
            }

            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
