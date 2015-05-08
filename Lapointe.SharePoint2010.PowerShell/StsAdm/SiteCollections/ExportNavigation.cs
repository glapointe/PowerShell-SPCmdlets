using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    public class ExportNavigation : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExportNavigation"/> class.
        /// </summary>
        public ExportNavigation()
        {
            SPParamCollection parameters = new SPParamCollection();
            StringBuilder sb = new StringBuilder();

#if MOSS
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site collection."));
            parameters.Add(new SPParam("outputfile", "output", true, null, new SPDirectoryExistsAndValidFileNameValidator(), "Make sure the output directory exists and a valid filename is provided."));
            parameters.Add(new SPParam("scope", "s", false, "site", new SPRegexValidator("(?i:^Site$|^Web$)")));

            sb.Append("\r\n\r\nExports the site navigation for all publishing sites within a site collection or for a specific web.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <site collection url>");
            sb.Append("\r\n\t-outputfile <file to output results to>");
            sb.Append("\r\n\t[-scope <Site | Web> (defaults to Site)]");
#else
            sb.Append(NOT_VALID_FOR_FOUNDATION);
#endif

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
#if !MOSS
            output = NOT_VALID_FOR_FOUNDATION;
            return (int)ErrorCodes.GeneralError;
#endif

            string url = Params["url"].Value.TrimEnd('/');
            string scope = Params["scope"].Value.ToLowerInvariant();

            XmlDocument xmlDoc = Common.SiteCollections.ExportNavigation.GetNavigation(url, scope);
            File.WriteAllText(Params["outputfile"].Value, Utilities.GetFormattedXml(xmlDoc));

            return (int)ErrorCodes.NoError;
        }

        #endregion


     
    }
}
