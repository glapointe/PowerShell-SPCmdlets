using System;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using System.Xml;
using System.IO;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class ExportListSecurity : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExportListSecurity"/> class.
        /// </summary>
        public ExportListSecurity()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the list view URL to export security from."));
            parameters.Add(new SPParam("outputfile", "file", true, null, new SPDirectoryExistsAndValidFileNameValidator()));
            parameters.Add(new SPParam("quiet", "quiet", false, null, null));
            parameters.Add(new SPParam("scope", "s", true, "Web", new SPRegexValidator("(?i:^Web$|^List$)")));
            parameters.Add(new SPParam("includeitemsecurity", "items"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nExports a list's security settings to an XML file.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <list view URL or Web URL to export security from>");
            sb.Append("\r\n\t-outputfile <file to output settings to>");
            sb.Append("\r\n\t-scope <Web | List>");
            sb.Append("\r\n\t[-quiet]");
            sb.Append("\r\n\t[-includeitemsecurity]");
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
            string scope = Params["scope"].Value.ToLowerInvariant();
            string outputFile = Params["outputfile"].Value;
            bool includeItemSecurity = Params["includeitemsecurity"].UserTypedIn;

            Logger.Verbose = !Params["quiet"].UserTypedIn;

            Common.Lists.ExportListSecurity.ExportSecurity(outputFile, scope, url, includeItemSecurity);

            return (int)ErrorCodes.NoError;
        }

        #endregion

    }
}
