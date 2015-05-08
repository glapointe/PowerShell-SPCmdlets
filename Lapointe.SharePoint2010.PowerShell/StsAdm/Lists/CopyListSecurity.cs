using System;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class CopyListSecurity : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CopyListSecurity"/> class.
        /// </summary>
        public CopyListSecurity()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("sourceurl", "sourceurl", true, null, new SPUrlValidator(), "Please specify the list view URL to copy security from."));
            parameters.Add(new SPParam("targeturl", "targeturl", true, null, new SPUrlValidator(), "Please specify the list view URL to copy security to."));
            parameters.Add(new SPParam("quiet", "quiet", false, null, null));
            parameters.Add(new SPParam("includeitemsecurity", "items"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nCopies a list's security settings from one list to another.\r\n\r\nParameters:\r\n\t");
            sb.Append("-sourceurl <list view url to copy security from>\r\n\t");
            sb.Append("-targeturl <list view url to copy security to>\r\n\t");
            sb.Append("[-quiet]\r\n\t");
            sb.Append("[-includeitemsecurity]");
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

            string sourceUrl = Params["sourceurl"].Value;
            string targetUrl = Params["targeturl"].Value;
            bool quiet = Params["quiet"].UserTypedIn;
            bool includeItemSecurity = Params["includeitemsecurity"].UserTypedIn;

            using (SPSite sourceSite = new SPSite(sourceUrl))
            using (SPSite targetSite = new SPSite(targetUrl))
            using (SPWeb sourceWeb = sourceSite.OpenWeb())
            using (SPWeb targetWeb = targetSite.OpenWeb())
            {
                SPList sourceList = Utilities.GetListFromViewUrl(sourceWeb, sourceUrl);
                SPList targetList = Utilities.GetListFromViewUrl(targetWeb, targetUrl);

                if (sourceList == null)
                    throw new SPException("Source list was not found.");
                if (targetList == null)
                    throw new SPException("Target list was not found.");

                Common.Lists.CopyListSecurity.CopySecurity(sourceList, targetList, targetWeb, includeItemSecurity, quiet);
            }
            return (int)ErrorCodes.NoError;
        }


        #endregion

    }
}
