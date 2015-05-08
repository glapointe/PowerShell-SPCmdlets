using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using System.Xml;
using System.IO;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class ImportListSecurity : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ImportListSecurity"/> class.
        /// </summary>
        public ImportListSecurity()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the web url import security to."));
            parameters.Add(new SPParam("inputfile", "file", true, null, new SPDirectoryExistsAndValidFileNameValidator()));
            parameters.Add(new SPParam("quiet", "quiet", false, null, null));
            parameters.Add(new SPParam("includeitemsecurity", "items"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nImports security settings using data output from gl-exportlistsecurity.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <url to import security to>");
            sb.Append("\r\n\t-inputfile <file to import settings from>");
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
            Logger.Verbose = !Params["quiet"].UserTypedIn;
            string inputFile = Params["inputfile"].Value;
            bool includeItemSecurity = Params["includeitemsecurity"].UserTypedIn;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputFile);
            Common.Lists.ImportListSecurity.ImportSecurity(xmlDoc, url, includeItemSecurity);
            return (int)ErrorCodes.NoError;
        }

        #endregion



    }
}
