using System;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Deployment;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class ImportList : SPOperation
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="ImportList"/> class.
        /// </summary>
        public ImportList()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the URL to import the list to."));
            parameters.Add(new SPParam("filename", "f", true, null, new SPNonEmptyValidator(), "Please specify the filename."));
            parameters.Add(new SPParam("quiet", "quiet"));
            parameters.Add(new SPParam("includeusersecurity", "security"));
            parameters.Add(new SPParam("haltonwarning", "warning"));
            parameters.Add(new SPParam("haltonfatalerror", "error"));
            parameters.Add(new SPParam("nologfile", "nolog"));
            parameters.Add(new SPParam("updateversions", "updatev", false, SPUpdateVersions.Append.ToString(), new SPIntRangeValidator(1, 3), "Please specify the updateversions setting."));
            parameters.Add(new SPParam("nofilecompression", "nofilecompression"));
            parameters.Add(new SPParam("retargetlinks", "retargetlinks"));
            parameters.Add(new SPParam("sourceurl", "source", false, null, new SPUrlValidator(), "Please specify the URL of the source list (must be on the same farm)."));
            parameters.Add(new SPParam("retainobjectidentity", "retainid"));
            parameters.Add(new SPParam("copysecuritysettings", "copysecurity"));
            parameters.Add(new SPParam("suppressafterevents", "sae"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nImports a list.\r\n\r\nParameters:\r\n\t");
            sb.Append("-url <url of a web site to import to>\r\n\t");
            sb.Append("-filename <import file name>\r\n\t");
            sb.Append("[-includeusersecurity]\r\n\t");
            sb.Append("[-haltonwarning]\r\n\t");
            sb.Append("[-haltonfatalerror]\r\n\t");
            sb.Append("[-nologfile]\r\n\t");
            sb.Append("[-updateversions <1-3>\r\n");
            sb.Append("\t\t1 - Add new versions to the current file (default)\r\n");
            sb.Append("\t\t2 - Overwrite the file and all its versions (delete then insert)\r\n");
            sb.Append("\t\t3 - Ignore the file if it exists on the destination]\r\n\t");
            sb.Append("[-nofilecompression]\r\n\t");
            sb.Append("[-quiet]\r\n\t");
            sb.Append("[-retargetlinks (resets links pointing to the source to now point to the target)]\r\n\t");
            sb.Append("[-sourceurl <url to a view of the original list> (use if retargetlinks)]\r\n\t");
            sb.Append("[-retainobjectidentity]\r\n\t");
            sb.Append("[-suppressafterevents (disable the firing of \"After\" events when creating or modifying list items)]\r\n\t");
            sb.Append("[-copysecuritysettings (must provide sourceurl and includeusersecurity)]");
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

            

            bool compressFile = !Params["nofilecompression"].UserTypedIn;
            string filename = Params["filename"].Value;
            bool quiet = Params["quiet"].UserTypedIn;
            bool haltOnWarning = Params["haltonwarning"].UserTypedIn;
            bool haltOnFatalError = Params["haltonfatalerror"].UserTypedIn;
            bool includeusersecurity = Params["includeusersecurity"].UserTypedIn;
            bool logFile = !Params["nologfile"].UserTypedIn;
            bool retainObjectIdentity = Params["retainobjectidentity"].UserTypedIn;
            bool copySecurity = Params["copysecuritysettings"].UserTypedIn;
            bool suppressAfterEvents = Params["suppressafterevents"].UserTypedIn;

            bool retargetLinks = Params["retargetlinks"].UserTypedIn;
            string sourceUrl = Params["sourceurl"].Value;
            string targetUrl = Params["url"].Value;

            Common.Lists.ImportList importList = new Common.Lists.ImportList(sourceUrl, targetUrl, retargetLinks);

            SPUpdateVersions updateVersions = SPUpdateVersions.Append;
            if (Params["updateversions"].UserTypedIn)
                updateVersions = (SPUpdateVersions)Enum.Parse(typeof(SPUpdateVersions), Params["updateversions"].Value);

            importList.PerformImport(compressFile, filename, quiet, haltOnWarning, haltOnFatalError, includeusersecurity, logFile, retainObjectIdentity, copySecurity, suppressAfterEvents, updateVersions);
            return (int)ErrorCodes.NoError;
        }


        #endregion

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(System.Collections.Specialized.StringDictionary keyValues)
        {
            if (Params["retargetlinks"].UserTypedIn)
                Params["sourceurl"].IsRequired = true;

            if (Params["copysecuritysettings"].UserTypedIn)
            {
                Params["sourceurl"].IsRequired = true;
                Params["includeusersecurity"].IsRequired = true;
            }

            base.Validate(keyValues);
        }



    }
}
