using System;
using System.IO;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Deployment;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class CopyList : ImportList
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CopyList"/> class.
        /// </summary>
        public CopyList()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("sourceurl", "sourceurl", true, null, new SPUrlValidator(), "Please specify the list view URL to copy from."));
            parameters.Add(new SPParam("targeturl", "targeturl", true, null, new SPUrlValidator(), "Please specify the site to copy to."));
            parameters.Add(new SPParam("quiet", "quiet", false, null, null));
            parameters.Add(new SPParam("includeusersecurity", "security", false, null, null));
            parameters.Add(new SPParam("haltonwarning", "warning", false, null, null));
            parameters.Add(new SPParam("haltonfatalerror", "error", false, null, null));
            parameters.Add(new SPParam("nologfile", "nolog", false, null, null));
            parameters.Add(new SPParam("versions", "v", false, SPIncludeVersions.All.ToString(), new SPIntRangeValidator(1, 4), "Please specify the version settings."));
            parameters.Add(new SPParam("updateversions", "updatev", false, SPUpdateVersions.Append.ToString(), new SPIntRangeValidator(1, 3), "Please specify the updateversions setting."));
            parameters.Add(new SPParam("retargetlinks", "retargetlinks", false, null, null));
            parameters.Add(new SPParam("deletesource", "delete", false, null, null));
            parameters.Add(new SPParam("copysecuritysettings", "copysecurity")); // Not used by this class but must be added for base class validation
            parameters.Add(new SPParam("temppath", "temppath", false, null, new SPDirectoryExistsValidator()));
            parameters.Add(new SPParam("includedescendants", "descendants", false, SPIncludeDescendants.All.ToString(), new SPEnumValidator(typeof(SPIncludeDescendants))));
            parameters.Add(new SPParam("excludedependencies", "exdep", false, null, null));
            parameters.Add(new SPParam("nofilecompression", "nofilecompression", false, null, null));
            parameters.Add(new SPParam("suppressafterevents", "sae"));
            parameters.Add(new SPParam("excludechildren", "excludechildren"));
            parameters.Add(new SPParam("cabsize", "csize", false, null, new SPIntRangeValidator(1, 0x400), "Please specify the cab size."));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nCopies a list to a new web.\r\n\r\nParameters:\r\n\t");
            sb.Append("-sourceurl <list view url to copy>\r\n\t");
            sb.Append("-targeturl <url of a web site to copy to>\r\n\t");
            sb.Append("[-includeusersecurity]\r\n\t");
            sb.Append("[-haltonwarning]\r\n\t");
            sb.Append("[-haltonfatalerror]\r\n\t");
            sb.Append("[-nologfile]\r\n\t");
            sb.Append("[-versions <1-4>\r\n");
            sb.Append("\t\t1 - Last major version for files and list items\r\n");
            sb.Append("\t\t2 - The current version, either the last major or the last minor\r\n");
            sb.Append("\t\t3 - Last major and last minor version for files and list items\r\n");
            sb.Append("\t\t4 - All versions for files and list items (default)]\r\n\t");
            sb.Append("[-updateversions <1-3>\r\n");
            sb.Append("\t\t1 - Add new versions to the current file (default)\r\n");
            sb.Append("\t\t2 - Overwrite the file and all its versions (delete then insert)\r\n");
            sb.Append("\t\t3 - Ignore the file if it exists on the destination]\r\n\t");
            sb.Append("[-quiet]\r\n\t");
            sb.Append("[-retargetlinks (resets links pointing to the source to now point to the target)]\r\n\t");
            sb.Append("[-deletesource]\r\n\t");
            sb.Append("[-temppath <temporary folder path for storing of export files>]\r\n\t");
            sb.Append("[-includedescendants <All | Content | None>]\r\n\t");
            sb.Append("[-excludedependencies (Specifies whether to exclude dependencies from the export package when exporting objects of type SPFile or SPListItem)]\r\n\t");
            sb.Append("[-nofilecompression]\r\n\t");
            sb.Append("[-cabsize <integer from 1-1024 megabytes> (default: 25)]\r\n\t");
            sb.Append("[-suppressafterevents (disable the firing of \"After\" events when creating or modifying list items)]\r\n\t");
            sb.Append("[-excludechildren]\r\n\t");
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
            bool compressFile = !Params["nofilecompression"].UserTypedIn;
            bool quiet = Params["quiet"].UserTypedIn;
            bool haltOnWarning = Params["haltonwarning"].UserTypedIn;
            bool haltOnFatalError = Params["haltonfatalerror"].UserTypedIn;
            bool includeusersecurity = Params["includeusersecurity"].UserTypedIn;
            bool excludeDependencies = Params["excludedependencies"].UserTypedIn;
            bool copySecurity = includeusersecurity;
            bool logFile = !Params["nologfile"].UserTypedIn;
            bool deleteSource = Params["deletesource"].UserTypedIn;
            string directory = null;
            if (Params["temppath"].UserTypedIn)
                directory = Params["temppath"].Value;
            bool suppressAfterEvents = Params["suppressafterevents"].UserTypedIn;
            bool retargetLinks = Params["retargetlinks"].UserTypedIn;
            SPIncludeVersions versions = SPIncludeVersions.All;
            if (Params["versions"].UserTypedIn)
                versions = (SPIncludeVersions)Enum.Parse(typeof(SPIncludeVersions), Params["versions"].Value);
            SPUpdateVersions updateVersions = SPUpdateVersions.Append;
            if (Params["updateversions"].UserTypedIn)
                updateVersions = (SPUpdateVersions)Enum.Parse(typeof(SPUpdateVersions), Params["updateversions"].Value);
            SPIncludeDescendants includeDescendents = (SPIncludeDescendants)Enum.Parse(typeof(SPIncludeDescendants), Params["includedescendants"].Value, true);
            bool useSqlSnapshot = Params["usesqlsnapshot"].UserTypedIn;
            bool excludeChildren = Params["excludechildren"].UserTypedIn;
            int cabSize = 0;
            if (Params["cabsize"].UserTypedIn)
            {
                cabSize = int.Parse(Params["cabsize"].Value);
            }

            Common.Lists.ImportList importList = new Common.Lists.ImportList(sourceUrl, targetUrl, retargetLinks);

            importList.Copy(directory, compressFile, cabSize, includeusersecurity, excludeDependencies, haltOnFatalError, haltOnWarning, versions, updateVersions, suppressAfterEvents, copySecurity, deleteSource, logFile, quiet, includeDescendents, useSqlSnapshot, excludeChildren, false);

            return (int)ErrorCodes.NoError;
        }

     

        #endregion

    }
}
