using System;
using System.Collections;
using System.IO;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Deployment;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.Administration.Backup;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class ExportList : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExportList"/> class.
        /// </summary>
        public ExportList()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the list view URL."));
            parameters.Add(new SPParam("filename", "f", true, null, new SPNonEmptyValidator(), "Please specify the filename."));
            parameters.Add(new SPParam("overwrite", "over", false, null, null));
            parameters.Add(new SPParam("quiet", "quiet", false, null, null));
            parameters.Add(new SPParam("includeusersecurity", "security", false, null, null));
            parameters.Add(new SPParam("haltonwarning", "warning", false, null, null));
            parameters.Add(new SPParam("haltonfatalerror", "error", false, null, null));
            parameters.Add(new SPParam("nologfile", "nolog", false, null, null));
            parameters.Add(new SPParam("versions", "v", false, SPIncludeVersions.All.ToString(), new SPIntRangeValidator(1, 4), "Please specify the version settings."));
            parameters.Add(new SPParam("cabsize", "csize", false, null, new SPIntRangeValidator(1, 0x400), "Please specify the cab size."));
            parameters.Add(new SPParam("nofilecompression", "nofilecompression", false, null, null));
            parameters.Add(new SPParam("includedescendants", "descendants", false, SPIncludeDescendants.All.ToString(), new SPEnumValidator(typeof(SPIncludeDescendants))));
            parameters.Add(new SPParam("excludedependencies", "exdep", false, null, null));
            parameters.Add(new SPParam("usesqlsnapshot", "snapshot"));
            parameters.Add(new SPParam("excludechildren", "excludechildren"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nExports a list.\r\n\r\nParameters:\r\n\t");
            sb.Append("-url <list view url>\r\n\t");
            sb.Append("-filename <export file name>\r\n\t");
            sb.Append("[-overwrite]\r\n\t");
            sb.Append("[-includeusersecurity]\r\n\t");
            sb.Append("[-haltonwarning]\r\n\t");
            sb.Append("[-haltonfatalerror]\r\n\t");
            sb.Append("[-nologfile]\r\n\t");
            sb.Append("[-versions <1-4>\r\n");
            sb.Append("\t\t1 - Last major version for files and list items\r\n");
            sb.Append("\t\t2 - The current version, either the last major or the last minor\r\n");
            sb.Append("\t\t3 - Last major and last minor version for files and list items\r\n");
            sb.Append("\t\t4 - All versions for files and list items (default)]\r\n\t");
            sb.Append("[-cabsize <integer from 1-1024 megabytes> (default: 25)]\r\n\t");
            sb.Append("[-nofilecompression]\r\n\t");
            sb.Append("[-includedescendants <All | Content | None>]\r\n\t");
            sb.Append("[-excludedependencies (Specifies whether to exclude dependencies from the export package when exporting objects of type SPFile or SPListItem)]\r\n\t");
            sb.Append("[-quiet]\r\n\t");
            sb.Append("[-usesqlsnapshot]\r\n\t");
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

            string url = Params["url"].Value;
            bool compressFile = !Params["nofilecompression"].UserTypedIn;
            string filename = Params["filename"].Value;
            bool overwrite = Params["overwrite"].UserTypedIn;
            bool quiet = Params["quiet"].UserTypedIn;
            bool haltOnWarning = Params["haltonwarning"].UserTypedIn;
            bool haltOnFatalError = Params["haltonfatalerror"].UserTypedIn;
            bool includeusersecurity = Params["includeusersecurity"].UserTypedIn;
            bool excludeDependencies = Params["excludedependencies"].UserTypedIn;
            bool logFile = !Params["nologfile"].UserTypedIn;
            bool useSqlSnapshot = Params["usesqlsnapshot"].UserTypedIn;
            bool excludeChildren = Params["excludechildren"].UserTypedIn;

            int cabSize = 0;
            if (Params["cabsize"].UserTypedIn)
            {
                cabSize = int.Parse(Params["cabsize"].Value);
            }

            SPIncludeVersions versions = SPIncludeVersions.All;
            if (Params["versions"].UserTypedIn)
                versions = (SPIncludeVersions)Enum.Parse(typeof(SPIncludeVersions), Params["versions"].Value);
            SPIncludeDescendants includeDescendents = (SPIncludeDescendants)Enum.Parse(typeof (SPIncludeDescendants), Params["includedescendants"].Value, true);

            Common.Lists.ExportList.PerformExport(url, filename, compressFile, haltOnFatalError, haltOnWarning, includeusersecurity, cabSize, logFile, overwrite, quiet, versions, includeDescendents, excludeDependencies, useSqlSnapshot, excludeChildren);


            return (int)ErrorCodes.NoError;
        }
        #endregion


    }
}
