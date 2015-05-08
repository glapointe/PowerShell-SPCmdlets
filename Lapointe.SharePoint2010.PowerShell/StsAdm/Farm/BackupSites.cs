using System;
using System.Collections.Specialized;
using System.DirectoryServices;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.Common;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Farm
{
    public class BackupSites : SPOperation
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="BackupSites"/> class.
        /// </summary>
        public BackupSites()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", false, null, new SPUrlValidator()));
            parameters.Add(new SPParam("scope", "s", false, "site", new SPRegexValidator("(?i:^Farm$|^WebApplication$|^Site$)")));
            parameters.Add(new SPParam("path", "p", true, null, new SPDirectoryExistsValidator()));
            parameters.Add(new SPParam("overwrite", "overwrite"));
            parameters.Add(new SPParam("includeiis", "iis"));
            parameters.Add(new SPParam("nositelock", "nositelock"));
            parameters.Add(new SPParam("usesnapshot", "usesnapshot"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nBackup all sites within the specified scope.  If the scope is farm or webapplication then IIS will also be backed up.\r\n\r\nParameters:");
            sb.Append("\r\n\t-path <path to backup directory (all backups will be placed in a folder beneath this directory)>");
            sb.Append("\r\n\t[-scope <Farm | WebApplication | Site>]");
            sb.Append("\r\n\t[-url <url of web application or site to backup (not required if the scope if farm)>]");
            sb.Append("\r\n\t[-includeiis]");
            sb.Append("\r\n\t[-overwrite]");
            sb.Append("\r\n\t[-nositelock]");
            sb.Append("\r\n\t[-usesnapshot]");

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

            string scope = Params["scope"].Value.ToLowerInvariant();
            bool overwrite = Params["overwrite"].UserTypedIn;
            bool includeIis = Params["includeiis"].UserTypedIn;
            bool noSiteLock = Params["nositelock"].UserTypedIn;
            bool useSnapshot = Params["usesnapshot"].UserTypedIn;
            string path = Path.Combine(Params["path"].Value, DateTime.Now.ToString("yyyyMMdd_"));
            string url = null;
            if (Params["url"].UserTypedIn)
                url = Params["url"].Value.TrimEnd('/');

            Common.Farm.BackupSites backup = new Common.Farm.BackupSites(overwrite, path, includeIis, noSiteLock, useSnapshot);
            if (scope == "farm")
            {
                backup.BackupFarm(true);
            }
            else if (scope == "webapplication")
            {
                backup.BackupWebApplication(url, true);
            }
            else
            {
                // scope == "site"
                backup.BackupSite(url, true);
            }

            return (int)ErrorCodes.NoError;
        }

       /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            if (Params["scope"].Validate())
                Params["url"].IsRequired = (Params["scope"].Value.ToLowerInvariant() != "farm");

            base.Validate(keyValues);
        }

        #endregion
    }
}
