using System;
using System.Collections.Specialized;
using System.DirectoryServices;
using System.IO;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Farm
{
    public class Backup2 : SPOperation
    {
        private const int FLAG_EXPORT_INHERITED_SETTINGS = 1;

        /// <summary>
        /// Initializes a new instance of the <see cref="BackupSites"/> class.
        /// </summary>
        public Backup2()
        {
            SPParamCollection parameters = new SPParamCollection();

            parameters.Add(new SPParam("directory", "dir", false, null, null));
            parameters.Add(new SPParam("backupmethod", "method", false, "None", null));
            parameters.Add(new SPParam("item", "item", false, null, null));
            parameters.Add(new SPParam("quiet", "quiet", false, null, null));
            parameters.Add(new SPParam("percentage", "update", false, null, new SPIntRangeValidator(1, 100)));
            parameters.Add(new SPParam("backupthreads", "threads", false, null, new SPIntRangeValidator(1, 10)));
            parameters.Add(new SPParam("showtree", "tree", false, "False", null));
            parameters.Add(new SPParam("url", "url", false, null, null));
            parameters.Add(new SPParam("filename", "f", false, null, new SPValidator()));
            parameters.Add(new SPParam("overwrite", "overwrite"));
            parameters.Add(new SPParam("includeiis", "iis"));
            parameters.Add(new SPParam("configurationonly", "configuration"));
            parameters.Add(new SPParam("usesqlsnapshot", "snapshot"));
            parameters.Add(new SPParam("nositelock", "nosl"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\nFor site collection backup:");
            sb.Append("\r\n    stsadm.exe -o gl-backup ");
            sb.Append("\r\n        -url <url>");
            sb.Append("\r\n        -filename <filename>");
            sb.Append("\r\n        [-overwrite]");
            sb.Append("\r\n        [-nositelock]");
            sb.Append("\r\n        [-usesqlsnapshot]");
            sb.Append("\r\n\r\nFor catastrophic backup:");
            sb.Append("\r\n    stsadm.exe -o gl-backup");
            sb.Append("\r\n        -directory <UNC path>");
            sb.Append("\r\n        -backupmethod <full | differential>");
            sb.Append("\r\n        [-item <created path from tree>]");
            sb.Append("\r\n        [-percentage <integer between 1 and 100>]");
            sb.Append("\r\n        [-backupthreads <integer between 1 and 10>]");
            sb.Append("\r\n        [-showtree]");
            sb.Append("\r\n        [-configurationonly]");
            sb.Append("\r\n        [-quiet]");
            sb.Append("\r\n        [-includeiis]");

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

            string args = "-o backup";
            foreach (string key in keyValues.Keys)
            {
                if (key.ToLowerInvariant() == "includeiis" || key.ToLowerInvariant() == "o")
                    continue;

                args += string.Format(" -{0} \"{1}\"", key, keyValues[key]);
            }

            if (Utilities.RunStsAdmOperation(args, false) != 0)
                return (int)ErrorCodes.GeneralError;


            if (Params["includeiis"].UserTypedIn && Params["directory"].UserTypedIn)
            {
                bool verbose = !Params["quiet"].UserTypedIn;
                XmlDocument xmlToc = new XmlDocument();
                xmlToc.Load(Path.Combine(Params["directory"].Value, "spbrtoc.xml"));
                string iisBakPath = Path.Combine(xmlToc.DocumentElement.FirstChild.SelectSingleNode("SPBackupDirectory").InnerText, "iis.bak");
                using (DirectoryEntry de = new DirectoryEntry("IIS://localhost"))
                {
                    if (verbose)
                        Console.WriteLine("Flushing IIS metadata to disk....");
                    de.Invoke("SaveData", new object[0]);
                    if (verbose)
                        Console.WriteLine("IIS metadata successfully flushed to disk.");
                }
                using (DirectoryEntry de = new DirectoryEntry("IIS://localhost"))
                {
                    if (verbose)
                        Console.WriteLine("Exporting full IIS settings to {0}....", iisBakPath);
                    string decryptionPwd = string.Empty;
                    de.Invoke("Export", new object[] { decryptionPwd, iisBakPath, "/lm", FLAG_EXPORT_INHERITED_SETTINGS });
                }
            }

            return (int)ErrorCodes.NoError;
        }

        #endregion

    }
}
