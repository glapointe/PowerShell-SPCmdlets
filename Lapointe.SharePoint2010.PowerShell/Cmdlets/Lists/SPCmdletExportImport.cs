using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using Microsoft.SharePoint.Deployment;
using System.IO;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    public abstract class SPCmdletExportImport : SPCmdletCustom
    {

        [Parameter(Mandatory = false,
            HelpMessage = "Forcefully overwrites the export package if it already exists.")]
        public SwitchParameter Force
        {
            get { return base.GetSwitch("Force"); }
            set { base.SetProp("Force", value); }
        }

        [Parameter(Mandatory = false,
            HelpMessage = "Stops the export/import process when an error occurs.")]
        public SwitchParameter HaltOnError { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Stops the export/import process when a warning occurs.")]
        public SwitchParameter HaltOnWarning { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Preserves the user security settings.")]
        public SwitchParameter IncludeUserSecurity { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Either enables or disables file compression in the export package. The export package is stored in the folder specified by thePath parameter or Identity parameter. We recommend that you use this parameter for performance reasons. If compression is enabled, the export process can increase by approximately 30 percent.")]
        public SwitchParameter NoFileCompression { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Suppresses the generation of an export log file. If this parameter is not specified, the cmdlet will generate an export log file in the same location as the export package. The log file uses Unified Logging Service (ULS).\r\n\r\nIt is recommended to use this parameter. However, for performance reasons, you might not want to generate a log file.")]
        public SwitchParameter NoLogFile { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "Specifies the name of the export file.\r\n\r\nIf the NoFileCompression parameter is used, a directory must be specified; otherwise, any file format is valid.")]
        public virtual string Path { get; set; }

        internal bool Verbose
        {
            get
            {
                return (base.MyInvocation.Line.IndexOf("-Verbose", StringComparison.InvariantCultureIgnoreCase) >= 0);
            }
        }

        protected void SetDeploymentSettings(SPDeploymentSettings settings)
        {
            string filePath;
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }
            settings.CommandLineVerbose = this.Verbose;
            settings.HaltOnNonfatalError = (bool)this.HaltOnError;
            settings.HaltOnWarning = (bool)this.HaltOnWarning;
            //settings.WarnOnUsingLastMajor = true;
            settings.FileCompression = !NoFileCompression.IsPresent;
            string filename = string.Empty;
            if (!NoFileCompression.IsPresent)
            {
                SplitPathFile(this.Path, out filePath, out filename);
                settings.FileLocation = filePath;
                settings.BaseFileName = filename;
            }
            else
            {
                settings.FileLocation = filePath = this.Path;
            }
            if (!NoLogFile.IsPresent)
            {
                string logFileName;
                if (this is SPCmdletExport)
                {
                    logFileName = "export.log";
                }
                else
                {
                    logFileName = "import.log";
                }
                if (NoFileCompression.IsPresent)
                {
                    settings.LogFilePath = System.IO.Path.Combine(settings.FileLocation, filename + logFileName);
                }
                else
                {
                    settings.LogFilePath = System.IO.Path.Combine(settings.FileLocation, filename + "." + logFileName);
                }
                bool fileExists = File.Exists(settings.LogFilePath) && !(this is SPCmdletImportWeb2);
                if ((!this.Force.IsPresent) && fileExists)
                {
                    throw new SPException(SPResource.GetString("DataFileExists", new object[] { settings.LogFilePath }));
                }
                if (fileExists)
                {
                    File.Delete(settings.LogFilePath);
                }
            }
        }

        internal static void SplitPathFile(string fullPathFile, out string path, out string filename)
        {
            FileInfo info = new FileInfo(fullPathFile);
            path = info.Directory.FullName;
            filename = info.Name;
        }

    }


}
