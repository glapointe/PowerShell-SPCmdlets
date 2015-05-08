using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using Microsoft.SharePoint.Deployment;
using System.IO;
using Microsoft.SharePoint.Administration.Backup;
using System.Collections;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    public abstract class SPCmdletExport : SPCmdletExportImport
    {
        private SPIncludeVersions m_IncludeVersions = SPIncludeVersions.All;


        [Parameter(Mandatory = false,
            HelpMessage = "Sets the maximum file size for the compressed export files. If the total size of the exported package is greater than this size, the exported package will be split into multiple files.")]
        public int CompressionSize { get; set; }

        public abstract SPExportObject ExportObject { get; }

        [Parameter(Mandatory = false)]
        public SwitchParameter IncludeDependencies { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter ExcludeChildren { get; set; }

        [Parameter(Mandatory = false, 
            HelpMessage = "Indicates the type of file and list item version history to be included in the export operation. If the IncludeVersions parameter is absent, the cmdlet by default uses a value of \"LastMajor\".\r\n\r\nThe type must be any one of the following versions:\r\n\r\n-Last major version for files and list items (default)\r\n-The current version, either the last major version or the last minor version\r\n-Last major and last minor version for files and list items\r\n-All versions for files and list items")]
        public SPIncludeVersions IncludeVersions
        {
            get { return this.m_IncludeVersions; }
            set { this.m_IncludeVersions = value; }
        }

        [Parameter(Mandatory = false)]
        public SPIncludeDescendants IncludeDescendants { get; set; }


        [Parameter(Mandatory = false,
            HelpMessage = "Specifies the URL of the Web application, GUID, or object to be exported.\r\n\r\nThe type must be a valid URL; for example, http://server_name.")]
        public string ItemUrl { get; set; }

        public abstract SPSite Site { get; }

        public abstract string SiteUrl { get; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies a SQL Database Snapshot will be created when the export process begins, and all exported data will be retrieved directly from the database snapshot. This snapshot will be automatically deleted when export completes.")]
        public SwitchParameter UseSqlSnapshot { get; set; }

        protected override void InternalProcessRecord()
        {
            bool createLogFile = !base.NoLogFile.IsPresent;
            if (!base.NoFileCompression.IsPresent)
            {
                this.ValidateDirectory(base.Path);
            }
            if ((!base.Force.IsPresent) && File.Exists(base.Path))
            {
                throw new SPException(string.Format("PSCannotOverwriteExport,{0}", base.Path));
            }
            SPExportSettings settings = new SPExportSettings();
            SPExport export = new SPExport(settings);
            base.SetDeploymentSettings(settings);
            settings.ExportMethod = SPExportMethodType.ExportAll;
            settings.ExcludeDependencies = !IncludeDependencies.IsPresent;
            SPDatabaseSnapshot snapshot = null;
            if (!this.UseSqlSnapshot.IsPresent)
            {
                settings.SiteUrl = this.SiteUrl;
            }
            else
            {
                snapshot = this.Site.ContentDatabase.Snapshots.CreateSnapshot();
                SPContentDatabase database2 = SPContentDatabase.CreateUnattachedContentDatabase(snapshot.ConnectionString);
                settings.UnattachedContentDatabase = database2;
                settings.SiteUrl = database2.Sites[this.Site.ServerRelativeUrl].Url;
            }
            SPExportObject exportObject = this.ExportObject;
            if (((exportObject.Type != SPDeploymentObjectType.Web) || 
                base.ShouldProcess(string.Format("ShouldProcessExportWeb,{0},{1}", this.SiteUrl, base.Path ))) && 
                ((exportObject.Type != SPDeploymentObjectType.List) || 
                base.ShouldProcess(string.Format("ShouldProcessExportList,{0},{1}", this.SiteUrl + "/" + this.ItemUrl, base.Path))))
            {
                if (exportObject != null)
                {
                    exportObject.ExcludeChildren = ExcludeChildren.IsPresent;
                    exportObject.IncludeDescendants = IncludeDescendants;
                    settings.ExportObjects.Add(exportObject);
                }
                settings.IncludeVersions = this.IncludeVersions;
                if (base.IncludeUserSecurity.IsPresent)
                {
                    settings.IncludeSecurity = SPIncludeSecurity.All;
                }
                settings.OverwriteExistingDataFile = (bool)base.Force;
                settings.FileMaxSize = this.CompressionSize;
                try
                {
                    export.Run();
                }
                finally
                {
                    if (base.Verbose && createLogFile)
                    {
                        Console.WriteLine();
                        Console.WriteLine(SPResource.GetString("ExportOperationLogFile", new object[0]));
                        Console.WriteLine("\t{0}", settings.LogFilePath);
                    }
                    if ((this.UseSqlSnapshot.IsPresent) && (snapshot != null))
                    {
                        snapshot.Delete();
                    }
                }
                if (base.Verbose)
                {
                    string fileLocation = settings.FileLocation;
                    ArrayList dataFiles = settings.DataFiles;
                    if (dataFiles != null)
                    {
                        if (((fileLocation != null) && (fileLocation.Length > 0)) && (fileLocation[fileLocation.Length - 1] != System.IO.Path.DirectorySeparatorChar))
                        {
                            fileLocation = fileLocation + System.IO.Path.DirectorySeparatorChar;
                        }
                        Console.WriteLine();
                        Console.WriteLine(SPResource.GetString("ExportOperationFilesGenerated", new object[0]));
                        for (int i = 0; i < dataFiles.Count; i++)
                        {
                            Console.WriteLine("\t{0}{1}", fileLocation, dataFiles[i]);
                            Console.WriteLine();
                        }
                        if (base.NoFileCompression.IsPresent)
                        {
                            DirectoryInfo info = new DirectoryInfo(base.Path);
                            Console.WriteLine("\t{0}", info.FullName);
                        }
                        Console.WriteLine();
                    }
                }
            }
        }

        private bool ValidateDirectory(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }
            try
            {
                FileInfo info = new FileInfo(path);
                if (!info.Directory.Exists)
                {
                    throw new DirectoryNotFoundException();
                }
                switch (info.Name[info.Name.Length - 1])
                {
                    case '\\':
                    case '/':
                        throw new SPException(SPResource.GetString("StsadmReqFileName", new object[0]));
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

    }

}
