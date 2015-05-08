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
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.Copy, "SPList", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Copy a list from one site to another. Wraps the functionality of the Export-SPWeb and Import-SPWeb cmdlets.")]
    [RelatedCmdlets(typeof(SPCmdletExportWeb2), typeof(SPCmdletImportWeb2), typeof(SPCmdletGetList), typeof(SPCmdletCopyListSecurity),
        ExternalCmdlets = new[] {"Export-SPWeb", "Import-SPWeb"})]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/sites/site1/lists/mylist\" | Copy-SPList -TargetWeb \"http://server_name/sites/sites2\"",
        Remarks = "This example copies the mylist list from site1 to site2.")]
    public class SPCmdletCopyList : SPCmdletExportImport
    {

        [Parameter(Mandatory = true, 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The source list to copy.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPListPipeBind SourceList { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the URL or GUID of the Web containing the list to be copied.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind TargetWeb { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Delete the source list after the copy operation completes.")]
        public SwitchParameter DeleteSource { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies the name of the export file.\r\n\r\nIf the NoFileCompression parameter is used, a directory must be specified; otherwise, any file format is valid.")]
        public override string Path { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter SuppressAfterEvents { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Iterate through all links pointing to the source and retarget to the new location.")]
        public SwitchParameter RetargetLinks { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "If specified, create the new list with the same ID as the source list. The target location must be in a different content database from the source list for this to work.")]
        public SwitchParameter RetainObjectIdentity { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Indicates the type of file and list item version history to be included in the export operation. If the IncludeVersions parameter is absent, the Copy-SPList cmdlet by default uses a value of \"LastMajor\".\r\n\r\nThe type must be any one of the following versions:\r\n\r\n-Last major version for files and list items (default)\r\n-The current version, either the last major version or the last minor version\r\n-Last major and last minor version for files and list items\r\n-All versions for files and list items")]
        public SPIncludeVersions? IncludeVersions { get; set; }

        [Parameter(Mandatory = false)]
        public SPIncludeDescendants? IncludeDescendants { get; set; }

        [Parameter(Mandatory = false)]
        public SPUpdateVersions? UpdateVersions { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Sets the maximum file size for the compressed export files. If the total size of the exported package is greater than this size, the exported package will be split into multiple files.")]
        public int CompressionSize { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter IncludeDependencies { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter ExcludeChildren { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies a SQL Database Snapshot will be created when the export process begins, and all exported data will be retrieved directly from the database snapshot. This snapshot will be automatically deleted when export completes.")]
        public SwitchParameter UseSqlSnapshot { get; set; }

        protected override void InternalProcessRecord()
        {
            using (SPWeb targetWeb = TargetWeb.Read())
            {
                SPList sourceList = SourceList.Read();
                try
                {
                    bool compressFile = !NoFileCompression.IsPresent;
                    bool quiet = !Verbose;
                    bool haltOnWarning = HaltOnWarning.IsPresent;
                    bool haltOnFatalError = HaltOnError.IsPresent;
                    bool includeusersecurity = IncludeUserSecurity.IsPresent;
                    bool excludeDependencies = !IncludeDependencies.IsPresent;
                    bool copySecurity = includeusersecurity;
                    bool logFile = !NoLogFile.IsPresent;
                    bool deleteSource = DeleteSource.IsPresent;
                    string directory = null;
                    if (!string.IsNullOrEmpty(Path))
                        directory = Path;
                    bool suppressAfterEvents = SuppressAfterEvents.IsPresent;
                    bool retargetLinks = RetargetLinks.IsPresent;

                    SPIncludeVersions versions = SPIncludeVersions.All;
                    if (IncludeVersions.HasValue)
                        versions = IncludeVersions.Value;

                    SPUpdateVersions updateVersions = SPUpdateVersions.Append;
                    if (UpdateVersions.HasValue)
                        updateVersions = UpdateVersions.Value;

                    SPIncludeDescendants includeDescendents = SPIncludeDescendants.All;
                    if (IncludeDescendants.HasValue)
                        includeDescendents = IncludeDescendants.Value;

                    bool useSqlSnapshot = UseSqlSnapshot.IsPresent;
                    bool excludeChildren = ExcludeChildren.IsPresent;
                    Common.Lists.ImportList importList = new Common.Lists.ImportList(sourceList, targetWeb, retargetLinks);

                    importList.Copy(directory, compressFile, CompressionSize, includeusersecurity, excludeDependencies, haltOnFatalError, haltOnWarning, versions, updateVersions, suppressAfterEvents, copySecurity, deleteSource, logFile, quiet, includeDescendents, useSqlSnapshot, excludeChildren, RetainObjectIdentity);
                }
                finally
                {
                    targetWeb.Site.Dispose();
                    sourceList.ParentWeb.Dispose();
                    sourceList.ParentWeb.Site.Dispose();
                }
            }
        }

    }

}
