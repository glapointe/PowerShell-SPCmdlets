using System;
using System.Collections.Generic;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using Microsoft.SharePoint.Administration;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Farm
{
    [Cmdlet("Backup", "SPSite2", SupportsShouldProcess = false, DefaultParameterSetName = "SPWeb"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Farm")]
    [CmdletDescription("The Backup-SPSite cmdlet performs a backup of the site collection when the Identity parameter is used.\r\n\r\nBy default, the site collection will be set to read-only for the duration of the backup to reduce the potential for user activity during the backup operation which could corrupt the backup. If you have SQL Server Enterprise Edition, we recommend that UseSqlSnapshot parameter be used because this ensures a valid backup while it allows users to continue reading and writing to the site collection during the backup.")]
    [RelatedCmdlets(ExternalCmdlets = new[] { "Get-SPSite", "Backup-SPSite" })]
    [Example(Code = "PS C:\\> Get-SPSite \"http://server_name\" | Backup-SPSite2 -IncludeIis",
        Remarks = "This example backs up the site located at http://server_name along with the IIS settings for the host web application.")]
    [Example(Code = "PS C:\\> Get-SPWebApplication \"http://server_name\" | Backup-SPSite2 -IncludeIis",
        Remarks = "This example backs up the web application located at http://server_name along with the IIS settings for the web application.")]
    public class SPCmdletBackupSites : SPCmdletCustom
    {
        Common.Farm.BackupSites _backup;

        [Parameter(ParameterSetName = "SPSite",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The site to backup.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [ValidateNotNull]
        [Alias("Site")]
        public SPSitePipeBind[] Identity { get; set; }

        [Parameter(ParameterSetName = "SPWebApplication",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The web application to backup.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        [ValidateNotNull]
        public SPWebApplicationPipeBind[] WebApplication { get; set; }
        
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "SPFarm",
            Position = 0,
            HelpMessage = "A valid SPFarm object.")]
        public SPFarmPipeBind Farm { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Include IIS settings in the backup.")]
        public SwitchParameter IncludeIis { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Overwrite existing backups.")]
        public SwitchParameter Overwrite { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies the full path to the backup file (that is, C:\\Backup\\site_name.bak).")]
        [ValidateDirectoryExists]
        public string Path { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies the site collection to remain read and write during the backup.\r\n\r\nIf  the NoSiteLock parameter is not specified, then a site collection that has a site collection lock setting of \"none\" or \"no additions\" will be temporarily set to \"read only\" while the site collection backup is performed. Once the backup has completed, the site collection lock will return to its original state. The backup package will record the original site collection lock state so that it is restored to that state.\r\n\r\nIf users are writing to the site collection while the site collection is being backed up, then the NoSiteLock parameter is not recommended for potential impact to backup integrity")]
        public SwitchParameter NoSiteLock { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies a SQL Database Snapshot will be created when the backup begins, and all site collection data will be retrieved directly from the database snapshot. This snapshot will be deleted automatically when the backup completes.\r\n\r\nThe UseSqlSnapshot parameter is recommended if the database server hosting your content database supports database snapshots such as such as SQL Server Enterprise Edition and SQL Server Developer Edition. This is because it will ensure a valid backup while allowing users to continue reading and writing to the site collection during the backup. It is not necessary to specify the NoSiteLock parameter when specifying the UseSqlSnapshot parameter.")]
        public SwitchParameter UseSqlSnapshot { get; set; }

        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();
            _backup = new Common.Farm.BackupSites(Overwrite.IsPresent, Path, IncludeIis.IsPresent, NoSiteLock.IsPresent, UseSqlSnapshot.IsPresent);
            _backup.PrepareSystem();
        }

        protected override void InternalEndProcessing()
        {
            base.InternalEndProcessing();
        }

        protected override void InternalProcessRecord()
        {
            if (Identity != null)
            {
                List<SPSite> sites = new List<SPSite>();
                foreach (SPSitePipeBind sitePipeBind in Identity)
                {
                    sites.Add(sitePipeBind.Read());
                }
                try
                {
                    _backup.BackupSite(sites, false);
                }
                finally
                {
                    foreach (SPSite site in sites)
                    {
                        site.Dispose();
                    }
                }
            }
            else if (WebApplication != null)
            {
                List<SPWebApplication> webApps = new List<SPWebApplication>();
                foreach (SPWebApplicationPipeBind webAppPipeBind in WebApplication)
                {
                    webApps.Add(webAppPipeBind.Read());
                }
                _backup.BackupWebApplication(webApps, false);
            }
            else
            {
                _backup.BackupFarm(Farm.Read(), false);
            }
        }
    }
}
