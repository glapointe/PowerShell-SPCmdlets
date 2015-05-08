using System.Text;
using System.Collections.Generic;
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
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

namespace Lapointe.SharePoint.PowerShell.Cmdlets.SiteCollections
{
    [Cmdlet("Import", "SPNavigation", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Site Collections")]
    [CmdletDescription("Export the security settings and permissions from a list.")]
    [RelatedCmdlets(typeof(SPCmdletExportNavigation))]
    [Example(Code = "PS C:\\> Get-SPSite \"http://server_name/\" | Import-SPNavigation -InputFile \"c:\\nav.xml\"",
        Remarks = "This example imports the navigation settings from the c:\\nav.xml to the specified site collection.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name/subsite\" | Import-SPNavigation -InputFile \"c:\\nav.xml\"",
        Remarks = "This example imports the navigation settings from the c:\\nav.xml to http://server_name/subsite.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name/subsite\" | Import-SPNavigation -InputFile \"c:\\nav.xml\" -IncludeChildren",
        Remarks = "This example imports the navigation settings from the c:\\nav.xml to http://server_name/subsite and all of its sub-sites.")]
    public sealed class SPCmdletImportNavigation : SPCmdletCustom
    {
        [Parameter(Mandatory = true, ParameterSetName = "SPSite",
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 0,
        HelpMessage = "Specifies the URL or GUID of the Site whose navigation settings will be updated.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "SPWeb",
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 1,
        HelpMessage = "Specifies the URL or GUID of the Web whose navigation settings will be updated.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The path to the file containing the navigation settings to import.")]
        public string InputFile { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "SPWeb",
            HelpMessage = "Include all child webs of the specified Web.")]
        public SwitchParameter IncludeChildren { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Delete the existing glboal navigation settings.")]
        public SwitchParameter DeleteExistingGlobal { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Delete the existing current navigation settings.")]
        public SwitchParameter DeleteExistingCurrent { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The path to the backup file to save the navigation settings to prior to importing.")]
        [ValidateDirectoryExistsAndValidFileName]
        [Alias("Backup")]
        public string BackupFile { get; set; }


        protected override void InternalProcessRecord()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(InputFile);

            switch (ParameterSetName)
            {
                case "SPWeb":
                    using (SPWeb web = Web.Read())
                    {
                        try
                        {
                            WriteVerbose("Importing navigation settings to " + web.Url);
                            Common.SiteCollections.ImportNavigation.SetNavigation(web, xmlDoc, DeleteExistingGlobal, DeleteExistingCurrent, IncludeChildren);
                        }
                        finally
                        {
                            web.Dispose();
                            web.Site.Dispose();
                        }
                    }
                    break;
                case "SPSite":
                    using (SPSite site = Site.Read())
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(BackupFile))
                            {
                                XmlDocument xmlBackupDoc = Common.SiteCollections.ExportNavigation.GetNavigation(site);
                                xmlBackupDoc.Save(BackupFile);
                            }

                            WriteVerbose("Importing navigation settings to  " + site.Url);
                            Common.SiteCollections.ImportNavigation.SetNavigation(site, xmlDoc, DeleteExistingGlobal, DeleteExistingCurrent);
                        }
                        finally
                        {
                            site.Dispose();
                        }
                        
                    }
                    break;
            }
        }
    }

}
