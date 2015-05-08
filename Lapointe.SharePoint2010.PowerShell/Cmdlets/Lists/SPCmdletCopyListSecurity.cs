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
    [Cmdlet(VerbsCommon.Copy, "SPListSecurity", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Copy the security settings and permissions from one list to another.")]
    [RelatedCmdlets(typeof(SPCmdletExportListSecurity), typeof(SPCmdletImportListSecurity), typeof(SPCmdletGetList))]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/lists/list1\" | Copy-SPListSecurity -TargetList (Get-SPList \"http://server_name/lists/list2\")",
        Remarks = "This example copies the security settings and permissions from list1 to list2.")]
    public class SPCmdletCopyListSecurity : SPCmdletCustom
    {

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The source list whose security will be copied.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPListPipeBind SourceList { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 1,
            HelpMessage = "The target list to apply the new security to.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPListPipeBind TargetList { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "In addition to the list level security copy the security settings of individual items.")]
        public SwitchParameter IncludeItemSecurity { get; set; }
        
        protected override void InternalProcessRecord()
        {
            SPList sourceList = SourceList.Read();
            SPList targetList = TargetList.Read();

            try
            {
                Common.Lists.CopyListSecurity.CopySecurity(sourceList, targetList, targetList.ParentWeb, IncludeItemSecurity, false);
            }
            finally
            {
                sourceList.ParentWeb.Dispose();
                sourceList.ParentWeb.Site.Dispose();
                targetList.ParentWeb.Dispose();
                targetList.ParentWeb.Site.Dispose();
            }

        }
    }
}
