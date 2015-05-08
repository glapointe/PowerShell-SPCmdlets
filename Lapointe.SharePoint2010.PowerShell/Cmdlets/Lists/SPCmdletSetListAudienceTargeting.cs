using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.Win32;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation.Internal;
using Lapointe.SharePoint.PowerShell.Common;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.Set, "SPListAudienceTargeting", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true,RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Sets audience targeting for the specified list.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/lists/mylist\" | Set-SPListAudienceTargeting -Enabled $true",
        Remarks = "This example enables audience targeting on the mylist list.")]
    public class SPCmdletSetListAudienceTargeting : SPSetCmdletBaseCustom<SPList>
    {
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The list to update.")]
        public SPListPipeBind Identity { get; set; }

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The web containing the list. This parameter is required if the Identity parameter is a relative URL to a list.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "Enable or disable audience targeting. Valid values are $true or $false.")]
        public bool Enabled { get; set; }


        protected override void InternalValidate()
        {
            if (Identity != null)
            {
                if (Web != null)
                    DataObject = Identity.Read(Web.Read());
                else
                    DataObject = Identity.Read();
            }
        }

        protected override void UpdateDataObject()
        {
            Common.Lists.ListAudienceTargeting.SetTargeting(Enabled, DataObject);
        }
    }
}
