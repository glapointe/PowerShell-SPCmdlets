using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Groups
{
    [Cmdlet(VerbsCommon.Remove, "SPGroup", SupportsShouldProcess = true),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Groups")]
    [CmdletDescription("Removes a SharePoint Group from a Site.")]
    [RelatedCmdlets(typeof(SPCmdletNewGroup), typeof(SPCmdletGetGroup))]
    [Example(Code = "PS C:\\> Remove-SPGroup -Web \"http://demo\" -Identity \"Approvers\"",
        Remarks = "This example removes the Approvers group from the http://demo site.")]
    public class SPCmdletRemoveGroup : SPRemoveCmdletBaseCustom<SPGroup>
    {

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the group to be removed.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            Position = 1,
            HelpMessage = "Specifies the Name or ID of the Group to be removed.\r\n\r\nThe type must be a valid integer, or a valid group name (for example, Approvers); or an instance of a valid SPGroup object. If not an SPGroup object then the -Web parameter is required.")]
        [ValidateNotNull]
        public PipeBindObjects.SPGroupPipeBind Identity { get; set; }

        protected override void InternalValidate()
        {
            if (Web != null)
            {
                using (SPWeb web = Web.Read())
                {
                    DataObject = Identity.Read(web);
                }
            }
            else
            {
                DataObject = Identity.Read();
            }


            if (DataObject == null)
            {
                WriteError(new PSArgumentException("The specified group could not be found."), ErrorCategory.InvalidArgument, null);
                SkipProcessCurrentRecord();
            }
        }

        protected override void DeleteDataObject()
        {
            bool test = false;
            ShouldProcessReason reason;
            if (!base.ShouldProcess(null, null, null, out reason))
            {
                if (reason == ShouldProcessReason.WhatIf)
                {
                    test = true;
                }
            }
            if (test)
                Logger.Verbose = true;

            if (DataObject != null && !test)
            {
                DataObject.ParentWeb.SiteGroups.RemoveByID(DataObject.ID);
            }
        }
    }
}
