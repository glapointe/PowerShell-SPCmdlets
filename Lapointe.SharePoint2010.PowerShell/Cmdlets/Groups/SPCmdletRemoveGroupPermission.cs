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
    [Cmdlet(VerbsCommon.Remove, "SPGroupPermission", SupportsShouldProcess = true),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Groups")]
    [CmdletDescription("Removes a permission from a SharePoint Group.")]
    [RelatedCmdlets(typeof(SPCmdletNewGroup), typeof(SPCmdletGetGroup))]
    [Example(Code = "PS C:\\> Remove-SPGroupPermission -Web \"http://demo\" -Group \"My Group\" -Permission \"Approve\",\"Contribute\"",
        Remarks = "This example removes the Approve and Contribute permissions from the My Group group from the http://demo site.")]
    public class SPCmdletRemoveGroupPermission : SPRemoveCmdletBaseCustom<SPGroup>
    {

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the group to be created.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            Position = 1,
            HelpMessage = "Specifies the Name or ID of the Group to add the permission to.\r\n\r\nThe type must be a valid integer, or a valid group name (for example, Approvers); or an instance of a valid SPGroup object. If not an SPGroup object then the -Web parameter is required.")]
        [ValidateNotNull]
        public PipeBindObjects.SPGroupPipeBind Identity { get; set; }

        [Parameter(Mandatory = true,
            Position = 2,
            HelpMessage = "The permissions to assign to the group")]
        [ValidateNotNullOrEmpty]
        public string[] Permission { get; set; }

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

            SPGroup group = DataObject;
            if (group != null)
            {
                SPRoleAssignment ra = group.ParentWeb.RoleAssignments.GetAssignmentByPrincipal(group);
                foreach (string permission in Permission) 
                {
                    SPRoleDefinition rd = null;
                    try
                    {
                        rd = group.ParentWeb.RoleDefinitions[permission];
                    }
                    catch (SPException)
                    {
                        throw new SPException(string.Format("Permission level \"{0}\" cannot be found.", permission));
                    }
                    if (ra.RoleDefinitionBindings.Contains(rd))
                        ra.RoleDefinitionBindings.Remove(rd);
                }
                if (!test)
                {
                    ra.Update();
                    group.Update();
                }
            }
        }
    }
}
