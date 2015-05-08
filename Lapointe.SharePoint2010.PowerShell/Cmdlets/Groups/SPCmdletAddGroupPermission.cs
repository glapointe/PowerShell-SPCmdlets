using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Groups
{
    [Cmdlet(VerbsCommon.Add, "SPGroupPermission", SupportsShouldProcess=true),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Groups")]
    [CmdletDescription("Adds a permission to a SharePoint Group.")]
    [RelatedCmdlets(typeof(SPCmdletGetGroup), typeof(SPCmdletNewGroup))]
    [Example(Code = "PS C:\\> Get-SPGroup -Web \"http://demo\" -Name \"My Group\" | Add-SPGroupPermission -Permission \"Approve\",\"Contribute\"",
        Remarks = "This example adds the Approve and Contribute permissions to the \"My Group\" group in the http://demo site.")]
    public class SPCmdletAddGroupPermission : SPCmdletCustom
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

        protected override void InternalProcessRecord()
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

            SPGroup group = null;
            if (Web != null)
            {
                using (SPWeb web = Web.Read())
                {
                    group = Identity.Read(web);
                }
            }
            else
            {
                group = Identity.Read();
            }

            if (group == null)
            {
                WriteError(new PSArgumentException("The specified group could not be found."), ErrorCategory.InvalidArgument, null);
                SkipProcessCurrentRecord();
                return;
            }

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
                if (!ra.RoleDefinitionBindings.Contains(rd))
                    ra.RoleDefinitionBindings.Add(rd);
            }
            if (!test)
            {
                ra.Update();
                group.Update();
            }
        }
    }
}
