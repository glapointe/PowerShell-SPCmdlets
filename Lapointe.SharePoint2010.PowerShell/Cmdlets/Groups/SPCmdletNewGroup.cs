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
    [Cmdlet(VerbsCommon.New, "SPGroup"),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Groups")]
    [CmdletDescription("Creates a SharePoint Group in a Site.")]
    [RelatedCmdlets(typeof(SPCmdletGetGroup), typeof(SPCmdletRemoveGroup))]
    [Example(Code = "PS C:\\> $group = New-SPGroup -Web \"http://demo\" -Name \"My Group\" -Owner \"domain\\user\" -Member \"domain\\user\"",
        Remarks = "This example creates the \"My Group\" group in the http://demo site.")]
    public class SPCmdletNewGroup : SPNewCmdletBaseCustom<SPGroup>
    {

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the group to be created.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            Position = 1,
            HelpMessage = "Specifies the name of the group to create.")]
        [ValidateNotNullOrEmpty]
        public string Name { get; set; }

        [Parameter(Mandatory = true,
            Position = 2,
            HelpMessage = "Specifies the owner of the new group.")]
        [ValidateNotNull]
        public SPUserPipeBind Owner { get; set; }

        [Parameter(Mandatory = false,
            Position = 3,
            HelpMessage = "Specifies the default user to add to the new group. If not specified then the Owner will be used.")]
        [ValidateNotNull]
        public SPUserPipeBind DefaultUser { get; set; }

        [Parameter(Mandatory = false,
            Position = 4,
            HelpMessage = "Specifies the description of the group to create.")]
        public string Description { get; set; }

        protected override SPGroup CreateDataObject()
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


            using (SPWeb web = Web.Read())
            {
                SPGroup group = null;
                try
                {
                    group = web.SiteGroups[Name];
                }
                catch {}
                if (group != null)
                {
                    throw new SPCmdletException("Group " + Name + " already exists!");
                }

                SPUser owner = Owner.Read(web);
                SPUser defaultUser = owner;
                if (DefaultUser != null)
                    defaultUser = DefaultUser.Read(web);

                if (!test)
                {
                    web.SiteGroups.Add(Name, owner, defaultUser, Description);
                    group = web.SiteGroups[Name];
                    web.RoleAssignments.Add(group);
                    return group;
                }
            }
            return null;
        }
    }
}
