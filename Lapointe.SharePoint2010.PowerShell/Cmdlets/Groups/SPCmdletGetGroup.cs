using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.Quotas;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Groups
{
    [Cmdlet(VerbsCommon.Get, "SPGroup"),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Groups")]
    [CmdletDescription("Retrieve a SharePoint Group from a Site.")]
    [RelatedCmdlets(typeof(SPCmdletNewGroup), typeof(SPCmdletRemoveGroup))]
    [Example(Code = "PS C:\\> $group = Get-SPGroup -Web \"http://demo\" -Identity \"Approvers\"",
        Remarks = "This example retrieves the Approvers group from the http://demo site.")]
    public class SPCmdletGetGroup : SPGetCmdletBaseCustom<SPGroup>
    {

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the group to be retrieved.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            Position = 1,
            HelpMessage = "Specifies the Name or ID of the Group to be retrieved.\r\n\r\nThe type must be a valid integer, or a valid group name (for example, Approvers); or an instance of a valid SPGroup object. If not an SPGroup object then the -Web parameter is required.")]
        [ValidateNotNull]
        public PipeBindObjects.SPGroupPipeBind[] Identity { get; set; }

        protected override IEnumerable<SPGroup> RetrieveDataObjects()
        {
            List<SPGroup> groups = new List<SPGroup>();
            if (Web != null)
            {
                using (SPWeb web = Web.Read())
                {
                    foreach (PipeBindObjects.SPGroupPipeBind gpb in Identity)
                        groups.Add(gpb.Read(web));
                }
            }
            else
            {
                foreach (PipeBindObjects.SPGroupPipeBind gpb in Identity)
                    groups.Add(gpb.Read());
            }
            return groups;
        }
    }
}
