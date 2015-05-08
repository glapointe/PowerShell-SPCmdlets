using System.Collections.Generic;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Farm
{
    [Cmdlet(VerbsCommon.Get, "SPDeveloperDashboard", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Farm")]
    [CmdletDescription("Retrieves the Developer Dashboard Settings object.")]
    [RelatedCmdlets(typeof(SPCmdletSetDeveloperDashboard))]
    [Example(Code = "PS C:\\> $dash = Get-SPDeveloperDashboard",
        Remarks = "This example returns back the developer dashboard settings object.")]
    public class SPCmdletGetDeveloperDashboard : SPGetCmdletBaseCustom<SPDeveloperDashboardSettings>
    {
        protected override IEnumerable<SPDeveloperDashboardSettings> RetrieveDataObjects()
        {
            WriteObject(SPWebService.ContentService.DeveloperDashboardSettings);

            return null;
        }
    }
}
