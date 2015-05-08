using System.Collections.Generic;
using Lapointe.SharePoint.PowerShell.Common.Lists;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.Get, "SPListOverlays", SupportsShouldProcess = false), 
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false, RequireUserMachineAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Retrieve all SPList objects set as a calendar overlay on the given list.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), typeof(SPCmdletSetListOverlay), ExternalCmdlets = new[] {"Get-SPWeb"})]
    [Example(Code = "PS C:\\> $lists = Get-SPListOverlays \"http://server_name/lists/mylist\"",
        Remarks = "This example retrieves the calendar overlays for the calendar at http://server_name/lists/mycalendar.")]
    public class SPCmdletGetListOverlays : SPGetCmdletBaseCustom<SPList>
    {
        #region Parameters

        [Parameter(Mandatory = true, 
            ValueFromPipeline = true,
            HelpMessage = "The calendar whose calendar overlays will be retrieved.\r\n\r\nThe value must be a valid URL in the form http://server_name/lists/listname or /lists/listname. If a server relative URL is provided then the Web parameter must be provided.",
            Position = 0)]
        [Alias("List")]
        public SPListPipeBind Identity { get; set; }

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            HelpMessage = "Specifies the URL or GUID of the Web containing the calendar whose overlays will be retrieved.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        #endregion

        protected override IEnumerable<SPList> RetrieveDataObjects()
        {
            SPWeb web = null;
            if (this.Web != null)
                web = this.Web.Read();

            return GetListOverlays.GetOverlayLists(Identity.Read(web));
        }
    }
}
