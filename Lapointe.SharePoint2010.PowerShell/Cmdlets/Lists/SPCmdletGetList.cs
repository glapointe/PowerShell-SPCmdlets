using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.Get, "SPList", SupportsShouldProcess = false), 
    SPCmdlet(RequireLocalFarmExist = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Retrieve an SPList object by name or type. Use the AssignmentCollection parameter to ensure parent objects are properly disposed.")]
    [RelatedCmdlets(typeof(SPCmdletDeleteList), typeof(SPCmdletCopyList), typeof(SPCmdletCopyListSecurity),
        typeof(SPCmdletExportListSecurity), ExternalCmdlets = new[] {"Get-SPWeb", "Start-SPAssignment", "Stop-SPAssignment"})]
    [Example(Code = "PS C:\\> $list = Get-SPList \"http://server_name/lists/mylist\"",
        Remarks = "This example retrieves the list at http://server_name/lists/mylist.")]
    public class SPCmdletGetList : SPGetCmdletBaseCustom<SPList>
    {
        #region Parameters

        [Parameter(Mandatory = false, 
            ValueFromPipeline = true, 
            Position = 0, 
            ParameterSetName = "AllListsInIdentity")]
        public SPListPipeBind Identity { get; set; }

        [Parameter(Mandatory = false, 
            ValueFromPipeline = false, 
            ParameterSetName = "AllListsByType")]
        public SPBaseType ListType { get; set; }

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            HelpMessage = "Specifies the URL or GUID of the Web containing the list to be retrieved.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        #endregion

        protected override IEnumerable<SPList> RetrieveDataObjects()
        {
            List<SPList> lists = new List<SPList>();
            SPWeb web = null;
            if (this.Web != null)
                web = this.Web.Read();

            if (Identity == null && ParameterSetName != "AllListsByType")
            {
                foreach (SPList list in web.Lists)
                    lists.Add(list);
            }
            else if (Identity == null && ParameterSetName == "AllListsByType")
            {
                foreach (SPList list in web.GetListsOfType(ListType))
                    lists.Add(list);
            }
            else
            {
                SPList list = this.Identity.Read(web);
                if (list != null)
                    lists.Add(list);
            }

            AssignmentCollection.Add(web);
            foreach (SPList list1 in lists)
            {
                AssignmentCollection.Add(list1.ParentWeb);
                AssignmentCollection.Add(list1.ParentWeb.Site);
            }

            return lists;
        }
    }
}
