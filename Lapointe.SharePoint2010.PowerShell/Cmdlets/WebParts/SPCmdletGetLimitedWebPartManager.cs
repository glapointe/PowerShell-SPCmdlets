using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using Microsoft.SharePoint.WebPartPages;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet(VerbsCommon.Get, "SPLimitedWebPartManager"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Retrieves an SPLimitedWebPartManager object used for managing web parts on a page.")]
    [RelatedCmdlets(typeof(Pages.SPCmdletGetPublishingPage), typeof(Lists.SPCmdletGetFile))]
    [Example(Code = "PS C:\\> Start-SPAssignment -Global\r\nPS C:\\> $mgr = Get-SPLimitedWebPartManager \"http://portal/pages/default.aspx\"\r\nPS C:\\> Stop-SPAssignment -Global",
        Remarks = "This exmple retrieves the SPLimitedWebPartManager associated with http://portal/pages/default.aspx.")]
    public class SPCmdletGetLimitedWebPartManager : SPGetCmdletBaseCustom<SPLimitedWebPartManager>
    {

        [Parameter(Mandatory = true, 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPLimitedWebPartManager object.")]
        [Alias(new string[] { "Url", "Page", "Manager" })]
        public SPLimitedWebPartManagerPipeBind Identity { get; set; }

        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();

            DisposeOutputObjects = true;
        }

        protected override void InternalValidate()
        {
            if (this.Identity != null)
            {
                base.DataObject = this.Identity.Read();
                if (base.DataObject == null)
                {
                    base.WriteError(new PSArgumentException("The web part manager could not be found."), ErrorCategory.InvalidArgument, this.Identity);
                    base.SkipProcessCurrentRecord();
                }
            }
        }

        protected override IEnumerable<SPLimitedWebPartManager> RetrieveDataObjects()
        {
            List<SPLimitedWebPartManager> managers = new List<SPLimitedWebPartManager>();
            if (base.DataObject != null)
            {
                managers.Add(base.DataObject);
                AssignmentCollection.Add(base.DataObject.Web);
                AssignmentCollection.Add(base.DataObject.Web.Site);

                return managers;
            }

            return managers;
        }


    }
}
