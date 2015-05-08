using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.Remove, "SPList", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Delete a list from a web site.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), ExternalCmdlets = new[] {"Get-SPWeb"})]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/lists/mylist\" | Remove-SPList -BackupDirectory \"c:\\backups\\mylist\"",
        Remarks = "This example deletes the list mylist and creates a backup of the list in the c:\\backups\\mylist folder.")]
    public class SPCmdletDeleteList : SPRemoveCmdletBaseCustom<SPList>
    {
        private SPWeb m_Web;

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The list to delete.\r\n\r\nThe value must be a valid URL in the form http://server_name/lists/listname or /lists/listname. If a server relative URL is provided then the Web parameter must be provided.")]
        [ValidateNotNull]
        public SPListPipeBind Identity { get; set; }

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            HelpMessage = "Specifies the URL or GUID of the Web containing the list to be deleted.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Force the deletion of the list by overriding the AllowDeletion flag.")]
        public SwitchParameter Force { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Export the list prior to deleting. The type must be a valid directory path.")]
        public string BackupDirectory { get; set; }

        protected override void InternalValidate()
        {
            if (this.Web != null)
            {
                this.m_Web = this.Web.Read();
            }
            if (this.Identity != null)
            {
                base.DataObject = (this.m_Web != null) ? this.Identity.Read(this.m_Web) : this.Identity.Read();
            }
            if (base.DataObject == null)
            {
                base.WriteError(new PSArgumentException("A valid SPList object must be provided."), ErrorCategory.InvalidArgument, null);
                base.SkipProcessCurrentRecord();
            }
        }

        protected override void DeleteDataObject()
        {
            if (DataObject != null)
            {
                try
                {
                    Common.Lists.DeleteList.Delete(Force.IsPresent, BackupDirectory, DataObject);
                }
                finally
                {
                    DataObject.ParentWeb.Dispose();
                    DataObject.ParentWeb.Site.Dispose();
                    if (m_Web != null)
                    {
                        m_Web.Dispose();
                        m_Web.Site.Dispose();
                    }
                }
            }
        }

        protected override string ConfirmationMessage
        {
            get
            {
                if (base.DataObject == null)
                {
                    return this.Identity.ToString();
                }
                return base.DataObject.ToString();
            }
        }
 



    }
}
