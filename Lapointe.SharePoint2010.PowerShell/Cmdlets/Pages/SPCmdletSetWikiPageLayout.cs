using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using Microsoft.SharePoint.Publishing;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Collections;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Enums;
using Lapointe.SharePoint.PowerShell.Common.Pages;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;


namespace Lapointe.SharePoint.PowerShell.Cmdlets.Pages
{
    [Cmdlet("Set", "SPOWikiPageLayout", SupportsShouldProcess = false, DefaultParameterSetName = "List")]
    [CmdletGroup("Wiki Pages")]
    [CmdletDescription("Sets the layout for a wiki page within an existing List.")]
    [RelatedCmdlets(typeof(SPCmdletSetWikiPageLayout), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> Get-SPWeb http://server_name | Set-SPWikiPageLayout -List \"Site Pages\" -WikiPageName \"MyWikiPage.aspx\" -WikiPageLayout \"OneColumnSideBar\"",
        Remarks = "This example sets the layout of an existing wiki page within the Site Pages list under the root Site of the root Site Collection.")]
    public class SPCmdletSetWikiPageLayout : SPSetCmdletBaseCustom<SPFile>
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to create the wiki page in.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }


        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "List",
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the identity of the Site containing the file to update.")]
        public SPListPipeBind List { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Folder",
            Position = 2,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the Folder containing the file to update.")]
        public SPFolderPipeBind Folder { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 3,
            Mandatory = true,
            HelpMessage = "The name of the wiki page to update.")]
        public string WikiPageName { get; set; }

        [Parameter(
            Position = 4,
            Mandatory = true,
            HelpMessage = "The page layout to set the wiki page to.")]
        public WikiPageLayout WikiPageLayout { get; set; }


        protected override void UpdateDataObject()
        {
            SPFile page = null;
            if (ParameterSetName == "List")
            {
                var web = Web.Read();
                SPList list = List.Read(web);
                page = list.RootFolder.Files[WikiPageName];
            }
            else
            {
                page = Folder.Read().Files[WikiPageName];
            }
            if (page != null)
            {
                WikiPageUtilities.SetWikiPageLayout(page.Item, WikiPageLayout);
            }
        }
    }
}
