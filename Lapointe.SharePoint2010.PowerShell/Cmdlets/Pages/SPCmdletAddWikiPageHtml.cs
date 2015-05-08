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
    [Cmdlet("Add", "SPWikiPageHtml", SupportsShouldProcess = false, DefaultParameterSetName = "List")]
    [CmdletGroup("Wiki Pages")]
    [CmdletDescription("Sets the layout for a wiki page within an existing List.")]
    [RelatedCmdlets(typeof(SPCmdletSetWikiPageLayout), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> Get-SPWeb http://server_name | Add-SPWikiPageHtml -List \"Site Pages\" -WikiPageName \"MyWikiPage.aspx\" -Html \"<h1>Welcome!</h1>\" -Prepend -Row 1 -Column 1",
        Remarks = "This example adds a welcome header to the wiki page.")]
    public class SPCmdletAddWikiPageHtml : SPSetCmdletBaseCustom<SPFile>
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
            HelpMessage = "The HTML to add to the wiki page.")]
        public string Html { get; set; }

        [Parameter(
            Position = 5,
            Mandatory = true,
            HelpMessage = "The zone row to add the HTML to.")]
        public int Row { get; set; }

        [Parameter(
            Position = 6,
            Mandatory = true,
            HelpMessage = "The zone column to add the HTML to.")]
        public int Column { get; set; }

        [Parameter(
            Position = 7,
            Mandatory = false,
            HelpMessage = "Add some space before the web part.")]
        public SwitchParameter Prepend { get; set; }

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
                WikiPageUtilities.AddHtmlToWikiPage(page.Item, Html, Row, Column, Prepend);
            }
        }
    }
}
