using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using Microsoft.SharePoint.Publishing;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Collections;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Enums;
using Lapointe.SharePoint.PowerShell.Common.Pages;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Pages
{
    [Cmdlet(VerbsCommon.New, "SPWikiPage", SupportsShouldProcess = false, DefaultParameterSetName = "List"), 
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Pages")]
    [CmdletDescription("Creates a new wiki page.")]
    [RelatedCmdlets(typeof(SPCmdletSetWikiPageLayout), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> Get-SPWeb http://server_name | New-SPWikiPage -List \"Site Pages\" -File \"MyWikiPage.aspx\"",
        Remarks = "This example creates a new wiki page within the Site Pages list under the root Site of the root Site Collection.")]
    public class SPCmdletNewWikiPage : SPNewCmdletBaseCustom<SPFile>
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
            HelpMessage = "Specifies the List to add the file to.")]
        public SPListPipeBind List { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Folder",
            Position = 2,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the Folder to add the file to.")]
        public SPFolderPipeBind Folder { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 3,
            Mandatory = true,
            HelpMessage = "The name of the wiki page to create.")]
        public string WikiPageName { get; set; }

        [Parameter(Position = 4, Mandatory = false, HelpMessage = "The page layout to set the wiki page to. The default value is \"ThreeColumnsHeaderFooter\".")]
        public WikiPageLayout WikiPageLayout { get; set; }

        protected override SPFile CreateDataObject()
        {
            SPFile page = null;
            try
            {
                if (ParameterSetName == "List")
                {
                    var web = Web.Read();
                    SPList list = List.Read(web);
                    page = WikiPageUtilities.AddWikiPage(list, WikiPageName, true);
                }
                else
                {
                    page = WikiPageUtilities.AddWikiPage(Folder.Read(), WikiPageName, true);
                }
            }
            catch (Exception)
            {
                throw new Exception("The specified wiki page already exists and will not be overwritten.");
            }
            if (page != null)
                WikiPageUtilities.AddLayoutToWikiPage(page.Item, WikiPageLayout);

            return page;
        }
    }
}
