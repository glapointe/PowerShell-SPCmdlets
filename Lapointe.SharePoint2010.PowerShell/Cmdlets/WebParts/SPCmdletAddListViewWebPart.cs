using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.WebPartPages;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet(VerbsCommon.Add, "SPListViewWebPart", DefaultParameterSetName = "File"),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Adds an XSLT List View Web Part to the specified web part page.")]
    [RelatedCmdlets(typeof(SPCmdletGetLimitedWebPartManager), typeof(Pages.SPCmdletGetPublishingPage), typeof(Lists.SPCmdletGetFile))]
    [Example(Code = "PS C:\\> Add-SPListViewWebPart -File \"http://portal/pages/default.aspx\" -List http://portal/Lists/News -WebPartTitle \"News\" -ViewTitle \"All Items\" -Zone \"Left\" -ZoneIndex 0 -Publish",
        Remarks = "This example adds a news list view web part to the page http://portal/pages/default.aspx.")]
    [Example(Code = "PS C:\\> Add-SPListViewWebPart -File \"http://portal/sitepages/home.aspx\" -List http://portal/Lists/News -WebPartTitle \"News\" -ViewTitle \"All Items\" -Row \"1\" -Column 1 -AddSpace -Publish",
        Remarks = "This example adds a news list view web part to the wiki page http://portal/sitepages/home.aspx.")]
    public class SPCmdletAddListViewWebPart : SPNewCmdletBaseCustom<Microsoft.SharePoint.WebPartPages.WebPart>
    {
        [Parameter(Mandatory = true,
            ParameterSetName = "Manager_WebPartPage",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPLimitedWebPartManager object.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Manager_WikiPage",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPLimitedWebPartManager object.")]
        public SPLimitedWebPartManagerPipeBind Manager { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "File_WebPartPage",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "File_WikiPage",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Alias(new string[] { "Url", "Page" })]
        public SPFilePipeBind File { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 1,
            HelpMessage = "The list to whose view will be added to the page.\r\n\r\nThe value must be a valid URL in the form http://server_name/lists/listname or /lists/listname. If a server relative URL is provided then the Web parameter must be provided.")]
        [ValidateNotNull]
        public SPListPipeBind List { get; set; }

        [Parameter(Mandatory = false,
            Position = 2,
            ValueFromPipeline = true,
            HelpMessage = "Specifies the URL or GUID of the Web containing the list whose view will be added to the page.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = false, Position = 3, HelpMessage = "The title to set the web part to. The list title will be used if not provided.")]
        public string WebPartTitle { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "File_WebPartPage",
            Position = 4, HelpMessage = "The name of the web part zone to add the list view web part to.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Manager_WebPartPage",
            Position = 4, HelpMessage = "The name of the web part zone to add the list view web part to.")]
        public string Zone { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "File_WebPartPage",
             Position = 5, HelpMessage = "The index within the web part zone to add the list view web part to.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_WebPartPage",
             Position = 5, HelpMessage = "The index within the web part zone to add the list view web part to.")]
        public int ZoneIndex { get; set; }

        [Parameter(Mandatory = false, Position = 6, HelpMessage = "The name of the view to use for the list view web part.")]
        public string ViewTitle { get; set; }

        [Parameter(Mandatory = false, Position = 7, HelpMessage = "If specified, the page will be published after adding the list view web part.")]
        public SwitchParameter Publish { get; set; }

        [Parameter(Mandatory = false, Position = 8, HelpMessage = "If specified, the web part title will link back to the default view for the list.")]
        public SwitchParameter LinkTitle { get; set; }


        [Parameter(
            ParameterSetName = "File_WikiPage",
            Position = 9,
            Mandatory = true,
            HelpMessage = "The zone to add the web part to.")]
        [Parameter(
            ParameterSetName = "Manager_WikiPage",
            Position = 9,
            Mandatory = true,
            HelpMessage = "The zone to add the web part to.")]
        public int Row { get; set; }

        [Parameter(
            ParameterSetName = "File_WikiPage",
            Position = 10,
            Mandatory = true,
            HelpMessage = "The zone index to add the web part to.")]
        [Parameter(
            ParameterSetName = "Manager_WikiPage",
            Position = 10,
            Mandatory = true,
            HelpMessage = "The zone index to add the web part to.")]
        public int Column { get; set; }

        [Parameter(
            ParameterSetName = "File_WikiPage",
            Position = 11,
            Mandatory = false,
            HelpMessage = "Add some space before the web part.")]
        [Parameter(
            ParameterSetName = "Manager_WikiPage",
            Position = 11,
            Mandatory = false,
            HelpMessage = "Add some space before the web part.")]
        public SwitchParameter AddSpace { get; set; }

#if !SP2010
        [Parameter(
            Position = 12,
            Mandatory = false,
            HelpMessage = "Set a specific JSLink file.")]
#endif
        public string JSLink { get; set; }


        [Parameter(Mandatory = false, HelpMessage = "The chrome settings for the web part.")]
        public PartChromeType ChromeType
        {
            get
            {
                if (Fields["ChromeType"] == null)
                    return PartChromeType.Default;
                return (PartChromeType)Fields["ChromeType"];
            }
            set { Fields["ChromeType"] = value; }
        }

        protected override Microsoft.SharePoint.WebPartPages.WebPart CreateDataObject()
        {
            SPList list;
            if (Web == null)
                list = List.Read();
            else
                list = List.Read(Web.Read());

            string listUrl = list.ParentWeb.Site.MakeFullUrl(list.RootFolder.ServerRelativeUrl);
            switch (ParameterSetName)
            {
                case "Manager_WebPartPage":
                    return Common.WebParts.AddListViewWebPart.Add(Manager.PageUrl, listUrl, WebPartTitle, ViewTitle, Zone, ZoneIndex, LinkTitle, JSLink, ChromeType, Publish);
                case "Manager_WikiPage":
                    return Common.WebParts.AddListViewWebPart.AddToWikiPage(Manager.PageUrl, listUrl, WebPartTitle, ViewTitle, Row, Column, LinkTitle, JSLink, ChromeType, AddSpace, Publish);
                case "File_WebPartPage":
                    return Common.WebParts.AddListViewWebPart.Add(File.FileUrl, listUrl, WebPartTitle, ViewTitle, Zone, ZoneIndex, LinkTitle, JSLink, ChromeType, Publish);
                case "File_WikiPage":
                    return Common.WebParts.AddListViewWebPart.AddToWikiPage(File.FileUrl, listUrl, WebPartTitle, ViewTitle, Row, Column, LinkTitle, JSLink, ChromeType, AddSpace, Publish);
            }
            return null;
        }
    }
}
