using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Microsoft.SharePoint;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet(VerbsCommon.Set, "SPContentQueryWebPartTarget", SupportsShouldProcess = false, DefaultParameterSetName = "SPSiteWithTitle"),
    SPCmdlet(RequireLocalFarmExist = true,RequireUserFarmAdmin = false)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Retargets a Content Query web part (do not provide list or site if you wish to show items from all sites in the containing site collection).")]
    [RelatedCmdlets(typeof(SPCmdletGetList), typeof(SPCmdletGetFile), typeof(SPCmdletGetWebPartList), ExternalCmdlets = new[] { "Get-SPWeb", "Get-SPSite" })]
    [Example(Code = "PS C:\\> Set-SPContentQueryWebPartTarget -List \"http://server_name/lists/mylist\" -File \"http://server_name/pages/default.aspx\" -WebPartTitle \"My CQWP\" -AllMatching -Publish",
        Remarks = "This example targets the web part My CQWP to the list mylist.")]
    public class SPCmdletSetContentQueryWebPartTarget : SPCmdletCustom
    {
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Alias(new string[] { "Url", "Page" })]
        public SPFilePipeBind File { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "SPSiteWithID",
            HelpMessage = "The ID of the Web Part to update.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "SPListWithID",
            HelpMessage = "The ID of the Web Part to update.")]
        public string WebPartId { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "SPSiteWithTitle",
            HelpMessage = "The title of the Web Part to update.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "SPListWithTitle",
            HelpMessage = "The title of the Web Part to update.")]
        public string WebPartTitle { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "SPSiteWithTitle",
            HelpMessage = "If more than one Content Query Web Part is found matching the title then update all matches. If not specified then an exception is thrown.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "SPListWithTitle",
            HelpMessage = "If more than one Content Query Web Part is found matching the title then update all matches. If not specified then an exception is thrown.")]
        public SwitchParameter AllMatching { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "SPSiteWithID",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Site to show items from.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [Parameter(Mandatory = true, ParameterSetName = "SPSiteWithTitle",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Site to show items from.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "SPListWithID",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The list to point the CQWP to.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "SPListWithTitle",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The list to point the CQWP to.")]
        public SPListPipeBind List { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "SPListWithID",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The web containing the list. This parameter is required if the List parameter is a relative URL to a list.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "SPListWithTitle",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The web containing the list. This parameter is required if the List parameter is a relative URL to a list.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The list type, or template, to show items from.")]
        public string ListType { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "If specified the page will be published after adjusting the Web Part.")]
        public SwitchParameter Publish { get; set; }

        protected override void InternalProcessRecord()
        {
            string fileUrl = File.FileUrl;
            string listUrl = null;
            if (List != null)
            {
                SPList list = null;
                if (Web != null)
                    list = List.Read(Web.Read());
                else
                    list = List.Read();

                listUrl = list.ParentWeb.Site.MakeFullUrl(list.RootFolder.ServerRelativeUrl);
            }
            string siteUrl = null;
            if (Site != null)
                siteUrl = Site.Read().Url;

            switch (ParameterSetName)
            {
                case "SPSiteWithID":
                    Common.WebParts.RetargetContentQueryWebPart.Retarget(fileUrl, false, WebPartId, null, listUrl, ListType, siteUrl, Publish);
                    break;
                case "SPSiteWithTitle":
                    Common.WebParts.RetargetContentQueryWebPart.Retarget(fileUrl, AllMatching, null, WebPartTitle, listUrl, ListType, siteUrl, Publish);
                    break;
                case "SPListWithID":
                    Common.WebParts.RetargetContentQueryWebPart.Retarget(fileUrl, false, WebPartId, null, listUrl, ListType, siteUrl, Publish);
                    break;
                case "SPListWithTitle":
                    Common.WebParts.RetargetContentQueryWebPart.Retarget(fileUrl, AllMatching, null, WebPartTitle, listUrl, ListType, siteUrl, Publish);
                    break;
            }
            base.InternalProcessRecord();
        }
    }
}
