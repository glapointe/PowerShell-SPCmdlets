using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet(VerbsCommon.Move, "SPWebPart", SupportsShouldProcess = false, DefaultParameterSetName = "WebPartTitle"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Moves a web part on a page.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), typeof(SPCmdletGetFile), typeof(SPCmdletGetWebPartList), ExternalCmdlets = new[] { "Get-SPWeb", "Get-SPSite" })]
    [Example(Code = "PS C:\\> Move-SPWebPart -File \"http://server_name/pages/default.aspx\" -WebPartTitle \"My CQWP\" -Zone \"LeftZone\" -ZoneIndex 0 -Publish",
        Remarks = "This example moves the web part My CQWP to the top of the LeftZone.")]
    public class SPCmdletMoveWebPart : SPCmdletCustom
    {
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Alias(new string[] { "Url", "Page" })]
        public SPFilePipeBind File { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "WebPartID",
            HelpMessage = "The ID of the Web Part to move.")]
        public string WebPartId { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "WebPartTitle",
            HelpMessage = "The title of the Web Part to move.")]
        public string WebPartTitle { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The name of the web part zone to move the web part to.")]
        public string Zone { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The index within the web part zone to move the web part to.")]
        public int? ZoneIndex { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "If specified, the page will be published after moving the web part.")]
        public SwitchParameter Publish { get; set; }

        protected override void InternalValidate()
        {
            base.InternalValidate();

            if (string.IsNullOrEmpty(Zone) && !ZoneIndex.HasValue)
                throw new SPCmdletException("You must specify at least the -Zone or -ZoneIndex parameters.");
        }

        protected override void InternalProcessRecord()
        {
            string fileUrl = File.FileUrl;
            string zoneIndex = null;
            if (ZoneIndex.HasValue)
                zoneIndex = ZoneIndex.Value.ToString();

            switch (ParameterSetName)
            {
                case "WebPartId":
                    Common.WebParts.MoveWebPart.MoveById(fileUrl, WebPartId, Zone, zoneIndex, Publish);
                    break;
                case "WebPartTitle":
                    Common.WebParts.MoveWebPart.MoveByTitle(fileUrl, WebPartTitle, Zone, zoneIndex, Publish);
                    break;
            }
            base.InternalProcessRecord();
        }
    }
}
