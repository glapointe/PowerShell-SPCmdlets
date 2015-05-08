using System;
using System.Collections;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Common.WebParts;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet(VerbsCommon.Set, "SPWebPart", SupportsShouldProcess = false, DefaultParameterSetName = "WebPartTitle_File_PropertyHash"),
    SPCmdlet(RequireLocalFarmExist = true,RequireUserFarmAdmin = false)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Sets the state of the web part including adding, moving, opening, closing, deleting, or updating web part properties.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), typeof(SPCmdletGetFile), typeof(SPCmdletGetWebPartList), ExternalCmdlets = new[] { "Get-SPWeb", "Get-SPSite" })]
    [Example(Code = "PS C:\\> Set-SPWebPart -File \"http://server_name/pages/default.aspx\" -WebPartTitle \"My CQWP\" -Action Update -Properties @{\"ItemXslLink\"=\"/Style Library/MyXsl/ItemStyle.xsl\"} -Publish",
        Remarks = "This sets the ItemXslLink property of a content query web part to point to a custom XSL file.")]
    public class SPCmdletSetWebPart : SPCmdletCustom
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
            HelpMessage = "The ID of the Web Part to update.")]
        public string WebPartId { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "WebPartTitle",
            HelpMessage = "The title of the Web Part to update.")]
        public string WebPartTitle { get; set; }

        
        [Parameter(Mandatory = false,
            HelpMessage = "A hash table where each key represents the name of a property whose value should be set to the corresponding key value. Alternatively, provide the path to a file with XML property settings (<Properties><Property Name=\"Name1\">Value1</Property><Property Name=\"Name2\">Value2</Property></Properties>).")]
        public PropertiesPipeBind Properties { get; set; }

        
        [Parameter(Mandatory = true, HelpMessage = "Specify whether to add, close, open, delete, or update the web part.")]
        public SetWebPartStateAction Action { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The name of the web part zone to move the web part to.")]
        public string Zone { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The index within the web part zone to move the web part to.")]
        public int? ZoneIndex { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "If specified the page will be published after adjusting the Web Part.")]
        public SwitchParameter Publish { get; set; }

        protected override void InternalProcessRecord()
        {
            string fileUrl = File.FileUrl;
            string zoneIndex = null;
            if (ZoneIndex.HasValue)
                zoneIndex = ZoneIndex.Value.ToString();

            Hashtable props = null;
            if (Properties != null)
                props = Properties.Read();

            switch (ParameterSetName)
            {
                case "WebPartTitle":
                    SetWebPartState.SetWebPartByTitle(fileUrl, Action, WebPartTitle, Zone, zoneIndex, props, Publish);
                    break;
                case "WebPartID":
                    SetWebPartState.SetWebPartById(fileUrl, Action, WebPartId, Zone, zoneIndex, props, Publish);
                    break;
            }
            base.InternalProcessRecord();
        }
    }
}
