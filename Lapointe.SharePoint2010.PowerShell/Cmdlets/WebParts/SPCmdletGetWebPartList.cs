using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using Microsoft.SharePoint.WebPartPages;
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet(VerbsCommon.Get, "SPWebPartList", DefaultParameterSetName="File"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Retrieves a listing of web parts on a page in XML format.")]
    [RelatedCmdlets(typeof(SPCmdletGetLimitedWebPartManager), typeof(Pages.SPCmdletGetPublishingPage), typeof(Lists.SPCmdletGetFile))]
    [Example(Code = "PS C:\\> $xml = Get-SPWebPartList -File \"http://portal/pages/default.aspx\"",
        Remarks = "This example retrieves a listing of web parts associated with the page http://portal/pages/default.aspx.")]
    [Example(Code = "PS C:\\> $xml = Get-SPLimitedWebPartManager \"http://portal/pages/default.aspx\" | Get-SPWebPartList",
        Remarks = "This example retrieves a listing of web parts associated with the page http://portal/pages/default.aspx.")]
    public class SPCmdletGetWebPartList : SPCmdletCustom
    {
        [Parameter(Mandatory = true, 
            ParameterSetName = "Manager", 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPLimitedWebPartManager object.")]
        public SPLimitedWebPartManagerPipeBind Manager { get; set; }

        [Parameter(Mandatory = true, 
            ParameterSetName = "File", 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Alias(new string[] { "Url", "Page" })]
        public SPFilePipeBind File { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Retrieves a minimal XML listing of the web parts on the page.")]
        public SwitchParameter Minimal { get; set; }

        protected override void InternalProcessRecord()
        {
            string xml = null;
            switch (ParameterSetName)
            {
                case "Manager":
                    xml = Common.WebParts.EnumPageWebParts.GetWebPartXml(Manager.PageUrl, Minimal.IsPresent);
                    break;
                case "File":
                    xml = Common.WebParts.EnumPageWebParts.GetWebPartXml(File.FileUrl, Minimal.IsPresent);
                    break;
            }
            XmlDocument doc = new XmlDocument();
            if (xml != null)
                doc.LoadXml(xml);
            WriteObject(doc);
                    
        }

    }
}
