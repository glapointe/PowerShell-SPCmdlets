using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using System.IO;
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.SiteCollections
{
    [Cmdlet("Export", "SPNavigation", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Site Collections")]
    [CmdletDescription("Export the navigation settings of a publishing site.")]
    [RelatedCmdlets(typeof(SPCmdletImportNavigation))]
    [Example(Code = "PS C:\\> Get-SPSite \"http://server_name/\" | Export-SPNavigation -OutputFile \"c:\\nav.xml\"",
        Remarks = "This example exports the navigation settings from the site collection and saves to c:\\nav.xml.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name/subsite\" | Export-SPNavigation -OutputFile \"c:\\nav.xml\"",
        Remarks = "This example exports the navigation settings for http://server_name/subsite and saves to c:\\nav.xml.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name/subsite\" | Export-SPNavigation -OutputFile \"c:\\nav.xml\" -IncludeChildren",
        Remarks = "This example exports the navigation settings for http://server_name/subsite and all its sub-sites and saves to c:\\nav.xml.")]
    public sealed class SPCmdletExportNavigation : SPCmdletCustom
    {
        [Parameter(Mandatory = true, ParameterSetName = "SPSite",
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 0,
        HelpMessage = "Specifies the URL or GUID of the Site whose navigation settings will be exported.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "SPWeb",
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 1,
        HelpMessage = "Specifies the URL or GUID of the Web whose navigation settings will be exported.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The path to the file to save the navigation settings to.")]
        [ValidateDirectoryExistsAndValidFileName]
        [Alias("Path")]
        public string OutputFile { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "SPWeb",
            HelpMessage = "Include all child webs of the specified Web.")]
        public SwitchParameter IncludeChildren { get; set; }


        protected override void InternalProcessRecord()
        {
            XmlDocument xmlDoc = null;
            switch (ParameterSetName)
            {
                case "SPWeb":
                    using (SPWeb web = Web.Read())
                    {
                        try
                        {
                            WriteVerbose("Exporting navigation settings from Web " + web.Url);
                            xmlDoc = Common.SiteCollections.ExportNavigation.GetNavigation(web, IncludeChildren);
                        }
                        finally
                        {
                            web.Dispose();
                            web.Site.Dispose();
                        }
                    }
                    break;
                case "SPSite":
                    using (SPSite site = Site.Read())
                    {
                        try
                        {
                            WriteVerbose("Exporting navigation settings from Site Collection " + site.Url);
                            xmlDoc = Common.SiteCollections.ExportNavigation.GetNavigation(site);
                        }
                        finally
                        {
                            site.Dispose();
                        }
                    }
                    break;
            }
            if (xmlDoc != null && !string.IsNullOrEmpty(OutputFile))
                File.WriteAllText(OutputFile, Utilities.GetFormattedXml(xmlDoc));
            
            WriteResult(xmlDoc);
        }


    }

}
