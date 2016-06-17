using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using System.IO;
using System.Collections;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Publishing;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.SiteCollections
{
#if SP2010
    [Cmdlet("Repair", "SPSite", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletDescription("Repair a site collection that was created using the export of a web.")]
    [Example(Code = "PS C:\\> Get-SPSite http://portal/sites/newsitecoll | Repair-SPSite -SourceSite \"http://portal/\"",
        Remarks = "This example repairs the site collection located at http://portal/sites/newsitecoll using the site collection http://portal as the model site for the repairs.")]
#else
    [Cmdlet("Repair", "SPMigratedSite", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletDescription("Repair a site collection that was created using the export of a web. This cmdlet is the equivalent to the SharePoint 2010 Repair-SPSite cmdlet - it was renamed because SharePoint 2013 introduced a Repair-SPSite cmdlet which does different things.")]
    [Example(Code = "PS C:\\> Get-SPSite http://portal/sites/newsitecoll | Repair-SPMigratedSite -SourceSite \"http://portal/\"",
        Remarks = "This example repairs the site collection located at http://portal/sites/newsitecoll using the site collection http://portal as the model site for the repairs.")]
#endif
    [CmdletGroup("Site Collections")]
    [RelatedCmdlets(typeof(SPCmdletConvertToSite), typeof(Lists.SPCmdletExportWeb2), typeof(Lists.SPCmdletImportWeb2), ExternalCmdlets = new[] { "Export-SPWeb", "Import-SPWeb" })]
    public class SPCmdletRepairSite : SPCmdletCustom
    {
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The site to repair.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind TargetSite { get; set; }


        [Parameter(Mandatory = true,
            HelpMessage = "The model or source site collection to use as the basis for the repairs.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind SourceSite { get; set; }


        protected override void InternalProcessRecord()
        {
            Common.SiteCollections.RepairSiteCollectionImportedFromSubSite.RepairSite(SourceSite.SiteUrl, TargetSite.SiteUrl);
        }
    }
}
