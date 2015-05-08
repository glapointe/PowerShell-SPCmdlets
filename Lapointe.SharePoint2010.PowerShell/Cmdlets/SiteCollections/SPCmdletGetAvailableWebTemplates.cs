using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using Microsoft.SharePoint;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.SiteCollections
{
    [Cmdlet(VerbsCommon.Get, "SPAvailableWebTemplates"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Site Collections")]
    [CmdletDescription("")]
    [RelatedCmdlets(ExternalCmdlets = new[] {"Get-SPWeb"})]
    [Example(Code = "PS C:\\> Get-SPWeb http://portal | Get-SPAvailableWebTemplates",
        Remarks = "This example returns back all available site templates for http://portal.")]
    public class SPCmdletGetAvailableWebTemplates : SPGetCmdletBaseCustom<SPWebTemplate>
    {
        [Parameter(Mandatory = true, 
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web for which the available site templates will be returned.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        protected override IEnumerable<SPWebTemplate> RetrieveDataObjects()
        {
            List<SPWebTemplate> templates = new List<SPWebTemplate>();

            using (SPWeb web = Web.Read())
            {
                foreach (SPWebTemplate template in web.GetAvailableCrossLanguageWebTemplates())
                {
                    templates.Add(template);
                }

                foreach (SPLanguage lang in web.RegionalSettings.InstalledLanguages)
                {
                    foreach (SPWebTemplate template in web.GetAvailableWebTemplates((uint)lang.LCID))
                    {
                        templates.Add(template);
                    }
                }
            }
            return templates;
        }
    }
}
