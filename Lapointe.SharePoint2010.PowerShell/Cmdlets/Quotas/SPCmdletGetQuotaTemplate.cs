using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Quotas
{
    [Cmdlet(VerbsCommon.Get, "SPQuotaTemplate"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Quotas")]
    [CmdletDescription("Retrieve one or more quota templates.")]
    [RelatedCmdlets(typeof(SPCmdletNewQuotaTemplate), typeof(SPCmdletSetQuota), typeof(SPCmdletSetQuotaTemplate))]
    [Example(Code = "PS C:\\> $quotas = Get-SPQuotaTemplate",
        Remarks = "This example retrieves all quota templates.")]
    [Example(Code = "PS C:\\> $quota = Get-SPQuotaTemplate \"Portal\"",
        Remarks = "This example retrieves the \"Portal\" quota template.")]
    public class SPCmdletGetQuotaTemplate : SPGetCmdletBaseCustom<SPQuotaTemplate>
    {

        [Parameter(Mandatory = false, 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The name of the quota template to return.")]
        [Alias(new string[] { "Name" })]
        public SPQuotaTemplatePipeBind Identity { get; set; }

        protected override void InternalValidate()
        {
            if (this.Identity != null)
            {
                base.DataObject = this.Identity.Read();
                if (base.DataObject == null)
                {
                    base.WriteError(new PSArgumentException("The quota template does not exist."), ErrorCategory.InvalidArgument, this.Identity);
                    base.SkipProcessCurrentRecord();
                }
            }
        }

        protected override IEnumerable<SPQuotaTemplate> RetrieveDataObjects()
        {
            List<SPQuotaTemplate> list = new List<SPQuotaTemplate>();
            if (base.DataObject != null)
            {
                list.Add(base.DataObject);
                return list;
            }
            SPWebService webService = SPWebService.ContentService;
            if (webService != null)
            {
                foreach (SPQuotaTemplate quota in webService.QuotaTemplates)
                {
                    list.Add(quota);
                }
            }

            return list;
        }


    }
}
