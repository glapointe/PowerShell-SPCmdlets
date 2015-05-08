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
    [Cmdlet(VerbsCommon.New, "SPQuotaTemplate", SupportsShouldProcess = true), 
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Quotas")]
    [CmdletDescription("Creates a new quota template.")]
    [RelatedCmdlets(typeof(SPCmdletGetQuotaTemplate), typeof(SPCmdletSetQuota), typeof(SPCmdletSetQuotaTemplate))]
    [Example(Code = "PS C:\\> $quota = New-SPQuotaTemplate -Name \"Portal\" -StorageMaximumLevel 15GB -StorageWarningLevel 13GB",
        Remarks = "This example creates a new quota template with a max storage value of 15GB and a warning value of 13GB.")]
    [Example(Code = "PS C:\\> $quota = Get-SPQuotaTemplate \"Portal\" | New-SPQuotaTemplate -Name \"Teams\"",
        Remarks = "This example creates a new quota template based on the existing Portal quota template.")]
    public class SPCmdletNewQuotaTemplate : SPNewCmdletBaseCustom<SPQuota>
    {

        [Parameter(Mandatory = false,
            HelpMessage = "Specify whether to limit the amount of storage available on a Site Collection, and set the maximum amount of storage.  When the warning level or maximum storage level is reached, an e-mail is sent to the site administrator to inform them of the issue.")]
        public long? StorageMaximumLevel { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specify whether to limit the amount of storage available on a Site Collection, and set the warning level.  When the warning level or maximum storage level is reached, an e-mail is sent to the site administrator to inform them of the issue.")]
        public long? StorageWarningLevel { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies whether sandboxed solutions with code are allowed for this site collection.  When the maximum usage limit is reached, sandboxed solutions with code are disabled for the rest of the day and an e-mail is sent to the site administrator.")]
        public double? UserCodeMaximumLevel { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies whether sandboxed solutions with code are allowed for this site collection.  When the warning level is reached, an e-mail is sent.")]
        public double? UserCodeWarningLevel { get; set; }

        [Parameter(Mandatory = true, 
            Position = 0,
            HelpMessage = "The name of the new quota template to create.")]
        [ValidateNotNullOrEmpty]
        public string Name { get; set; }

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "An existing quota template to base the new template off of.")]
        public SPQuotaTemplatePipeBind Quota { get; set; }

        protected override SPQuota CreateDataObject()
        {
            SPQuotaTemplate quota = new SPQuotaTemplate();
            quota.Name = Name;

            SPWebService webService = SPWebService.ContentService;
            webService.QuotaTemplates.Add(quota);

            if (Quota != null)
            {
                SPQuota clone = Quota.Read();
                quota.StorageMaximumLevel = clone.StorageMaximumLevel;
                quota.StorageWarningLevel = clone.StorageWarningLevel;
                quota.UserCodeMaximumLevel = clone.UserCodeMaximumLevel;
                quota.UserCodeWarningLevel = clone.UserCodeWarningLevel;
            }

            if (StorageMaximumLevel.HasValue)
            {
                if (StorageMaximumLevel.Value > quota.StorageWarningLevel)
                {
                    quota.StorageMaximumLevel = StorageMaximumLevel.Value;
                    quota.StorageWarningLevel = StorageWarningLevel.Value;
                }
                else
                {
                    quota.StorageWarningLevel = StorageWarningLevel.Value;
                    quota.StorageMaximumLevel = StorageMaximumLevel.Value;
                }
            }

            if (UserCodeMaximumLevel.HasValue)
            {
                if (UserCodeMaximumLevel.Value > quota.UserCodeWarningLevel)
                {
                    quota.UserCodeMaximumLevel = UserCodeMaximumLevel.Value;
                    quota.UserCodeWarningLevel = UserCodeWarningLevel.Value;
                }
                else
                {
                    quota.UserCodeWarningLevel = UserCodeWarningLevel.Value;
                    quota.UserCodeMaximumLevel = UserCodeMaximumLevel.Value;
                }
            }
            webService.Update();

            return quota;
        }

        protected override void InternalValidate()
        {
            if (StorageMaximumLevel.HasValue != StorageWarningLevel.HasValue)
            {
                base.WriteError(new PSArgumentException("Both StorageMaximumLevel and StorageWarningLevel are required if either parameter is provided."), ErrorCategory.SyntaxError, null);
                base.SkipProcessCurrentRecord();
            }
            if (UserCodeMaximumLevel.HasValue != UserCodeWarningLevel.HasValue)
            {
                base.WriteError(new PSArgumentException("Both UserCodeMaximumLevel and UserCodeWarningLevel are required if either parameter is provided."), ErrorCategory.SyntaxError, null);
                base.SkipProcessCurrentRecord();
            }

            if (StorageMaximumLevel.HasValue && StorageWarningLevel.HasValue)
            {
                if (StorageMaximumLevel.Value < 0 || StorageWarningLevel.Value < 0 || StorageWarningLevel.Value > StorageMaximumLevel.Value)
                {
                    base.WriteError(new PSArgumentOutOfRangeException("Storage maximum and warning values must be greater than zero and the warning value must be less than the maximum value."), ErrorCategory.InvalidData, null);
                    base.SkipProcessCurrentRecord();
                }
            }

            if (UserCodeMaximumLevel.HasValue && UserCodeWarningLevel.HasValue)
            {
                if (UserCodeMaximumLevel.Value < 0 || UserCodeWarningLevel.Value < 0 || UserCodeWarningLevel.Value > UserCodeMaximumLevel.Value)
                {
                    base.WriteError(new PSArgumentOutOfRangeException("User code maximum and warning values must be greater than zero and the warning value must be less than the maximum value."), ErrorCategory.InvalidData, null);
                    base.SkipProcessCurrentRecord();
                }
            }

            foreach (SPQuotaTemplate quota in SPWebService.ContentService.QuotaTemplates)
            {
                if (quota.Name.ToLower() == Name)
                {
                    base.WriteError(new PSArgumentException("The quota template name specified already exists."), ErrorCategory.ResourceExists, null);
                    base.SkipProcessCurrentRecord();
                    break;
                }
            }

        }
        
    }
}
