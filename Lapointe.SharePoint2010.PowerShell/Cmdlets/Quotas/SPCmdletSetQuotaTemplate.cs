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
    [Cmdlet(VerbsCommon.Set, "SPQuotaTemplate", SupportsShouldProcess = true), 
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Quotas")]
    [CmdletDescription("Updates an existing quota template.")]
    [RelatedCmdlets(typeof(SPCmdletGetQuotaTemplate), typeof(SPCmdletSetQuota), typeof(SPCmdletNewQuotaTemplate))]
    [Example(Code = "PS C:\\> $quota = Set-SPQuotaTemplate -Name \"Portal\" -StorageMaximumLevel 15GB -StorageWarningLevel 13GB",
        Remarks = "This example updates the Portal quota template and sets the max storage value to 15GB and warning value to 13GB.")]
    [Example(Code = "PS C:\\> Get-SPQuotaTemplate \"Portal\" | Set-SPQuotaTemplate -StorageMaximumLevel 15GB -StorageWarningLevel 13GB",
        Remarks = "This example updates the Portal quota template and sets the max storage value to 15GB and warning value to 13GB.")]
    public class SPCmdletSetQuotaTemplate : SPSetCmdletBaseCustom<SPQuotaTemplate>
    {

        #region Parameters

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
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The name of the quota template to update.")]
        public SPQuotaTemplatePipeBind Identity { get; set; }

        #endregion


        protected override void UpdateDataObject()
        {
            SPQuotaTemplate quota = new SPQuotaTemplate();

            if (Identity != null)
            {
                SPQuotaTemplate clone = Identity.Read();
                quota.Name = clone.Name;
                quota.StorageMaximumLevel = clone.StorageMaximumLevel;
                quota.StorageWarningLevel = clone.StorageWarningLevel;
                quota.UserCodeMaximumLevel = clone.UserCodeMaximumLevel;
                quota.UserCodeWarningLevel = clone.UserCodeWarningLevel;
            }
            else
                throw new SPCmdletException("A quota template is required.");


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
            SPWebService webService = SPWebService.ContentService;
            webService.QuotaTemplates[quota.Name] = quota;

            webService.Update();
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
        }


    }
}
