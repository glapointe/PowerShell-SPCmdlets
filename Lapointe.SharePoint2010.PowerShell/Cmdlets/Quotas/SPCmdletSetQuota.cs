using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.Win32;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation.Internal;
using Lapointe.SharePoint.PowerShell.Common;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Quotas
{
    [Cmdlet(VerbsCommon.Set, "SPQuota", SupportsShouldProcess = true, DefaultParameterSetName = "SPSite2"), 
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Quotas")]
    [CmdletDescription("Synchronizes all assigned quotas with the corresponding quota template.")]
    [RelatedCmdlets(typeof(SPCmdletGetQuotaTemplate), typeof(SPCmdletSetQuota), typeof(SPCmdletNewQuotaTemplate))]
    [Example(Code = "PS C:\\> Get-SPWebApplication http://server_name | Set-SPQuota -SyncExistingOnly",
        Remarks = "This example synchronizes all site collections in http://server_name with the assigned quota template.")]
    [Example(Code = "PS C:\\> Get-SPWebApplication http://server_name | Set-SPQuota -QuotaTempate (Get-SPQuotaTemplate \"Portal\")",
        Remarks = "This example sets all site collections in http://server_name with the Portal quota template.")]
    public class SPCmdletSetQuota : SPSetCmdletBaseCustom<PSObject>
    {
        private SPQuotaTemplateCollection m_quotaColl;
        private SPQuotaTemplate m_quota;
        private bool m_setQuota;
        private bool m_whatIf = false;

        #region Parameters

        [Parameter(ParameterSetName = "SPSite2",
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Synchronize site collection quota settings with the associated template only (do not assign the template if missing).")]
        [Parameter(ParameterSetName = "SPWebApplication2",
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Synchronize site collection quota settings with the associated template only (do not assign the template if missing).")]
        [Parameter(ParameterSetName = "SPFarm2",
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Synchronize site collection quota settings with the associated template only (do not assign the template if missing).")]
        public SwitchParameter SyncExistingOnly { get; set; }

        [Parameter(ParameterSetName = "SPWebApplication1",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The web application containing the sites whose quota will be set or synchronized.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        [Parameter(ParameterSetName = "SPWebApplication2",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The web application containing the sites whose quota will be set or synchronized.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        public SPWebApplicationPipeBind WebApplication { get; set; }

        [Parameter(ParameterSetName = "SPSite1",
            Mandatory = true, 
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The site whose quota will be set or synchronized.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [Parameter(ParameterSetName = "SPSite2",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The site whose quota will be set or synchronized.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        [Parameter(ParameterSetName = "SPFarm1",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Set or synchronize all site collection quotas within the farm.")]
        [Parameter(ParameterSetName = "SPFarm2",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Set or synchronize all site collection quotas within the farm.")]
        public SPFarmPipeBind Farm { get; set; }

        [Parameter(ParameterSetName = "SPSite1",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The quota template to associate with the site collections in the scope.")]
        [Parameter(ParameterSetName = "SPWebApplication1",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The quota template to associate with the site collections in the scope.")]
        [Parameter(ParameterSetName = "SPFarm1",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The quota template to associate with the site collections in the scope.")]
        public SPQuotaTemplatePipeBind QuotaTemplate { get; set; }

        #endregion

        protected override void UpdateDataObject()
        {

            ShouldProcessReason reason;
            if (!base.ShouldProcess("Set Quotas", null, null, out reason))
            {
                if (reason == ShouldProcessReason.WhatIf)
                {
                    m_whatIf = true;
                }
           }

            SPFarm farm = SPFarm.Local;
            SPWebService webService = farm.Services.GetValue<SPWebService>("");

            m_quotaColl = webService.QuotaTemplates;
            m_setQuota = !SyncExistingOnly.IsPresent;


            if (QuotaTemplate != null)
            {
                m_quota = QuotaTemplate.Read();
                if (m_quota == null)
                    throw new SPCmdletException("The specified quota template name could not be found.");
            }

            SPEnumerator enumerator = null;
            switch (ParameterSetName)
            {
                case "SPSite1":
                case "SPSite2":
                    Sync(Site.Read());
                    return;
                case "SPWebApplication1":
                case "SPWebApplication2":
                    enumerator = new SPEnumerator(WebApplication.Read());
                    break;
                default:
                    enumerator = new SPEnumerator(Farm.Read());
                    break;
            }

            enumerator.SPSiteEnumerated += enumerator_SPSiteEnumerated;
            enumerator.Enumerate();
        }


        protected override void InternalValidate()
        {
            if (!SyncExistingOnly.IsPresent && (QuotaTemplate == null || QuotaTemplate.Read() == null))
            {
                throw new SPCmdletException("A valid quota template is required if not synchronizing templates.");
            }
        }

        /// <summary>
        /// Handles the SPSiteEnumerated event of the enumerator control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="Lapointe.SharePoint.STSADM.Commands.OperationHelpers.SPEnumerator.SPSiteEventArgs"/> instance containing the event data.</param>
        private void enumerator_SPSiteEnumerated(object sender, SPEnumerator.SPSiteEventArgs e)
        {
            Sync(e.Site);
        }

        /// <summary>
        /// Syncs the specified site quota with the quota template.
        /// </summary>
        /// <param name="site">The site.</param>
        private void Sync(SPSite site)
        {
            SPQuota currentQuota = site.Quota;
            SPQuotaTemplate currentTemplate = null;
            foreach (SPQuotaTemplate quota in m_quotaColl)
            {
                if (currentQuota.QuotaID == quota.QuotaID)
                {
                    currentTemplate = quota;
                    break;
                }
            }

            if (currentTemplate == null)
            {
                if (!m_setQuota)
                {
                    WriteWarning("No quota template has been assigned to site {0}.  Use the -setquota parameter to assign a quota.", site.Url);
                    return;
                }
                WriteVerbose(string.Format("PROGRESS: Synchronizing {0}", site.Url));
                WriteVerbose(string.Format("PROGRESS: No quota template assigned to site.  Assigning template...", site.Url));
                
                if (!m_whatIf)
                    site.Quota = m_quota;

                WriteVerbose(string.Format("PROGRESS: Template \"{0}\" assigned to site.", m_quota.Name));
                return;
            }

            if (m_quota == null || (m_quota != null && currentQuota.QuotaID == m_quota.QuotaID))
            {
                WriteVerbose(string.Format("PROGRESS: Synchronizing {0}", site.Url));
                WriteVerbose(string.Format("PROGRESS: Currently using template \"{0}\".", currentTemplate.Name));

                if (site.Quota.InvitedUserMaximumLevel == currentTemplate.InvitedUserMaximumLevel &&
                    site.Quota.StorageMaximumLevel == currentTemplate.StorageMaximumLevel &&
                    site.Quota.StorageWarningLevel == currentTemplate.StorageWarningLevel &&
                    site.Quota.UserCodeMaximumLevel == currentTemplate.UserCodeMaximumLevel &&
                    site.Quota.UserCodeWarningLevel == currentTemplate.UserCodeWarningLevel)
                {
                    WriteVerbose("PROGRESS: No changes necessary, quota already synchronized with template.");
                    return;
                }
                if (!m_whatIf)
                    site.Quota = currentTemplate;

                WriteVerbose(string.Format("PROGRESS: Storage maximum updated from {0}MB to {1}MB",
                    ((currentQuota.StorageMaximumLevel / 1024) / 1024).ToString(),
                    ((site.Quota.StorageMaximumLevel / 1024) / 1024).ToString()));
                WriteVerbose(string.Format("PROGRESS: Storage warning updated from {0}MB to {1}MB",
                    ((currentQuota.StorageWarningLevel / 1024) / 1024).ToString(),
                    ((site.Quota.StorageWarningLevel / 1024) / 1024).ToString()));
                WriteVerbose(string.Format("PROGRESS: User Code Maximum Level updated from {0} to {1}",
                    currentQuota.UserCodeMaximumLevel.ToString(),
                    site.Quota.UserCodeMaximumLevel.ToString()));
                WriteVerbose(string.Format("PROGRESS: User Code Maximum Level updated from {0} to {1}",
                    currentQuota.UserCodeWarningLevel.ToString(),
                    site.Quota.UserCodeWarningLevel.ToString()));
                WriteVerbose(string.Format("PROGRESS: Invited user maximum updated from {0} to {1}",
                    currentQuota.InvitedUserMaximumLevel.ToString(),
                    site.Quota.InvitedUserMaximumLevel.ToString()));
            }
        }
    }
}
