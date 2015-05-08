using System;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using System.Collections.Specialized;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.Common;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Quotas
{
    public class SyncQuotas : SPOperation
    {
        private SPQuotaTemplateCollection m_quotaColl;
        private SPQuotaTemplate m_quota;
        private bool m_setQuota;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncQuotas"/> class.
        /// </summary>
        public SyncQuotas()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", false, null, new SPUrlValidator()));
            parameters.Add(new SPParam("scope", "s", false, "site", new SPRegexValidator("(?i:^Farm$|^WebApplication$|^Site$)")));
            parameters.Add(new SPParam("quota", "q", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("setquota", "set"));


            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nSynchronizes site quota settings with those defined in the quota templates.\r\n\r\nParameters:");
            sb.Append("\r\n\t[-scope <Farm | WebApplication | Site (default)>]");
            sb.Append("\r\n\t[-url <url>]");
            sb.Append("\r\n\t[-quota <quota template name to synchronize (sites using other quotas will not be affected)>]");
            sb.Append("\r\n\t[-setquota (if scope is site or web application and a quota is specified apply the quota to any site that does not currently have a quota assigned)]");

            Init(parameters, sb.ToString());
        }

        /// <summary>
        /// Gets the help message.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <returns></returns>
        public override string GetHelpMessage(string command)
        {
            return HelpMessage;
        }

        /// <summary>
        /// Executes the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public override int Execute(string command, StringDictionary keyValues, out string output)
        {
            output = string.Empty;
            Logger.Verbose = true;

            string scope = Params["scope"].Value.ToLowerInvariant();

            SPFarm farm = SPFarm.Local;
            SPWebService webService = farm.Services.GetValue<SPWebService>("");

            m_quotaColl = webService.QuotaTemplates;
            m_setQuota = Params["setquota"].UserTypedIn;

            if (Params["quota"].UserTypedIn)
            {
                m_quota = m_quotaColl[Params["quota"].Value];
                if (m_quota == null)
                    throw new ArgumentException("The specified quota template name could not be found.");
            }

            SPEnumerator enumerator;
            if (scope == "farm")
            {
                enumerator = new SPEnumerator(SPFarm.Local);
            }
            else if (scope == "webapplication")
            {
                enumerator = new SPEnumerator(SPWebApplication.Lookup(new Uri(Params["url"].Value.TrimEnd('/'))));
            }
            else
            {
                // scope == "site"
                using (SPSite site = new SPSite(Params["url"].Value.TrimEnd('/')))
                {
                    Sync(site);
                }
                return (int)ErrorCodes.NoError;
            }
            
            enumerator.SPSiteEnumerated += enumerator_SPSiteEnumerated;
            enumerator.Enumerate();

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            if (Params["scope"].Validate())
            {
                Params["url"].IsRequired = true;
                Params["url"].Enabled = true;
                if (Params["scope"].Value.ToLowerInvariant() == "farm")
                {
                    Params["url"].IsRequired = false;
                    Params["url"].Enabled = false;
                }
            }
            if (Params["setquota"].UserTypedIn)
            {
                Params["quota"].IsRequired = true;
                if (Params["scope"].Value.ToLowerInvariant() == "farm")
                    throw new SPSyntaxException("Scope of \"farm\" is not valid when \"setquota\" is specified.");
            }
            base.Validate(keyValues);
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
                    Logger.WriteWarning("No quota template has been assigned to site {0}.  Use the -setquota parameter to assign a quota.", site.Url);
                    return;
                }
                Logger.Write("PROGRESS: Synchronizing {0}", site.Url);
                Logger.Write("PROGRESS: No quota template assigned to site.  Assigning template...", site.Url);
                site.Quota = m_quota;

                Logger.Write("PROGRESS: Template \"{0}\" assigned to site.", m_quota.Name);
                return;
            }

            if (m_quota == null || (m_quota != null && currentQuota.QuotaID == m_quota.QuotaID))
            {
                Logger.Write("PROGRESS: Synchronizing {0}", site.Url);
                Logger.Write("PROGRESS: Currently using template \"{0}\".", currentTemplate.Name);

                if (site.Quota.InvitedUserMaximumLevel == currentTemplate.InvitedUserMaximumLevel &&
                    site.Quota.StorageMaximumLevel == currentTemplate.StorageMaximumLevel &&
                    site.Quota.StorageWarningLevel == currentTemplate.StorageWarningLevel &&
                    site.Quota.UserCodeMaximumLevel == currentTemplate.UserCodeMaximumLevel &&
                    site.Quota.UserCodeWarningLevel == currentTemplate.UserCodeWarningLevel)
                {
                    Logger.Write("PROGRESS: No changes necessary, quota already synchronized with template.");
                    return;
                }
                site.Quota = currentTemplate;

                Logger.Write("PROGRESS: Storage maximum updated from {0}MB to {1}MB", 
                    ((currentQuota.StorageMaximumLevel / 1024) / 1024).ToString(),
                    ((site.Quota.StorageMaximumLevel / 1024) / 1024).ToString());
                Logger.Write("PROGRESS: Storage warning updated from {0}MB to {1}MB",
                    ((currentQuota.StorageWarningLevel / 1024) / 1024).ToString(),
                    ((site.Quota.StorageWarningLevel / 1024) / 1024).ToString());
                Logger.Write("PROGRESS: User Code Maximum Level updated from {0} to {1}",
                    currentQuota.UserCodeMaximumLevel.ToString(),
                    site.Quota.UserCodeMaximumLevel.ToString());
                Logger.Write("PROGRESS: User Code Maximum Level updated from {0} to {1}",
                    currentQuota.UserCodeWarningLevel.ToString(),
                    site.Quota.UserCodeWarningLevel.ToString());
                Logger.Write("PROGRESS: Invited user maximum updated from {0} to {1}",
                    currentQuota.InvitedUserMaximumLevel.ToString(),
                    site.Quota.InvitedUserMaximumLevel.ToString());
            }
        }
    }
}
