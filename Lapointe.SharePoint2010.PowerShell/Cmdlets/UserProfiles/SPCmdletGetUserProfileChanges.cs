using System;
using System.Collections.Generic;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.UserProfiles
{
    [SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true, RequireUserMachineAdmin = false)]
    [Cmdlet(VerbsCommon.Get, "SPUserProfileChanges", DefaultParameterSetName = "SPSite")]
    [CmdletGroup("User Profiles")]
    [CmdletDescription("Retrieves user profile property changes for the given interval and context.")]
    [RelatedCmdlets(ExternalCmdlets = new[] { "Get-SPProfileServiceApplication" })]
    [Example(Code = "PS C:\\> Get-SPUserProfileChanges -Site \"http://site/\" -IncludeSingleValueProperty -IncludeMultiValueProperty",
        Remarks = "This example retrieves all single and multi-value property changes for the User Profile Application associated with the \"http://site/\" context.")]
    public class SPCmdletGetUserProfileChanges : SPGetCmdletBase<UserProfileChange>
    {
        [Parameter(HelpMessage = "The interval, in minutes, to retrieve changes from.")]
        public int Interval { get; set; }

        [Parameter(Mandatory = true, 
            ParameterSetName = "UPA", 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
        HelpMessage = "Specifies the service application that contains the user profile manager to retrieve.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a service application (for example, ServiceApp1); or an instance of a valid SPServiceApplication object.")]
        [ValidateNotNull]
        public SPServiceApplicationPipeBind UserProfileServiceApplication { get; set; }

        [Parameter(Mandatory = false, 
            ParameterSetName = "UPA",
            HelpMessage = "Specifies the site subscription containing the user profile manager to retrieve.\r\n\r\nThe type must be a valid URL, in the form http://server_name; a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a site subscription (for example, SiteSubscription1); or an instance of a valid SiteSubscription object.")]
        public SPSiteSubscriptionPipeBind SiteSubscription { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "SPSite",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Site to use for retrieving the service context. Use this parameter when the service application is not associated with the default proxy group or more than one custom proxy groups.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind ContextSite { get; set; }

        [Parameter(HelpMessage = "Include single-value property changes.")]
        public SwitchParameter IncludeSingleValuePropertyChanges { get; set; }

        [Parameter(HelpMessage = "Include multi-value property changes.")]
        public SwitchParameter IncludeMultiValuePropertyChanges { get; set; }

        [Parameter(HelpMessage = "Include anniversary property changes.")]
        public SwitchParameter IncludeAnniversaryChanges { get; set; }

        [Parameter(HelpMessage = "Include colleague property changes.")]
        public SwitchParameter IncludeColleagueChanges { get; set; }

        [Parameter(HelpMessage = "Include organization membership property changes.")]
        public SwitchParameter IncludeOrganizationMembershipChanges { get; set; }

        [Parameter(HelpMessage = "Include distribution list membership property changes.")]
        public SwitchParameter IncludeDistributionListMembershipChanges { get; set; }

        [Parameter(HelpMessage = "Include personalization site property changes.")]
        public SwitchParameter IncludePersonalizationSiteChanges { get; set; }

        [Parameter(HelpMessage = "Include site membership property changes.")]
        public SwitchParameter IncludeSiteMembershipChanges { get; set; }

        protected override IEnumerable<UserProfileChange> RetrieveDataObjects()
        {
            SPServiceContext context = null;
            if (ParameterSetName == "UPA")
            {
                SPSiteSubscriptionIdentifier subId;
                if (SiteSubscription != null)
                {
                    SPSiteSubscription siteSub = SiteSubscription.Read();
                    subId = siteSub.Id;
                }
                else
                    subId = SPSiteSubscriptionIdentifier.Default;

                SPServiceApplication svcApp = UserProfileServiceApplication.Read();
                context = SPServiceContext.GetContext(svcApp.ServiceApplicationProxyGroup, subId);

            }
            else
            {
                using (SPSite site = ContextSite.Read())
                {
                    context = SPServiceContext.GetContext(site);
                }
            }
            UserProfileManager profileManager = new UserProfileManager(context);
            
            DateTime tokenStart = DateTime.UtcNow.Subtract(TimeSpan.FromMinutes(Interval));
            UserProfileChangeToken changeToken = new UserProfileChangeToken(tokenStart);

            UserProfileChangeQuery changeQuery = new UserProfileChangeQuery(false, true);
            changeQuery.ChangeTokenStart = changeToken;
            changeQuery.UserProfile = true;
            changeQuery.SingleValueProperty = IncludeSingleValuePropertyChanges;
            changeQuery.MultiValueProperty = IncludeMultiValuePropertyChanges;
            changeQuery.Anniversary = IncludeAnniversaryChanges;
            changeQuery.Colleague = IncludeColleagueChanges;
            changeQuery.OrganizationMembership = IncludeOrganizationMembershipChanges;
            changeQuery.DistributionListMembership = IncludeDistributionListMembershipChanges;
            changeQuery.PersonalizationSite = IncludePersonalizationSiteChanges;
            changeQuery.SiteMembership = IncludeSiteMembershipChanges;

            UserProfileChangeCollection changes = profileManager.GetChanges(changeQuery);
            
            foreach (UserProfileChange change in changes)
            {
                WriteResult(change);
            }
            return null;
        }
    }
}
