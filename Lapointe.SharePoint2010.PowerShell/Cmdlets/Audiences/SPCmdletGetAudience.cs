using System.Collections.Generic;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.Office.Server.Audience;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Audiences
{
    [Cmdlet(VerbsCommon.Get, "SPAudience", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Audiences")]
    [CmdletDescription("The Get-SPAudience cmdlet returns back one or more audiences for a given User Profile Service Application.")]
    [Example(Code = "PS C:\\> Get-SPAudience -UserProfileServiceApplication \"30daa535-b0fe-4d10-84b0-fb04029d161a\" -Name \"Human Resources\"",
        Remarks = "This example gets the \"Human Resources\" audience from the user profile service application with ID \"30daa535-b0fe-4d10-84b0-fb04029d161a\".")]
    [Example(Code = "PS C:\\> Get-SPServiceApplication | where {$_.TypeName -eq \"User Profile Service Application\"} | Get-SPAudience -Name \"Human Resources\"",
        Remarks = "This example gets the \"Human Resources\" audience(s) from all configured user profile service applications.")]
    [Example(Code = "PS C:\\> Get-SPServiceApplication | where {$_.TypeName -eq \"User Profile Service Application\"} | Get-SPAudience",
        Remarks = "This example gets all audiences from all configured user profile service applications.")]
    [RelatedCmdlets(
        typeof(SPCmdletExportAudienceRules), typeof(SPCmdletExportAudiences),
        typeof(SPCmdletImportAudiences), typeof(SPCmdletNewAudience), typeof(SPCmdletNewAudienceRule),
        typeof(SPCmdletRemoveAudience), typeof(SPCmdletSetAudience),
        ExternalCmdlets = new[] { "Get-SPServiceApplication" })]
    public class SPCmdletGetAudience : SPGetCmdletBaseCustom<Audience>
    {
        [Parameter(Mandatory = true,
            ParameterSetName = "UPA",
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the service application that contains the audience to retrieve.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a service application (for example, ServiceApp1); or an instance of a valid SPServiceApplication object.")]
        [ValidateNotNull]
        public SPServiceApplicationPipeBind UserProfileServiceApplication { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "UPA",
            HelpMessage = "Specifies the site subscription containing the audiences to retrieve.\r\n\r\nThe type must be a valid URL, in the form http://server_name; a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a site subscription (for example, SiteSubscription1); or an instance of a valid SiteSubscription object.")]
        public SPSiteSubscriptionPipeBind SiteSubscription { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "SPSite",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Site to use for retrieving the service context. Use this parameter when the service application is not associated with the default proxy group or more than one custom proxy groups.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind ContextSite { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The name of the audience to retrieve.")]
        [Alias("Name")]
        public string Identity { get; set; }

        protected override IEnumerable<Audience> RetrieveDataObjects()
        {
            List<Audience> audiences = new List<Audience>();

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
            AudienceManager manager = new AudienceManager(context);

            if (!string.IsNullOrEmpty(Identity) && !manager.Audiences.AudienceExist(Identity))
            {
                throw new SPException("Audience name does not exist");
            }
            if (!string.IsNullOrEmpty(Identity))
            {
                audiences.Add(manager.Audiences[Identity]);
                return audiences;
            }

            foreach (Audience audience in manager.Audiences)
            {
                audiences.Add(audience);
            }
            return audiences;
        }

       
    }
}
