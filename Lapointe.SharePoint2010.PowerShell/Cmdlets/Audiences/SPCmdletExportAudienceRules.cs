using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System.Xml;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Audiences
{
    [Cmdlet("Export", "SPAudienceRules", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Audiences")]
    [CmdletDescription("Exports the rules for an audience.")]
    [Example(Code = "PS C:\\> $xml = Export-SPAudienceRules -UserProfileServiceApplication \"30daa535-b0fe-4d10-84b0-fb04029d161a\" -Name \"Human Resources\"",
        Remarks = "This example gets the rules for the \"Human Resources\" audience from the user profile service application with ID \"30daa535-b0fe-4d10-84b0-fb04029d161a\".")]
    [Example(Code = "PS C:\\> $xml = Get-SPServiceApplication | where {$_.TypeName -eq \"User Profile Service Application\"} | Export-SPAudienceRules -Name \"Human Resources\"",
        Remarks = "This example gets the rules for the \"Human Resources\" audience(s) from all configured user profile service applications.")]
    [RelatedCmdlets(
        typeof(SPCmdletExportAudiences), typeof(SPCmdletImportAudiences),
        ExternalCmdlets = new[] { "Get-SPServiceApplication" })]
    public class SPCmdletExportAudienceRules : SPCmdletCustom
    {
        [Parameter(Mandatory = true, 
            ParameterSetName = "UPA",
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the service application that contains the audience rules to export.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a service application (for example, ServiceApp1); or an instance of a valid SPServiceApplication object.")]
        [ValidateNotNull]
        public SPServiceApplicationPipeBind UserProfileServiceApplication { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "UPA",
            HelpMessage = "Specifies the site subscription containing the audience rules to export.\r\n\r\nThe type must be a valid URL, in the form http://server_name; a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a site subscription (for example, SiteSubscription1); or an instance of a valid SiteSubscription object.")]
        public SPSiteSubscriptionPipeBind SiteSubscription { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "SPSite",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Site to use for retrieving the service context. Use this parameter when the service application is not associated with the default proxy group or more than one custom proxy groups.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind ContextSite { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The name of the audience to export.")]
        [Alias("Name")]
        [ValidateNotNullOrEmpty]
        public string Identity { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Shows field and value attributes for every rule.")]
        public SwitchParameter Explicit { get; set; }

        protected override void InternalProcessRecord()
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


            string xml = Common.Audiences.EnumAudienceRules.EnumRules(context, Identity, Explicit.IsPresent);
            if (string.IsNullOrEmpty(xml))
                return;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);
            WriteResult(xml);
        }
    }
}
