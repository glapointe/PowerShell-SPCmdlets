using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.Common.Audiences;
using Microsoft.Office.Server.Audience;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using System.ComponentModel;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Audiences
{
    [Cmdlet(VerbsCommon.New, "SPAudience", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Audiences")]
    [CmdletDescription("Create a new audience in the specified user profile service application.")]
    [Example(Code = "PS C:\\> New-SPAudience -UserProfileServiceApplication \"30daa535-b0fe-4d10-84b0-fb04029d161a\" -Identity \"Human Resources\" -Membership Any -Owner domain\\user -Description \"All members of the Human Resources department.\"",
        Remarks = "This example creates a new audience named \"Human Resources\" in the user profile service application with ID \"30daa535-b0fe-4d10-84b0-fb04029d161a\".")]
    [RelatedCmdlets(
        typeof(SPCmdletNewAudienceRule), typeof(SPCmdletSetAudience),
        ExternalCmdlets = new[] { "Get-SPServiceApplication" })]
    public class SPCmdletNewAudience : SPNewCmdletBaseCustom<Audience>
    {
        RuleEnum _membership = RuleEnum.Any;

        [Parameter(Mandatory = true,
            ParameterSetName = "UPA",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the service application where the new audience will be created.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a service application (for example, ServiceApp1); or an instance of a valid SPServiceApplication object.")]
        [ValidateNotNull]
        public SPServiceApplicationPipeBind UserProfileServiceApplication { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "UPA",
            HelpMessage = "Specifies the site subscription containing the audiences to update.\r\n\r\nThe type must be a valid URL, in the form http://server_name; a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a site subscription (for example, SiteSubscription1); or an instance of a valid SiteSubscription object.")]
        public SPSiteSubscriptionPipeBind SiteSubscription { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "SPSite",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Site to use for retrieving the service context. Use this parameter when the service application is not associated with the default proxy group or more than one custom proxy groups.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind ContextSite { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The name of the audience to create.")]
        [Alias("Name")]
        [ValidateNotNullOrEmpty]
        public string Identity { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The condition under which membership is met. Valid values are None, Any, All, Mix (default is Any)")]
        public RuleEnum Membership
        {
            get
            {
                return _membership;
            }
            set
            {
                _membership = value;
            }
        }

        [Parameter(Mandatory = false, HelpMessage = "The audience owner in the form domain\\user")]
        public string Owner { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The audience description")]
        public string Description { get; set; }


        protected override Audience CreateDataObject()
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

            return Common.Audiences.CreateAudience.Create(context, Identity, Description, Membership, Owner, false);
        }

    }
}
