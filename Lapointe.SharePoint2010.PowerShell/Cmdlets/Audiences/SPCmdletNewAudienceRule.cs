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
    [Cmdlet(VerbsCommon.New, "SPAudienceRule", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Audiences")]
    [CmdletDescription("Create a new audience in the specified user profile service application.")]
    [Example(Code = 
        "PS C:\\> [xml]$rules = \"<rules><rule op='Member of' value='sales department' /><rule op='AND' /><rule op='Contains' field='Department' value='Human Resources' /></rules>\"\r\n" + 
        "PS C:\\> New-SPAudienceRule -UserProfileServiceApplication \"30daa535-b0fe-4d10-84b0-fb04029d161a\" -Identity \"Human Resources\" -Clear -Compile\" -Rules $rules",
        Remarks = "This example adds a new rule to the audience named \"Human Resources\" in the user profile service application with ID \"30daa535-b0fe-4d10-84b0-fb04029d161a\".")]
    [RelatedCmdlets(
        typeof(SPCmdletNewAudience), typeof(SPCmdletSetAudience),
        ExternalCmdlets = new[] { "Get-SPServiceApplication" })]
    public class SPCmdletNewAudienceRule : SPCmdletCustom
    {
        AppendOp? _appendOp = AppendOp.AND;

        [Parameter(Mandatory = true,
            ParameterSetName = "UPA",
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the service application that contains the audience to update.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a service application (for example, ServiceApp1); or an instance of a valid SPServiceApplication object.")]
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

        [Parameter(Mandatory = true, HelpMessage = "The name of the audience to update.")]
        [Alias("Name")]
        [ValidateNotNullOrEmpty]
        public string Identity { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Clear existing rules.")]
        public SwitchParameter Clear { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Wraps any existing rules in parantheses.")]
        public SwitchParameter GroupExisting { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Compile the audiences after update.")]
        public SwitchParameter Compile { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The rules XML should be in the following format: " +
            "<rules><rule op='' field='' value='' /></rules>\r\n\r\n" +
            "Values for the \"op\" attribute can be any of \"=,>,>=,<,<=,<>,Contains,Not contains,Reports Under,Member Of,AND,OR,(,)\"\r\n\r\n" +
            "The \"field\" attribute is not required if \"op\" is any of \"Reports Under,Member Of,AND,OR,(,)\"\r\n\r\n" +
            "The \"value\" attribute is not required if \"op\" is any of \"AND,OR,(,)\"\r\n\r\n" +
            "Note that if your rules contain any grouping or mixed logic then you will not be able to manage the rule via the browser.\r\n\r\n" +
            "Example: <rules><rule op='Member of' value='sales department' /><rule op='AND' /><rule op='Contains' field='Department' value='Sales' /></rules>")]
        [ValidateNotNull]
        public XmlDocument Rules { get; set; }


        [Parameter(Mandatory = false, HelpMessage = "Operator used to append to existing rules. Valid values are AND or OR. Default is AND")]
        public AppendOp? AppendOperator
        {
            get { return _appendOp; }
            set { _appendOp = value; }
        }

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

            Common.Audiences.AddAudienceRule.AddRules(context, Identity, Rules.OuterXml, Clear.IsPresent, Compile.IsPresent, GroupExisting.IsPresent, AppendOperator.Value);
        }

    }
}
