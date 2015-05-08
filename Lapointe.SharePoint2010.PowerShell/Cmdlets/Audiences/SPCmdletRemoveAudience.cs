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
using Lapointe.PowerShell.MamlGenerator.Attributes;
using System.ComponentModel;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Audiences
{
    [Cmdlet("Remove", "SPAudience", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Audiences")]
    [CmdletDescription("The Remove-SPAudience cmdlet deletes the specified audience from the given User Profile Service Application.")]
    [Example(Code = "PS C:\\> Remove-SPAudience -UserProfileServiceApplication \"30daa535-b0fe-4d10-84b0-fb04029d161a\" -Name \"Human Resources\"",
        Remarks = "This example deletes the \"Human Resources\" audience from the user profile service application with ID \"30daa535-b0fe-4d10-84b0-fb04029d161a\".")]
    [Example(Code = "PS C:\\> Remove-SPAudience -UserProfileServiceApplication \"30daa535-b0fe-4d10-84b0-fb04029d161a\" -Name \"Human Resources\" -DeleteRulesOnly",
        Remarks = "This example deletes only the rules from the \"Human Resources\" audience located in the user profile service application with ID \"30daa535-b0fe-4d10-84b0-fb04029d161a\".")]
    [RelatedCmdlets(
        typeof(SPCmdletExportAudienceRules), typeof(SPCmdletExportAudiences), 
        typeof(SPCmdletImportAudiences), typeof(SPCmdletNewAudience),
        typeof(SPCmdletNewAudienceRule), typeof(SPCmdletSetAudience),
        ExternalCmdlets = new[] { "Get-SPServiceApplication" })]
    public class SPCmdletRemoveAudience : SPRemoveCmdletBaseCustom<string>
    {
        SPServiceContext _context = null;

        [Parameter(Mandatory = true,
            ParameterSetName = "UPA",
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the service application that contains the audience to remove.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a service application (for example, ServiceApp1); or an instance of a valid SPServiceApplication object.")]
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

        [Parameter(Mandatory = true, HelpMessage = "The name of the audience to remove.")]
        [Alias("Name")]
        [ValidateNotNullOrEmpty]
        public string Identity { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Delete only the rules associated with the audience.")]
        public SwitchParameter DeleteRulesOnly { get; set; }

        protected override void InternalValidate()
        {
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
                _context = SPServiceContext.GetContext(svcApp.ServiceApplicationProxyGroup, subId);

            }
            else
            {
                using (SPSite site = ContextSite.Read())
                {
                    _context = SPServiceContext.GetContext(site);
                }
            }

            if (!string.IsNullOrEmpty(Identity))
            {
                base.DataObject = Identity;
            }
            if (base.DataObject == null)
            {
                base.WriteError(new PSArgumentException("A valid audience name and service application must be provided."), ErrorCategory.InvalidArgument, null);
                base.SkipProcessCurrentRecord();
            }
        }

        protected override void DeleteDataObject()
        {
            if (DataObject != null)
            {
                Common.Audiences.DeleteAudience.Delete(_context, DataObject, DeleteRulesOnly.IsPresent);
            }
        }
    }
}
