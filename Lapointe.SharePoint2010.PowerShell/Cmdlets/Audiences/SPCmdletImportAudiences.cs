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
using System.IO;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using System.ComponentModel;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Audiences
{
    [Cmdlet("Import", "SPAudiences", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Audiences")]
    [CmdletDescription("Imports the audience and corresponding rules from an XML file to the specified user profile service application.")]
    [Example(Code = "PS C:\\> Import-SPAudiences -UserProfileServiceApplication \"30daa535-b0fe-4d10-84b0-fb04029d161a\" -DeleteExisting -Compile -InputFile c:\\audiences.xml",
        Remarks = "This example imports the audience definitions and rules from the audiences.xml file to the user profile service application with ID \"30daa535-b0fe-4d10-84b0-fb04029d161a\".")]
    [RelatedCmdlets(
        typeof(SPCmdletExportAudienceRules), typeof(SPCmdletExportAudiences),
        ExternalCmdlets = new[] { "Get-SPServiceApplication" })]
    public class SPCmdletImportAudiences : SPCmdletCustom
    {
        [Parameter(Mandatory = true,
            ParameterSetName = "UPA",
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the service application that contains the audiences to import.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a service application (for example, ServiceApp1); or an instance of a valid SPServiceApplication object.")]
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

        [Parameter(Mandatory = false, HelpMessage = "Delete existing audiences prior to import.")]
        public SwitchParameter DeleteExisting { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Compile the audiences after import.")]
        public SwitchParameter Compile { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "XML file containing all audiences to import. The file can be generated using Export-SPAudiences.")]
        [ValidateFileExists]
        public string InputFile { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Generate a map file to use for search and replace of Audience IDs. Must be a valid filename.")]
        [ValidateDirectoryExistsAndValidFileName]
        public string MapFile { get; set; }

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

            string xml = File.ReadAllText(InputFile);
            Common.Audiences.ImportAudiences.Import(xml, context, DeleteExisting.IsPresent, Compile.IsPresent, MapFile);
        }
    }
}
