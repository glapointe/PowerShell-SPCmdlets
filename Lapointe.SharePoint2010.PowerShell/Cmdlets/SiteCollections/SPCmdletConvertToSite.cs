using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using System.IO;
using System.Collections;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Publishing;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.SiteCollections
{
    [Cmdlet("ConvertTo", "SPSite", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Site Collections")]
    [CmdletDescription("Converts a sub-site to a site collection.")]
    [RelatedCmdlets(typeof(SPCmdletRepairSite), typeof(Lists.SPCmdletExportWeb2), typeof(Lists.SPCmdletImportWeb2), ExternalCmdlets = new[] {"Export-SPWeb", "Import-SPWeb"})]
    [Example(Code = "PS C:\\> Get-SPWeb http://portal/subweb1 | ConvertTo-SPSite -TargetUrl \"http://portal/sites/newsitecoll\" -OwnerLogin domain\\user",
        Remarks = "This example converts the web located at http://portal/subweb1 to a site collection located at http://portal/sites/newsitecoll")]
    public class SPCmdletConvertToSite : SPNewCmdletBaseCustom<SPSite>
    {
        SPWebApplication webApp;
        SPContentDatabase contentDb;
        string quotaTemplate;
        SPSiteSubscription siteSubscription;
        bool useHostHeaderAsSiteName;
        Uri siteUri;

        [Parameter(Mandatory = true, 
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to be converted.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind SourceWeb { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The target URL of the new site collection. Example: http://server_name/sites/newsite1")]
        public string TargetUrl { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The site collection's owner login in the form domain\\user.")]
        public string OwnerLogin { get; set; }

        [Parameter(Mandatory = false, 
            HelpMessage = "The owner's email in the form user@domain.com")]
        public string OwnerEmail { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The secondary owner's email in the form user@domain.com")]
        public string SecondaryEmail { get; set; }

        [Parameter(Mandatory = false, 
            HelpMessage = "The seconary owner's login in the form domain\\user.")]
        public string SecondaryOwnerLogin { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The content database to store the new site collection within.")]
        public SPContentDatabasePipeBind ContentDatabase { get; set; }

        [Parameter(Mandatory = false)]
        public SPSiteSubscriptionPipeBind SiteSubscription { get; set; }

        [Parameter(Mandatory = false, 
            HelpMessage = "The name or title of the new site collection.")]
        public string Name { get; set; }

        [Parameter]
        public uint Language { get; set; }

        [Parameter]
        public string Description { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "If an export of the source web was already performed provide the path to the exported file to prevent another export from occuring.")]
        public string ExportedFile { get; set; }

        [Parameter]
        public SwitchParameter NoFileCompression { get; set; }

        [Parameter]
        public SwitchParameter SuppressAfterEvents { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Delete the source web after the conversion completes.")]
        public SwitchParameter DeleteSource { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Create a managed path for the new site collection.")]
        public SwitchParameter CreateManagedPath { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Halt the export or import operation on warning.")]
        public SwitchParameter HaltOnWarning { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Halt the export or import operation on error.")]
        public SwitchParameter HaltOnFatalError { get; set; }

        [Parameter]
        public SwitchParameter IncludeUserSecurity { get; set; }

        [Parameter]
        public SPWebApplicationPipeBind HostHeaderWebApplication { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The quota template to assign to the new site collection.")]
        public SPQuotaTemplatePipeBind QuotaTemplate { get; set; }

        protected override SPSite CreateDataObject()
        {
            if (useHostHeaderAsSiteName)
            {
                bool foundMatch = false;
                foreach (SPIisSettings settings in this.webApp.IisSettings.Values)
                {
                    foreach (SPSecureBinding binding in settings.SecureBindings)
                    {
                        if (binding.Port == this.siteUri.Port)
                        {
                            foundMatch = true;
                            break;
                        }
                    }
                    if (foundMatch)
                    {
                        break;
                    }
                    foreach (SPServerBinding binding2 in settings.ServerBindings)
                    {
                        if (binding2.Port == this.siteUri.Port)
                        {
                            foundMatch = true;
                            break;
                        }
                    }
                    if (foundMatch)
                    {
                        break;
                    }
                }
                if (!foundMatch)
                {
                    base.WriteWarning(SPResource.GetString("HostHeaderDoesNotMatchWebAppPort", new object[0]));
                }
            }

            if ((null != this.contentDb) && !this.contentDb.WebApplication.Id.Equals(this.webApp.Id))
            {
                throw new SPCmdletException("The specified content database does not belong to the specified web application.");
            }
            string ownerLogin = OwnerLogin;
            string secondaryContactLogin = SecondaryOwnerLogin;
            string databaseName = null;
            bool createSiteInDB = false;
            if (ContentDatabase != null)
            {
                databaseName = contentDb.Name;
                createSiteInDB = true;
            }
            string quota = null;
            if (QuotaTemplate != null)
                quota = QuotaTemplate.Read().Name;

            SPSite site = Common.SiteCollections.ConvertSubSiteToSiteCollection.ConvertWebToSite(SourceWeb.WebUrl, TargetUrl, siteSubscription, SuppressAfterEvents.IsPresent,
                NoFileCompression.IsPresent, ExportedFile, createSiteInDB, databaseName, CreateManagedPath.IsPresent, 
                HaltOnWarning.IsPresent, HaltOnFatalError.IsPresent, DeleteSource.IsPresent, Name, Description, Language, 
                null, OwnerEmail, ownerLogin, null, secondaryContactLogin, SecondaryEmail, quota, useHostHeaderAsSiteName);

            return site;
        }


        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();

            if ((this.OwnerEmail != null) && !Utilities.ValidateEmail(this.OwnerEmail))
            {
                base.ThrowTerminatingError(new ErrorRecord(new ArgumentException("The Owner Email specified is invalid."), "", ErrorCategory.InvalidData, this));
            }
            if ((this.SecondaryEmail != null) && !Utilities.ValidateEmail(this.SecondaryEmail))
            {
                base.ThrowTerminatingError(new ErrorRecord(new ArgumentException("The Secondary Owner Email is invalid."), "", ErrorCategory.InvalidData, this));
            }
            base.DisposeOutputObjects = true;
        }

 

        protected override void InternalValidate()
        {
            if (TargetUrl != null)
            {
                TargetUrl = TargetUrl.Trim();
                siteUri = new Uri(TargetUrl, UriKind.Absolute);
                if (!Uri.UriSchemeHttps.Equals(siteUri.Scheme, StringComparison.OrdinalIgnoreCase) && !Uri.UriSchemeHttp.Equals(siteUri.Scheme, StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("The specified target URL is not valid.");
                }
            }
            string serverRelUrlFromFullUrl = Utilities.GetServerRelUrlFromFullUrl(TargetUrl);
            string siteRoot = null;
            if (this.HostHeaderWebApplication == null)
            {
                webApp = new SPWebApplicationPipeBind(TargetUrl).Read(false);
                siteRoot = Utilities.FindSiteRoot(webApp.Prefixes, serverRelUrlFromFullUrl);
                if ((siteRoot == null) || !siteRoot.Equals(serverRelUrlFromFullUrl, StringComparison.OrdinalIgnoreCase))
                {
                    throw new SPCmdletException("A managed path for the site could not be found.");
                }
            }
            else
            {
                webApp = this.HostHeaderWebApplication.Read();
                useHostHeaderAsSiteName = true;
                SPWebService service = SPFarm.Local.Services.GetValue<SPWebService>();
                if (service == null)
                {
                    throw new InvalidOperationException("A default web service could not be found.");
                }
                siteRoot = Utilities.FindSiteRoot(service.HostHeaderPrefixes, serverRelUrlFromFullUrl);
                if ((siteRoot == null) || !siteRoot.Equals(serverRelUrlFromFullUrl, StringComparison.OrdinalIgnoreCase))
                {
                    throw new SPCmdletException("A managed path for the site could not be found.");
                }
            }
            if (this.ContentDatabase != null)
            {
                this.contentDb = this.ContentDatabase.Read();
                if (null == this.contentDb)
                {
                    throw new SPException("The specified content database could not be found.");
                }
            }
            if (this.QuotaTemplate != null)
            {
                quotaTemplate = this.QuotaTemplate.Read().Name;
            }
            if (this.SiteSubscription != null)
            {
                this.siteSubscription = this.SiteSubscription.Read();
                if (this.siteSubscription == null)
                {
                    base.WriteError(new ArgumentException("The provided site subscription object is invalid."), ErrorCategory.InvalidArgument, this);
                    base.SkipProcessCurrentRecord();
                }
            }
        }


    }
}
