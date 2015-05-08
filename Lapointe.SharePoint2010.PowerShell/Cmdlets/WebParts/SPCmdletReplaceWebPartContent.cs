using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using Microsoft.SharePoint.Deployment;
using System.IO;
using Microsoft.SharePoint.Administration.Backup;
using System.Collections;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet("Replace", "SPWebPartContent", SupportsShouldProcess = true, DefaultParameterSetName="Site"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Replaces all occurances of the search string with the replacement string. Supports the use of regular expressions. Use -WhatIf to verify your replacements before executing.")]
    [RelatedCmdlets(typeof(SPCmdletGetLimitedWebPartManager), typeof(SPCmdletGetWebPartList), typeof(Pages.SPCmdletGetPublishingPage), typeof(Lists.SPCmdletGetFile),
        ExternalCmdlets = new[] {"Get-SPFarm", "Get-SPWebApplication", "Get-SPSite", "Get-SPWeb"})]
    [Example(Code = "PS C:\\> Get-SPWeb http://portal | Replace-SPWebPartContent -SearchString \"(?i:old product name)\" -ReplaceString \"New Product Name\" -Publish",
        Remarks = "This example does a case-insensitive search for \"old product name\" and replaces with \"New Product Name\" and publishes the changes after completion.")]
    public class SPCmdletReplaceWebPartContent : SPCmdletCustom
    {

        [Parameter(Mandatory = true,
            HelpMessage = "A regular expression search string.")]
        public string SearchString { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The string to replace the match with.")]
        public string ReplaceString { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The name of the web part to update within the scope.")]
        public string WebPartName { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Publish or check-in the file after updating the contents.")]
        public SwitchParameter Publish { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Perform a replace on XML fields without iterating through the elements of the XML (do a straight search and replace on the XML string without going through the DOM)")]
        public SwitchParameter UnsafeXml { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The log file to store all change records to.")]
        [ValidateDirectoryExistsAndValidFileName]
        public string LogFile { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Page",
            Position = 0,
            HelpMessage = "The URL to the page containing the web parts whose content will be replaced.")]
        public SPFilePipeBind Page { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Web",
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the web parts whose content will be replaced.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Site",
            Position = 0,
            HelpMessage = "The site containing the web parts whose content will be replaced.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "WebApplication",
            Position = 0,
            HelpMessage = "The web application containing the web parts whose content will be replaced.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        public SPWebApplicationPipeBind WebApplication { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Farm",
            Position = 0,
            HelpMessage = "Provide the SPFarm object to replace matching content in all web parts throughout the farm.")]
        public SPFarmPipeBind Farm { get; set; }

        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();

            Logger.LogFile = LogFile;
        }

        protected override void InternalProcessRecord()
        {
            bool test = false;
            ShouldProcessReason reason;
            if (!base.ShouldProcess(null, null, null, out reason))
            {
                if (reason == ShouldProcessReason.WhatIf)
                {
                    test = true;
                }
            }
            if (test)
                Logger.Verbose = true;

            Common.WebParts.ReplaceWebPartContent.Settings settings = new Common.WebParts.ReplaceWebPartContent.Settings();
            settings.SearchString = SearchString;
            settings.ReplaceString = ReplaceString;
            settings.WebPartName = WebPartName;
            settings.Publish = Publish.IsPresent;
            settings.Test = test;
            settings.UnsafeXml = UnsafeXml;

            switch (ParameterSetName)
            {
                case "WebApplication":
                    SPWebApplication webApp1 = WebApplication.Read();
                    if (webApp1 == null)
                        throw new SPException("Web Application not found.");
                    Common.WebParts.ReplaceWebPartContent.ReplaceValues(webApp1, settings);
                    break;
                case "Site":
                    using (SPSite site = Site.Read())
                    {
                        Common.WebParts.ReplaceWebPartContent.ReplaceValues(site, settings);
                    }
                    break;
                case "Web":
                    using (SPWeb web = Web.Read())
                    {
                        try
                        {
                            Common.WebParts.ReplaceWebPartContent.ReplaceValues(web, settings);
                        }
                        finally
                        {
                            web.Site.Dispose();
                        }
                    }
                    break;
                case "Page":
                    SPFile file = Page.Read();
                    try
                    {
                        Common.WebParts.ReplaceWebPartContent.ReplaceValues(file.Web, file, settings);
                    }
                    finally
                    {
                        file.Web.Dispose();
                        file.Web.Site.Dispose();
                    }
                    break;
                default:
                    SPFarm farm = Farm.Read();
                    foreach (SPService svc in farm.Services)
                    {
                        if (!(svc is SPWebService))
                            continue;

                        foreach (SPWebApplication webApp2 in ((SPWebService)svc).WebApplications)
                        {
                            Common.WebParts.ReplaceWebPartContent.ReplaceValues(webApp2, settings);
                        }
                    }
                    break;
            }
        }

    }
}
