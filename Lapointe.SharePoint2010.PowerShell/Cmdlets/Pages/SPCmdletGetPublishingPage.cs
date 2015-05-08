using System;
using System.Collections.Generic;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint.Publishing;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Pages
{
    [Cmdlet(VerbsCommon.Get, "SPPublishingPage", SupportsShouldProcess = false, DefaultParameterSetName = "SPFile"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Pages")]
    [CmdletDescription("Retrieves all publishing pages from the specified source.")]
    [RelatedCmdlets(typeof(SPCmdletGetPublishingPageLayout), typeof(SPCmdletRepairPageLayoutUrl), ExternalCmdlets = new[] {"Get-SPWeb"})]
    [Example(Code = "PS C:\\> $pages = Get-SPWeb http://server_name | Get-SPPublishingPage",
        Remarks = "This example returns back all publishing pages in the http://server_name web.")]
    [Example(Code = "PS C:\\> $page = Get-SPPublishingPage \"http://server_name/pages/default.aspx\"",
        Remarks = "This example returns back the default.aspx publishing pages from the http://server_name web.")]
    [Example(Code = "PS C:\\> $page = $Get-SPWeb http://server_name | Get-SPPublishingPage -PageName \"default.aspx\"",
        Remarks = "This example returns back the default.aspx publishing pages from the http://server_name web.")]
    public class SPCmdletGetPublishingPage : SPGetCmdletBaseCustom<PublishingPage>
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the URL or GUID of the Web to retrieve publishing pages from.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind[] Web { get; set; }

        /// <summary>
        /// Gets or sets the page name.
        /// </summary>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = false,
            HelpMessage = "The name of the publishing page to return. Example: default.aspx")]
        [ValidateNotNullOrEmpty]
        public string[] PageName { get; set; }

        /// <summary>
        /// Gets or sets the url to the publishing page.
        /// </summary>
        /// <value>The name of the contentType.</value>
        [Parameter(ParameterSetName = "SPFile", 
            Mandatory = true, 
            Position = 0, 
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The path to the publishing page to return."),
        Alias("File")]
        [ValidateNotNull]
        public SPFilePipeBind[] Identity { get; set; }

        protected override IEnumerable<PublishingPage> RetrieveDataObjects()
        {
            List<PublishingPage> publishingPages = new List<PublishingPage>();

            switch (ParameterSetName)
            {
                case "SPWeb":
                    foreach (SPWebPipeBind webPipe in Web)
                    {
                        SPWeb web = webPipe.Read();

                        WriteVerbose("Getting publishing page from " + web.Url);
                        if (!PublishingWeb.IsPublishingWeb(web))
                        {
                            WriteWarning(string.Format("Web \"{0}\" is not a publishing web and will be skipped.", web.Url));
                            continue;
                        }
                        PublishingWeb pubWeb = PublishingWeb.GetPublishingWeb(web);
                        

                        if (PageName == null || PageName.Length == 0)
                        {

                            foreach (PublishingPage page in pubWeb.GetPublishingPages())
                                publishingPages.Add(page);
                        }
                        else
                        {
                            foreach (string pageName in PageName)
                            {
                                string pageUrl = string.Format("{0}/{1}/{2}", web.Url.TrimEnd('/'), pubWeb.PagesListName, pageName.Trim('/'));

                                try
                                {
                                    PublishingPage page = pubWeb.GetPublishingPage(pageUrl);
                                    if (page != null)
                                    {
                                        publishingPages.Add(page);
                                    }
                                    else
                                        WriteWarning("Could not locate the specified page: " + pageUrl);
                                }
                                catch (ArgumentException)
                                {
                                    WriteWarning("Could not locate the specified page: " + pageUrl);
                                }
                           }
                        }
                    }
                    break;
                case "SPFile":
                    foreach (SPFilePipeBind filePipe in Identity)
                    {
                        SPFile file = filePipe.Read();
                        if (!PublishingWeb.IsPublishingWeb(file.Web))
                        {
                            WriteWarning(string.Format("Web \"{0}\" is not a publishing web and will be skipped.", file.Web.Url));
                            continue;
                        }

                        WriteVerbose("Getting publishing page from " + file.Url);
                        try
                        {
                            PublishingPage page = null;
                            if (file.Exists)
                                page = PublishingPage.GetPublishingPage(file.Item);

                            if (page != null)
                                publishingPages.Add(page);
                            else
                                WriteWarning("Could not locate the specified page: " + file.Url);
                        }
                        catch (ArgumentException)
                        {
                            WriteWarning("Could not locate the specified page: " + file.Url);
                        }
                    }
                    break;
            }

            foreach (PublishingPage page in publishingPages)
            {
                AssignmentCollection.Add(page.PublishingWeb.Web);
                AssignmentCollection.Add(page.PublishingWeb.Web.Site);
                WriteResult(page);
            }

            return null;
        }
    }
}
