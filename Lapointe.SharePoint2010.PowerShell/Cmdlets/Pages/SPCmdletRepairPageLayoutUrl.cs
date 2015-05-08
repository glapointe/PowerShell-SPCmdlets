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

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Pages
{
    [Cmdlet("Repair", "SPPageLayoutUrl", SupportsShouldProcess = true),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Pages")]
    [CmdletDescription("Fixes the Page Layout URL property of publishing pages which can get messed up during an upgrade or from importing into a new farm.")]
    [RelatedCmdlets(typeof(SPCmdletGetPublishingPage), typeof(SPCmdletGetPublishingPageLayout),
        ExternalCmdlets = new[] {"Get-SPWebApplication", "Get-SPSite", "Get-SPWeb"})]
    [Example(Code = "PS C:\\> Get-SPWebApplication \"http://server_name\" | Repair-SPPageLayoutUrl",
        Remarks = "This example automatically repairs all pages in http://server_name by analyzing the page layout name and fixing the path so it points to the appropriate location.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name\" | Repair-SPPageLayoutUrl -RegexSearchString \"(?:http://testserver/)\" -RegexReplaceString \"http://prodserver\"",
        Remarks = "This example automatically repairs all pages in http://server_name replacing the value http://testserver with http://prodserver.")]
    public class SPCmdletRepairPageLayoutUrl : SPCmdletCustom
    {
        [Parameter(Mandatory = false,
            HelpMessage = "Search pattern to use for a regular expression replacement of the page layout URL.")]
        public string RegexSearchString { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Replace pattern to use for a regular expression replacement of the page layout URL.")]
        public string RegexReplaceString { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Attempt to fix any page contacts that are invalid.")]
        public SwitchParameter FixContact { get; set; }

        [Parameter(Mandatory = true, 
            ParameterSetName = "WebApplication", 
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The web application containing the pages to repair.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        public SPWebApplicationPipeBind WebApplication { get; set; }

        [Parameter(Mandatory = true, 
            ParameterSetName = "Site",
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The site containing the pages to repair.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        [Parameter(Mandatory = true, 
            ParameterSetName = "Web",
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the pages to repair.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true, 
            ParameterSetName = "Page",
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The URL to the page to repair.")]
        public SPFilePipeBind Page { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "URL of page layout to retarget page(s) to.")]
        public SPPageLayoutPipeBind PageLayout { get; set; }

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

            Regex regex = null;
            string replaceString = null;
            if (!string.IsNullOrEmpty(RegexSearchString))
            {
                regex = new Regex(RegexSearchString);
                replaceString = RegexReplaceString;
            }


            string pageName = null;
            SPFile page = null;

            try
            {
                if (Page != null)
                {
                    page = Page.Read();
                    pageName = page.Name;
                }
            }
            catch
            {
                if (page != null)
                {
                    page.Web.Dispose();
                    page.Web.Site.Dispose();
                }
                throw;
            }

            try
            {
                if (WebApplication != null)
                {
                    SPWebApplication webApp = WebApplication.Read();
                    Logger.Write("Progress: Begin processing web application '{0}'.", webApp.GetResponseUri(SPUrlZone.Default).ToString());
                    foreach (SPSite site in webApp.Sites)
                    {
                        Logger.Write("Progress: Begin processing site '{0}'.", site.ServerRelativeUrl);
                        try
                        {
                            foreach (SPWeb web in site.AllWebs)
                            {
                                Logger.Write("Progress: Begin processing web '{0}'.", web.ServerRelativeUrl);

                                try
                                {
                                    PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                                    Common.Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb, pageName, GetPageLayout(pubweb, page), regex, replaceString, FixContact.IsPresent, test);
                                }
                                finally
                                {
                                    Logger.Write("Progress: Finished processing web '{0}'.", web.ServerRelativeUrl);

                                    web.Dispose();
                                }
                            }
                        }
                        finally
                        {
                            Logger.Write("Progress: Finished processing site '{0}'.", site.ServerRelativeUrl);
                            site.Dispose();
                        }
                    }
                    Logger.Write("Progress: Finished processing web application '{0}'.", webApp.GetResponseUri(SPUrlZone.Default).ToString());
                }
                else if (Site != null)
                {
                    using (SPSite site = new SPSite(Site.SiteGuid))
                    {
                        Logger.Write("Progress: Begin processing site '{0}'.", site.ServerRelativeUrl);

                        foreach (SPWeb web in site.AllWebs)
                        {
                            Logger.Write("Progress: Begin processing web '{0}'.", web.ServerRelativeUrl);

                            try
                            {
                                PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                                Common.Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb, pageName, GetPageLayout(pubweb, page), regex, replaceString, FixContact.IsPresent, test);
                            }
                            finally
                            {
                                Logger.Write("Progress: Finished processing web '{0}'.", web.ServerRelativeUrl);

                                web.Dispose();
                            }
                        }
                        Logger.Write("Progress: Finished processing site '{0}'.", site.ServerRelativeUrl);
                    }
                }
                else if (Web != null)
                {
                    using (SPWeb web = Web.Read())
                    {
                        Logger.Write("Progress: Begin processing web '{0}'.", web.Url);

                        PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                        Common.Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb, pageName, GetPageLayout(pubweb, page), regex, replaceString, FixContact.IsPresent, test);

                        Logger.Write("Progress: Finished processing web '{0}'.", web.Url);
                    }
                }
                else if (Page != null)
                {
                    try
                    {
                        PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(page.Web);

                        Common.Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb, pageName, GetPageLayout(pubweb, page), regex, replaceString, FixContact.IsPresent, test);
                    }
                    finally
                    {
                        page.Web.Dispose();
                        page.Web.Site.Dispose();
                    }
                }
            }
            finally
            {
                if (page != null)
                {
                    page.Web.Dispose();
                    page.Web.Site.Dispose();
                }
            }
        }

        protected override void InternalValidate()
        {
            base.InternalValidate();

            if (PageLayout != null)
            {
                if (WebApplication != null)
                    throw new SPCmdletException("The PageLayout parameter is incompatible with the WebApplication parameter");

                if (!string.IsNullOrEmpty(RegexSearchString) || !string.IsNullOrEmpty(RegexReplaceString))
                    throw new SPCmdletException("The PageLayout parameter is incompatible with the RegexSearchString and RegexReplaceString parameters.");
            }
            if (!string.IsNullOrEmpty(RegexSearchString) || !string.IsNullOrEmpty(RegexReplaceString))
            {
                if (string.IsNullOrEmpty(RegexSearchString) || string.IsNullOrEmpty(RegexReplaceString))
                    throw new SPCmdletException("Both the search and replace strings are required if either is provided.");
            }
        }

        private string GetPageLayout(PublishingWeb web, SPFile page)
        {
            PageLayout pageLayout = null;
            try
            {
                if (PageLayout != null)
                {
                    if (web == null)
                        pageLayout = PageLayout.Read();
                    else
                    {
                        pageLayout = PageLayout.Read(web);
                    }

                    if (page != null && pageLayout != null)
                    {
                        if (pageLayout.ListItem.Web.Site.ID != page.Web.Site.ID)
                        {
                            throw new SPException("The specified page layout and page are not in the same site collection.");
                        }
                    }
                }
                if (pageLayout == null)
                    return null;

                string pageLayoutValue = pageLayout.ListItem.Web.Site.MakeFullUrl(pageLayout.ServerRelativeUrl) + ", " + pageLayout.Title;
                return pageLayoutValue;
            }
            catch
            {
                if (web.IsRoot) throw;
                return GetPageLayout(web.ParentPublishingWeb, page);
            }
            finally
            {
                if (pageLayout != null)
                {
                    pageLayout.ListItem.Web.Dispose();
                    pageLayout.ListItem.Web.Site.Dispose();
                }
            }

        }

    }
}
