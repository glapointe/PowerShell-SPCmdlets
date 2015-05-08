using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Publishing;
using Microsoft.SharePoint.Utilities;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Themes
{
    [Cmdlet(VerbsCommon.Reset, "SPTheme", DefaultParameterSetName = "SPSite"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Themes")]
    [CmdletDescription("Resets a theme by applying all user specified theme configuration settings to the original source files. This is particularly helpful when the original source files have changed to a Feature upgrade.")]
    [RelatedCmdlets(ExternalCmdlets = new[] { "Get-SPWeb", "Get-SPSite"})]
    [Example(Code = "PS C:\\> Get-SPWeb http://server_name/sub-web | Reset-SPTheme",
        Remarks = "This example resets the theme for the web http://server_name/sub-web.")]
    [Example(Code = "PS C:\\> Get-SPSite http://server_name | Reset-SPTheme -SetSubWebsToInherit",
        Remarks = "This example resets the theme for the site collection http://server_name and resets all child webs to inherit from the root web.")]
    public class SPCmdletResetTheme : SPCmdletCustom
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the theme to reset.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        /// <summary>
        /// Gets or sets the site.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPSite",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The site containing the theme(s) to reset. If specified, all webs that have a unique theme will be reset.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [ValidateNotNull]
        public SPSitePipeBind Site { get; set; }


        /// <summary>
        /// Gets or sets whether to set all sub-webs to inherit.
        /// </summary>
        [Parameter(HelpMessage = "If specified, all child webs will be reset to inherit the theme of the specified web or root web.")]
        public SwitchParameter SetSubWebsToInherit { get; set; }


        protected override void InternalProcessRecord()
        {

            switch (ParameterSetName)
            {
                case "SPWeb":
                    SPWeb web = Web.Read();
                    if (web == null)
                        throw new FileNotFoundException(string.Format("The specified web could not be found. {0}", Web.WebUrl), Web.WebUrl);

                    try
                    {
                        ResetSPTheme(web, SetSubWebsToInherit);
                    }
                    finally
                    {
                        web.Dispose();
                        web.Site.Dispose();
                    }
                    break;
                case "SPSite":
                    SPSite site = Site.Read();
                    if (site == null)
                        throw new FileNotFoundException(string.Format("The specified site collection could not be found. {0}", Site.SiteUrl), Site.SiteUrl);

                    try
                    {
                        ResetSPTheme(site.RootWeb, SetSubWebsToInherit);
                        if (!SetSubWebsToInherit)
                        {
                            foreach (SPWeb web2 in site.AllWebs)
                            {
                                if (!web2.IsRootWeb)
                                {
                                    string themesUrlForWeb = ThmxTheme.GetThemeUrlForWeb(web2);
                                    string themesUrlForParentWeb = ThmxTheme.GetThemeUrlForWeb(web2.ParentWeb);
                                    if (!string.IsNullOrEmpty(themesUrlForWeb) && !string.IsNullOrEmpty(themesUrlForParentWeb) && themesUrlForWeb.ToLower() != themesUrlForParentWeb.ToLower())
                                    {
                                        // We found a web that does not have it's theme inheriting from the parent so reset it against the source files.
                                        ResetSPTheme(web2, false);
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        site.Dispose();
                    }
                    break;
            }
        }


        private void ResetSPTheme(SPWeb web, bool resetChildren)
        {
            try
            {
                // Store some variables for later use
                string tempFolderName = Guid.NewGuid().ToString("N").ToUpper();
                string themedFolderName = SPUrlUtility.CombineUrl(web.Site.ServerRelativeUrl, "/_catalogs/theme/Themed");
                string themesUrlForWeb = ThmxTheme.GetThemeUrlForWeb(web);

                if (string.IsNullOrEmpty(themesUrlForWeb))
                {
                    Logger.WriteWarning("The web {0} does not have a theme set and will be ignored.", web.Url);
                    return;
                }
                if (!web.IsRootWeb)
                {
                    string themesUrlForParentWeb = ThmxTheme.GetThemeUrlForWeb(web.ParentWeb);
                    if (themesUrlForWeb.ToLower() == themesUrlForParentWeb.ToLower())
                    {
                        Logger.WriteWarning("The web {0} inherits it's theme from it's parent. The theme will not be reset.", web.Url);
                        return;
                    }
                }

                // Open the existing theme
                ThmxTheme webThmxTheme = ThmxTheme.Open(web.Site, themesUrlForWeb);

                // Generate a new theme using the settings defined for the existing theme
                // (this will generate a temporary theme folder that we'll need to delete)
                webThmxTheme.GenerateThemedStyles(true, web.Site.RootWeb, tempFolderName);

                // Apply the newly generated theme to the web
                ThmxTheme.SetThemeUrlForWeb(web, SPUrlUtility.CombineUrl(themedFolderName, tempFolderName) + "/theme.thmx", true);

                // Delete the temp folder created earlier
                web.Site.RootWeb.GetFolder(SPUrlUtility.CombineUrl(themedFolderName, tempFolderName)).Delete();

                // Reset the theme URL just in case it has changed (sometimes it will)
                using (SPSite site = new SPSite(web.Site.ID))
                using (SPWeb web2 = site.OpenWeb(web.ID))
                {
                    string updatedThemesUrlForWeb = ThmxTheme.GetThemeUrlForWeb(web2);
                    if (resetChildren)
                    {
                        bool isPublishingWeb = false;
#if MOSS
                        isPublishingWeb = PublishingWeb.IsPublishingWeb(web2);
#endif
                        // Set all child webs to inherit.
                        if (isPublishingWeb)
                        {
#if MOSS
                            PublishingWeb pubWeb = PublishingWeb.GetPublishingWeb(web2);
                            pubWeb.ThemedCssFolderUrl.SetValue(pubWeb.Web.ThemedCssFolderUrl, true);
#endif
                        }
                        else
                        {
                            ResetChildWebs(web2, updatedThemesUrlForWeb);
                        }
                    }
                    else
                    {
                        if (themesUrlForWeb != updatedThemesUrlForWeb)
                        {
                            UpdateInheritingWebs(web2, themesUrlForWeb, updatedThemesUrlForWeb);
                        }
                    }
                }
            }
            finally
            {
                if (web != null)
                {
                    web.Site.Dispose();
                }
            }
        }
        private void UpdateInheritingWebs(SPWeb web, string oldThemesUrlForWeb, string newThemesUrlForWeb)
        {
            foreach (SPWeb childWeb in web.Webs)
            {
                string themesUrlForWeb = ThmxTheme.GetThemeUrlForWeb(childWeb);
                if (!string.IsNullOrEmpty(themesUrlForWeb) && themesUrlForWeb.ToLower() == oldThemesUrlForWeb.ToLower())
                    ThmxTheme.SetThemeUrlForWeb(childWeb, newThemesUrlForWeb);
                if (childWeb.Webs.Count > 0)
                    UpdateInheritingWebs(childWeb, oldThemesUrlForWeb, newThemesUrlForWeb);
                childWeb.Dispose();
            }
        }

        private void ResetChildWebs(SPWeb web, string themesUrlForWeb)
        {
            foreach (SPWeb childWeb in web.Webs)
            {
                ThmxTheme.SetThemeUrlForWeb(childWeb, themesUrlForWeb);
                if (childWeb.Webs.Count > 0)
                    ResetChildWebs(childWeb, themesUrlForWeb);
                childWeb.Dispose();
            }
        }


    }
}
