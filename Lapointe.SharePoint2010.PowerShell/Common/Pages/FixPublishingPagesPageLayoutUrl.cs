using System;
using System.Management.Automation;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Publishing;

namespace Lapointe.SharePoint.PowerShell.Common.Pages
{
    class FixPublishingPagesPageLayoutUrl
    {

        /// <summary>
        /// Fixes the pages page layout url so that it points to the page layout in the container site collections master page gallery.
        /// </summary>
        /// <param name="publishingWeb">The publishing web.</param>
        public static void FixPages(PublishingWeb publishingWeb)
        {
            FixPages(publishingWeb, null, null, null, null, false, false);
        }

        /// <summary>
        /// Fixes the pages page layout url so that it points to the page layout in the container site collections master page gallery.
        /// </summary>
        /// <param name="publishingWeb">The target publishing web.</param>
        /// <param name="pageName">Name of the page.</param>
        /// <param name="pageLayoutUrl">The page layout URL.</param>
        /// <param name="searchRegex">The search regex.</param>
        /// <param name="replaceString">The replace string.</param>
        /// <param name="fixContact">if set to <c>true</c> [fix contact].</param>
        /// <param name="test">if set to <c>true</c> [test].</param>
        public static void FixPages(PublishingWeb publishingWeb, string pageName, string pageLayoutUrl, Regex searchRegex, string replaceString, bool fixContact, bool test)
        {
            if (!PublishingWeb.IsPublishingWeb(publishingWeb.Web))
                return;

            PublishingPageCollection pages;
            int tryCount = 0;
            while (true)
            {
                try
                {
                    tryCount++;
                    pages = publishingWeb.GetPublishingPages();
                    break;
                }
                catch (InvalidPublishingWebException)
                {
                    // The following is meant to deal with a timing issue when using this method in conjuction with other commands.  When
                    // used independently this should be unnecessary.
                    if (tryCount > 4)
                        throw;
                    Thread.Sleep(10000);
                    SPWeb web = publishingWeb.Web;
                    SPSite site = web.Site;
                    string url = site.MakeFullUrl(web.ServerRelativeUrl);
                    site.Close();
                    site.Dispose();
                    web.Close();
                    web.Dispose();
                    publishingWeb.Close();
                    site = new SPSite(url);
                    web = site.OpenWeb(Utilities.GetServerRelUrlFromFullUrl(url));
                    publishingWeb = PublishingWeb.GetPublishingWeb(web);
                }
            }

            foreach (PublishingPage page in pages)
            {
                if (!(string.IsNullOrEmpty(pageName) || page.Name.ToLower() == pageName.ToLower()))
                    continue;

                if (page.ListItem[FieldId.PageLayout] == null)
                    continue;

                Logger.Write("Progress: Begin processing {0}.", page.Url);
                Logger.Write("Progress: Current layout set to {0}.", page.ListItem[FieldId.PageLayout].ToString());

                // Can't edit items that are checked out.
                if (Utilities.IsCheckedOut(page.ListItem) && !Utilities.IsCheckedOutByCurrentUser(page.ListItem))
                {
                    Logger.WriteWarning("WARNING: Page is already checked out by another user - skipping.");
                    continue;
                }

                SPFieldUrlValue url;
                if (string.IsNullOrEmpty(pageLayoutUrl))
                {
                    if (searchRegex == null)
                    {
                        if (page.ListItem[FieldId.PageLayout] == null || string.IsNullOrEmpty(page.ListItem[FieldId.PageLayout].ToString().Trim()))
                        {
                            Logger.WriteWarning("WARNING: Current page layout is empty - skipping.  Use the 'pagelayout' parameter to set a page layout.");

                            continue;
                        }

                        // We didn't get a layout url passed in or a regular expression so try and fix the existing url
                        url = new SPFieldUrlValue(page.ListItem[FieldId.PageLayout].ToString());
                        if (string.IsNullOrEmpty(url.Url) ||
                            url.Url.IndexOf("/_catalogs/") < 0)
                        {
                            Logger.WriteWarning("WARNING: Current page layout does not point to a _catalogs folder or is empty - skipping.  Use the 'pagelayout' parameter to set a page layout  Layout Url: {0}", url.ToString());
                            continue;
                        }


                        string newUrl = publishingWeb.Web.Site.ServerRelativeUrl.TrimEnd('/') +
                                      url.Url.Substring(url.Url.IndexOf("/_catalogs/"));

                        string newDesc = publishingWeb.Web.Site.MakeFullUrl(newUrl);

                        if (url.Url.ToLowerInvariant() == newUrl.ToLowerInvariant())
                        {
                            Logger.Write("Progress: Current layout matches new evaluated layout - skipping.");
                            continue;
                        }
                        url.Url = newUrl;
                        url.Description = newDesc;
                    }
                    else
                    {
                        if (page.ListItem[FieldId.PageLayout] == null || string.IsNullOrEmpty(page.ListItem[FieldId.PageLayout].ToString().Trim()))
                            Logger.Write("Progress: Current page layout is empty - skipping.  Use the pagelayout parameter to set a page layout.");

                        // A regular expression was passed in so use it to fix the page layout url if we find a match.
                        if (searchRegex.IsMatch((string)page.ListItem[FieldId.PageLayout]))
                        {
                            url = new SPFieldUrlValue(page.ListItem[FieldId.PageLayout].ToString());
                            string newUrl = searchRegex.Replace((string)page.ListItem[FieldId.PageLayout], replaceString);
                            if (url.ToString().ToLowerInvariant() == newUrl.ToLowerInvariant())
                            {
                                Logger.Write("Progress: Current layout matches new evaluated layout - skipping.");
                                continue;
                            }
                            url = new SPFieldUrlValue(newUrl);
                        }
                        else
                        {
                            Logger.Write("Progress: Existing page layout url does not match provided regular expression - skipping.");
                            continue;
                        }
                    }
                }
                else
                {
                    // The user passed in an url string so use it.
                    if (pageLayoutUrl.ToLowerInvariant() == (string)page.ListItem[FieldId.PageLayout])
                    {
                        Logger.Write("Progress: Current layout matches provided layout - skipping.");
                        continue;
                    }

                    url = new SPFieldUrlValue(pageLayoutUrl);
                }

                string fileName = url.Url.Substring(url.Url.LastIndexOf('/'));
                // Make sure that the URLs are server relative instead of absolute.
                //if (url.Description.ToLowerInvariant().StartsWith("http"))
                //    url.Description = Utilities.GetServerRelUrlFromFullUrl(url.Description) + fileName;
                //if (url.Url.ToLowerInvariant().StartsWith("http"))
                //    url.Url = Utilities.GetServerRelUrlFromFullUrl(url.Url) + fileName;

                if (page.ListItem[FieldId.PageLayout] != null && url.ToString().ToLowerInvariant() == page.ListItem[FieldId.PageLayout].ToString().ToLowerInvariant())
                    continue; // No difference detected so move on.

                Logger.Write("Progress: Changing layout url from \"{0}\" to \"{1}\"", page.ListItem[FieldId.PageLayout].ToString(), url.ToString());


                if (fixContact)
                {
                    SPUser contact = null;
                    try
                    {
                        contact = page.Contact;
                    }
                    catch (SPException)
                    {
                    }
                    if (contact == null)
                    {
                        Logger.Write("Progress: Page contact ('{0}') does not exist - assigning current user as contact.", page.ListItem[FieldId.Contact].ToString());
                        page.Contact = publishingWeb.Web.CurrentUser;

                        if (!test)
                            page.ListItem.SystemUpdate();
                    }
                }

                if (test)
                    continue;

                try
                {
                    bool publish = false;
                    if (!Utilities.IsCheckedOut(page.ListItem))
                    {
                        page.CheckOut();
                        publish = true;
                    }
                    page.ListItem[FieldId.PageLayout] = url;
                    page.ListItem.UpdateOverwriteVersion();

                    if (publish)
                    {
                        Common.Lists.PublishItems itemPublisher = new Common.Lists.PublishItems();
                        itemPublisher.PublishListItem(page.ListItem, page.ListItem.ParentList, false, "Automated fix of publishing pages page layout URL.", null, null);
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteException(new ErrorRecord(ex, null, ErrorCategory.NotSpecified, page));
                }
            }
        }
    }
}
