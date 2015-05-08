using System;
using System.Globalization;
using System.IO;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing.WebControls;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.WebPartPages;
using System.Web;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;

namespace Lapointe.SharePoint.PowerShell.Common.WebParts
{
    public class RetargetContentQueryWebPart
    {
        public static void Retarget(string url, bool allMatching, string webPartId, string webPartTitle, string listUrl, string listType, string siteUrl, bool publish)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb()) // The url contains a filename so AllWebs[] will not work unless we want to try and parse which we don't
            {
                SPFile file = web.GetFile(url);

                // file.Item will throw "The object specified does not belong to a list." if the url passed
                // does not correspond to a file in a list.

                bool wasCheckedOut = true;
                if (file.CheckOutStatus == SPFile.SPCheckOutStatus.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                else if (!Utilities.IsCheckedOutByCurrentUser(file.Item))
                {
                    // We don't want to mess with files that are checked out to other users so exit out.
                    throw new SPException("The file is currently checked out by another user.");
                }

                string displayTitle = string.Empty;
                bool modified = false;
                SPLimitedWebPartManager manager = null;
                try
                {
                    if (!allMatching)
                    {
                        bool cleanupContext = false;
                        if (HttpContext.Current == null)
                        {
                            HttpRequest request = new HttpRequest("", web.Url, "");
                            HttpContext.Current = new HttpContext(request, new HttpResponse(new StringWriter()));
                            SPControl.SetContextWeb(HttpContext.Current, web);
                            cleanupContext = true;
                        }

                        WebPart wp = null;
                        try
                        {
                            if (!string.IsNullOrEmpty(webPartId))
                            {
                                wp = Utilities.GetWebPartById(web, url, webPartId, out manager);
                            }
                            else
                            {
                                wp = Utilities.GetWebPartByTitle(web, url, webPartTitle, out manager);
                                if (wp == null)
                                {
                                    throw new SPException(
                                        "Unable to find specified web part. Try specifying the -ID parameter instead (use enumpagewebparts to get the ID) or use -AllMatching to adjust all web parts that match the specified title.");
                                }
                            }
                            if (wp == null)
                            {
                                throw new SPException("Unable to find specified web part.");
                            }

                            AdjustWebPart(web, wp, manager, listUrl, listType, siteUrl);
                            modified = true;

                            // Set this so that we can add it to the check-in comment.
                            displayTitle = wp.DisplayTitle;
                        }
                        finally
                        {
                            if (cleanupContext)
                                HttpContext.Current = null;

                            if (wp != null)
                                wp.Dispose();
                        }

                    }
                    else
                    {
                        manager = web.GetLimitedWebPartManager(url, PersonalizationScope.Shared);
                        foreach (WebPart tempWP in manager.WebParts)
                        {
                            try
                            {
                                if (!(tempWP is ContentByQueryWebPart))
                                    continue;

                                if (tempWP.DisplayTitle.ToLowerInvariant() == webPartTitle.ToLowerInvariant())
                                {
                                    AdjustWebPart(web, tempWP, manager, listUrl, listType, siteUrl);
                                    displayTitle = tempWP.DisplayTitle;
                                    modified = true;
                                }
                            }
                            finally
                            {
                                tempWP.Dispose();
                            }
                        }

                    }

                }
                finally
                {
                    if (manager != null)
                    {
                        manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                        manager.Dispose();
                    }

                    if (modified)
                        file.CheckIn("Checking in changes to page layout due to retargeting of content query web part " + displayTitle);
                    else if (!wasCheckedOut)
                        file.UndoCheckOut();

                    if (modified && publish)
                    {
                        file.Publish("Publishing changes to page layout due to retargeting of content query web part " + displayTitle);
                        if (file.Item.ModerationInformation != null)
                        {
                            file.Approve("Approving changes to page layout due to retargeting of content query web part " + displayTitle);
                        }
                    }
                }
            }
        }



        /// <summary>
        /// Adjusts the web part.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="wp">The web part.</param>
        /// <param name="manager">The web part manager.</param>
        /// <param name="listUrl">The list URL.</param>
        /// <param name="listType">Type of the list (list template).</param>
        /// <param name="siteUrl">The site URL.</param>
        internal static void AdjustWebPart(SPWeb web, WebPart wp, SPLimitedWebPartManager manager, string listUrl, string listType, string siteUrl)
        {
            ContentByQueryWebPart cqwp = wp as ContentByQueryWebPart;
            if (cqwp == null)
                throw new SPException("Web part is not a Content Query web part.");

            if (listUrl != null)
            {
                SPList list = Utilities.GetListFromViewUrl(listUrl);
                if (list == null)
                    throw new SPException("Specified List was not found.");

                AdjustWebPart(cqwp, list, web);
            }
            else if (siteUrl != null)
            {
                try
                {
                    using (SPWeb web2 = web.Site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(siteUrl)])
                    {
                        if (!web2.Exists)
                            throw new SPException(string.Format("The web {0} does not exist.", siteUrl));
                    }
                }
                catch (ArgumentException)
                {
                    throw new SPException(siteUrl + " either does not exist is is invalid or does not belong to the web part's container site collection.");
                }
                catch (SPException)
                {
                    throw new SPException(siteUrl + " either does not exist is is invalid or does not belong to the web part's container site collection.");
                }
                catch (FileNotFoundException)
                {
                    throw new SPException(siteUrl + " either does not exist is is invalid or does not belong to the web part's container site collection.");
                }

                cqwp.WebUrl = siteUrl;
                cqwp.ListGuid = string.Empty;
                cqwp.ListName = string.Empty;
            }
            else
            {
                cqwp.WebUrl = string.Empty;
                cqwp.ListGuid = string.Empty;
            }

            if (listType != null)
                ApplyListTypeChanges(web, cqwp, listType);


            manager.SaveChanges(cqwp);
        }

        /// <summary>
        /// Adjusts the web part.
        /// </summary>
        /// <param name="cqwp">The CQWP.</param>
        /// <param name="list">The list.</param>
        /// <param name="web">The web.</param>
        internal static void AdjustWebPart(ContentByQueryWebPart cqwp, SPList list, SPWeb web)
        {
            if (list.ContentTypes["Listing"] != null)
            {
                // The list is a special list - it was upgraded from v2 and corresponds 
                // to the grouped listings web part so we need to set some additional
                // properties that cannot be set via the browser.

                ApplyListTypeChanges(web, cqwp, "Links");

                cqwp.AdditionalGroupAndSortFields = "Modified,Modified;Created,Created";
                cqwp.DataColumnRenames = "SummaryTitle,Title;Comments,Description;URL,LinkUrl;SummaryImage,ImageUrl";
                cqwp.SortByFieldType = "Number";
                cqwp.ChromeType = PartChromeType.None;
                cqwp.CommonViewFields = "SummaryTitle,Text;Comments,Note;URL,URL;SummaryImage,URL;SummaryIcon,URL;SummaryType,Integer;_TargetItemID,Note";
                cqwp.FilterType1 = "ModStat";
                cqwp.Xsl = "<xsl:stylesheet xmlns:x=\"http://www.w3.org/2001/XMLSchema\" version=\"1.0\" xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\" xmlns:cmswrt=\"http://schemas.microsoft.com/WebPart/v3/Publishing/runtime\" exclude-result-prefixes=\"xsl cmswrt x\" > <xsl:import href=\"/Style Library/XSL Style Sheets/Header.xsl\" /> <xsl:import href=\"/Style Library/XSL Style Sheets/ItemStyle.xsl\" /> <xsl:import href=\"/Style Library/XSL Style Sheets/ContentQueryMain.xsl\" /> </xsl:stylesheet>";
                cqwp.FilterValue1 = "Approved";
                cqwp.ShowUntargetedItems = false;
                cqwp.FilterField1 = list.Fields["Approval Status"].Id.ToString();
                cqwp.Filter1ChainingOperator = ContentByQueryWebPart.FilterChainingOperator.And;
                cqwp.ItemLimit = -1;
                cqwp.SortBy = list.Fields["Order"].Id.ToString();
                cqwp.SortByDirection = ContentByQueryWebPart.SortDirection.Asc;
                cqwp.GroupByDirection = ContentByQueryWebPart.SortDirection.Asc;
                cqwp.Description = "Show listings in the portal.";
                cqwp.GroupBy = "SummaryGroup";
                cqwp.Hidden = false;
            }
            using (SPWeb parentWeb = list.ParentWeb)
                cqwp.WebUrl = parentWeb.ServerRelativeUrl;
            cqwp.ListGuid = list.ID.ToString();
            cqwp.ListName = list.Title;
        }

        /// <summary>
        /// Applies the list type changes.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="cqwp">The content query web part.</param>
        /// <param name="listType">Type of the list.</param>
        private static void ApplyListTypeChanges(SPWeb web, ContentByQueryWebPart cqwp, string listType)
        {
            if (listType == null)
                return;

            using (SPWeb rootWeb = web.Site.RootWeb)
            {
                SPListTemplateCollection listTemplates = rootWeb.ListTemplates;

                SPListTemplate template = listTemplates[listType];
                if (template == null)
                    throw new SPException("List template (type) not found.");

                cqwp.BaseType = string.Empty;
                cqwp.ServerTemplate = Convert.ToString((int)template.Type, CultureInfo.InvariantCulture);

                bool isGenericList = template.BaseType == SPBaseType.GenericList;
                bool isIssueList = template.BaseType == SPBaseType.Issue;
                bool isLinkList = template.Type == SPListTemplateType.Links;
                cqwp.UseCopyUtil = !isGenericList ? isIssueList : (isLinkList ? false : true);
            }
        }
    }
}
