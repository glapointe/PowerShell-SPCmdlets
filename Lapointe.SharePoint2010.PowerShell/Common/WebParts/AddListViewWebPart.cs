using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebPartPages;
using Lapointe.SharePoint.PowerShell.Common.Pages;

namespace Lapointe.SharePoint.PowerShell.Common.WebParts
{
    class AddListViewWebPart
    {

        /// <summary>
        /// Adds a List View Web Part to the specified page.
        /// </summary>
        /// <param name="pageUrl">The page URL.</param>
        /// <param name="listUrl">The list URL.</param>
        /// <param name="title">The title.</param>
        /// <param name="viewTitle">Title of the view.</param>
        /// <param name="zoneId">The zone ID.</param>
        /// <param name="zoneIndex">Index within the zone.</param>
        /// <param name="linkTitle">if set to <c>true</c> [link title].</param>
        /// <param name="chromeType">Type of the chrome.</param>
        /// <param name="publish">if set to <c>true</c> [publish].</param>
        /// <returns></returns>
        public static Microsoft.SharePoint.WebPartPages.WebPart Add(string pageUrl, string listUrl, string title, string viewTitle, string zoneId, int zoneIndex, bool linkTitle, string jsLink, PartChromeType chromeType, bool publish)
        {
            using (SPSite site = new SPSite(pageUrl))
            using (SPWeb web = site.OpenWeb())
            // The url contains a filename so AllWebs[] will not work unless we want to try and parse which we don't
            {
                SPFile file = web.GetFile(pageUrl);

                // file.Item will throw "The object specified does not belong to a list." if the url passed
                // does not correspond to a file in a list.

                SPList list = Utilities.GetListFromViewUrl(listUrl);
                if (list == null)
                    throw new ArgumentException("List not found.");

                SPView view = null;
                if (!string.IsNullOrEmpty(viewTitle))
                {
                    view = list.Views.Cast<SPView>().FirstOrDefault(v => v.Title == viewTitle);
                    if (view == null)
                        throw new ArgumentException("The specified view was not found.");
                }

                bool checkBackIn = false;
                if (file.InDocumentLibrary)
                {
                    if (!Utilities.IsCheckedOut(file.Item) || !Utilities.IsCheckedOutByCurrentUser(file.Item))
                    {
                        checkBackIn = true;
                        file.CheckOut();
                    }
                    // If it's checked out by another user then this will throw an informative exception so let it do so.
                }
                string displayTitle = string.Empty;
                Microsoft.SharePoint.WebPartPages.WebPart lvw = null;
                
                SPLimitedWebPartManager manager = null;
                try
                {
                    manager = web.GetLimitedWebPartManager(pageUrl, PersonalizationScope.Shared);
                    lvw = new XsltListViewWebPart();
                    if (list.BaseTemplate == SPListTemplateType.Events)
                        lvw = new ListViewWebPart();

                    if (lvw is ListViewWebPart)
                    {
                        ((ListViewWebPart)lvw).ListName = list.ID.ToString("B").ToUpperInvariant();
                        ((ListViewWebPart)lvw).WebId = list.ParentWeb.ID;
                        if (view != null)
                            ((ListViewWebPart)lvw).ViewGuid = view.ID.ToString("B").ToUpperInvariant();
                    }
                    else
                    {
                        ((XsltListViewWebPart)lvw).ListName = list.ID.ToString("B").ToUpperInvariant();
#if !SP2010
                        if (!string.IsNullOrEmpty(jsLink))
                            ((XsltListViewWebPart)lvw).JSLink = jsLink;
#endif
                        ((XsltListViewWebPart)lvw).WebId = list.ParentWeb.ID;
                        if (view != null)
                            ((XsltListViewWebPart)lvw).ViewGuid = view.ID.ToString("B").ToUpperInvariant();
                    }

                    if (linkTitle)
                    {
                        if (view != null)
                            lvw.TitleUrl = view.Url;
                        else
                            lvw.TitleUrl = list.DefaultViewUrl;
                    }

                    if (!string.IsNullOrEmpty(title))
                        lvw.Title = title;

                    lvw.ChromeType = chromeType;

                    displayTitle = lvw.DisplayTitle;

                    manager.AddWebPart(lvw, zoneId, zoneIndex);
                }
                finally
                {
                    if (manager != null)
                    {
                        manager.Web.Dispose();
                        manager.Dispose();
                    }
                    if (lvw != null)
                        lvw.Dispose();

                    if (file.InDocumentLibrary && Utilities.IsCheckedOut(file.Item) && (checkBackIn || publish))
                        file.CheckIn("Checking in changes to page due to new web part being added: " + displayTitle);

                    if (publish && file.InDocumentLibrary)
                    {
                        file.Publish("Publishing changes to page due to new web part being added: " + displayTitle);
                        if (file.Item.ModerationInformation != null)
                        {
                            file.Approve("Approving changes to page due to new web part being added: " + displayTitle);
                        }
                    }
                }
                return lvw;
            }
        }

        public static Microsoft.SharePoint.WebPartPages.WebPart AddToWikiPage(string pageUrl, string listUrl, string title, string viewTitle, int row, int column, bool linkTitle, string jsLink, PartChromeType chromeType, bool addSpace, bool publish)
        {
            using (SPSite site = new SPSite(pageUrl))
            using (SPWeb web = site.OpenWeb())
            // The url contains a filename so AllWebs[] will not work unless we want to try and parse which we don't
            {
                SPFile file = web.GetFile(pageUrl);

                // file.Item will throw "The object specified does not belong to a list." if the url passed
                // does not correspond to a file in a list.

                SPList list = Utilities.GetListFromViewUrl(listUrl);
                if (list == null)
                    throw new ArgumentException("List not found.");

                SPView view = null;
                if (!string.IsNullOrEmpty(viewTitle))
                {
                    view = list.Views.Cast<SPView>().FirstOrDefault(v => v.Title == viewTitle);
                    if (view == null)
                        throw new ArgumentException("The specified view was not found.");
                }

                bool checkBackIn = false;
                if (file.InDocumentLibrary)
                {
                    if (!Utilities.IsCheckedOut(file.Item) || !Utilities.IsCheckedOutByCurrentUser(file.Item))
                    {
                        checkBackIn = true;
                        file.CheckOut();
                    }
                    // If it's checked out by another user then this will throw an informative exception so let it do so.
                }
                string displayTitle = string.Empty;
                Microsoft.SharePoint.WebPartPages.WebPart lvw = null;
                try
                {
                    lvw = new XsltListViewWebPart();
                    if (list.BaseTemplate == SPListTemplateType.Events)
                        lvw = new ListViewWebPart();

                    if (lvw is ListViewWebPart)
                    {
                        ((ListViewWebPart)lvw).ListName = list.ID.ToString("B").ToUpperInvariant();
                        ((ListViewWebPart)lvw).WebId = list.ParentWeb.ID;
                        if (view != null)
                            ((ListViewWebPart)lvw).ViewGuid = view.ID.ToString("B").ToUpperInvariant();
                    }
                    else
                    {
                        ((XsltListViewWebPart)lvw).ListName = list.ID.ToString("B").ToUpperInvariant();
#if !SP2010
                        if (!string.IsNullOrEmpty(jsLink))
                            ((XsltListViewWebPart)lvw).JSLink = jsLink;
#endif
                        ((XsltListViewWebPart)lvw).WebId = list.ParentWeb.ID;
                        if (view != null)
                            ((XsltListViewWebPart)lvw).ViewGuid = view.ID.ToString("B").ToUpperInvariant();
                    }

                    if (linkTitle)
                    {
                        if (view != null)
                            lvw.TitleUrl = view.Url;
                        else
                            lvw.TitleUrl = list.DefaultViewUrl;
                    }

                    if (!string.IsNullOrEmpty(title))
                        lvw.Title = title;


                    lvw.ChromeType = chromeType;

                    displayTitle = lvw.DisplayTitle;

                    WikiPageUtilities.AddWebPartToWikiPage(file.Item, lvw, displayTitle, row, column, addSpace, chromeType, publish);
                }
                finally
                {
                    if (lvw != null)
                        lvw.Dispose();

                    if (file.InDocumentLibrary && Utilities.IsCheckedOut(file.Item) && (checkBackIn || publish))
                        file.CheckIn("Checking in changes to page due to new web part being added: " + displayTitle);

                    if (publish && file.InDocumentLibrary)
                    {
                        file.Publish("Publishing changes to page due to new web part being added: " + displayTitle);
                        if (file.Item.ModerationInformation != null)
                        {
                            file.Approve("Approving changes to page due to new web part being added: " + displayTitle);
                        }
                    }
                }
                return lvw;
            }
        }

    }
}
