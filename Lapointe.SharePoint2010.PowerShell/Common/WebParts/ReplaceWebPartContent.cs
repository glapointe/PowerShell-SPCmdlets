using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Text.RegularExpressions;
#if MOSS
using Microsoft.SharePoint.Portal.WebControls;
using Microsoft.SharePoint.Publishing;
using Microsoft.SharePoint.Publishing.Fields;
using Microsoft.SharePoint.Publishing.WebControls;
#endif
using Microsoft.SharePoint.WebPartPages;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;
using Microsoft.Office.Server.Search.WebControls; //Microsoft.SharePoint.WebPartPages.WebPart;
namespace Lapointe.SharePoint.PowerShell.Common.WebParts
{
    class ReplaceWebPartContent
    {

        #region ReplaceValues Methods

        #region Looping Methods

        /// <summary>
        /// Replaces the content of the various web parts on every file of every web of every site of a given web application.
        /// </summary>
        /// <param name="webApp">The web app that should be searched.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        internal static void ReplaceValues(SPWebApplication webApp, Settings settings)
        {
            Logger.Write("Processing Web Application: " + webApp.DisplayName);

            foreach (SPSite site in webApp.Sites)
            {
                try
                {
                    ReplaceValues(site, settings);
                }
                finally
                {
                    site.Dispose();
                }
            }

            Logger.Write("Finished Processing Web Application: " + webApp.DisplayName + "\r\n");
        }

        /// <summary>
        /// Replaces the content of the various web parts on every file of every web of a given site.
        /// </summary>
        /// <param name="site">The site that should be searched.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        internal static void ReplaceValues(SPSite site, Settings settings)
        {
            Logger.Write("Processing Site: " + site.ServerRelativeUrl);

            foreach (SPWeb web in site.AllWebs)
            {
                try
                {
                    ReplaceValues(web, settings);
                }
                finally
                {
                    web.Dispose();
                }
            }

            Logger.Write("Finished Processing Site: " + site.ServerRelativeUrl + "\r\n");
        }

        /// <summary>
        /// Replaces the content of the various web parts on every file of a given web.
        /// </summary>
        /// <param name="web">The web which should be searched.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        internal static void ReplaceValues(SPWeb web, Settings settings)
        {
            Logger.Write("Processing Web: " + web.ServerRelativeUrl);

            foreach (SPFile file in web.Files)
            {
                ReplaceValues(web, file, settings);
            }

            foreach (SPList list in web.Lists)
            {
                foreach (SPListItem item in list.Items)
                {
                    if (item.File != null && item.File.Url.ToLowerInvariant().EndsWith(".aspx"))
                    {
                        ReplaceValues(web, item.File, settings);
                    }
                }
            }

            Logger.Write("Finished Processing Web: " + web.ServerRelativeUrl + "\r\n");
        }

        /// <summary>
        /// Replaces the content of the various web parts on a given page (file).  This is the main
        /// decision maker method which calls the various worker methods based on web part type.
        /// The following web part types are currently supported (all others will be ignored):
        /// <see cref="ContentEditorWebPart"/>, <see cref="PageViewerWebPart"/>, <see cref="ImageWebPart"/>,
        /// <see cref="SiteDocuments"/>, <see cref="SummaryLinkWebPart"/>, <see cref="DataFormWebPart" />,
        /// <see cref="ContentByQueryWebPart"/>
        /// </summary>
        /// <param name="web">The web to which the file belongs.</param>
        /// <param name="file">The file containing the web parts to search.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        internal static void ReplaceValues(SPWeb web, SPFile file, Settings settings)
        {
            if (file == null)
            {
                return; // This should never be the case.
            }

            if (!Utilities.EnsureAspx(file.Url, true, false))
                return; // We can only handle aspx and master pages.

            SPLimitedWebPartManager manager = null;
            try
            {
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                Logger.Write("Processing File: " + manager.ServerRelativeUrl);
                Regex regex = new Regex(settings.SearchString);

                if (file.InDocumentLibrary && Utilities.IsCheckedOut(file.Item) &&
                    !Utilities.IsCheckedOutByCurrentUser(file.Item))
                {
                    return; // The item is checked out by a different user so leave it alone.
                }

                bool fileModified = false;
                bool wasCheckedOut = true;

                SPLimitedWebPartCollection webParts = manager.WebParts;
                for (int i = 0; i < webParts.Count; i++)
                {
                    WebPart webPart = null;
                    try
                    {
                        webPart = webParts[i] as WebPart;

                        if (webPart == null)
                            continue;

                        bool modified = false;
                        string webPartName = webPart.Title.ToLowerInvariant();
                        if (settings.WebPartName == null || settings.WebPartName.ToLowerInvariant() == webPartName)
                        {
                            // As every web part has different requirements we are only going to consider a small subset.
                            // Custom web parts will not be addressed as there's no interface that can utilized.
                            if (webPart is ContentEditorWebPart)
                            {
                                ContentEditorWebPart wp = (ContentEditorWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp, regex, ref manager, ref wasCheckedOut, ref modified);
                            }
                            else if (webPart is PageViewerWebPart)
                            {
                                PageViewerWebPart wp = (PageViewerWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp, regex, ref manager, ref wasCheckedOut, ref modified);
                            }
                            else if (webPart is ImageWebPart)
                            {
                                ImageWebPart wp = (ImageWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp, regex, ref manager, ref wasCheckedOut, ref modified);
                            }
#if MOSS
                            else if (webPart is MediaWebPart)
                            {
                                MediaWebPart wp = (MediaWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp, regex, ref manager, ref wasCheckedOut, ref modified);
                            }
                            else if (webPart is SiteDocuments)
                            {
                                SiteDocuments wp = (SiteDocuments)webPart;
                                webPart = ReplaceValues(web, file, settings, wp, regex, ref manager, ref wasCheckedOut, ref modified);
                            }
                            else if (webPart is SummaryLinkWebPart)
                            {
                                SummaryLinkWebPart wp = (SummaryLinkWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp, regex, ref manager, ref wasCheckedOut, ref modified);
                            }
                            else if (webPart is ContentByQueryWebPart)
                            {
                                DataFormWebPart wp1 = (DataFormWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp1, regex, ref manager, ref wasCheckedOut, ref modified);
                                if (modified && !settings.Test)
                                    manager.SaveChanges(webPart);

                                CmsDataFormWebPart wp2 = (CmsDataFormWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp2, regex, ref manager, ref wasCheckedOut, ref modified);
                                if (modified && !settings.Test)
                                    manager.SaveChanges(webPart);

                                ContentByQueryWebPart wp3 = (ContentByQueryWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp3, regex, ref manager, ref wasCheckedOut, ref modified);
                            }
                            else if (webPart is CmsDataFormWebPart)
                            {
                                DataFormWebPart wp1 = (DataFormWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp1, regex, ref manager, ref wasCheckedOut, ref modified);
                                if (modified && !settings.Test)
                                    manager.SaveChanges(webPart);

                                CmsDataFormWebPart wp2 = (CmsDataFormWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp2, regex, ref manager, ref wasCheckedOut, ref modified);
                            }
                            else if (webPart is PeopleCoreResultsWebPart)
                            {
                                DataFormWebPart wp1 = (DataFormWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp1, regex, ref manager, ref wasCheckedOut, ref modified);
                                if (modified && !settings.Test)
                                    manager.SaveChanges(webPart);

                                PeopleCoreResultsWebPart wp2 = (PeopleCoreResultsWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp2, regex, ref manager, ref wasCheckedOut, ref modified);
                            }

#endif
                            else if (webPart is DataFormWebPart)
                            {
                                DataFormWebPart wp = (DataFormWebPart)webPart;
                                webPart = ReplaceValues(web, file, settings, wp, regex, ref manager, ref wasCheckedOut, ref modified);
                            }

                            if (modified && !settings.Test)
                                manager.SaveChanges(webPart);

                            if (modified)
                                fileModified = true;
                        }
                    }
                    finally
                    {
                        if (webPart != null)
                            webPart.Dispose();
                    }
                }

                if (!settings.Test)
                {
                    if (fileModified)
                        file.CheckIn("Checking in changes to list item due to automated search and replace (\"" +
                                     settings.SearchString + "\" replaced with \"" + settings.ReplaceString + "\").");

                    if (file.InDocumentLibrary && fileModified && settings.Publish && !wasCheckedOut)
                    {
                        Common.Lists.PublishItems itemPublisher = new Common.Lists.PublishItems();
                        itemPublisher.PublishListItem(file.Item, file.Item.ParentList, settings.Test,
                                                     "\"Automated web part content replacement.\"", null, null);
                    }
                }
                Logger.Write("Finished Processing File: " + manager.ServerRelativeUrl + "\r\n");
            }
            finally
            {
                if (manager != null)
                {
                    manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                    manager.Dispose();
                }
            }
        }

        #endregion

        #region Primary Worker Methods

        /// <summary>
        /// Replaces the content of a <see cref="ContentEditorWebPart"/>.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>The modified web part.  This returned web part is what must be used when saving any changes.</returns>
        internal static WebPart ReplaceValues(SPWeb web,
            SPFile file,
            Settings settings,
            ContentEditorWebPart wp,
            Regex regex,
            ref SPLimitedWebPartManager manager,
            ref bool wasCheckedOut,
            ref bool modified)
        {
            if (wp.Content.FirstChild == null && string.IsNullOrEmpty(wp.ContentLink))
                return wp;

            // The first child of a the content XmlElement for a ContentEditorWebPart is a CDATA section
            // so we want to work with that to make sure we don't accidentally replace the CDATA text itself.
            bool isContentMatch = false;
            if (wp.Content.FirstChild != null)
                isContentMatch = regex.IsMatch(wp.Content.FirstChild.InnerText);
            bool isLinkMatch = false;
            if (!string.IsNullOrEmpty(wp.ContentLink))
                isLinkMatch = regex.IsMatch(wp.ContentLink);

            if (!isContentMatch && !isLinkMatch)
                return wp;

            string content;
            if (isContentMatch)
                content = wp.Content.FirstChild.InnerText;
            else
                content = wp.ContentLink;

            string result = content;
            if (!string.IsNullOrEmpty(content))
                result = regex.Replace(content, settings.ReplaceString);

            Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                              file.ServerRelativeUrl, wp.Title, content, result);
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (ContentEditorWebPart)manager.WebParts[wp.ID];

                if (isContentMatch)
                    wp.Content = GetDataAsXmlElement("Content", "http://schemas.microsoft.com/WebPart/v2/ContentEditor", result);
                else
                    wp.ContentLink = result;

                modified = true;
            }
            return wp;
        }

        /// <summary>
        /// Replaces the content of a <see cref="PageViewerWebPart" /> web part.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>The modified web part.  This returned web part is what must be used when saving any changes.</returns>
        internal static WebPart ReplaceValues(SPWeb web,
            SPFile file,
            Settings settings,
            PageViewerWebPart wp,
            Regex regex,
            ref SPLimitedWebPartManager manager,
            ref bool wasCheckedOut,
            ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.ContentLink))
                return wp;

            bool isLinkMatch = regex.IsMatch(wp.ContentLink);

            if (!isLinkMatch)
                return wp;

            string content = wp.ContentLink;
            string result = content;

            if (!string.IsNullOrEmpty(content))
                result = regex.Replace(content, settings.ReplaceString);

            Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                              file.ServerRelativeUrl, wp.Title, content, result);
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (PageViewerWebPart)manager.WebParts[wp.ID];

                wp.ContentLink = result;

                modified = true;
            }
            return wp;
        }


        /// <summary>
        /// Replaces the content of a <see cref="DataFormWebPart"/> web part.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>
        /// The modified web part.  This returned web part is what must be used when saving any changes.
        /// </returns>
        internal static WebPart ReplaceValues(SPWeb web,
           SPFile file,
           Settings settings,
           DataFormWebPart wp,
           Regex regex,
           ref SPLimitedWebPartManager manager,
           ref bool wasCheckedOut,
           ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.DataSourcesString) && string.IsNullOrEmpty(wp.ParameterBindings) && string.IsNullOrEmpty(wp.ListName) && string.IsNullOrEmpty(wp.Xsl) && string.IsNullOrEmpty(wp.XslLink))
                return wp;

            bool isDataSourcesMatch = false;
            if (!string.IsNullOrEmpty(wp.DataSourcesString))
                isDataSourcesMatch = regex.IsMatch(wp.DataSourcesString);

            bool isParameterBindingsMatch = false;
            if (!string.IsNullOrEmpty(wp.ParameterBindings))
                isParameterBindingsMatch = regex.IsMatch(wp.ParameterBindings);

            bool isListNameMatch = false;
            if (!string.IsNullOrEmpty(wp.ListName))
                isListNameMatch = regex.IsMatch(wp.ListName);

            bool isXslMatch = false;
            if (!string.IsNullOrEmpty(wp.Xsl))
                isXslMatch = regex.IsMatch(wp.Xsl);

            bool isXslLinkMatch = false;
            if (!string.IsNullOrEmpty(wp.XslLink))
                isXslLinkMatch = regex.IsMatch(wp.XslLink);

            if (!isDataSourcesMatch && !isParameterBindingsMatch && !isListNameMatch && !isXslMatch && !isXslLinkMatch)
                return wp;

            string dataSourcesContent = wp.DataSourcesString;
            string parameterBindingsContent = wp.ParameterBindings;
            string listNameContent = wp.ListName;
            string xslContent = wp.Xsl;
            string xslLinkContent = wp.XslLink;

            string dataSourcesResult = dataSourcesContent;
            string parameterBindingsResult = parameterBindingsContent;
            string listNameResult = listNameContent;
            string xslResult = xslContent;
            string xslLinkResult = xslLinkContent;

            if (!string.IsNullOrEmpty(dataSourcesContent))
                dataSourcesResult = regex.Replace(dataSourcesContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(listNameContent))
                listNameResult = regex.Replace(listNameContent, settings.ReplaceString);
            try
            {
                if (!string.IsNullOrEmpty(parameterBindingsContent))
                {
                    if (!settings.UnsafeXml)
                        parameterBindingsResult = ReplaceXmlValues(regex, settings.ReplaceString, parameterBindingsContent);
                    else
                        parameterBindingsResult = regex.Replace(parameterBindingsContent, settings.ReplaceString);
                }
            }
            catch (XmlException ex)
            {
                isParameterBindingsMatch = false;
                string msg = string.Format(
                    "WARNING: An error occured replacing data in a DataFormWebPart: File={0}, WebPart={1}\r\n{2}",
                    file.ServerRelativeUrl, wp.Title, Utilities.FormatException(ex));
                Logger.WriteWarning(msg);
            }
            if (!string.IsNullOrEmpty(xslContent))
                xslResult = regex.Replace(xslContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(xslLinkContent))
                xslLinkResult = regex.Replace(xslLinkContent, settings.ReplaceString);

            if (isDataSourcesMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, dataSourcesContent, dataSourcesResult);
            if (isParameterBindingsMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, parameterBindingsContent, parameterBindingsResult);
            if (isListNameMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, listNameContent, listNameResult);
            if (isXslMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, xslContent, xslResult);
            if (isXslLinkMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, xslLinkContent, xslLinkResult);
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (DataFormWebPart)manager.WebParts[wp.ID];

                if (isDataSourcesMatch)
                    wp.DataSourcesString = dataSourcesResult;
                if (isParameterBindingsMatch)
                    wp.ParameterBindings = parameterBindingsResult;
                if (isListNameMatch)
                    wp.ListName = listNameResult;
                if (isXslMatch)
                    wp.Xsl = xslResult;
                if (isXslLinkMatch)
                    wp.XslLink = xslLinkResult;

                modified = true;
            }
            return wp;
        }

        /// <summary>
        /// Replaces the content of a <see cref="ImageWebPart"/> web part.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>The modified web part.  This returned web part is what must be used when saving any changes.</returns>
        internal static WebPart ReplaceValues(SPWeb web,
           SPFile file,
           Settings settings,
           ImageWebPart wp,
           Regex regex,
           ref SPLimitedWebPartManager manager,
           ref bool wasCheckedOut,
           ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.ImageLink) && string.IsNullOrEmpty(wp.AlternativeText))
                return wp;

            bool isAltTextMatch = false;
            if (!string.IsNullOrEmpty(wp.AlternativeText))
                isAltTextMatch = regex.IsMatch(wp.AlternativeText);

            bool isLinkMatch = false;
            if (!string.IsNullOrEmpty(wp.ImageLink))
                isLinkMatch = regex.IsMatch(wp.ImageLink);

            if (!isAltTextMatch && !isLinkMatch)
                return wp;

            string altTextContent = wp.AlternativeText;
            string linkContent = wp.ImageLink;

            string altTextResult = altTextContent;
            string linkResult = linkContent;

            if (!string.IsNullOrEmpty(altTextContent))
                altTextResult = regex.Replace(altTextContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(linkContent))
                linkResult = regex.Replace(linkContent, settings.ReplaceString);

            if (isAltTextMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, altTextContent, altTextResult);
            if (isLinkMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, linkContent, linkResult);
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (ImageWebPart)manager.WebParts[wp.ID];

                if (isAltTextMatch)
                    wp.AlternativeText = altTextResult;
                if (isLinkMatch)
                    wp.ImageLink = linkResult;

                modified = true;
            }
            return wp;
        }

#if MOSS

        internal static WebPart ReplaceValues(SPWeb web,
           SPFile file,
           Settings settings,
           PeopleCoreResultsWebPart wp,
           Regex regex,
           ref SPLimitedWebPartManager manager,
           ref bool wasCheckedOut,
           ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.FixedQuery) && string.IsNullOrEmpty(wp.Xsl))
                return wp;

            bool isFixedQueryMatch = false;
            if (!string.IsNullOrEmpty(wp.FixedQuery))
                isFixedQueryMatch = regex.IsMatch(wp.FixedQuery);

            bool isXslMatch = false;
            if (!string.IsNullOrEmpty(wp.Xsl))
                isXslMatch = regex.IsMatch(wp.Xsl);

            if (!isFixedQueryMatch && !isXslMatch)
                return wp;

            string fixedQueryContent = wp.FixedQuery;
            string xslContent = wp.Xsl;

            string fixedQueryResult = fixedQueryContent;
            string xslResult = xslContent;

            if (!string.IsNullOrEmpty(fixedQueryContent))
                fixedQueryResult = regex.Replace(fixedQueryContent, settings.ReplaceString);
            try
            {
                if (!string.IsNullOrEmpty(xslContent))
                {
                    if (!settings.UnsafeXml)
                        xslResult = ReplaceXmlValues(regex, settings.ReplaceString, xslContent);
                    else
                        xslResult = regex.Replace(xslContent, settings.ReplaceString);
                }
            }
            catch (XmlException ex)
            {
                isXslMatch = false;
                string msg = string.Format(
                    "WARNING: An error occured replacing data in a PeopleCoreResultsWebPart: File={0}, WebPart={1}\r\n{2}",
                    file.ServerRelativeUrl, wp.Title, Utilities.FormatException(ex));
                Logger.WriteWarning(msg);
            }
            if (isFixedQueryMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, fixedQueryContent, fixedQueryResult);
            if (isXslMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, xslContent, xslResult);
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (PeopleCoreResultsWebPart)manager.WebParts[wp.ID];

                if (isFixedQueryMatch)
                    wp.FixedQuery = fixedQueryResult;
                if (isXslMatch)
                    wp.Xsl = xslResult;

                modified = true;
            }
            return wp;
        }

        /// <summary>
        /// Replaces the content of a <see cref="ContentByQueryWebPart"/> web part.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>The modified web part.  This returned web part is what must be used when saving any changes.</returns>
        internal static WebPart ReplaceValues(SPWeb web,
           SPFile file,
           Settings settings,
           ContentByQueryWebPart wp,
           Regex regex,
           ref SPLimitedWebPartManager manager,
           ref bool wasCheckedOut,
           ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.ListGuid) && string.IsNullOrEmpty(wp.WebUrl) && string.IsNullOrEmpty(wp.GroupBy))
                return wp;

            bool isListGuidMatch = false;
            if (!string.IsNullOrEmpty(wp.ListGuid))
                isListGuidMatch = regex.IsMatch(wp.ListGuid);

            bool isWebUrlMatch = false;
            if (!string.IsNullOrEmpty(wp.WebUrl))
                isWebUrlMatch = regex.IsMatch(wp.WebUrl);

            bool isGroupByMatch = false;
            if (!string.IsNullOrEmpty(wp.GroupBy))
                isGroupByMatch = regex.IsMatch(wp.GroupBy);

            if (!isListGuidMatch && !isWebUrlMatch && !isGroupByMatch)
                return wp;

            string listGuidContent = wp.ListGuid;
            string webUrlContent = wp.WebUrl;
            string groupByContent = wp.GroupBy;

            string listGuidResult = listGuidContent;
            string webUrlResult = webUrlContent;
            string groupByResult = groupByContent;

            if (!string.IsNullOrEmpty(listGuidContent))
                listGuidResult = regex.Replace(listGuidContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(webUrlContent))
                webUrlResult = regex.Replace(webUrlContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(groupByContent))
                groupByResult = regex.Replace(groupByContent, settings.ReplaceString);

            if (isListGuidMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, listGuidContent, listGuidResult);
            if (isWebUrlMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, webUrlContent, webUrlResult);

            if (isGroupByMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, groupByContent, groupByResult);

            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (ContentByQueryWebPart)manager.WebParts[wp.ID];

                if (isListGuidMatch)
                    wp.ListGuid = listGuidResult;
                if (isWebUrlMatch)
                    wp.WebUrl = webUrlResult;
                if (isGroupByMatch)
                    wp.GroupBy = groupByResult;

                modified = true;
            }
            return wp;
        }

        /// <summary>
        /// Replaces the content of a <see cref="CmsDataFormWebPart"/> web part.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>
        /// The modified web part.  This returned web part is what must be used when saving any changes.
        /// </returns>
        internal static WebPart ReplaceValues(SPWeb web,
           SPFile file,
           Settings settings,
           CmsDataFormWebPart wp,
           Regex regex,
           ref SPLimitedWebPartManager manager,
           ref bool wasCheckedOut,
           ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.HeaderXslLink) && string.IsNullOrEmpty(wp.ItemXslLink) && string.IsNullOrEmpty(wp.MainXslLink))
                return wp;

            bool isHeaderXslLinkMatch = false;
            if (!string.IsNullOrEmpty(wp.HeaderXslLink))
                isHeaderXslLinkMatch = regex.IsMatch(wp.HeaderXslLink);

            bool isItemXslLinkMatch = false;
            if (!string.IsNullOrEmpty(wp.ParameterBindings))
                isItemXslLinkMatch = regex.IsMatch(wp.ItemXslLink);

            bool isMainXslLinkMatch = false;
            if (!string.IsNullOrEmpty(wp.ListName))
                isMainXslLinkMatch = regex.IsMatch(wp.MainXslLink);

            if (!isHeaderXslLinkMatch && !isItemXslLinkMatch && !isMainXslLinkMatch)
                return wp;

            string headerXslLinkContent = wp.HeaderXslLink;
            string itemXslLinkContent = wp.ItemXslLink;
            string mainXslLinkContent = wp.MainXslLink;

            string headerXslLinkResult = headerXslLinkContent;
            string itemXslLinkResult = itemXslLinkContent;
            string mainXslLinkResult = mainXslLinkContent;

            if (!string.IsNullOrEmpty(headerXslLinkContent))
                headerXslLinkResult = regex.Replace(headerXslLinkContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(itemXslLinkContent))
                itemXslLinkResult = regex.Replace(itemXslLinkContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(mainXslLinkContent))
                mainXslLinkResult = regex.Replace(mainXslLinkContent, settings.ReplaceString);

            if (isHeaderXslLinkMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, headerXslLinkContent, headerXslLinkResult);
            if (isItemXslLinkMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, itemXslLinkContent, itemXslLinkResult);
            if (isMainXslLinkMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, mainXslLinkContent, mainXslLinkResult);
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (CmsDataFormWebPart)manager.WebParts[wp.ID];

                if (isHeaderXslLinkMatch)
                    wp.HeaderXslLink = headerXslLinkResult;
                if (isItemXslLinkMatch)
                    wp.ItemXslLink = itemXslLinkResult;
                if (isMainXslLinkMatch)
                    wp.MainXslLink = mainXslLinkResult;

                modified = true;
            }
            return wp;
        }

        /// <summary>
        /// Replaces the content of a <see cref="MediaWebPart"/> web part.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>The modified web part.  This returned web part is what must be used when saving any changes.</returns>
        internal static WebPart ReplaceValues(SPWeb web,
           SPFile file,
           Settings settings,
           MediaWebPart wp,
           Regex regex,
           ref SPLimitedWebPartManager manager,
           ref bool wasCheckedOut,
           ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.MediaSource) && string.IsNullOrEmpty(wp.TemplateSource) && string.IsNullOrEmpty(wp.PreviewImageSource))
                return wp;

            bool isMediaSourceMatch = false;
            if (!string.IsNullOrEmpty(wp.MediaSource))
                isMediaSourceMatch = regex.IsMatch(wp.MediaSource);

            bool isTemplateSourceMatch = false;
            if (!string.IsNullOrEmpty(wp.TemplateSource))
                isTemplateSourceMatch = regex.IsMatch(wp.TemplateSource);

            bool isPreviewImageSourceMatch = false;
            if (!string.IsNullOrEmpty(wp.PreviewImageSource))
                isPreviewImageSourceMatch = regex.IsMatch(wp.PreviewImageSource);

            if (!isMediaSourceMatch && !isTemplateSourceMatch && !isPreviewImageSourceMatch)
                return wp;

            string mediaSourceContent = wp.MediaSource;
            string templateSourceContent = wp.TemplateSource;
            string previewImageSourceContent = wp.PreviewImageSource;

            string mediaSourceResult = mediaSourceContent;
            string templateSourceResult = templateSourceContent;
            string previewImageSourceResult = previewImageSourceContent;

            if (!string.IsNullOrEmpty(mediaSourceContent))
                mediaSourceResult = regex.Replace(mediaSourceContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(templateSourceContent))
                templateSourceResult = regex.Replace(templateSourceContent, settings.ReplaceString);
            if (!string.IsNullOrEmpty(previewImageSourceContent))
                previewImageSourceResult = regex.Replace(previewImageSourceContent, settings.ReplaceString);

            if (isMediaSourceMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, mediaSourceContent, mediaSourceResult);
            if (isTemplateSourceMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, templateSourceContent, templateSourceResult);
            if (isPreviewImageSourceMatch)
                Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                                  file.ServerRelativeUrl, wp.Title, previewImageSourceContent, previewImageSourceResult);
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (MediaWebPart)manager.WebParts[wp.ID];

                if (isMediaSourceMatch)
                    wp.MediaSource = mediaSourceResult;
                if (isTemplateSourceMatch)
                    wp.TemplateSource = templateSourceResult;
                if (isPreviewImageSourceMatch)
                    wp.PreviewImageSource = previewImageSourceResult;

                modified = true;
            }
            return wp;
        }

        /// <summary>
        /// Replaces the content of a <see cref="SiteDocuments"/> web part.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>The modified web part.  This returned web part is what must be used when saving any changes.</returns>
        internal static WebPart ReplaceValues(SPWeb web,
            SPFile file,
            Settings settings,
            SiteDocuments wp,
            Regex regex,
            ref SPLimitedWebPartManager manager,
            ref bool wasCheckedOut,
            ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.UserTabs))
                return wp;

            bool isLinkMatch = regex.IsMatch(wp.UserTabs);

            if (!isLinkMatch)
                return wp;

            string content = wp.UserTabs;
            string result = content;
            if (!string.IsNullOrEmpty(content))
                result = regex.Replace(content, settings.ReplaceString);

            Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                              file.ServerRelativeUrl, wp.Title, content, result);
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (SiteDocuments)manager.WebParts[wp.ID];

                wp.UserTabs = result;

                modified = true;
            }
            return wp;
        }

        /// <summary>
        /// Replaces the content of a <see cref="SummaryLinkWebPart"/> web part.
        /// </summary>
        /// <param name="web">The web that the file belongs to.</param>
        /// <param name="file">The file that the web part is associated with.</param>
        /// <param name="settings">The settings object containing user provided parameters.</param>
        /// <param name="wp">The web part whose content will be replaced.</param>
        /// <param name="regex">The regular expression object which contains the search pattern.</param>
        /// <param name="manager">The web part manager.  This value may get updated during this method call.</param>
        /// <param name="wasCheckedOut">if set to <c>true</c> then the was checked out prior to this method being called.</param>
        /// <param name="modified">if set to <c>true</c> then the web part was modified as a result of this method being called.</param>
        /// <returns>The modified web part.  This returned web part is what must be used when saving any changes.</returns>
        internal static WebPart ReplaceValues(SPWeb web,
            SPFile file,
            Settings settings,
            SummaryLinkWebPart wp,
            Regex regex,
            ref SPLimitedWebPartManager manager,
            ref bool wasCheckedOut,
            ref bool modified)
        {
            if (string.IsNullOrEmpty(wp.SummaryLinkStore))
                return wp;

            bool isLinkMatch = false;
            string originalContent = wp.SummaryLinkStore;

            SummaryLinkFieldValue links = wp.SummaryLinkValue;

            // Make all appropriate changes here and then reset the value below.
            // We don't want to manipulate the store itself because we risk messing up the XML.
            foreach (SummaryLink link in links.SummaryLinks)
            {
                if (!string.IsNullOrEmpty(link.Description) && regex.IsMatch(link.Description))
                {
                    link.Description = regex.Replace(link.Description, settings.ReplaceString);
                    isLinkMatch = true;
                }
                if (!string.IsNullOrEmpty(link.ImageUrl) && regex.IsMatch(link.ImageUrl))
                {
                    link.ImageUrl = regex.Replace(link.ImageUrl, settings.ReplaceString);
                    isLinkMatch = true;
                }
                if (!string.IsNullOrEmpty(link.ImageUrlAltText) && regex.IsMatch(link.ImageUrlAltText))
                {
                    link.ImageUrlAltText = regex.Replace(link.ImageUrlAltText, settings.ReplaceString);
                    isLinkMatch = true;
                }
                if (!string.IsNullOrEmpty(link.LinkUrl) && regex.IsMatch(link.LinkUrl))
                {
                    link.LinkUrl = regex.Replace(link.LinkUrl, settings.ReplaceString);
                    isLinkMatch = true;
                }
                if (!string.IsNullOrEmpty(link.Title) && regex.IsMatch(link.Title))
                {
                    link.Title = regex.Replace(link.Title, settings.ReplaceString);
                    isLinkMatch = true;
                }
                if (!string.IsNullOrEmpty(link.LinkToolTip) && regex.IsMatch(link.LinkToolTip))
                {
                    link.LinkToolTip = regex.Replace(link.LinkToolTip, settings.ReplaceString);
                    isLinkMatch = true;
                }
            }
            if (!isLinkMatch)
                return wp;

            Logger.Write("Match found: File={0}, WebPart={1}, Replacement={2} => {3}",
                              file.ServerRelativeUrl, wp.Title, originalContent, links.ToString());
            if (!settings.Test)
            {
                if (file.CheckOutType == SPFile.SPCheckOutType.None)
                {
                    file.CheckOut();
                    wasCheckedOut = false;
                }
                // We need to reset the manager and the web part because a checkout (now or from an earlier call) 
                // could mess things up so safest to just reset every time.
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);

                wp.Dispose();
                wp = (SummaryLinkWebPart)manager.WebParts[wp.ID];

                wp.SummaryLinkValue = links;

                modified = true;
            }
            return wp;
        }
#endif

        #endregion

        #endregion

        #region Utility Methods

        /// <summary>
        /// Replaces XML attribute and text node values.
        /// </summary>
        /// <param name="regex">The regular expression object.</param>
        /// <param name="replaceString">The replacement string.</param>
        /// <param name="xmlContent">The XML to replace.</param>
        /// <returns></returns>
        private static string ReplaceXmlValues(Regex regex, string replaceString, string xmlContent)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.PreserveWhitespace = true;

            try
            {
                // The incomming xml may not have a root node so we need to add a temp root in order to load it.
                xmlDoc.LoadXml("<STSADM_TEMP>" + xmlContent + "</STSADM_TEMP>");
            }
            catch (XmlException ex)
            {
                ex.Data.Add("XmlContent", xmlContent);

                throw;
            }

            ReplaceXmlValues(regex, replaceString, xmlDoc.DocumentElement);

            return xmlDoc.DocumentElement.InnerXml;
        }

        /// <summary>
        /// Replaces XML attribute and text node values.
        /// </summary>
        /// <param name="regex">The regular expression object.</param>
        /// <param name="replaceString">The replacement string.</param>
        /// <param name="node">The node.</param>
        private static void ReplaceXmlValues(Regex regex, string replaceString, XmlNode node)
        {
            foreach (XmlNode childNode in node.ChildNodes)
            {
                ReplaceXmlValues(regex, replaceString, childNode);
            }

            if (node.Attributes != null)
            {
                foreach (XmlAttribute attr in node.Attributes)
                {
                    if (!string.IsNullOrEmpty(attr.Value) && regex.IsMatch(attr.Value))
                    {
                        attr.Value = regex.Replace(attr.Value, replaceString);
                    }
                }
            }
            if (node.ChildNodes.Count == 0 && !string.IsNullOrEmpty(node.Value))
            {
                if (regex.IsMatch(node.Value))
                {
                    node.Value = regex.Replace(node.Value, replaceString);
                }
            }
        }

        /// <summary>
        /// Gets the data as XML element.  This method is critical for the <see cref="ContentEditorWebPart"/>.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="ns">The ns.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        internal static XmlElement GetDataAsXmlElement(string name, string ns, string data)
        {
            XmlDocument xdoc = new XmlDocument();
            XmlElement element = xdoc.CreateElement(name, ns);
            if (data != null)
            {
                if (data.Length <= 0)
                {
                    return element;
                }
                if (data.IndexOf("]]>") != -1)
                {
                    element.InnerText = data;
                    return element;
                }
                element.InnerXml = "<![CDATA[" + data + "]]>";
            }
            return element;
        }


        #endregion

        #region Internal Classes

        /// <summary>
        /// Data class used for passing user provided settings into the various worker methods.
        /// </summary>
        public class Settings
        {
            public string SearchString;
            public string ReplaceString;
            public string WebPartName;
            public bool Publish;
            public bool Test;
            public bool UnsafeXml = false;
        }

        #endregion
    }
}
