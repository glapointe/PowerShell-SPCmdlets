using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing;
using System.Management.Automation;

namespace Lapointe.SharePoint.PowerShell.Common.Pages
{
    internal class CreatePublishingPage
    {
        /// <summary>
        /// Creates the page.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="pageName">Name of the page.</param>
        /// <param name="title">The title.</param>
        /// <param name="layoutName">Name of the layout.</param>
        /// <param name="fieldDataCollection">The field data collection.</param>
        public static string CreatePage(string url, string pageName, string title, string layoutName, Dictionary<string, string> fieldDataCollection, bool test)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            {
                PublishingPage page = CreatePage(web, pageName, title, layoutName, fieldDataCollection, test);
                return site.MakeFullUrl(page.Url);
            }
        }

        public static PublishingPage CreatePage(SPWeb web, string pageName, string title, string layoutName, Dictionary<string, string> fieldDataCollection, bool test)
        {
            if (!PublishingWeb.IsPublishingWeb(web))
                throw new ArgumentException("The specified web is not a publishing web.");

            PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);
            PageLayout layout = null;
            string availableLayouts = string.Empty;
            foreach (PageLayout lo in pubweb.GetAvailablePageLayouts())
            {
                availableLayouts += "\t" + lo.Name + "\r\n";
                if (lo.Name.ToLowerInvariant() == layoutName.ToLowerInvariant())
                {
                    layout = lo;
                    break;
                }
            }
            if (layout == null)
            {
                if (PublishingSite.IsPublishingSite(web.Site))
                {
                    Logger.WriteWarning("The specified page layout could not be found among the list of available page layouts for the web. Available layouts are:\r\n" + availableLayouts);
                    availableLayouts = string.Empty;
                    foreach (PageLayout lo in (new PublishingSite(web.Site).PageLayouts))
                    {
                        availableLayouts += "\t" + lo.Name + "\r\n";
                        if (lo.Name.ToLowerInvariant() == layoutName.ToLowerInvariant())
                        {
                            layout = lo;
                            break;
                        }
                    }
                }
                if (layout == null)
                    throw new ArgumentException("The layout specified could not be found. Available layouts are:\r\n" + availableLayouts);
            }

            if (!pageName.ToLowerInvariant().EndsWith(".aspx"))
                pageName += ".aspx";

            PublishingPage page = null;
            SPListItem item = null;
            if (test)
            {
                Logger.Write("Page to be created at {0}", pubweb.Url);
            }
            else
            {
                page = pubweb.GetPublishingPages().Add(pageName, layout);
                page.Title = title;
                item = page.ListItem;
            }

            foreach (string fieldName in fieldDataCollection.Keys)
            {
                string fieldData = fieldDataCollection[fieldName];

                try
                {
                    SPField field = item.Fields.GetFieldByInternalName(fieldName);

                    if (field.ReadOnlyField)
                    {
                        Logger.Write("Field '{0}' is read only and will not be updated.", field.InternalName);
                        continue;
                    }

                    if (field.Type == SPFieldType.Computed)
                    {
                        Logger.Write("Field '{0}' is a computed column and will not be updated.", field.InternalName);
                        continue;
                    }

                    if (!test)
                    {
                        if (field.Type == SPFieldType.URL)
                            item[field.Id] = new SPFieldUrlValue(fieldData);
                        else if (field.Type == SPFieldType.User)
                            Common.Pages.CreatePublishingPage.SetUserField(web, item, field, fieldData);
                        else
                            item[field.Id] = fieldData;
                    }
                    else
                    {
                        Logger.Write("Field '{0}' would be set to '{1}'.", field.InternalName, fieldData);
                    }
                }
                catch (ArgumentException ex)
                {
                    Logger.WriteException(new ErrorRecord(new Exception(string.Format("Could not set field {0} for item {1}.", fieldName, item.ID), ex), null, ErrorCategory.InvalidArgument, item));
                }
            }
            if (page != null)
                page.Update();
            return page;
        }

        /// <summary>
        /// Sets the user field.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="item">The item.</param>
        /// <param name="field">The field.</param>
        /// <param name="fieldData">The field data.</param>
        internal static void SetUserField(SPWeb web, SPListItem item, SPField field, string fieldData)
        {
            string[] accounts = fieldData.Split(new string[] { ";#" }, StringSplitOptions.RemoveEmptyEntries);
            string accountsToAdd = string.Empty;
            for (int i = 1; i < accounts.Length; i = i + 2)
            {
                bool found = false;
                foreach (SPUser user in web.AllUsers)
                {
                    if (user.Name == accounts[i])
                    {
                        found = true;
                        accountsToAdd += ";#" + user.ID + ";#" + user.Name;
                        break;
                    }
                }
                if (!found)
                {
                    foreach (SPUser user in web.SiteUsers)
                    {
                        if (user.Name == accounts[i])
                        {
                            accountsToAdd += ";#" + user.ID + ";#" + user.Name;
                            break;
                        }
                    }
                }
            }
            if (!string.IsNullOrEmpty(accountsToAdd))
            {
                item[field.Id] = accountsToAdd.Trim(new char[] { ';', '#' });
            }
        }
    }
}
