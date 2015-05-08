using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Deployment;
using Microsoft.SharePoint.Navigation;
#if MOSS
using Microsoft.SharePoint.Publishing;
using Microsoft.SharePoint.Publishing.Fields;
using Microsoft.SharePoint.Publishing.Navigation;
using Microsoft.SharePoint.Publishing.WebControls;
#endif
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebPartPages;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart; //Microsoft.SharePoint.WebPartPages.WebPart;
using Lapointe.SharePoint.PowerShell.Common.Features;

namespace Lapointe.SharePoint.PowerShell.Common.SiteCollections
{
    class RepairSiteCollectionImportedFromSubSite
    {

        /// <summary>
        /// Repairs the site.
        /// </summary>
        /// <param name="sourceurl">The sourceurl.</param>
        /// <param name="targeturl">The targeturl.</param>
        public static void RepairSite(string sourceurl, string targeturl)
        {
            using (SPSite targetSite = new SPSite(targeturl.TrimEnd('/')))
            using (SPWeb targetWeb = targetSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(targeturl.TrimEnd('/'))])
            using (SPSite sourceSite = new SPSite(sourceurl.TrimEnd('/')))
            using (SPWeb sourceWeb = sourceSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(sourceurl.TrimEnd('/'))])
            {
                AddMissingFeatures(sourceSite, sourceWeb, targetSite, targetWeb);

                try
                {
                    Logger.Write("Progress: Begin copying content types...");
                    Common.ContentTypes.CopyContentTypes ctCopier = new Common.ContentTypes.CopyContentTypes();
                    ctCopier.Copy(sourceurl, targeturl);
                }
                finally
                {
                    Logger.Write("Progress: End copying content types.");
                }
            }

            // We need to re-open all the objects as some values such as content types need to be refreshed.
            using (SPSite targetSite = new SPSite(targeturl.TrimEnd('/')))
            using (SPWeb targetWeb = targetSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(targeturl.TrimEnd('/'))])
            using (SPSite sourceSite = new SPSite(sourceurl.TrimEnd('/')))
            using (SPWeb sourceWeb = sourceSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(sourceurl.TrimEnd('/'))])
            {
                SetMasterPageGallerySettings(sourceSite, targetSite, targetWeb);

#if MOSS
                PublishingWeb targetPublishingWeb = PublishingWeb.GetPublishingWeb(targetWeb);
                PublishingSite targetPublishingSite = new PublishingSite(targetSite);
                PublishingWeb sourcePublishingWeb = PublishingWeb.GetPublishingWeb(sourceWeb);
                PublishingSite sourcePublishingSite = new PublishingSite(sourceSite);

                FixPageLayoutsAndSiteTemplates(sourcePublishingSite, sourcePublishingWeb, targetPublishingSite, targetPublishingWeb, targetSite, targetWeb);
#endif
            }

            Common.TimerJobs.ExecAdmSvcJobs.Execute(false, true);

            using (SPSite targetSite = new SPSite(targeturl.TrimEnd('/')))
            using (SPWeb targetWeb = targetSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(targeturl.TrimEnd('/'))])
            using (SPSite sourceSite = new SPSite(sourceurl.TrimEnd('/')))
            using (SPWeb sourceWeb = sourceSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(sourceurl.TrimEnd('/'))])
            {
#if MOSS
                PublishingWeb targetPublishingWeb = PublishingWeb.GetPublishingWeb(targetWeb);
                
                FixPublishingPages(targetSite);

                SetGlobalNavigation(targetPublishingWeb);

                //RetargetGroupedListingsWebPart.RetargetGroupedListings(targetWeb, "grouped listings");
#endif
                //RepairDiscussionLists(targetSite);

                RetargetMiscWebParts(sourceSite, sourceWeb, targetWeb);
            }
        }

        #region Worker Methods

        #region Retargart Miscellaneous Web Parts

        /// <summary>
        /// Retargets the misc web parts.
        /// </summary>
        /// <param name="sourceSite">The source site.</param>
        /// <param name="sourceWeb">The source web.</param>
        /// <param name="targetWeb">The target web.</param>
        internal static void RetargetMiscWebParts(SPSite sourceSite, SPWeb sourceWeb, SPWeb targetWeb)
        {
            StringDictionary listMap = new StringDictionary();

            // Handle the primary web first then address the child webs
            FindMatchingLists(listMap, sourceWeb, targetWeb);

            string parentTargetUrl = targetWeb.ServerRelativeUrl;
            string parentSourceUrl = sourceWeb.ServerRelativeUrl;

            RetargetMiscWebPartsRecursiveHelper(listMap, parentSourceUrl, parentTargetUrl, sourceSite, targetWeb);

            RetargetMiscWebParts(targetWeb, listMap);
        }

        /// <summary>
        /// Recursive helper for retargetting the misc web parts.
        /// </summary>
        /// <param name="listMap">The list map.</param>
        /// <param name="parentSourceUrl">The parent source URL.</param>
        /// <param name="parentTargetUrl">The parent target URL.</param>
        /// <param name="sourceSite">The source site.</param>
        /// <param name="targetWeb">The target web.</param>
        private static void RetargetMiscWebPartsRecursiveHelper(StringDictionary listMap, string parentSourceUrl, string parentTargetUrl, SPSite sourceSite, SPWeb targetWeb)
        {
            foreach (SPWeb childTargetWeb in targetWeb.Webs)
            {
                try
                {
                    string targetParentRelativeUrl = childTargetWeb.ServerRelativeUrl.Substring(parentTargetUrl.Length).TrimEnd('/');
                    string sourceUrl = Utilities.ConcatServerRelativeUrls(parentSourceUrl, targetParentRelativeUrl).TrimEnd('/');


                    using (SPWeb childSourceWeb = sourceSite.AllWebs[sourceUrl])
                    {
                        if (childSourceWeb.Exists)
                        {
                            FindMatchingLists(listMap, childSourceWeb, childTargetWeb);
                        }
                    }
                    RetargetMiscWebPartsRecursiveHelper(listMap, parentSourceUrl, parentTargetUrl, sourceSite, childTargetWeb);

                }
                finally
                {
                    childTargetWeb.Dispose();
                }
            }
        }

        /// <summary>
        /// Finds the matching lists.
        /// </summary>
        /// <param name="listMap">The list map.</param>
        /// <param name="sourceWeb">The source web.</param>
        /// <param name="targetWeb">The target web.</param>
        private static void FindMatchingLists(StringDictionary listMap, SPWeb sourceWeb, SPWeb targetWeb)
        {
            foreach (SPList targetList in targetWeb.Lists)
            {
                try
                {
                    // Try to find a matching list in the source web
                    SPList sourceList = sourceWeb.Lists[targetList.Title];
                    listMap.Add(sourceList.ID.ToString(), targetList.ID.ToString());
                }
                catch (ArgumentException)
                {
                    // There is no matching source list - this should happen if our source is the actual source and not a model.
                    // If the source is just a model then we're wasting our time here but unfortunately there's no way to
                    // determine that without asking the user.
                }
            }
        }

        /// <summary>
        /// Retargets the misc web parts for the target web.  Loops through all sub-webs and all files.
        /// </summary>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="listMap">The list map.</param>
        private static void RetargetMiscWebParts(SPWeb targetWeb, StringDictionary listMap)
        {
            foreach (SPWeb subweb in targetWeb.Webs)
            {
                try
                {
                    RetargetMiscWebParts(subweb, listMap);
                }
                finally
                {
                    subweb.Dispose();
                }
            }

            foreach (SPFile file in targetWeb.Files)
            {
                RetargetMiscWebParts(targetWeb, listMap, file);
            }

            foreach (SPList list in targetWeb.Lists)
            {
                foreach (SPListItem item in list.Items)
                {
                    if (item.File != null && item.File.Url.ToLowerInvariant().EndsWith(".aspx"))
                    {
                        RetargetMiscWebParts(targetWeb, listMap, item.File);
                    }
                }
            }
            targetWeb.Dispose();
        }

        /// <summary>
        /// Retargets the misc web parts on a specific file.  Loops through all web parts on the file.  Currently only
        /// <see cref="DataFormWebPart"/> and <see cref="ContentByQueryWebPart"/> are considered.
        /// </summary>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="listMap">The list map.</param>
        /// <param name="file">The file.</param>
        private static void RetargetMiscWebParts(SPWeb targetWeb, StringDictionary listMap, SPFile file)
        {
            if (file == null)
            {
                return; // This should never be the case.
            }
            if (!Utilities.EnsureAspx(file.Url, true, false))
                return; // We can only handle aspx and master pages.

            if (file.InDocumentLibrary && Utilities.IsCheckedOut(file.Item) && !Utilities.IsCheckedOutByCurrentUser(file.Item))
            {
                return; // The item is checked out by a different user so leave it alone.
            }

            bool fileModified = false;
            bool wasCheckedOut = true;

            SPLimitedWebPartManager manager = null;
            try
            {
                manager = targetWeb.GetLimitedWebPartManager(file.Url, PersonalizationScope.Shared);
                SPLimitedWebPartCollection webParts = manager.WebParts;

                for (int i = 0; i < webParts.Count; i++)
                {
                    WebPart webPart = null;
                    try
                    {
                        webPart = webParts[i] as WebPart;
                        if (webPart == null)
                        {
                            webParts[i].Dispose();
                            continue;
                        }

#if MOSS
                        if (!(webPart is ContentByQueryWebPart || webPart is DataFormWebPart))
                        {
                            webPart.Dispose();
                            continue;
                        }
#else
                        if (!(webPart is DataFormWebPart))
                        {
                            webPart.Dispose();
                            continue;
                        }
#endif
                        foreach (string sourceId in listMap.Keys)
                        {
                            bool modified = false;

                            Common.WebParts.ReplaceWebPartContent.Settings settings = new Common.WebParts.ReplaceWebPartContent.Settings();
                            settings.Test = false;
                            settings.SearchString = string.Format("(?i:{0})", sourceId);
                            settings.ReplaceString = listMap[sourceId];
                            settings.Publish = true;
                            settings.UnsafeXml = true;

                            Regex regex = new Regex(settings.SearchString);

                            // As every web part has different requirements we are only going to consider a small subset.
                            // Custom web parts will not be addressed as there's no interface that can be utilized.
#if MOSS
                            if (webPart is ContentByQueryWebPart)
                            {
                                ContentByQueryWebPart wp = (ContentByQueryWebPart)webPart;
                                webPart = Common.WebParts.ReplaceWebPartContent.ReplaceValues(
                                    targetWeb,
                                    file,
                                    settings,
                                    wp,
                                    regex,
                                    ref manager,
                                    ref wasCheckedOut,
                                    ref modified);
                            }
                            else 
#endif
                            if (webPart is DataFormWebPart)
                            {
                                DataFormWebPart wp = (DataFormWebPart)webPart;
                                webPart = Common.WebParts.ReplaceWebPartContent.ReplaceValues(
                                    targetWeb,
                                    file,
                                    settings,
                                    wp,
                                    regex,
                                    ref manager,
                                    ref wasCheckedOut,
                                    ref modified);
                            }

                            if (modified)
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
                if (fileModified)
                    file.CheckIn(
                        "Checking in changes to list item due to retargetting of web part as a result of converting a sub-site to a site collection.");

                if (file.InDocumentLibrary && fileModified && !wasCheckedOut)
                {
                    Lists.PublishItems itemPublisher = new Lists.PublishItems();
                    itemPublisher.PublishListItem(file.Item, file.Item.ParentList, false,
                        "\"Command-line fix of SPSite.\"", null, null);
                }
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

#if MOSS

        #region Set Global Navigation

        /// <summary>
        /// Sets the global navigation.
        /// </summary>
        /// <param name="targetPublishingWeb">The target publishing web.</param>
        private static void SetGlobalNavigation(PublishingWeb targetPublishingWeb)
        {
            SPNavigationNodeCollection globalNodes = targetPublishingWeb.Navigation.GlobalNavigationNodes;
            SPNavigationNodeCollection currentNodes = targetPublishingWeb.Navigation.CurrentNavigationNodes;

            if (globalNodes.Count > 0)
            {
                return;
            }

            SetGlobalNavigationRecursiveHelper(currentNodes, globalNodes);
        }

        /// <summary>
        /// Recursive routine for setting the global navigation (necessary to handle child elements).
        /// </summary>
        /// <param name="currentNodes">The current nodes.</param>
        /// <param name="globalNodes">The global nodes.</param>
        private static void SetGlobalNavigationRecursiveHelper(SPNavigationNodeCollection currentNodes, SPNavigationNodeCollection globalNodes)
        {
            // We're going to copy the current nodes in order to set the global nodes (so use current as the default)
            for (int i = 0; i < currentNodes.Count; i++)
            {
                SPNavigationNode currentNode = currentNodes[i];
                NodeTypes type = NodeTypes.None;
                if (currentNode.Properties["NodeType"] != null)
                    type = (NodeTypes)Enum.Parse(typeof(NodeTypes), (string)currentNode.Properties["NodeType"]);

                SPNavigationNode node = SPNavigationSiteMapNode.CreateSPNavigationNode(
                    currentNode.Title, currentNode.Url, type, globalNodes);
                foreach (DictionaryEntry de in currentNode.Properties)
                {
                    node.Properties[de.Key] = de.Value;
                }
                node.Update();
                node.MoveToLast(globalNodes);

                if (currentNode.Children.Count > 0)
                    SetGlobalNavigationRecursiveHelper(currentNode.Children, node.Children);
            }
        }

        #endregion


        /// <summary>
        /// Fixes the publishing pages.
        /// </summary>
        /// <param name="targetSite">The target site.</param>
        private static void FixPublishingPages(SPSite targetSite)
        {
            Logger.Write("Progress: Begin fixing publishing pages...");
            try
            {
                foreach (SPWeb web in targetSite.AllWebs)
                {
                    try
                    {
                        if (!PublishingWeb.IsPublishingWeb(web))
                            continue;

                        PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                        Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb);
                    }
                    finally
                    {
                        web.Dispose();
                    }
                }
            }
            finally
            {
                Logger.Write("Progress: End fixing publishing pages.");
            }
        }

        /// <summary>
        /// Fixes the page layouts and sets the site templates.
        /// </summary>
        /// <param name="sourcePublishingSite">The source publishing site.</param>
        /// <param name="sourcePublishingWeb">The source publishing web.</param>
        /// <param name="targetPublishingSite">The target publishing site.</param>
        /// <param name="targetPublishingWeb">The target publishing web.</param>
        /// <param name="targetSite">The target site.</param>
        /// <param name="targetWeb">The target web.</param>
        private static void FixPageLayoutsAndSiteTemplates(PublishingSite sourcePublishingSite, PublishingWeb sourcePublishingWeb, PublishingSite targetPublishingSite, PublishingWeb targetPublishingWeb, SPSite targetSite, SPWeb targetWeb)
        {
            try
            {
                Logger.Write("Progress: Begin fixing page layouts and site templates...");
                // Next thing we need to do is reset the "__PageLayouts" property - after doing the import
                // this value (targetWeb.AllProperties["__PageLayouts"]) will equal "__inherit" but it should
                // be either an empty string or an xml list of layouts.  Fixing this resolves the xml error
                // that we receive when going to "targetSite/_layouts/AreaTemplateSettings.aspx":
                // "Data at the root level is invalid. Line 1, position 1".  This error occurs when GetAvailablePageLayouts()
                // is called - within that method there's call to get to GetEffectiveAvailablePageLayoutsAsString()
                // which in this case returns back "__inherit" but should return back "" - as a result the next
                // statement after that is a call to IsPropertyAllowingAll - this returns back false because it's not
                // an empty string - as a result an XmlDocument.LoadXml() call is made which fails because "__inherit"
                // is not valid xml.

                // GL 2/7/2010: Change from SP2007 to SP2010 - Use __DefaultPageLayout instead of __PageLayouts.
                if (!string.IsNullOrEmpty(targetWeb.AllProperties["__DefaultPageLayout"] as string))
                {
                    Logger.Write(
                        "Progress: Setting \"__PageLayouts\" web property to string.Empty (current value is \"{0}\")...",
                        targetWeb.AllProperties["__DefaultPageLayout"] as string);
                    targetWeb.AllProperties["__DefaultPageLayout"] = string.Empty;
                    // We need to update the web so that the changes above stick and the following code can execute.
                    targetWeb.Update();
                }

                // Reset to make sure everything gets propagated correctly. (Note - this may need to be moved to after we copy the page layouts)
                Logger.Write("Progress: Setting available page layouts...");
                if (sourcePublishingWeb.IsAllowingAllPageLayouts)
                    targetPublishingWeb.AllowAllPageLayouts(true);
                else
                    targetPublishingWeb.SetAvailablePageLayouts(sourcePublishingWeb.GetAvailablePageLayouts(), true);

                // Reset the site templates settings.
                if (sourcePublishingWeb.IsAllowingAllWebTemplates)
                {
                    Logger.Write("Progress: Allowing all site templates...");
                    targetPublishingWeb.AllowAllWebTemplates(true);
                }
                else
                {
                    Logger.Write("Progress: Setting explicit site template settings...");
                    // Handle site templates set explicitly.
                    Collection<SPWebTemplate> list = new Collection<SPWebTemplate>();
                    foreach (SPWebTemplate template in sourcePublishingWeb.GetAvailableCrossLanguageWebTemplates())
                    {
                        list.Add(template);
                    }
                    targetPublishingWeb.SetAvailableCrossLanguageWebTemplates(list, true);

                    foreach (SPLanguage lang in sourcePublishingWeb.Web.RegionalSettings.InstalledLanguages)
                    {
                        list = new Collection<SPWebTemplate>();
                        foreach (
                            SPWebTemplate template in sourcePublishingWeb.GetAvailableWebTemplates((uint) lang.LCID))
                        {
                            list.Add(template);
                        }
                        if (list.Count > 0)
                            targetPublishingWeb.SetAvailableWebTemplates(list, (uint) lang.LCID, true);
                    }
                }

                Logger.Write("Progress: Setting page layouts...");
                foreach (PageLayout layout in sourcePublishingSite.GetPageLayouts(false))
                {
                    // We need to reset all the page layouts to match that of the source site
                    // (after the import all the layouts are messed up and treated as plain files and not page layouts
                    // because the content type is no longer associated).
                    PageLayoutCollection targetSiteLayouts = targetPublishingSite.GetPageLayouts(false);
                    PageLayout tempPageLayout = null;

                    try
                    {
                        tempPageLayout = targetSiteLayouts[
                            targetPublishingSite.PageLayouts.LayoutsDocumentLibrary.RootFolder.ServerRelativeUrl.
                                TrimStart('/') + "/" + layout.Name];
                    }
                    catch (ArgumentException)
                    {
                    }

                    if (tempPageLayout == null)
                    {
                        // We didn't find an item in the collection so let's attempt to add it back.
                        Logger.Write("Progress: Source layout {0} not found in target layout collection, searching files...",
                            layout.Name);
                        // First we need to see if the file is already there.
                        SPFile file = null;
                        foreach (SPListItem item in targetPublishingSite.PageLayouts.LayoutsDocumentLibrary.Items)
                        {
                            if (item.Name == layout.Name)
                            {
                                // We found the file so exit out of the loop.
                                file = item.File;
                                Logger.Write("Progress: Layout file found.");
                                break;
                            }
                        }

                        if (file == null)
                        {
                            Logger.Write("Progress: Layout file Not found.");
                            SPList targetMasterGallery =
                                targetSite.RootWeb.GetCatalog(SPListTemplateType.MasterPageCatalog);
                            SPFolder targetMasterGalleryFolder = targetMasterGallery.RootFolder;

                            try
                            {
                                Logger.Write("Progress: Copying page layout {0} from source site...", layout.Name);
                                // We couldn't find a file so copy the file from the source.
                                file = targetMasterGalleryFolder.Files.Add(layout.Name, layout.ListItem.File.OpenBinary());
                            }
                            catch (Exception ex)
                            {
                                Logger.WriteException(new System.Management.Automation.ErrorRecord(new SPException("Unable to copy page layout from source location", ex),
                                    null, System.Management.Automation.ErrorCategory.NotSpecified, layout));
                                continue;
                            }
                        }
                        if (file == null)
                        {
                            Logger.WriteWarning("Unable to copy page layout from source location.");
                            continue;
                        }

                        // Reset all the properties on the file so that it's flagged as a PageLayout content type.
                        Logger.Write("Progress: Setting page layout properties...");
                        SPListItem listItem = file.Item;
                        if (layout.ListItem.Fields.Contains(FieldId.Hidden) && listItem.Fields.Contains(FieldId.Hidden))
                        {
                            try
                            {
                                listItem[FieldId.Hidden] = layout.ListItem[FieldId.Hidden];
                            }
                            catch (ArgumentException) {}
                        }
                        if (layout.ListItem.Fields.Contains(FieldId.Title) && listItem.Fields.Contains(FieldId.Title))
                            listItem[FieldId.Title] = layout.ListItem[FieldId.Title];

                        if (layout.ListItem.Fields.Contains(FieldId.ContentType) && listItem.Fields.Contains(FieldId.ContentType))
                            listItem[FieldId.ContentType] = GetContentType(targetSite.RootWeb, ContentTypeId.PageLayout).Name;

                        if (layout.ListItem.Fields.Contains(FieldId.AssociatedContentType) && listItem.Fields.Contains(FieldId.AssociatedContentType))
                            listItem[FieldId.AssociatedContentType] = new ContentTypeIdFieldValue(GetContentType(targetSite.RootWeb,
                                                                       layout.AssociatedContentType.Id));

                        if (layout.ListItem.Fields.Contains(FieldId.AssociatedVariations) && listItem.Fields.Contains(FieldId.AssociatedVariations))
                            listItem[FieldId.AssociatedVariations] = layout.ListItem[FieldId.AssociatedVariations];

                        if (layout.ListItem.Fields.Contains(SPBuiltInFieldId.File_x0020_Type) && listItem.Fields.Contains(SPBuiltInFieldId.File_x0020_Type))
                        {
                            try
                            {
                                listItem[SPBuiltInFieldId.File_x0020_Type] = layout.ListItem[SPBuiltInFieldId.File_x0020_Type];
                            }
                            catch (ArgumentException) {}
                        }
                        // Add any additional properties specific to the page layout
                        PageLayout newLayout = new PageLayout(listItem);
                        newLayout.Description = layout.Description;
                        newLayout.Title = layout.Title;
                        if (!string.IsNullOrEmpty(layout.PreviewImageUrl))
                        {
                            string previewImageUrl =
                                targetWeb.ServerRelativeUrl +
                                layout.PreviewImageUrl.Substring(
                                    layout.PreviewImageUrl.IndexOf("/_catalogs/"));
                            newLayout.PreviewImageUrl = targetSite.MakeFullUrl(previewImageUrl);
                        }
                        listItem.SystemUpdate();

                        if (Utilities.IsCheckedOut(file.Item))
                        {
                            Logger.Write("Progress: Checking in page layout file...");
                            file.CheckIn("", SPCheckinType.MajorCheckIn);
                            if (file.Item.ModerationInformation != null)
                                file.Approve("");
                        }
                    }
                }
            }
            finally
            {
                Logger.Write("Progress: End fixing page layouts and site templates.");
            }
        }
#endif

        /// <summary>
        /// Adds the missing features.
        /// </summary>
        /// <param name="sourceSite">The source site.</param>
        /// <param name="sourceWeb">The source web.</param>
        /// <param name="targetSite">The target site.</param>
        /// <param name="targetWeb">The target web.</param>
        public static void AddMissingFeatures(SPSite sourceSite, SPWeb sourceWeb, SPSite targetSite, SPWeb targetWeb)
        {
            Logger.Write("Progress: Begin adding missing features...");
            try
            {
                // Set any features that need to be enabled.
                Dictionary<SPFeatureScope, SPFeatureCollection> activeFeatures =
                    new Dictionary<SPFeatureScope, SPFeatureCollection>();
                if (sourceSite != null)
                {
                    activeFeatures[SPFeatureScope.WebApplication] = sourceSite.WebApplication.Features;
                    activeFeatures[SPFeatureScope.Site] = sourceSite.Features;
                }
                if (sourceWeb != null)
                    activeFeatures[SPFeatureScope.Web] = sourceWeb.Features;

                // Note that you should be able to use the ActivationDependencies property of the SPDefinition,
                // however, I found that this property is not reliable (good example: Publishing is dependent
                // on PublishingSite but when you view the ActivationDependencies for the Publishing definition
                // it shows zero dependencies.
                Queue<SPFeatureDefinition> queuedFeatures = new Queue<SPFeatureDefinition>();
                // For some reason these need to be added before the rest...
                if (SPFarm.Local.FeatureDefinitions["BaseSite"] != null)
                    queuedFeatures.Enqueue(SPFarm.Local.FeatureDefinitions["BaseSite"]);
                if (SPFarm.Local.FeatureDefinitions["PremiumSite"] != null)
                    queuedFeatures.Enqueue(SPFarm.Local.FeatureDefinitions["PremiumSite"]);
                if (SPFarm.Local.FeatureDefinitions["PublishingSite"] != null)
                    queuedFeatures.Enqueue(SPFarm.Local.FeatureDefinitions["PublishingSite"]);

                foreach (SPFeatureDefinition definition in SPFarm.Local.FeatureDefinitions)
                {
                    try
                    {
                        if (definition.Scope == SPFeatureScope.Farm)
                            continue;

                        if (!queuedFeatures.Contains(definition))
                            queuedFeatures.Enqueue(definition);
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteException(new System.Management.Automation.ErrorRecord(ex, null, System.Management.Automation.ErrorCategory.NotSpecified, definition));
                    }
                }
                while (queuedFeatures.Count > 0)
                {
                    SPFeatureDefinition definition = queuedFeatures.Dequeue();
                    if (definition == null)
                        continue;

                    SPFeatureScope scope = SPFeatureScope.ScopeInvalid;
                    try
                    {
                        scope = definition.Scope;
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteException(new System.Management.Automation.ErrorRecord(ex, null, System.Management.Automation.ErrorCategory.NotSpecified, definition));
                        continue;
                    }
                    Guid featureID = definition.Id;
                    if (activeFeatures[scope] != null)
                    {
                        bool isActive = (activeFeatures[scope][featureID] != null);
                        if (!isActive)
                            continue;

                        try
                        {
                            switch (scope)
                            {
                                case SPFeatureScope.Site:
                                    if (targetSite != null && targetSite.Features[featureID] == null)
                                    {
                                        Logger.Write("Progress: Activating site scoped feature \"{0}\"...", definition.DisplayName);
                                        targetSite.Features.Add(featureID);
                                    }
                                    break;
                                case SPFeatureScope.Web:
                                    if (targetWeb != null && targetWeb.Features[featureID] == null)
                                    {
                                        Logger.Write("Progress: Activating web scoped feature \"{0}\"...", definition.DisplayName);
                                        targetWeb.Features.Add(featureID);
                                    }
                                    break;
                                case SPFeatureScope.WebApplication:
                                    if (sourceSite != null && sourceSite.WebApplication.Features[featureID] == null)
                                    {
                                        Logger.Write("Progress: Activating web application scoped feature \"{0}\"...", definition.DisplayName);
                                        sourceSite.WebApplication.Features.Add(featureID);
                                    }
                                    break;
                                default:
                                    continue;
                            }
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            // We couldn't add the item most likely due a dependent content type that has not yet been added.
                            queuedFeatures.Enqueue(definition);
                        }
                        catch (InvalidOperationException)
                        {
                            // We couldn't add the item most likely due to dependencies so add to the back of the queue.
                            queuedFeatures.Enqueue(definition);
                        }
                        catch (ArgumentException)
                        {
                            // We couldn't add the item most likely due to dependencies so add to the back of the queue.
                            queuedFeatures.Enqueue(definition);
                        }
                        catch (DuplicateNameException)
                        {
                            
                        }
                        catch (SPException ex)
                        {
                            // This can occur occasionally if the feature does some external modifications (should be rare)
                            if (ex.Message == "The web being updated was changed by an external process.")
                            {
                                if (targetWeb != null)
                                    targetWeb.Close();
                                if (targetSite != null)
                                    targetWeb = targetSite.OpenWeb();
                                queuedFeatures.Enqueue(definition);
                            }
                            else
                            {
                                Logger.WriteWarning("Unable to activate feature '{0} ({1})'\r\n{2}",
                                                  definition.DisplayName, definition.Name, ex.Message);
                            }
                        }
                    }
                }
            }
            finally
            {
                Logger.Write("Progress: End adding missing features.");
            }
        }


        /// <summary>
        /// Sets the master page gallery settings.
        /// </summary>
        /// <param name="sourceSite">The source site.</param>
        /// <param name="targetSite">The target site.</param>
        /// <param name="targetWeb">The target web.</param>
        internal static void SetMasterPageGallerySettings(SPSite sourceSite, SPSite targetSite, SPWeb targetWeb)
        {
            Logger.Write("Progress: Begin setting master page gallery settings...");
            try
            {
                SPList targetMasterGallery = targetSite.RootWeb.GetCatalog(SPListTemplateType.MasterPageCatalog);
                SPList sourceMasterGallery = sourceSite.RootWeb.GetCatalog(SPListTemplateType.MasterPageCatalog);

                // Need to make sure the master gallery has it's content types enabled (get's disabled during the import) 
                // and make sure all other settings are set to match the source.
                Logger.Write("Progress: Setting versioning properties...");
                targetMasterGallery.ContentTypesEnabled = true;
                targetMasterGallery.EnableModeration = sourceMasterGallery.EnableModeration;
                targetMasterGallery.EnableVersioning = sourceMasterGallery.EnableVersioning;
                targetMasterGallery.EnableMinorVersions = sourceMasterGallery.EnableMinorVersions;
                targetMasterGallery.MajorVersionLimit = sourceMasterGallery.MajorVersionLimit;
                try
                {
                    targetMasterGallery.MajorWithMinorVersionsLimit = sourceMasterGallery.MajorWithMinorVersionsLimit;
                }
                catch (NotSupportedException)
                {
                    // If we're here then something is wrong with the source list so just ignore the error.
                }
                targetMasterGallery.DraftVersionVisibility = sourceMasterGallery.DraftVersionVisibility;
                targetMasterGallery.ForceCheckout = sourceMasterGallery.ForceCheckout;

                targetMasterGallery.Update();

                SPField contentTypeField = targetMasterGallery.Fields.GetFieldByInternalName("ContentType");
                if (contentTypeField.Type != SPFieldType.Choice)
                {
                    Logger.Write("Progress: ContentType field must be a Choice column (currently is {0})", contentTypeField.TypeAsString);
                    if (contentTypeField.Type == SPFieldType.Text)
                    {
#if MOSS
                        if (PublishingWeb.IsPublishingWeb(targetWeb))
                        {
                            Logger.Write("Progress: ContentType field is a Text type in a publishing web, attempting to activate PublishingResources to correct...");
                            Guid publishingResourcesFeatureId = new Guid("AEBC918D-B20F-4a11-A1DB-9ED84D79C87E");
                            FeatureHelper fh = new FeatureHelper();
                            fh.ActivateDeactivateFeatureAtSite(targetSite, true, publishingResourcesFeatureId, true, false);

                            // Get the list again to make sure we're not dealing with a cached copy
                            targetMasterGallery = targetSite.RootWeb.GetCatalog(SPListTemplateType.MasterPageCatalog);
                            contentTypeField = targetMasterGallery.Fields.GetFieldByInternalName("ContentType");
                            if (contentTypeField.Type != SPFieldType.Choice)
                            {
                                Logger.Write("Progress: Failed to fix ContentType field via PublishingResources feature activation, attempting to copy entire gallery from source...");

                                Common.Lists.ImportList importList = new Common.Lists.ImportList(
                                    sourceSite.MakeFullUrl(sourceMasterGallery.RootFolder.ServerRelativeUrl),
                                    targetSite.Url, false);

                                importList.Copy(null, true, 0, false, true, true, false, SPIncludeVersions.All, SPUpdateVersions.Overwrite, true, false, false, false, false, SPIncludeDescendants.All, false, false, false);
                            }
                        }
#endif
                        /*****  The code below was necessary prior to SP2 (can't be sure of the exact update)
                     *      Activating the publishingresources feature seems to now resolve the issue
                     *      that this code was fixing.
                     *      
                     *
                    // Need to reset the content type field – after the import it gets changed to a 
                    // Text field but it needs to be a Choice field otherwise no Page Layout objects 
                    // will be returned.

                    string childXml =
                        string.Format(@"<Default>{0}</Default>
                          <CHOICES>
                            <CHOICE>{0}</CHOICE>
                            <CHOICE>{1}</CHOICE>
                            <CHOICE>{2}</CHOICE>
                            <CHOICE>{3}</CHOICE>
                          </CHOICES>",
                                     SPUtility.GetLocalizedString("$Resources:cmscore,contenttype_pagelayout_name;", null, (uint)targetWeb.Locale.LCID),
                                     SPUtility.GetLocalizedString("$Resources:cmscore,contenttype_masterpage_name;", null, (uint)targetWeb.Locale.LCID),
                                     SPUtility.GetLocalizedString("$Resources:core,MasterPage", null, (uint)targetWeb.Locale.LCID),
                                     SPUtility.GetLocalizedString("$Resources:core,Folder", null, (uint)targetWeb.Locale.LCID));

                    using (SqlConnection connection = new SqlConnection(targetMasterGallery.ParentWeb.Site.ContentDatabase.DatabaseConnectionString))
                    {
                        try
                        {
                            string sql = "select tp_fields from alllists where tp_id=@listID";
                            SqlCommand sqlCommand = new SqlCommand(sql, connection);
                            connection.Open();
                            sqlCommand.Parameters.AddWithValue("listID", targetMasterGallery.ID.ToString());
                            string sourceXml = (string)sqlCommand.ExecuteScalar();
                            string header = sourceXml.Substring(0, sourceXml.IndexOf('<'));
                            sourceXml = "<Fields>" + sourceXml.Substring(sourceXml.IndexOf('<')) + "</Fields>";

                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.LoadXml(sourceXml);
                            XmlElement fieldElement = (XmlElement)xmlDoc.SelectSingleNode("//Field[@ID='{" + contentTypeField.Id + "}']");
                            fieldElement.InnerXml = childXml;
                            fieldElement.SetAttribute("RowOrdinal", "0");
                            fieldElement.SetAttribute("Type", "Choice");
                            fieldElement.SetAttribute("Format", "Dropdown"); 
                            fieldElement.SetAttribute("FillInChoice", "FALSE"); 
                            fieldElement.SetAttribute("Sealed", "FALSE"); 
                            fieldElement.SetAttribute("Name", "ContentType"); 
                            fieldElement.SetAttribute("ColName", "tp_ContentType"); 
                            fieldElement.SetAttribute("SourceID", "http://schemas.microsoft.com/sharepoint/v3"); 
                            fieldElement.SetAttribute("ID", "{c042a256-787d-4a6f-8a8a-cf6ab767f12d}");
                            fieldElement.SetAttribute("DisplayName",
                                                      SPUtility.GetLocalizedString("$Resources:core,Content_Type;", null,
                                                                                   (uint) targetWeb.Locale.LCID));
                            fieldElement.SetAttribute("StaticName", "ContentType"); 
                            fieldElement.SetAttribute("Group", "_Hidden"); 
                            fieldElement.SetAttribute("PITarget", "MicrosoftWindowsSharePointServices"); 
                            fieldElement.SetAttribute("PIAttribute", "ContentTypeID");

                            sourceXml = header + xmlDoc.DocumentElement.InnerXml;

                            sql = "update alllists set tp_fields=@fieldXml where tp_id=@listID";
                            sqlCommand = new SqlCommand(sql, connection);
                            sqlCommand.Parameters.AddWithValue("listID", targetMasterGallery.ID.ToString());
                            sqlCommand.Parameters.AddWithValue("fieldXml", sourceXml);
                            sqlCommand.ExecuteNonQuery();
                        }
                        finally
                        {
                            if (connection.State != ConnectionState.Closed)
                                connection.Close();
                        }
                    }
                     * */
                    }
                }
            }
            finally
            {
                Logger.Write("Progress: End setting master page gallery settings.");
            }
        }
        /*
        #region Fix Discussion Lists

        /// <summary>
        /// Repairs the discussion lists for all webs belonging to the site collection.
        /// </summary>
        /// <param name="site">The site.</param>
        internal static void RepairDiscussionLists(SPSite site)
        {
            foreach (SPWeb web in site.AllWebs)
            {
                try
                {
                    RepairDiscussionLists(site, web);
                }
                finally
                {
                    web.Dispose();
                }
            }
        }

        /// <summary>
        /// Repairs the discussion list for all lists belonging to the web.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        internal static void RepairDiscussionLists(SPSite site, SPWeb web)
        {
            foreach (SPList list in web.GetListsOfType(SPBaseType.DiscussionBoard))
            {
                RepairDiscussionList(site, list);
            }
        }

        /// <summary>
        /// Repairs the discussion list.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="list">The list.</param>
        internal static void RepairDiscussionList(SPSite site, SPList list)
        {
            if (list.ContentTypes["Discussion"] == null)
            {
                return;
            }
            foreach (SPListItem item in list.Items)
            {
                SPListItem folder = GetFolderById(list, (int)item["ParentFolderId"]);
                if (folder == null)
                {
                    Console.WriteLine("WARNING: Unable to find parent folder for '{0}'", item.Url);
                    continue;
                }
                string parentFolder = (string)folder["FileRef"];
                string fileRef = parentFolder + "/" + item["FileLeafRef"];

                if (fileRef == (string)item["FileRef"])
                {
                    continue;
                }


                Guid parentFolderGuid = folder.UniqueId;

                using (SqlConnection connection = new SqlConnection(site.ContentDatabase.DatabaseConnectionString))
                {
                    try
                    {
                        string sql = "select ItemChildCount from AllDocs where id=@itemID";
                        SqlCommand sqlCommand = new SqlCommand(sql, connection);
                        connection.Open();
                        sqlCommand.Parameters.AddWithValue("itemID", parentFolderGuid);
                        int currentCount = (int)sqlCommand.ExecuteScalar();
                        currentCount++;

                        
                        sql = "update AllDocs set DirName=@dir, ParentId=@parentId where Id=@itemID";
                        sqlCommand = new SqlCommand(sql, connection);
                        sqlCommand.Parameters.AddWithValue("dir", parentFolder.Trim('/'));
                        sqlCommand.Parameters.AddWithValue("parentId", parentFolderGuid);
                        sqlCommand.Parameters.AddWithValue("itemID", item.UniqueId);
                        sqlCommand.ExecuteNonQuery();


                        sql = "update AllDocs set ItemChildCount=@count where Id=@itemID";
                        sqlCommand = new SqlCommand(sql, connection);
                        sqlCommand.Parameters.AddWithValue("itemID", parentFolderGuid);
                        sqlCommand.Parameters.AddWithValue("count", currentCount);
                        sqlCommand.ExecuteNonQuery();


                        sql = "update AllUserData set tp_DirName=@dir where tp_ID=@itemID and tp_ListId=@listID and tp_SiteId=@siteID";
                        sqlCommand = new SqlCommand(sql, connection);
                        sqlCommand.Parameters.AddWithValue("dir", parentFolder.Trim('/'));
                        sqlCommand.Parameters.AddWithValue("itemID", item.ID);
                        sqlCommand.Parameters.AddWithValue("listID", list.ID);
                        sqlCommand.Parameters.AddWithValue("siteID", site.ID);
                        sqlCommand.ExecuteNonQuery();

                    }
                    finally
                    {
                        if (connection.State != ConnectionState.Closed)
                            connection.Close();
                    }
                }
            }
        }

        #endregion
        */
        #endregion



        #region Helper Methods

        /// <summary>
        /// Gets the type of the content.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="contentTypeId">The content type id.</param>
        /// <returns></returns>
        private static SPContentType GetContentType(SPWeb web, SPContentTypeId contentTypeId)
        {
            SPContentType type = web.AvailableContentTypes[contentTypeId];
            if (type == null)
            {
                throw new SPException("Content Type Not Found In Web " + contentTypeId + ", " + web.Url);
            }
            return type;
        }

        /// <summary>
        /// Gets the folder by id.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="id">The id.</param>
        /// <returns></returns>
        private static SPListItem GetFolderById(SPList list, int id)
        {
            foreach (SPListItem folder in list.Folders)
            {
                if (id == folder.ID)
                    return folder;
            }
            return null;
        }


        #endregion
    }
}
