using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Navigation;
using Microsoft.SharePoint.Publishing;
using Microsoft.SharePoint.Publishing.Navigation;

namespace Lapointe.SharePoint.PowerShell.Common.SiteCollections
{
    internal class ImportNavigation
    {
        public static void SetNavigation(SPSite site, XmlDocument xmlDoc, bool deleteExistingGlobal, bool deleteExistingCurrent)
        {
            SetNavigation(site.RootWeb, xmlDoc, deleteExistingGlobal, deleteExistingCurrent, true);

        }

        /// <summary>
        /// Sets the navigation.
        /// </summary>
        /// <param name="web">The web site.</param>
        /// <param name="xmlDoc">The XML doc containing the navigation nodes.</param>
        /// <param name="deleteExistingGlobal">if set to <c>true</c> [delete existing global nodes].</param>
        /// <param name="deleteExistingCurrent">if set to <c>true</c> [delete existing current nodes].</param>
        public static void SetNavigation(SPWeb web, XmlDocument xmlDoc, bool deleteExistingGlobal, bool deleteExistingCurrent, bool includeChildren)
        {
            PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

            AutoAddSubSites(web, xmlDoc);
            AutoAddSiteCollectionUrl(web.Site, xmlDoc);
            AutoAddWebUrl(web, xmlDoc);

            XmlElement webElement = xmlDoc.SelectSingleNode("//Web[translate(@serverRelativeUrl, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='" + web.ServerRelativeUrl.ToLower() + "']") as XmlElement;
            if (webElement == null)
                throw new Exception("Cannot find Web with Server Relative URL \"" + web.ServerRelativeUrl + "\".");

            XmlElement globalNode = webElement.SelectSingleNode("./Navigation/Global") as XmlElement;
            XmlElement currentNode = webElement.SelectSingleNode("./Navigation/Current") as XmlElement;


            // First need to set whether or not we show sub-sites and pages
            bool globalShowSubSites = pubweb.Navigation.GlobalIncludeSubSites;
            bool globalShowPages = pubweb.Navigation.GlobalIncludePages;
            bool currentShowSubSites = pubweb.Navigation.CurrentIncludeSubSites;
            bool currentShowPages = pubweb.Navigation.CurrentIncludePages;

            if (globalNode != null)
            {
                globalShowSubSites = bool.Parse(globalNode.GetAttribute("IncludeSubSites"));
                globalShowPages = bool.Parse(globalNode.GetAttribute("IncludePages")); ;

                pubweb.Navigation.GlobalIncludeSubSites = globalShowSubSites;
                pubweb.Navigation.GlobalIncludePages = globalShowPages;
            }
            if (currentNode != null)
            {
                currentShowSubSites = bool.Parse(currentNode.GetAttribute("IncludeSubSites"));
                currentShowPages = bool.Parse(currentNode.GetAttribute("IncludePages")); ;

                pubweb.Navigation.CurrentIncludeSubSites = currentShowSubSites;
                pubweb.Navigation.CurrentIncludePages = currentShowPages;
            }
            pubweb.Update();

            List<SPNavigationNode> existingGlobalNodes = new List<SPNavigationNode>();
            List<SPNavigationNode> existingCurrentNodes = new List<SPNavigationNode>();
            // We can't delete the navigation items until we've added the new ones so store the existing 
            // ones for later deletion (note that we don't have to store all of them - just the top level).
            // I have no idea why this is the case - but when I tried to clear everything out first I got
            // all kinds of funky errors that just made no sense to me - this works so....
            foreach (SPNavigationNode node in pubweb.Navigation.GlobalNavigationNodes)
                existingGlobalNodes.Add(node);
            foreach (SPNavigationNode node in pubweb.Navigation.CurrentNavigationNodes)
                existingCurrentNodes.Add(node);

            XmlNodeList newGlobalNodes = null;
            if (globalNode != null)
                newGlobalNodes = globalNode.SelectNodes("./Node");

            XmlNodeList newCurrentNodes = null;
            if (currentNode != null)
                newCurrentNodes = currentNode.SelectNodes("./Node");

            if (newGlobalNodes != null && newGlobalNodes.Count > 0)
            {
                pubweb.Navigation.InheritGlobal = false;
                pubweb.Update();
            }
            if (newCurrentNodes != null && newCurrentNodes.Count > 0)
            {
                pubweb.Navigation.InheritCurrent = false;
                pubweb.Update();
            }
            pubweb = PublishingWeb.GetPublishingWeb(web);

            // If we've got global or current nodes in the xml then the intent is to reset those elements.
            // If we've also specified to delete any existing elements then we need to first hide all the
            // sub-sites and pages (you can't delete them because they don't exist as a node).  Note that
            // we are only doing this if showSubSites is true - if it's false we don't see them so no point
            // in hiding them.  Any non-sub-site or non-page will be deleted after we've added the new nodes.
            foreach (SPWeb tempWeb in pubweb.Web.Webs)
            {
                try
                {
                    if (newGlobalNodes != null && newGlobalNodes.Count > 0 && deleteExistingGlobal && globalShowSubSites)
                    {
                        // Initialize the sub-sites (forces the provided XML to specify whether any should be visible)
                        pubweb.Navigation.ExcludeFromNavigation(true, tempWeb.ID);
                    }
                    if (newCurrentNodes != null && newCurrentNodes.Count > 0 && deleteExistingCurrent && currentShowSubSites)
                    {
                        pubweb.Navigation.ExcludeFromNavigation(false, tempWeb.ID);
                    }
                }
                finally
                {
                    tempWeb.Dispose();
                }
            }
            pubweb.Update();

            // Now we need to add all the global nodes (if any - if the collection is empty the following will just return and do nothing)
            AddNodes(pubweb, true, pubweb.Navigation.GlobalNavigationNodes, newGlobalNodes);
            // Update the web as the above may have made modifications
            pubweb.Update();

            // Now delete all the previously existing global nodes.
            if (newGlobalNodes != null && newGlobalNodes.Count > 0 && deleteExistingGlobal)
            {
                foreach (SPNavigationNode node in existingGlobalNodes)
                {
                    node.Delete();
                }
            }

            // Now we need to add all the current nodes (if any)
            AddNodes(pubweb, false, pubweb.Navigation.CurrentNavigationNodes, newCurrentNodes);
            // Update the web as the above may have made modifications
            pubweb.Update();

            // Now delete all the previously existing current nodes.
            if (newCurrentNodes != null && newCurrentNodes.Count > 0 && deleteExistingCurrent)
            {
                foreach (SPNavigationNode node in existingCurrentNodes)
                {
                    node.Delete();
                }
            }

            if (includeChildren)
            {
                foreach (SPWeb childWeb in web.Webs)
                    SetNavigation(childWeb, xmlDoc, deleteExistingGlobal, deleteExistingCurrent, true);
            }
        }

        /// <summary>
        /// Replaces any occurance of "SiteCollectionUrl" nodes with text corresponding to the 
        /// server relative url of the site collection.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="xmlDoc">The XML doc.</param>
        private static void AutoAddSiteCollectionUrl(SPSite site, XmlDocument xmlDoc)
        {
            XmlNodeList autoAddList = xmlDoc.SelectNodes("//SiteCollectionUrl");
            if (autoAddList.Count == 0)
                return;

            foreach (XmlElement autoAddElement in autoAddList)
            {
                XmlText textNode = xmlDoc.CreateTextNode(site.ServerRelativeUrl);
                autoAddElement.ParentNode.ReplaceChild(textNode, autoAddElement);
            }
        }

        /// <summary>
        /// Replaces any occurance of "WebUrl" nodes with text corresponding to the 
        /// server relative url of the web site.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="xmlDoc">The XML doc.</param>
        private static void AutoAddWebUrl(SPWeb web, XmlDocument xmlDoc)
        {
            XmlNodeList autoAddList = xmlDoc.SelectNodes("//WebUrl");
            if (autoAddList.Count == 0)
                return;

            foreach (XmlElement autoAddElement in autoAddList)
            {
                XmlText textNode = xmlDoc.CreateTextNode(web.ServerRelativeUrl);
                autoAddElement.ParentNode.ReplaceChild(textNode, autoAddElement);
            }
        }

        /// <summary>
        /// Replaces any occurance of "AutoAddSubSites" nodes with nodes corresponding to each of the sub-sites for the web.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="xmlDoc">The XML doc.</param>
        private static void AutoAddSubSites(SPWeb web, XmlDocument xmlDoc)
        {
            XmlNodeList autoAddList = xmlDoc.SelectNodes("//AutoAddSubSites");
            if (autoAddList.Count == 0)
                return;

            foreach (XmlElement autoAddElement in autoAddList)
            {
                foreach (SPWeb childWeb in web.Webs)
                {
                    try
                    {
                        XmlElement navNode = xmlDoc.CreateElement("Node");
                        navNode.SetAttribute("Title", childWeb.Title);
                        navNode.SetAttribute("IsVisible", "True");

                        XmlElement prop = xmlDoc.CreateElement("Url");
                        prop.InnerText = childWeb.ServerRelativeUrl;
                        navNode.AppendChild(prop);

                        prop = xmlDoc.CreateElement("vti_navsequencechild");
                        prop.InnerText = "true";
                        navNode.AppendChild(prop);

                        prop = xmlDoc.CreateElement("NodeType");
                        prop.InnerText = "Area";
                        navNode.AppendChild(prop);

                        prop = xmlDoc.CreateElement("Description");
                        prop.InnerText = childWeb.Description;
                        navNode.AppendChild(prop);

                        autoAddElement.ParentNode.AppendChild(navNode);
                    }
                    finally
                    {
                        childWeb.Dispose();
                    }
                }
                autoAddElement.ParentNode.RemoveChild(autoAddElement);
            }
        }


        /// <summary>
        /// Adds the nodes.
        /// </summary>
        /// <param name="pubWeb">The publishing web.</param>
        /// <param name="isGlobal">if set to <c>true</c> [is global].</param>
        /// <param name="existingNodes">The existing nodes.</param>
        /// <param name="newNodes">The new nodes.</param>
        private static void AddNodes(PublishingWeb pubWeb, bool isGlobal, SPNavigationNodeCollection existingNodes, XmlNodeList newNodes)
        {
            if (newNodes.Count == 0)
                return;

            for (int i = 0; i < newNodes.Count; i++)
            {
                XmlElement newNodeXml = (XmlElement)newNodes[i];
                string url = newNodeXml.SelectSingleNode("Url").InnerText;
                string title = newNodeXml.GetAttribute("Title");
                NodeTypes type = NodeTypes.None;
                if (newNodeXml.SelectSingleNode("NodeType") != null && !string.IsNullOrEmpty(newNodeXml.SelectSingleNode("NodeType").InnerText))
                    type = (NodeTypes)Enum.Parse(typeof(NodeTypes), newNodeXml.SelectSingleNode("NodeType").InnerText);

                bool isVisible = true;
                if (!string.IsNullOrEmpty(newNodeXml.GetAttribute("IsVisible")))
                    isVisible = bool.Parse(newNodeXml.GetAttribute("IsVisible"));

                if (type == NodeTypes.Area)
                {
                    // You can't just add an "Area" node (which represents a sub-site) to the current web if the
                    // url does not correspond with an actual sub-site (the code will appear to work but you won't
                    // see anything when you load the page).  So we need to check and see if the node actually
                    // points to a sub-site - if it does not then change it to "AuthoredLinkToWeb".
                    SPWeb web = null;
                    try
                    {
                        string name = url.Trim('/');
                        if (name.Length != 0 && name.IndexOf("/") > 0)
                        {
                            name = name.Substring(name.LastIndexOf('/') + 1);
                        }
                        try
                        {
                            // pubWeb.Web.Webs[] does not return null if the item doesn't exist - it simply throws an exception (I hate that!)
                            web = pubWeb.Web.Webs[name];
                        }
                        catch (ArgumentException)
                        {
                        }
                        if (web == null || !web.Exists || web.ServerRelativeUrl.ToLower() != url.ToLower())
                        {
                            // The url doesn't correspond with a sub-site for the current web so change the node type.
                            // This is most likely due to copying navigation elements from another site
                            type = NodeTypes.AuthoredLinkToWeb;
                        }
                        else if (web.Exists && web.ServerRelativeUrl.ToLower() == url.ToLower())
                        {
                            // We did find a matching sub-site so now we need to set the visibility
                            if (isVisible)
                                pubWeb.Navigation.IncludeInNavigation(isGlobal, web.ID);
                            else
                                pubWeb.Navigation.ExcludeFromNavigation(isGlobal, web.ID);
                        }
                    }
                    finally
                    {
                        if (web != null)
                            web.Dispose();
                    }

                }
                else if (type == NodeTypes.Page)
                {
                    // Adding links to pages has the same limitation as sub-sites (Area nodes) so we need to make
                    // sure it actually exists and if it doesn't then change the node type.
                    PublishingPage page = null;
                    try
                    {
                        // Note that GetPublishingPages()[] does not return null if the item doesn't exist - it simply throws an exception (I hate that!)
                        page = pubWeb.GetPublishingPages()[url];
                    }
                    catch (ArgumentException)
                    {
                    }
                    if (page == null)
                    {
                        // The url doesn't correspond with a page for the current web so change the node type.
                        // This is most likely due to copying navigation elements from another site
                        type = NodeTypes.AuthoredLinkToPage;
                        url = pubWeb.Web.Site.MakeFullUrl(url);
                    }
                    else
                    {
                        // We did find a matching page so now we need to set the visibility
                        if (isVisible)
                            pubWeb.Navigation.IncludeInNavigation(isGlobal, page.ListItem.UniqueId);
                        else
                            pubWeb.Navigation.ExcludeFromNavigation(isGlobal, page.ListItem.UniqueId);
                    }
                }

                // If it's not a sub-site or a page that's part of the current web and it's set to
                // not be visible then just move on to the next (there is no visibility setting for
                // nodes that are not of type Area or Page).
                if (!isVisible && type != NodeTypes.Area && type != NodeTypes.Page)
                    continue;

                // Finally, can add the node to the collection.
                SPNavigationNode node = SPNavigationSiteMapNode.CreateSPNavigationNode(
                    title, url, type, existingNodes);


                // Now we need to set all the other properties
                foreach (XmlElement property in newNodeXml.ChildNodes)
                {
                    // We've already set these so don't set them again.
                    if (property.Name == "Url" || property.Name == "Node" || property.Name == "NodeType")
                        continue;

                    // CreatedDate and LastModifiedDate need to be the correct type - all other properties are strings
                    if (property.Name == "CreatedDate" && !string.IsNullOrEmpty(property.InnerText))
                    {
                        node.Properties["CreatedDate"] = DateTime.Parse(property.InnerText);
                        continue;
                    }
                    if (property.Name == "LastModifiedDate" && !string.IsNullOrEmpty(property.InnerText))
                    {
                        node.Properties["LastModifiedDate"] = DateTime.Parse(property.InnerText);
                        continue;
                    }

                    node.Properties[property.Name] = property.InnerText;
                }
                // If we didn't have a CreatedDate or LastModifiedDate then set them to now.
                if (node.Properties["CreatedDate"] == null)
                    node.Properties["CreatedDate"] = DateTime.Now;
                if (node.Properties["LastModifiedDate"] == null)
                    node.Properties["LastModifiedDate"] = DateTime.Now;

                // Save our changes to the node.
                node.Update();
                node.MoveToLast(existingNodes); // Should already be at the end but I prefer to make sure :)

                XmlNodeList childNodes = newNodeXml.SelectNodes("Node");

                // If we have child nodes then make a recursive call passing in the current nodes Children property as the collection to add to.
                if (childNodes.Count > 0)
                    AddNodes(pubWeb, isGlobal, node.Children, childNodes);
            }

        }
    }
}
