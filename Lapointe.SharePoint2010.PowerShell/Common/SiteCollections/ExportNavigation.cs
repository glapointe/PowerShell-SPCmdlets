using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Navigation;
using Microsoft.SharePoint.Publishing;

namespace Lapointe.SharePoint.PowerShell.Common.SiteCollections
{
    internal class ExportNavigation
    {
        public static XmlDocument GetNavigation(SPSite site)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement rootNode = (XmlElement)xmlDoc.AppendChild(xmlDoc.CreateElement("Site"));
            rootNode.SetAttribute("url", site.Url);
            GetNavigation(xmlDoc, site.RootWeb, true);
            return xmlDoc;
        }

        public static XmlDocument GetNavigation(SPWeb web, bool includeChildren)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement rootNode = (XmlElement)xmlDoc.AppendChild(xmlDoc.CreateElement("Web"));
            rootNode.SetAttribute("url", web.Url);
            GetNavigation(xmlDoc, web, includeChildren);

            return xmlDoc;
        }

        /// <summary>
        /// Gets the navigation.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="scope">The scope.</param>
        /// <returns></returns>
        public static XmlDocument GetNavigation(string url, string scope)
        {
            XmlDocument xmlDoc = new XmlDocument();
            using (SPSite site = new SPSite(url))
            {
                XmlElement rootNode = (XmlElement)xmlDoc.AppendChild(xmlDoc.CreateElement("Site"));
                rootNode.SetAttribute("url", site.Url);

                if (scope == "site")
                {
                    GetNavigation(xmlDoc, site.RootWeb, true);
                }
                else
                {
                    using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                    {
                        GetNavigation(xmlDoc, web, false);
                    }
                }
            }
            return xmlDoc;
        }

        public static void GetNavigation(XmlDocument xmlDoc, SPWeb web, bool includeChildren)
        {
            if (PublishingWeb.IsPublishingWeb(web))
            {

                PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);
                XmlElement webNode = (XmlElement) xmlDoc.DocumentElement.AppendChild(xmlDoc.CreateElement("Web"));
                webNode.SetAttribute("serverRelativeUrl", web.ServerRelativeUrl);

                GetNavigation(pubweb, webNode);
            }
            if (includeChildren)
            {
                foreach (SPWeb childWeb in web.Webs)
                    GetNavigation(xmlDoc, childWeb, true);
            }
        }

        /// <summary>
        /// Gets the navigation as XML.
        /// </summary>
        /// <param name="pubweb">The pubweb.</param>
        /// <param name="webNode">The web node.</param>
        internal static void GetNavigation(PublishingWeb pubweb, XmlElement webNode)
        {
            XmlDocument xmlDoc = webNode.OwnerDocument;
            XmlElement navNode = (XmlElement)webNode.AppendChild(xmlDoc.CreateElement("Navigation"));
            XmlElement globalNode = (XmlElement)navNode.AppendChild(xmlDoc.CreateElement("Global"));
            globalNode.SetAttribute("IncludeSubSites", pubweb.Navigation.GlobalIncludeSubSites.ToString());
            globalNode.SetAttribute("IncludePages", pubweb.Navigation.GlobalIncludePages.ToString());

            XmlElement currentNode = (XmlElement)navNode.AppendChild(xmlDoc.CreateElement("Current"));
            currentNode.SetAttribute("IncludeSubSites", pubweb.Navigation.CurrentIncludeSubSites.ToString());
            currentNode.SetAttribute("IncludePages", pubweb.Navigation.CurrentIncludePages.ToString());

            EnumerateCollection(pubweb, true, pubweb.Navigation.GlobalNavigationNodes, globalNode);
            EnumerateCollection(pubweb, false, pubweb.Navigation.CurrentNavigationNodes, currentNode);
        }

        /// <summary>
        /// Enumerates the collection.
        /// </summary>
        /// <param name="pubWeb">The publishing web.</param>
        /// <param name="isGlobal">if set to <c>true</c> [is global].</param>
        /// <param name="nodes">The nodes.</param>
        /// <param name="xmlParentNode">The XML parent node.</param>
        private static void EnumerateCollection(PublishingWeb pubWeb, bool isGlobal, SPNavigationNodeCollection nodes, XmlElement xmlParentNode)
        {
            if (nodes == null || nodes.Count == 0)
                return;

            foreach (SPNavigationNode node in nodes)
            {
                NodeTypes type = NodeTypes.None;
                if (node.Properties["NodeType"] != null && !string.IsNullOrEmpty(node.Properties["NodeType"].ToString()))
                    type = (NodeTypes)Enum.Parse(typeof(NodeTypes), node.Properties["NodeType"].ToString());

                if (isGlobal)
                {
                    if ((type == NodeTypes.Area && !pubWeb.Navigation.GlobalIncludeSubSites) ||
                        (type == NodeTypes.Page && !pubWeb.Navigation.GlobalIncludePages))
                        continue;
                }
                else
                {
                    if ((type == NodeTypes.Area && !pubWeb.Navigation.CurrentIncludeSubSites) ||
                        (type == NodeTypes.Page && !pubWeb.Navigation.CurrentIncludePages))
                        continue;
                }

                XmlElement xmlChildNode = xmlParentNode.OwnerDocument.CreateElement("Node");
                xmlParentNode.AppendChild(xmlChildNode);

                xmlChildNode.SetAttribute("Id", node.Id.ToString());
                xmlChildNode.SetAttribute("Title", node.Title);


                // Set the default visibility to true.
                xmlChildNode.SetAttribute("IsVisible", node.IsVisible.ToString());

                #region Determine Visibility

                if (type == NodeTypes.Area)
                {
                    SPWeb web = null;
                    try
                    {
                        string name = node.Url.Trim('/');
                        if (name.Length != 0 && name.IndexOf("/") > 0)
                        {
                            name = name.Substring(name.LastIndexOf('/') + 1);
                        }
                        try
                        {
                            web = pubWeb.Web.Webs[name];
                        }
                        catch (ArgumentException)
                        {
                        }

                        if (web != null && web.Exists && web.ServerRelativeUrl.ToLower() == node.Url.ToLower() &&
                            PublishingWeb.IsPublishingWeb(web))
                        {
                            PublishingWeb tempPubWeb = PublishingWeb.GetPublishingWeb(web);
                            if (isGlobal)
                                xmlChildNode.SetAttribute("IsVisible", tempPubWeb.IncludeInGlobalNavigation.ToString());
                            else
                                xmlChildNode.SetAttribute("IsVisible", tempPubWeb.IncludeInCurrentNavigation.ToString());
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
                    PublishingPage page = null;
                    try
                    {
                        page = pubWeb.GetPublishingPages()[node.Url];
                    }
                    catch (ArgumentException)
                    {
                    }
                    if (page != null)
                    {
                        if (isGlobal)
                            xmlChildNode.SetAttribute("IsVisible", page.IncludeInGlobalNavigation.ToString());
                        else
                            xmlChildNode.SetAttribute("IsVisible", page.IncludeInCurrentNavigation.ToString());
                    }
                }
                #endregion

                XmlElement xmlProp = xmlParentNode.OwnerDocument.CreateElement("Url");
                xmlProp.InnerText = node.Url;
                xmlChildNode.AppendChild(xmlProp);

                foreach (DictionaryEntry d in node.Properties)
                {
                    xmlProp = xmlParentNode.OwnerDocument.CreateElement(d.Key.ToString());
                    xmlProp.InnerText = d.Value.ToString();
                    xmlChildNode.AppendChild(xmlProp);
                }
                EnumerateCollection(pubWeb, isGlobal, node.Children, xmlChildNode);
            }
        }
    }
}
