using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebPartPages;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;
using System.Web;
using Microsoft.SharePoint.WebControls;
using System.Collections.Generic;

namespace Lapointe.SharePoint.PowerShell.Common.WebParts
{
    internal static class EnumPageWebParts
    {
        internal static string GetWebPartXml(string pageUrl, bool simple)
        {
            using (SPSite site = new SPSite(pageUrl))
            using (SPWeb web = site.OpenWeb()) // The url contains a filename so AllWebs[] will not work unless we want to try and parse which we don't
            {
                SPLimitedWebPartManager webPartMngr = null;
                bool cleanupContext = false;
                try
                {
                    if (HttpContext.Current == null)
                    {
                        cleanupContext = true;
                        HttpRequest httpRequest = new HttpRequest("", web.Url, "");
                        HttpContext.Current = new HttpContext(httpRequest, new HttpResponse(new StringWriter()));
                        SPControl.SetContextWeb(HttpContext.Current, web);
                    }

                    webPartMngr = web.GetLimitedWebPartManager(pageUrl, PersonalizationScope.Shared);

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.AppendChild(xmlDoc.CreateElement("WebParts"));
                    xmlDoc.DocumentElement.SetAttribute("page", web.Site.MakeFullUrl(webPartMngr.ServerRelativeUrl));

                    XmlElement shared = xmlDoc.CreateElement("Shared");
                    xmlDoc.DocumentElement.AppendChild(shared);


                    string tempXml = string.Empty;
                    foreach (WebPart wp in webPartMngr.WebParts)
                    {
                        if (!simple)
                            tempXml += GetWebPartDetails(wp, webPartMngr);
                        else
                            tempXml += GetWebPartDetailsSimple(wp, webPartMngr);
                    }
                    shared.InnerXml = tempXml;

                    XmlElement user = xmlDoc.CreateElement("User");
                    xmlDoc.DocumentElement.AppendChild(user);

                    webPartMngr.Web.Dispose();
                    webPartMngr.Dispose();

                    webPartMngr = web.GetLimitedWebPartManager(pageUrl, PersonalizationScope.User);
                    tempXml = string.Empty;
                    foreach (WebPart wp in webPartMngr.WebParts)
                    {
                        if (!simple)
                            tempXml += GetWebPartDetails(wp, webPartMngr);
                        else
                            tempXml += GetWebPartDetailsSimple(wp, webPartMngr);
                    }
                    user.InnerXml = tempXml;
                    return Utilities.GetFormattedXml(xmlDoc);
                }
                finally
                {
                    if (webPartMngr != null)
                    {
                        webPartMngr.Web.Dispose();
                        webPartMngr.Dispose();
                    }
                    if (HttpContext.Current != null && cleanupContext)
                    {
                        HttpContext.Current = null;
                    }

                }
            }
        }


        /// <summary>
        /// Gets the web part details.
        /// </summary>
        /// <param name="wp">The web part.</param>
        /// <param name="manager">The web part manager.</param>
        /// <returns></returns>
        internal static string GetWebPartDetails(WebPart wp, SPLimitedWebPartManager manager)
        {
            if (wp.ExportMode == WebPartExportMode.None)
            {
                Logger.WriteWarning("Unable to export {0}", wp.Title);
                return "";
            }
            StringBuilder sb = new StringBuilder();

            XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
            xmlWriter.Formatting = Formatting.Indented;
            manager.ExportWebPart(wp, xmlWriter);
            xmlWriter.Flush();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(sb.ToString());

            XmlElement elem = xmlDoc.DocumentElement;
            if (xmlDoc.DocumentElement.Name == "webParts")
            {
                elem = (XmlElement)xmlDoc.DocumentElement.ChildNodes[0];

                // We've found a v3 web part but the export method does not export what the zone ID is so we
                // have to manually add that in.  Unfortunately the Zone property is always null because we are
                // using a SPLimitedWebPartManager so we have to use the helper method GetZoneID to set the zone ID.
                XmlElement property = xmlDoc.CreateElement("property");
                property.SetAttribute("name", "ZoneID");
                property.SetAttribute("type", "string");
                property.InnerText = manager.GetZoneID(wp);
                elem.ChildNodes[1].ChildNodes[0].AppendChild(property);
            }

            return elem.OuterXml.Replace(" xmlns=\"\"", ""); // Just some minor cleanup to deal with erroneous namespace tags added due to the zoneID being added manually.
        }

        /// <summary>
        /// Gets the web part details.
        /// </summary>
        /// <param name="wp">The web part.</param>
        /// <param name="manager">The web part manager.</param>
        internal static string GetWebPartDetailsSimple(WebPart wp, SPLimitedWebPartManager manager)
        {
            XmlDocument xmlDoc = new XmlDocument();

            XmlElement wpXml = xmlDoc.CreateElement("WebPart");
            xmlDoc.AppendChild(wpXml);

            wpXml.SetAttribute("id", wp.ID);

            XmlElement prop = xmlDoc.CreateElement("Title");
            prop.InnerText = wp.Title;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("AllowClose");
            prop.InnerText = wp.AllowClose.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("AllowConnect");
            prop.InnerText = wp.AllowConnect.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("AllowEdit");
            prop.InnerText = wp.AllowEdit.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("AllowHide");
            prop.InnerText = wp.AllowHide.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("TitleUrl");
            prop.InnerText = wp.TitleUrl;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("AllowMinimize");
            prop.InnerText = wp.AllowMinimize.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("AllowZoneChange");
            prop.InnerText = wp.AllowZoneChange.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("CatalogIconImageUrl");
            prop.InnerText = wp.CatalogIconImageUrl;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("ChromeState");
            prop.InnerText = wp.ChromeState.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("ChromeType");
            prop.InnerText = wp.ChromeType.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("Description");
            prop.InnerText = wp.Description;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("DisplayTitle");
            prop.InnerText = wp.DisplayTitle;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("HasSharedData");
            prop.InnerText = wp.HasSharedData.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("HasUserData");
            prop.InnerText = wp.HasUserData.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("Hidden");
            prop.InnerText = wp.Hidden.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("IsClosed");
            prop.InnerText = wp.IsClosed.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("IsShared");
            prop.InnerText = wp.IsShared.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("IsStandalone");
            prop.InnerText = wp.IsStandalone.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("IsStatic");
            prop.InnerText = wp.IsStatic.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("Subtitle");
            prop.InnerText = wp.Subtitle;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("TitleUrl");
            prop.InnerText = wp.TitleUrl;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("TitleIconImageUrl");
            prop.InnerText = wp.TitleIconImageUrl;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("Zone");
            if (wp.Zone != null)
                prop.InnerText = wp.Zone.ToString();
            else
                prop.InnerText = manager.GetZoneID(wp);
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("ZoneIndex");
            prop.InnerText = wp.ZoneIndex.ToString();
            wpXml.AppendChild(prop);

            return Utilities.GetFormattedXml(xmlDoc);
        }


        /// <summary>
        /// Gets the web part details.
        /// </summary>
        /// <param name="wp">The web part.</param>
        /// <param name="manager">The web part manager.</param>
        internal static string GetWebPartDetailsMinimal(WebPart wp, SPLimitedWebPartManager manager)
        {
            XmlDocument xmlDoc = new XmlDocument();

            XmlElement wpXml = xmlDoc.CreateElement("WebPart");
            xmlDoc.AppendChild(wpXml);

            wpXml.SetAttribute("id", wp.ID);

            XmlElement prop = xmlDoc.CreateElement("Title");
            prop.InnerText = wp.Title;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("Description");
            prop.InnerText = wp.Description;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("DisplayTitle");
            prop.InnerText = wp.DisplayTitle;
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("Hidden");
            prop.InnerText = wp.Hidden.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("IsClosed");
            prop.InnerText = wp.IsClosed.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("IsShared");
            prop.InnerText = wp.IsShared.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("IsStandalone");
            prop.InnerText = wp.IsStandalone.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("IsStatic");
            prop.InnerText = wp.IsStatic.ToString();
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("Zone");
            if (wp.Zone != null)
                prop.InnerText = wp.Zone.ToString();
            else
                prop.InnerText = manager.GetZoneID(wp);
            wpXml.AppendChild(prop);

            prop = xmlDoc.CreateElement("ZoneIndex");
            prop.InnerText = wp.ZoneIndex.ToString();
            wpXml.AppendChild(prop);

            return Utilities.GetFormattedXml(xmlDoc);
        }
    }
}
