using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Lapointe.SharePoint.PowerShell.Common.Lists;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Portal.WebControls;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.WebPartPages;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;

namespace Lapointe.SharePoint.PowerShell.Common.WebParts
{
    public enum SetWebPartStateAction { Delete, Open, Close, Update }
    public class SetWebPartState
    {
        public static void SetWebPartByTitle(string url, SetWebPartStateAction action, string webPartTitle, string webPartZone, string webPartZoneIndex, Hashtable props, bool publish)
        {
            SetWebPart(url, action, null, webPartTitle, webPartZone, webPartZoneIndex, props, publish);
        }
        public static void SetWebPartById(string url, SetWebPartStateAction action, string webPartId, string webPartZone, string webPartZoneIndex, Hashtable props, bool publish)
        {
            SetWebPart(url, action, webPartId, null, webPartZone, webPartZoneIndex, props, publish);
        }


        internal static void SetWebPart(string url, SetWebPartStateAction action, string webPartId, string webPartTitle, string webPartZone, string webPartZoneIndex, Hashtable props, bool publish)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb()) // The url contains a filename so AllWebs[] will not work unless we want to try and parse which we don't
            {
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


                    SPFile file = web.GetFile(url);

                    // file.Item will throw "The object specified does not belong to a list." if the url passed
                    // does not correspond to a file in a list.

                    bool checkIn = false;
                    if (file.InDocumentLibrary)
                    {
                        if (!Utilities.IsCheckedOut(file.Item) || !Utilities.IsCheckedOutByCurrentUser(file.Item))
                        {
                            file.CheckOut();
                            checkIn = true;
                            // If it's checked out by another user then this will throw an informative exception so let it do so.
                        }
                    }

                    string displayTitle = string.Empty;
                    WebPart wp = null;
                    SPLimitedWebPartManager manager = null;
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
                                    "Unable to find specified web part using title \"" + webPartTitle + "\". Try specifying the -id parameter instead (use Get-SPWebPartList to get the ID)");
                            }
                        }

                        if (wp == null)
                        {
                            throw new SPException("Unable to find specified web part.");
                        }

                        // Set this so that we can add it to the check-in comment.
                        displayTitle = wp.DisplayTitle;

                        if (action == SetWebPartStateAction.Delete)
                            manager.DeleteWebPart(wp);
                        else if (action == SetWebPartStateAction.Close)
                            manager.CloseWebPart(wp);
                        else if (action == SetWebPartStateAction.Open)
                            manager.OpenWebPart(wp);


                        if (action != SetWebPartStateAction.Delete)
                        {
                            string zoneID = manager.GetZoneID(wp);
                            int zoneIndex = wp.ZoneIndex;

                            if (!string.IsNullOrEmpty(webPartZone))
                                zoneID = webPartZone;
                            if (!string.IsNullOrEmpty(webPartZoneIndex))
                                zoneIndex = int.Parse(webPartZoneIndex);

                            manager.MoveWebPart(wp, zoneID, zoneIndex);

                            if (props != null && props.Count > 0)
                            {
                                SetWebPartProperties(wp, props);
                            }
                            manager.SaveChanges(wp);
                        }

                    }
                    finally
                    {
                        if (manager != null)
                        {
                            manager.Web.Dispose();
                            manager.Dispose();
                        }
                        if (wp != null)
                            wp.Dispose();

                        if (checkIn)
                            file.CheckIn("Checking in changes to page due to state change of web part " + displayTitle);
                        if (publish && file.InDocumentLibrary)
                        {
                            PublishItems pi = new PublishItems();
                            pi.PublishListItem(file.Item, file.Item.ParentList, false, "Set-SPWebPart", "Checking in changes to page due to state change of web part " + displayTitle, null);
                        }

                    }
                }
                finally
                {
                    if (HttpContext.Current != null && cleanupContext)
                    {
                        HttpContext.Current = null;
                    }
                }

            }
        }

        /// <summary>
        /// Gets the properties array.
        /// </summary>
        /// <param name="properties">The properties.</param>
        /// <param name="propertiesFile">The properties file.</param>
        /// <param name="seperator">The seperator.</param>
        /// <returns></returns>
        public static Hashtable GetPropertiesArray(string properties, string propertiesFile, string seperator)
        {
            if (string.IsNullOrEmpty(properties) && string.IsNullOrEmpty(propertiesFile))
                return null;

            Hashtable props = new Hashtable();
            if (!string.IsNullOrEmpty(properties))
            {
                properties = properties.Replace(seperator + seperator, "[STSADM_COMMA]");
                foreach (string s in properties.Split(seperator.ToCharArray()))
                {
                    string[] prop = s.Split(new[] { '=' }, 2);
                    props.Add(prop[0].Trim(), prop[1].Trim().Replace("[STSADM_COMMA]", seperator));
                }
            }
            else
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(propertiesFile);
                foreach (XmlElement propElement in xmlDoc.DocumentElement.ChildNodes)
                {
                    props.Add(propElement.Attributes["Name"].Value, propElement.InnerText.Trim());
                }
            }
            return props;
        }
        public static Hashtable GetPropertiesArray(XmlDocument xmlDoc)
        {
            Hashtable props = new Hashtable();
            foreach (XmlElement propElement in xmlDoc.DocumentElement.ChildNodes)
            {
                props.Add(propElement.Attributes["Name"].Value, propElement.InnerText.Trim());
            }
            return props;
        }


        /// <summary>
        /// Sets the web part properties.
        /// </summary>
        /// <param name="xWp">The x wp.</param>
        /// <param name="properties">The properties.</param>
        private static void SetWebPartProperties(WebPart xWp, Hashtable properties)
        {
            if (properties == null)
                return;

            foreach (string prop in properties.Keys)
            {
                try
                {
                    object val = properties[prop];
                    if (xWp is Microsoft.SharePoint.WebPartPages.WebPart)
                    {
                        if (prop == "Width")
                        {
                            ((Microsoft.SharePoint.WebPartPages.WebPart)xWp).Width = val.ToString();
                            return;
                        }
                        if (prop == "Height")
                        {
                            ((Microsoft.SharePoint.WebPartPages.WebPart)xWp).Height = val.ToString();
                            return;
                        }
                    }
                    PropertyInfo propInfo = xWp.GetType().GetProperty(prop);
                    if (propInfo == null)
                        throw new SPException(string.Format("Unable to set property '{0}'.  The property could not be found.", prop));

                    val = Convert.ChangeType(val, propInfo.PropertyType);
                    if (val == null)
                        throw new SPException(string.Format("Unable to convert '{0}' to type '{1}'", properties[prop], propInfo.PropertyType));

#if MOSS
                    if (xWp is KPIListWebPart && prop == "TitleUrl")
                    {
                        if (val.ToString().Contains("://"))
                            val = Utilities.GetServerRelUrlFromFullUrl((string)val);
                    }
                    if (xWp is KPIListWebPart && prop == "ListURL")
                    {
                        if (val.ToString().Contains("://"))
                            val = Utilities.GetServerRelUrlFromFullUrl((string)val);

                        Utilities.SetFieldValue(xWp, typeof(KPIListWebPart), "listUrl", val);
                        Utilities.SetFieldValue(xWp, typeof(KPIListWebPart), "spList", null);
                        continue;
                    }
#endif

                    propInfo.SetValue(xWp, val, null);
                }
                catch (Exception ex)
                {
                    throw new SPException(string.Format("Error setting property value {0} to {1}:\r\n{2}", prop, properties[prop], Utilities.FormatException(ex)));
                }
            }
        }

    }
}
