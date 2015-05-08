using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.SharePoint.PowerShell.Enums;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebPartPages;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.WebParts;
using System.Web.UI.WebControls.WebParts;

namespace Lapointe.SharePoint.PowerShell.Common.Pages
{
    public class WikiPageUtilities
    {
        // Reference code - OfficeDevPnP: https://github.com/OfficeDev/PnP/blob/master/OfficeDevPnP.Core/OfficeDevPnP.Core/AppModelExtensions/PageExtensions.cs
        private const string WikiPage_OneColumn = @"<div class=""ExternalClassC1FD57BEDB8942DC99A06C02F9A98241""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;100%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,1</span></div>";
        private const string WikiPage_OneColumnSideBar = @"<div class=""ExternalClass47565ACDF7974263AA4A556DA974B687""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;66.6%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,2</span></div>";
        private const string WikiPage_TwoColumns = @"<div class=""ExternalClass3811C839E5984CCEA4C8CF738462AFD8""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,2</span></div>";
        private const string WikiPage_TwoColumnsHeader = @"<div class=""ExternalClass850251EB51394304A07A64A05C0BB0F1""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,false,2</span></div>";
        private const string WikiPage_TwoColumnsHeaderFooter = @"<div class=""ExternalClass71C5527252AD45859FA774445D4909A2""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;49.95%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td colspan=""2""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,true,2</span></div>";
        private const string WikiPage_ThreeColumns = @"<div class=""ExternalClass833D1FA704C94892A26C4069C3FE5FE9""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">false,false,3</span></div>";
        private const string WikiPage_ThreeColumnsHeader = @"<div class=""ExternalClassD1A150D6187F449B8A6C4BEA2D4913BB""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,false,3</span></div>";
        private const string WikiPage_ThreeColumnsHeaderFooter = @"<div class=""ExternalClass5849C2C61FEC44E9B249C60F7B0ACA38""><table id=""layoutsTable"" style=""width&#58;100%;""><tbody><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td><td class=""ms-wiki-columnSpacing"" style=""width&#58;33.3%;""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr><tr style=""vertical-align&#58;top;""><td colspan=""3""><div class=""ms-rte-layoutszone-outer"" style=""width&#58;100%;""><div class=""ms-rte-layoutszone-inner"" role=""textbox"" aria-haspopup=""true"" aria-autocomplete=""both"" aria-multiline=""true""></div>&#160;</div></td></tr></tbody></table><span id=""layoutsData"" style=""display&#58;none;"">true,true,3</span></div>";


        private static List<string> GetWikiPageHtmlFromZones(SPListItem item)
        {
            string wikiField = (string)item["WikiField"];

            XmlDocument xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            List<string> html = new List<string>();
            var nodes = xd.SelectNodes("//div[@class=\"ms-rte-layoutszone-inner\"]");
            foreach (var node in nodes)
            {
                html.Add(((XmlElement)node).InnerXml);
            }
            return html;
        }

        public static void AddHtmlToWikiPage(SPListItem item, string html)
        {
            item["WikiField"] = html;
            item.Update();
        }

        public static void AddHtmlToWikiPage(SPListItem item, string html, int row, int col, bool prepend)
        {
            // Reference code - OfficeDevPnP: https://github.com/OfficeDev/PnP/blob/master/OfficeDevPnP.Core/OfficeDevPnP.Core/AppModelExtensions/PageExtensions.cs
            string wikiField = (string)item["WikiField"];

            XmlDocument xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            // Sometimes the wikifield content seems to be surrounded by an additional div? 
            XmlElement layoutsTable = xd.SelectSingleNode("div/div/table") as XmlElement;
            if (layoutsTable == null)
            {
                layoutsTable = xd.SelectSingleNode("div/table") as XmlElement;
            }

            // Add the html content 
            XmlElement layoutsZoneInner = layoutsTable.SelectSingleNode(string.Format("tbody/tr[{0}]/td[{1}]/div/div", row, col)) as XmlElement;
            if (layoutsZoneInner == null)
                throw new ArgumentException("Unable to locate row and/or column to insert HTML into.");

            XmlText text = xd.CreateTextNode("!!123456789!!");
            if (prepend)
                layoutsZoneInner.PrependChild(text);
            else
                layoutsZoneInner.AppendChild(text);

            item["WikiField"] = xd.OuterXml.Replace("!!123456789!!", html); ;
            item.Update();
        }

        public static void SetWikiPageLayout(SPListItem item, WikiPageLayout layout)
        {
            var html = GetWikiPageHtmlFromZones(item);
            AddLayoutToWikiPage(item, layout);
            foreach (string zoneHtml in html)
            {
                AddHtmlToWikiPage(item, zoneHtml, 1, 1, false);
            }
        }

        public static void AddLayoutToWikiPage(SPListItem item, WikiPageLayout layout)
        {
            // Reference code - OfficeDevPnP: https://github.com/OfficeDev/PnP/blob/master/OfficeDevPnP.Core/OfficeDevPnP.Core/AppModelExtensions/PageExtensions.cs
            string html = "";

            switch (layout)
            {
                case WikiPageLayout.OneColumn:
                    html = WikiPage_OneColumn;
                    break;
                case WikiPageLayout.OneColumnSideBar:
                    html = WikiPage_OneColumnSideBar;
                    break;
                case WikiPageLayout.TwoColumns:
                    html = WikiPage_TwoColumns;
                    break;
                case WikiPageLayout.TwoColumnsHeader:
                    html = WikiPage_TwoColumnsHeader;
                    break;
                case WikiPageLayout.TwoColumnsHeaderFooter:
                    html = WikiPage_TwoColumnsHeaderFooter;
                    break;
                case WikiPageLayout.ThreeColumns:
                    html = WikiPage_ThreeColumns;
                    break;
                case WikiPageLayout.ThreeColumnsHeader:
                    html = WikiPage_ThreeColumnsHeader;
                    break;
                case WikiPageLayout.ThreeColumnsHeaderFooter:
                    html = WikiPage_ThreeColumnsHeaderFooter;
                    break;
                default:
                    break;
            }

            AddHtmlToWikiPage(item, html);
        }

        public static SPFile AddWikiPage(SPWeb web, string wikiPageLibraryUrlOrTitle, string wikiPageName)
        {
            return AddWikiPage(web, wikiPageLibraryUrlOrTitle, wikiPageName, false);
        }
        internal static SPFile AddWikiPage(SPWeb web, string wikiPageLibraryUrlOrTitle, string wikiPageName, bool throwIfExists)
        {
            SPListPipeBind lpb = new SPListPipeBind(wikiPageLibraryUrlOrTitle);
            SPList list = lpb.Read(web);
            return AddWikiPage(list, wikiPageName, throwIfExists);
        }

        public static SPFile AddWikiPage(SPList list, string wikiPageName)
        {
            return AddWikiPage(list, wikiPageName, false);
        }
        internal static SPFile AddWikiPage(SPList list, string wikiPageName, bool throwIfExists)
        {
            var folder = list.RootFolder;
            return AddWikiPage(folder, wikiPageName, throwIfExists);
        }

        public static SPFile AddWikiPage(SPFolder folder, string wikiPageName)
        {
            return AddWikiPage(folder, wikiPageName, false);
        }

        internal static SPFile AddWikiPage(SPFolder folder, string wikiPageName, bool throwIfExists)
        {
            return folder.Files.Add(wikiPageName, SPTemplateFileType.WikiPage);
        }

        public static WebPart AddWebPartToWikiPage(SPListItem item, WebPart webPart, string title, int row, int col, bool addSpace, PartChromeType chromeType, bool publish)
        {
            string wikiField = (string)item["WikiField"];
            XmlDocument xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            // Sometimes the wikifield content seems to be surrounded by an additional div? 
            XmlElement layoutsTable = xd.SelectSingleNode("div/div/table") as XmlElement;
            if (layoutsTable == null)
            {
                layoutsTable = xd.SelectSingleNode("div/table") as XmlElement;
            }

            XmlElement layoutsZoneInner = layoutsTable.SelectSingleNode(string.Format("tbody/tr[{0}]/td[{1}]/div/div", row, col)) as XmlElement;
            if (layoutsZoneInner == null)
                throw new ArgumentException("Unable to locate row and/or column to insert HTML into.");


            bool checkBackIn = false;
            if (item.File.InDocumentLibrary)
            {
                bool ignoreCheckOut = item.File.CheckOutType == SPFile.SPCheckOutType.None && !item.ParentList.ForceCheckout;
                if (!ignoreCheckOut)
                {
                    if (!Utilities.IsCheckedOut(item) || !Utilities.IsCheckedOutByCurrentUser(item))
                    {
                        checkBackIn = true;
                        item.File.CheckOut();
                    }
                }
                // If it's checked out by another user then this will throw an informative exception so let it do so.
            }

            SPLimitedWebPartManager limitedWebPartManager = null;
            try
            {
                limitedWebPartManager = item.File.GetLimitedWebPartManager(System.Web.UI.WebControls.WebParts.PersonalizationScope.Shared);
                webPart.ChromeType = chromeType;
                limitedWebPartManager.AddWebPart(webPart, "wpz", 0);
                
                Guid storageKey = limitedWebPartManager.GetStorageKey(webPart);

                // - space element
                XmlElement space = xd.CreateElement("p");
                XmlText text = xd.CreateTextNode(" ");
                space.AppendChild(text);

                // - wpBoxDiv
                XmlElement wpBoxDiv = xd.CreateElement("div");
                layoutsZoneInner.AppendChild(wpBoxDiv);

                if (addSpace)
                {
                    layoutsZoneInner.AppendChild(space);
                }

                XmlAttribute attribute = xd.CreateAttribute("class");
                wpBoxDiv.Attributes.Append(attribute);
                attribute.Value = "ms-rtestate-read ms-rte-wpbox";
                attribute = xd.CreateAttribute("contentEditable");
                wpBoxDiv.Attributes.Append(attribute);
                attribute.Value = "false";
                // - div1
                XmlElement div1 = xd.CreateElement("div");
                wpBoxDiv.AppendChild(div1);
                div1.IsEmpty = false;
                attribute = xd.CreateAttribute("class");
                div1.Attributes.Append(attribute);
                attribute.Value = "ms-rtestate-read " + storageKey.ToString("D");
                attribute = xd.CreateAttribute("id");
                div1.Attributes.Append(attribute);
                attribute.Value = "div_" + storageKey.ToString("D");
                // - div2
                XmlElement div2 = xd.CreateElement("div");
                wpBoxDiv.AppendChild(div2);
                div2.IsEmpty = false;
                attribute = xd.CreateAttribute("style");
                div2.Attributes.Append(attribute);
                attribute.Value = "display:none";
                attribute = xd.CreateAttribute("id");
                div2.Attributes.Append(attribute);
                attribute.Value = "vid_" + storageKey.ToString("D");

                item["WikiField"] = xd.OuterXml;
                item.Update();

                if (webPart.Title != title)
                {
                    webPart.Title = title;
                    limitedWebPartManager.SaveChanges(webPart);
                }
            }
            finally
            {
                if (limitedWebPartManager != null) 
                    limitedWebPartManager.Dispose();

                SPFile file = item.File;
                if (file.InDocumentLibrary && Utilities.IsCheckedOut(file.Item) && (checkBackIn || publish))
                    file.CheckIn("Checking in changes to page due to new web part being added: " + title);

                if (publish && file.InDocumentLibrary)
                {
                    file.Publish("Publishing changes to page due to new web part being added: " + title);
                    if (file.Item.ModerationInformation != null)
                    {
                        file.Approve("Approving changes to page due to new web part being added: " + title);
                    }
                }
            }
            return webPart;
        }


        public static WebPart AddWebPartToWikiPage(SPListItem item, string webPartXml, string title, int row, int col, bool addSpace, Hashtable customReplaceText, PartChromeType chromeType, bool publish)
        {
            string wikiField = (string)item["WikiField"];
            XmlDocument xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            // Sometimes the wikifield content seems to be surrounded by an additional div? 
            XmlElement layoutsTable = xd.SelectSingleNode("div/div/table") as XmlElement;
            if (layoutsTable == null)
            {
                layoutsTable = xd.SelectSingleNode("div/table") as XmlElement;
            }

            XmlElement layoutsZoneInner = layoutsTable.SelectSingleNode(string.Format("tbody/tr[{0}]/td[{1}]/div/div", row, col)) as XmlElement;
            if (layoutsZoneInner == null)
                throw new ArgumentException("Unable to locate row and/or column to insert HTML into.");

            bool checkBackIn = false;
            if (item.File.InDocumentLibrary)
            {
                bool ignoreCheckOut = item.File.CheckOutType == SPFile.SPCheckOutType.None && !item.ParentList.ForceCheckout;
                if (!ignoreCheckOut)
                {
                    if (!Utilities.IsCheckedOut(item) || !Utilities.IsCheckedOutByCurrentUser(item))
                    {
                        checkBackIn = true;
                        item.File.CheckOut();
                    }
                }
                // If it's checked out by another user then this will throw an informative exception so let it do so.
            }

            SPLimitedWebPartManager limitedWebPartManager = null;
            WebPart wpdNew = null;
            try
            {
                limitedWebPartManager = item.File.GetLimitedWebPartManager(System.Web.UI.WebControls.WebParts.PersonalizationScope.Shared);
                wpdNew = SPCmdletImportWebPart.AddWebPart(limitedWebPartManager, item.File, webPartXml, customReplaceText);
                wpdNew.ChromeType = chromeType;
                limitedWebPartManager.AddWebPart(wpdNew, "wpz", 0);
                Guid storageKey = limitedWebPartManager.GetStorageKey(wpdNew);


                // - space element
                XmlElement space = xd.CreateElement("p");
                XmlText text = xd.CreateTextNode(" ");
                space.AppendChild(text);

                // - wpBoxDiv
                XmlElement wpBoxDiv = xd.CreateElement("div");
                layoutsZoneInner.AppendChild(wpBoxDiv);

                if (addSpace)
                {
                    layoutsZoneInner.AppendChild(space);
                }

                XmlAttribute attribute = xd.CreateAttribute("class");
                wpBoxDiv.Attributes.Append(attribute);
                attribute.Value = "ms-rtestate-read ms-rte-wpbox";
                attribute = xd.CreateAttribute("contentEditable");
                wpBoxDiv.Attributes.Append(attribute);
                attribute.Value = "false";
                // - div1
                XmlElement div1 = xd.CreateElement("div");
                wpBoxDiv.AppendChild(div1);
                div1.IsEmpty = false;
                attribute = xd.CreateAttribute("class");
                div1.Attributes.Append(attribute);
                attribute.Value = "ms-rtestate-read " + storageKey.ToString("D");
                attribute = xd.CreateAttribute("id");
                div1.Attributes.Append(attribute);
                attribute.Value = "div_" + storageKey.ToString("D");
                // - div2
                XmlElement div2 = xd.CreateElement("div");
                wpBoxDiv.AppendChild(div2);
                div2.IsEmpty = false;
                attribute = xd.CreateAttribute("style");
                div2.Attributes.Append(attribute);
                attribute.Value = "display:none";
                attribute = xd.CreateAttribute("id");
                div2.Attributes.Append(attribute);
                attribute.Value = "vid_" + storageKey.ToString("D");

                item["WikiField"] = xd.OuterXml;
                item.Update();

                if (wpdNew.Title != title)
                {
                    wpdNew.Title = title;
                    limitedWebPartManager.SaveChanges(wpdNew);
                }
            }
            finally
            {
                if (limitedWebPartManager != null) 
                    limitedWebPartManager.Dispose();

                SPFile file = item.File;
                if (file.InDocumentLibrary && Utilities.IsCheckedOut(file.Item) && (checkBackIn || publish))
                    file.CheckIn("Checking in changes to page due to new web part being added: " + title);

                if (publish && file.InDocumentLibrary)
                {
                    file.Publish("Publishing changes to page due to new web part being added: " + title);
                    if (file.Item.ModerationInformation != null)
                    {
                        file.Approve("Approving changes to page due to new web part being added: " + title);
                    }
                }

            }
            return wpdNew;
        }
    }

}
