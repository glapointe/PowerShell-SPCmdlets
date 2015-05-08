using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    public enum CalendarOverlayColor
    {
        LightYellow = 1,
        LightGreen = 2,
        Orange = 3,
        LightTurquise = 4,
        Pink = 5,
        LightBlue = 6,
        IceBlue1 = 7,
        IceBlue2 = 8,
        White = 9
    }

    public class SetListOverlay
    {
        public static void AddCalendarOverlay(SPList targetList, string viewName, string owaUrl, string exchangeUrl, string overlayName, string overlayDescription, CalendarOverlayColor color, bool alwaysShow, bool clearExisting)
        {
            AddCalendarOverlay(targetList, viewName, owaUrl, exchangeUrl, null, overlayName, overlayDescription, color, alwaysShow, clearExisting);
        }
        public static void AddCalendarOverlay(SPList targetList, string viewName, SPList overlayList, string overlayName, string overlayDescription, CalendarOverlayColor color, bool alwaysShow, bool clearExisting)
        {
            AddCalendarOverlay(targetList, viewName, null, null, overlayList, overlayName, overlayDescription, color, alwaysShow, clearExisting);
        }
        private static void AddCalendarOverlay(SPList targetList, string viewName, string owaUrl, string exchangeUrl, SPList overlayList, string overlayName, string overlayDescription, CalendarOverlayColor color, bool alwaysShow, bool clearExisting)
        {
            bool sharePoint = overlayList != null;
            string linkUrl = owaUrl;
            if (sharePoint)
                linkUrl = overlayList.DefaultViewUrl;

            SPView targetView = targetList.DefaultView;
            if (!string.IsNullOrEmpty(viewName))
                targetView = targetList.Views[viewName];

            XmlDocument xml = new XmlDocument();
            XmlElement aggregationElement = null;
            int count = 0;
            if (string.IsNullOrEmpty(targetView.CalendarSettings) || clearExisting)
            {
                xml.AppendChild(xml.CreateElement("AggregationCalendars"));
                aggregationElement = xml.CreateElement("AggregationCalendar");
                xml.DocumentElement.AppendChild(aggregationElement);
            }
            else
            {
                xml.LoadXml(targetView.CalendarSettings);
                XmlNodeList calendars = xml.SelectNodes("/AggregationCalendars/AggregationCalendar");
                if (calendars != null)
                    count = calendars.Count;
                aggregationElement = xml.SelectSingleNode(string.Format("/AggregationCalendars/AggregationCalendar[@CalendarUrl='{0}']", linkUrl)) as XmlElement;
                if (aggregationElement == null)
                {
                    if (count >= 10)
                        throw new SPException(string.Format("10 calendar ovarlays already exist for the calendar {0}.",targetList.RootFolder.ServerRelativeUrl));
                    aggregationElement = xml.CreateElement("AggregationCalendar");
                    xml.DocumentElement.AppendChild(aggregationElement);
                }
            }
            if (!aggregationElement.HasAttribute("Id"))
                aggregationElement.SetAttribute("Id", Guid.NewGuid().ToString("B", CultureInfo.InvariantCulture));

            aggregationElement.SetAttribute("Type", sharePoint ? "SharePoint" : "Exchange");
            aggregationElement.SetAttribute("Name", !string.IsNullOrEmpty(overlayName) ? overlayName : (overlayList == null ? "" : overlayList.Title));
            aggregationElement.SetAttribute("Description", !string.IsNullOrEmpty(overlayDescription) ? overlayDescription : (overlayList == null ? "" : overlayList.Description));
            aggregationElement.SetAttribute("Color", ((int)color).ToString());
            aggregationElement.SetAttribute("AlwaysShow", alwaysShow.ToString());
            aggregationElement.SetAttribute("CalendarUrl", linkUrl);

            XmlElement settingsElement = aggregationElement.SelectSingleNode("./Settings") as XmlElement;
            if (settingsElement == null)
            {
                settingsElement = xml.CreateElement("Settings");
                aggregationElement.AppendChild(settingsElement);
            }
            if (sharePoint)
            {
                settingsElement.SetAttribute("WebUrl", overlayList.ParentWeb.Site.MakeFullUrl(overlayList.ParentWebUrl));
                settingsElement.SetAttribute("ListId", overlayList.ID.ToString("B", CultureInfo.InvariantCulture));
                settingsElement.SetAttribute("ViewId", overlayList.DefaultView.ID.ToString("B", CultureInfo.InvariantCulture));
                settingsElement.SetAttribute("ListFormUrl", overlayList.Forms[PAGETYPE.PAGE_DISPLAYFORM].ServerRelativeUrl);
            }
            else
            {
                settingsElement.SetAttribute("ServiceUrl", exchangeUrl);
            }
            targetView.CalendarSettings = xml.OuterXml;
            targetView.Update();
            /*
            <AggregationCalendars>
                <AggregationCalendar 
                    Id="{cfc22c0b-688e-4555-b1d0-784081a91464}" 
                    Type="SharePoint" 
                    Name="My Overlay Calendar"
                    Description="" 
                    Color="1" 
                    AlwaysShow="True" 
                    CalendarUrl="/Lists/MyOverlayCalendar/calendar.aspx">
                    <Settings 
                        WebUrl="http://demo" 
                        ListId="{4a15e596-674f-4af7-a548-0b01470e8d75}" 
                        ViewId="{594c2916-14e7-4b08-ba36-1126b825bf45}" 
                        ListFormUrl="/Lists/MyOverlayCalendar/DispForm.aspx" />
                </AggregationCalendar>
                <AggregationCalendar 
                    Id="{cfc22c0b-688e-4555-b1d0-784081a91465}" 
                    Type="Exchange" 
                    Name="My Overlay Calendar"
                    Description="" 
                    Color="1" 
                    AlwaysShow="True" 
                    CalendarUrl="<url>">
                    <Settings ServiceUrl="<url>" />
                </AggregationCalendar>
            </AggregationCalendars>
            */
        }
    }
}
