using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{

    public class GetListOverlays
    {
        public static List<SPList> GetOverlayLists(SPList list)
        {
            XmlDocument xml = new XmlDocument();
            SPView view = list.DefaultView;
            if (string.IsNullOrEmpty(view.CalendarSettings))
                return null;
            xml.LoadXml(view.CalendarSettings);
            var nodes = xml.SelectNodes("//AggregationCalendar[@Type='SharePoint']/Settings");
            if (nodes == null)
                return null;
            SPSite site = null;
            string webUrl = null;
            List<SPList> lists = new List<SPList>();
            try
            {
                foreach (XmlElement listNode in nodes)
                {
                    try
                    {
                        string currentWebUrl = listNode.GetAttribute("WebUrl");
                        if (currentWebUrl != webUrl || site == null)
                        {
                            if (site != null)
                                site.Dispose();

                            webUrl = currentWebUrl;
                            site = new SPSite(webUrl);
                        }
                        SPWeb web = site.OpenWeb();
                        lists.Add(web.Lists[new Guid(listNode.GetAttribute("ListId"))]);
                    }
                    catch (Exception)
                    {
                        Logger.WriteWarning("Unable to retrieve calendar overlay: {0}", listNode.OuterXml);
                        continue;
                    }
                }
            }
            finally
            {
                if (site != null)
                    site.Dispose();
            }
            return lists;
        }

    }
}
