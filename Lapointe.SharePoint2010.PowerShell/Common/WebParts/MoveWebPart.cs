using Lapointe.SharePoint.PowerShell.Common.Lists;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebPartPages;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;

namespace Lapointe.SharePoint.PowerShell.Common.WebParts
{
    public class MoveWebPart
    {
        public static void MoveById(string url, string webPartId, string webPartZone, string webPartZoneIndex, bool publish)
        {
            Move(url, webPartId, null, webPartZone, webPartZoneIndex, publish);
        }
        public static void MoveByTitle(string url, string webPartTitle, string webPartZone, string webPartZoneIndex, bool publish)
        {
            Move(url, null, webPartTitle, webPartZone, webPartZoneIndex, publish);
        }
        internal static void Move(string url, string webPartId, string webPartTitle, string webPartZone, string webPartZoneIndex, bool publish)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb()) // The url contains a filename so AllWebs[] will not work unless we want to try and parse which we don't
            {
                SPFile file = web.GetFile(url);
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
                try
                {
                    WebPart wp;
                    SPLimitedWebPartManager manager;
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
                                "Unable to find specified web part. Try specifying the -id parameter instead (use Get-SPWebPartList to get the ID)");
                        }
                    }

                    if (wp == null)
                    {
                        throw new SPException("Unable to find specified web part.");
                    }

                    string zoneID = manager.GetZoneID(wp);
                    int zoneIndex = wp.ZoneIndex;

                    if (!string.IsNullOrEmpty(webPartZone))
                        zoneID = webPartZone;
                    if (!string.IsNullOrEmpty(webPartZoneIndex))
                        zoneIndex = int.Parse(webPartZoneIndex);

                    // Set this so that we can add it to the check-in comment.
                    displayTitle = wp.DisplayTitle;

                    manager.MoveWebPart(wp, zoneID, zoneIndex);
                    manager.SaveChanges(wp);
                }
                finally
                {
                    if (checkIn)
                        file.CheckIn("Checking in changes to page due to moving of web part " + displayTitle);
                    if (publish && file.InDocumentLibrary)
                    {
                        PublishItems pi = new PublishItems();
                        pi.PublishListItem(file.Item, file.Item.ParentList, false, "Move-SPWebPart", "Checking in changes to page due to moving of web part " + displayTitle, null);
                    }
                }
            }
        }

    }
}
