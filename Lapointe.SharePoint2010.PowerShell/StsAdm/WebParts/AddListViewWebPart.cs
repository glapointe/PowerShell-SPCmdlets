using System;
using System.Collections.Specialized;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using Lapointe.SharePoint.PowerShell.StsAdm.Lists;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;
using Microsoft.SharePoint.WebPartPages;
using PublishItems = Lapointe.SharePoint.PowerShell.Common.Lists.PublishItems;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebParts
{
    public class AddListViewWebPart : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AddListViewWebPart"/> class.
        /// </summary>
        public AddListViewWebPart()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the page url."));
            parameters.Add(new SPParam("listurl", "lu", true, null, new SPUrlValidator(), "Please specify the page url."));
            parameters.Add(new SPParam("title", "t", false, null, new SPNonEmptyValidator(), "Please specify the web part title."));
            parameters.Add(new SPParam("viewtitle", "v", false, null, new SPNonEmptyValidator(), "Please specify the name of the view to use."));
            parameters.Add(new SPParam("zone", "z", true, null, new SPNonEmptyValidator(), "Please specify the zone to add the web part to."));
            parameters.Add(new SPParam("zoneindex", "zi", true, null, new SPIntRangeValidator(0, int.MaxValue)));
            parameters.Add(new SPParam("linktitle", "lt", false, null, null));
            parameters.Add(new SPParam("publish", "p", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nAdds a list view web part to a page.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web part page URL>");
            sb.Append("\r\n\t-listurl <list url>");
            sb.Append("\r\n\t-viewtitle <view title>");
            sb.Append("\r\n\t-zone <zone ID>");
            sb.Append("\r\n\t-zoneindex <zone index>");
            sb.Append("\r\n\t[-title <web part title>]");
            sb.Append("\r\n\t[-linktitle]");
            sb.Append("\r\n\t[-publish]");

            Init(parameters, sb.ToString());
        }

        
        #region ISPStsadmCommand Members

        /// <summary>
        /// Gets the help message.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <returns></returns>
        public override string GetHelpMessage(string command)
        {
            return HelpMessage;
        }

        /// <summary>
        /// Runs the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public override int Execute(string command, StringDictionary keyValues, out string output)
        {
            output = string.Empty;

            string pageUrl = Params["url"].Value;
            string listUrl = Params["listurl"].Value;
            string title = Params["title"].Value;
            string zoneID = Params["zone"].Value;
            int zoneIndex = int.Parse(Params["zoneindex"].Value);
            bool publish = Params["publish"].UserTypedIn;
            string viewTitle = Params["viewtitle"].Value;
            bool linkTitle = Params["linktitle"].UserTypedIn;
            Common.WebParts.AddListViewWebPart.Add(pageUrl, listUrl, title, viewTitle, zoneID, zoneIndex, linkTitle, null, PartChromeType.Default, publish);

            return (int)ErrorCodes.NoError;
        }


        #endregion
    }
}