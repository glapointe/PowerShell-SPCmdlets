#if MOSS
using System;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing.WebControls;
using Microsoft.SharePoint.StsAdmin;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.WebPartPages;
using System.Web;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebParts
{
    public class RetargetContentQueryWebPart : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RetargetContentQueryWebPart"/> class.
        /// </summary>
        public RetargetContentQueryWebPart()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPNonEmptyValidator(), "Please specify the page url."));
            parameters.Add(new SPParam("id", "id", false, null, new SPNonEmptyValidator(), "Please specify the web part ID."));
            parameters.Add(new SPParam("title", "t", false, null, new SPNonEmptyValidator(), "Please specify the web part title."));
            parameters.Add(new SPParam("list", "l", false, null, new SPNonEmptyValidator(), "Please specify the list view URL."));
            parameters.Add(new SPParam("listtype", "type", false, null, new SPNonEmptyValidator(), "Please specify the list type."));
            parameters.Add(new SPParam("site", "s", false, null, new SPUrlValidator(), "Please specify the site to show items from."));
            parameters.Add(new SPParam("allmatching", "all", false, null, null));
            parameters.Add(new SPParam("publish", "p", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nRetargets a Content Query web part (do not provide list or site if you wish to show items from all sites in the containing site collection).\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web part page URL>");
            sb.Append("\r\n\t{-id <web part ID> |\r\n\t -title <web part title>}");
            sb.Append("\r\n\t[-list <list view URL>]");
            sb.Append("\r\n\t[-listtype <list type (template) to show items from>]");
            sb.Append("\r\n\t[-site <show items from this site and all subsites>]");
            sb.Append("\r\n\t[-allmatching (if title specified and more than one match found, adjust all matches)");
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

            string url = Params["url"].Value;
            bool allMatching = Params["allmatching"].UserTypedIn;
            string webPartId = Params["id"].Value;
            string webPartTitle = Params["title"].Value;
            string listUrl = Params["list"].Value;
            string listType = Params["listtype"].Value;
            string siteUrl = Params["site"].Value;
            bool publish = Params["publish"].UserTypedIn;

            Common.WebParts.RetargetContentQueryWebPart.Retarget(url, allMatching, webPartId, webPartTitle, listUrl, listType, siteUrl, publish);

            return (int)ErrorCodes.NoError;
        }


        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            SPBinaryParameterValidator.Validate("id", Params["id"].Value, "title", Params["title"].Value);
            if (Params["allmatching"].UserTypedIn && Params["id"].UserTypedIn)
            {
                throw new ArgumentException(SPResource.GetString("IncompatibleParametersSpecified", new object[] { "allmatching", "id" }));
            }
            if (Params["list"].UserTypedIn && Params["site"].UserTypedIn)
            {
                throw new ArgumentException(SPResource.GetString("IncompatibleParametersSpecified", new object[] { "list", "site" }));
            }

            base.Validate(keyValues);
        }
        #endregion

    }
}
#endif