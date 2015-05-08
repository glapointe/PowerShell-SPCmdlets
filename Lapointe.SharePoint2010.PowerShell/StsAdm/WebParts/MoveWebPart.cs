using System.Collections.Specialized;
using System.Text;
using Lapointe.SharePoint.PowerShell.Common.Lists;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;
using Microsoft.SharePoint.WebPartPages;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebParts
{
    public class MoveWebPart : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MoveWebPart"/> class.
        /// </summary>
        public MoveWebPart()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPNonEmptyValidator(), "Please specify the page url."));
            parameters.Add(new SPParam("id", "id", false, null, new SPNonEmptyValidator(), "Please specify the web part ID."));
            parameters.Add(new SPParam("title", "t", false, null, new SPNonEmptyValidator(), "Please specify the web part title."));
            parameters.Add(new SPParam("zone", "z", false, null, new SPNonEmptyValidator(), "Please specify the zone to move to."));
            parameters.Add(new SPParam("zoneindex", "zi", false, null, new SPIntRangeValidator(0, int.MaxValue)));
            parameters.Add(new SPParam("publish", "p", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nMoves a web part on a page.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web part page URL>");
            sb.Append("\r\n\t{-id <web part ID> |\r\n\t -title <web part title>}");
            sb.Append("\r\n\t[-zone <zone ID>]");
            sb.Append("\r\n\t[-zoneindex <zone index>]");
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

            

            SPBinaryParameterValidator.Validate("id", Params["id"].Value, "title", Params["title"].Value);
            if (!Params["zone"].UserTypedIn && !Params["zoneindex"].UserTypedIn)
                throw new SPSyntaxException("You must specify at least the zone or zoneindex parameters.");

            string url = Params["url"].Value;
            string webPartId = Params["id"].Value;
            string webPartTitle = Params["title"].Value;
            string webPartZone = Params["zone"].Value;
            string webPartZoneIndex = Params["zoneindex"].Value;
            bool publish = Params["publish"].UserTypedIn;

            if (Params["title"].UserTypedIn)
                Common.WebParts.MoveWebPart.MoveByTitle(url, webPartTitle, webPartZone, webPartZoneIndex, publish);
            else if (Params["id"].UserTypedIn)
                Common.WebParts.MoveWebPart.MoveById(url, webPartId, webPartZone, webPartZoneIndex, publish);

            return (int)ErrorCodes.NoError;
        }


        #endregion
    }
}
