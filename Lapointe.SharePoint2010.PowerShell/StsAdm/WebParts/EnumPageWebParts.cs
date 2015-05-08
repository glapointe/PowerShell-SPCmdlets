using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebPartPages;
using WebPart=System.Web.UI.WebControls.WebParts.WebPart;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebParts
{
    public class EnumPageWebParts : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EnumPageWebParts"/> class.
        /// </summary>
        public EnumPageWebParts()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPNonEmptyValidator(), "Please specify the page url."));
            parameters.Add(new SPParam("verbose", "v", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nLists all the web parts that have been added to the specified page.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web part page URL>");
            sb.Append("\r\n\t[-verbose]");

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
            bool verbose = Params["verbose"].UserTypedIn;

            string xml = Common.WebParts.EnumPageWebParts.GetWebPartXml(url, verbose);
            output += xml;

            return (int)ErrorCodes.NoError;
        }

        #endregion

    }
}
