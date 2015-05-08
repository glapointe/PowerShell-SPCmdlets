using System;
using System.Collections.Specialized;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class ListAudienceTargeting : SPOperation
    {
         /// <summary>
        /// Initializes a new instance of the <see cref="ListAudienceTargeting"/> class.
        /// </summary>
        public ListAudienceTargeting()
        {
            SPParamCollection parameters = new SPParamCollection();
            StringBuilder sb = new StringBuilder();

#if MOSS
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the list view URL."));
            parameters.Add(new SPParam("enabled", "e", true, null, new SPTrueFalseValidator()));

            sb.Append("\r\n\r\nEnabling audience targeting will create a targeting column for the list. Web parts, such as the Content Query Web Part, can use this data to filter list contents based on the user's context.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <list view url>");
            sb.Append("\r\n\t-enabled <true|false>");
#else
            sb.Append(NOT_VALID_FOR_FOUNDATION);
#endif
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

#if !MOSS
            output = NOT_VALID_FOR_FOUNDATION;
            return (int)ErrorCodes.GeneralError;
#endif

            string url = Params["url"].Value;
            bool enabled = bool.Parse(Params["enabled"].Value);

            Common.Lists.ListAudienceTargeting.SetTargeting(url, enabled);

            return (int)ErrorCodes.NoError;
        }

        #endregion

    }
}
