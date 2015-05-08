using System;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.StsAdmin;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Permissions
{
    public class SetSelfServiceSiteCreation : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SetSelfServiceSiteCreation"/> class.
        /// </summary>
        public SetSelfServiceSiteCreation()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the web application url."));
            parameters.Add(new SPParam("enabled", "e", false, null, new SPTrueFalseValidator(), "Please specify \"true\" or \"false\" for enabled."));
            parameters.Add(new SPParam("requiresecondarycontact", "r", false, null, new SPTrueFalseValidator(), "Please specify \"true\" or \"false\" for requiresecondarycontact."));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nSets whether self service site creation is enabled for the web application.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web application>");
            sb.Append("\r\n\t[-enabled <true|false>]");
            sb.Append("\r\n\t[-requiresecondarycontact <true|false>]");
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

            bool enabledProvided = Params["enabled"].UserTypedIn;
            bool requireContactProvided = Params["requiresecondarycontact"].UserTypedIn;
            bool enabled = false;
            bool requireSecondaryContact = false;

            if (enabledProvided)
                enabled = bool.Parse(Params["enabled"].Value);
            if (requireContactProvided)
                requireSecondaryContact = bool.Parse(Params["requiresecondarycontact"].Value);

            if (!enabledProvided && !requireContactProvided)
            {
                output = "Please specify at least one parameter.";
                return (int)ErrorCodes.SyntaxError;
            }

            string url = Params["url"].Value.TrimEnd('/');

            SPWebApplication wa = SPWebApplication.Lookup(new Uri(url));

            if (enabledProvided)
                wa.SelfServiceSiteCreationEnabled = enabled;
            if (requireContactProvided)
                wa.RequireContactForSelfServiceSiteCreation = requireSecondaryContact;

            wa.Update();

            return (int)ErrorCodes.NoError;
        }

        #endregion

    }
}
