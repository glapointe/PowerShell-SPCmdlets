using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebApplications
{
    public class AddUserPolicyForWebApp : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AddUserPolicyForWebApp"/> class.
        /// </summary>
        public AddUserPolicyForWebApp()
        {
            SPEnumValidator urlZoneValidator = new SPEnumValidator(typeof(SPUrlZone), new string[] {"All"});
            
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the web application url."));
            parameters.Add(new SPParam("zone", "z", true, null, urlZoneValidator));
            parameters.Add(new SPParam("userlogin", "u", true, null, new SPNonEmptyValidator(), "Please specify the users login."));
            parameters.Add(new SPParam("username", "username", true, null, new SPNonEmptyValidator(), "Please specify the users name."));
            parameters.Add(new SPParam("permissions", "p", true, null, new SPNonEmptyValidator()));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nAdds a user policy for a web application.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web application url>");
            sb.AppendFormat("\r\n\t-zone <{0}>", urlZoneValidator.DisplayValue);
            sb.Append("\r\n\t-userlogin <DOMAIN\\user>");
            sb.Append("\r\n\t-username <display name>");
            sb.Append("\r\n\t-permissions <comma separated list of policy permissions>");

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

            string url = Params["url"].Value.TrimEnd('/');
            string login = Params["userlogin"].Value;
            string username = Params["username"].Value;
            string[] permissions = Params["permissions"].Value.Split(',');
            string zone = Params["zone"].Value;

            Common.WebApplications.AddUserPolicyForWebApp.AddUserPolicy(url, login, username, permissions, zone);

            return (int)ErrorCodes.NoError;
        }

        #endregion

        


    }
}
