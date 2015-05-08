using System;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Permissions
{
    public class DisableUserPermissionForWebApp : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DisableUserPermissionForWebApp"/> class.
        /// 
        /// To remove a permission using PowerShell use -bxor:
        /// $webapp.RightsMask = $webapp.RightsMask -bxor [Microsoft.SharePoint.SPBasePermissions]::EditMyUserInfo
        /// 
        /// To add a permission using PowerShell use -bor:
        /// $webapp.RightsMask = $webapp.RightsMask -bor [Microsoft.SharePoint.SPBasePermissions]::EditMyUserInfo
        /// </summary>
        public DisableUserPermissionForWebApp()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the web application url."));
            foreach (string name in Enum.GetNames(typeof(SPBasePermissions)))
            {
                parameters.Add(new SPParam(name, name.ToLowerInvariant(), false, null, null));
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nDisable permissions that can be used in permission levels within the web application.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web application>");
            foreach (string name in Enum.GetNames(typeof(SPBasePermissions)))
            {
                sb.Append("\r\n\t[-");
                sb.Append(name);
                sb.Append("]");
            }
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

            SPWebApplication wa = SPWebApplication.Lookup(new Uri(url));

            foreach (SPBasePermissions perm in Enum.GetValues(typeof(SPBasePermissions)))
            {
                if (Params[perm.ToString()].UserTypedIn)
                    wa.RightsMask = wa.RightsMask & ~perm;
            }

            wa.Update();

            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
