using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Reflection;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Permissions
{
    public class EnumEffectiveBasePerms : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EnumEffectiveBasePerms"/> class.
        /// </summary>
        public EnumEffectiveBasePerms()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site collection."));
            parameters.Add(new SPParam("user", "user", false, null, new SPNonEmptyValidator(), "Please specify the username."));
            parameters.Add(new SPParam("invert", "i", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nLists the effective base permissions for a user.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web url>");
            sb.Append("\r\n\t[-user <DOMAIN\\name>]");
            sb.Append("\r\n\t[-invert (shows what base permissions the user is missing)]");
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
            if (url != null)
                url = url.TrimEnd('/');

            SPBasePermissions perms;

            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            {
                if (Params["user"].UserTypedIn)
                {
                    perms = web.GetUserEffectivePermissions(Params["user"].Value);
                }
                else
                {
                    perms = web.EffectiveBasePermissions;
                }
            }

            if (!Params["invert"].UserTypedIn)
                output += OutputPerms(perms);
            else
            {
                List<string> permsInverted = new List<string>();

                foreach (SPBasePermissions perm in Enum.GetValues(typeof(SPBasePermissions)))
                {
                    if ((perms & perm) != perm)
                    {
                        permsInverted.Add(perm.ToString());
                    }
                }
                output += string.Join(", ", permsInverted.ToArray());
            }
            if (output == string.Empty)
            {
                output += "No permissions were found.";
            }

            return (int)ErrorCodes.NoError;
        }

        #endregion

        /// <summary>
        /// Outputs the perms.
        /// </summary>
        /// <param name="perms">The perms.</param>
        /// <returns></returns>
        private static string OutputPerms(SPBasePermissions perms)
        {
            string output = string.Empty;
            if ((SPBasePermissions.FullMask & perms) == perms)
            {
                return SPBasePermissions.FullMask.ToString();
            }
            foreach (SPBasePermissions perm in Enum.GetValues(typeof(SPBasePermissions)))
            {
                if ((perms & perm) == perm)
                {
                    output += perm + ", ";
                }
            }
            return output.Substring(0, output.Length - 2);
        }
    }
}
