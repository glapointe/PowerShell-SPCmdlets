using System;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Security
{
    public class AddUser2 : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AddUser2"/> class.
        /// </summary>
        public AddUser2()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the web url."));
            parameters.Add(new SPParam("userlogin", "u", true, null, new SPNonEmptyValidator(), "Please specify the users login."));
            parameters.Add(new SPParam("useremail", "email", false, null, new SPRegexValidator(@"^[^ \r\t\n\f@]+@[^ \r\t\n\f@]+$"), "Please specify the users email address."));
            parameters.Add(new SPParam("username", "username", false, null, new SPNonEmptyValidator(), "Please specify the users name."));
            parameters.Add(new SPParam("role", "r", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("group", "g", false, null, new SPNonEmptyValidator()));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nAdds a user to a site (allows for useremail and username to be optional).\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web url>");
            sb.Append("\r\n\t-userlogin <DOMAIN\\user>");
            sb.Append("\r\n\t[-useremail <someone@example.com>]");
            sb.Append("\r\n\t[-username <display name>]");
            sb.Append("\r\n\t[-role <role name> / -group <group name>]");
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

            if (Params["role"].UserTypedIn && Params["group"].UserTypedIn)
                throw new SPException(SPResource.GetString("ExclusiveArgs", new object[] { "role, group" }));

            string url = Params["url"].Value.TrimEnd('/');
            string login = Params["userlogin"].Value;
            string email = Params["useremail"].Value;
            string username = Params["username"].Value;

            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            {
                
                login = Utilities.TryGetNT4StyleAccountName(login, web.Site.WebApplication);
                // First lets see if our user already exists.
                SPUser user = null;
                try
                {
                    user = web.AllUsers[login];
                }
                catch (SPException) { }

                if (user == null)
                {
                    web.SiteUsers.Add(login, email, username, string.Empty);
                    user = web.AllUsers[login];
                }

                if (Params["role"].UserTypedIn)
                {
                    SPRoleDefinition roleDefinition = null;
                    try
                    {
                        roleDefinition = web.RoleDefinitions[Params["role"].Value];
                    }
                    catch (ArgumentException) {}

                    if (roleDefinition == null)
                        throw new SPException("The specified role does not exist.");

                    SPRoleDefinitionBindingCollection roleDefinitionBindings = new SPRoleDefinitionBindingCollection();
                    roleDefinitionBindings.Add(roleDefinition);
                    SPRoleAssignment roleAssignment = new SPRoleAssignment(user);
                    roleAssignment.ImportRoleDefinitionBindings(roleDefinitionBindings);
                    web.RoleAssignments.Add(roleAssignment);
                }
                else if (Params["group"].UserTypedIn)
                {
                    SPGroup group = null;
                    try
                    {
                        group = web.SiteGroups[Params["group"].Value];
                    }
                    catch (ArgumentException) {}

                    if (group == null)
                        throw new SPException("The specified group does not exist.");

                    group.AddUser(user);
                }
            }

            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
