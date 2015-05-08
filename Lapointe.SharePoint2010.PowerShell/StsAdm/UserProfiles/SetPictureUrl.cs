#if MOSS
using System;
using System.Collections.Specialized;
using System.Management.Automation;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.Office.Server;
using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;
using System.Net;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.UserProfiles
{
    public class SetPictureUrl : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SetPictureUrl"/> class.
        /// </summary>
        public SetPictureUrl()
        {
            SPParamCollection parameters = new SPParamCollection();
#if MOSS

            parameters.Add(new SPParam("path", "p", true, null, new SPNullOrNonEmptyValidator(), "Please specify the path."));
            parameters.Add(new SPParam("serviceappname", "app", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("sitesubscriptionid", "subid", false, Guid.Empty.ToString(), new SPGuidValidator()));
            parameters.Add(new SPParam("contextsite", "site", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("username", "u", false, null, new SPNonEmptyValidator(), "Please specify the username."));
            parameters.Add(new SPParam("overwrite", "ow"));
            parameters.Add(new SPParam("ignoremissingdata", "ignore"));
            parameters.Add(new SPParam("validateurl", "validate"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nSets the picture URL path for user profiles.  The following variables may be used for dynamic replacement: \"$(username)\", \"$(domain)\", \"$(email)\", \"$(firstname)\", \"$(lastname)\", \"$(employeeid)\".\r\n\r\nParameters:");
            sb.Append("\r\n\t-path <path to new photo (i.e., \"http://intranet/hr/EmployeePictures/$(username).jpg\") - leave blank to clear>");
            sb.Append("\r\n\t[-serviceappname <user profile service application name>]");
            sb.Append("\r\n\t[-sitesubscriptionid <GUID>]");
            sb.Append("\r\n\t[-contextsite <URL of the site to use for service context>]");
            sb.Append("\r\n\t[-username <DOMAIN\\name>]");
            sb.Append("\r\n\t[-overwrite]");
            sb.Append("\r\n\t[-ignoremissingdata]");
            sb.Append("\r\n\t[-validateurl]");
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

            string username = null;

            if (Params["username"].UserTypedIn)
                username = Params["username"].Value;
            string path = Params["path"].Value;

            SPServiceContext context = null;
            if (Params["serviceappname"].UserTypedIn)
            {
                SPSiteSubscriptionIdentifier subId = Utilities.GetSiteSubscriptionId(new Guid(Params["sitesubscriptionid"].Value));
                SPServiceApplication svcApp = Utilities.GetUserProfileServiceApplication(Params["serviceappname"].Value);
                Utilities.GetServiceContext(svcApp, subId);
            }
            else
            {
                using (SPSite site = new SPSite(Params["contextsite"].Value))
                    context = SPServiceContext.GetContext(site);
            }

            bool overwrite = Params["overwrite"].UserTypedIn;
            bool ignoreMissingData = Params["ignoremissingdata"].UserTypedIn;
            bool validateUrl = Params["validateurl"].UserTypedIn;

            UserProfileManager profManager = new UserProfileManager(context);

            if (string.IsNullOrEmpty(username))
                Common.UserProfiles.SetPictureUrl.SetPictures(profManager, path, overwrite, ignoreMissingData, validateUrl);
            else
                Common.UserProfiles.SetPictureUrl.SetPicture(profManager, username, path, overwrite, ignoreMissingData, validateUrl);

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            SPBinaryParameterValidator.Validate("serviceappname", Params["serviceappname"].Value, "contextsite", Params["contextsite"].Value);

            base.Validate(keyValues);
        }
        #endregion
    }
}
#endif