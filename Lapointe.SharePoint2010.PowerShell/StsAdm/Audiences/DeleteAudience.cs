using System.Collections;
using System.Collections.Specialized;
using System.Text;
#if MOSS
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
#endif
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using System;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Audiences
{
    public class DeleteAudience : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteAudience"/> class.
        /// </summary>
        public DeleteAudience()
        {
            SPParamCollection parameters = new SPParamCollection();
            StringBuilder sb = new StringBuilder();

#if MOSS
            parameters.Add(new SPParam("name", "n", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("serviceappname", "app", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("sitesubscriptionid", "subid", false, Guid.Empty.ToString(), new SPGuidValidator()));
            parameters.Add(new SPParam("contextsite", "site", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("rulesonly", "ro"));

            sb.Append("\r\n\r\nDeletes an audience.  To delete only the rules associated with the audience pass in the -rulesonly parameter.\r\n\r\nParameters:");
            sb.Append("\r\n\t-name <audience name>");
            sb.Append("\r\n\t[-serviceappname <user profile service application name>]");
            sb.Append("\r\n\t[-sitesubscriptionid <GUID>]");
            sb.Append("\r\n\t[-contextsite <URL of the site to use for service context>]");
            sb.Append("\r\n\t[-rulesonly (delete only the rules)]");
#else
            sb.Append(NOT_VALID_FOR_FOUNDATION);
#endif

            Init(parameters, sb.ToString());
        }

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
        /// Executes the specified command.
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

            Common.Audiences.DeleteAudience.Delete(context,
                   Params["name"].Value,
                   Params["rulesonly"].UserTypedIn);


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
    }
}
