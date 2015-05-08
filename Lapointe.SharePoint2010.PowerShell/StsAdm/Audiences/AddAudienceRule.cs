using System;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
#if MOSS
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
using Microsoft.Office.Server.Search.Administration;
#endif
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.Common.Audiences;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Audiences
{
    public class AddAudienceRule : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AddAudienceRule"/> class.
        /// </summary>
        public AddAudienceRule()
        {
            SPParamCollection parameters = new SPParamCollection();
            StringBuilder sb = new StringBuilder();

#if MOSS
            SPEnumValidator appendOpValidator = new SPEnumValidator(typeof(AppendOp));

            parameters.Add(new SPParam("name", "n", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("serviceappname", "app", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("sitesubscriptionid", "subid", false, Guid.Empty.ToString(), new SPGuidValidator()));
            parameters.Add(new SPParam("contextsite", "site", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("rules", "r", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("rulesfile", "rf", false, null, new SPFileExistsValidator()));
            parameters.Add(new SPParam("clear", "cl"));
            parameters.Add(new SPParam("compile", "co"));
            parameters.Add(new SPParam("groupexisting", "group"));
            parameters.Add(new SPParam("appendop", "op", false, "and", appendOpValidator));

            sb.Append("\r\n\r\nAdds simple or complex rules to an existing audience.  The rules XML should be in the following format: ");
            sb.Append("<rules><rule op='' field='' value='' /></rules>\r\n");
            sb.Append("Values for the \"op\" attribute can be any of \"=,>,>=,<,<=,<>,Contains,Not contains,Reports Under,Member Of,AND,OR,(,)\"\r\n");
            sb.Append("The \"field\" attribute is not required if \"op\" is any of \"Reports Under,Member Of,AND,OR,(,)\"\r\n");
            sb.Append("The \"value\" attribute is not required if \"op\" is any of \"AND,OR,(,)\"\r\n");
            sb.Append("Note that if your rules contain any grouping or mixed logic then you will not be able to manage the rule via the browser.\r\n");
            sb.Append("Example: <rules><rule op='Member of' value='sales department' /><rule op='AND' /><rule op='Contains' field='Department' value='Sales' /></rules>");
            sb.Append("\r\n\r\nParameters:");
            sb.Append("\r\n\t-name <audience name>");
            sb.Append("\r\n\t-rules <rules xml> | -rulesfile <xml file containing the rules>");
            sb.Append("\r\n\t[-serviceappname <user profile service application name>]");
            sb.Append("\r\n\t[-sitesubscriptionid <GUID>]");
            sb.Append("\r\n\t[-contextsite <URL of the site to use for service context>]");
            sb.Append("\r\n\t[-clear (clear existing rules)]");
            sb.Append("\r\n\t[-compile]");
            sb.Append("\r\n\t[-groupexisting (wraps any existing rules in parantheses)]");
            sb.Append("\r\n\t[-appendop <and (default) | or> (operator used to append to existing rules)]");
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

            string rules;
            if (Params["rules"].UserTypedIn)
                rules = Params["rules"].Value;
            else
                rules = File.ReadAllText(Params["rulesfile"].Value);

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
            Common.Audiences.AddAudienceRule.AddRules(context,
                     Params["name"].Value,
                     rules,
                     Params["clear"].UserTypedIn,
                     Params["compile"].UserTypedIn,
                     Params["groupexisting"].UserTypedIn,
                     (AppendOp)Enum.Parse(typeof(AppendOp), Params["appendop"].Value, true));

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            SPBinaryParameterValidator.Validate("rules", Params["rules"].Value, "rulesfile", Params["rulesfile"].Value);
            SPBinaryParameterValidator.Validate("serviceappname", Params["serviceappname"].Value, "contextsite", Params["contextsite"].Value);
            
            if (Params["clear"].UserTypedIn && (Params["appendop"].UserTypedIn || Params["groupexisting"].UserTypedIn))
                throw new SPSyntaxException("The -clear parameter cannot be used with the -appendop or -groupexisting parameters.");

            base.Validate(keyValues);
        }

    }
}
