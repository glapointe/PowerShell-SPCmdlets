using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class ReplaceFieldValues : SPOperation
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ReplaceFieldValues"/> class.
        /// </summary>
        public ReplaceFieldValues()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", false, null, new SPNonEmptyValidator(), "Please specify the url to search."));
            parameters.Add(new SPParam("searchstring", "search", false, null, new SPNonEmptyValidator(), "Please specify the search string."));
            parameters.Add(new SPParam("replacestring", "replace", false, null, new SPNullOrNonEmptyValidator(), "Please specify the replace string."));
            parameters.Add(new SPParam("scope", "scope", true, null, new SPRegexValidator("(?i:^Farm$|^WebApplication$|^Site$|^Web$|^List$)")));
            parameters.Add(new SPParam("field", "f", false, null, new SPNonEmptyValidator(), "Please enter the field name."));
            parameters.Add(new SPParam("useinternalfieldname", "userinternal", false, null, null));
            parameters.Add(new SPParam("inputfile", "input", false, null, new SPDirectoryExistsAndValidFileNameValidator()));
            parameters.Add(new SPParam("inputfiledelimiter", "delimiter", false, "|", new SPNonEmptyValidator()));
            parameters.Add(new SPParam("inputfileisxml", "isxml"));
            parameters.Add(new SPParam("quiet", "q"));
            parameters.Add(new SPParam("test", "t"));
            parameters.Add(new SPParam("publish", "p"));
            parameters.Add(new SPParam("logfile", "log", false, null, new SPDirectoryExistsAndValidFileNameValidator()));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nReplaces all occurrences of the search string with the replacement string.  Supports use of regular expressions.  Use -test to verify your replacements before executing.\r\n\r\nParameters:");
            sb.Append("\r\n\t[-url <url to search>]");
            sb.Append("\r\n\t{-inputfile <input file> |");
            sb.Append("\r\n\t -searchstring <regular expression string to search for>");
            sb.Append("\r\n\t -replacestring <replacement string>}");
            sb.Append("\r\n\t-scope <Farm | WebApplication | Site | Web | List>");
            sb.Append("\r\n\t[-field <field name>]");
            sb.Append("\r\n\t[-useinternalfieldname (if not present then the display name will be used)]");
            sb.Append("\r\n\t[-inputfiledelimiter <delimiter character to use in the input file (default is \"|\")>]");
            sb.Append("\r\n\t[-inputfileisxml (input is XML in the following format: <Replacements><Replacement><SearchString>string</SearchString><ReplaceString>string</ReplaceString></Replacement></Replacements>)");
            sb.Append("\r\n\t[-quiet]");
            sb.Append("\r\n\t[-test]");
            sb.Append("\r\n\t[-logfile <log file>]");
            sb.Append("\r\n\t[-publish]");

            Init(parameters, sb.ToString());
        }

        #endregion

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

            Common.Lists.ReplaceFieldValues rfv = new Common.Lists.ReplaceFieldValues();
            rfv.LogFile = Params["logfile"].Value;
            if (Params["field"].UserTypedIn)
                rfv.FieldName.Add(Params["field"].Value);
            rfv.Publish = Params["publish"].UserTypedIn;
            rfv.Quiet = Params["quiet"].UserTypedIn;
            rfv.Test = Params["test"].UserTypedIn;
            rfv.UseInternalFieldName = Params["useinternalfieldname"].UserTypedIn;
            
            if (Params["inputfile"].UserTypedIn)
            {
                rfv.ParseInputFile(Params["inputfile"].Value, Params["inputfiledelimiter"].Value, Params["inputfileisxml"].UserTypedIn);
            }
            else
            {
                rfv.SearchStrings.Add(new Common.Lists.ReplaceFieldValues.SearchReplaceData(Params["searchstring"].Value, Params["replacestring"].Value));
            }
            if (rfv.SearchStrings.Count == 0)
                throw new SPException("No search strings were specified.");

            string scope = Params["scope"].Value.ToLowerInvariant();

            if (scope == "farm")
            {
                foreach (SPService svc in SPFarm.Local.Services)
                {
                    if (!(svc is SPWebService))
                        continue;

                    foreach (SPWebApplication webApp in ((SPWebService)svc).WebApplications)
                    {
                        rfv.ReplaceValues(webApp);
                    }
                }
            }
            else if (scope == "webapplication")
            {
                SPWebApplication webApp = SPWebApplication.Lookup(new Uri(url));
                rfv.ReplaceValues(webApp);
            }
            else if (scope == "site")
            {
                using (SPSite site = new SPSite(url))
                {
                    rfv.ReplaceValues(site);
                }
            }
            else if (scope == "web")
            {
                using (SPSite site = new SPSite(url))
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {
                    rfv.ReplaceValues(web);
                }
            }
            else if (scope == "list")
            {
                SPList list = Utilities.GetListFromViewUrl(url);
                rfv.ReplaceValues(list);
            }

            return (int)ErrorCodes.NoError;
        }

        #endregion

        #region SPOperation Misc Overrides

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            if (Params["quiet"].UserTypedIn && Params["test"].UserTypedIn && !Params["logfile"].UserTypedIn)
                throw new SPSyntaxException("The quiet parameter and the test parameter are incompatible if no log file is specified.");

            if (Params["scope"].Value.ToLowerInvariant() == "farm" && Params["url"].UserTypedIn)
                throw new SPSyntaxException("The url parameter is not compatible with a scope of Farm.");
            if (Params["scope"].Value.ToLowerInvariant() != "farm" && !Params["url"].UserTypedIn)
                throw new SPSyntaxException("The url parameter is required if the scope is not Farm.");

            if (Params["inputfile"].UserTypedIn && !File.Exists(Params["inputfile"].Value))
                throw new SPSyntaxException("The inputfile was not found.");

            if ((Params["searchstring"].UserTypedIn || Params["replacestring"].UserTypedIn) && Params["inputfile"].UserTypedIn)
                throw new SPSyntaxException("The searchstring and replacestring parameters are incompatible with the inputfile parameter.");

            if (!(Params["searchstring"].UserTypedIn || Params["replacestring"].UserTypedIn) && !Params["inputfile"].UserTypedIn)
                throw new SPSyntaxException("You must specify at least the searchstring and replacestring or the inputfile parameter.");

            if (!Params["inputfile"].UserTypedIn)
            {
                Params["searchstring"].IsRequired = true;
                Params["replacestring"].IsRequired = true;
            }

            if (Params["searchstring"].UserTypedIn)
            {
                Params["inputfile"].IsRequired = false;
                Params["replacestring"].IsRequired = true;
            }

            if (Params["inputfiledelimiter"].UserTypedIn && Params["inputfileisxml"].UserTypedIn)
                throw new SPSyntaxException("The parameters inputfileisxml and inputfiledelimiter are incompatible.");

            if (!Params["inputfile"].UserTypedIn && (Params["inputfiledelimiter"].UserTypedIn || Params["inputfileisxml"].UserTypedIn))
                throw new SPSyntaxException(
                    "The paramerts inputfiledelimiter and inputfileisxml are only valid when inputfile is used.");

            base.Validate(keyValues);
        }

        #endregion

    }
}
