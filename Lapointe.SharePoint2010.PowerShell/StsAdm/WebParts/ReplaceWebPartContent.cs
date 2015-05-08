using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Text.RegularExpressions;
#if MOSS
using Microsoft.SharePoint.Portal.WebControls;
using Microsoft.SharePoint.Publishing;
using Microsoft.SharePoint.Publishing.Fields;
using Microsoft.SharePoint.Publishing.WebControls;
#endif
using Microsoft.SharePoint.WebPartPages;
using WebPart=Microsoft.SharePoint.WebPartPages.WebPart;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebParts
{
    public class ReplaceWebPartContent : SPOperation
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ReplaceWebPartContent"/> class.
        /// </summary>
        public ReplaceWebPartContent()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", false, null, new SPNonEmptyValidator(), "Please specify the url to search."));
            parameters.Add(new SPParam("searchstring", "search", true, null, new SPNonEmptyValidator(), "Please specify the search string."));
            parameters.Add(new SPParam("replacestring", "replace", true, null, new SPNullOrNonEmptyValidator(), "Please specify the search string."));
            parameters.Add(new SPParam("scope", "scope", true, null, new SPRegexValidator("(?i:^Farm$|^WebApplication$|^Site$|^Web$|^Page$)")));
            parameters.Add(new SPParam("webpartname", "part", false, null, new SPNonEmptyValidator(), "Please enter the web part name."));
            parameters.Add(new SPParam("quiet", "q", false, null, null));
            parameters.Add(new SPParam("test", "t", false, null, null));
            parameters.Add(new SPParam("publish", "p", false, null, null));
            parameters.Add(new SPParam("logfile", "log", false, null, new SPDirectoryExistsAndValidFileNameValidator()));
            parameters.Add(new SPParam("unsafexml", "unsafexml"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nReplaces all occurances of the search string with the replacement string.  Supports use of regular expressions.  Use -test to verify your replacements before executing.\r\n\r\nParameters:");
            sb.Append("\r\n\t[-url <url to search>]");
            sb.Append("\r\n\t-searchstring <regular expression string to search for>");
            sb.Append("\r\n\t-replacestring <replacement string>");
            sb.Append("\r\n\t-scope <Farm | WebApplication | Site | Web | Page>");
            sb.Append("\r\n\t[-webpartname <web part name>]");
            sb.Append("\r\n\t[-quiet]");
            sb.Append("\r\n\t[-test]");
            sb.Append("\r\n\t[-logfile <log file>]");
            sb.Append("\r\n\t[-publish]");
            sb.Append("\r\n\t[-unsafexml (treats known XML data as a string)]");

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

            Logger.LogFile = Params["logfile"].Value;
            Logger.Verbose = !Params["quiet"].UserTypedIn;

            Common.WebParts.ReplaceWebPartContent.Settings settings = new Common.WebParts.ReplaceWebPartContent.Settings();
            settings.SearchString = Params["searchstring"].Value;
            settings.ReplaceString = Params["replacestring"].Value;
            settings.WebPartName = Params["webpartname"].Value;
            settings.Publish = Params["publish"].UserTypedIn;
            settings.Test = Params["test"].UserTypedIn;
            settings.UnsafeXml = Params["unsafexml"].UserTypedIn;

            string scope = Params["scope"].Value.ToLowerInvariant();

            if (scope == "farm")
            {
                foreach (SPService svc in SPFarm.Local.Services)
                {
                    if (!(svc is SPWebService))
                        continue;

                    foreach (SPWebApplication webApp in ((SPWebService)svc).WebApplications)
                    {
                        Common.WebParts.ReplaceWebPartContent.ReplaceValues(webApp, settings);
                    }
                }
            }
            else if (scope == "webapplication")
            {
                SPWebApplication webApp = SPWebApplication.Lookup(new Uri(url));
                Common.WebParts.ReplaceWebPartContent.ReplaceValues(webApp, settings);
            }
            else if (scope == "site")
            {
                using (SPSite site = new SPSite(url))
                {
                    Common.WebParts.ReplaceWebPartContent.ReplaceValues(site, settings);
                }
            }
            else if (scope == "web")
            {
                using (SPSite site = new SPSite(url))
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {
                    Common.WebParts.ReplaceWebPartContent.ReplaceValues(web, settings);
                }
            }
            else if (scope == "page")
            {
                using (SPSite site = new SPSite(url))
                using (SPWeb web = site.OpenWeb())
                {
                    Common.WebParts.ReplaceWebPartContent.ReplaceValues(web, web.GetFile(url), settings);
                }

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

            if (Params["scope"].UserTypedIn)
            {
                if (Params["scope"].Value.ToLowerInvariant() == "farm" && Params["url"].UserTypedIn)
                    throw new SPSyntaxException("The url parameter is not compatible with a scope of Farm.");
                if (Params["scope"].Value.ToLowerInvariant() != "farm" && !Params["url"].UserTypedIn)
                    throw new SPSyntaxException("The url parameter is required if the scope is not Farm.");
            }
            base.Validate(keyValues);
        }

        #endregion

    }
}
