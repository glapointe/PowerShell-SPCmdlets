using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Workflow;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class PublishItems : SPOperation
    {

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PublishItems"/> class.
        /// </summary>
        public PublishItems()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", false, null, new SPNonEmptyValidator(), "Please specify the url to search."));
            parameters.Add(new SPParam("scope", "scope", true, null, new SPRegexValidator("(?i:^Farm$|^WebApplication$|^Site$|^Web$|^List$)")));
            parameters.Add(new SPParam("quiet", "q", false, null, null));
            parameters.Add(new SPParam("test", "t", false, null, null));
            parameters.Add(new SPParam("logfile", "log", false, null, new SPDirectoryExistsAndValidFileNameValidator()));
            parameters.Add(new SPParam("takeoverfileswithnocheckin", "takeover"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nPublishes all items at a given scope.  Use -test to verify what will be published before executing.\r\n\r\nParameters:");
            sb.Append("\r\n\t[-url <url to publish>]");
            sb.Append("\r\n\t-scope <Farm | WebApplication | Site | Web | List>");
            sb.Append("\r\n\t[-quiet]");
            sb.Append("\r\n\t[-test]");
            sb.Append("\r\n\t[-logfile <log file>]");
            sb.Append("\r\n\t[-takeoverfileswithnocheckin (Take over ownership of any files that do not have an existing check-in)]");

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
            bool test = Params["test"].UserTypedIn;
            bool takeOver = Params["takeoverfileswithnocheckin"].UserTypedIn;

            Common.Lists.PublishItems itemsPublisher = new Common.Lists.PublishItems();

            string scope = Params["scope"].Value.ToLowerInvariant();

            if (scope == "farm")
            {
                foreach (SPService svc in SPFarm.Local.Services)
                {
                    if (!(svc is SPWebService))
                        continue;

                    foreach (SPWebApplication webApp in ((SPWebService)svc).WebApplications)
                    {
                        itemsPublisher.Publish(webApp, test, null, takeOver, null);
                    }
                }
            }
            else if (scope == "webapplication")
            {
                SPWebApplication webApp = SPWebApplication.Lookup(new Uri(url));
                itemsPublisher.Publish(webApp, test, null, takeOver, null);
            }
            else if (scope == "site")
            {
                using (SPSite site = new SPSite(url))
                {
                    itemsPublisher.Publish(site, test, null, takeOver, null);
                }
            }
            else if (scope == "web")
            {
                using (SPSite site = new SPSite(url))
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {
                    itemsPublisher.Publish(web, test, null, takeOver, null);
                }
            }
            else if (scope == "list")
            {
                SPList list = Utilities.GetListFromViewUrl(url);
                if (list == null)
                    throw new SPException("List not found.");
                itemsPublisher.Publish(list, test, null, takeOver, null);
            }

            Logger.Write(string.Format("\r\n\r\nFinished Process: {0} Errors, {1} Items(s) Checked In, {2} Item(s) Published, {3} Item(s) Approved",
                itemsPublisher.TaskCounts.Errors, itemsPublisher.TaskCounts.Checkin, itemsPublisher.TaskCounts.Publish, itemsPublisher.TaskCounts.Approve));

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
