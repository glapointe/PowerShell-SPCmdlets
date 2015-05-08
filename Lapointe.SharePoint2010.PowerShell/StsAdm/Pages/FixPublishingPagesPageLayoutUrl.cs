using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Publishing;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Pages
{
    public class FixPublishingPagesPageLayoutUrl : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FixPublishingPagesPageLayoutUrl"/> class.
        /// </summary>
        public FixPublishingPagesPageLayoutUrl()
        {
            SPParamCollection parameters = new SPParamCollection();
            StringBuilder sb = new StringBuilder();

#if MOSS
            parameters.Add(new SPParam("url", "url", false, null, new SPNonEmptyValidator(), "Please specify the url to search."));
            parameters.Add(new SPParam("scope", "scope", true, null, new SPRegexValidator("(?i:^WebApplication$|^Site$|^Web$|^Page$)")));
            parameters.Add(new SPParam("pagelayout", "layout", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("regexsearchstring", "search", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("regexreplacestring", "replace", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("pagename", "page", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("verbose", "v"));
            parameters.Add(new SPParam("fixcontact", "fixcontact"));
            parameters.Add(new SPParam("test", "t"));


            sb.Append("\r\n\r\nFixes the Page Layout URL property of publishing pages which can get messed up during an upgrade or from importing into a new farm.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <url>");
            sb.Append("\r\n\t-scope <WebApplication | Site | Web | Page>");
            sb.Append("\r\n\t[-pagename <if scope is Page, the name of the page to update>]");
            sb.Append("\r\n\t{[-pagelayout <url of page layout to retarget page(s) to (format: \"url, desc\")>] /");
            sb.Append("\r\n\t [-regexsearchstring <search pattern to use for a regular expression replacement of the page layout url>]");
            sb.Append("\r\n\t [-regexreplacestring <replace pattern to use for a regular expression replacement of the page layout url>]}");
            sb.Append("\r\n\t[-fixcontact]");
            sb.Append("\r\n\t[-verbose]");
            sb.Append("\r\n\t[-test]");
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
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(System.Collections.Specialized.StringDictionary keyValues)
        {
#if !MOSS
            return;
#endif
            if (Params["scope"].UserTypedIn)
            {
                string scope = Params["scope"].Value.ToLowerInvariant();

                if (scope == "page")
                    Params["pagename"].IsRequired = true;

                if (scope == "webapplication" && Params["pagelayout"].UserTypedIn)
                    throw new SPSyntaxException(
                        "The pagelayout parameter is incompatible with a scope of WebApplication");
            }
            if (Params["regexsearchstring"].UserTypedIn || Params["regexreplacestring"].UserTypedIn)
            {
                Params["regexsearchstring"].IsRequired = true;
                Params["regexreplacestring"].IsRequired = true;
            }


            if (Params["pagelayout"].UserTypedIn && (Params["regexsearchstring"].UserTypedIn || Params["regexreplacestring"].UserTypedIn))
                throw new SPSyntaxException("The pagelayout parameter is incompatible with the regexsearchstring and regexreplacestring parameters.");

            base.Validate(keyValues);
        }
        /// <summary>
        /// Runs the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public override int Execute(string command, System.Collections.Specialized.StringDictionary keyValues, out string output)
        {
            output = string.Empty;

#if !MOSS
            output = NOT_VALID_FOR_FOUNDATION;
            return (int)ErrorCodes.GeneralError;
#endif

            string url = Params["url"].Value;
            if (url != null)
                url = url.TrimEnd('/');

            string scope = Params["scope"].Value.ToLowerInvariant();
            bool verbose = Params["verbose"].UserTypedIn;
            bool test = Params["test"].UserTypedIn;
            if (test)
                verbose = true;
            Logger.Verbose = verbose;

            Regex regex = null;
            string replaceString = null;
            if (Params["regexsearchstring"].UserTypedIn)
            {
                regex = new Regex(Params["regexsearchstring"].Value);
                replaceString = Params["regexreplacestring"].Value;
            }
            if (scope == "webapplication")
            {
                SPWebApplication webapp = SPWebApplication.Lookup(new Uri(url));
                Logger.Write("Progress: Begin processing web application '{0}'.", url);

                foreach (SPSite site in webapp.Sites)
                {
                    Logger.Write("Progress: Begin processing site '{0}'.", site.ServerRelativeUrl);
                    try
                    {
                        foreach (SPWeb web in site.AllWebs)
                        {
                            Logger.Write("Progress: Begin processing web '{0}'.", web.ServerRelativeUrl);
                    
                            try
                            {
                                PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                                Common.Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb, Params["pagename"].Value, null, regex, replaceString, Params["fixcontact"].UserTypedIn, test);
                            }
                            finally
                            {
                                Logger.Write("Progress: Finished processing web '{0}'.", web.ServerRelativeUrl);

                                web.Dispose();
                            }
                        }
                    }
                    finally
                    {
                        Logger.Write("Progress: Finished processing site '{0}'.", site.ServerRelativeUrl);
                        site.Dispose();
                    }
                }
                Logger.Write("Progress: Finished processing web application '{0}'.", url);
            }
            else if (scope == "site")
            {
                using (SPSite site = new SPSite(url))
                {
                    Logger.Write("Progress: Begin processing site '{0}'.", site.ServerRelativeUrl);

                    foreach (SPWeb web in site.AllWebs)
                    {
                        Logger.Write("Progress: Begin processing web '{0}'.", web.ServerRelativeUrl);

                        try
                        {
                            PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                            Common.Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb, Params["pagename"].Value, Params["pagelayout"].Value, regex, replaceString, Params["fixcontact"].UserTypedIn, test);
                        }
                        finally
                        {
                            Logger.Write("Progress: Finished processing web '{0}'.", web.ServerRelativeUrl);
                            
                            web.Dispose();
                        }
                    }
                    Logger.Write("Progress: Finished processing site '{0}'.", site.ServerRelativeUrl);
                }
            }
            else if (scope == "web")
            {
                using (SPSite site = new SPSite(url))
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {
                    Logger.Write("Progress: Begin processing web '{0}'.", web.ServerRelativeUrl);

                    PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                    Common.Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb, Params["pagename"].Value, Params["pagelayout"].Value, regex, replaceString, Params["fixcontact"].UserTypedIn, test);

                    Logger.Write("Progress: Finished processing web '{0}'.", web.ServerRelativeUrl);
                }
            }
            else if (scope == "page")
            {
                using (SPSite site = new SPSite(url))
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {
                    PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                    Common.Pages.FixPublishingPagesPageLayoutUrl.FixPages(pubweb, Params["pagename"].Value, Params["pagelayout"].Value, regex, replaceString, Params["fixcontact"].UserTypedIn, test);
                }
            }
            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
