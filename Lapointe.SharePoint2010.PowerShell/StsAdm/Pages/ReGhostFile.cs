using System;
using System.IO;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using Lapointe.SharePoint.PowerShell.StsAdm.Lists;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using System.Management.Automation;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Pages
{
    public class ReGhostFile : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ReGhostFile"/> class.
        /// </summary>
        public ReGhostFile()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the file url."));
            parameters.Add(new SPParam("force", "force", false, null, null)); // Left if for compatibility but not used.
            parameters.Add(new SPParam("scope", "scope", false, "file", new SPRegexValidator("(?i:^WebApplication$|^Site$|^Web$|^List$|^File$)")));
            parameters.Add(new SPParam("recursewebs", "recurse"));
            parameters.Add(new SPParam("haltonerror", "halt"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nReghosts a file (use force to override CustomizedPageStatus check).\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <url to analyze>");
            //sb.Append("\r\n\t[-force]");
            sb.Append("\r\n\t[-scope <WebApplication | Site | Web | List | File>]");
            sb.Append("\r\n\t[-recursewebs (applies to Web scope only)]");
            sb.Append("\r\n\t[-haltonerror]");
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
        public override int Execute(string command, System.Collections.Specialized.StringDictionary keyValues, out string output)
        {
            output = string.Empty;
            Logger.Verbose = true;
            

            string url = Params["url"].Value;
            bool force = false; // Params["force"].UserTypedIn;
            string scope = Params["scope"].Value.ToLowerInvariant();
            bool haltOnError = Params["haltonerror"].UserTypedIn;

            switch (scope)
            {
                case "file":
                    using (SPSite site = new SPSite(url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPFile file = web.GetFile(url);
                        if (file == null)
                        {
                            throw new FileNotFoundException(string.Format("File '{0}' not found.", url), url);
                        }

                        Common.Pages.ReGhostFile.Reghost(site, web, file, force, haltOnError);
                    }
                    break;
                case "list":
                    using (SPSite site = new SPSite(url))
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList list = Utilities.GetListFromViewUrl(web, url);
                        Common.Pages.ReGhostFile.ReghostFilesInList(site, web, list, force, haltOnError);
                    }
                    break;
                case "web":
                    bool recurseWebs = Params["recursewebs"].UserTypedIn;
                    using (SPSite site = new SPSite(url))
                    using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                    {
                        Common.Pages.ReGhostFile.ReghostFilesInWeb(site, web, recurseWebs, force, haltOnError);
                    }
                    break;
                case "site":
                    using (SPSite site = new SPSite(url))
                    {
                        Common.Pages.ReGhostFile.ReghostFilesInSite(site, force, haltOnError);
                    }
                    break;
                case "webapplication":
                    SPWebApplication webApp = SPWebApplication.Lookup(new Uri(url));
                    Logger.Write("Progress: Analyzing files in web application '{0}'.", url);

                    foreach (SPSite site in webApp.Sites)
                    {
                        try
                        {
                            Common.Pages.ReGhostFile.ReghostFilesInSite(site, force, haltOnError);
                        }
                        finally
                        {
                            site.Dispose();
                        }
                    }
                    break;
                    
            }
            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
