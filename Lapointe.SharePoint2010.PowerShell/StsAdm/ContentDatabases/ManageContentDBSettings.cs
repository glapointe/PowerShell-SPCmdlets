using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
#if SP2010
using Microsoft.SharePoint.Search.Administration;
#elif SP2013
using Microsoft.SharePoint.Search.Administration;
#endif
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.ContentDatabases
{
    public class ManageContentDBSettings : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ManageContentDBSettings"/> class.
        /// </summary>
        public ManageContentDBSettings()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("dbname", "db", true, null, new SPNonEmptyValidator(), "Please specify the database name."));
            parameters.Add(new SPParam("webapp", "web", true, null, new SPUrlValidator(), "Please specify the web application."));
            parameters.Add(new SPParam("status", "s", false, null, new SPRegexValidator("^online$|^disabled$")));
            parameters.Add(new SPParam("maxsites", "max", false, null, new SPIntRangeValidator(0, Int32.MaxValue)));
            parameters.Add(new SPParam("setmaxsitestocurrent", "settocurrent", false, null, null));
            parameters.Add(new SPParam("warningsitecount", "warn", false, null, new SPIntRangeValidator(0, Int32.MaxValue)));
            parameters.Add(new SPParam("searchserver", "search", false, null, new SPNullOrNonEmptyValidator()));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nSets the status and site limits for a content database.\r\n\r\nParameters:");
            sb.Append("\r\n\t-dbname <content database name>");
            sb.Append("\r\n\t-webapp <web application url>");
            sb.Append("\r\n\t[-status <online | disabled>]");
            sb.Append("\r\n\t[-maxsites <maximum number of sites allowed in the db> / -setmaxsitestocurrent (sets the max value to the current value to keep the database online but prevent new sites from beeing added)]");
            sb.Append("\r\n\t[-warningsitecount <number of sites before a warning event is generated>]");
            sb.Append("\r\n\t[-searchserver <search server (leave empty to clear the search server)>]");
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
            Logger.Verbose = true;

            if (Params["setmaxsitestocurrent"].UserTypedIn && Params["maxsites"].UserTypedIn)
            {
                throw new SPException(SPResource.GetString("ExclusiveArgs", new object[] { "setmaxsitestocurrent, maxsites" }));
            }

            string dbname = Params["dbname"].Value;

            SPWebApplication webApp = SPWebApplication.Lookup(new Uri(Params["webapp"].Value));
            SPContentDatabase db = null;
            foreach (SPContentDatabase tempDB in webApp.ContentDatabases)
            {
                if (tempDB.Name.ToLower() == dbname.ToLower())
                {
                    db = tempDB;
                    break;
                }
            }
            if (db == null)
                throw new Exception("Content database not found.");

            bool modified = false;
            if (Params["status"].UserTypedIn)
            {
                db.Status = (SPObjectStatus)Enum.Parse(typeof(SPObjectStatus), Params["status"].Value, true);
                modified = true;
            }

            if (Params["maxsites"].UserTypedIn)
            {
                db.MaximumSiteCount = int.Parse(Params["maxsites"].Value);
                modified = true;
            }
            else if (Params["setmaxsitestocurrent"].UserTypedIn)
            {
                if (db.CurrentSiteCount < db.WarningSiteCount)
                    db.WarningSiteCount = db.CurrentSiteCount - 1;
                db.MaximumSiteCount = db.CurrentSiteCount;
                
                modified = true;
            }

            if (Params["warningsitecount"].UserTypedIn)
            {
                db.WarningSiteCount = int.Parse(Params["warningsitecount"].Value);
                modified = true;
            }

            if (Params["searchserver"].UserTypedIn && !string.IsNullOrEmpty(Params["searchserver"].Value))
            {
                // If they specified a search server then we need to try and find a valid
                // matching search server using the server address property.
#if SP2010
                SPSearchService service = SPFarm.Local.Services.GetValue<SPSearchService>("SPSearch");
#elif SP2013
                SPSearchService service = SPFarm.Local.Services.GetValue<SPSearchService>("SPSearch4");
#endif
                SPServiceInstance searchServiceServer = null;
                foreach (SPServiceInstance tempsvc in service.Instances)
                {
                    if (!(tempsvc is SPSearchServiceInstance))
                        continue;

                    if (tempsvc.Status != SPObjectStatus.Online)
                        continue;

                    if (tempsvc.Server.Address.ToLowerInvariant() == Params["searchserver"].Value.ToLowerInvariant())
                    {
                        // We found a match so bug out of the loop.
                        searchServiceServer = tempsvc;
                        break;
                    }
                }
                if (searchServiceServer != null)
                {
                    db.SearchServiceInstance = searchServiceServer;
                    modified = true;
                }
                else
                    Logger.Write("Search server \"{0}\" not found.", EventLogEntryType.Warning, Params["searchserver"].Value);
            }
            else if (Params["searchserver"].UserTypedIn)
            {
                // The user specified the searchserver switch with no value which is what we use to indicate
                // clearing the value.
                db.SearchServiceInstance = null;
                modified = true;
            }

            if (modified)
                db.Update();

            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
