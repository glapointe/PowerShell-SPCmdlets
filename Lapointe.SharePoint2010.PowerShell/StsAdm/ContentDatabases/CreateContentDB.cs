using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Text;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Search.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.ContentDatabases
{
    public class CreateContentDB : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CreateContentDB"/> class.
        /// </summary>
        public CreateContentDB()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("dbname", "db", true, null, new SPNonEmptyValidator(), "Please specify the database name."));
            parameters.Add(new SPParam("dbserver", "server", false, null, new SPNonEmptyValidator(), "Please specify the database server."));
            parameters.Add(new SPParam("webapp", "web", true, null, new SPUrlValidator(), "Please specify the web application."));
            parameters.Add(new SPParam("maxsites", "max", false, "15000", new SPIntRangeValidator(0, Int32.MaxValue)));
            parameters.Add(new SPParam("warningsitecount", "warn", false, "9000", new SPIntRangeValidator(0, Int32.MaxValue)));
            parameters.Add(new SPParam("searchserver", "search", false, null, new SPNullOrNonEmptyValidator()));
            parameters.Add(new SPParam("dbuser", "dbuser", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("dbpwd", "dbpwd", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("status", "s", false, "online", new SPRegexValidator("^online$|^disabled$")));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nCreates a new content database.\r\n\r\nParameters:");
            sb.Append("\r\n\t-dbname <content database name>");
            sb.Append("\r\n\t-webapp <web application url>");
            sb.Append("\r\n\t[-dbserver <content database server>]");
            sb.Append("\r\n\t[-maxsites <maximum number of sites allowed in the db (default is 15000)>]");
            sb.Append("\r\n\t[-warningsitecount <number of sites before a warning event is generated (default is 9000)>]");
            sb.Append("\r\n\t[-searchserver <search server (leave empty to clear the search server)>]");
            sb.Append("\r\n\t[-dbuser <database username (if using SQL Authentication and not Windows Authentication)>]");
            sb.Append("\r\n\t[-dbpwd <database password (if using SQL Authentication and not Windows Authentication)>]");
            sb.Append("\r\n\t[-status <online | disabled>]");
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

            string dbserver = Params["dbserver"].Value;
            string dbname = Params["dbname"].Value;

            if (string.IsNullOrEmpty(dbserver))
            {
                dbserver = SPWebService.ContentService.DefaultDatabaseInstance.NormalizedDataSource;
            }
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
            if (db != null)
                throw new Exception("Content database already exists.");

            SPObjectStatus status = (SPObjectStatus)Enum.Parse(typeof(SPObjectStatus), Params["status"].Value, true);

            db = webApp.ContentDatabases.Add(dbserver, dbname, null, null, 
                int.Parse(Params["warningsitecount"].Value),
                int.Parse(Params["maxsites"].Value), (status == SPObjectStatus.Online?0:1));
            
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
                }
                else
                    Logger.Write("Search server \"{0}\" not found.", EventLogEntryType.Warning, Params["searchserver"].Value);
            }
            else if (Params["searchserver"].UserTypedIn)
            {
                // The user specified the searchserver switch with no value which is what we use to indicate
                // clearing the value.
                db.SearchServiceInstance = null;
            }

            db.Update();

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            if (Params["dbuser"].UserTypedIn || Params["dbpwd"].UserTypedIn)
            {
                Params["dbuser"].IsRequired = true;
                Params["dbpwd"].IsRequired = true;
            }
            base.Validate(keyValues);
        }
        #endregion
    }
}
