using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Threading;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Deployment;
using Microsoft.SharePoint.PowerShell;
#if MOSS
using Microsoft.SharePoint.Publishing;
#endif
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    public class ConvertSubSiteToSiteCollection : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConvertSubSiteToSiteCollection"/> class.
        /// </summary>
        public ConvertSubSiteToSiteCollection()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("sourceurl", "source", true, null, new SPUrlValidator(), "Please specify the source url."));
            parameters.Add(new SPParam("targeturl", "target", true, null, new SPUrlValidator(), "Please specify the target url."));
            parameters.Add(new SPParam("owneremail", "oe", true, null, new SPRegexValidator(@"^[^ \r\t\n\f@]+@[^ \r\t\n\f@]+$"), "Please specify the owner email."));
            parameters.Add(new SPParam("createmanagedpath", "createpath", false, null, null));
            parameters.Add(new SPParam("haltonwarning", "warning", false, null, null));
            parameters.Add(new SPParam("haltonfatalerror", "error", false, null, null));
            parameters.Add(new SPParam("includeusersecurity", "security", false, null, null));
            parameters.Add(new SPParam("exportedfile", "file", false, null, new SPNonEmptyValidator(), "Please specify the exported filename."));
            parameters.Add(new SPParam("nofilecompression", "nofilecompression", false, null, null));
            parameters.Add(new SPParam("ownerlogin", "ol", false, null, new SPNonEmptyValidator(), "Please specify the owner login."));
            parameters.Add(new SPParam("ownername", "on", false, null, new SPNonEmptyValidator(), "Please specify the owner name."));
            parameters.Add(new SPParam("secondarylogin", "sl", false, null, new SPNonEmptyValidator(), "Please specify the secondary owner login."));
            parameters.Add(new SPParam("secondaryname", "sn", false, null, new SPNonEmptyValidator(), "Please specify the secondary owner name."));
            parameters.Add(new SPParam("secondaryemail", "se", false, null, new SPRegexValidator(@"^[^ \r\t\n\f@]+@[^ \r\t\n\f@]+$"), "Please specify the secondary owner email."));
            parameters.Add(new SPParam("lcid", "lcid", false, null, new SPRegexValidator("^[0-9]+$"), "Please specify the language identifier or LCID (example: 1033 for US English)."));
            parameters.Add(new SPParam("title", "t", false, null, new SPNonEmptyValidator(), "Please specify the title for the new site."));
            parameters.Add(new SPParam("description", "desc", false, null, new SPNonEmptyValidator(), "Please specify the description for the new site."));
            parameters.Add(new SPParam("hostheaderwebapplicationurl", "hhurl", false, null, new SPUrlValidator(), "Please specify the host header web application url."));
            parameters.Add(new SPParam("quota", "quota", false, null, new SPNonEmptyValidator(), "Please specify the quota template to assign the new site."));
            parameters.Add(new SPParam("deletesource", "deletesource", false, null, null));
            parameters.Add(new SPParam("createsiteindb", "db", false, null, null));
            parameters.Add(new SPParam("verbose", "v"));

            SPWebService contentService = SPWebService.ContentService;
            string normalizedDataSource = contentService.DefaultDatabaseInstance.NormalizedDataSource;
            string defaultDatabaseUsername = contentService.DefaultDatabaseUsername;
            string defaultDatabasePassword = contentService.DefaultDatabasePassword;
            if (string.IsNullOrEmpty(normalizedDataSource))
            {
                normalizedDataSource = Environment.MachineName;
            }
            if (defaultDatabaseUsername == null)
            {
                defaultDatabaseUsername = "";
            }
            if (defaultDatabasePassword == null)
            {
                defaultDatabasePassword = "";
            }
            parameters.Add(new SPParam("databaseserver", "ds", false, normalizedDataSource, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("databasename", "dn", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("databaseuser", "du", false, defaultDatabaseUsername, null));
            parameters.Add(new SPParam("databasepassword", "dp", false, defaultDatabasePassword, null));
            parameters.Add(new SPParam("suppressafterevents", "sae"));


            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nConverts a sub-site to a top level site collection via a managed path.\r\n\r\nParameters:\r\n\t");
            sb.Append("\r\n\t-sourceurl <source location of the existing sub-site or model site collection>");
            sb.Append("\r\n\t-targeturl <target location for the new site collection>");
            sb.Append("\r\n\t-owneremail <someone@example.com>");
            // Removed 6/25/2009
            //sb.Append("\r\n\t{-sitetemplate <site template if exportedfile is specified - must match the template of the exported site> |");
            //sb.Append("\r\n\t -nositetemplate (will automatically apply the sources site template) |");
            //sb.Append("\r\n\t -donotcreatesite (specify if the site to import to already exists)}");
            sb.Append("\r\n\t[-createmanagedpath]");
            sb.Append("\r\n\t[-haltonwarning]");
            sb.Append("\r\n\t[-haltonfatalerror]");
            sb.Append("\r\n\t[-includeusersecurity]");
            sb.Append("\r\n\t[-suppressafterevents (disable the firing of \"After\" events when creating or modifying list items)]");
            sb.Append("\r\n\t[-exportedfile <filename of exported site if previously exported>]");
            sb.Append("\r\n\t[-nofilecompression]");
            sb.Append("\r\n\t[-ownerlogin <DOMAIN\\name>]");
            sb.Append("\r\n\t[-ownername <display name>]");
            sb.Append("\r\n\t[-secondaryemail <someone@example.com>]");
            sb.Append("\r\n\t[-secondarylogin <DOMAIN\\name>]");
            sb.Append("\r\n\t[-secondaryname <display name>]");
            sb.Append("\r\n\t[-lcid <language>]");
            sb.Append("\r\n\t[-title <site title>]");
            sb.Append("\r\n\t[-description <site description>]");
            sb.Append("\r\n\t[-hostheaderwebapplicationurl <web application url>]");
            sb.Append("\r\n\t[-quota <quota template>]");
            sb.Append("\r\n\t[-deletesource]");
            sb.Append("\r\n\t[-createsiteindb]");
            sb.Append("\r\n\t[-databaseuser <database username>]");
            sb.Append("\r\n\t[-databasepassword <database password>]");
            sb.Append("\r\n\t[-databaseserver <database server name>]");
            sb.Append("\r\n\t[-databasename <database name>]");
            sb.Append("\r\n\t[-verbose]");

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
           

            Logger.Verbose = Params["verbose"].UserTypedIn;
            string sourceurl = Params["sourceurl"].Value.TrimEnd('/');
            string targeturl = Params["targeturl"].Value.TrimEnd('/');
            bool suppressAfterEvents = Params["suppressafterevents"].UserTypedIn;
            bool noFileCompression = Params["nofilecompression"].UserTypedIn;
            string exportedFile = Params["exportedfile"].Value;
            bool createSiteInDB = Params["createsiteindb"].UserTypedIn;
            string databaseName = Params["databasename"].Value;
            bool createManagedPath = Params["createmanagedpath"].UserTypedIn;
            bool haltOnWarning = Params["haltonwarning"].UserTypedIn;
            bool haltOnFatalError = Params["haltonfatalerror"].UserTypedIn;
            bool deleteSource = Params["deletesource"].UserTypedIn;
            string title = Params["title"].Value;
            string description = Params["description"].Value;
            uint nLCID = uint.Parse(Params["lcid"].Value);
            string ownerName = Params["ownername"].Value;
            string ownerEmail = Params["owneremail"].Value;
            string ownerLogin = Params["ownerlogin"].Value;
            string secondaryContactName = Params["secondaryname"].Value;
            string secondaryContactLogin = Params["secondarylogin"].Value;
            string secondaryContactEmail = Params["secondaryemail"].Value;
            string quota = Params["quota"].Value;
            bool useHostHeaderAsSiteName = Params["hostheaderwebapplicationurl"].UserTypedIn;

            SPSite site = Common.SiteCollections.ConvertSubSiteToSiteCollection.ConvertWebToSite(sourceurl, targeturl, null, suppressAfterEvents,
                noFileCompression, exportedFile, createSiteInDB, databaseName, createManagedPath, haltOnWarning, 
                haltOnFatalError, deleteSource, title, description, nLCID, ownerName, ownerEmail, ownerLogin, secondaryContactName, 
                secondaryContactLogin, secondaryContactEmail, quota, useHostHeaderAsSiteName);
            if (site != null)
                site.Dispose();

            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
