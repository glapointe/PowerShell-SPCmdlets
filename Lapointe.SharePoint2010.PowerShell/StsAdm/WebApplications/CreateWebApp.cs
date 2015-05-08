using System;
using System.IO;
using System.Reflection;
using System.Security;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebApplications
{
    public class CreateWebApp : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CreateWebApp"/> class.
        /// </summary>
        public CreateWebApp()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator()));
            parameters.Add(new SPParam("directory", "dir", false, null, new SPNonEmptyValidator(),
                                       "Please specify the virtual directory for the web application."));
            parameters.Add(new SPParam("port", "p", false, "80", new SPIntRangeValidator(0, int.MaxValue),
                                       "Please specify the port for the web application."));
            parameters.Add(new SPParam("timezone", "tz", false, null, new SPIntRangeValidator(0, ushort.MaxValue)));
            parameters.Add(new SPParam("description", "desc", false, null, new SPNullOrNonEmptyValidator()));
            parameters.Add(new SPParam("sethostheader", "sethh", false, null, new SPNullOrNonEmptyValidator()));
            parameters.Add(new SPParam("exclusivelyusentlm", "ntlm"));
            parameters.Add(new SPParam("allowanonymous", "anon"));
            parameters.Add(new SPParam("ssl", "ssl"));

            string normalizedDataSource;
            string defaultDatabaseUsername;
            string defaultDatabasePassword;
            SPFarm local = SPFarm.Local;
            if (local != null)
            {
                SPWebService service = local.Services.GetValue<SPWebService>();
                normalizedDataSource = service.DefaultDatabaseInstance.NormalizedDataSource;
                defaultDatabaseUsername = service.DefaultDatabaseUsername;
                defaultDatabasePassword = service.DefaultDatabasePassword;
                if ((normalizedDataSource == null) || (normalizedDataSource.Length == 0))
                {
                    normalizedDataSource = Environment.MachineName;
                }
            }
            else
            {
                Console.WriteLine(SPResource.GetString("NoFarmObject", new object[0]));
                return;
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
            parameters.Add(new SPParam("databasename", "dn", false, null, new SPNullOrNonEmptyValidator()));
            parameters.Add(new SPParam("databaseuser", "du", false, defaultDatabaseUsername, null));
            parameters.Add(new SPParam("databasepassword", "dp", false, defaultDatabasePassword, null));

            parameters.Add(new SPParam("apidname", "apid", false, "DefaultAppPool", new SPNonEmptyValidator()));
            parameters.Add(new SPParam("apidtype", "apidtype", false, "NetworkService",
                                       new SPRegexValidator("^configurableid$|^networkservice$")));
            parameters.Add(new SPParam("apidlogin", "apu", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("apidpwd", "app", true, null, new SPNonEmptyValidator()));

            parameters.Add(new SPParam("donotcreatesite", "nosite"));
            parameters.Add(new SPParam("ownerlogin", "ol", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("ownername", "on", false, null, null));
            parameters.Add(new SPParam("owneremail", "oe", true, null,
                                       new SPRegexValidator(@"^[^ \r\t\n\f@]+@[^ \r\t\n\f@]+$")));
            parameters.Add(new SPParam("sitetemplate", "st", false, null, new SPNullOrNonEmptyValidator()));
            parameters.Add(new SPParam("lcid", "lcid", false, "0", new SPRegexValidator("^[0-9]+$")));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nCreates a web application.\r\n\r\nParameters:\r\n");
            sb.Append("\t-url <url>\r\n");
            sb.Append("\t[-directory <virtual directory path>]\r\n");
            sb.Append("\t[-port <web application port>]\r\n");
            sb.Append("\t[-ownerlogin <domain\\name>]\r\n");
            sb.Append("\t[-owneremail <someone@example.com>]\r\n");
            sb.Append("\t[-exclusivelyusentlm]\r\n");
            sb.Append("\t[-ownername <display name>]\r\n");
            sb.Append("\t[-databaseuser <database user>]\r\n");
            sb.Append("\t[-databaseserver <database server>]\r\n");
            sb.Append("\t[-databasename <database name>]\r\n");
            sb.Append("\t[-databasepassword <database user password>]\r\n");
            sb.Append("\t[-lcid <language>]\r\n");
            sb.Append("\t[-sitetemplate <site template>]\r\n");
            sb.Append("\t[-donotcreatesite]\r\n");
            sb.Append("\t[-description <iis web site name>]\r\n");
            sb.Append("\t[-sethostheader <host header name>]\r\n");
            sb.Append("\t[-apidname <app pool name>]\r\n");
            sb.Append("\t[-apidtype <configurableid/NetworkService>]\r\n");
            sb.Append("\t[-apidlogin <DOMAIN\\name>]\r\n");
            sb.Append("\t[-apidpwd <app pool password>]\r\n");
            sb.Append("\t[-allowanonymous]\r\n");
            sb.Append("\t[-ssl]\r\n");
            sb.Append("\t[-timezone <time zone ID>]\r\n");

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
        public override int Execute(string command, System.Collections.Specialized.StringDictionary keyValues,
                                    out string output)
        {
            output = string.Empty;
            
            Uri uri = new Uri(Params["url"].Value);

            SPWebApplicationBuilder builder = GetWebAppBuilder(uri);

            SPWebApplication app = builder.Create();

            SPAdministrationWebApplication local = SPAdministrationWebApplication.Local;

            // Set the TimeZone of the Application
            if (Params["timezone"].UserTypedIn)
                app.DefaultTimeZone = ushort.Parse(Params["timezone"].Value);

            app.Update();
            app.ProvisionGlobally();

            // Execute pending timer jobs before moving on.
            Common.TimerJobs.ExecAdmSvcJobs.Execute(false, true);
            // Recreate the web application object to avoid update conflicts.
            app = SPWebApplication.Lookup(uri);

            // Upload the newly created WebApplication to the List 'Web Application List' in Central Administration:
            SPWebService.AdministrationService.WebApplications.Add(app);

            if (!Params["donotcreatesite"].UserTypedIn)
            {
                uint nLCID = uint.Parse(Params["lcid"].Value);
                string webTemplate = Params["sitetemplate"].Value;
                string ownerLogin = Params["ownerlogin"].Value;
                ownerLogin = Utilities.TryGetNT4StyleAccountName(ownerLogin, app);
                string ownerName = Params["ownername"].Value;
                string ownerEmail = Params["owneremail"].Value;

                app.Sites.Add(uri.AbsolutePath, null, null, nLCID, webTemplate, ownerLogin, ownerName, ownerEmail, null, null, null);
            }


            Console.WriteLine(SPResource.GetString("PendingRestartInExtendWebFarm", new object[0]));
            Console.WriteLine();

            if (!Params["donotcreatesite"].UserTypedIn)
                Console.WriteLine(SPResource.GetString("AccessSiteAt", new object[] {uri.ToString()}));
            Console.WriteLine();

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(System.Collections.Specialized.StringDictionary keyValues)
        {
            bool isUserAccount;
            if (Params["donotcreatesite"].UserTypedIn)
            {
                Params["ownerlogin"].Enabled = false;
                Params["ownername"].Enabled = false;
                Params["owneremail"].Enabled = false;
                Params["sitetemplate"].Enabled = false;
                Params["lcid"].Enabled = false;
            }
            else
            {
                if (Params["ownerlogin"].UserTypedIn)
                {
                    string ownerLogin = Utilities.TryGetNT4StyleAccountName(Params["ownerlogin"].Value, null);
                    if (!SPUtility.IsLoginValid(null, ownerLogin, out isUserAccount))
                        throw new ArgumentException(
                            SPResource.GetString("InvalidLoginAccount", new object[] {ownerLogin}));
                    if (!isUserAccount)
                        throw new ArgumentException(SPResource.GetString("OwnerNotUserAccount", new object[0]));
                }
            }
            if (Params["apidtype"].Value.ToLowerInvariant() == "networkservice")
            {
                Params["apidlogin"].Enabled = false;
                Params["apidpwd"].Enabled = false;
            }
            else
            {
                if (Params["apidlogin"].UserTypedIn)
                {
                    string apidlogin = Utilities.TryGetNT4StyleAccountName(Params["apidlogin"].Value, null);
                    if (!SPUtility.IsLoginValid(null, apidlogin, out isUserAccount))
                        throw new ArgumentException(
                            SPResource.GetString("InvalidLoginAccount", new object[] {apidlogin}));
                }
            }
            base.Validate(keyValues);
        }

        #endregion

        /// <summary>
        /// Gets the web app builder.
        /// </summary>
        /// <param name="uri">The URI.</param>
        /// <returns></returns>
        private SPWebApplicationBuilder GetWebAppBuilder(Uri uri)
        {
            SPWebApplicationBuilder builder = new SPWebApplicationBuilder(SPFarm.Local);

            //Set the Port and the RootDirectory where you want to install the Application, e.g:
            builder.Port = int.Parse(Params["port"].Value);
            if (Params["directory"].UserTypedIn)
                builder.RootDirectory = new DirectoryInfo(Params["directory"].Value);

            // Set the ServerComment for the Application which will be the Name of the Application in the SharePoint-List And IIS. If you do not set this Property, the Name of the Application will be 'SharePoint - <Default given Portnumber from System>' 
            if (Params["description"].UserTypedIn && !string.IsNullOrEmpty(Params["description"].Value))
                builder.ServerComment = Params["description"].Value;

            // Create the content database for this Application
            builder.CreateNewDatabase = true;
            if (Params["databasename"].UserTypedIn)
                builder.DatabaseName = Params["databasename"].Value;
            if (Params["databaseserver"].UserTypedIn)
                builder.DatabaseServer = Params["databaseserver"].Value;
            if (Params["databaseuser"].UserTypedIn)
                builder.DatabaseUsername = Params["databaseuser"].Value;
            if (Params["databasepassword"].UserTypedIn)
                builder.DatabasePassword = Params["databasepassword"].Value;

            // Host Header settings
            if (Params["sethostheader"].UserTypedIn)
            {
                if (string.IsNullOrEmpty(Params["sethostheader"].Value))
                    builder.HostHeader = uri.Host;
                else
                    builder.HostHeader = Params["sethostheader"].Value;
            }
            builder.DefaultZoneUri = uri;


            // App pool settings
            builder.ApplicationPoolId = Params["apidname"].Value;
            if (Params["apidtype"].Value.ToLowerInvariant() == "networkservice")
                builder.IdentityType = IdentityType.NetworkService;
            else
            {
                builder.IdentityType = IdentityType.SpecificUser;
                builder.ApplicationPoolUsername = Params["apidlogin"].Value;
                builder.ApplicationPoolPassword = Utilities.CreateSecureString(Params["apidpwd"].Value);
            }

            // Some additional Settings
            builder.UseNTLMExclusively = Params["exclusivelyusentlm"].UserTypedIn;
            builder.AllowAnonymousAccess = Params["allowanonymous"].UserTypedIn;
            builder.UseSecureSocketsLayer = Params["ssl"].UserTypedIn;
            return builder;
        }

    }
}