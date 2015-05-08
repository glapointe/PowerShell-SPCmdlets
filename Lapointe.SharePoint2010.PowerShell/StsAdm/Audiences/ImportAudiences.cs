using System;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
#if MOSS
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
#endif
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Audiences
{
    public class ImportAudiences : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ImportAudiences"/> class.
        /// </summary>
        public ImportAudiences()
        {
            SPParamCollection parameters = new SPParamCollection();
            StringBuilder sb = new StringBuilder();

#if MOSS
            parameters.Add(new SPParam("serviceappname", "app", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("sitesubscriptionid", "subid", false, Guid.Empty.ToString(), new SPGuidValidator()));
            parameters.Add(new SPParam("contextsite", "site", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("deleteexisting", "delete"));
            parameters.Add(new SPParam("inputfile", "input", false, null, new SPFileExistsValidator()));
            parameters.Add(new SPParam("compile", "c"));
            parameters.Add(new SPParam("mapfile", "map", false, null, new SPDirectoryExistsAndValidFileNameValidator()));

            sb.Append("\r\n\r\nImports all audiences given the provided input file.\r\n\r\nParameters:");
            sb.Append("\r\n\t-inputfile <file to input results from>");
            sb.Append("\r\n\t[-deleteexisting <delete existing audiences>]");
            sb.Append("\r\n\t[-serviceappname <user profile service application name>]");
            sb.Append("\r\n\t[-sitesubscriptionid <GUID>]");
            sb.Append("\r\n\t[-contextsite <URL of the site to use for service context>]");
            sb.Append("\r\n\t[-compile]");
            sb.Append("\r\n\t[-mapfile <generate a map file to use for search and replace of Audience IDs>]");
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

            string inputFile = Params["inputfile"].Value;
            bool deleteExisting = Params["deleteexisting"].UserTypedIn;
            bool compile = Params["compile"].UserTypedIn;
            string mapFile = default(string);
            if (Params["mapfile"].UserTypedIn)
                mapFile = Params["mapfile"].Value;

            string xml = File.ReadAllText(inputFile);
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

            Common.Audiences.ImportAudiences.Import(xml, context, deleteExisting, compile, mapFile);

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            SPBinaryParameterValidator.Validate("serviceappname", Params["serviceappname"].Value, "contextsite", Params["contextsite"].Value);

            base.Validate(keyValues);
        }
    }
}
