using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.ContentTypes
{
    public class ExportContentTypes : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CopyContentTypes"/> class.
        /// </summary>
        public ExportContentTypes()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the source site collection."));
            parameters.Add(new SPParam("name", "n", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("group", "g", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("listname", "list", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("outputfile", "output", false, null, new SPDirectoryExistsAndValidFileNameValidator()));
            parameters.Add(new SPParam("includelistbindings", "ilb"));
            parameters.Add(new SPParam("includefielddefinitions", "ifd"));
            parameters.Add(new SPParam("excludeparentfields", "epf"));
            parameters.Add(new SPParam("removeencodedspaces", "res"));
            parameters.Add(new SPParam("featuresafe", "safe"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nExports Content Types to an XML file.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <url containing the content types>");
            sb.Append("\r\n\t-outputfile <file to output results to>");
            sb.Append("\r\n\t[-name <name of an individual content type to export>]");
            sb.Append("\r\n\t[-group <content type group name to filter results by>]");
            sb.Append("\r\n\t[-listname <name of a list to export content types from>]");
            sb.Append("\r\n\t[-includelistbindings]");
            sb.Append("\r\n\t[-includefielddefinitions]");
            sb.Append("\r\n\t[-excludeparentfields]");
            sb.Append("\r\n\t[-removeencodedspaces (removes '_x0020_' in field names)]");
            sb.Append("\r\n\t[-featuresafe]");
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

            string url = Params["url"].Value.TrimEnd('/');
            string contentTypeName = null;
            if (Params["name"].UserTypedIn)
                contentTypeName = Params["name"].Value;
            string contentTypeGroup = null;
            if (Params["group"].UserTypedIn)
                contentTypeGroup = Params["group"].Value.ToLowerInvariant();
            if (Params["group"].UserTypedIn && Params["listname"].UserTypedIn)
                throw new SPSyntaxException("The parameters group and listname are incompatible");
            bool excludeParentFields = Params["excludeparentfields"].UserTypedIn;
            bool includeFieldDefinitions = Params["includefielddefinitions"].UserTypedIn;
            bool includeListBindings = Params["includelistbindings"].UserTypedIn;
            bool removeEncodedSpaces = Params["removeencodedspaces"].UserTypedIn;
            bool featureSafe = Params["featuresafe"].UserTypedIn;
            string outputFile = Params["outputfile"].Value;
            string listName = null;
            if (Params["listname"].UserTypedIn)
                listName = Params["listname"].Value;

            Common.ContentTypes.ExportContentTypes.Export(url, contentTypeGroup, contentTypeName, excludeParentFields, includeFieldDefinitions, includeListBindings, listName, removeEncodedSpaces, featureSafe, outputFile);
            
            return (int)ErrorCodes.NoError;
        }
        #endregion


    }
}
