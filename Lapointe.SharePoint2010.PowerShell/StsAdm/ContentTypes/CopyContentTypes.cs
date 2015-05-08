using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;
using Microsoft.SharePoint.Workflow;
using System.Runtime.InteropServices;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;

namespace Lapointe.SharePoint.PowerShell.StsAdm.ContentTypes
{
    public class CopyContentTypes : SPOperation
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="CopyContentTypes"/> class.
        /// </summary>
        public CopyContentTypes()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("sourceurl", "sourceurl", true, null, new SPUrlValidator(), "Please specify the source site collection."));
            parameters.Add(new SPParam("targeturl", "targeturl", true, null, new SPUrlValidator(), "Please specify the target site collection"));
            parameters.Add(new SPParam("sourcename", "sourcename", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("noworkflows", "noworkflows", false, null, null));
            parameters.Add(new SPParam("nocolumns", "nocolumns", false, null, null));
            parameters.Add(new SPParam("nodocconversions", "nodocconversions", false, null, null));
            parameters.Add(new SPParam("nodocinfopanel", "nodocinfopanel", false, null, null));
            parameters.Add(new SPParam("nopolicies", "nopolicies", false, null, null));
            parameters.Add(new SPParam("nodoctemplate", "nodoctemplate", false, null, null));
            parameters.Add(new SPParam("verbose", "v"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nCopies all Content Types from one gallery to another.\r\n\r\nParameters:");
            sb.Append("\r\n\t-sourceurl <site collection url containing the source content types>");
            sb.Append("\r\n\t-targeturl <site collection url where the content types will be copied to>");
            sb.Append(
                "\r\n\t[-sourcename <name of an individual content type to copy - if ommitted all content types are copied if they don't already exist>]");
            sb.Append("\r\n\t[-noworkflows]");
            sb.Append("\r\n\t[-nocolumns]");
            sb.Append("\r\n\t[-nodocconversions]");
            sb.Append("\r\n\t[-nodocinfopanel]");
            sb.Append("\r\n\t[-nopolicies]");
            sb.Append("\r\n\t[-nodoctemplate]");
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
        public override int Execute(string command, System.Collections.Specialized.StringDictionary keyValues, out string output)
        {
            output = string.Empty;

            string sourceUrl = Params["sourceurl"].Value.TrimEnd('/');
            string targetUrl = Params["targeturl"].Value.TrimEnd('/');
            string sourceContentTypeName = null;
            if (Params["sourcename"].UserTypedIn)
                sourceContentTypeName = Params["sourcename"].Value;
            bool verbose = Params["verbose"].UserTypedIn;


            bool copyWorkflows = !Params["noworkflows"].UserTypedIn;
            bool copyColumns = !Params["nocolumns"].UserTypedIn;
            bool copyDocConversions = !Params["nodocconversions"].UserTypedIn;
            bool copyDocInfoPanel = !Params["nodocinfopanel"].UserTypedIn;
            bool copyPolicies = !Params["nopolicies"].UserTypedIn;
            bool copyDocTemplate = !Params["nodoctemplate"].UserTypedIn;

            Logger.Verbose = verbose;
            Logger.Write("Start Time: {0}", DateTime.Now.ToString());

            Common.ContentTypes.CopyContentTypes ctCopier = new Common.ContentTypes.CopyContentTypes(
                copyWorkflows, copyColumns, copyDocConversions, copyDocInfoPanel, copyPolicies, copyDocTemplate);

            ctCopier.Copy(sourceUrl, targetUrl, sourceContentTypeName);

            Logger.Write("Finish Time: {0}", DateTime.Now.ToString());

            return (int)ErrorCodes.NoError;
        }
        #endregion

    }
}
