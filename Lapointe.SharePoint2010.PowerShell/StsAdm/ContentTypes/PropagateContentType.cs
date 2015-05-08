using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.ContentTypes
{
    /// <summary>
    /// This code was derived from Søren Nielsen's code that he provides on his blog: 
    /// http://soerennielsen.wordpress.com/2007/09/11/propagate-site-content-types-to-list-content-types/
    /// </summary>
    public class PropagateContentType : SPOperation
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="PropagateContentType"/> class.
        /// </summary>
        public PropagateContentType()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("contenttype", "ct", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("verbose", "v"));
            parameters.Add(new SPParam("updatefields", "uf"));
            parameters.Add(new SPParam("removefields", "rf"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nPropagates a site scoped content type to list scoped instances of that content type.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <site collection url>");
            sb.Append("\r\n\t-contenttype <content type name>");
            sb.Append("\r\n\t[-verbose]");
            sb.Append("\r\n\t[-updatefields]");
            sb.Append("\r\n\t[-removefields]");

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

            using (SPSite site = new SPSite(Params["url"].Value.TrimEnd('/')))
            {
                Common.ContentTypes.PropagateContentType.Execute(
                    site, Params["contenttype"].Value, Params["updatefields"].UserTypedIn, Params["removefields"].UserTypedIn);
            }

            return (int)ErrorCodes.NoError;
        }
    }
}
