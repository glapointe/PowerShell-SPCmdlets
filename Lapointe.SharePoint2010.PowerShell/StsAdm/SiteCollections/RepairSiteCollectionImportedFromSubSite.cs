using System;
using System.Collections;
using System.Collections.Generic;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using System.Text;
using System.Collections.Specialized;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    /// <summary>
    /// Repairs a site collection that has been imported from an exported sub-site.  
    /// Note that the sourceurl can be the actual source site or any site collection 
    /// that can be used as a model for the target.
    /// </summary>
    public class RepairSiteCollectionImportedFromSubSite : SPOperation
    {
        public RepairSiteCollectionImportedFromSubSite()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("sourceurl", "source", true, null, new SPUrlValidator(), "Please specify the source sub-site or model site collection."));
            parameters.Add(new SPParam("targeturl", "target", true, null, new SPUrlValidator(), "Please specify the target location of the new site collection to repair."));
            
            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nRepairs a site collection that has been imported from an exported sub-site.  Note that the sourceurl can be the actual source site or any site collection that can be used as a model for the target.\r\n\r\nParameters:");
            sb.Append("\r\n\t-sourceurl <source location of the existing sub-site or model site collection>");
            sb.Append("\r\n\t-targeturl <target location for the new site collection>");
            
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

            string sourceurl = Params["sourceurl"].Value.TrimEnd('/');
            string targeturl = Params["targeturl"].Value.TrimEnd('/');
            Logger.Verbose = true;

            Common.SiteCollections.RepairSiteCollectionImportedFromSubSite.RepairSite(sourceurl, targeturl);

            return (int)ErrorCodes.NoError;
        }

        #endregion

    }
}
