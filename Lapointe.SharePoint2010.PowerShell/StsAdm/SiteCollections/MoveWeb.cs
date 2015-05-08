using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Deployment;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    public class MoveWeb : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MoveWeb"/> class.
        /// </summary>
        public MoveWeb()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the URL of the web to move."));
            parameters.Add(new SPParam("parenturl", "parent", true, null, new SPUrlValidator(), "Please specify the parent url."));
            parameters.Add(new SPParam("haltonwarning", "warning", false, null, null));
            parameters.Add(new SPParam("haltonfatalerror", "error", false, null, null));
            parameters.Add(new SPParam("includeusersecurity", "security", false, null, null));
            parameters.Add(new SPParam("retainobjectidentity", "retainid", false, null, null));
            parameters.Add(new SPParam("suppressafterevents", "sae"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nMoves a web.\r\n\r\nParameters:\r\n\t-url <url of web to move>");
            sb.Append("\r\n\t-parenturl <url of parent web>");
            sb.Append("\r\n\t[-haltonwarning (only considered if moving to a new site collection)]");
            sb.Append("\r\n\t[-haltonfatalerror (only considered if moving to a new site collection)]");
            sb.Append("\r\n\t[-includeusersecurity (only considered if moving to a new site collection)]");
            sb.Append("\r\n\t[-retainobjectidentity (only considered if moving to a new site collection)]");
            sb.Append("\r\n\t[-suppressafterevents (disable the firing of \"After\" events when creating or modifying list items)]");

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
            string parentUrl = Params["parenturl"].Value.TrimEnd('/');
            bool retainObjectIdentity = Params["retainobjectidentity"].UserTypedIn;
            bool suppressAfterEvents = Params["suppressafterevents"].UserTypedIn;

            using (SPSite sourceSite = new SPSite(url))
            using (SPWeb sourceWeb = sourceSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            using (SPSite parentSite = new SPSite(parentUrl))
            using (SPWeb parentWeb = parentSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(parentUrl)])
            using (SPWeb sourceParentWeb = sourceWeb.ParentWeb)
            {
                if (sourceWeb.ID == parentWeb.ID)
                {
                    throw new Exception("Source web and parent web cannot be the same.");
                }
                if (sourceParentWeb != null && sourceParentWeb.ID == parentWeb.ID)
                {
                    throw new Exception(
                        "Parent web specified matches the source web's current parent - move is not necessary.");
                }
                if (sourceWeb.IsRootWeb && sourceSite.ID == parentSite.ID)
                {
                    throw new Exception("Cannot move root web within the same site collection.");
                }
                if (sourceWeb.IsRootWeb && retainObjectIdentity)
                {
                    // If we allow retainobjectidentity when moving a root web the import will attempt to import over
                    // the parent web's site collection which would be really bad.
                    throw new Exception("Cannot move a root web when \"-retainobjectidentity\" is used.");
                }

                if (sourceSite.ID == parentSite.ID)
                {
                    // This ones the easy one - just need to set the property and update the web.
                    sourceWeb.ServerRelativeUrl = Utilities.ConcatServerRelativeUrls(parentWeb.ServerRelativeUrl, sourceWeb.Name);
                    sourceWeb.Update();
                }
                else
                {
                    // Now for the hard one - we need to move to another site collection which requires using the export/import commands.
                    bool haltOnWarning = Params["haltonwarning"].UserTypedIn;
                    bool haltOnFatalError = Params["haltonfatalerror"].UserTypedIn;
                    bool includeUserSecurity = Params["includeusersecurity"].UserTypedIn;


                    Common.SiteCollections.MoveWeb.MoveWebOutsideSite(sourceWeb, parentWeb, null, retainObjectIdentity, haltOnWarning, haltOnFatalError, includeUserSecurity, suppressAfterEvents, null);
                }
            }

            return (int)ErrorCodes.NoError;
        }


        #endregion

        

    }
}
