using System;
using System.Collections.Specialized;
using System.IO;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.Deployment;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class DeleteList : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteList"/> class.
        /// </summary>
        public DeleteList()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the list view URL."));
            parameters.Add(new SPParam("listname", "list", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("force", "f"));
            parameters.Add(new SPParam("backupdir", "backup", false, null, new SPDirectoryExistsValidator()));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nDeletes a list.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <list view URL or web url>");
            sb.Append("\r\n\t[-listname <list name if url is a web url and not a list view url>]");
            sb.Append("\r\n\t[-force]");
            sb.Append("\r\n\t[-backupdir <directory to backup the list to>]");
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
            bool force = Params["force"].UserTypedIn;
            string backupDir = Params["backupdir"].Value;

            SPList list = null;
            if (Utilities.EnsureAspx(url, false, false) && !Params["listname"].UserTypedIn)
                list = Utilities.GetListFromViewUrl(url);
            else if (Params["listname"].UserTypedIn)
            {
                using (SPSite site = new SPSite(url))
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        list = web.Lists[Params["listname"].Value];
                    }
                    catch (ArgumentException)
                    {
                        throw new SPException("List not found.");
                    }
                }
            }

            if (list == null)
                throw new SPException("List not found.");


            Common.Lists.DeleteList.Delete(force, backupDir, list);

            return (int)ErrorCodes.NoError;
        }


        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            if (Params["url"].UserTypedIn && !Utilities.EnsureAspx(Params["url"].Value, false, false))
                Params["listname"].IsRequired = true;

            base.Validate(keyValues);
        }

        #endregion


    }
}
