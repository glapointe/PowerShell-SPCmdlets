using System;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Security
{
    public class DeleteAllUsers : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteAllUsers"/> class.
        /// </summary>
        public DeleteAllUsers()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site collection url."));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nDeletes all site collection users.  Will not delete site administrators.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <site collection url>");
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

            string url = Params["url"].Value;
            if (url != null)
                url = url.TrimEnd('/');

            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb())
            {
                int offsetIndex = 0;
                int count = 0;
                int err = 0;
                Logger.Write("Starting user deletion...");
                while (web.SiteUsers.Count > offsetIndex)
                {

                    if (web.SiteUsers[offsetIndex].IsSiteAdmin || web.SiteUsers[offsetIndex].ID == web.CurrentUser.ID)
                    {
                        offsetIndex++;
                        continue;
                    }
                    Logger.Write("Progress: Deleting {0}", web.SiteUsers[offsetIndex].LoginName);
                    try
                    {
                        web.SiteUsers.Remove(offsetIndex);
                        count++;
                    }
                    catch (Exception ex)
                    {
                        err++;
                        offsetIndex++;
                        Logger.Write("ERROR: Unable to delete user.\r\n{0}\r\n{1}", ex.Message, ex.StackTrace);
                    }
                }
                Logger.Write("Finished user deletion.  {0} Users deleted, {1} errors.", count.ToString(), err.ToString());
            }


            return (int)ErrorCodes.NoError;
        }

    }
}
