using System.Collections.Generic;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.StsAdm.Lists;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Pages
{
    public class EnumUnGhostedFiles : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EnumUnGhostedFiles"/> class.
        /// </summary>
        public EnumUnGhostedFiles()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site url."));
            parameters.Add(new SPParam("recursesubwebs", "recurse", false, null, null));
            Init(parameters, "\r\n\r\nReturns a list of all unghosted (customized) files for a web.\r\n\r\nParameters:\r\n\t-url <web site url>\r\n\t[-recursesubwebs]");
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

            string url = Params["url"].Value.TrimEnd('/');
            bool recurse = Params["recursesubwebs"].UserTypedIn;
            List<object> unghostedFiles = new List<object>();

            using (SPSite site = new SPSite(url))
            {
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {
                    if (recurse)
                    {
                        Common.Pages.EnumUnGhostedFiles.RecurseSubWebs(web, ref unghostedFiles, true);
                    }
                    else
                        Common.Pages.EnumUnGhostedFiles.CheckFoldersForUnghostedFiles(web.RootFolder, ref unghostedFiles, true);
                }
            }

            if (unghostedFiles.Count == 0)
            {
                output += "There are no unghosted (customized) files on the current web.\r\n";
            }
            else
            {
                output += "The following files are unghosted:";

                foreach (string fileName in unghostedFiles)
                {
                    output += "\r\n\t" + fileName;
                }
            }

            return (int)ErrorCodes.NoError;
        }

        #endregion

      

    }
}
