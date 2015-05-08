using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.Win32;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebApplications
{
    public class SetBackConnectionHostNames : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SetBackConnectionHostNames"/> class.
        /// </summary>
        public SetBackConnectionHostNames()
        {
            
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("updatefarm", "uf"));
            parameters.Add(new SPParam("username", "user", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("password", "pwd", false, null, new SPNullOrNonEmptyValidator()));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nSets the BackConnectionHostNames registry key with the URLs associated with each web application.\r\n\r\nParameters:");
            sb.Append("\r\n\t[-updatefarm (update all servers in the farm)]");
            sb.Append("\r\n\t[-username <DOMAIN\\user (must have rights to update the registry on each server)>]");
            sb.Append("\r\n\t[-password <password>]");

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

            if (!Params["updatefarm"].UserTypedIn)
                Common.WebApplications.SetBackConnectionHostNames.SetBackConnectionRegKey(Common.WebApplications.SetBackConnectionHostNames.GetUrls());
            else
            {
                SPTimerService timerService = SPFarm.Local.TimerService;
                if (null == timerService)
                {
                    throw new SPException("The Farms timer service cannot be found.");
                }
                Common.WebApplications.SetBackConnectionHostNamesTimerJob job = new Common.WebApplications.SetBackConnectionHostNamesTimerJob(timerService);

                string user = Params["username"].Value;
                if (user.IndexOf('\\') < 0)
                    user = Environment.UserDomainName + "\\" + user;

                List<string> urls = Common.WebApplications.SetBackConnectionHostNames.GetUrls();

                job.SubmitJob(user, Params["password"].Value + "", urls);

                output += "Timer job successfully created.";
            }

            return (int)ErrorCodes.NoError;
        }

        public override void Validate(StringDictionary keyValues)
        {
            base.Validate(keyValues);

            if (Params["updatefarm"].UserTypedIn)
            {
                if (!Params["username"].UserTypedIn)
                    throw new SPSyntaxException("A valid username with rights to edit the registry is required.");
            }
        }

        #endregion

       

    }
}
