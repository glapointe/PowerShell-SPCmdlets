using System;
using System.Collections.Generic;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Threading;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.TimerJobs
{
    public class ExecAdmSvcJobs : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExecAdmSvcJobs"/> class.
        /// </summary>
        public ExecAdmSvcJobs()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("local", "l"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nExecutes pending timer jobs on all servers in the farm.\r\n\r\n\r\n\r\nParameters:");
            sb.Append("\r\n\t[-local]");
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
        public override int Execute(string command, System.Collections.Specialized.StringDictionary keyValues, out string output)
        {
            output = string.Empty;

            Common.TimerJobs.ExecAdmSvcJobs.Execute(Params["local"].UserTypedIn);

            return (int)ErrorCodes.NoError;
        }
    }
}
