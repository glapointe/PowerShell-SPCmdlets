using System.Collections.Specialized;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Quotas
{
	public class CreateQuotaTemplate : SPOperation
	{
        /// <summary>
        /// Initializes a new instance of the <see cref="CreateQuotaTemplate"/> class.
        /// </summary>
        public CreateQuotaTemplate()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("name", "n", true, null, new SPNonEmptyValidator(), "Please specify the quota template name."));
            parameters.Add(new SPParam("storagemaxlevel", "max", false, null, new SPLongRangeValidator(0, (long.MaxValue / 1024) / 1024), "Please specify the maximum storage level."));
            parameters.Add(new SPParam("storagewarninglevel", "warn", false, null, new SPLongRangeValidator(0, (long.MaxValue / 1024) / 1024), "Please specify the level at which a warning email should be sent."));
            
            StringBuilder sb = new StringBuilder();
            sb.Append("\n\r\n\r\nCreates a new quota template\r\n\r\nParameters:");
            sb.Append("\r\n\t-name <quota name>");
            sb.Append("\r\n\t[-storagemaxlevel <maximum storage level in megabytes>]");
            sb.Append("\r\n\t[-storagewarninglevel <warning level in megabytes>]");
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

            

			string name = Params["name"].Value;
			long storagemaxlevel = 0;
			long storagewarninglevel = 0;
			try
			{
                if (Params["storagemaxlevel"].UserTypedIn)
                    storagemaxlevel = long.Parse(Params["storagemaxlevel"].Value) * 1024 * 1024; // Convert to bytes
                if (Params["storagewarninglevel"].UserTypedIn)
                    storagewarninglevel = long.Parse(Params["storagewarninglevel"].Value) * 1024 * 1024;
			}
			catch
			{
                throw new SPException("Please specify levels in terms of whole numbers between 0 and " + (long.MaxValue / 1024) / 1024 + ".");
			}

            SPFarm farm = SPFarm.Local;
            SPWebService webService = farm.Services.GetValue<SPWebService>("");

			SPQuotaTemplateCollection quotaColl = webService.QuotaTemplates;

            if (quotaColl[name] != null)
            {
                throw new SPException("The template specified already exists.");
            }

            SPQuotaTemplate newTemplate = new SPQuotaTemplate();

            newTemplate.Name = name;
            if (Params["storagemaxlevel"].UserTypedIn)
                newTemplate.StorageMaximumLevel = storagemaxlevel;
            if (Params["storagewarninglevel"].UserTypedIn)
                newTemplate.StorageWarningLevel = storagewarninglevel;

            quotaColl.Add(newTemplate); 

			return (int)ErrorCodes.NoError;
		}

		#endregion
	}
}
