using System;
using System.Collections.Specialized;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class AddList : SPOperation
    {
         /// <summary>
        /// Initializes a new instance of the <see cref="AddList"/> class.
        /// </summary>
        public AddList()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the web url to add the list to."));
            parameters.Add(new SPParam("urlname", "name", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("title", "title", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("featureid", "fid", true, null, new SPGuidValidator()));
            parameters.Add(new SPParam("templatetype", "type", true, null, new SPIntRangeValidator(0, int.MaxValue)));
            parameters.Add(new SPParam("description", "desc", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("doctemplatetype", "doc", false, null, new SPIntRangeValidator(0, int.MaxValue)));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nAdds a list to a web.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web url to add the list to>");
            sb.Append("\r\n\t-urlname <the name that will appear in the URL>");
            sb.Append("\r\n\t-title <list title>");
            sb.Append("\r\n\t-featureid <the feature ID to which the list definition belongs>");
            sb.Append("\r\n\t-templatetype <an integer corresponding to the list definition type>");
            sb.Append("\r\n\t[-description <list description>]");
            sb.Append("\r\n\t[-doctemplatetype <the ID for the document template type>]");
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

            try
            {
                string url = Params["url"].Value.TrimEnd('/');
                string urlName = Params["urlname"].Value;
                string title = Params["title"].Value;
                string desc = Params["description"].Value;
                Guid featureId = new Guid(Params["featureid"].Value);
                int templateType = int.Parse(Params["templatetype"].Value);
                string docTemplateType = Params["doctemplatetype"].Value;

                Common.Lists.AddList.Add(url, urlName, title, desc, featureId, templateType, docTemplateType);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured executing the command: {0}\r\n{1}", ex.Message, ex.StackTrace);
                return (int)ErrorCodes.GeneralError;
            }
            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
