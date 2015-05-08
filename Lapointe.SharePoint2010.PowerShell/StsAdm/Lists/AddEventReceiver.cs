using System;
using System.Collections.Specialized;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Lists
{
    public class AddEventReceiver : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AddEventReceiver"/> class.
        /// </summary>
        public AddEventReceiver()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the URL to the web or list."));
            parameters.Add(new SPParam("assembly", "a", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("class", "c", true, null, new SPNonEmptyValidator()));
            SPEnumValidator typeValidator = new SPEnumValidator(typeof(SPEventReceiverType));
            parameters.Add(new SPParam("type", "type", true, null, typeValidator));
            SPEnumValidator targetValidator = new SPEnumValidator(typeof(Common.Lists.AddEventReceiver.TargetEnum));
            parameters.Add(new SPParam("target", "target", false, "list", targetValidator));
            parameters.Add(new SPParam("contenttype", "ct", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("sequence", "seq", false, null, new SPIntRangeValidator(0, int.MaxValue)));
            parameters.Add(new SPParam("name", "n", false, null, new SPNonEmptyValidator()));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nAdds an event receiver to a list, web, or content type.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web or list URL>");
            sb.Append("\r\n\t-assembly <assembly>");
            sb.Append("\r\n\t-class <class name>");
            sb.AppendFormat("\r\n\t-type <{0}>", typeValidator.DisplayValue);
            sb.AppendFormat("\r\n\t-target <{0}>", targetValidator.DisplayValue);
            sb.Append("\r\n\t[-contenttype <content type name if target is ContentType>]");
            sb.Append("\r\n\t[-sequence <sequence number>]");
            sb.Append("\r\n\t[-name <the name to give to the event receiver>]");
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
            Common.Lists.AddEventReceiver.TargetEnum target = (Common.Lists.AddEventReceiver.TargetEnum) Enum.Parse(typeof (Common.Lists.AddEventReceiver.TargetEnum), Params["target"].Value, true);
            SPEventReceiverType type = (SPEventReceiverType)Enum.Parse(typeof(SPEventReceiverType), Params["type"].Value, true);
            string assembly = Params["assembly"].Value;
            string className = Params["class"].Value;
            string contentTypeName = Params["contenttype"].Value;
            string name = Params["name"].Value;
            int sequence = -1;
            if (Params["sequence"].UserTypedIn)
                sequence = int.Parse(Params["sequence"].Value);

            Common.Lists.AddEventReceiver.Add(url, contentTypeName, target, assembly, className, type, sequence, name);

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            if (Params["target"].Validate())
            {
                Common.Lists.AddEventReceiver.TargetEnum target = (Common.Lists.AddEventReceiver.TargetEnum) Enum.Parse(typeof (Common.Lists.AddEventReceiver.TargetEnum), Params["target"].Value, true);
                if (target == Common.Lists.AddEventReceiver.TargetEnum.ContentType)
                    Params["contenttype"].IsRequired = true;
            }
            base.Validate(keyValues);
        }
        #endregion
    }
}
