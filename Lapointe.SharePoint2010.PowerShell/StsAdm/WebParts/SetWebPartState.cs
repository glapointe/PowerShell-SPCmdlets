using System.Collections;
using System.Collections.Specialized;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint.StsAdmin;
using Lapointe.SharePoint.PowerShell.Common.WebParts;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebParts
{
    public class SetWebPartState : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SetWebPartState"/> class.
        /// </summary>
        public SetWebPartState()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPNonEmptyValidator(), "Please specify the page url."));
            parameters.Add(new SPParam("id", "id", false, null, new SPNonEmptyValidator(), "Please specify the web part ID."));
            parameters.Add(new SPParam("title", "t", false, null, new SPNonEmptyValidator(), "Please specify the web part title."));
            parameters.Add(new SPParam("delete", "del", false, null, null));
            parameters.Add(new SPParam("open", "open", false, null, null));
            parameters.Add(new SPParam("close", "close", false, null, null));
            parameters.Add(new SPParam("zone", "z", false, null, new SPNonEmptyValidator(), "Please specify the zone to move to."));
            parameters.Add(new SPParam("zoneindex", "zi", false, null, new SPIntRangeValidator(0, int.MaxValue)));
            parameters.Add(new SPParam("properties", "props", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("propertyseperator", "propsep", false, ",", new SPNonEmptyValidator()));
            parameters.Add(new SPParam("propertiesfile", "propsfile", false, null, new SPFileExistsValidator()));
            parameters.Add(new SPParam("publish", "p", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nOpens, Closes, Adds, or Deletes a web part on a page.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <web part page URL>");
            sb.Append("\r\n\t{-id <web part ID> |\r\n\t -title <web part title>}");
            sb.Append("\r\n\t{-delete |\r\n\t -open |\r\n\t -close}");
            sb.Append("\r\n\t[-zone <zone ID>]");
            sb.Append("\r\n\t[-zoneindex <zone index>]");
            sb.Append("\r\n\t{[-properties <comma separated list of key value pairs: \"Prop1=Val1,Prop2=Val2\">] | ");
            sb.Append("\r\n\t [-propertiesfile <path to a file with xml property settings (<Properties><Property Name=\"Name1\">Value1</Property><Property Name=\"Name2\">Value2</Property></Properties>)>]}");
            sb.Append("\r\n\t[-propertyseperator <string to use as a property seperator if the properties parameter is used - default is a comma>]");
            sb.Append("\r\n\t[-publish]");

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

            string url = Params["url"].Value;
            SetWebPartStateAction action = SetWebPartStateAction.Update;
            if (Params["open"].UserTypedIn)
                action = SetWebPartStateAction.Open;
            else if (Params["close"].UserTypedIn)
                action = SetWebPartStateAction.Close;
            else if (Params["delete"].UserTypedIn)
                action = SetWebPartStateAction.Delete;

            string webPartId = Params["id"].Value;
            string webPartTitle = Params["title"].Value;
            string webPartZone = Params["zone"].Value;
            string webPartZoneIndex = Params["zoneindex"].Value;
            bool publish = Params["publish"].UserTypedIn;
            string properties = Params["properties"].Value;
            string propertiesFile = Params["propertiesfile"].Value;
            string propertySeperator = Params["propertyseperator"].Value;

            Hashtable props = null;
            if (!string.IsNullOrEmpty(properties) || !string.IsNullOrEmpty(propertiesFile))
            {
                props = Common.WebParts.SetWebPartState.GetPropertiesArray(properties, propertiesFile, propertySeperator);
            }
            Common.WebParts.SetWebPartState.SetWebPart(url, action, webPartId, webPartTitle, webPartZone, webPartZoneIndex, props, publish);

            return (int)ErrorCodes.NoError;
        }


        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            SPMultiParameterValidator.Validate(Params, new string[] {"add", "close", "open", "delete"}, 1, 1);

            SPBinaryParameterValidator.Validate("id", Params["id"].Value, "title", Params["title"].Value);
            if (Params["properties"].UserTypedIn)
            {
                string properties = Params["properties"].Value;
                string seperator = Params["propertyseperator"].Value;
                properties = properties.Replace(seperator + seperator, "[STSADM_COMMA]");
                string[] props = properties.Split(seperator.ToCharArray());
                foreach (string prop in props)
                {
                    if (prop.Split(new[] {'='}, 2).Length != 2)
                        throw new SPSyntaxException(
                            "The format of the properties parameter is incorrect: \"prop1=val1" + seperator + "prop2=val2\"");
                }
            }
            if (Params["properties"].UserTypedIn && Params["propertiesfile"].UserTypedIn)
                throw new SPSyntaxException("The properties parameter and propertiesfile parameters are incompatible.");

            base.Validate(keyValues);
        }

        #endregion
    }
}
