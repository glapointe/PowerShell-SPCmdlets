using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing;
using System.Management.Automation;
using Lapointe.SharePoint.PowerShell.StsAdm.Lists;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Pages
{
    public class CreatePublishingPage : SPOperation
    {
         /// <summary>
        /// Initializes a new instance of the <see cref="CreatePublishingPage"/> class.
        /// </summary>
        public CreatePublishingPage()
        {
            SPParamCollection parameters = new SPParamCollection();
            StringBuilder sb = new StringBuilder();

#if MOSS
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator()));
            parameters.Add(new SPParam("name", "n", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("title", "t", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("layout", "l", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("fielddata", "fd", false, null, new SPNonEmptyValidator()));

            sb.Append("\r\n\r\nCreates a new publishing page.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <url to the publishing web within which to create the page>");
            sb.Append("\r\n\t-name <the filename of the page to create (do not include the extension)>");
            sb.Append("\r\n\t-title <the page title>");
            sb.Append("\r\n\t-layout <the filename of the page layout to use>");
            sb.Append("\r\n\t[-fielddata <semi-colon separated list of key value pairs: \"Field1=Val1;Field2=Val2\"> (use ';;' to escape semi-colons in data values)]");

#else
            sb.Append(NOT_VALID_FOR_FOUNDATION);
#endif
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
        /// Executes the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public override int Execute(string command, StringDictionary keyValues, out string output)
        {
            output = string.Empty;

#if !MOSS
            output = NOT_VALID_FOR_FOUNDATION;
            return (int)ErrorCodes.GeneralError;
#endif


            string url = Params["url"].Value;
            if (url != null)
                url = url.TrimEnd('/');

            string pageName = Params["name"].Value;
            string title = Params["title"].Value;
            string layout = Params["layout"].Value;

            string fieldData = Params["fielddata"].Value;
            Dictionary<string, string> fieldDataCollection = new Dictionary<string, string>();
            if (!string.IsNullOrEmpty(fieldData))
            {
                fieldData = fieldData.Replace(";;", "_STSADM_CREATEPUBLISHINGPAGE_SEMICOLON_");
                foreach (string s in fieldData.Split(';'))
                {
                    string[] data = s.Split(new char[] { '=' }, 2);
                    fieldDataCollection.Add(data[0].Trim(), data[1].Trim().Replace("_STSADM_CREATEPUBLISHINGPAGE_SEMICOLON_", ";"));
                }
            }


            Common.Pages.CreatePublishingPage.CreatePage(url, pageName, title, layout, fieldDataCollection, false);

            return (int)ErrorCodes.NoError;
        }


        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
#if !MOSS
            return;
#endif

            if (Params["fielddata"].UserTypedIn)
            {
                string fieldData = Params["fielddata"].Value;
                fieldData = fieldData.Replace(";;", "_STSADM_CREATEPUBLISHINGPAGE_SEMICOLON_");

                foreach (string prop in fieldData.Split(';'))
                {
                    if (prop.Split(new char[] { '=' }, 2).Length != 2)
                        throw new SPSyntaxException(
                            "The format of the fielddata parameter is incorrect: \"Field1=Val1;Field2=Val2\" (use ';;' to escape semi-colons in data values)");
                }
            }

            base.Validate(keyValues);
        }

        #endregion

       
    }
}
