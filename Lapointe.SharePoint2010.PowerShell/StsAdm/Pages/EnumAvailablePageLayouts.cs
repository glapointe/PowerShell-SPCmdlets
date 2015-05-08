using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Pages
{
    public class EnumAvailablePageLayouts : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EnumAvailablePageLayouts"/> class.
        /// </summary>
        public EnumAvailablePageLayouts()
        {
            SPParamCollection parameters = new SPParamCollection();
            StringBuilder sb = new StringBuilder();

#if MOSS
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site collection"));
            
            sb.Append("\r\n\r\nReturns the list of page layouts available for the given site collection.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <site collection url>");
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
        /// Runs the specified command.
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

            string url = Params["url"].Value.TrimEnd('/');

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.AppendChild(xmlDoc.CreateElement("PageLayouts"));

            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            {
                PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                foreach (PageLayout layout in pubweb.GetAvailablePageLayouts())
                {
                    XmlElement layoutNode = xmlDoc.CreateElement("PageLayout");

                    XmlElement node = xmlDoc.CreateElement("Name");
                    node.InnerText = layout.Name;
                    layoutNode.AppendChild(node);

                    node = xmlDoc.CreateElement("Title");
                    node.InnerText = layout.Title;
                    layoutNode.AppendChild(node);

                    node = xmlDoc.CreateElement("Id");
                    node.InnerText = layout.ListItem.ID.ToString();
                    layoutNode.AppendChild(node);

                    node = xmlDoc.CreateElement("AssociatedContentType");
                    if (layout.AssociatedContentType != null)
                        node.InnerText = layout.AssociatedContentType.Name;
                    layoutNode.AppendChild(node);

                    node = xmlDoc.CreateElement("ContentType");
                    if (layout.ListItem[FieldId.ContentType] != null)
                        node.InnerText = layout.ListItem[FieldId.ContentType].ToString();
                    layoutNode.AppendChild(node);

                    node = xmlDoc.CreateElement("Hidden");
                    if (layout.ListItem[FieldId.Hidden] != null)
                        node.InnerText = layout.ListItem[FieldId.Hidden].ToString();
                    else
                        node.InnerText = "false";
                    layoutNode.AppendChild(node);

                    node = xmlDoc.CreateElement("FileUrl");
                    if (layout.ListItem.File != null)
                        node.InnerText = layout.ListItem.File.Url;
                    layoutNode.AppendChild(node);

                    xmlDoc.DocumentElement.AppendChild(layoutNode);
                }
            }

            StringBuilder sb = new StringBuilder();
            XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
            xmlWriter.Formatting = Formatting.Indented;
            xmlDoc.WriteContentTo(xmlWriter);
            xmlWriter.Flush();
            output += sb.ToString();

            return (int)ErrorCodes.NoError;
        }

        #endregion

    }
}
