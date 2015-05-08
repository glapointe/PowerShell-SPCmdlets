using System;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Common.Audiences
{
    internal class ExportAudiences
    {
        /// <summary>
        /// Returns an XML structure containing all the audience details.
        /// </summary>
        /// <param name="audienceName">Name of the audience.</param>
        /// <param name="includeAllAttributes">if set to <c>true</c> [include all attributes].</param>
        /// <returns></returns>
        internal static string Export(SPServiceContext context, string audienceName, bool includeAllAttributes)
        {
            AudienceManager manager = new AudienceManager(context);

            if (!string.IsNullOrEmpty(audienceName) && !manager.Audiences.AudienceExist(audienceName))
            {
                throw new SPException("Audience name does not exist");
            }

            StringBuilder sb = new StringBuilder();
            XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
            xmlWriter.Formatting = Formatting.Indented;

            xmlWriter.WriteStartElement("Audiences");

            if (!string.IsNullOrEmpty(audienceName))
            {
                Audience audience = manager.Audiences[audienceName];
                ExportAudience(xmlWriter, audience, includeAllAttributes);
            }
            else
            {
                foreach (Audience audience in manager.Audiences)
                    ExportAudience(xmlWriter, audience, includeAllAttributes);
            }

            xmlWriter.WriteEndElement(); // Audiences
            xmlWriter.Flush();
            return sb.ToString();
        }

        /// <summary>
        /// Exports the audience.
        /// </summary>
        /// <param name="xmlWriter">The XML writer.</param>
        /// <param name="audience">The audience.</param>
        /// <param name="includeAllAttributes">if set to <c>true</c> [include all attributes].</param>
        private static void ExportAudience(XmlWriter xmlWriter, Audience audience, bool includeAllAttributes)
        {
            xmlWriter.WriteStartElement("Audience");
            xmlWriter.WriteAttributeString("AudienceDescription", audience.AudienceDescription);
            xmlWriter.WriteAttributeString("AudienceID", audience.AudienceID.ToString());
            xmlWriter.WriteAttributeString("AudienceName", audience.AudienceName);
            xmlWriter.WriteAttributeString("CreateTime", audience.CreateTime.ToString());
            xmlWriter.WriteAttributeString("GroupOperation", audience.GroupOperation.ToString());
            xmlWriter.WriteAttributeString("LastCompilation", audience.LastCompilation.ToString());
            xmlWriter.WriteAttributeString("LastError", audience.LastError);
            xmlWriter.WriteAttributeString("LastPropertyUpdate", audience.LastPropertyUpdate.ToString());
            xmlWriter.WriteAttributeString("LastRuleUpdate", audience.LastRuleUpdate.ToString());
            xmlWriter.WriteAttributeString("MemberShipCount", audience.MemberShipCount.ToString());
            xmlWriter.WriteAttributeString("OwnerAccountName", audience.OwnerAccountName);


            ArrayList audienceRules = audience.AudienceRules;
            xmlWriter.WriteStartElement("rules");
            if (audienceRules != null && audienceRules.Count > 0)
            {
                foreach (AudienceRuleComponent rule in audienceRules)
                {
                    xmlWriter.WriteStartElement("rule");
                    if (includeAllAttributes)
                    {
                        xmlWriter.WriteAttributeString("field", rule.LeftContent);
                        xmlWriter.WriteAttributeString("op", rule.Operator);
                        xmlWriter.WriteAttributeString("value", rule.RightContent);
                    }
                    else
                    {
                        switch (rule.Operator.ToLowerInvariant())
                        {
                            case "=":
                            case ">":
                            case ">=":
                            case "<":
                            case "<=":
                            case "contains":
                            case "<>":
                            case "not contains":
                                xmlWriter.WriteAttributeString("field", rule.LeftContent);
                                xmlWriter.WriteAttributeString("op", rule.Operator);
                                xmlWriter.WriteAttributeString("value", rule.RightContent);
                                break;
                            case "reports under":
                            case "member of":
                                xmlWriter.WriteAttributeString("op", rule.Operator);
                                xmlWriter.WriteAttributeString("value", rule.RightContent);
                                break;
                            case "and":
                            case "or":
                            case "(":
                            case ")":
                                xmlWriter.WriteAttributeString("op", rule.Operator);
                                break;
                        }
                    }
                    xmlWriter.WriteEndElement(); // rule
                }
            }
            xmlWriter.WriteEndElement(); // rules
            xmlWriter.WriteEndElement(); // Audience
        }
    }
}
