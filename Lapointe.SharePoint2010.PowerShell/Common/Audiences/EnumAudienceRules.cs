using System;
using System.Collections;
using System.Collections.Specialized;
using System.Text;
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Common.Audiences
{
    internal class EnumAudienceRules
    {
        /// <summary>
        /// Returns an XML structure containing all the rules associated with the audience.
        /// </summary>
        /// <param name="audienceName">Name of the audience.</param>
        /// <param name="includeAllAttributes">if set to <c>true</c> [include all attributes].</param>
        /// <returns></returns>
        internal static string EnumRules(SPServiceContext context, string audienceName, bool includeAllAttributes)
        {
            AudienceManager manager = new AudienceManager(context);

            if (!manager.Audiences.AudienceExist(audienceName))
            {
                throw new SPException("Audience name does not exist");
            }

            Audience audience = manager.Audiences[audienceName];

            ArrayList audienceRules = audience.AudienceRules;

            if (audienceRules == null || audienceRules.Count == 0)
                return "The audience contains no rules.";

            string rulesXml = "<rules>\r\n";
            foreach (AudienceRuleComponent rule in audienceRules)
            {
                if (includeAllAttributes)
                {
                    rulesXml += string.Format("\t<rule field=\"{1}\" op=\"{0}\" value=\"{2}\" />\r\n", rule.Operator, rule.LeftContent, rule.RightContent);
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
                            rulesXml += string.Format("\t<rule field=\"{1}\" op=\"{0}\" value=\"{2}\" />\r\n", rule.Operator, rule.LeftContent, rule.RightContent);
                            break;
                        case "reports under":
                        case "member of":
                            rulesXml += string.Format("\t<rule op=\"{0}\" value=\"{1}\" />\r\n", rule.Operator, rule.RightContent);
                            break;
                        case "and":
                        case "or":
                        case "(":
                        case ")":
                            rulesXml += string.Format("\t<rule op=\"{0}\" />\r\n", rule.Operator);
                            break;
                    }
                }
            }
            rulesXml += "</rules>";
            return rulesXml;
        }

    }
}
