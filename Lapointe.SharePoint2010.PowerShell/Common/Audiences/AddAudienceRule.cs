using System;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
using Microsoft.Office.Server.Search.Administration;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Common.Audiences
{
    public enum AppendOp
    {
        AND, OR
    }

    internal class AddAudienceRule
    {
        /// <summary>
        /// Adds the rules.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="audienceName">Name of the audience.</param>
        /// <param name="rules">The rules.</param>
        /// <param name="clearExistingRules">if set to <c>true</c> [clear existing rules].</param>
        /// <param name="compile">if set to <c>true</c> [compile].</param>
        /// <param name="groupExisting">if set to <c>true</c> [group existing].</param>
        /// <param name="appendOp">The append op.</param>
        /// <returns></returns>
        internal static ArrayList AddRules(SPServiceContext context, string audienceName, string rules, bool clearExistingRules, bool compile, bool groupExisting, AppendOp appendOp)
        {
            AudienceManager manager = new AudienceManager(context);

            if (!manager.Audiences.AudienceExist(audienceName))
            {
                throw new SPException("Audience name does not exist");
            }

            Audience audience = manager.Audiences[audienceName];
            /*
            Operator        Need left and right operands (not a group operator) 
            =               Yes 
            >               Yes 
            >=              Yes 
            <               Yes 
            <=              Yes 
            Contains        Yes 
            Reports Under   Yes (Left operand must be 'Everyone') 
            <>              Yes 
            Not contains    Yes 
            AND             No 
            OR              No 
            (               No 
            )               No 
            Member Of       Yes (Left operand must be 'DL') 
            */
            XmlDocument rulesDoc = new XmlDocument();
            rulesDoc.LoadXml(rules);

            ArrayList audienceRules = audience.AudienceRules;
            bool ruleListNotEmpty = false;

            if (audienceRules == null || clearExistingRules)
                audienceRules = new ArrayList();
            else
                ruleListNotEmpty = true;

            //if the rule is not emply, start with a group operator 'AND' to append
            if (ruleListNotEmpty)
            {
                if (groupExisting)
                {
                    audienceRules.Insert(0, new AudienceRuleComponent(null, "(", null));
                    audienceRules.Add(new AudienceRuleComponent(null, ")", null));
                }

                audienceRules.Add(new AudienceRuleComponent(null, appendOp.ToString(), null));
            }

            if (rulesDoc.SelectNodes("//rule") == null || rulesDoc.SelectNodes("//rule").Count == 0)
                throw new ArgumentException("No rules were supplied.");

            foreach (XmlElement rule in rulesDoc.SelectNodes("//rule"))
            {
                string op = rule.GetAttribute("op").ToLowerInvariant();
                string field = null;
                string val = null;
                bool valIsRequired = true;
                bool fieldIsRequired = false;

                switch (op)
                {
                    case "=":
                    case ">":
                    case ">=":
                    case "<":
                    case "<=":
                    case "contains":
                    case "<>":
                    case "not contains":
                        field = rule.GetAttribute("field");
                        val = rule.GetAttribute("value");
                        fieldIsRequired = true;
                        break;
                    case "reports under":
                        field = "Everyone";
                        val = rule.GetAttribute("value");
                        break;
                    case "member of":
                        field = "DL";
                        val = rule.GetAttribute("value");
                        break;
                    case "and":
                    case "or":
                    case "(":
                    case ")":
                        valIsRequired = false;
                        break;
                    default:
                        throw new ArgumentException(string.Format("Rule operator is invalid: {0}", rule.GetAttribute("op")));
                }
                if (valIsRequired && string.IsNullOrEmpty(val))
                    throw new ArgumentNullException(string.Format("Rule value attribute is missing or invalid: {0}", rule.GetAttribute("value")));

                if (fieldIsRequired && string.IsNullOrEmpty(field))
                    throw new ArgumentNullException(string.Format("Rule field attribute is missing or invalid: {0}", rule.GetAttribute("field")));
                
                AudienceRuleComponent r0 = new AudienceRuleComponent(field, op, val);
                audienceRules.Add(r0);
            }

            audience.AudienceRules = audienceRules;
            audience.Commit();
            
            if (compile)
            {
                SPServiceApplication svcApp = Utilities.GetUserProfileServiceApplication(context);
                if (svcApp != null)
                    CompileAudience(svcApp.Id, audience.AudienceName);
            }

            return audienceRules;
        }

        /// <summary>
        /// Compiles the audience.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="audienceName">Name of the audience.</param>
        internal static void CompileAudience(Guid id, string audienceName)
        {
            string[] args = new string[4];
            args[0] = id.ToString();
            args[1] = "1"; //"1" = start job, "0" = stop job 
            args[2] = "1"; //"1" = full compilation, "0" = incremental compilation (optional, default = 0) 
            args[3] = audienceName;

            AudienceJob.RunAudienceJob(args);
        }


    }
}
