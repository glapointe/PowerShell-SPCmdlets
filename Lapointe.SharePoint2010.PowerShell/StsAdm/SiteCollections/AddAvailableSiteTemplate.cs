using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
#if MOSS
using Microsoft.SharePoint.Publishing;
#endif
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    public class AddAvailableSiteTemplate : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AddAvailableSiteTemplate"/> class.
        /// </summary>
        public AddAvailableSiteTemplate()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site collection."));
            parameters.Add(new SPParam("template", "t", false, null, new SPNonEmptyValidator(), "Please specify the template name (use enuminstalledsitetemplates to see what is installed)."));
            parameters.Add(new SPParam("lcid", "l", false, null, new SPRegexValidator(@"^\d{4}$"), "Please specify the locale id (defaults to cross language)."));
            parameters.Add(new SPParam("resetallsubsites", "reset", false, null, null));
            parameters.Add(new SPParam("allowalltemplates", "allowall", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append(
                "\r\n\r\nAdds a site template to the list of available templates for the given site collection.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <site collection url>");
            sb.Append("\r\n\t[-template <template name> / -allowalltemplates]");
            sb.Append("\r\n\t[-lcid <locale id>]");
#if MOSS
            sb.Append("\r\n\t[-resetallsubsites]");
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

            

            SPBinaryParameterValidator.Validate("template", Params["template"].Value, "allowalltemplates",
                                    (Params["allowalltemplates"].UserTypedIn ? "true" : Params["allowalltemplates"].Value));

            string url = Params["url"].Value.TrimEnd('/');
            string templateName = Params["template"].Value;
            bool resetAllSubsites = Params["resetallsubsites"].UserTypedIn;

            using (SPSite site = new SPSite(url))
            {
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {
                    uint lcid = web.Language;
                    if (!string.IsNullOrEmpty(keyValues["lcid"]))
                        lcid = uint.Parse(keyValues["lcid"]);
                    bool localeProvided = keyValues.ContainsKey("lcid");

#if MOSS
                    PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);

                    if (Params["allowalltemplates"].UserTypedIn)
                    {
                        pubweb.AllowAllWebTemplates(resetAllSubsites);
                        pubweb.Update();
                        return (int)ErrorCodes.NoError;
                    }
#else
                    if (Params["allowalltemplates"].UserTypedIn)
                    {
                        web.AllowAllWebTemplates();
                        web.Update();
                        return (int)ErrorCodes.NoError;
                    }
#endif
                    SPWebTemplateCollection templateColl;
                    if (localeProvided)
                    {
                        templateColl = web.GetAvailableWebTemplates(lcid);
                    }
                    else
                    {
                        templateColl = web.GetAvailableCrossLanguageWebTemplates();
                    }
                    
                    bool exists;
                    try
                    {
                        exists = (templateColl[templateName] != null);
                    }
                    catch (ArgumentException)
                    {
                        exists = false;
                    }
                    if (exists && !web.AllWebTemplatesAllowed)
                    {
                        output = "Template is already installed.";
                        return (int)ErrorCodes.GeneralError;
                    }

                    Collection<SPWebTemplate> list = new Collection<SPWebTemplate>();
                    if (!web.AllWebTemplatesAllowed)
                    {
                        foreach (SPWebTemplate existingTemplate in templateColl)
                        {
                            list.Add(existingTemplate);
                        }
                    }
                    SPWebTemplate newTemplate = GetWebTemplate(site, lcid, templateName);
                    if (newTemplate == null)
                    {
                        output = "Template not found.";
                        return (int)ErrorCodes.GeneralError;
                    }
                    else
                        list.Add(newTemplate);

#if MOSS
                    if (!localeProvided)
                    {
                        pubweb.SetAvailableCrossLanguageWebTemplates(list, resetAllSubsites);
                    }
                    else
                    {
                        pubweb.SetAvailableWebTemplates(list, lcid, resetAllSubsites);
                    }
#else
                    if (!localeProvided)
                    {
                        web.SetAvailableCrossLanguageWebTemplates(list);
                    }
                    else
                    {
                        web.SetAvailableWebTemplates(list, lcid);
                    }

#endif
                }
            }

            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Gets the web template.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="lcid">The lcid.</param>
        /// <param name="webTemplateName">Name of the web template.</param>
        /// <returns></returns>
        internal static SPWebTemplate GetWebTemplate(SPSite site, uint lcid, string webTemplateName)
        {
            SPWebTemplateCollection webTemplates = site.GetWebTemplates(lcid);
            SPWebTemplateCollection customWebTemplates = site.GetCustomWebTemplates(lcid);
            SPWebTemplate template = null;
            try
            {
                template = webTemplates[webTemplateName];
            }
            catch (ArgumentException)
            {
                try
                {
                    return customWebTemplates[webTemplateName];
                }
                catch (ArgumentException)
                {
                    return template;
                }
            }
            return template;
        }


        #endregion
    }
}
