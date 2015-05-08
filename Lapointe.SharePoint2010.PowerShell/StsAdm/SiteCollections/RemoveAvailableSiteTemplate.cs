using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
#if MOSS
using Microsoft.SharePoint.Publishing;
#endif
using Microsoft.SharePoint.StsAdmin;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    public class RemoveAvailableSiteTemplate : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RemoveAvailableSiteTemplate"/> class.
        /// </summary>
        public RemoveAvailableSiteTemplate()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site collection"));
            parameters.Add(new SPParam("template", "t", false, null, new SPNonEmptyValidator(), "Please specify the template name (use enuminstalledsitetemplates to see what is installed)."));
            parameters.Add(new SPParam("lcid", "l", false, null, new SPRegexValidator(@"^\d{4}$"), "Please specify the locale id (defaults to cross language)."));
            parameters.Add(new SPParam("resetallsubsites", "reset", false, null, null));
            parameters.Add(new SPParam("allowalltemplates", "allowall", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nRemoves a site template from the list of available templates for the given site collection.\r\n\r\nParameters:");
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
                    bool localeProvided = Params["lcid"].UserTypedIn;
                    if (localeProvided)
                        lcid = uint.Parse(Params["lcid"].Value);
                    bool exists = false;

#if MOSS
                    PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);
                    
                    if (Params["allowalltemplates"].UserTypedIn)
                    {
                        pubweb.AllowAllWebTemplates(resetAllSubsites);
                        pubweb.Update();
                        return (int)ErrorCodes.NoError;
                    }
                    SPWebTemplateCollection existingLanguageNeutralTemplatesCollection = pubweb.GetAvailableCrossLanguageWebTemplates();
                    SPWebTemplateCollection existingLanguageSpecificTemplatesCollection = pubweb.GetAvailableWebTemplates(lcid);
#else
                    if (Params["allowalltemplates"].UserTypedIn)
                    {
                        web.AllowAllWebTemplates();
                        web.Update();
                        return (int)ErrorCodes.NoError;
                    }
                    SPWebTemplateCollection existingLanguageNeutralTemplatesCollection = web.GetAvailableCrossLanguageWebTemplates();
                    SPWebTemplateCollection existingLanguageSpecificTemplatesCollection = web.GetAvailableWebTemplates(lcid);
#endif


                    Collection<SPWebTemplate> newLanguageNeutralTemplatesCollection = new Collection<SPWebTemplate>();
                    Collection<SPWebTemplate> newLanguageSpecificTemplatesCollection = new Collection<SPWebTemplate>();

                    foreach (SPWebTemplate existingTemplate in existingLanguageNeutralTemplatesCollection)
                    {
                        if (existingTemplate.Name == templateName && !localeProvided)
                        {
                            exists = true;
                            continue;
                        }
                        newLanguageNeutralTemplatesCollection.Add(existingTemplate);
                    }
                    foreach (SPWebTemplate existingTemplate in existingLanguageSpecificTemplatesCollection)
                    {
                        if (existingTemplate.Name == templateName && localeProvided)
                        {
                            exists = true;
                            continue;
                        }
                        newLanguageSpecificTemplatesCollection.Add(existingTemplate);
                    }
                   

                    if (!exists)
                    {
                        output = "Template is not assigned.";
                        return (int)ErrorCodes.GeneralError;
                    }

                    if (newLanguageSpecificTemplatesCollection.Count == 0 && newLanguageNeutralTemplatesCollection.Count == 0)
                    {
                        output = "There must be at least one template available.";
                        return (int)ErrorCodes.GeneralError;
                    }
#if MOSS
                    pubweb.SetAvailableCrossLanguageWebTemplates(newLanguageNeutralTemplatesCollection, resetAllSubsites);
                    pubweb.SetAvailableWebTemplates(newLanguageSpecificTemplatesCollection, lcid, resetAllSubsites);
#else
                    web.SetAvailableCrossLanguageWebTemplates(newLanguageNeutralTemplatesCollection);
                    web.SetAvailableWebTemplates(newLanguageSpecificTemplatesCollection, lcid);
#endif
                }
            }

            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
