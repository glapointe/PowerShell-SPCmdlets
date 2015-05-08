using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Themes
{
    public class ApplyTheme : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ApplyTheme"/> class.
        /// </summary>
        public ApplyTheme()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator()));
            parameters.Add(new SPParam("theme", "theme", true, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("recurse", "r"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nApplies the specified theme to the specified site.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <url of the web to update>");

            sb.AppendFormat("\r\n\t-theme <id of the theme to apply (see {0}\\Layouts\\[LCID]\\SPThemes.xml for template IDs>", Utilities.GetGenericSetupPath("Template"));
            sb.Append("\r\n\t[-recurse (applies change to all sub-webs)]");
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
            string theme = Params["theme"].Value;

            ApplyThemeToWeb(theme, url, Params["recurse"].UserTypedIn);

            return (int)ErrorCodes.NoError;
        }

        #endregion

        /// <summary>
        /// Applies the theme to web.
        /// </summary>
        /// <param name="theme">The theme.</param>
        /// <param name="url">The URL.</param>
        /// <param name="recurse">if set to <c>true</c> [recurse].</param>
        internal static void ApplyThemeToWeb(string theme, string url, bool recurse)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            {
                ApplyThemeToWeb(theme, web, recurse);
            }
        }

        /// <summary>
        /// Applies the theme to web.
        /// </summary>
        /// <param name="theme">The theme.</param>
        /// <param name="web">The web.</param>
        /// <param name="recurse">if set to <c>true</c> [recurse].</param>
        internal static void ApplyThemeToWeb(string theme, SPWeb web, bool recurse)
        {
            web.ApplyTheme(theme);

            if (recurse)
            {
                foreach (SPWeb subWeb in web.Webs)
                {
                    try
                    {
                        ApplyThemeToWeb(theme, subWeb, recurse);
                    }
                    finally
                    {
                        subWeb.Dispose();
                    }
                }

            }
        }

    }
}
