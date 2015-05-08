using System;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;

#if MOSS
using Microsoft.SharePoint.Publishing;
#endif

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    public class SetNavigationSettings : SPOperation
    {
        internal enum CurrentNavSettingsEnum
        {
            InheritParent,
            CurrentSiteAndSiblings,
            CurrentSiteOnly
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SetNavigationSettings"/> class.
        /// </summary>
        public SetNavigationSettings()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator()));
            parameters.Add(new SPParam("treeviewenabled", "tree", false, null, new SPTrueFalseValidator(), "Please specify either \"true\" or \"false\"."));
            parameters.Add(new SPParam("quicklaunchenabled", "ql", false, null, new SPTrueFalseValidator(), "Please specify either \"true\" or \"false\"."));
#if MOSS
            parameters.Add(new SPParam("currentshowsubsites", "currentsubsites", false, null, new SPTrueFalseValidator(), "Please specify either \"true\" or \"false\"."));
            parameters.Add(new SPParam("globalshowsubsites", "globalsubsites", false, null, new SPTrueFalseValidator(), "Please specify either \"true\" or \"false\"."));
            parameters.Add(new SPParam("currentshowpages", "currentpages", false, null, new SPTrueFalseValidator(), "Please specify either \"true\" or \"false\"."));
            parameters.Add(new SPParam("globalshowpages", "globalpages", false, null, new SPTrueFalseValidator(), "Please specify either \"true\" or \"false\"."));
            SPEnumValidator sortMethodValidator = new SPEnumValidator(typeof(OrderingMethod));
            parameters.Add(new SPParam("sortmethod", "sm", false, null, sortMethodValidator));
            SPEnumValidator autoSortMethodValidator = new SPEnumValidator(typeof(AutomaticSortingMethod));
            parameters.Add(new SPParam("autosortmethod", "asm", false, null, autoSortMethodValidator));
            parameters.Add(new SPParam("sortascending", "sa", false, null, new SPTrueFalseValidator(), "Please specify either \"true\" or \"false\"."));
            parameters.Add(new SPParam("inheritglobalnav", "ig", false, null, new SPTrueFalseValidator(), "Please specify either \"true\" or \"false\"."));
            SPEnumValidator currentNavValidator = new SPEnumValidator(typeof(CurrentNavSettingsEnum));
            parameters.Add(new SPParam("currentnav", "c", false, null, currentNavValidator));
#endif

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nSets the navigation settings for a web site (use gl-setnavigationnodes to change the actual nodes that appear).\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <site collection url>");
            sb.Append("\r\n\t[-treeviewenabled <true | false>]");
            sb.Append("\r\n\t[-quicklaunchenabled <true | false>]");
#if MOSS
            sb.Append("\r\n\t[-currentshowsubsites <true | false>]");
            sb.Append("\r\n\t[-globalshowsubsites <true | false>]");
            sb.Append("\r\n\t[-currentshowpages <true | false>]");
            sb.Append("\r\n\t[-globalshowpages <true | false>]");
            sb.AppendFormat("\r\n\t[-sortmethod <{0}>]", sortMethodValidator.DisplayValue);
            sb.AppendFormat("\r\n\t[-autosortmethod <{0}>]", autoSortMethodValidator.DisplayValue);
            sb.Append("\r\n\t[-sortascending <true | false>]");
            sb.Append("\r\n\t[-inheritglobalnav <true | false>]");
            sb.AppendFormat("\r\n\t[-currentnav <{0}>]", currentNavValidator.DisplayValue);
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
        public override int Execute(string command, System.Collections.Specialized.StringDictionary keyValues, out string output)
        {
            output = string.Empty;

            string url = Params["url"].Value.TrimEnd('/');

            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            {
                
                if (Params["treeviewenabled"].UserTypedIn)
                    web.TreeViewEnabled = bool.Parse(Params["treeviewenabled"].Value);

                if (Params["quicklaunchenabled"].UserTypedIn)
                    web.QuickLaunchEnabled = bool.Parse(Params["quicklaunchenabled"].Value);

#if MOSS
                PublishingWeb pubweb = PublishingWeb.GetPublishingWeb(web);
                if (Params["currentshowsubsites"].UserTypedIn)
                    pubweb.Navigation.CurrentIncludeSubSites = bool.Parse(Params["currentshowsubsites"].Value);
                if (Params["globalshowsubsites"].UserTypedIn)
                    pubweb.Navigation.GlobalIncludeSubSites = bool.Parse(Params["globalshowsubsites"].Value);

                if (Params["currentshowpages"].UserTypedIn)
                    pubweb.Navigation.CurrentIncludePages = bool.Parse(Params["currentshowpages"].Value);
                if (Params["globalshowpages"].UserTypedIn)
                    pubweb.Navigation.GlobalIncludePages = bool.Parse(Params["globalshowpages"].Value);

                OrderingMethod sortMethod = pubweb.Navigation.OrderingMethod;
                if (Params["sortmethod"].UserTypedIn)
                {
                    sortMethod = (OrderingMethod)Enum.Parse(typeof (OrderingMethod), Params["sortmethod"].Value, true);
                    pubweb.Navigation.OrderingMethod = sortMethod;
                }

                if (sortMethod != OrderingMethod.Manual)
                {
                    if (Params["autosortmethod"].UserTypedIn)
                        pubweb.Navigation.AutomaticSortingMethod = (AutomaticSortingMethod)Enum.Parse(typeof(AutomaticSortingMethod), Params["autosortmethod"].Value, true);
                    if (Params["sortascending"].UserTypedIn)
                        pubweb.Navigation.SortAscending = bool.Parse(Params["sortascending"].Value);
                }
                else
                {
                    if (Params["autosortmethod"].UserTypedIn)
                        Console.WriteLine("WARNING: parameter autosortmethod is incompatible with sortmethod {0}.  The parameter will be ignored.", sortMethod);
                    if (Params["sortascending"].UserTypedIn)
                        Console.WriteLine("WARNING: parameter sortascending is incompatible with sortmethod {0}.  The parameter will be ignored.", sortMethod);
                }

                if (Params["inheritglobalnav"].UserTypedIn)
                    pubweb.Navigation.InheritGlobal = bool.Parse(Params["inheritglobalnav"].Value);

                if (Params["currentnav"].UserTypedIn)
                {
                    CurrentNavSettingsEnum currentNav = (CurrentNavSettingsEnum)Enum.Parse(typeof (CurrentNavSettingsEnum), Params["currentnav"].Value, true);
                    if (currentNav == CurrentNavSettingsEnum.InheritParent)
                    {
                        pubweb.Navigation.InheritCurrent = true;
                        pubweb.Navigation.ShowSiblings = false;
                    }
                    else if (currentNav == CurrentNavSettingsEnum.CurrentSiteAndSiblings)
                    {
                        pubweb.Navigation.InheritCurrent = false;
                        pubweb.Navigation.ShowSiblings = true;
                    }
                    else if (currentNav == CurrentNavSettingsEnum.CurrentSiteOnly)
                    {
                        pubweb.Navigation.InheritCurrent = false;
                        pubweb.Navigation.ShowSiblings = false;
                    }
                }

                pubweb.Update();
#else
                web.Update();
#endif
            }

            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
