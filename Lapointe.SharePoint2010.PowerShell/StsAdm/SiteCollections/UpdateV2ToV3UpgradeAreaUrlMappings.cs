using System;
using System.Collections.Specialized;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.SiteCollections
{
    // This class is currently not used as it is specific to 2007 and simply not needed for 2010; I've left it here for reference purposes only (just in case I need it in the future)
    public class UpdateV2ToV3UpgradeAreaUrlMappings : SPOperation
    {
         /// <summary>
        /// Initializes a new instance of the <see cref="UpdateV2ToV3UpgradeAreaUrlMappings"/> class.
        /// </summary>
        public UpdateV2ToV3UpgradeAreaUrlMappings()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("webapp", "w", true, null, new SPUrlValidator(), "Please specify the upgraded web application."));
            parameters.Add(new SPParam("sourceurl", "s", true, null, new SPNonEmptyValidator(), "Please specify the source url."));
            parameters.Add(new SPParam("targeturl", "t", true, null, new SPNonEmptyValidator(), "Please specify the target url."));
            Init(parameters, "\r\n\r\nUpdates the server relative URL corresponding to the source URL to reflect the new target URL in the Upgrade Area URL Mappings list: \"http://[portal]/Lists/98d3057cd9024c27b2007643c1/AllItems.aspx\".  All sites below the site are also updated.\r\n\r\nParameters:\r\n\t-webapp <web application>\r\n\t-sourceurl <source V3 url (this is not the V2 bucket URL)>\r\n\t-targeturl <target V3 URL>");
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

            

            string webApp = Params["webapp"].Value.TrimEnd('/');
            string sourceUrl = Params["sourceurl"].Value.TrimEnd('/');
            string targetUrl = Params["targeturl"].Value.TrimEnd('/');

            SPWebApplication webApplication = SPWebApplication.Lookup(new Uri(webApp));

            FixUrls(webApplication, ref sourceUrl, ref targetUrl);

            bool foundAList = false;
            foreach (SPSite site in webApplication.Sites)
            {
                try
                {
                    SPList list = GetUpgradeAreaUrlMappingsList(site);
                    if (list != null)
                    {
                        foundAList = true;
                        FixUpgradeAreaUrlMappings(list, sourceUrl, targetUrl);
                    }
                }
                finally
                {
                    site.Dispose();
                }
            }

            if (!foundAList)
            {
                output += "Source web application does not contain an Upgrade Area URL Mapping list.\r\n";
                return (int)ErrorCodes.NoError; // This isn't really an error condition - only upgraded sites will have it so simply bow out gracefully.
            }


            return (int)ErrorCodes.NoError;
        }

        /// <summary>
        /// Fixes the input urls so that they are properly formatted as server relative.
        /// </summary>
        /// <param name="webApplication">The web application.</param>
        /// <param name="sourceUrl">The source URL.</param>
        /// <param name="targetUrl">The target URL.</param>
        private static void FixUrls(SPWebApplication webApplication, ref string sourceUrl, ref string targetUrl)
        {
            SPSite sourceSite = null;
            SPSite targetSite = null;
            try
            {
                // We need to account for either an absolute or relative url being passed in and we want to be able
                // to address scenarios in which a source or target site either no longer exists or hasn't be moved yet.
                // I imagine there's probably an easier way to do this but this works.
                if (sourceUrl.StartsWith("http://") || sourceUrl.StartsWith("https://"))
                {
                    try
                    {
                        sourceSite = new SPSite(sourceUrl);
                        sourceUrl = Utilities.GetServerRelUrlFromFullUrl(sourceUrl);
                    }
                    catch (Exception)
                    {
                        sourceSite = null;
                    }
                }
                else
                {
                    using (SPSite rootSite = webApplication.Sites[0])
                    {
                        try
                        {
                            sourceSite = new SPSite(rootSite.MakeFullUrl(sourceUrl));
                        }
                        catch (Exception)
                        {
                            sourceSite = null;
                        }
                    }
                }
                if (targetUrl.StartsWith("http://") || targetUrl.StartsWith("https://"))
                {
                    try
                    {
                        targetSite = new SPSite(targetUrl);
                        targetUrl = Utilities.GetServerRelUrlFromFullUrl(targetUrl);
                    }
                    catch (Exception)
                    {
                        targetSite = null;
                    }
                }
                else
                {
                    using (SPSite rootSite = webApplication.Sites[0])
                    {
                        try
                        {
                            targetSite = new SPSite(rootSite.MakeFullUrl(targetUrl));
                        }
                        catch (Exception)
                        {
                            targetSite = null;
                        }
                    }
                }
                if (sourceSite != null)
                {
                    using (SPWeb sourceWeb = sourceSite.AllWebs[sourceUrl])
                    {
                        if (sourceWeb.Exists)
                            sourceUrl = sourceWeb.ServerRelativeUrl;
                    }
                }
                if (targetSite != null)
                {
                    using (SPWeb targetWeb = targetSite.AllWebs[targetUrl])
                    {
                        if (targetWeb.Exists)
                        {
                            targetUrl = targetWeb.ServerRelativeUrl;
                        }
                    }
                }
                if (sourceSite != null && targetSite != null)
                {
                    if (sourceSite.WebApplication.Id != targetSite.WebApplication.Id)
                        throw new Exception(
                            "Source and target web applications must be the same (spsredirect.aspx does not work across web applications).");
                    if (sourceUrl.ToLowerInvariant() == targetUrl.ToLowerInvariant())
                        throw new Exception("Source web and target web are the same.");
                }
                
            }
            finally
            {
                if (sourceSite != null)
                {
                    sourceSite.Dispose();
                }
                if (targetSite != null)
                {
                    targetSite.Dispose();
                }
            }
        }

        #endregion


        /// <summary>
        /// Fixes the upgrade area URL mappings.  This overload is used when moving a web between site collections on a single web application.
        /// Currently does not support going between web applications due to the possibility of changing the wrong urls.
        /// </summary>
        /// <param name="webApp">The web app.</param>
        /// <param name="sourceWebUrl">The source web URL.</param>
        /// <param name="targetWebUrl">The target web URL.</param>
        internal static void FixUpgradeAreaUrlMappings(SPWebApplication webApp, string sourceWebUrl, string targetWebUrl)
        {
            foreach (SPSite site in webApp.Sites)
            {
                try
                {
                    FixUpgradeAreaUrlMappings(site, sourceWebUrl, targetWebUrl);
                }
                finally
                {
                    site.Dispose();
                }
            }
        }


        /// <summary>
        /// Fixes the upgraded area URL mappings.  This overload is used when moving a web within a single site collection.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="sourceWebUrl">The source web URL.</param>
        /// <param name="targetWebUrl">The target web URL.</param>
        private static void FixUpgradeAreaUrlMappings(SPSite site, string sourceWebUrl, string targetWebUrl)
        {
            SPList list = GetUpgradeAreaUrlMappingsList(site);

            if (list == null)
                return;

            FixUpgradeAreaUrlMappings(list, sourceWebUrl, targetWebUrl);
        }

        /// <summary>
        /// Gets the upgrade area URL mappings list.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <returns></returns>
        internal static SPList GetUpgradeAreaUrlMappingsList(SPSite site)
        {
            try
            {
                // 98d3057cd9024c27b2007643c1 is a special hard coded name for a list that Microsoft uses to store the mapping
                // of URLs from v2 to v3 (maps the bucket urls to the new urls).
                using (SPWeb rootWeb = site.RootWeb)
                {
                    return Utilities.GetListByUrl(rootWeb, UPGRADE_AREA_URL_LIST, false);
                }
            }
            catch (Exception)
            {
                return null;
            }

        }
        /// <summary>
        /// Fixes the upgraded Area URL mappings.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="sourceWebUrl">The source web URL.</param>
        /// <param name="targetWebUrl">The target web URL.</param>
        private static void FixUpgradeAreaUrlMappings(SPList list, string sourceWebUrl, string targetWebUrl)
        {
            targetWebUrl = targetWebUrl.TrimEnd('/');
            sourceWebUrl = sourceWebUrl.TrimEnd('/');

            using (SPSite rootSite = list.ParentWeb.Site)
            {
                foreach (SPListItem item in list.Items)
                {
                    string v3Url = item["V3ServerRelativeUrl"] as string;
                    if (string.IsNullOrEmpty(v3Url))
                        continue;

                    if (v3Url.ToLowerInvariant().StartsWith(sourceWebUrl.ToLowerInvariant()))
                    {
                        v3Url = targetWebUrl + v3Url.Substring(sourceWebUrl.Length);
                        item["V3ServerRelativeUrl"] = v3Url;


                        SPSite targetSite;
                        try
                        {
                            targetSite = new SPSite(rootSite.MakeFullUrl(v3Url));
                        }
                        catch (Exception)
                        {
                            targetSite = null;
                        }
                        try
                        {
                            if (targetSite != null)
                            {
                                using (SPWeb targetWeb = targetSite.AllWebs[v3Url])
                                {
                                    if (targetWeb.Exists)
                                    {
                                        item["V3WebId"] = targetWeb.ID.ToString();
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (targetSite != null)
                                targetSite.Dispose();
                        }
                        item.Update();
                    }
                }
            }
        }
    }
}
