using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Deployment;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.StsAdmin;
using Microsoft.SharePoint.Utilities;

namespace Lapointe.SharePoint.PowerShell.Common.SiteCollections
{
    public class MoveWeb
    {
        public static void MoveWebWithinSite(SPWeb sourceWeb, SPWeb parentWeb, string urlName)
        {
            // All we're doing is changing the URL name of the web.
            if (sourceWeb.IsRootWeb)
                throw new SPCmdletException("Cannot change the URL name of the root of the Site Collection.");
            sourceWeb.ServerRelativeUrl = SPUrlUtility.CombineUrl(parentWeb.ServerRelativeUrl, urlName);
            sourceWeb.Update();
        }

        public static void MoveWebOutsideSite(SPWeb sourceWeb, SPWeb parentWeb, string newUrlName, bool retainObjectIdentity, bool haltOnWarning, bool haltOnFatalError, bool includeUserSecurity, bool suppressAfterEvents, string tempPath)
        {
            string webName = sourceWeb.Name;
            if (string.IsNullOrEmpty(webName))
                webName = sourceWeb.ServerRelativeUrl.Substring(sourceWeb.ServerRelativeUrl.LastIndexOf('/') + 1);
            if (string.IsNullOrEmpty(webName))
                webName = sourceWeb.ID.ToString();

            string originalWebName = webName;
            string tempWebName = Guid.NewGuid().ToString();
            bool wasMoved = false;
            if (retainObjectIdentity && sourceWeb.ParentWeb != null && !sourceWeb.ParentWeb.IsRootWeb)
            {
                Console.WriteLine("Moving source web to child of root web...");

                sourceWeb.ServerRelativeUrl = Utilities.ConcatServerRelativeUrls(sourceWeb.Site.ServerRelativeUrl, tempWebName);
                sourceWeb.Update();
                wasMoved = true;

                Console.WriteLine("Source web moved to: " + sourceWeb.ServerRelativeUrl + "\r\n");
            }

            Console.WriteLine("Exporting Web...");
            if (string.IsNullOrEmpty(tempPath))
                tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            else
                tempPath = Path.Combine(tempPath, Guid.NewGuid().ToString());
            string fileName = ExportHelper.ExportSite(sourceWeb.Site.MakeFullUrl(sourceWeb.ServerRelativeUrl),
                                                      haltOnWarning, haltOnFatalError, true, includeUserSecurity, true, SPIncludeVersions.All, 0, false, true, tempPath);
            Console.WriteLine("Web exported.\r\n");

            string newWebUrl = SPUrlUtility.CombineUrl(parentWeb.Url, webName);
            if (!retainObjectIdentity)
            {
                Console.WriteLine("Creating web for import...");
                CreateWeb(newWebUrl, sourceWeb);
                Console.WriteLine("Web created.\r\n");
            }

            // Need to delete the web before importing if retaining the object identity.
            if (retainObjectIdentity)
                DeleteWeb(sourceWeb);

            Console.WriteLine("Importing web...");
            ImportHelper import = new ImportHelper();

            import.ImportSite(fileName, newWebUrl, haltOnWarning, haltOnFatalError, true,
                              includeUserSecurity, true, true, retainObjectIdentity, SPUpdateVersions.Append, suppressAfterEvents);
            Console.WriteLine("Web imported.\r\n");

            if (retainObjectIdentity && wasMoved)
            {
                Console.WriteLine("Renaming imported web to original name...");

                // We first need to make sure that a web with the same name doesn't already exist.
                // If it does then simply add an index to the name of the web.
                int i = 0;
                string suffix = "";
                if (!string.IsNullOrEmpty(newUrlName))
                {
                    originalWebName = newUrlName;
                }
                while (true)
                {
                    i++;
                    try
                    {
                        using (SPWeb tempWeb = parentWeb.Site.AllWebs[Utilities.ConcatServerRelativeUrls(parentWeb.ServerRelativeUrl, originalWebName + suffix)])
                        {
                            if (tempWeb.Exists)
                            {
                                suffix = i.ToString();
                                continue;
                            }
                            break;
                        }
                    }
                    catch (Exception)
                    {
                        suffix = i.ToString();
                    }
                    break;
                }
                using (SPWeb targetWeb = parentWeb.Site.AllWebs[Utilities.ConcatServerRelativeUrls(parentWeb.ServerRelativeUrl, tempWebName)])
                {
                    originalWebName = originalWebName + suffix;
                    targetWeb.ServerRelativeUrl = Utilities.ConcatServerRelativeUrls(parentWeb.ServerRelativeUrl, originalWebName);
                    targetWeb.Update();

                    Console.WriteLine("Imported web renamed: " + targetWeb.ServerRelativeUrl + "\r\n");
                }

            }

            string newWebServerRelativeUrl = SPUrlUtility.CombineUrl(parentWeb.ServerRelativeUrl, originalWebName);
            if (!retainObjectIdentity)
            {
                Console.WriteLine("Retargetting web parts...");
                using (SPWeb targetWeb = parentWeb.Site.AllWebs[newWebServerRelativeUrl])
                {
                    RepairSiteCollectionImportedFromSubSite.RetargetMiscWebParts(sourceWeb.Site, sourceWeb, targetWeb);
                }
                Console.WriteLine("Retargetting of web parts complete.");
            }

            // Need to delete the web after repairing if not retaining the object identity.
            if (!retainObjectIdentity)
                DeleteWeb(sourceWeb);

            Directory.Delete(fileName, true);
        }

        /// <summary>
        /// Deletes the web.
        /// </summary>
        /// <param name="sourceWeb">The source web.</param>
        private static void DeleteWeb(SPWeb sourceWeb)
        {
            Console.WriteLine("Deleting source web...");
            if (!sourceWeb.IsRootWeb)
            {
                DeleteSubWebs(sourceWeb.Webs);

                sourceWeb.Delete();
                Console.WriteLine("Source web deleted.\r\n");
            }
            else
            {
                sourceWeb.Site.Delete();
            }
        }

        /// <summary>
        /// Deletes the sub webs.
        /// </summary>
        /// <param name="webs">The webs.</param>
        private static void DeleteSubWebs(SPWebCollection webs)
        {
            foreach (SPWeb web in webs)
            {
                if (web.Webs.Count > 0)
                    DeleteSubWebs(web.Webs);

                web.Delete();
            }
        }

        /// <summary>
        /// Creates the web.
        /// </summary>
        /// <param name="fullWebUrl">The full url of the web to create.</param>
        /// <param name="sourceWeb">The source web.</param>
        private static void CreateWeb(string fullWebUrl, SPWeb sourceWeb)
        {
            string strWebTemplate = null;
            string strTitle = null;
            string strDescription = null;
            uint nLCID = sourceWeb.Language;
            bool convert = false;
            bool useUniquePermissions = false;

            using (SPSiteAdministration administration = new SPSiteAdministration(fullWebUrl))
                administration.AddWeb(Utilities.GetServerRelUrlFromFullUrl(fullWebUrl), strTitle, strDescription, nLCID, strWebTemplate, useUniquePermissions, convert);

        }

    }
}
