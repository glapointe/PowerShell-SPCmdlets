using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Threading;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Deployment;
#if MOSS
using Microsoft.SharePoint.Publishing;
#endif
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;

namespace Lapointe.SharePoint.PowerShell.Common.SiteCollections
{
    internal class ConvertSubSiteToSiteCollection
    {

        internal static SPSite ConvertWebToSite(string sourceurl, string targeturl, SPSiteSubscription siteSubscription, bool suppressAfterEvents, 
            bool noFileCompression, string exportedFile, bool createSiteInDB, string databaseName, bool createManagedPath, 
            bool haltOnWarning, bool haltOnFatalError, bool deleteSource, string title, string description, uint nLCID, string ownerName, 
            string ownerEmail, string ownerLogin, string secondaryContactName, string secondaryContactLogin, string secondaryContactEmail, 
            string quota, bool useHostHeaderAsSiteName)
        {

            try
            {
                if (!string.IsNullOrEmpty(exportedFile))
                {
                    if (noFileCompression)
                    {
                        if (!Directory.Exists(exportedFile))
                            throw new SPException(SPResource.GetString("DirectoryNotFoundExceptionMessage", new object[] { exportedFile }));
                    }
                    else
                    {
                        if (!File.Exists(exportedFile))
                            throw new SPException(SPResource.GetString("FileNotFoundExceptionMessage", new object[] { exportedFile }));
                    }
                }

                if (createSiteInDB && string.IsNullOrEmpty(databaseName))
                {
                    throw new SPSyntaxException("databasename is required if creating the site in a new or existing database.");
                }


                using (SPSite sourceSite = new SPSite(sourceurl))
                using (SPWeb sourceWeb = sourceSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(sourceurl)])
                {
                    if (createManagedPath)
                    {
                        Logger.Write("Adding managed path...");
                        AddManagedPath(targeturl, haltOnWarning);
                        Logger.Write("Managed path added.\r\n");
                    }

                    if (string.IsNullOrEmpty(exportedFile))
                    {
                        Logger.Write("Exporting site...");
                        exportedFile = ExportHelper.ExportSite(sourceurl,
                                              haltOnWarning,
                                              haltOnFatalError,
                                              noFileCompression,
                                              true, !Logger.Verbose, SPIncludeVersions.All, 0, false, true);
                        Logger.Write("Site exported.\r\n");
                    }

                    Logger.Write("Creating site for import...");

                    SPWebApplicationPipeBind webAppBind = new SPWebApplicationPipeBind(targeturl);
                    SPWebApplication webApp = webAppBind.Read(false);


                    SPSite targetSite = CreateSite(webApp, databaseName, siteSubscription, null, title, description, ownerName, ownerEmail, quota, secondaryContactName,
                        secondaryContactEmail, useHostHeaderAsSiteName, nLCID, new Uri(targeturl), ownerLogin, secondaryContactLogin);

                    SPWeb targetWeb = targetSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(targeturl)];
                    Logger.Write("Site created.\r\n");

                    Logger.Write("Importing site...");

                    ImportHelper import = new ImportHelper();
                    import.ImportSite(exportedFile,
                                      targeturl,
                                      haltOnWarning,
                                      haltOnFatalError,
                                      noFileCompression, true, !Logger.Verbose, true,
                                      false, SPUpdateVersions.Append, suppressAfterEvents);

                    Logger.Write("Site imported.\r\n");


#if MOSS
                    Logger.Write("Repairing imported site...");
                    // Need to add a small delay here as the repair seems to fail occassionally due to a timing issue.
                    Common.TimerJobs.ExecAdmSvcJobs.Execute(false, true);
                    int tryCount = 0;
                    while (true)
                    {
                        try
                        {
                            tryCount++;
                            Common.SiteCollections.RepairSiteCollectionImportedFromSubSite.RepairSite(sourceurl, targeturl);
                            break;
                        }
                        catch (InvalidPublishingWebException)
                        {
                            if (haltOnWarning)
                                throw;
                            if (tryCount > 3)
                            {
                                Logger.WriteWarning("Repair of site collection failed - unable to get Pages library.  Manually run 'repairsitecollectionimportedfromsubsite' command to try again.");
                                break;
                            }
                            else
                                Thread.Sleep(10000);
                        }
                        catch (OutOfMemoryException)
                        {
                            if (haltOnWarning)
                                throw;

                            Logger.WriteWarning("Server ran out of memory and was not able to complete the repair operation.  Manually run 'repairsitecollectionimportedfromsubsite' command to try again.");
                            break;
                        }
                    }
                    Logger.Write("Imported site repaired.\r\n");
#endif

                    // Upgrade any redirect links if present.
                    if (sourceSite.WebApplication.Id == targetSite.WebApplication.Id)
                    {
                        //Console.WriteLine("Repairing Area upgrade URLs...");
                        // The spsredirect.aspx page which uses the upgrade list only supports server relative urls so
                        // if we change the link to be absolute the redirect won't work so don't bother changing it.
                        // Note that we could get around this if we had a simple redirect page that could take in
                        // a target url - then we could have spsredirect.aspx load the server relative redirect
                        // page which would have the actual url passed into it (unfortunately redirect.aspx does
                        // some funky stuff so we can't use it).
                        //UpdateV2ToV3UpgradeAreaUrlMappings.FixUpgradeAreaUrlMappings(sourceSite.WebApplication,
                        //                                                             sourceWeb.ServerRelativeUrl,
                        //                                                             targetWeb.ServerRelativeUrl);
                        //Console.WriteLine("Area upgrade URLs repaired.\r\n");
                    }

                    if (deleteSource)
                    {
                        Logger.Write("Deleting source web...");
                        if (!sourceWeb.IsRootWeb)
                        {
                            DeleteSubWebs(sourceWeb.Webs);

                            sourceWeb.Delete();
                            Logger.Write("Source web deleted.\r\n");
                            Logger.Write("You can find the exported web at " + exportedFile + "\r\n");
                        }
                        else
                            Logger.Write("Source web is a root web - cannot delete.");
                    }
                    return targetSite;
                }
            }
            catch (Exception ex)
            {
                Logger.WriteException(new System.Management.Automation.ErrorRecord(ex, null, System.Management.Automation.ErrorCategory.NotSpecified, null));
            }
            return null;
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
        /// Adds the managed path.
        /// </summary>
        /// <param name="targetUrl">The target URL.</param>
        /// <param name="haltOnWarning">if set to <c>true</c> [halt on warning].</param>
        private static void AddManagedPath(string targetUrl, bool haltOnWarning)
        {
            string serverRelUrlFromFullUrl = Utilities.GetServerRelUrlFromFullUrl(targetUrl);
            serverRelUrlFromFullUrl = serverRelUrlFromFullUrl.Trim(new char[] { '/', '*' });

            SPWebApplication webApp = SPWebApplication.Lookup(new Uri(targetUrl));
            SPPrefixCollection prefixCollection = webApp.Prefixes;
            if (prefixCollection.Contains(serverRelUrlFromFullUrl))
            {
                if (haltOnWarning)
                    throw new SPException("Managed path already exists.");
                else 
                    Logger.WriteWarning("Managed path already exists.");
                return;
            }
            prefixCollection.Add(serverRelUrlFromFullUrl, SPPrefixType.ExplicitInclusion);
        }

        private static SPSite CreateSite(SPWebApplication webApp, string dbname, SPSiteSubscription siteSubscription, string webTemplate, string title, string description, string ownerName, string ownerEmail, string quota, string secondaryContactName, string secondaryContactEmail, bool useHostHeaderAsSiteName, uint nLCID, Uri uri, string ownerLogin, string secondaryContactLogin)
        {

            Logger.Write("PROGRESS: Getting content database...");
            SPContentDatabase database = null;
            if (!string.IsNullOrEmpty(dbname))
            {
                foreach (SPContentDatabase tempDB in webApp.ContentDatabases)
                {
                    if (tempDB.Name.ToLower() == dbname.ToLower())
                    {
                        database = tempDB;
                        break;
                    }
                }
                if (database == null)
                    throw new SPException("Content database not found.");
            }

            SPSite site = CreateSite(webApp, siteSubscription, webTemplate, title, description, ownerName, ownerEmail, quota, secondaryContactName, secondaryContactEmail, useHostHeaderAsSiteName, nLCID, uri, ownerLogin, secondaryContactLogin, database);

            if (useHostHeaderAsSiteName && !webApp.IisSettings[SPUrlZone.Default].DisableKerberos)
            {
                Logger.Write(SPResource.GetString("WarnNoDefaultNTLM", new object[0]));
            }

            return site;
        }

        private static SPSite CreateSite(SPWebApplication webApp, SPSiteSubscription siteSubscription, string webTemplate, string title, string description, string ownerName, string ownerEmail, string quota, string secondaryContactName, string secondaryContactEmail, bool useHostHeaderAsSiteName, uint nLCID, Uri uri, string ownerLogin, string secondaryContactLogin, SPContentDatabase database)
        {

            if (database != null && database.MaximumSiteCount <= database.CurrentSiteCount)
                throw new SPException("The maximum site count for the specified database has been exceeded.  Increase the maximum site count or specify another database.");

            Logger.Write("PROGRESS: Creating site collection...");
            SPSite site = null;
            if (database != null)
            {
                site = database.Sites.Add(siteSubscription, uri.OriginalString, title, description, nLCID, webTemplate, ownerLogin,
                     ownerName, ownerEmail, secondaryContactLogin, secondaryContactName, secondaryContactEmail,
                     useHostHeaderAsSiteName);
            }
            else
            {
                site = webApp.Sites.Add(siteSubscription, uri.OriginalString, title, description, nLCID, webTemplate, ownerLogin,
                     ownerName, ownerEmail, secondaryContactLogin, secondaryContactName, secondaryContactEmail,
                     useHostHeaderAsSiteName);
            }
            Logger.Write("PROGRESS: Site collection successfully created.");

            if (!string.IsNullOrEmpty(quota))
            {
                Logger.Write("PROGRESS: Associating quota template with site collection...");
                using (SPSiteAdministration administration = new SPSiteAdministration(site.Url))
                {
                    SPFarm farm = SPFarm.Local;
                    SPWebService webService = farm.Services.GetValue<SPWebService>("");

                    SPQuotaTemplateCollection quotaColl = webService.QuotaTemplates;
                    administration.Quota = quotaColl[quota];
                }
            }
            if (!string.IsNullOrEmpty(webTemplate))
            {
                Logger.Write("PROGRESS: Creating default security groups...");
                using (SPWeb web = site.RootWeb)
                {
                    web.CreateDefaultAssociatedGroups(ownerLogin, secondaryContactLogin, string.Empty);
                }
            }

            return site;
        }
    }
}
