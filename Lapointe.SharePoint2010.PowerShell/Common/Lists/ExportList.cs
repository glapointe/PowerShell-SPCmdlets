using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Collections;
using Microsoft.SharePoint.Deployment;
using Microsoft.SharePoint.Administration.Backup;
using System.IO;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    internal class ExportList
    {

        /// <summary>
        /// Performs the export.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="filename">The filename.</param>
        /// <param name="compressFile">if set to <c>true</c> [compress file].</param>
        /// <param name="haltOnFatalError">if set to <c>true</c> [halt on fatal error].</param>
        /// <param name="haltOnWarning">if set to <c>true</c> [halt on warning].</param>
        /// <param name="includeusersecurity">if set to <c>true</c> [includeusersecurity].</param>
        /// <param name="cabSize">Size of the CAB.</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <param name="versions">The versions.</param>
        /// <param name="includeDescendents">The include descendents.</param>
        /// <param name="excludeDependencies">if set to <c>true</c> [exclude dependencies].</param>
        public static void PerformExport(string url, string filename, bool compressFile, bool haltOnFatalError, bool haltOnWarning, bool includeusersecurity, int cabSize, bool logFile, bool overwrite, bool quiet, SPIncludeVersions versions, SPIncludeDescendants includeDescendents, bool excludeDependencies, bool useSqlSnapshot, bool excludeChildren)
        {
            SPExportObject exportObject = new SPExportObject();
            SPExportSettings settings = new SPExportSettings();
            settings.ExcludeDependencies = excludeDependencies;
            SPExport export = new SPExport(settings);


            exportObject.Type = SPDeploymentObjectType.List;
            exportObject.IncludeDescendants = includeDescendents;
            exportObject.ExcludeChildren = excludeChildren;
            StsAdm.OperationHelpers.ExportHelper.SetupExportObjects(settings, cabSize, compressFile, filename, haltOnFatalError, haltOnWarning, includeusersecurity, logFile, overwrite, quiet, versions);



            PerformExport(export, exportObject, settings, logFile, quiet, url, useSqlSnapshot);
        }


        /// <summary>
        /// Performs the export.
        /// </summary>
        /// <param name="export">The export.</param>
        /// <param name="exportObject">The export object.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <param name="url">The URL.</param>
        internal static void PerformExport(SPExport export, SPExportObject exportObject, SPExportSettings settings, bool logFile, bool quiet, string url, bool useSqlSnapshot)
        {
            SPDatabaseSnapshot snapshot = null;
            using (SPSite site = new SPSite(url))
            {
                ValidateUser(site);

                if (!useSqlSnapshot)
                {
                    settings.SiteUrl = site.Url;
                }
                else
                {
                    snapshot = site.ContentDatabase.Snapshots.CreateSnapshot();
                    SPContentDatabase database2 = SPContentDatabase.CreateUnattachedContentDatabase(snapshot.ConnectionString);
                    settings.UnattachedContentDatabase = database2;
                    settings.SiteUrl = database2.Sites[site.ServerRelativeUrl].Url;
                }

                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = Utilities.GetListFromViewUrl(web, url);

                    if (list == null)
                    {
                        throw new Exception("List not found.");
                    }

                    settings.SiteUrl = web.Url;
                    exportObject.Id = list.ID;
                }

                settings.ExportObjects.Add(exportObject);


                try
                {
                    export.Run();
                    if (!quiet)
                    {
                        ArrayList dataFiles = settings.DataFiles;
                        if (dataFiles != null)
                        {
                            Console.WriteLine();
                            Console.WriteLine("File(s) generated: ");
                            for (int i = 0; i < dataFiles.Count; i++)
                            {
                                Console.WriteLine("\t{0}", Path.Combine(settings.FileLocation, (string)dataFiles[i]));
                                Console.WriteLine();
                            }
                            Console.WriteLine();

                        }
                    }
                }
                finally
                {
                    if (useSqlSnapshot && (snapshot != null))
                    {
                        snapshot.Delete();
                    }

                    if (logFile)
                    {
                        Console.WriteLine();
                        Console.WriteLine("Log file generated: ");
                        Console.WriteLine("\t{0}", settings.LogFilePath);
                    }
                }
            }

        }


        /// <summary>
        /// Validates the user.
        /// </summary>
        /// <param name="site">The site.</param>
        internal static void ValidateUser(SPSite site)
        {
            if (site == null)
                throw new ArgumentNullException("site");

            bool isSiteAdmin = false;
            using (SPWeb rootWeb = site.RootWeb)
            {
                if ((rootWeb != null) && (rootWeb.CurrentUser != null))
                {
                    isSiteAdmin = rootWeb.CurrentUser.IsSiteAdmin;
                }
            }
            if (isSiteAdmin)
            {
                return;
            }
            try
            {
                isSiteAdmin = SPFarm.Local.CurrentUserIsAdministrator();
            }
            catch
            {
                isSiteAdmin = false;
            }
            if (!isSiteAdmin)
                throw new SPException(SPResource.GetString("AccessDenied", new object[0]));
        }

    }
}
