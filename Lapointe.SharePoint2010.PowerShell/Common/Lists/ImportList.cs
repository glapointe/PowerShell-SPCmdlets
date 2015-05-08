using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Deployment;
using System.IO;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    internal class ImportList
    {
        protected string m_targetUrl;
        protected string m_sourceUrl;
        protected bool m_retargetLinks = false;
        protected SPList m_targetList = null;
        protected SPSite m_targetSite = null;
        protected SPWeb m_targetWeb = null;

        public ImportList(SPList sourceList, SPWeb targetWeb, bool retargetLinks)
        {
            m_targetUrl = targetWeb.Url;
            m_sourceUrl = sourceList.ParentWeb.Site.MakeFullUrl(sourceList.RootFolder.ServerRelativeUrl);
            m_retargetLinks = retargetLinks;
        }

        public ImportList(string sourceUrl, string targetUrl, bool retargetLinks)
        {
            m_sourceUrl = sourceUrl;
            m_targetUrl = targetUrl;
            m_retargetLinks = retargetLinks;
        }

        public void PerformImport(bool compressFile, string filename, bool quiet, bool haltOnWarning, bool haltOnFatalError, bool includeusersecurity, bool logFile, bool retainObjectIdentity, bool copySecurity, bool suppressAfterEvents, SPUpdateVersions updateVersions)
        {
            SPImportSettings settings = new SPImportSettings();
            SPImport import = new SPImport(settings);

            StsAdm.OperationHelpers.ImportHelper.SetupImportObject(settings, compressFile, filename, haltOnFatalError, haltOnWarning, includeusersecurity, logFile, quiet, updateVersions, retainObjectIdentity, suppressAfterEvents);

            try
            {
                m_targetSite = new SPSite(m_targetUrl);
                m_targetWeb = m_targetSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(m_targetUrl)];

                PerformImport(import, settings, logFile, m_targetUrl);

                // If the list is a discussion list then attempt to resolve flattened threads.
                //if (m_targetList != null)
                //    SiteCollectionSettings.RepairSiteCollectionImportedFromSubSite.RepairDiscussionList(m_targetSite, m_targetList);

                if (includeusersecurity && !string.IsNullOrEmpty(m_sourceUrl) && copySecurity)
                {
                    using (SPSite sourceSite = new SPSite(m_sourceUrl))
                    using (SPWeb sourceWeb = sourceSite.OpenWeb())
                    {
                        SPList sourceList = Utilities.GetListFromViewUrl(sourceWeb, m_sourceUrl);

                        if (sourceList != null)
                        {
                            if (m_targetList != null)
                                Common.Lists.CopyListSecurity.CopySecurity(sourceList, m_targetList, m_targetWeb, true, quiet);
                        }
                    }
                }
            }
            finally
            {
                if (m_targetSite != null)
                    m_targetSite.Dispose();
                if (m_targetWeb != null)
                    m_targetWeb.Dispose();
            }
        }

        /// <summary>
        /// Performs the import.
        /// </summary>
        /// <param name="import">The import.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <param name="targetUrl">The Source URL.</param>
        internal void PerformImport(SPImport import, SPImportSettings settings, bool logFile, string targetUrl)
        {
            using (SPSite site = new SPSite(targetUrl))
            {
                ExportList.ValidateUser(site);

                using (SPWeb web = site.OpenWeb())
                {
                    settings.SiteUrl = site.Url;
                    settings.WebUrl = web.Url;
                }
            }
            import.ObjectImported += new EventHandler<SPObjectImportedEventArgs>(OnImported);

            try
            {
                import.Run();
            }
            finally
            {
                if (logFile)
                {
                    Console.WriteLine();
                    Console.WriteLine("Log file generated: ");
                    Console.WriteLine("\t{0}", settings.LogFilePath);
                    Console.WriteLine();
                }
            }
        }


        /// <summary>
        /// Called when [imported].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="eventArgs">The <see cref="Microsoft.SharePoint.Deployment.SPObjectImportedEventArgs"/> instance containing the event data.</param>
        public void OnImported(object sender, SPObjectImportedEventArgs eventArgs)
        {
            if (m_targetList == null && eventArgs.Type == SPDeploymentObjectType.List)
            {
                // Get the list for later processing.
                m_targetList = m_targetWeb.GetList(eventArgs.TargetUrl);
            }

            if (!m_retargetLinks)
                return;

            if (eventArgs.Type != SPDeploymentObjectType.ListItem)
                return;

            SPImport import = sender as SPImport;
            if (import == null)
                return;

            try
            {
                string url = eventArgs.SourceUrl; // This is not fully qualified so we need the user specified url for the site.
                using (SPSite site = new SPSite(m_sourceUrl))
                using (SPWeb web = site.OpenWeb(url, false))
                {
                    string targetUrl = m_targetSite.MakeFullUrl(eventArgs.TargetUrl);
                    SPListItem li = web.GetListItem(url);
                    int count = li.BackwardLinks.Count;
                    for (int i = count - 1; i >= 0; i--)
                    {
                        SPLink link = li.BackwardLinks[i];
                        using (SPWeb rweb = site.OpenWeb(link.ServerRelativeUrl, false))
                        {
                            object o = rweb.GetObject(link.ServerRelativeUrl);
                            if (o is SPFile)
                            {
                                SPFile f = o as SPFile;
                                f.ReplaceLink(eventArgs.SourceUrl, targetUrl);
                            }
                            if (o is SPListItem)
                            {
                                SPListItem l = o as SPListItem;
                                l.ReplaceLink(eventArgs.SourceUrl, targetUrl);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Links could not be retargeted for " + eventArgs.SourceUrl + "\r\n" + ex.Message);
            }
        }

        /// <summary>
        /// Copies the specified source list to the target URL.
        /// </summary>
        /// <param name="directory">The directory.</param>
        /// <param name="compressFile">if set to <c>true</c> [compress file].</param>
        /// <param name="includeusersecurity">if set to <c>true</c> [includeusersecurity].</param>
        /// <param name="excludeDependencies">if set to <c>true</c> [exclude dependencies].</param>
        /// <param name="haltOnFatalError">if set to <c>true</c> [halt on fatal error].</param>
        /// <param name="haltOnWarning">if set to <c>true</c> [halt on warning].</param>
        /// <param name="versions">The versions.</param>
        /// <param name="updateVersions">The update versions.</param>
        /// <param name="suppressAfterEvents">if set to <c>true</c> [suppress after events].</param>
        /// <param name="copySecurity">if set to <c>true</c> [copy security].</param>
        /// <param name="deleteSource">if set to <c>true</c> [delete source].</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        internal void Copy(string directory, bool compressFile, int cabSize, bool includeusersecurity, bool excludeDependencies, bool haltOnFatalError, bool haltOnWarning, SPIncludeVersions versions, SPUpdateVersions updateVersions, bool suppressAfterEvents, bool copySecurity, bool deleteSource, bool logFile, bool quiet, SPIncludeDescendants includeDescendents, bool useSqlSnapshot, bool excludeChildren, bool retainObjectIdentity)
        {
            if (string.IsNullOrEmpty(directory))
                directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            string filename = directory;
            if (compressFile)
                filename = Path.Combine(directory, "temp.cmp");

            SPExportObject exportObject = new SPExportObject();
            SPExportSettings exportSettings = new SPExportSettings();
            exportSettings.ExcludeDependencies = excludeDependencies;
            SPExport export = new SPExport(exportSettings);

            exportObject.Type = SPDeploymentObjectType.List;
            exportObject.IncludeDescendants = includeDescendents;
            exportObject.ExcludeChildren = excludeChildren;
            StsAdm.OperationHelpers.ExportHelper.SetupExportObjects(exportSettings, cabSize, compressFile, filename, haltOnFatalError, haltOnWarning, includeusersecurity, logFile, true, quiet, versions);

            ExportList.PerformExport(export, exportObject, exportSettings, logFile, quiet, m_sourceUrl, useSqlSnapshot);

            SPImportSettings importSettings = new SPImportSettings();
            SPImport import = new SPImport(importSettings);

            StsAdm.OperationHelpers.ImportHelper.SetupImportObject(importSettings, compressFile, filename, haltOnFatalError, haltOnWarning, includeusersecurity, logFile, quiet, updateVersions, retainObjectIdentity, suppressAfterEvents);

            try
            {
                m_targetSite = new SPSite(m_targetUrl);
                m_targetWeb = m_targetSite.AllWebs[Utilities.GetServerRelUrlFromFullUrl(m_targetUrl)];

                PerformImport(import, importSettings, logFile, m_targetUrl);

                // If the list is a discussion list then attempt to resolve flattened threads.
                //if (m_targetList != null)
                //    SiteCollectionSettings.RepairSiteCollectionImportedFromSubSite.RepairDiscussionList(m_targetSite, m_targetList);

                if (!logFile && !deleteSource)
                {
                    Directory.Delete(directory, true);
                }
                else if (logFile && !deleteSource)
                {
                    foreach (string s in Directory.GetFiles(directory))
                    {
                        FileInfo file = new FileInfo(s);
                        if (file.Extension == ".log")
                            continue;
                        file.Delete();
                    }
                }

                if (deleteSource || copySecurity)
                {
                    using (SPSite sourceSite = new SPSite(m_sourceUrl))
                    using (SPWeb sourceWeb = sourceSite.OpenWeb())
                    {
                        SPList sourceList = Utilities.GetListFromViewUrl(sourceWeb, m_sourceUrl);

                        if (sourceList != null)
                        {
                            // If the user has chosen to include security then assume they mean for all the settings to match
                            // the source - copy those settings using the CopyListSecurity operation.
                            if (copySecurity)
                            {
                                Common.Lists.CopyListSecurity.CopySecurity(sourceList, m_targetList, m_targetWeb, true, quiet);
                            }

                            // If the user wants the source deleted (move operation) then delete using the DeleteList operation.
                            if (deleteSource)
                            {
                                DeleteList.Delete(sourceList, true);
                                Console.WriteLine("Source list deleted.  You can find the exported list here: " +
                                                  directory);
                            }
                        }
                    }
                }
            }
            finally
            {
                if (m_targetSite != null)
                    m_targetSite.Dispose();
                if (m_targetWeb != null)
                    m_targetWeb.Dispose();
            }
        }
    }
}
