using System;
using System.Collections.Specialized;
using System.DirectoryServices;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.Common;
using System.Collections.Generic;
using Microsoft.SharePoint.Administration.Backup;

namespace Lapointe.SharePoint.PowerShell.Common.Farm
{
    public class BackupSites
    {
        private bool _overwrite = false;
        private string _path = null;
        private bool _includeIis = false;
        private bool _noSiteLock = false;
        private bool _useSnapshot = false;
        private const int FLAG_EXPORT_INHERITED_SETTINGS = 1;

        public BackupSites(bool overwrite, string path, bool includeIis, bool noSiteLock, bool useSnapshot)
        {
            _overwrite = overwrite;
            _path = path;
            _includeIis = includeIis;
            _noSiteLock = noSiteLock;
            _useSnapshot = useSnapshot;
        }


        public void BackupSite(SPSite site, bool prepare)
        {
            if (prepare)
                PrepareSystem();

            BackupSiteCore(site);
        }

        public void BackupSite(List<SPSite> sites, bool prepare)
        {
            if (prepare)
                PrepareSystem();

            foreach (SPSite site in sites)
            {
                BackupSiteCore(site);
            }
        }

        public void BackupSite(string url, bool prepare)
        {
            using (SPSite site = new SPSite(url))
            {
                BackupSite(site, prepare);
            }
        }

        public void BackupWebApplication(SPWebApplication webApp, bool prepare)
        {
            if (prepare)
                PrepareSystem();

            SPEnumerator enumerator = new SPEnumerator(webApp);
            InitiateEnumerator(enumerator);
        }

        public void BackupWebApplication(List<SPWebApplication> webApps, bool prepare)
        {
            if (prepare)
                PrepareSystem();

            foreach (SPWebApplication webApp in webApps)
            {
                SPEnumerator enumerator = new SPEnumerator(webApp);
                InitiateEnumerator(enumerator);
            }
        }

        public void BackupWebApplication(string url, bool prepare)
        {
            BackupWebApplication(SPWebApplication.Lookup(new Uri(url)), prepare);
        }

        public void BackupFarm(bool prepare)
        {
            BackupFarm(SPFarm.Local, prepare);
        }

        public void BackupFarm(SPFarm farm, bool prepare)
        {
            if (prepare)
                PrepareSystem();

            if (_includeIis)
            {
                // Export the IIS settings.
                string iisBakPath = Path.Combine(_path, "iis_full.bak");
                if (_overwrite && File.Exists(iisBakPath))
                    File.Delete(iisBakPath);
                if (!_overwrite && File.Exists(iisBakPath))
                    throw new SPException(
                        string.Format("The IIS backup file '{0}' already exists - specify '-overwrite' to replace the file.", iisBakPath));

                //Utilities.RunCommand("cscript", string.Format("{0}\\iiscnfg.vbs /export /f \"{1}\" /inherited /children /sp /lm", Environment.SystemDirectory, iisBakPath), false);
                using (DirectoryEntry de = new DirectoryEntry("IIS://localhost"))
                {
                    Logger.Write("Exporting full IIS settings....");
                    string decryptionPwd = string.Empty;
                    de.Invoke("Export", new object[] { decryptionPwd, iisBakPath, "/lm", FLAG_EXPORT_INHERITED_SETTINGS });
                }
            }
            SPEnumerator enumerator = new SPEnumerator(farm);
            InitiateEnumerator(enumerator);
        }

        public void PrepareSystem()
        {
            int index = 0;
            while (Directory.Exists(_path + index) && !_overwrite)
                index++;

            _path += index;

            if (!Directory.Exists(_path))
                Directory.CreateDirectory(_path);

            if (_includeIis)
            {
                // Flush any in memory changes to the file system so we can capture them on an export.
                //Utilities.RunCommand("cscript", Environment.SystemDirectory + "\\iiscnfg.vbs /save", false);

                using (DirectoryEntry de = new DirectoryEntry("IIS://localhost"))
                {
                    Logger.Write("Flushing IIS metadata to disk....");
                    de.Invoke("SaveData", new object[0]);
                    Logger.Write("IIS metadata successfully flushed to disk.");
                }
            }
        }



        private void InitiateEnumerator(SPEnumerator enumerator)
        {
            if (enumerator != null)
            {
                // Listen for web application events so that we can export the settings for the specified application.
                enumerator.SPWebApplicationEnumerated += new SPEnumerator.SPWebApplicationEnumeratedEventHandler(OnSPWebApplicationEnumerated);

                enumerator.SPSiteEnumerated += new SPEnumerator.SPSiteEnumeratedEventHandler(OnSPSiteEnumerated);
                enumerator.Enumerate();
            }
        }

        /// <summary>
        /// Handles the SPWebApplicationEnumerated event of the enumerator control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="Lapointe.SharePoint.STSADM.Commands.OperationHelpers.SPEnumerator.SPWebApplicationEventArgs"/> instance containing the event data.</param>
        private void OnSPWebApplicationEnumerated(object sender, SPEnumerator.SPWebApplicationEventArgs e)
        {
            if (!_includeIis)
                return;

            foreach (SPIisSettings iis in e.WebApplication.IisSettings.Values)
            {
                string iisBakPath = Path.Combine(_path, string.Format("iis_w3svc_{0}.bak", iis.PreferredInstanceId));
                if (_overwrite && File.Exists(iisBakPath))
                    File.Delete(iisBakPath);
                if (!_overwrite && File.Exists(iisBakPath))
                    throw new SPException(string.Format("The IIS backup file '{0}' already exists - specify '-overwrite' to replace the file.", iisBakPath));

                //Utilities.RunCommand(
                //    "cscript", 
                //    string.Format("{0}\\iiscnfg.vbs /export /f \"{1}\" /inherited /children /sp /lm/w3svc/{2}", 
                //        Environment.SystemDirectory, 
                //        iisBakPath,
                //        iis.PreferredInstanceId), 
                //    false);

                using (DirectoryEntry de = new DirectoryEntry("IIS://localhost"))
                {
                    Logger.Write("Exporting IIS settings for web application '{0}'....", iis.ServerComment);
                    string decryptionPwd = string.Empty;
                    string path = string.Format("/lm/w3svc/{0}", iis.PreferredInstanceId);
                    de.Invoke("Export", new object[] { decryptionPwd, iisBakPath, path, FLAG_EXPORT_INHERITED_SETTINGS });
                }
            }

        }

        /// <summary>
        /// Handles the SPSiteEnumerated event of the enumerator control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="Lapointe.SharePoint.STSADM.Commands.OperationHelpers.SPEnumerator.SPSiteEventArgs"/> instance containing the event data.</param>
        private void OnSPSiteEnumerated(object sender, SPEnumerator.SPSiteEventArgs e)
        {
            BackupSiteCore(e.Site);
        }

        /// <summary>
        /// Backups the site.
        /// </summary>
        /// <param name="site">The site.</param>
        private void BackupSiteCore(SPSite site)
        {
            string path = Path.Combine(_path, EncodePath(site.Url.ToString())) + ".bak";
            Logger.Write("Backing up site '{0}' to '{1}'", site.Url, path);

            if (!_useSnapshot)
            {
                bool writeLock = site.WriteLocked;
                if (!_noSiteLock)
                    site.WriteLocked = true;
                try
                {
                    site.WebApplication.Sites.Backup(site.Url, path, _overwrite);
                }
                finally
                {
                    if (!_noSiteLock)
                        site.WriteLocked = writeLock;
                }
            }
            else
            {
                SPDatabaseSnapshot snapshot = null;
                try
                {
                    snapshot = site.ContentDatabase.Snapshots.CreateSnapshot();
                    SPContentDatabase database = SPContentDatabase.CreateUnattachedContentDatabase(snapshot.ConnectionString);
                    string strSiteUrl = site.HostHeaderIsSiteName ? site.Url.ToString() : site.ServerRelativeUrl;
                    database.Sites.Backup(strSiteUrl, path, _overwrite);
                }
                finally
                {
                    if (snapshot != null)
                    {
                        snapshot.Delete();
                    }
                }
            }
        }

        /// <summary>
        /// Encodes the path.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <returns></returns>
        private static string EncodePath(string path)
        {
            Regex reg = new Regex("(?i:http://|https://)");
            path = reg.Replace(path, "");
            reg = new Regex("(?i: |\\.|:|/|%20|\\\\)");
            return reg.Replace(path, "_");
        }

    }
}
