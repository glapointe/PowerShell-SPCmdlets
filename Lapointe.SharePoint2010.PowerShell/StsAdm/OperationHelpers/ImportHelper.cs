using System;
using System.IO;
using Lapointe.SharePoint.PowerShell.StsAdm.Lists;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Deployment;

namespace Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers
{
    internal class ImportHelper
    {
        private string m_webName;
        private string m_webParentUrl;

        /// <summary>
        /// Initializes a new instance of the <see cref="ImportHelper"/> class.
        /// </summary>
        public ImportHelper()
        {
            m_webName = string.Empty;
            m_webParentUrl = string.Empty;
        }

        /// <summary>
        /// Sets up the import object.
        /// </summary>
        /// <param name="settings">The settings.</param>
        /// <param name="compressFile">if set to <c>true</c> [compress file].</param>
        /// <param name="filename">The filename.</param>
        /// <param name="haltOnFatalError">if set to <c>true</c> [halt on fatal error].</param>
        /// <param name="haltOnWarning">if set to <c>true</c> [halt on warning].</param>
        /// <param name="includeusersecurity">if set to <c>true</c> [includeusersecurity].</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <param name="updateVersions">The update versions.</param>
        /// <param name="retainObjectIdentity">if set to <c>true</c> [retain object identity].</param>
        /// <param name="suppressAfterEvents">if set to <c>true</c> [suppress after events].</param>
        internal static void SetupImportObject(SPImportSettings settings, bool compressFile, string filename, bool haltOnFatalError, bool haltOnWarning, bool includeusersecurity, bool logFile, bool quiet, SPUpdateVersions updateVersions, bool retainObjectIdentity, bool suppressAfterEvents)
        {
            settings.CommandLineVerbose = !quiet;
            settings.HaltOnNonfatalError = haltOnFatalError;
            settings.HaltOnWarning = haltOnWarning;
            settings.FileCompression = compressFile;
            settings.SuppressAfterEvents = suppressAfterEvents;

            if (!compressFile)
            {
                if (string.IsNullOrEmpty(filename) || !Directory.Exists(filename))
                {
                    throw new SPException(SPResource.GetString("DirectoryNotFoundExceptionMessage", new object[] { filename }));
                }
            }
            else if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
            {
                throw new SPException(SPResource.GetString("FileNotFoundExceptionMessage", new object[] { filename }));
            }

            if (!compressFile)
            {
                settings.FileLocation = filename;
            }
            else
            {
                string path;
                Utilities.SplitPathFile(filename, out path, out filename);
                settings.FileLocation = path;
                settings.BaseFileName = filename;
            }

            if (logFile)
            {
                if (!compressFile)
                {
                    settings.LogFilePath = Path.Combine(settings.FileLocation, "import.log");
                }
                else
                    settings.LogFilePath = Path.Combine(settings.FileLocation, filename + ".import.log");
            }


            if (includeusersecurity)
            {
                settings.IncludeSecurity = SPIncludeSecurity.All;
                settings.UserInfoDateTime = SPImportUserInfoDateTimeOption.ImportAll;
            }
            settings.UpdateVersions = updateVersions;
            settings.RetainObjectIdentity = retainObjectIdentity;
        }

        /// <summary>
        /// Imports the site providing the option to retain the source objects ID values.  The source object
        /// must no longer exist in the target content database.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <param name="targeturl">The targeturl.</param>
        /// <param name="haltOnWarning">if set to <c>true</c> [halt on warning].</param>
        /// <param name="haltOnFatalError">if set to <c>true</c> [halt on fatal error].</param>
        /// <param name="noFileCompression">if set to <c>true</c> [no file compression].</param>
        /// <param name="includeUserSecurity">if set to <c>true</c> [include user security].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <param name="retainObjectIdentity">if set to <c>true</c> [retain object identity].</param>
        /// <param name="updateVersions">The update versions.</param>
        /// <param name="suppressAfterEvents">if set to <c>true</c> [suppress after events].</param>
        public void ImportSite(string filename, string targeturl, bool haltOnWarning, bool haltOnFatalError, bool noFileCompression, bool includeUserSecurity, bool quiet, bool logFile, bool retainObjectIdentity, SPUpdateVersions updateVersions, bool suppressAfterEvents)
        {
            //if (!retainObjectIdentity)
            //{
            //    // Use the built in "import" command.
            //    ImportSite(filename, targeturl, haltOnWarning, haltOnFatalError, noFileCompression, includeUserSecurity,
            //               quiet, logFile, updateVersions);
            //    return;
            //}

            SPImportSettings settings = new SPImportSettings();
            SPImport import = new SPImport(settings);

            SetupImportObject(settings, !noFileCompression, filename, haltOnFatalError, haltOnWarning, includeUserSecurity, logFile, quiet, updateVersions, retainObjectIdentity, suppressAfterEvents);

            using (SPSite site = new SPSite(targeturl))
            {
                settings.SiteUrl = site.Url;
                Common.Lists.ExportList.ValidateUser(site);

                string dirName;
                Utilities.SplitUrl(Utilities.ConvertToServiceRelUrl(Utilities.GetServerRelUrlFromFullUrl(targeturl), site.ServerRelativeUrl), out dirName, out m_webName);
                m_webParentUrl = site.ServerRelativeUrl;
                if (!string.IsNullOrEmpty(dirName))
                {
                    if (!m_webParentUrl.EndsWith("/"))
                    {
                        m_webParentUrl = m_webParentUrl + "/";
                    }
                    m_webParentUrl = m_webParentUrl + dirName;
                }
                if (m_webName == null)
                {
                    m_webName = string.Empty;
                }
            }

            EventHandler<SPDeploymentEventArgs> handler = new EventHandler<SPDeploymentEventArgs>(OnSiteImportStarted);
            import.Started += handler;

            try
            {
                import.Run();
            }
            catch (SPException ex)
            {
                if (retainObjectIdentity && ex.Message.StartsWith("The Web site address ") && ex.Message.EndsWith(" is already in use."))
                {
                    throw new SPException(
                        "You cannot import the web because the source web still exists.  Either specify the \"-deletesource\" parameter or manually delete the source web and use the exported file.", ex);

                }
                else
                    throw;
            }
            finally
            {
                Console.WriteLine();
                Console.WriteLine("Log file generated: ");
                Console.WriteLine("\t{0}", settings.LogFilePath);
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Called when [site import started].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="args">The <see cref="Microsoft.SharePoint.Deployment.SPDeploymentEventArgs"/> instance containing the event data.</param>
        private void OnSiteImportStarted(object sender, SPDeploymentEventArgs args)
        {
            SPImportObjectCollection rootObjects = args.RootObjects;
            if (rootObjects.Count != 0)
            {
                if (rootObjects.Count != 1)
                {
                    for (int i = 0; i < rootObjects.Count; i++)
                    {
                        if (rootObjects[i].Type == SPDeploymentObjectType.Web)
                        {
                            rootObjects[i].TargetParentUrl = m_webParentUrl;
                            rootObjects[i].TargetName = m_webName;
                            return;
                        }
                    }
                }
                else
                {
                    rootObjects[0].TargetParentUrl = m_webParentUrl;
                    rootObjects[0].TargetName = m_webName;
                }
            }
        }


    }
}
