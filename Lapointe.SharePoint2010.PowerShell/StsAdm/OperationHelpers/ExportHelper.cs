using System;
using System.IO;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Deployment;
using Microsoft.SharePoint.Administration.Backup;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers
{
    internal class ExportHelper
    {
        /// <summary>
        /// Sets up the export objects.
        /// </summary>
        /// <param name="settings">The settings.</param>
        /// <param name="cabSize">Size of the CAB.</param>
        /// <param name="compressFile">if set to <c>true</c> [compress file].</param>
        /// <param name="filename">The filename.</param>
        /// <param name="haltOnFatalError">if set to <c>true</c> [halt on fatal error].</param>
        /// <param name="haltOnWarning">if set to <c>true</c> [halt on warning].</param>
        /// <param name="includeUserSecurity">if set to <c>true</c> [include user security].</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <param name="versions">The versions.</param>
        internal static void SetupExportObjects(SPExportSettings settings, int cabSize, bool compressFile, string filename, bool haltOnFatalError, bool haltOnWarning, bool includeUserSecurity, bool logFile, bool overwrite, bool quiet, SPIncludeVersions versions)
        {
            if (compressFile)
            {
                new SPDirectoryExistsAndValidFileNameValidator().Validate(filename);
            }
            if (!overwrite && File.Exists(filename))
                throw new SPException(SPResource.GetString("NotOverwriteExportError", new object[] { filename }));

            settings.ExportMethod = SPExportMethodType.ExportAll;
            settings.HaltOnNonfatalError = haltOnFatalError;
            settings.HaltOnWarning = haltOnWarning;
            settings.CommandLineVerbose = !quiet;
            settings.IncludeVersions = versions;
            settings.IncludeSecurity = includeUserSecurity ? SPIncludeSecurity.All : SPIncludeSecurity.None;
            settings.OverwriteExistingDataFile = overwrite;
            settings.FileMaxSize = cabSize;
            settings.FileCompression = compressFile;

            settings.FileLocation = filename;
            if (!compressFile)
            {
                settings.FileLocation = filename;
            }
            else
            {
                string fileLocation;
                Utilities.SplitPathFile(filename, out fileLocation, out filename);
                settings.FileLocation = fileLocation;
                settings.BaseFileName = filename;
            }

            if (logFile)
            {
                if (!compressFile)
                {
                    settings.LogFilePath = Path.Combine(settings.FileLocation, "export.log");
                }
                else
                {
                    settings.LogFilePath = Path.Combine(settings.FileLocation, filename + ".export.log");
                }
                bool fileExists = File.Exists(settings.LogFilePath);
                if (!overwrite && fileExists)
                {
                    throw new SPException(SPResource.GetString("DataFileExists", new object[] { settings.LogFilePath }));
                }
                if (fileExists)
                {
                    File.Delete(settings.LogFilePath);
                }
            }
        }

        /// <summary>
        /// Exports the site.
        /// </summary>
        /// <param name="sourceurl">The sourceurl.</param>
        /// <param name="haltOnWarning">if set to <c>true</c> [halt on warning].</param>
        /// <param name="haltOnFatalError">if set to <c>true</c> [halt on fatal error].</param>
        /// <param name="noFileCompression">if set to <c>true</c> [no file compression].</param>
        /// <param name="includeUserSecurity">if set to <c>true</c> [include user security].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <param name="versions">The versions.</param>
        /// <param name="cabSize">Size of the CAB.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <returns></returns>
        public static string ExportSite(string sourceurl,
            bool haltOnWarning,
            bool haltOnFatalError,
            bool noFileCompression,
            bool includeUserSecurity,
            bool quiet,
            SPIncludeVersions versions,
            int cabSize,
            bool overwrite,
            bool logFile)
        {
            string filename = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());

            return
                ExportSite(sourceurl, haltOnWarning, haltOnFatalError, noFileCompression, includeUserSecurity, quiet,
                           versions, cabSize, overwrite, logFile, filename);
        }

        /// <summary>
        /// Exports the site.
        /// </summary>
        /// <param name="sourceurl">The sourceurl.</param>
        /// <param name="haltOnWarning">if set to <c>true</c> [halt on warning].</param>
        /// <param name="haltOnFatalError">if set to <c>true</c> [halt on fatal error].</param>
        /// <param name="noFileCompression">if set to <c>true</c> [no file compression].</param>
        /// <param name="includeUserSecurity">if set to <c>true</c> [include user security].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <param name="versions">The versions.</param>
        /// <param name="cabSize">Size of the CAB.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <param name="logFile">if set to <c>true</c> [log file].</param>
        /// <param name="filename">The filename.</param>
        /// <returns></returns>
        public static string ExportSite(string sourceurl, 
            bool haltOnWarning, 
            bool haltOnFatalError, 
            bool noFileCompression, 
            bool includeUserSecurity, 
            bool quiet,
            SPIncludeVersions versions,
            int cabSize,
            bool overwrite,
            bool logFile,
            string filename)
        {
            /**
            stsadm.exe -o export
            -url <URL to be exported>
            -filename <export file name>
            [-overwrite]
            [-includeusersecurity]
            [-haltonwarning]
            [-haltonfatalerror]
            [-nologfile]
            [-versions <1-4>
               1 - Last major version for files and list items (default)
               2 - The current version, either the last major or the last minor
               3 - Last major and last minor version for files and list items
               4 - All versions for files and list items]
            [-cabsize <integer from 1-1024 megabytes> (default: 25)]
            [-nofilecompression]
            [-quiet]
             * */

            string command = string.Format(" -o export -url \"{0}\"{2}{3}{4}{5} -versions {6}{7}{8}{9}{10} -filename \"{1}\"",
                sourceurl,
                filename,
                (haltOnWarning ? " -haltonwarning" : ""),
                (haltOnFatalError ? " -haltonfatalerror" : ""),
                (noFileCompression ? " -nofilecompression" : ""),
                (includeUserSecurity ? " -includeusersecurity" : ""),
                (int)versions,
                (cabSize > 0 ? " -cabsize " + cabSize : ""),
                (quiet ? " -quiet" : ""),
                (!logFile ? " -nologfile" : ""),
                (overwrite ? " -overwrite" : ""));

            if (Utilities.RunStsAdmOperation(command, quiet) != 0)
                throw new SPException("Error occured exporting site.\r\nCOMMAND: " + command);

            // If we're using file compression then we need to add the "cmp" extension that the export tool automatically appends.
            if (!noFileCompression)
                filename += ".cmp";

            return filename;
        }

    }
}
