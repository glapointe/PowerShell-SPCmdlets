using System.Collections.Specialized;
using System.Text;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Logging
{
    public class TraceLog : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TraceLog"/> class.
        /// </summary>
        public TraceLog()
        {
            SPDiagnosticsService diagnosticsService = SPDiagnosticsService.Local;

            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("logdirectory", "log", false, diagnosticsService.LogLocation, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("logfilecount", "num", false, diagnosticsService.DaysToKeepLogs.ToString() , new SPIntRangeValidator(0, 1024)));
            parameters.Add(new SPParam("logfileminutes", "min", false, diagnosticsService.LogCutInterval.ToString(), new SPIntRangeValidator(0, 1440)));
            parameters.Add(new SPParam("allowlegacytraceproviders", "legacytracers", false, diagnosticsService.AllowLegacyTraceProviders.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("customerexperienceimprovementprogramenabled", "ceipenabled", false, diagnosticsService.CEIPEnabled.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("errorreportingenabled", "errreportenabled", false, diagnosticsService.ErrorReportingEnabled.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("errorreportingautomaticuploadenabled", "errreportautoupload", false, diagnosticsService.ErrorReportingAutomaticUpload.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("downloaderrorreportingupdatesenabled", "errreportupdates", false, diagnosticsService.DownloadErrorReportingUpdates.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("logmaxdiskspaceusageenabled", "maxdiskspaceusage", false, diagnosticsService.LogMaxDiskSpaceUsageEnabled.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("eventlogfloodprotectionenabled", "floodprotection", false, diagnosticsService.EventLogFloodProtectionEnabled.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("scripterrorreportingenabled", "scripterrorreporting", false, diagnosticsService.ScriptErrorReportingEnabled.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("scripterrorreportingrequireauth", "scripterrorreqauth", false, diagnosticsService.ScriptErrorReportingRequireAuth.ToString(), new SPTrueFalseValidator()));
            parameters.Add(new SPParam("eventlogfloodprotectionthreshold", "floodprotectionthreshold", false, diagnosticsService.EventLogFloodProtectionThreshold.ToString(), new SPIntRangeValidator(0, 100)));
            parameters.Add(new SPParam("eventlogfloodprotectiontriggerperiod", "floodprotectiontrigger", false, diagnosticsService.EventLogFloodProtectionTriggerPeriod.ToString(), new SPIntRangeValidator(1, 1440)));
            parameters.Add(new SPParam("scripterrorreportingdelay", "scriptreportingdelay", false, diagnosticsService.ScriptErrorReportingDelay.ToString(), new SPIntRangeValidator(1, 1440)));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nSets the log file location (note that the location must exist on each server) and the maximum number of log files to maintain and how long to capture events to a single file.\r\n\r\nParameters:");
            sb.Append("\r\n\t[-logdirectory <log file location>]");
            sb.Append("\r\n\t[-logfilecount <number of log files to create (0-1024)>]");
            sb.Append("\r\n\t[-logfileminutes <number of minutes to use a log file (0-1440)>]");
            sb.Append("\r\n\t[-allowlegacytraceproviders <true|false>]");
            sb.Append("\r\n\t[-customerexperienceimprovementprogramenabled <true|false>]");
            sb.Append("\r\n\t[-errorreportingenabled <true|false>]");
            sb.Append("\r\n\t[-errorreportingautomaticuploadenabled <true|false>]");
            sb.Append("\r\n\t[-downloaderrorreportingupdatesenabled <true|false>]");
            sb.Append("\r\n\t[-logmaxdiskspaceusageenabled <true|false>]");
            sb.Append("\r\n\t[-eventlogfloodprotectionenabled <true|false>]");
            sb.Append("\r\n\t[-scripterrorreportingenabled <true|false>]");
            sb.Append("\r\n\t[-scripterrorreportingrequireauth <true|false>]");
            sb.Append("\r\n\t[-eventlogfloodprotectionthreshold <number between 0 and 100>]");
            sb.Append("\r\n\t[-eventlogfloodprotectiontriggerperiod <number between 1 and 1440>]");
            sb.Append("\r\n\t[-scripterrorreportingdelay <number between 1 and 1440>]");

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

            string logDirectory = Params["logdirectory"].Value;
            int logFileCount = int.Parse(Params["logfilecount"].Value);
            int logFileMinutes = int.Parse(Params["logfileminutes"].Value);
            bool allowLegacyTraceProviders = bool.Parse(Params["allowlegacytraceproviders"].Value);
            bool customerExperienceImprovementProgramEnabled = bool.Parse(Params["customerexperienceimprovementprogramenabled"].Value);
            bool errorReportingEnabled = bool.Parse(Params["errorreportingenabled"].Value);
            bool errorReportingAutomaticUploadEnabled = bool.Parse(Params["errorreportingautomaticuploadenabled"].Value);
            bool downloadErrorReportingUpdatesEnabled = bool.Parse(Params["downloaderrorreportingupdatesenabled"].Value);
            bool logMaxDiskSpaceUsageEnabled = bool.Parse(Params["logmaxdiskspaceusageenabled"].Value);
            bool eventLogFloodProtectionEnabled = bool.Parse(Params["eventlogfloodprotectionenabled"].Value);
            bool scriptErrorReportingEnabled = bool.Parse(Params["scripterrorreportingenabled"].Value);
            bool scriptErrorReportingRequireAuth = bool.Parse(Params["scripterrorreportingrequireauth"].Value);
            int eventLogFloodProtectionThreshold = int.Parse(Params["eventlogfloodprotectionthreshold"].Value);
            int eventLogFloodProtectionTriggerPeriod = int.Parse(Params["eventlogfloodprotectiontriggerperiod"].Value);
            int scriptErrorReportingDelay = int.Parse(Params["scripterrorreportingdelay"].Value);

            SPDiagnosticsService local = SPDiagnosticsService.Local;
            local.LogLocation = logDirectory;
            local.DaysToKeepLogs = logFileCount;
            local.LogCutInterval = logFileMinutes;
            local.AllowLegacyTraceProviders = allowLegacyTraceProviders;
            local.CEIPEnabled = customerExperienceImprovementProgramEnabled;
            local.ErrorReportingEnabled = errorReportingEnabled;
            local.ErrorReportingAutomaticUpload = errorReportingAutomaticUploadEnabled;
            local.DownloadErrorReportingUpdates = downloadErrorReportingUpdatesEnabled;
            local.LogMaxDiskSpaceUsageEnabled = logMaxDiskSpaceUsageEnabled;
            local.EventLogFloodProtectionEnabled = eventLogFloodProtectionEnabled;
            local.ScriptErrorReportingEnabled = scriptErrorReportingEnabled;
            local.ScriptErrorReportingRequireAuth = scriptErrorReportingRequireAuth;
            local.EventLogFloodProtectionThreshold = eventLogFloodProtectionThreshold;
            local.EventLogFloodProtectionTriggerPeriod = eventLogFloodProtectionTriggerPeriod;
            local.ScriptErrorReportingDelay = scriptErrorReportingDelay;
            
            local.Update();

            return (int)ErrorCodes.NoError;
        }

        #endregion
    }
}
