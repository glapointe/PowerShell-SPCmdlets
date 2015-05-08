using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell
{
    internal static class Logger
    {
        private static SPCmdlet _executingCmdlet;

        public static SPCmdlet ExecutingCmdlet
        {
            get { return _executingCmdlet; }
            set
            {
                _executingCmdlet = value;
                if (value != null)
                    Verbose = (value.MyInvocation.Line.IndexOf("-Verbose", StringComparison.InvariantCultureIgnoreCase) >= 0);
            }
        }


        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="Logger"/> is verbose.
        /// </summary>
        /// <value><c>true</c> if verbose; otherwise, <c>false</c>.</value>
        public static bool Verbose
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the log file.
        /// </summary>
        /// <value>The log file.</value>
        public static string LogFile
        {
            get;
            set;
        }

        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="args">The args.</param>
        internal static void Write(string message, params string[] args)
        {
            Write(message, EventLogEntryType.Information, args);
        }

        internal static void WriteException(System.Management.Automation.ErrorRecord error)
        {
            if (ExecutingCmdlet != null)
                ExecutingCmdlet.WriteError(error);
            else
                Console.WriteLine(error.ToString());

            if (!string.IsNullOrEmpty(LogFile))
                File.AppendAllText(LogFile, error.ToString() + "\r\n");
        }

        internal static void WriteWarning(string message, params string[] args)
        {
            Write(message, EventLogEntryType.Warning, args);
        }

        /// <summary>
        /// Logs the specified STR.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="msgType">Type of the MSG.</param>
        /// <param name="args">The args.</param>
        internal static void Write(string message, EventLogEntryType msgType, params string[] args)
        {
            if (string.IsNullOrEmpty(message))
                return;

            message = string.Format(message, args);

            if (ExecutingCmdlet != null)
            {
                if (msgType == EventLogEntryType.Warning)
                {
                    ExecutingCmdlet.WriteWarning(message);
                }
                else if (msgType == EventLogEntryType.Error)
                {
                    ExecutingCmdlet.WriteError(new System.Management.Automation.ErrorRecord(new Exception(message), null, System.Management.Automation.ErrorCategory.NotSpecified, null));
                }
                ExecutingCmdlet.WriteVerbose(message);
            }
            else
            {
                if (msgType == EventLogEntryType.Warning && !message.ToUpper().StartsWith("WARNING:"))
                    message = "WARNING: " + message;

                if (msgType != EventLogEntryType.Information || Verbose)
                {
                    Console.WriteLine(message);
                }
            }
            if (!string.IsNullOrEmpty(LogFile))
                File.AppendAllText(LogFile, message + "\r\n");
        }
    }
}
