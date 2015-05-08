using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Threading;

namespace Lapointe.SharePoint.PowerShell.Common.TimerJobs
{
    internal class ExecAdmSvcJobs
    {

        /// <summary>
        /// Executes the timer jobs.
        /// </summary>
        /// <param name="local">if set to <c>true</c> [local].</param>
        public static void Execute(bool local)
        {
            Execute(local, false);
        }

        /// <summary>
        /// Executes the timer jobs.
        /// </summary>
        /// <param name="local">if set to <c>true</c> [local].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        public static void Execute(bool local, bool quiet)
        {
            // First run the OOTB execadmsvcjobs on the local machine to make sure that any local jobs get executed
            if (!quiet)
                Console.WriteLine("\r\nExecuting jobs on {0}", SPServer.Local.Name);

            Utilities.RunStsAdmOperation("-o execadmsvcjobs", quiet);
            // If local was passed in then we're basically just using the OOTB command - I included this for testing only - it's not
            // really helpful otherwise.
            if (!local)
            {
                foreach (SPServer server in SPFarm.Local.Servers)
                {
                    // Only look at servers with a valid role.
                    if (server.Role == SPServerRole.Invalid)
                        continue;

                    // Don't need to check locally as we just ran the OOTB command locally so skip the local server.
                    if (server.Id.Equals(SPServer.Local.Id))
                        continue;
                    
                    bool stillExecuting;
                    if (!quiet)
                        Console.WriteLine("\r\nChecking jobs on {0}", server.Name);

                    do
                    {
                        List<SPJobDefinition> jobs = GetAdminCmdLineJobs(server);

                        stillExecuting = jobs.Count > 0; // CheckApplicableRunningJobs(server, quiet);

                        // If jobs are still executing then sleep for 1 second.
                        if (stillExecuting)
                        {
                            foreach (SPJobDefinition job in jobs)
                            {
                                Logger.Write("Waiting on {0}", job.Name);
                            }
                            Thread.Sleep(1000);
                        }
                    } while (stillExecuting);

                }
            }
        }



        private static List<SPJobDefinition> GetAdminCmdLineJobs(SPServer server)
        {
            List<SPJobDefinition> list = new List<SPJobDefinition>();
            foreach (KeyValuePair<Guid, SPService> pair in GetLocalProvisionedServices())
            {
                SPService service = pair.Value;
                foreach (SPJobDefinition definition in service.JobDefinitions)
                {
                    if (ShouldRunAdminCmdLineJob(definition, server))
                    {
                        list.Add(definition);
                    }
                }
                SPWebService service2 = service as SPWebService;
                if (service2 != null)
                {
                    foreach (SPWebApplication application in service2.WebApplications)
                    {
                        foreach (SPJobDefinition definition2 in application.JobDefinitions)
                        {
                            if (ShouldRunAdminCmdLineJob(definition2, server))
                            {
                                list.Add(definition2);
                            }
                        }
                    }
                    continue;
                }
            }
            return list;
        }

        private static bool ShouldRunAdminCmdLineJob(SPJobDefinition jd, SPServer server)
        {
            if (!(jd is SPAdministrationServiceJobDefinition))
            {
                return false;
            }
            if (jd.GetType().Name == "SPTimerRecycleJobDefinition")
            {
                return false;
            }
            if (!IsApplicableJob(jd, server, false, false))
            {
                return false;
            }
            if (jd.LockType == SPJobLockType.ContentDatabase)
            {
                return false;
            }
            return true;
        }

        private static bool IsApplicableJob(SPJobDefinition jd, SPServer server, bool isRefresh, bool logSpadminDisabledEvent)
        {
            if ((jd.Server != null) && !jd.Server.Id.Equals(server.Id))
            {
                return false;
            }
            if ((jd.LockType == SPJobLockType.ContentDatabase) && !SupportsContentDatabaseJobs(server))
            {
                return false;
            }
            if (!isRefresh)
            {
                return false; // ((jd.Flags & 1) == 0);
            }
            return true;
        }

        private static bool SupportsContentDatabaseJobs(SPServer server)
        {
            if (server == null)
            {
                return false;
            }
            SPWebServiceInstance instance = server.ServiceInstances.GetValue<SPWebServiceInstance>();
            if ((instance == null) || (instance.Status != SPObjectStatus.Online))
            {
                return false;
            }
            SPTimerServiceInstance instance2 = server.ServiceInstances.GetValue<SPTimerServiceInstance>();
            return ((instance2 != null) && instance2.AllowContentDatabaseJobs);
        }

        private static Dictionary<Guid, SPService> GetLocalProvisionedServices()
        {
            Dictionary<Guid, SPService> dictionary = new Dictionary<Guid, SPService>(8);
            foreach (SPServiceInstance instance in SPServer.Local.ServiceInstances)
            {
                SPService service = instance.Service;
                if (instance.Status == SPObjectStatus.Online)
                {
                    if (!dictionary.ContainsKey(service.Id))
                    {
                        dictionary.Add(service.Id, service);
                    }
                }
            }
            return dictionary;
        }
    }
}
