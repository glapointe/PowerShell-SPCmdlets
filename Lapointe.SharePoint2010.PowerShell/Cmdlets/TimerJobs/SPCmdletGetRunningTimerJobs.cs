using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using Microsoft.SharePoint;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.TimerJobs
{
    [Cmdlet(VerbsCommon.Get, "SPRunningTimerJobs"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Timer Jobs")]
    [CmdletDescription("")]
    [RelatedCmdlets(ExternalCmdlets = new[] {"Get-SPTimerJob"})]
    [Example(Code = "PS C:\\> Get-SPRunningTimerJobs | ? {((Get-Date).ToUniversalTime() - $_.StartTime).TotalHours -gt 4}",
        Remarks = "This example returns back all running timer jobs and filters to the results to show only those that have been running for more than four hours.")]
    [Example(Code = "PS C:\\> Get-SPTimerJob | ?{$_.Name -like {\"*User Profile*\"} | Get-SPRunningTimerJobs",
        Remarks = "This example returns back all running user profile timer jobs.")]
    public class SPCmdletGetRunningTimerJobs : SPGetCmdletBaseCustom<SPRunningJob>
    {
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 0)]
        public SPTimerJobPipeBind Identity { get; set; }

        protected override IEnumerable<SPRunningJob> RetrieveDataObjects()
        {
            List<SPRunningJob> runningJobs = new List<SPRunningJob>();

            SPJobDefinition definition = null;
            if (Identity != null)
            {
                definition = this.Identity.Read();
                if (null == definition)
                {
                    base.ThrowTerminatingError(new SPCmdletException("Job definition was not found."), ErrorCategory.ObjectNotFound, null);
                }

            }
            foreach (var svc in SPFarm.Local.Services) 
            {
                if (svc.RunningJobs.Count == 0) continue;
                foreach (SPRunningJob job in svc.RunningJobs)
                {
                    if (null == definition || job.JobDefinitionId == definition.Id)
                        runningJobs.Add(job);
                }
            }

            return runningJobs;
        }
    }
}
