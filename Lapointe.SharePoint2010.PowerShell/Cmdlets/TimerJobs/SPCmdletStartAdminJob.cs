using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.TimerJobs
{
    [Cmdlet("Start", "SPAdminJob2", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Timer Jobs")]
    [CmdletDescription("Immediately starts any waiting administrative job on the local computer and triggers the jobs to run on all other servers in the farm if applicable.",
        "Use the Start-SPAdminJob2 cmdlet to execute all administrative timer jobs immediately rather than waiting for the timer job to run.\r\n\r\nWhen the process account for the SharePoint 2010 Administration service (SPAdminV4)) is disabled (necessary in some installations for security reasons), the Start-SPAdminJob2 cmdlet triggers all administrative tasks to run immediately on all servers to allow provisioning and other administrative tasks that SPAdmin ordinarily handles.\r\n\r\nWhen you run this cmdlet in person (not in script), use the Verbose parameter to see the individual administrative operations that are run.")]
    [RelatedCmdlets(ExternalCmdlets = new[] {"Start-SPAdminJob"})]
    [Example(Code = "PS C:\\> Start-SPAdminJob2 -Verbose",
        Remarks = "This example runs all waiting administrative jobs on all servers and shows verbose output to the administrator.")]
    public class SPCmdletStartAdminJob : SPCmdletCustom
    {
        protected override void InternalProcessRecord()
        {
            Common.TimerJobs.ExecAdmSvcJobs.Execute(Local.IsPresent);
        }

        [Parameter(Mandatory = false,
            HelpMessage = "Providing this parameter causes the cmdlet to behave just like the built-in Start-SPAdminJob cmdlet.")]
        public SwitchParameter Local { get; set; }
    }

}
