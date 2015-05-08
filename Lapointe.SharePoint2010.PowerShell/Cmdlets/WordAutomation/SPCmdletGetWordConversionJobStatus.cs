using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.Office.Word.Server.Conversions;
using Microsoft.Office.Word.Server.Service;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WordAutomation
{
    [Cmdlet(VerbsCommon.Get, "SPWordConversionJobStatus", SupportsShouldProcess = false),
     SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Word Automation")]
    [CmdletDescription("Retrieves the conversion job status details.")]
    [RelatedCmdlets(typeof(SPCmdletNewWordConversionJob))]
    [Example(Code = "PS C:\\> $job = New-SPWordConversionJob -InputList \"http://server_name/WordDocs\" -OutputList \"http://server_name/PDFDocs\" -OutputFormt PDF -OutputSaveBehavior AppendIfPossible\r\nPS C:\\> Get-SPWordConversionJobStatus $job",
        Remarks = "This example converts the documents located in http://server_name/WordDocs and stores the converted items in http://server_name/PDFDocs and then outputs the jobs status.")]
    public class SPCmdletGetWordConversionJobStatus : SPGetCmdletBaseCustom<ConversionJobStatus>
    {
        [Parameter(Mandatory = true, 
            Position = 0, 
            ValueFromPipeline = true,
            HelpMessage = "The conversion job whose status should be retrieved.")]
        public SPWordConversionJobPipeBind Job { get; set; }

        protected override IEnumerable<ConversionJobStatus> RetrieveDataObjects()
        {
            List<ConversionJobStatus> statuses = new List<ConversionJobStatus>();
            foreach (WordServiceApplicationProxy proxy in GetProxies())
            {
                ConversionJobStatus jobStatus = new ConversionJobStatus(proxy, Job.Read().ID, null);
                if (jobStatus != null)
                    statuses.Add(jobStatus);
            }
            return statuses;
        }

        private List<WordServiceApplicationProxy> GetProxies()
        {
            SPFarm local = SPFarm.Local;
            List<WordServiceApplicationProxy> proxies = new List<WordServiceApplicationProxy>();
            if (null != local)
            {
                foreach (SPServiceProxy svcProcies in local.ServiceProxies)
                {
                    if (null != svcProcies)
                    {
                        foreach (SPServiceApplicationProxy appProxies in svcProcies.ApplicationProxies)
                        {
                            if (null == appProxies || typeof(WordServiceApplicationProxy) != appProxies.GetType())
                            {
                                continue;
                            }
                            proxies.Add((WordServiceApplicationProxy)appProxies);
                        }
                    }
                }
            }
            return proxies;
        }

 

 

    }


}
