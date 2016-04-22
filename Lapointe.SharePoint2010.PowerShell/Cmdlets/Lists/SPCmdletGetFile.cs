using System;
using System.Collections.Generic;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint.Publishing;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.Get, "SPFile", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Retrieves the SPFile object associated with the specified URL. Use the AssignmentCollection parameter to handle disposal of parent web and site objects.")]
    [RelatedCmdlets(typeof(Pages.SPCmdletGetPublishingPage), ExternalCmdlets = new[] {"Start-SPAssignment", "Stop-SPAssignment"})]
    [Example(Code = "PS C:\\> $file = Get-SPFile \"http://server_name/pages/default.aspx\"",
        Remarks = "This example returns back the default.aspx file from the http://server_name/pages library.")]
    [Example(Code = "PS C:\\> $file = Get-SPFile \"http://server_name/documents/HRHandbook.doc\"",
        Remarks = "This example returns back the HRHandbook.doc file from the http://server_name/documents library.")]
    public class SPCmdletGetFile : SPGetCmdletBaseCustom<SPFile>
    {
        [Parameter(Mandatory = true, 
            Position = 0, 
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The path to the publishing page to return."),
        Alias("File")]
        [ValidateNotNull]
        public SPFilePipeBind[] Identity { get; set; }

        protected override IEnumerable<SPFile> RetrieveDataObjects()
        {
            foreach (SPFilePipeBind filePipe in Identity)
            {
                SPFile file = filePipe.Read();
                AssignmentCollection.Add(file.Web);
                AssignmentCollection.Add(file.Web.Site);
                WriteResult(file);
            }

            return null;
        }
    }
}
