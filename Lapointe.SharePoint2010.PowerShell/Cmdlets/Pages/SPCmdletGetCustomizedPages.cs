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

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Pages
{
    [Cmdlet(VerbsCommon.Get, "SPCustomizedPages", SupportsShouldProcess = false, DefaultParameterSetName = "SPSite"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Pages")]
    [CmdletDescription("Retrieves all customized (unghosted) pages at the given scope. If not returning as a string using the Start-SPAssignment and Stop-SPAssignment cmdlets to dispose of all parent objects.")]
    [RelatedCmdlets(typeof(SPCmdletResetCustomizedPages), ExternalCmdlets = new[] { "Get-SPWeb", "Get-SPSite", "Start-SPAssignment", "Stop-SPAssignment" })]
    [Example(Code = "PS C:\\> $pages = Get-SPWeb http://server_name/subweb | Get-SPCustomizedPages -Recurse",
        Remarks = "This example returns back all customized, or unghosted, pages from all webs under http://server_name/subweb, inclusive.")]
    [Example(Code = "PS C:\\> $pages = Get-SPSite http://server_name/ | Get-SPCustomizedPages",
        Remarks = "This example returns back all customized, or unghosted, pages from all webs at the site collection http://server_name/.")]
    public class SPCmdletGetCustomizedPages : SPGetCmdletBaseCustom<object>
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to retrieve customized pages from.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind[] Web { get; set; }

        /// <summary>
        /// Gets or sets the site.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPSite",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The site to retrieve customized pages from. All sub-webs will be iterated through.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [ValidateNotNull]
        public SPSitePipeBind[] Site { get; set; }

        /// <summary>
        /// Gets or sets whether to recurse all webs
        /// </summary>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = false,
            HelpMessage = "Iterate through all child-webs of the specified web.")]
        public SwitchParameter Recurse{ get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Returns the results as a string. If this parameter is not provided the results will be returned as SPFile objects and the caller will be responsible for disposing the parent objects using the Start-SPAssignment and Stop-SPAssignment cmdlets.")]
        public SwitchParameter AsString { get; set; }

        protected override IEnumerable<object> RetrieveDataObjects()
        {
            List<object> customizedPages = new List<object>();

            switch (ParameterSetName)
            {
                case "SPWeb":
                    foreach (SPWebPipeBind webPipe in Web)
                    {
                        SPWeb web = webPipe.Read();

                        if (Recurse.IsPresent)
                        {
                            Common.Pages.EnumUnGhostedFiles.RecurseSubWebs(web, ref customizedPages, AsString.IsPresent);
                        }
                        else
                            Common.Pages.EnumUnGhostedFiles.CheckFoldersForUnghostedFiles(web.RootFolder, ref customizedPages, AsString.IsPresent);
                    }
                    break;
                case "SPSite":
                    foreach (SPSitePipeBind sitePipe in Site)
                    {
                        SPSite site = sitePipe.Read();
                        Common.Pages.EnumUnGhostedFiles.RecurseSubWebs(site.RootWeb, ref customizedPages, AsString.IsPresent);
                    }
                    break;
            }

            foreach (object page in customizedPages)
            {
                if (!AsString.IsPresent)
                {
                    AssignmentCollection.Add(((SPFile)page).Web);
                    AssignmentCollection.Add(((SPFile)page).Web.Site);
                }
                WriteResult(page);
            }

            return null;
        }
    }
}
