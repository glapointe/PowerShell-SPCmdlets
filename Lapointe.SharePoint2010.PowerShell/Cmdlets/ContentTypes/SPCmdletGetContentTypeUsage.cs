using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.ContentTypes
{

    [Cmdlet(VerbsCommon.Get, "SPContentTypeUsage", SupportsShouldProcess = true, DefaultParameterSetName = "Url")]
    [SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = false)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Retrieve an SPContentTypeUsage object corresponding to the List or Site where a Content Type is utilized.")]
    [RelatedCmdlets(typeof(SPCmdletGetContentType), typeof(SPCmdletGetList), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> $ct = Get-SPWeb \"http://server_name\" | Get-SPContentTypeUsage -Identity \"ContentType1\"",
        Remarks = "This example retrieves SPContentTypeUsage objects for all the Lists and Webs where ContentType1 is used.")]
    public class SPCmdletGetContentTypeUsage : SPGetCmdletBaseCustom<SPContentTypeUsage>
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The site.</value>
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The source web containing the content type usages to retrieve.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        public SPWebPipeBind[] Web { get; set; }

        /// <summary>
        /// Gets or sets the name of the content type.
        /// </summary>
        /// <value>The name of the contentType.</value>
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = false,
            HelpMessage = "The name or ID of the Content Type to return associated usages.")]
        [ValidateNotNullOrEmpty]
        public SPContentTypePipeBind Identity { get; set; }

        /// <summary>
        /// Gets or sets whether to only return only list scoped content types.
        /// </summary>
        /// <value>The list scope only.</value>
        [Parameter(Mandatory = false, Position = 1, ValueFromPipeline = false,
            HelpMessage = "Specify to return only Content Type usages scoped to a list.")]
        public SwitchParameter ListScopeOnly { get; set; }

        /// <summary>
        /// Processes the record.
        /// </summary>
        protected override IEnumerable<SPContentTypeUsage> RetrieveDataObjects()
        {
            foreach (SPWebPipeBind webPipeBind in Web)
            {
                using (SPWeb web = webPipeBind.Read())
                {
                    SPContentType ct = Identity.Read(web);
                    IList<SPContentTypeUsage> contentTypes = SPContentTypeUsage.GetUsages(ct);

                    WriteResult(contentTypes.Where(ctu => ListScopeOnly && ctu.IsUrlToList || !ListScopeOnly));

                }
            }
            return null;
        }
    }

}
