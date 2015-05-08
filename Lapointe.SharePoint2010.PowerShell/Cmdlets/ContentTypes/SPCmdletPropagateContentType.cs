using System;
using System.Collections.Generic;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.ContentTypes
{
    [Cmdlet("Propagate", "SPContentType", SupportsShouldProcess = false, DefaultParameterSetName = "SPWeb"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = false)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Propagates changes to a parent content type to all child content types.")]
    [RelatedCmdlets(typeof(SPCmdletCopyContentType), typeof(SPCmdletExportContentType), typeof(SPCmdletGetContentType),
        ExternalCmdlets = new[] { "Get-SPSite" })]
    [Example(Code = "PS C:\\> Get-SPSite \"http://server_name\" | Propagate-SPContentType -Identity \"ContentType1\" -UpdateFields",
        Remarks = "This example propgates changes to ContentType1 found in http://server_name to all child content types.")]
    public class SPCmdletPropagateContentType : SPCmdletCustom
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The source site containing the content types to propagate.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [ValidateNotNull]
        public SPSitePipeBind Site { get; set; }

        /// <summary>
        /// Gets or sets the name of the content type.
        /// </summary>
        /// <value>The name of the contentType.</value>
        [Parameter(Mandatory = true, 
            Position = 0, 
            ValueFromPipeline = false, 
            HelpMessage = "The content type to propagate changes. The type must be a valid content type name; a valid content type ID, in the form 0x0123...; or an instance of an SPContentType object.")]
        public SPContentTypePipeBind Identity { get; set; }


        [Parameter(Mandatory = false, HelpMessage = "Propagate changes to all fields.")]
        public SwitchParameter UpdateFields { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Remove fields defined in child content types which do not exist in the source content type.")]
        public SwitchParameter RemoveFields { get; set; }

        protected override void InternalProcessRecord()
        {
            SPContentType ct = null;

            if (Site != null)
            {
                using (SPSite site = Site.Read(true))
                {
                    ct = Identity.Read(site.RootWeb);
                    Common.ContentTypes.PropagateContentType.Execute(ct, UpdateFields.IsPresent, RemoveFields.IsPresent);
                }
            }
            else
            {
                ct = Identity.Read();
                Common.ContentTypes.PropagateContentType.Execute(ct, UpdateFields.IsPresent, RemoveFields.IsPresent);
            }
        }
    }
}
