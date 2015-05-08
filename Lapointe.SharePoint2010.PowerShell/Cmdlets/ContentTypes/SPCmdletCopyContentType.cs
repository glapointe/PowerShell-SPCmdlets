using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using Microsoft.SharePoint.Deployment;
using System.IO;
using Microsoft.SharePoint.Administration.Backup;
using System.Collections;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.ContentTypes
{
    [Cmdlet(VerbsCommon.Copy, "SPContentType", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Copies all Content Types from one gallery to another.")]
    [RelatedCmdlets(typeof(SPCmdletExportContentType), typeof(SPCmdletGetContentType), typeof(SPCmdletPropagateContentType), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> Copy-SPContentType -SourceContentType ContentType1 -SourceWeb \"http://server_name/sites/site1\" -TargetWeb \"http://server_name/sites/site2\"",
        Remarks = "This example copies the ContentType1 content type from site1 to site2.")]
    public class SPCmdletCopyContentType : SPCmdletCustom
    {
        [Parameter(Mandatory = true,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        HelpMessage = "The content type to copy. The type must be a valid content type name; a valid content type ID, in the form 0x0123...; or an instance of an SPContentType object.")]
        public SPContentTypePipeBind SourceContentType
        {
            get;
            set;
        }

        [Parameter(Mandatory = false,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        HelpMessage = "The source web containing the content type to copy.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPWebPipeBind SourceWeb
        {
            get;
            set;
        }

        [Parameter(Mandatory = true,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        HelpMessage = "The target web containing to copy the content type to.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPWebPipeBind TargetWeb
        {
            get;
            set;
        }

        [Parameter(Mandatory = false, HelpMessage = "Do not copy any workflow associations.")]
        public SwitchParameter NoWorkflows
        {
            get;
            set;
        }

        [Parameter(Mandatory = false, HelpMessage = "Do not copy the document template.")]
        public SwitchParameter NoDocTemplate
        {
            get;
            set;
        }

        [Parameter(Mandatory = false, HelpMessage = "Do not copy the information rights policies.")]
        public SwitchParameter NoPolicies
        {
            get;
            set;
        }

        [Parameter(Mandatory = false, HelpMessage = "Do not copy the document information panel.")]
        public SwitchParameter NoDocInfoPanel
        {
            get;
            set;
        }

        [Parameter(Mandatory = false, HelpMessage = "Do not copy the document conversion settings.")]
        public SwitchParameter NoDocConversions
        {
            get;
            set;
        }

        [Parameter(Mandatory = false, HelpMessage = "Do not copy fields associated with the content type.")]
        public SwitchParameter NoColumns
        {
            get;
            set;
        }       
        
        protected override void InternalProcessRecord()
        {
            bool copyWorkflows = !NoWorkflows.IsPresent;
            bool copyColumns = !NoColumns.IsPresent;
            bool copyDocConversions = !NoDocConversions.IsPresent;
            bool copyDocInfoPanel = !NoDocInfoPanel.IsPresent;
            bool copyPolicies = !NoPolicies.IsPresent;
            bool copyDocTemplate = !NoDocTemplate.IsPresent;

            SPWeb sourceWeb = null;
            SPContentType sourceCT = null;

            if (SourceWeb != null)
                sourceWeb = SourceWeb.Read();
            try
            {
                sourceCT = SourceContentType.Read(sourceWeb);
            }
            finally
            {
                if (sourceWeb != null)
                    sourceWeb.Dispose();
            }
            using (SPWeb targetWeb = TargetWeb.Read())
            {
                Logger.Write("Start Time: {0}", DateTime.Now.ToString());

                Common.ContentTypes.CopyContentTypes ctCopier = new Common.ContentTypes.CopyContentTypes(
                    copyWorkflows, copyColumns, copyDocConversions, copyDocInfoPanel, copyPolicies, copyDocTemplate);

                ctCopier.Copy(sourceCT, targetWeb);
            }
            Logger.Write("Finish Time: {0}", DateTime.Now.ToString());
        }


    }

}
