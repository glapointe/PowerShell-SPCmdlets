using System.Text;
using System.Collections.Generic;
using Lapointe.SharePoint.PowerShell.Common.ManagedMetadata;
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
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint.Taxonomy;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.ManagedMetadata
{
    [Cmdlet("Export", "SPTerms", SupportsShouldProcess = false, DefaultParameterSetName = "TaxonomySession"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Export the Managed Metadata Terms.")]
    [RelatedCmdlets(typeof(SPCmdletImportTerms), ExternalCmdlets = new[] { "Get-SPTaxonomySession" })]
    [Example(Code = "PS C:\\> Export-SPTerms -TaxonomySession \"http://site/\" -OutputFile \"c:\\terms.xml\"",
        Remarks = "This example exports the terms for all term stores associated with the site and saves to c:\\terms.xml.")]
    [Example(Code = "PS C:\\> Export-SPTerms -Group (Get-SPTaxonomySession -Site \"http://site/\").TermStores[0].Groups[0] -OutputFile \"c:\\terms.xml\"",
        Remarks = "This example exports the first Group of the first Term Store and saves to c:\\terms.xml.")]
    public sealed class SPCmdletExportTerms : SPCmdletCustom
    {
        [Parameter(ParameterSetName = "TaxonomySession",
            Mandatory = true,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The TaxonomySession object containing the Term Stores to export.")]
        public SPTaxonomySessionPipeBind TaxonomySession { get; set; }

        [Parameter(ParameterSetName = "TermStore",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermStore object containing the terms to export.")]
        public SPTaxonomyTermStorePipeBind TermStore { get; set; }

        [Parameter(ParameterSetName = "Group",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Group object containing the terms to export.")]
        public SPTaxonomyGroupPipeBind Group { get; set; }

        [Parameter(ParameterSetName = "TermSet",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermSet object containing the terms to export.")]
        public SPTaxonomyTermSetPipeBind TermSet { get; set; }

        [Parameter(ParameterSetName = "Term",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Term object containing the terms to export.")]
        public SPTaxonomyTermPipeBind Term { get; set; }


        [Parameter(Mandatory = false,
            HelpMessage = "The path to the file to save the terms to.",
            Position = 1)]
        [ValidateDirectoryExistsAndValidFileName]
        [Alias("Path")]
        public string OutputFile { get; set; }


        protected override void InternalProcessRecord()
        {
            ExportTerms export = new ExportTerms();
            XmlDocument xml = null;

            switch (ParameterSetName)
            {
                case "TaxonomySession":
                    xml = export.Export(TaxonomySession.Read());
                    break;
                case "TermStore":
                    xml = export.Export(TermStore.Read());
                    break;
                case "Group":
                    xml = export.Export(Group.Read());
                    break;
                case "TermSet":
                    xml = export.Export(TermSet.Read());
                    break;
                case "Term":
                    xml = export.Export(Term.Read());
                    break;
            }
            if (xml == null)
                return;

            if (!string.IsNullOrEmpty(OutputFile))
                xml.Save(OutputFile);
            else
                WriteResult(xml);
        }


    }

}
