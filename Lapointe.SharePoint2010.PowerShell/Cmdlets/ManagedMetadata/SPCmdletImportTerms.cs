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
    [Cmdlet("Import", "SPTerms", SupportsShouldProcess = true, DefaultParameterSetName = "TaxonomySession"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Import the Managed Metadata Terms.")]
    [RelatedCmdlets(typeof(SPCmdletExportTerms), ExternalCmdlets = new[] { "Get-SPTaxonomySession" })]
    [Example(Code = "PS C:\\> Import-SPTerms -TaxonomySession \"http://site/\" -InputFile \"c:\\terms.xml\"",
        Remarks = "This example imports the terms from c:\\terms.xml to the Term Store associated with http://site.")]
    [Example(Code = "PS C:\\> Import-SPTerms -ParentTermStore (Get-SPTaxonomySession -Site \"http://site/\").TermStores[0] -InputFile \"c:\\terms.xml\"",
        Remarks = "This example imports the Group from c:\\terms.xml to the first Term Store.")]
    public sealed class SPCmdletImportTerms : SPCmdletCustom
    {
        XmlDocument _xml = null;

        [Parameter(ParameterSetName = "TaxonomySession",
            Mandatory = true,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The TaxonomySession object to import Term Stores into.")]
        public SPTaxonomySessionPipeBind TaxonomySession { get; set; }

        [Parameter(ParameterSetName = "TermStore",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermStore object to import Groups into.")]
        public SPTaxonomyTermStorePipeBind ParentTermStore { get; set; }

        [Parameter(ParameterSetName = "Group",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Group object to import Term Sets into.")]
        public SPTaxonomyGroupPipeBind ParentGroup { get; set; }

        [Parameter(ParameterSetName = "TermSet",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermSet object to import Terms into.")]
        public SPTaxonomyTermSetPipeBind ParentTermSet { get; set; }

        [Parameter(ParameterSetName = "Term",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Term object to import Terms into.")]
        public SPTaxonomyTermPipeBind ParentTerm { get; set; }


        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            HelpMessage = "The path to the file containing the terms to import or an XmlDocument object or XML string.",
            Position = 1)]
        [Alias("Xml")]
        public XmlDocumentPipeBind InputFile { get; set; }

        protected override void InternalValidate()
        {
            base.InternalValidate();

            _xml = InputFile.Read();
            string rootElement = _xml.DocumentElement.Name;
            bool isValidRoot = false;

            switch (ParameterSetName)
            {
                case "TaxonomySession":
                    isValidRoot = rootElement == "TermStores" || rootElement == "TermStore";
                    break;
                case "TermStore":
                    isValidRoot = rootElement == "Groups" || rootElement == "Group";
                    break;
                case "Group":
                    isValidRoot = rootElement == "TermSets" || rootElement == "TermSet";
                    break;
                case "TermSet":
                    isValidRoot = rootElement == "Terms" || rootElement == "Term";
                    break;
                case "Term":
                    isValidRoot = rootElement == "Terms" || rootElement == "Term";
                    break;
            }
            if (!isValidRoot)
            {
                string msg = "The import file cannot be imported to the specified target location. The following details the allowed import targets:\r\n";
                msg += "\tTaxonomySession: Term Store\r\n";
                msg += "\tTerm Store: Group\r\n";
                msg += "\tGroup: Term Set\r\n";
                msg += "\tTerm Set: Term\r\n";
                msg += "\tTerm: Term";
                throw new SPCmdletException(msg);
            }
        }

        protected override void InternalProcessRecord()
        {
            try
            {
                ShouldProcessReason reason;
                bool whatIf = false;
                if (!base.ShouldProcess(null, null, null, out reason))
                {
                    if (reason == ShouldProcessReason.WhatIf)
                    {
                        whatIf = true;
                        Logger.Verbose = true;
                    }
                }
                Logger.Write("Start Time: {0}", DateTime.Now.ToString());

                ImportTerms import = new ImportTerms(_xml, whatIf);

                switch (ParameterSetName)
                {
                    case "TaxonomySession":
                        import.Import(TaxonomySession.Read());
                        break;
                    case "TermStore":
                        import.Import(ParentTermStore.Read());
                        break;
                    case "Group":
                        import.Import(ParentGroup.Read());
                        break;
                    case "TermSet":
                        import.Import(ParentTermSet.Read());
                        break;
                    case "Term":
                        import.Import(ParentTerm.Read());
                        break;
                }
            }
            finally
            {
                Logger.Write("Finish Time: {0}", DateTime.Now.ToString());
            }
        }


    }

}
