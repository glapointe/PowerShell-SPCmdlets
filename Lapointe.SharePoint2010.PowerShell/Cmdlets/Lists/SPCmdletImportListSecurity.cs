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

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Import", "SPListSecurity", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Export the security settings and permissions from a list.")]
    [RelatedCmdlets(typeof(SPCmdletCopyListSecurity), typeof(SPCmdletExportListSecurity), typeof(SPCmdletGetList))]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/lists/list1\" | Import-SPListSecurity -Path \"c:\\listsecurity.xml\"",
        Remarks = "This example imports the security settings and permissions from c:\\listsecurity.xml and applies them to list1.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name\" | Import-SPListSecurity -Path \"c:\\listsecurity.xml\"",
        Remarks = "This example imports the security settings and permissions from c:\\listsecurity.xml and applies them to matching lists found in http://server_name.")]
    public sealed class SPCmdletImportListSecurity : SPCmdletCustom
    {

        [Parameter(Mandatory = true, ParameterSetName = "SPList",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The list whose security will be updated.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPListPipeBind List { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "SPWeb",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 1,
            HelpMessage = "Specifies the URL or GUID of the Web containing the lists whose security will be updated.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            HelpMessage = "The path to the file containing the security settings to import.")]
        [Alias("Path")]
        public XmlDocumentPipeBind Input { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "In addition to the list level security copy the security settings of individual items.")]
        public SwitchParameter IncludeItemSecurity { get; set; }

        protected override void InternalProcessRecord()
        {
            switch (ParameterSetName)
            {
                case "SPWeb":
                    using (SPWeb web = Web.Read())
                    {
                        try
                        {
                            WriteVerbose("Importing list security to web " + web.Url);
                            Common.Lists.ImportListSecurity.ImportSecurity(Input.Read(), IncludeItemSecurity, web);
                        }
                        finally
                        {
                            web.Site.Dispose();
                        }
                    }
                    break;
                case "SPList":
                    SPList list = List.Read();
                    try
                    {
                        WriteVerbose("Importing list security to list " + list.RootFolder.Url);
                        Common.Lists.ImportListSecurity.ImportSecurity(Input.Read(), IncludeItemSecurity, list);
                    }
                    finally
                    {
                        list.ParentWeb.Dispose();
                        list.ParentWeb.Site.Dispose();
                    }
                    break;
            }
        }
    }

}
