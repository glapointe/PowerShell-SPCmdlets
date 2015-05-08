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
using System.Xml;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Export", "SPListSecurity", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Export the security settings and permissions from a list.")]
    [RelatedCmdlets(typeof(SPCmdletCopyListSecurity), typeof(SPCmdletImportListSecurity), typeof(SPCmdletGetList))]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/lists/list1\" | Export-SPListSecurity -OutputFile \"c:\\listsecurity.xml\"",
        Remarks = "This example exports the security settings and permissions from list1 and saves to c:\\listsecurity.xml.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name\" | Export-SPListSecurity -OutputFile \"c:\\listsecurity.xml\"",
        Remarks = "This example exports the security settings and permissions for all lists in http://server_name and saves to c:\\listsecurity.xml.")]
    public sealed class SPCmdletExportListSecurity : SPCmdletCustom
    {
        StringBuilder _sb;
        XmlTextWriter _xmlWriter;

        [Parameter(Mandatory = true, ParameterSetName = "SPList",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The list whose security will be exported.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPListPipeBind List { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "SPWeb",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 1,
            HelpMessage = "Specifies the URL or GUID of the Web containing the lists whose security will be exported.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The path to the file to save the security to.")]
        [ValidateDirectoryExistsAndValidFileName]
        [Alias("Path")]
        public string OutputFile { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "In addition to the list level security copy the security settings of individual items.")]
        public SwitchParameter IncludeItemSecurity { get; set; }

        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();
            _sb = new StringBuilder();
            _xmlWriter = Common.Lists.ExportListSecurity.OpenXmlWriter(_sb);
        }

        protected override void InternalEndProcessing()
        {
            base.InternalEndProcessing();

            Common.Lists.ExportListSecurity.CloseXmlWriter(_xmlWriter, OutputFile, _sb);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(_sb.ToString());
            WriteResult(xmlDoc);
        }

        protected override void InternalProcessRecord()
        {
            switch (ParameterSetName)
            {
                case "SPWeb":
                    using (SPWeb web = Web.Read())
                    {
                        try
                        {
                            WriteVerbose("Exporting list security from web " + web.Url);
                            foreach (SPList childList in web.Lists)
                            {
                                Common.Lists.ExportListSecurity.ExportSecurity(childList, web, _xmlWriter, IncludeItemSecurity.IsPresent);
                            }
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
                        WriteVerbose("Exporting list security from list " + list.RootFolder.Url);
                        Common.Lists.ExportListSecurity.ExportSecurity(list, list.ParentWeb, _xmlWriter, IncludeItemSecurity.IsPresent);
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
