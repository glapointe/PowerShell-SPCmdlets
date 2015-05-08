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
using Lapointe.PowerShell.MamlGenerator.Attributes;
using System.ComponentModel;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.ContentTypes
{
    [Cmdlet("Export", "SPContentType", SupportsShouldProcess = false, DefaultParameterSetName = "SPWeb"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = false)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Exports a Content Types to an XML file or the pipeline.")]
    [RelatedCmdlets(typeof(SPCmdletCopyContentType), typeof(SPCmdletGetContentType), typeof(SPCmdletPropagateContentType), 
        typeof(Lists.SPCmdletGetList), ExternalCmdlets = new[] {"Get-SPWeb"})]
    [Example(Code = "PS C:\\> Export-SPContentType -ContentType \"ContentType1\" -Web \"http://server_name/\" -OutputFile \"c:\\contentTypes.xml\"",
        Remarks = "This example exports the ContentType1 content type from http://server_name and saves it to contentTypes.xml.")]
    [Example(Code = "PS C:\\> Export-SPContentType -ContentType \"ContentType1\" -List \"http://server_name/lists/mylist\" -OutputFile \"c:\\contentTypes.xml\"",
        Remarks = "This example exports the ContentType1 content type from the list located at http://server_name/lists/mylist and saves it to contentTypes.xml.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name\" | Export-SPContentType -Identity \"ContentType1\" -OutputFile \"c:\\contentTypes.xml\"",
        Remarks = "This example exports ContentType1 from http://server_name and saves it to contentTypes.xml.")]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/lists/mylist/\" | Export-SPContentType -Identity \"ContentType1\" -OutputFile \"c:\\contentTypes.xml\"",
        Remarks = "This example exports ContentType1 from the list located at http://server_name/lists/mylist and saves them to contentTypes.xml.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name/*\" | Export-SPContentType -List \"Documents\" -OutputFile \"c:\\contentTypes.xml\"",
        Remarks = "This example exports all content types from all lists named Documents in all webs at http://server_name/* and saves them to contentTypes.xml.")]
    public class SPCmdletExportContentType : SPCmdletCustom
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The source web containing the content type to export.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        [Parameter(Mandatory = false, ParameterSetName = "SPContentType1",
            HelpMessage = "The source web containing the content type to export.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPWebPipeBind[] Web { get; set; }

        /// <summary>
        /// Gets or sets the list.
        /// </summary>
        /// <value>The list.</value>
        [Parameter(ParameterSetName = "SPList",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The source list containing the content type to export.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        [Parameter(Mandatory = false, ParameterSetName = "SPContentType2",
            HelpMessage = "The source list containing the content type to export.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPListPipeBind[] List { get; set; }

        /// <summary>
        /// Gets or sets the name of the content type.
        /// </summary>
        /// <value>The name of the contentType.</value>
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = false, ParameterSetName = "SPContentType1", HelpMessage = "The name of the content type to export. The type must be a valid content type name; a valid content type ID, in the form 0x0123...; or an instance of an SPContentType object.")]
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = false, ParameterSetName = "SPContentType2", HelpMessage = "The name of the content type to export. The type must be a valid content type name; a valid content type ID, in the form 0x0123...; or an instance of an SPContentType object.")]
        public SPContentTypePipeBind Identity { get; set; }

        [Parameter(Mandatory = false, Position = 0, ValueFromPipeline = false, ParameterSetName = "SPWeb", HelpMessage = "The name of the content type to export.")]
        [Parameter(Mandatory = false, Position = 0, ValueFromPipeline = false, ParameterSetName = "SPList", HelpMessage = "The name of the content type to export.")]
        [Alias("Name")]
        public string ContentType { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "SPWeb", HelpMessage = "The content type group to export. Specifying this will filter the exported content types to just those items in the group.")]
        public string Group { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "SPWeb", HelpMessage = "The list name to filter by.\r\n\r\nSpecifying this along with the Web parameter will export content types from the list with the specified name only.")]
        public string ListName { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Exclude all fields defined in parent content types.")]
        public SwitchParameter ExcludeParentFields { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Include the field definitions in the exported XML.")]
        public SwitchParameter IncludeFieldDefinitions { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Include the list bindings in the exported XML.")]
        public SwitchParameter IncludeListBindings { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Remove encoded spaces (_x0020_) in all field and content type names.")]
        public SwitchParameter RemoveEncodedSpaces { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Remove any elements/attributes which are invalid in a SharePoint Solution Package Feature manifest file.")]
        public SwitchParameter FeatureSafe { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The path to the output file where the exported XML will be saved.\r\n\r\nIf omitted the output will be dumped to the pipeline.")]
        [ValidateDirectoryExistsAndValidFileName]
        [Alias("Path")]
        public string OutputFile { get; set; }

        StringBuilder _sb;
        XmlTextWriter _xmlWriter;

        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();

            _sb = new StringBuilder();
            _xmlWriter = Common.ContentTypes.ExportContentTypes.OpenXmlWriter(_sb);
        }

        protected override void InternalEndProcessing()
        {
            base.InternalEndProcessing();

            Common.ContentTypes.ExportContentTypes.CloseXmlWriter(_xmlWriter, OutputFile, _sb);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(_sb.ToString());
            WriteResult(xmlDoc);
        }

        protected override void InternalProcessRecord()
        {
            List<SPContentType> contentTypes = new List<SPContentType>();
            List<SPWeb> webs = new List<SPWeb>();
            List<SPList> lists = new List<SPList>();

            try
            {
                switch (ParameterSetName)
                {
                    case "SPContentType1":
                        if (Web != null)
                        {
                            foreach (SPWebPipeBind webPipe in Web)
                            {
                                SPWeb web = webPipe.Read();
                                webs.Add(web);
                                WriteVerbose("Getting content type from " + web.Url);
                                try
                                {
                                    contentTypes.Add(Identity.Read(web));
                                }
                                catch (ArgumentException)
                                {
                                    WriteWarning("Could not locate the specified content type at " + web.Url);
                                }
                            }
                        }
                        else
                        {
                            SPContentType ct = Identity.Read();
                            contentTypes.Add(ct);
                        }
                        foreach (SPContentType ct in contentTypes)
                        {
                            Common.ContentTypes.ExportContentTypes.Export(ct.ParentWeb.Url, null, ct.Name,
                                ExcludeParentFields.IsPresent, IncludeFieldDefinitions.IsPresent, IncludeListBindings.IsPresent,
                                null, RemoveEncodedSpaces.IsPresent, FeatureSafe.IsPresent, _xmlWriter);
                        }
                        break;
                    case "SPContentType2":
                        if (List != null)
                        {
                            foreach (SPListPipeBind listPipe in List)
                            {
                                SPList list = listPipe.Read();
                                lists.Add(list);
                                WriteVerbose("Getting content type from " + list.RootFolder.Url);
                                try
                                {
                                    contentTypes.Add(Identity.Read(list));
                                }
                                catch (ArgumentException)
                                {
                                    WriteWarning("Could not locate the specified content type at " + list.RootFolder.Url);
                                }
                            }
                        }
                        else
                        {
                            SPContentType ct = Identity.Read();
                            contentTypes.Add(ct);
                        }
                        foreach (SPContentType ct in contentTypes)
                        {
                            Common.ContentTypes.ExportContentTypes.Export(ct.ParentWeb.Url, null, ct.Name,
                                ExcludeParentFields.IsPresent, IncludeFieldDefinitions.IsPresent, IncludeListBindings.IsPresent,
                                ct.ParentList.Title, RemoveEncodedSpaces.IsPresent, FeatureSafe.IsPresent, _xmlWriter);
                        }
                        break;
                    case "SPWeb":
                        foreach (SPWebPipeBind webPipe in Web)
                        {
                            SPWeb web = webPipe.Read();
                            webs.Add(web);

                            Common.ContentTypes.ExportContentTypes.Export(web, Group, ContentType,
                                ExcludeParentFields.IsPresent, IncludeFieldDefinitions.IsPresent, IncludeListBindings.IsPresent,
                                ListName, RemoveEncodedSpaces.IsPresent, FeatureSafe.IsPresent, _xmlWriter);
                        }
                        break;
                    case "SPList":
                        foreach (SPListPipeBind listPipe in List)
                        {
                            SPList list = listPipe.Read();
                            lists.Add(list);

                            Common.ContentTypes.ExportContentTypes.Export(list.ParentWeb.Url, null, ContentType,
                                ExcludeParentFields.IsPresent, IncludeFieldDefinitions.IsPresent, IncludeListBindings.IsPresent,
                                list.Title, RemoveEncodedSpaces.IsPresent, FeatureSafe.IsPresent, _xmlWriter);
                        }
                        break;
                }
            }
            finally
            {

                foreach (SPWeb web in webs)
                {
                    web.Site.Dispose();
                    web.Dispose();
                }
                foreach (SPList list in lists)
                {
                    list.ParentWeb.Site.Dispose();
                    list.ParentWeb.Dispose();
                }
                foreach (SPContentType ct in contentTypes)
                {
                    if (ct.ParentList != null)
                    {
                        ct.ParentList.ParentWeb.Site.Dispose();
                        ct.ParentList.ParentWeb.Dispose();
                    }
                    if (ct.ParentWeb != null)
                    {
                        ct.ParentWeb.Site.Dispose();
                        ct.ParentWeb.Dispose();
                    }
                }
            }
        }
    }
}
