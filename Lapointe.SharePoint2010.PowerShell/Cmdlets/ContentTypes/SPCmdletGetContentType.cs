using System;
using System.Collections.Generic;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.ContentTypes
{
    [Cmdlet(VerbsCommon.Get, "SPContentType", SupportsShouldProcess = false, DefaultParameterSetName = "SPWeb"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = false)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Retrieve an SPContentType object.")]
    [RelatedCmdlets(typeof(SPCmdletCopyContentType), typeof(SPCmdletExportContentType), typeof(SPCmdletPropagateContentType),
        typeof(Lists.SPCmdletGetList), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> $ct = Get-SPWeb \"http://server_name\" | Get-SPContentType -Identity \"ContentType1\"",
        Remarks = "This example retrieves ContentType1 from http://server_name.")]
    [Example(Code = "PS C:\\> $ct = Get-SPWeb \"http://server_name\" | Get-SPContentType",
        Remarks = "This example retrieves all content types from http://server_name.")]
    [Example(Code = "PS C:\\> $ct = Get-SPList \"http://server_name/lists/mylist\" | Get-SPContentType",
        Remarks = "This example retrieves all content types from http://server_name.")]
    public class SPCmdletGetContentType : SPGetCmdletBaseCustom<SPContentType>
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The source web containing the content type to retrieve.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        public SPWebPipeBind[] Web { get; set; }

        /// <summary>
        /// Gets or sets the list.
        /// </summary>
        /// <value>The list.</value>
        [Parameter(ParameterSetName = "SPList",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The source list containing the content type to retrieve.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        public SPListPipeBind[] List { get; set; }

        /// <summary>
        /// Gets or sets the name of the content type.
        /// </summary>
        /// <value>The name of the contentType.</value>
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = false,
            HelpMessage = "The name of the content type to return. The type must be a valid content type name; a valid content type ID, in the form 0x0123...; or an instance of an SPContentType object."),
        Alias("ContentType", "Name")]
        [ValidateNotNull]
        public SPContentTypePipeBind Identity { get; set; }

        protected override IEnumerable<SPContentType> RetrieveDataObjects()
        {
            List<SPContentType> contentTypes = new List<SPContentType>();

            switch (ParameterSetName)
            {
                case "SPWeb":
                    foreach (SPWebPipeBind webPipe in Web)
                    {
                        SPWeb web = webPipe.Read();
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
                    break;
                case "SPList":
                    foreach (SPListPipeBind listPipe in List)
                    {
                        SPList list = listPipe.Read();
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
                    break;
            }
            if (contentTypes.Count == 0)
                WriteWarning("No content types were found matching the specified name at the provided location.");

            foreach (SPContentType ct in contentTypes)
                WriteResult(ct);

            return null;
        }
    }
}
