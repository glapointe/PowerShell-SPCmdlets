using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using System.Collections;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.New, "SPFile", SupportsShouldProcess = false, DefaultParameterSetName = "List"),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Creates a new file within a list.")]
    [RelatedCmdlets(typeof(SPCmdletGetList))]
    [Example(Code = "PS C:\\> New-SPFile -List \"http://server_name/documents\" -File \"c:\\myfile.txt\" -FieldValues @{\"Title\"=\"My new file\"}",
         Remarks = "This example creates a new file within the List My List under the root Site of the current Site Collection.")]
    public class SPCmdletNewFile : SPNewCmdletBaseCustom<SPFile>
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "List",
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the List to add the file to.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPListPipeBind List { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Folder",
            Position = 2,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the Folder to add the file to.\r\n\r\nThe type must be a valid SPOFolder object.")]
        public SPFolderPipeBind Folder { get; set; }

        [ValidateNotNullOrEmpty,
        ValidateFileExists,
        Parameter(
            Position = 3,
            Mandatory = true,
            HelpMessage = "Specify the path to the file to add to the list.")]
        public string File { get; set; }

        [Parameter(Position = 4, HelpMessage = "Overwrite an existing file if present.")]
        public SwitchParameter Overwrite { get; set; }

        [Parameter(
            Position = 5,
            Mandatory = false,
            HelpMessage = "The collection of field values to set where the key is the internal field name. The type must be a hash table where each key represents the name of a field whose value should be set to the corresponding key value (e.g., @{\"Field1\"=\"Value1\";\"Field2\"=\"Value2\"}). Alternatively, provide the path to a file with XML property settings (<Properties><Property Name=\"Name1\">Value1</Property><Property Name=\"Name2\">Value2</Property></Properties>).")]
        public PropertiesPipeBind FieldValues { get; set; }


        protected override SPFile CreateDataObject()
        {
            Hashtable fieldValues = null;
            if (FieldValues != null)
                fieldValues = FieldValues.Read();

            byte[] content = System.IO.File.ReadAllBytes(File);
            System.IO.FileInfo file = new System.IO.FileInfo(File);
            string leafName = file.Name;

            if (ParameterSetName == "List")
            {
                SPList list = List.Read();
                return AddFile(list.RootFolder, content, Overwrite, leafName, fieldValues);
            }
            else
            {
                SPFolder folder = Folder.Read();
                return AddFile(folder, content, Overwrite, leafName, fieldValues);
            }
        }

        private SPFile AddFile(SPFolder folder, byte[] content, bool overwrite, string leafName, Hashtable fieldValues)
        {
            return folder.Files.Add(leafName, content, fieldValues, overwrite);
        }

    }


}
