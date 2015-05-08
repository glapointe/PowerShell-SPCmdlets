using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.New, "SPFolder", SupportsShouldProcess = false),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Creates a new folder within a list.")]
    [RelatedCmdlets(typeof(SPCmdletGetList))]
    [Example(Code = "PS C:\\> $folder = Get-SPList http://server_name/documents | New-SPFolder -Path \"TopLevelFolder/ChildFolder1\" ",
        Remarks = "This example creates a TopLevelFolder with a sub-folder called ChildFolder1 at http://server_name/documents.")]
    public class SPCmdletNewFolder : SPNewCmdletBaseCustom<SPFolder>
    {
        /// <summary>
        /// Gets or sets the list.
        /// </summary>
        /// <value>The list.</value>
        [Parameter(Mandatory = true,
            Position = 0,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The list to add the field to.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        public SPListPipeBind List { get; set; }


        [Parameter(Mandatory = true,
            HelpMessage = "The list relative folder path to create to create. Example: \"TopLevelFolder/ChildFolder1\"")]
        [ValidateNotNullOrEmpty]
        public string Path { get; set; }


        protected override SPFolder CreateDataObject()
        {
            SPList list = List.Read();
            if (list != null)
                return GetFolder(list, Path);

            return null;
        }

        public static SPFolder GetFolder(SPList targetList, string folderUrl)
        {
            if (string.IsNullOrEmpty(folderUrl))
                return targetList.RootFolder;

            SPFolder folder = targetList.ParentWeb.GetFolder(targetList.RootFolder.Url + "/" + folderUrl);

            if (!folder.Exists)
            {
                if (!targetList.EnableFolderCreation)
                {
                    targetList.EnableFolderCreation = true;
                    targetList.Update();
                }

                // We couldn't find the folder so create it
                string[] folders = folderUrl.Trim('/').Split('/');

                string folderPath = string.Empty;
                for (int i = 0; i < folders.Length; i++)
                {
                    folderPath += "/" + folders[i];
                    folder = targetList.ParentWeb.GetFolder(targetList.RootFolder.Url + folderPath);
                    if (!folder.Exists)
                    {
                        SPListItem newFolder = targetList.Items.Add("", SPFileSystemObjectType.Folder, folderPath.Trim('/'));
                        newFolder.Update();
                        folder = newFolder.Folder;
                    }
                }
            }
            // Still no folder so error out
            if (folder == null)
                throw new SPException(string.Format("The folder '{0}' could not be found.", folderUrl));
            return folder;
        }
    }


}
