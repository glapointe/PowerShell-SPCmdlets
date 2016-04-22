using System;
using System.Collections.Generic;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint.Publishing;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using System.IO;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Download", "SPFile", SupportsShouldProcess = false, DefaultParameterSetName = "File"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Save the file associated with the specified file or folder URL.")]
    [RelatedCmdlets(typeof(SPCmdletGetFile))]
    [Example(Code = "PS C:\\> Download-SPFile -File \"http://server_name/pages/default.aspx\" -TargetFolder \"c:\\temp\"",
        Remarks = "This saves the default.aspx file from the http://server_name/pages library.")]
    [Example(Code = "PS C:\\> Download-SPFile \"http://server_name/documents/\" -TargetFolder \"c:\\temp\" -Recursive",
        Remarks = "This example saves all files from the http://server_name/documents library.")]
    public class SPCmdletDownloadFile : SPCmdletCustom
    {
        [Parameter(Mandatory = true, 
            ParameterSetName = "File",
            Position = 0, 
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The path to the file to save.")]
        [ValidateNotNull]
        public SPFilePipeBind[] File { get; set; }

        [Parameter(Mandatory = true, 
            ParameterSetName = "Folder",
            Position = 0, 
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The path to the folder to save.")]
        [ValidateNotNull]
        public SPFolderPipeBind Folder { get; set; }

        [Parameter(ParameterSetName = "Folder",
            Position = 1,
            HelpMessage = "Save all child folders and their files.")]
        public SwitchParameter Recursive { get; set; }

        [Parameter(Mandatory = true,
            Position = 2,
            HelpMessage = "The path to the folder to save the files to.")]
        [ValidateDirectoryExists]
        [Alias("Path")]
        public string TargetFolder { get; set; }

        [Parameter(Position = 3, HelpMessage = "Overwrite existing files.")]
        public SwitchParameter Overwrite { get; set; }

        protected override void InternalProcessRecord()
        {
 	        base.InternalProcessRecord();
            if (ParameterSetName == "File")
            {
                foreach (SPFilePipeBind filePipe in File)
                {
                    SPFile file = filePipe.Read();
                    WriteFile(TargetFolder, file);
                    file.Web.Dispose();
                    file.Web.Site.Dispose();
                }
            }
            else if (ParameterSetName == "Folder")
            {
                var folder = Folder.Read();
                WriteFiles(TargetFolder, folder);
                folder.ParentWeb.Dispose();
                folder.ParentWeb.Site.Dispose();
            }
        }
        private void WriteFiles(string targetFolder, SPFolder folder)
        {
            if (!Directory.Exists(targetFolder))
            {
                WriteVerbose(string.Format("Creating folder {0}...", targetFolder));
                Directory.CreateDirectory(targetFolder);
            }

            foreach (SPFile file in folder.Files)
            {
                WriteFile(targetFolder, file);
            }
            if (Recursive)
            {
                foreach (SPFolder childFolder in folder.SubFolders)
                {
                    WriteFiles(Path.Combine(targetFolder, childFolder.Name), childFolder);
                }
            }
        }

        private void WriteFile(string targetFolder, SPFile file)
        {
            string fullName = Path.Combine(targetFolder, file.Name);
            if (System.IO.File.Exists(fullName) && !Overwrite)
            {
                WriteWarning(string.Format("Unable to save \"{0}\". File already exists. Use the -Overwrite parameter to overwrite.", fullName));
                return;
            }
            WriteVerbose(string.Format("Saving {0} to {1}...", file.Name, targetFolder));
            System.IO.File.WriteAllBytes(fullName, file.OpenBinary());
        }
    }
}
