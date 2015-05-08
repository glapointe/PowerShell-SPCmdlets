using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.SharePoint;
/*
namespace System.Runtime.CompilerServices
{
    /// <summary>
    /// Enable C# 3.0 extensions.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method)]
    public sealed class ExtensionAttribute : Attribute { }
}
*/
namespace Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers
{
    public static class ExtensionMethods
    {/*
        public static SPFolder GetFolder(this SPList targetList, string folderUrl)
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
        */
    }
}
