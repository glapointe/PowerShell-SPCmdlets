using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.Pages
{
    internal class EnumUnGhostedFiles
    {
        /// <summary>
        /// Recurses the sub webs.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="unghostedFiles">The unghosted files.</param>
        internal static void RecurseSubWebs(SPWeb web, ref List<object> unghostedFiles, bool asString)
        {
            foreach (SPWeb subweb in web.Webs)
            {
                try
                {
                    RecurseSubWebs(subweb, ref unghostedFiles, asString);
                }
                finally
                {
                    subweb.Dispose();
                }
            }
            CheckFoldersForUnghostedFiles(web.RootFolder, ref unghostedFiles, asString);
        }

        /// <summary>
        /// Checks the folders for unghosted files.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="unghostedFiles">The unghosted files.</param>
        internal static void CheckFoldersForUnghostedFiles(SPFolder folder, ref List<object> unghostedFiles, bool asString)
        {
            foreach (SPFolder sub in folder.SubFolders)
            {
                CheckFoldersForUnghostedFiles(sub, ref unghostedFiles, asString);
            }

            foreach (SPFile file in folder.Files)
            {
                if (file.CustomizedPageStatus == SPCustomizedPageStatus.Customized)
                {
                    if (asString)
                    {
                        string url = file.Web.Site.MakeFullUrl(file.ServerRelativeUrl);
                        if (!unghostedFiles.Contains(url))
                            unghostedFiles.Add(url);
                    }
                    else
                    {
                        if (!unghostedFiles.Contains(file))
                            unghostedFiles.Add(file);
                    }
                }
            }
        }
    }
}
