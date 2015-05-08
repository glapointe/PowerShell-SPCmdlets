using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Deployment;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    internal static class DeleteList
    {

        internal static void Delete(bool force, string backupDir, SPList list)
        {
            if (!string.IsNullOrEmpty(backupDir))
            {
                string path = Path.Combine(backupDir, list.RootFolder.Name.Replace(" ", "_"));
                int i = 0;
                while (Directory.Exists(path + i))
                    i++;
                path += i;
                Directory.CreateDirectory(path);
                try
                {
                    using (SPSite site = new SPSite(list.ParentWeb.Site.ID))
                    {
                        Common.Lists.ExportList.PerformExport(
                            site.MakeFullUrl(list.DefaultViewUrl),
                            Path.Combine(path, list.RootFolder.Name),
                            true, false, false, true, 0, true, false,
                            false, SPIncludeVersions.All, SPIncludeDescendants.All, true, false, false);
                    }
                }
                catch (Exception ex)
                {
                    throw new SPException("Unable to backup list.  List not deleted.", ex);
                }
            }
            Delete(list, force);
        }

        /// <summary>
        /// Deletes the specified list.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        internal static void Delete(SPList list, bool force)
        {
            if (list == null)
                throw new SPException("List not found.");

            if (!list.AllowDeletion && force)
            {
                list.AllowDeletion = true;
                list.Update();
            }
            else if (!list.AllowDeletion)
                throw new SPException("List cannot be deleted.  Try using the '-force' parameter to force the delete.");

            try
            {
                list.Delete();
            }
            catch (Exception)
            {
                if (force)
                {
                    using (SPSite site = new SPSite(list.ParentWeb.Site.ID))
                    {
                        Utilities.RunStsAdmOperation(
                            string.Format(" -o forcedeletelist -url \"{0}\"",
                                          site.MakeFullUrl(list.RootFolder.ServerRelativeUrl)), false);
                    }
                }
            }
        }
    }
}
