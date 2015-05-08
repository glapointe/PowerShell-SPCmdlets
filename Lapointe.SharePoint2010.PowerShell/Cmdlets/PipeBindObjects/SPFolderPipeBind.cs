using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPFolderPipeBind : SPCmdletPipeBind<SPFolder>
    {
        private string _folderUrl;

        public SPFolderPipeBind(SPFolder folder)
        {
            _folderUrl = folder.ParentWeb.Site.MakeFullUrl(folder.ServerRelativeUrl);
        }
        public SPFolderPipeBind(SPFile file) : this(file.ParentFolder) { }
        public SPFolderPipeBind(SPListItem item) : this(item.Folder) { }
        public SPFolderPipeBind(SPList list) : this(list.RootFolder) { }

        public SPFolderPipeBind(string folder)
        {
            _folderUrl = folder;
        }

        protected override void Discover(SPFolder instance)
        {
            _folderUrl = instance.ParentWeb.Site.MakeFullUrl(instance.ServerRelativeUrl);
        }
        public override SPFolder Read()
        {
            if (string.IsNullOrEmpty(_folderUrl)) throw new SPCmdletPipeBindException("The folder path was not specified.");

            using (SPSite site = new SPSite(_folderUrl))
            using (SPWeb web = site.OpenWeb())
            {
                SPFolder folder = web.GetFolder(_folderUrl);
                if (!folder.Exists) throw new SPCmdletPipeBindException("The specified folder does not exist: " + _folderUrl);
                return folder;
            }
        }

    }
}
