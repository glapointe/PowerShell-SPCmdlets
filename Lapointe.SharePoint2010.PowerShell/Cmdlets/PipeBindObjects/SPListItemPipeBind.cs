using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPListItemPipeBind : SPCmdletPipeBind<SPListItem>
    {
        private string _fileUrl;
        private string _listUrl;
        private int? _itemId;

        public SPListItemPipeBind(SPFile instance)
            : base(instance.Item)
        {
            _listUrl = instance.Item.ParentList.RootFolder.Url;
            _itemId = instance.Item.ID;
        }

        public SPListItemPipeBind(SPListItem instance)
            : base(instance)
        {
            Discover(instance);
        }

        public SPListItemPipeBind(int itemId)
        {
            _itemId = itemId;
        }

        public SPListItemPipeBind(string inputString)
        {
            _fileUrl = inputString.Trim();
        }

        public SPListItemPipeBind(Uri fileUri)
        {
            _fileUrl = fileUri.ToString();
        }

        protected override void Discover(SPListItem instance)
        {
            _listUrl = instance.ParentList.ParentWeb.Site.MakeFullUrl(instance.ParentList.RootFolder.ServerRelativeUrl);
            _itemId = instance.ID;
        }

        public SPListItem Read(SPList list)
        {
            if (_itemId.HasValue)
                return list.GetItemById(_itemId.Value);
            throw new SPCmdletPipeBindException("An item ID was not specified.");
        }

        public override SPListItem Read()
        {
            // We don't dispose here as we'll add these objects
            // to the SPAssignmentCollection

            if (!string.IsNullOrEmpty(_fileUrl))
            {
                SPSite site = new SPSite(_fileUrl);
                SPWeb web = site.OpenWeb();
                SPFile file = web.GetFile(_fileUrl);

                if (file == null)
                {
                    web.Dispose();
                    site.Dispose();
                    throw new SPCmdletPipeBindException(string.Format("SPListItem PipeBind object not found ({0})", _fileUrl));
                }
                return file.Item;
            }
            if (string.IsNullOrEmpty(_listUrl))
                throw new SPCmdletPipeBindException("The list URL was not specified.");
            if (!_itemId.HasValue)
                throw new SPCmdletPipeBindException("The list ID was not specified.");

            SPListPipeBind lp = new SPListPipeBind(_listUrl);
            SPList list = lp.Read();

            if (list == null)
                throw new SPCmdletPipeBindException(string.Format("SPList PipeBind object not found ({0})", _listUrl));

            return list.GetItemById(_itemId.Value);
        }

    }

}
