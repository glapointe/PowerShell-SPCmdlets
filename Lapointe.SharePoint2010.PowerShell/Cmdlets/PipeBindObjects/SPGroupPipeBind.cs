using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public class SPGroupPipeBind : SPCmdletPipeBind<SPGroup>
    {
        // Fields
        private int? _id;
        private string _groupName;
        private SPWeb _web;

        // Methods
        public SPGroupPipeBind(SPGroup group)
            : base(group)
        {
        }

        public SPGroupPipeBind(int id)
        {
            _id = id;
        }
        public SPGroupPipeBind(string groupName)
        {
            if (string.IsNullOrEmpty(groupName))
            {
                throw new ArgumentNullException("groupName");
            }
            _groupName = groupName;
        }

        protected override void Discover(SPGroup instance)
        {
            if (instance != null)
            {
                _web = instance.ParentWeb;
                _groupName = instance.Name;
                _id = instance.ID;
            }
        }

        public override SPGroup Read()
        {
            if (_web == null)
            {
                throw new InvalidOperationException("A valid SPWeb object must be provided to retrieve the specified group.");
            }
            return this.Read(_web);
        }

        public SPGroup Read(SPWeb web)
        {
            if (web == null)
            {
                throw new InvalidOperationException("A valid SPWeb object must be provided to retrieve the specified group.");
            }

            if (_id.HasValue)
            {
                try
                {
                    return web.SiteGroups.GetByID(_id.Value);
                }
                catch (Exception) {}
            }
            if (!string.IsNullOrEmpty(_groupName))
            {
                try
                {
                    return web.SiteGroups[_groupName];
                }
                catch (Exception) { }
                return null;
            }
            return null;
        }

    }
}
