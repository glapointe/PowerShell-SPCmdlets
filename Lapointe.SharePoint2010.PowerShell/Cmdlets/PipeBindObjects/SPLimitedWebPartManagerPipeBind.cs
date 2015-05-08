using System;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using System.Globalization;
using Microsoft.SharePoint.WebPartPages;
using System.Web.UI.WebControls.WebParts;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPLimitedWebPartManagerPipeBind : SPCmdletPipeBind<SPLimitedWebPartManager>
    {
        private string _pageUrl;

        public SPLimitedWebPartManagerPipeBind(SPLimitedWebPartManager instance) : base(instance)
        {
            this._pageUrl = instance.Web.Site.MakeFullUrl(instance.ServerRelativeUrl);
        }

        public SPLimitedWebPartManagerPipeBind(string inputString)
        {
            this._pageUrl = inputString.Trim();
        }

        public SPLimitedWebPartManagerPipeBind(Uri pageUri)
        {
            this._pageUrl = pageUri.ToString();
        }

        protected override void Discover(SPLimitedWebPartManager instance)
        {
            this._pageUrl = instance.Web.Site.MakeFullUrl(instance.ServerRelativeUrl);
        }

        public override SPLimitedWebPartManager Read()
        {
            // We don't dispose here as we'll add these objects
            // to the SPAssignmentCollection
            SPSite site = new SPSite(PageUrl);
            SPWeb web = site.OpenWeb();
            SPLimitedWebPartManager mgr = web.GetLimitedWebPartManager(PageUrl, PersonalizationScope.Shared);

            if (mgr == null)
            {
                web.Dispose();
                site.Dispose();
                throw new SPCmdletPipeBindException(string.Format("SPLimitedWebPartManager PipeBind object not found ({0})", PageUrl));
            }
            return mgr;
        }

        public string PageUrl
        {
            get { return this._pageUrl; }
        }
    }

}
