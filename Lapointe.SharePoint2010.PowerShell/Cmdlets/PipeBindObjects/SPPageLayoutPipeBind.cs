using System;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using System.Globalization;
using Microsoft.SharePoint.WebPartPages;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint.Publishing;
using System.IO;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPPageLayoutPipeBind : SPCmdletPipeBind<PageLayout>
    {
        private bool _isAbsoluteUrl;
        private string _fileUrl;

        public SPPageLayoutPipeBind(PageLayout instance)
            : base(instance)
        {
            this._fileUrl = instance.ListItem.Web.Site.MakeFullUrl(instance.ServerRelativeUrl);
            _isAbsoluteUrl = true;
        }

        public SPPageLayoutPipeBind(string inputString)
        {
            if (string.IsNullOrEmpty(inputString))
                throw new ArgumentException("Input string cannot be null or empty.", "inputString");

            this._fileUrl = inputString.Trim();
            if (this._fileUrl.StartsWith("http", true, CultureInfo.CurrentCulture) && this._fileUrl.Contains("://"))
                this._isAbsoluteUrl = true;

        }

        public SPPageLayoutPipeBind(Uri fileUri)
        {
            this._fileUrl = fileUri.ToString();
        }

        public SPPageLayoutPipeBind(SPFile file)
        {
            this._fileUrl = file.Web.Site.MakeFullUrl(file.ServerRelativeUrl);
        }

        protected override void Discover(PageLayout instance)
        {
            this._fileUrl = instance.ListItem.Web.Site.MakeFullUrl(instance.ServerRelativeUrl);
        }

        public override PageLayout Read()
        {
            return Read((PublishingWeb)null);
        }

        public PageLayout Read(SPWeb web)
        {
            if (web == null)
                return Read((PublishingWeb)null);

            if (!PublishingWeb.IsPublishingWeb(web))
                throw new ArgumentException("The specified web is not a publishing web.");
            PublishingWeb pubWeb = PublishingWeb.GetPublishingWeb(web);
            return Read(pubWeb);
        }

        public PageLayout Read(PublishingWeb pubWeb)
        {
            // We don't dispose here as we'll add these objects
            // to the SPAssignmentCollection
            SPSite site = null;
            SPWeb web = null;

            if (pubWeb == null && _isAbsoluteUrl)
            {
                site = new SPSite(_fileUrl);
                web = site.OpenWeb();
                if (!PublishingWeb.IsPublishingWeb(web))
                    throw new ArgumentException("The specified web is not a publishing web.");
                pubWeb = PublishingWeb.GetPublishingWeb(web);
            }
            else if (pubWeb == null && !_isAbsoluteUrl)
            {
                throw new FileNotFoundException("The specified layout could not be found.", _fileUrl);
            }
            else
            {
                web = pubWeb.Web;
            }

            PageLayout file = null;
            if (_isAbsoluteUrl)
            {
                foreach (PageLayout lo in pubWeb.GetAvailablePageLayouts())
                {
                    if (lo.ListItem.Web.Site.MakeFullUrl(lo.ServerRelativeUrl.ToLowerInvariant()) == _fileUrl.ToLowerInvariant())
                    {
                        file = lo;
                        break;
                    }
                }
            }
            else
            {
                foreach (PageLayout lo in pubWeb.GetAvailablePageLayouts())
                {
                    if (lo.Name.ToLowerInvariant() == _fileUrl.ToLowerInvariant())
                    {
                        file = lo;
                        break;
                    }
                }
            }

            if (file == null)
            {
                if (web != null)
                    web.Dispose();
                if (site != null)
                    site.Dispose();
                throw new SPCmdletPipeBindException(string.Format("SPPageLayoutPipeBind object not found ({0})", _fileUrl));
            }
            return file;
        }

        public string FileUrl
        {
            get { return this._fileUrl; }
        }
    }

}
