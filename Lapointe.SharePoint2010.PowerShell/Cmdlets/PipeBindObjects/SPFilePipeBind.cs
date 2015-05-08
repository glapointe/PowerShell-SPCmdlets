using System;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using System.Globalization;
using Microsoft.SharePoint.Publishing;
using Microsoft.SharePoint.WebPartPages;
using System.Web.UI.WebControls.WebParts;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPFilePipeBind : SPCmdletPipeBind<SPFile>
    {
        private string _fileUrl;

        public SPFilePipeBind(SPFile instance)
            : base(instance)
        {
            this._fileUrl = instance.Web.Site.MakeFullUrl(instance.ServerRelativeUrl);
        }
        public SPFilePipeBind(PublishingPage page) : this(page.ListItem.File)
        {
        }
        public SPFilePipeBind(string inputString)
        {
            this._fileUrl = inputString.Trim();
        }

        public SPFilePipeBind(Uri fileUri)
        {
            this._fileUrl = fileUri.ToString();
        }

        protected override void Discover(SPFile instance)
        {
            this._fileUrl = instance.Web.Site.MakeFullUrl(instance.ServerRelativeUrl);
        }

        public override SPFile Read()
        {
            // We don't dispose here as we'll add these objects
            // to the SPAssignmentCollection
            SPSite site = new SPSite(FileUrl);
            SPWeb web = site.OpenWeb();
            SPFile file = web.GetFile(FileUrl);

            if (file == null)
            {
                web.Dispose();
                site.Dispose();
                throw new SPCmdletPipeBindException(string.Format("SPFile PipeBind object not found ({0})", FileUrl));
            }
            return file;
        }

        public string FileUrl
        {
            get { return this._fileUrl; }
        }
    }

}
