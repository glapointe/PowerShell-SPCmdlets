using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Taxonomy;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPTaxonomySessionPipeBind : SPCmdletPipeBind<TaxonomySession>
    {
        private Guid _siteGuid;
        private string _siteUrl;
        private TaxonomySession _taxonomySession;

        public SPTaxonomySessionPipeBind(TaxonomySession instance)
            : base(instance)
        {
        }

        public SPTaxonomySessionPipeBind(SPSite site)
        {
            if (site == null)
            {
                throw new ArgumentNullException("site");
            }
            _siteGuid = site.ID;
        }

        public SPTaxonomySessionPipeBind(Guid guid)
        {
            this._siteGuid = guid;
        }

        public SPTaxonomySessionPipeBind(string inputString)
        {
            if (inputString != null)
            {
                inputString = inputString.Trim();
                try
                {
                    this._siteGuid = new Guid(inputString);
                }
                catch (FormatException)
                {
                }
                catch (OverflowException)
                {
                }
                if (this._siteGuid.Equals(Guid.Empty))
                {
                    this._siteUrl = inputString;
                }
            }
        }

        public SPTaxonomySessionPipeBind(Uri uri)
        {
            this._siteUrl = uri.ToString();
        }

        public SPTaxonomySessionPipeBind(SPSiteAdministration inputObject)
        {
            this._siteUrl = inputObject.Url;
        }

        protected override void Discover(TaxonomySession instance)
        {
            _taxonomySession = instance;
            try
            {
                var taxSessionContext = Utilities.GetPropertyValue(instance, "Context");
                SPSite site = Utilities.GetPropertyValue(taxSessionContext, "SiteOrNull") as SPSite;
                if (site != null)
                {
                    _taxonomySession = null;
                    _siteGuid = site.ID;
                }
            }
            catch {}
        }

        public override TaxonomySession Read()
        {
            if (_taxonomySession != null)
                return _taxonomySession;

            SPSite site = null;
            try
            {
                if (Guid.Empty != this._siteGuid)
                {
                    site = new SPSite(this._siteGuid);
                }
                else if (!string.IsNullOrEmpty(this._siteUrl))
                {
                    site = new SPSite(this._siteUrl);
                    string serverRelUrlFromFullUrl = Utilities.GetServerRelUrlFromFullUrl(this._siteUrl);
                    if (!site.ServerRelativeUrl.Equals(serverRelUrlFromFullUrl, StringComparison.OrdinalIgnoreCase))
                    {
                        site.Dispose();
                        site = null;
                    }
                }
            }
            catch (Exception exception)
            {
                throw new SPCmdletPipeBindException("The SPSite object was not found.", exception);
            }
            if (site != null)
            {
                _taxonomySession = new TaxonomySession(site, true);
                site.Dispose();
                return _taxonomySession;
            }
            throw new SPCmdletPipeBindException("Could not create TaxonomySession object.");
        }
    }
}
