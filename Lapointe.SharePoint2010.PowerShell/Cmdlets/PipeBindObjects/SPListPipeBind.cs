using System;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using System.Globalization;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPListPipeBind : SPCmdletPipeBind<SPList>
    {
        private bool _isAbsoluteUrl;
        private bool _isCollection;
        private Guid _siteGuid;
        private Guid _webGuid;
        private Guid _listGuid;
        private string _listUrl;

        public SPListPipeBind(SPList instance) : base(instance)
        { }

        public SPListPipeBind(Guid guid)
        {
            this._listGuid = guid;
        }

        public SPListPipeBind(string inputString)
        {
            if (inputString != null)
            {
                inputString = inputString.Trim();
                try
                {
                    this._listGuid = new Guid(inputString);
                }
                catch (FormatException) { }
                catch (OverflowException) { }

                if (this._listGuid.Equals(Guid.Empty))
                {
                    this._listUrl = inputString;
                    if (this._listUrl.StartsWith("http", true, CultureInfo.CurrentCulture) && _listUrl.Contains("://"))
                        this._isAbsoluteUrl = true;

                    if (WildcardPattern.ContainsWildcardCharacters(this._listUrl))
                        this._isCollection = true;
                }
            }
        }

        public SPListPipeBind(Uri listUri)
        {
            this._listUrl = listUri.ToString();
        }

        protected override void Discover(SPList instance)
        {
            this._listGuid = instance.ID;
            this._webGuid = instance.ParentWeb.ID;
            this._siteGuid = instance.ParentWeb.Site.ID;
        }

        public override SPList Read()
        {
            return this.Read(null);
        }

        public SPList Read(SPWeb web)
        {
            SPList list = null;
            string parameterDetails = string.Format(CultureInfo.CurrentCulture, "Id or Url : {0}", new object[] { "Empty or Null" });
            if (this.IsCollection)
            {
                // Not currently supporting wildcards so no collections.
                return null;
            }
            try
            {
                if (Guid.Empty != this.ListGuid)
                {
                    if (web == null && Guid.Empty != this._webGuid && Guid.Empty != this._siteGuid)
                    {
                        parameterDetails = string.Format(CultureInfo.CurrentCulture, "Id or Url: {0} and Web Id: {1}", new object[] { this.ListGuid.ToString(), this._webGuid.ToString() });
                        using (SPSite site = new SPSite(this._siteGuid))
                        {
                            web = site.OpenWeb(this._webGuid);
                            list = web.Lists[ListGuid];
                        }
                    }
                    else
                    {
                        parameterDetails = string.Format(CultureInfo.CurrentCulture, "Id or Url: {0} and web Url {1}", new object[] { this.ListUrl, web.Url });
                        list = web.Lists[ListGuid];
                    }
                }
                else if (!string.IsNullOrEmpty(this.ListUrl))
                {
                    string serverRelativeListUrl = null;
                    if (this._isAbsoluteUrl)
                        serverRelativeListUrl = Utilities.GetServerRelUrlFromFullUrl(this.ListUrl).Trim('/');
                    else
                        serverRelativeListUrl = this.ListUrl.Trim('/');

                    if (web == null)
                    {
                        parameterDetails = string.Format(CultureInfo.CurrentCulture, "Id or Url : {0}", new object[] { this.ListUrl });
                        using (SPSite site = new SPSite(this.ListUrl))
                        {
                            web = site.OpenWeb();
                        }
                    }
                    else
                        parameterDetails = string.Format(CultureInfo.CurrentCulture, "Id or Url : {0} and web Url {1}", new object[] { this.ListUrl, web.Url });

                    if (!web.Exists)
                        list = null;
                    else
                        list = web.GetList(serverRelativeListUrl);
                }
            }
            catch (Exception exception)
            {
                throw new SPCmdletPipeBindException(string.Format("The SPList Pipebind object could not be found ({0}).", parameterDetails), exception);
            }
            if (list == null)
                throw new SPCmdletPipeBindException(string.Format("The SPList Pipebind object could not be found ({0}).", parameterDetails));

            return list;
        }

        public bool IsCollection
        {
            get { return this._isCollection; }
        }

        public Guid ListGuid
        {
            get { return this._listGuid; }
        }

        public string ListUrl
        {
            get { return this._listUrl; }
        }
    }
}
