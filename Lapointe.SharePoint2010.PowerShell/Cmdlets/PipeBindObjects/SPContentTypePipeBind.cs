using System;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;
using System.Globalization;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects
{
    public sealed class SPContentTypePipeBind : SPCmdletPipeBind<SPContentType>
    {
        private Guid m_SiteGuid;
        private Guid m_WebGuid;

        private SPContentTypeId m_ctId;
        private string m_ContentTypeName;

        public SPContentTypePipeBind(SPContentType instance)
            : base(instance)
        {
        }

        public SPContentTypePipeBind(SPContentTypeId id)
        {
            this.m_ctId = id;
        }

        public SPContentTypePipeBind(string inputString)
        {
            if (inputString != null)
            {
                inputString = inputString.Trim();
                try
                {
                    this.m_ctId = new SPContentTypeId(inputString);
                }
                catch (ArgumentException)
                {
                }
                catch (OverflowException)
                {
                }
                if (this.m_ctId.ToString() == "0x")
                {
                    this.m_ContentTypeName = inputString;
                }
            }
        }

        protected override void Discover(SPContentType instance)
        {
            this.m_ctId = instance.Id;
            this.m_WebGuid = instance.ParentWeb.ID;
            this.m_SiteGuid = instance.ParentWeb.Site.ID;
        }

        public override SPContentType Read()
        {
            return this.Read((SPWeb)null);
        }

        public SPContentType Read(SPList list)
        {
            SPContentType ct = null;
            if (list == null)
            {
                if (Guid.Empty != m_WebGuid && Guid.Empty != m_SiteGuid)
                {
                    using (SPSite site = new SPSite(m_SiteGuid))
                    {
                        SPWeb web = site.OpenWeb(m_WebGuid);
                        ct = Read(web);
                    }
                }
            }
            else
            {
                if (m_ctId.ToString() != "0x")
                    ct = list.ContentTypes[m_ctId];
                else if (!string.IsNullOrEmpty(m_ContentTypeName))
                    ct = list.ContentTypes[m_ContentTypeName];
            }

            if (ct == null)
            {
                throw new SPCmdletPipeBindException("The SPContentType Pipebind object could not be found.");
            }
            return ct;
        }

        public SPContentType Read(SPWeb web)
        {
            SPContentType ct = null;
            if (web == null)
            {
                if (Guid.Empty != m_WebGuid && Guid.Empty != m_SiteGuid)
                {
                    using (SPSite site = new SPSite(m_SiteGuid))
                    {
                        web = site.OpenWeb(m_WebGuid);
                    }
                }
            }
            if (m_ctId.ToString() != "0x")
                ct = web.ContentTypes[m_ctId];
            else if (!string.IsNullOrEmpty(m_ContentTypeName))
                ct = web.ContentTypes[m_ContentTypeName];

            if (ct == null)
            {
                throw new SPCmdletPipeBindException("The SPContentType Pipebind object could not be found.");
            }
            return ct;
        }

        public SPContentTypeId ContentTypeId
        {
            get
            {
                return this.m_ctId;
            }
        }

        public string ContentTypeName
        {
            get
            {
                return this.m_ContentTypeName;
            }
        }
    }


}
