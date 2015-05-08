using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
#if MOSS
using Microsoft.SharePoint.Publishing;
#endif

namespace Lapointe.SharePoint.PowerShell.Common
{
    public class SPEnumerator
    {
        private object root;

        #region Events

        public delegate void SPWebApplicationEnumeratedEventHandler(object sender, SPWebApplicationEventArgs e);
        public delegate void SPSiteEnumeratedEventHandler(object sender, SPSiteEventArgs e);
        public delegate void SPWebEnumeratedEventHandler(object sender, SPWebEventArgs e);
        public delegate void SPListEnumeratedEventHandler(object sender, SPListEventArgs e);
        public delegate void PublishingPageEnumeratedEventHandler(object sender, PublishingPageEventArgs e);

        public event SPWebApplicationEnumeratedEventHandler SPWebApplicationEnumerated;
        public event SPSiteEnumeratedEventHandler SPSiteEnumerated;
        public event SPWebEnumeratedEventHandler SPWebEnumerated;
        public event SPListEnumeratedEventHandler SPListEnumerated;
        public event PublishingPageEnumeratedEventHandler PublishingPageEnumerated;

        /// <summary>
        /// Called when an SPWebApplication object is enumerated.
        /// </summary>
        /// <param name="webApp">The web app.</param>
        protected void OnSPWebApplicationEnumerated(SPWebApplication webApp)
        {
            if (SPWebApplicationEnumerated != null)
                SPWebApplicationEnumerated(this, new SPWebApplicationEventArgs(webApp));
        }

        /// <summary>
        /// Called when an SPSite object is enumerated.
        /// </summary>
        /// <param name="site">The site.</param>
        protected void OnSPSiteEnumerated(SPSite site)
        {
            if (SPSiteEnumerated != null)
                SPSiteEnumerated(this, new SPSiteEventArgs(site));
        }

        /// <summary>
        /// Called when an SPWeb object is enumerated.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        protected void OnSPWebEnumerated(SPSite site, SPWeb web)
        {
            if (SPWebEnumerated != null)
                SPWebEnumerated(this, new SPWebEventArgs(site, web));
        }

        /// <summary>
        /// Called when SPList object is enumerated.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        /// <param name="list">The list.</param>
        protected void OnSPListEnumerated(SPSite site, SPWeb web, SPList list)
        {
            if (SPListEnumerated != null)
                SPListEnumerated(this, new SPListEventArgs(site, web, list));
        }

#if MOSS
        /// <summary>
        /// Called when a PublishingPage object is enumerated.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        /// <param name="page">The page.</param>
        protected void OnPublishingPageEnumerated(SPSite site, SPWeb web, PublishingPage page)
        {
            if (PublishingPageEnumerated != null)
                PublishingPageEnumerated(this, new PublishingPageEventArgs(site, web, page));
        }
#endif
        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SPEnumerator"/> class.
        /// </summary>
        /// <param name="farm">The farm.</param>
        public SPEnumerator(SPFarm farm)
        {
            root = farm;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SPEnumerator"/> class.
        /// </summary>
        /// <param name="webApp">The web app.</param>
        public SPEnumerator(SPWebApplication webApp)
        {
            root = webApp;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SPEnumerator"/> class.
        /// </summary>
        /// <param name="site">The site.</param>
        public SPEnumerator(SPSite site)
        {
            root = site;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SPEnumerator"/> class.
        /// </summary>
        /// <param name="web">The web.</param>
        public SPEnumerator(SPWeb web)
        {
            root = web;
        }

        #endregion

        #region Enumeration Methods

        /// <summary>
        /// Enumerates through the object passed in via the constructor firing an event to indicate the object was enumerated when necessary.
        /// </summary>
        public void Enumerate()
        {
            if (root is SPFarm)
            {
                Enumerate((SPFarm) root);
            }
            else if (root is SPWebApplication)
            {
                Enumerate(((SPWebApplication) root));
            }
            else if (root is SPSite)
            {
                Enumerate(((SPSite)root));
            }
            else if (root is SPWeb)
            {
                using (SPSite site = ((SPWeb)root).Site)
                    Enumerate(site, ((SPWeb)root), true);
            }
        }

        /// <summary>
        /// Enumerates the specified farm.
        /// </summary>
        /// <param name="farm">The farm.</param>
        private void Enumerate(SPFarm farm)
        {
            foreach (SPService svc in farm.Services)
            {
                if (!(svc is SPWebService))
                    continue;

                foreach (SPWebApplication webApp in ((SPWebService)svc).WebApplications)
                {
                    Enumerate(webApp);
                }
            }
        }

        /// <summary>
        /// Enumerates the specified web app.
        /// </summary>
        /// <param name="webApp">The web app.</param>
        private void Enumerate(SPWebApplication webApp)
        {
            OnSPWebApplicationEnumerated(webApp);


            if (SPSiteEnumerated != null || SPWebEnumerated != null || SPListEnumerated != null || PublishingPageEnumerated != null)
            {
                foreach (SPSite site in webApp.Sites)
                {
                    try
                    {
                        Enumerate(site);
                    }
                    finally
                    {
                        site.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// Enumerates the specified site.
        /// </summary>
        /// <param name="site">The site.</param>
        private void Enumerate(SPSite site)
        {
            OnSPSiteEnumerated(site);

            if (SPWebEnumerated != null || SPListEnumerated != null || PublishingPageEnumerated != null)
            {
                foreach (SPWeb web in site.AllWebs)
                {
                    try
                    {
                        Enumerate(site, web, false);
                    }
                    finally
                    {
                        web.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// Enumerates the specified web.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        /// <param name="enumerateSubWebs">if set to <c>true</c> [enumerate sub webs].</param>
        private void Enumerate(SPSite site, SPWeb web, bool enumerateSubWebs)
        {
            OnSPWebEnumerated(site, web);

            if (SPListEnumerated != null)
            {
                for (int i = 0; i < web.Lists.Count; i++)
                {
                    OnSPListEnumerated(site, web, web.Lists[i]);
                }
            }

#if MOSS
            if (PublishingPageEnumerated != null)
            {
                if (PublishingWeb.IsPublishingWeb(web))
                {
                    PublishingWeb pubWeb = PublishingWeb.GetPublishingWeb(web);
                    foreach (PublishingPage page in pubWeb.GetPublishingPages())
                    {
                        OnPublishingPageEnumerated(site, web, page);
                    }
                }
            }
#endif
            if (enumerateSubWebs)
            {
                foreach (SPWeb subWeb in web.Webs)
                {
                    try
                    {
                        Enumerate(site, subWeb, enumerateSubWebs);
                    }
                    finally
                    {
                        subWeb.Dispose();
                    }
                }
            }
        }

        #endregion

        #region EventArgs Classes

        public class SPWebApplicationEventArgs : EventArgs
        {
            private SPWebApplication m_WebApp = null;
            internal SPWebApplicationEventArgs(SPWebApplication webApp)
            {
                m_WebApp = webApp;
            }
            public SPWebApplication WebApplication
            {
                get { return m_WebApp; }
            }
        }

        public class SPSiteEventArgs : SPWebApplicationEventArgs
        {
            private SPSite m_Site = null;
            internal SPSiteEventArgs(SPSite site) : base(site.WebApplication)
            {
                m_Site = site;
            }
            public SPSite Site
            {
                get { return m_Site; }
            }
        }

        public class SPWebEventArgs : SPSiteEventArgs
        {
            private SPWeb m_Web = null;

            internal SPWebEventArgs(SPSite site, SPWeb web) : base(site)
            {
                m_Web = web;
            }
            public SPWeb Web
            {
                get { return m_Web; }
            }
        }


        public class SPListEventArgs : SPWebEventArgs
        {
            private SPList m_List = null;

            internal SPListEventArgs(SPSite site, SPWeb web, SPList list) : base(site, web)
            {
                m_List = list;
            }
            public SPList List
            {
                get { return m_List; }
            }
        }

        public class PublishingPageEventArgs : SPListEventArgs
        {
#if MOSS
            private PublishingPage m_Page = null;
            private PublishingWeb m_PublishingWeb = null;

            internal PublishingPageEventArgs(SPSite site, SPWeb web, PublishingPage page) : base(site, web, page.ListItem.ParentList)
            {
                m_Page = page;
            }
            public PublishingPage Page
            {
                get { return m_Page; }
            }
            public PublishingWeb PublishingWeb
            {
                get
                {
                    if (m_PublishingWeb == null)
                        m_PublishingWeb = PublishingWeb.GetPublishingWeb(Web);
                    return m_PublishingWeb;
                }
            }
#else
            internal PublishingPageEventArgs(SPSite site, SPWeb web, SPList list) : base(site, web, list)
            {}
#endif
        }

        #endregion
    }
}
