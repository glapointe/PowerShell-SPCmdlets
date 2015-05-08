using System;
using System.IO;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.Common;

namespace Lapointe.SharePoint.PowerShell.Common.Features
{
    public enum ActivationScope
    {
        Farm, WebApplication, Site, Web, Feature
    }
    public class FeatureHelper
    {
        private string m_Url;
        private bool m_Force;
        private Guid m_FeatureId = Guid.Empty;
        private bool m_IgnoreNonActive;
        private bool m_Activate;

        /// <summary>
        /// Gets the feature id from params.
        /// </summary>
        /// <param name="Params">The params.</param>
        /// <returns></returns>
        internal static Guid GetFeatureIdFromParams(SPParamCollection Params)
        {
            Guid empty = Guid.Empty;
            if (!Params["id"].UserTypedIn)
            {
                SPFeatureScope scope;
                if (!Params["filename"].UserTypedIn)
                {
                    if (Params["name"].UserTypedIn)
                    {
                        SPFeatureScope scope2;
                        SPFeatureDefinition.GetFeatureIdAndScope(Params["name"].Value + @"\feature.xml", out empty, out scope2);
                    }
                    return empty;
                }
                SPFeatureDefinition.GetFeatureIdAndScope(Params["filename"].Value, out empty, out scope);
                return empty;
            }
            return new Guid(Params["id"].Value);
        }

        /// <summary>
        /// Activates or deactivates the feature at the specified scope.
        /// </summary>
        /// <param name="scope">The scope.</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="url">The URL.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        internal void ActivateDeactivateFeatureAtScope(ActivationScope scope, Guid featureId, bool activate, string url, bool force, bool ignoreNonActive)
        {
            Logger.Verbose = true;

            m_FeatureId = featureId;

            if (m_FeatureId.Equals(Guid.Empty))
                throw new SPException("Unable to locate Feature.");

            SPFeatureDefinition feature = SPFarm.Local.FeatureDefinitions[m_FeatureId];
            if (feature == null)
                throw new SPException("Unable to locate Feature.");

            if (scope == ActivationScope.Feature)
                scope = (ActivationScope)Enum.Parse(typeof(ActivationScope), feature.Scope.ToString().ToLowerInvariant(), true);

            ActivateDeactivateFeatureAtScope(feature, scope, activate, url, force, ignoreNonActive);
        }

        internal void ActivateDeactivateFeatureAtScope(SPFeatureDefinition feature, ActivationScope scope, bool activate, string url, bool force, bool ignoreNonActive)
        {
            m_IgnoreNonActive = ignoreNonActive;
            m_Activate = activate;
            m_Force = force;
            m_Url = url;
            m_FeatureId = feature.Id;

            if (feature.Scope == SPFeatureScope.Farm)
            {
                if (scope != ActivationScope.Farm)
                    throw new SPSyntaxException("The Feature specified is scoped to the Farm.  The -scope parameter must be \"Farm\".");
                ActivateDeactivateFeatureAtFarm(activate, m_FeatureId, m_Force, m_IgnoreNonActive);
            }
            else if (feature.Scope == SPFeatureScope.WebApplication)
            {
                if (scope != ActivationScope.Farm && scope != ActivationScope.WebApplication)
                    throw new SPSyntaxException("The Feature specified is scoped to the Web Application.  The -scope parameter must be either \"Farm\" or \"WebApplication\".");

                if (scope == ActivationScope.Farm)
                {
                    SPEnumerator enumerator = new SPEnumerator(SPFarm.Local);
                    enumerator.SPWebApplicationEnumerated += enumerator_SPWebApplicationEnumerated;
                    enumerator.Enumerate();
                }
                else
                {
                    if (string.IsNullOrEmpty(m_Url))
                        throw new SPSyntaxException("The -url parameter is required if the scope is \"WebApplication\".");
                    SPWebApplication webApp = SPWebApplication.Lookup(new Uri(m_Url));
                    ActivateDeactivateFeatureAtWebApplication(webApp, m_FeatureId, activate, m_Force, m_IgnoreNonActive);
                }
            }
            else if (feature.Scope == SPFeatureScope.Site)
            {
                if (scope == ActivationScope.Web)
                    throw new SPSyntaxException("The Feature specified is scoped to Site.  The -scope parameter cannot be \"Web\".");

                SPSite site = null;
                SPEnumerator enumerator = null;
                try
                {
                    if (scope == ActivationScope.Farm)
                        enumerator = new SPEnumerator(SPFarm.Local);
                    else if (scope == ActivationScope.WebApplication)
                    {
                        SPWebApplication webApp = SPWebApplication.Lookup(new Uri(m_Url));
                        enumerator = new SPEnumerator(webApp);
                    }
                    else if (scope == ActivationScope.Site)
                    {
                        site = new SPSite(m_Url);
                        ActivateDeactivateFeatureAtSite(site, activate, m_FeatureId, m_Force, m_IgnoreNonActive);
                    }
                    if (enumerator != null)
                    {
                        enumerator.SPSiteEnumerated += enumerator_SPSiteEnumerated;
                        enumerator.Enumerate();
                    }
                }
                finally
                {
                    if (site != null)
                        site.Dispose();
                }
            }
            else if (feature.Scope == SPFeatureScope.Web)
            {
                SPSite site = null;
                SPWeb web = null;
                SPEnumerator enumerator = null;
                try
                {
                    if (scope == ActivationScope.Farm)
                        enumerator = new SPEnumerator(SPFarm.Local);
                    else if (scope == ActivationScope.WebApplication)
                    {
                        SPWebApplication webApp = SPWebApplication.Lookup(new Uri(m_Url));
                        enumerator = new SPEnumerator(webApp);
                    }
                    else if (scope == ActivationScope.Site)
                    {
                        site = new SPSite(m_Url);
                        enumerator = new SPEnumerator(site);
                    }
                    else if (scope == ActivationScope.Web)
                    {

                        site = new SPSite(m_Url);
                        web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(m_Url)];
                        enumerator = new SPEnumerator(web);
                    }
                    if (enumerator != null)
                    {
                        enumerator.SPWebEnumerated += enumerator_SPWebEnumerated;
                        enumerator.Enumerate();
                    }
                }
                finally
                {
                    if (web != null)
                        web.Dispose();
                    if (site != null)
                        site.Dispose();
                }
            }
        }

        #region Event Handlers

        /// <summary>
        /// Handles the SPWebEnumerated event of the enumerator control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="Lapointe.SharePoint.STSADM.Commands.OperationHelpers.SPEnumerator.SPWebEventArgs"/> instance containing the event data.</param>
        private void enumerator_SPWebEnumerated(object sender, SPEnumerator.SPWebEventArgs e)
        {
            ActivateDeactivateFeatureAtWeb(e.Site, e.Web, m_Activate, m_FeatureId, m_Force, m_IgnoreNonActive);
        }

        /// <summary>
        /// Handles the SPSiteEnumerated event of the enumerator control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="Lapointe.SharePoint.STSADM.Commands.OperationHelpers.SPEnumerator.SPSiteEventArgs"/> instance containing the event data.</param>
        private void enumerator_SPSiteEnumerated(object sender, SPEnumerator.SPSiteEventArgs e)
        {
            ActivateDeactivateFeatureAtSite(e.Site, m_Activate, m_FeatureId, m_Force, m_IgnoreNonActive);
        }

        /// <summary>
        /// Handles the SPWebApplicationEnumerated event of the enumerator control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="Lapointe.SharePoint.STSADM.Commands.OperationHelpers.SPEnumerator.SPWebApplicationEventArgs"/> instance containing the event data.</param>
        private void enumerator_SPWebApplicationEnumerated(object sender, SPEnumerator.SPWebApplicationEventArgs e)
        {
            ActivateDeactivateFeatureAtWebApplication(e.WebApplication, m_FeatureId, m_Activate, m_Force, m_IgnoreNonActive);
        }

        #endregion

        /// <summary>
        /// Activates or deactivates the feature.
        /// </summary>
        /// <param name="features">The features.</param>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="urlScope">The URL scope.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        /// <returns></returns>
        private SPFeature ActivateDeactivateFeature(SPFeatureCollection features, bool activate, Guid featureId, string urlScope, bool force, bool ignoreNonActive)
        {
            if (features[featureId] == null && ignoreNonActive)
                return null;

            if (!activate)
            {
                if (features[featureId] != null || force)
                {
                    Logger.Write("Progress: Deactivating Feature {0} from {1}.", featureId.ToString(), urlScope);
                    try
                    {
                        features.Remove(featureId, force);
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteWarning("{0}", ex.Message);
                    }
                }
                else
                {
                    Logger.WriteWarning("" + SPResource.GetString("FeatureNotActivatedAtScope", new object[] { featureId }) + "  Use the -force parameter to force a deactivation.");
                }

                return null;
            }
            if (features[featureId] == null)
                Logger.Write("Progress: Activating Feature {0} on {1}.", featureId.ToString(), urlScope);
            else
            {
                if (!force)
                {
                    SPFeatureDefinition fd = features[featureId].Definition;
                    Logger.WriteWarning("" + SPResource.GetString("FeatureAlreadyActivated", new object[] { fd.DisplayName, fd.Id, urlScope }) + "  Use the -force parameter to force a reactivation.");
                    return features[featureId];
                }

                Logger.Write("Progress: Re-Activating Feature {0} on {1}.", featureId.ToString(), urlScope);
            }
            try
            {
                return features.Add(featureId, force);
            }
            catch(Exception ex)
            {
                Logger.WriteException(new System.Management.Automation.ErrorRecord(ex, null, System.Management.Automation.ErrorCategory.NotSpecified, features));
                return null;
            }
        }
        /// <summary>
        /// Activates or deactivates the farm scoped feature.
        /// </summary>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        /// <returns></returns>
        public SPFeature ActivateDeactivateFeatureAtFarm(bool activate, Guid featureId, bool force, bool ignoreNonActive)
        {
            SPWebService service = SPFarm.Local.Services.GetValue<SPWebService>(string.Empty);
            return ActivateDeactivateFeature(service.Features, activate, featureId, "Farm", force, ignoreNonActive);
        }

        /// <summary>
        /// Activates or deactivates the web application scoped feature.
        /// </summary>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="urlScope">The URL scope.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        /// <returns></returns>
        public SPFeature ActivateDeactivateFeatureAtWebApplication(bool activate, Guid featureId, string urlScope, bool force, bool ignoreNonActive)
        {
            SPWebApplication application = SPWebApplication.Lookup(new Uri(urlScope));
            if (application == null)
            {
                throw new FileNotFoundException(SPResource.GetString("WebApplicationLookupFailed", new object[] { urlScope }));
            }
            return ActivateDeactivateFeatureAtWebApplication(application, featureId, activate, force, ignoreNonActive);
        }

        /// <summary>
        /// Activates or deactivates the web application scoped feature.
        /// </summary>
        /// <param name="application">The application.</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        /// <returns></returns>
        public SPFeature ActivateDeactivateFeatureAtWebApplication(SPWebApplication application, Guid featureId, bool activate, bool force, bool ignoreNonActive)
        {
            return ActivateDeactivateFeature(application.Features, activate, featureId, application.GetResponseUri(SPUrlZone.Default).ToString(), force, ignoreNonActive);
        }

        /// <summary>
        /// Activates or deactivates the site scoped feature.
        /// </summary>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="urlScope">The URL scope.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        /// <returns></returns>
        public SPFeature ActivateDeactivateFeatureAtSite(bool activate, Guid featureId, string urlScope, bool force, bool ignoreNonActive)
        {
            using (SPSite site = new SPSite(urlScope))
            using (SPWeb web = site.OpenWeb(Utilities.GetServerRelUrlFromFullUrl(urlScope), true))
            {
                if (web.IsRootWeb)
                {
                    return ActivateDeactivateFeatureAtSite(site, activate, featureId, force, ignoreNonActive);
                }
                throw new SPException(SPResource.GetString("FeatureActivateDeactivateScopeAmbiguous", new object[] { site.Url }));
            }
        }

        /// <summary>
        /// Activates or deactivates the site scoped feature.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        /// <returns></returns>
        public SPFeature ActivateDeactivateFeatureAtSite(SPSite site, bool activate, Guid featureId, bool force, bool ignoreNonActive)
        {
            return ActivateDeactivateFeature(site.Features, activate, featureId, site.Url, force, ignoreNonActive);
        }

        /// <summary>
        /// Activates or deactivates the web scoped feature.
        /// </summary>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="urlScope">The URL scope.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        /// <returns></returns>
        public SPFeature ActivateDeactivateFeatureAtWeb(bool activate, Guid featureId, string urlScope, bool force, bool ignoreNonActive)
        {
            using (SPSite site = new SPSite(urlScope))
            using (SPWeb web = site.OpenWeb())
            {
                return ActivateDeactivateFeatureAtWeb(site, web, activate, featureId, force, ignoreNonActive);
            }
        }

        /// <summary>
        /// Activates or deactivates the web scoped feature.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        /// <param name="activate">if set to <c>true</c> [activate].</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="ignoreNonActive">if set to <c>true</c> [ignore non active].</param>
        /// <returns></returns>
        public SPFeature ActivateDeactivateFeatureAtWeb(SPSite site, SPWeb web, bool activate, Guid featureId, bool force, bool ignoreNonActive)
        {
            return ActivateDeactivateFeature(web.Features, activate, featureId, web.Url, force, ignoreNonActive);
        }

    }
}
