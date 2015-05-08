using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Common.WebApplications
{
    internal class AddUserPolicyForWebApp
    {
        internal static void AddUserPolicy(string url, string login, string username, string[] permissions, string zone)
        {
            SPWebApplication webApp = SPWebApplication.Lookup(new Uri(url));
            AddUserPolicy(login, username, permissions, zone, webApp);
        }

        internal static void AddUserPolicy(string login, string username, string[] permissions, string zone, SPWebApplication webApp)
        {
            SPPolicyCollection policies = GetZonePolicyCollection(zone, webApp);
            AddUserPolicy(login, username, permissions, webApp, policies);
        }

        internal static void AddUserPolicy(string login, string username, string[] permissions, SPWebApplication webApp, SPUrlZone[] zones)
        {
            foreach (SPUrlZone zone in zones)
                AddUserPolicy(login, username, permissions, webApp, webApp.ZonePolicies(zone));
        }

        internal static void AddUserPolicy(string login, string username, string[] permissions, SPWebApplication webApp)
        {
            AddUserPolicy(login, username, permissions, webApp, webApp.Policies);
        }

        internal static void AddUserPolicy(string login, string username, string[] permissions, SPWebApplication webApp, SPPolicyCollection policies)
        {
            login = Utilities.TryGetNT4StyleAccountName(login, webApp);

            List<SPPolicyRole> roles = new List<SPPolicyRole>();
            foreach (string roleName in permissions)
            {
                SPPolicyRole role = webApp.PolicyRoles[roleName.Trim()];
                if (role == null)
                    throw new SPException(string.Format("The policy permission '{0}' was not found.", roleName.Trim()));

                roles.Add(role);
            }
            SPPolicy policy = policies.Add(login, username);

            foreach (SPPolicyRole role in roles)
                policy.PolicyRoleBindings.Add(role);

            webApp.Update();
        }


        /// <summary>
        /// Gets the zone policy collection.
        /// </summary>
        /// <param name="zoneName">Name of the zone.</param>
        /// <param name="application">The application.</param>
        /// <returns></returns>
        internal static SPPolicyCollection GetZonePolicyCollection(string zoneName, SPWebApplication application)
        {
            zoneName = zoneName.ToLowerInvariant();
            return (!zoneName.Equals("all") ? application.ZonePolicies((SPUrlZone)Enum.Parse(typeof(SPUrlZone), zoneName, true)) : application.Policies);
        }
    }
}
