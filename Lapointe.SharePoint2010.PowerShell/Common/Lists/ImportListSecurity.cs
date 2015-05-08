using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Microsoft.SharePoint;
using System.Xml;
using System.IO;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    internal static class ImportListSecurity
    {
        public static void ImportSecurity(XmlDocument xmlDoc, bool includeItemSecurity, SPList list)
        {
            if (xmlDoc.SelectNodes("//List").Count > 1)
                throw new SPException("A single target list was specified but the input file contains multiple lists.");

            if (xmlDoc.SelectNodes("//List").Count == 0)
                throw new SPException("No list information was found in the input file.");


            Logger.Write("Start Time: {0}.", DateTime.Now.ToString());


            XmlElement listElement = xmlDoc.SelectSingleNode("//List") as XmlElement;
            ImportSecurity(list, list.ParentWeb, includeItemSecurity, listElement);

            Logger.Write("Finish Time: {0}.\r\n", DateTime.Now.ToString());
        }

        public static void ImportSecurity(XmlDocument xmlDoc, bool includeItemSecurity, SPWeb web)
        {
            Logger.Write("Start Time: {0}.", DateTime.Now.ToString());

            foreach (XmlElement listElement in xmlDoc.SelectNodes("//List"))
            {
                SPList list = null;

                try
                {
                    list = web.GetList(web.ServerRelativeUrl.TrimEnd('/') + "/" + listElement.GetAttribute("Url"));
                }
                catch (ArgumentException) { }
                catch (FileNotFoundException) { }

                if (list == null)
                {
                    Console.WriteLine("WARNING: List was not found - skipping.");
                    continue;
                }

                ImportSecurity(list, web, includeItemSecurity, listElement);

            }
            Logger.Write("Finish Time: {0}.\r\n", DateTime.Now.ToString());
        }

        /// <summary>
        /// Imports the security.
        /// </summary>
        /// <param name="xmlDoc">The XML doc.</param>
        /// <param name="url">The URL.</param>
        /// <param name="includeItemSecurity">if set to <c>true</c> [include item security].</param>
        public static void ImportSecurity(XmlDocument xmlDoc, string url, bool includeItemSecurity)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb())
            {
                ImportSecurity(xmlDoc, includeItemSecurity, web);
            }
        }

        /// <summary>
        /// Imports the security.
        /// </summary>
        /// <param name="targetList">The target list.</param>
        /// <param name="web">The web.</param>
        /// <param name="includeItemSecurity">if set to <c>true</c> [include item security].</param>
        /// <param name="listElement">The list element.</param>
        internal static void ImportSecurity(SPList targetList, SPWeb web, bool includeItemSecurity, XmlElement listElement)
        {
            Logger.Write("Progress: Processing list \"{0}\".", targetList.RootFolder.ServerRelativeUrl);

            try
            {
                int writeSecurity = int.Parse(listElement.GetAttribute("WriteSecurity"));
                int readSecurity = int.Parse(listElement.GetAttribute("ReadSecurity"));

                if (writeSecurity != targetList.WriteSecurity)
                    targetList.WriteSecurity = writeSecurity;

                if (readSecurity != targetList.ReadSecurity)
                    targetList.ReadSecurity = readSecurity;

                // Set the security on the list itself.
                SetObjectSecurity(web, targetList, targetList.RootFolder.ServerRelativeUrl, listElement);

                // Set the security on any folders in the list.
                SetFolderSecurity(web, targetList, listElement);

                // Set the security on any items in the list.
                if (includeItemSecurity)
                    SetItemSecurity(web, targetList, listElement);


                if (listElement.HasAttribute("AnonymousPermMask64"))
                {
                    SPBasePermissions anonymousPermMask64 = (SPBasePermissions)int.Parse(listElement.GetAttribute("AnonymousPermMask64"));
                    if (anonymousPermMask64 != targetList.AnonymousPermMask64 && targetList.HasUniqueRoleAssignments)
                        targetList.AnonymousPermMask64 = anonymousPermMask64;
                }

                if (listElement.HasAttribute("AllowEveryoneViewItems"))
                {
                    bool allowEveryoneViewItems = bool.Parse(listElement.GetAttribute("AllowEveryoneViewItems"));
                    if (allowEveryoneViewItems != targetList.AllowEveryoneViewItems)
                        targetList.AllowEveryoneViewItems = allowEveryoneViewItems;
                }

                targetList.Update();

            }
            finally
            {
                Logger.Write("Progress: Finished processing list \"{0}\".", targetList.RootFolder.ServerRelativeUrl);
            }
        }

        /// <summary>
        /// Sets the folder security.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="list">The list.</param>
        /// <param name="listElement">The list element.</param>
        private static void SetFolderSecurity(SPWeb web, SPList list, XmlElement listElement)
        {
            foreach (XmlElement folderElement in listElement.SelectNodes("Folder"))
            {
                string folderUrl = folderElement.GetAttribute("Url");
                SPListItem folder = null;
                foreach (SPListItem tempFolder in list.Folders)
                {
                    if (tempFolder.Folder.Url.ToLowerInvariant() == folderUrl.ToLowerInvariant())
                    {
                        folder = tempFolder;
                        break;
                    }
                }
                if (folder == null)
                {
                    Logger.WriteWarning("Progress: Unable to locate folder '{0}'.", folderUrl);
                    continue;
                }
                SetObjectSecurity(web, folder, folderUrl, folderElement);
            }
        }

        /// <summary>
        /// Sets the item security.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="list">The list.</param>
        /// <param name="listElement">The list element.</param>
        private static void SetItemSecurity(SPWeb web, SPList list, XmlElement listElement)
        {
            foreach (XmlElement itemElement in listElement.SelectNodes("Item"))
            {
                int itemId = int.Parse(itemElement.GetAttribute("Id"));
                SPListItem item = null;
                try
                {
                    item = list.GetItemById(itemId);
                }
                catch (ArgumentException)
                {
                    // no-op
                }
                if (item == null)
                {
                    Logger.WriteWarning("Progress: Unable to locate item '{0}'.", itemId.ToString());
                    continue;
                }
                SetObjectSecurity(web, item, "Item " + itemId, itemElement);
            }
        }

        /// <summary>
        /// Sets the object security.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="targetObject">The target object.</param>
        /// <param name="itemName">Name of the item.</param>
        /// <param name="sourceElement">The source element.</param>
        private static void SetObjectSecurity(SPWeb web, SPSecurableObject targetObject, string itemName, XmlElement sourceElement)
        {

            bool hasUniqueRoleAssignments = bool.Parse(sourceElement.GetAttribute("HasUniqueRoleAssignments"));

            if (!hasUniqueRoleAssignments && targetObject.HasUniqueRoleAssignments)
            {
                Logger.Write("Progress: Setting target object to inherit permissions from parent for \"{0}\".", itemName);
                targetObject.ResetRoleInheritance();
                return;
            }
            else if (hasUniqueRoleAssignments && !targetObject.HasUniqueRoleAssignments)
            {
                Logger.Write("Progress: Breaking target object inheritance from parent for \"{0}\".", itemName);
                targetObject.BreakRoleInheritance(false);
            }
            else if (!hasUniqueRoleAssignments && !targetObject.HasUniqueRoleAssignments)
            {
                Logger.Write("Progress: Ignoring \"{0}\".  Target object and source object both inherit from parent.", itemName);
                return; // Both are inheriting so don't change.
            }
            if (hasUniqueRoleAssignments && targetObject.HasUniqueRoleAssignments)
            {
                while (targetObject.RoleAssignments.Count > 0)
                    targetObject.RoleAssignments.Remove(0); // Clear out any existing permissions
            }

            foreach (XmlElement roleAssignmentElement in sourceElement.SelectNodes("RoleAssignments/RoleAssignment"))
            {
                string memberName = roleAssignmentElement.GetAttribute("Member");
                string userName = null;
                if (roleAssignmentElement.HasAttribute("LoginName"))
                    userName = roleAssignmentElement.GetAttribute("LoginName");

                SPRoleAssignment existingRoleAssignment = GetRoleAssignement(web, targetObject, memberName, userName);

                if (existingRoleAssignment != null)
                {
                    if (AddRoleDefinitions(web, existingRoleAssignment, roleAssignmentElement))
                    {
                        existingRoleAssignment.Update();

                        Logger.Write("Progress: Updated \"{0}\" at target object \"{1}\".", memberName, itemName);
                    }
                }
                else
                {
                    SPPrincipal principal = GetPrincipal(web, memberName, userName);
                    if (principal == null)
                    {
                        Logger.WriteWarning("Progress: Unable to add Role Assignment for \"{0}\" - Member \"{1}\" not found.", itemName, memberName);
                        continue;
                    }

                    SPRoleAssignment newRA = new SPRoleAssignment(principal);
                    AddRoleDefinitions(web, newRA, roleAssignmentElement);

                    if (newRA.RoleDefinitionBindings.Count == 0)
                    {
                        Logger.WriteWarning("Progress: Unable to add \"{0}\" to target object \"{1}\" (principals with only \"Limited Access\" cannot be added).", memberName, itemName);
                        continue;
                    }

                    Logger.Write("Progress: Adding new Role Assignment \"{0}\".", newRA.Member.Name);

                    targetObject.RoleAssignments.Add(newRA);

                    existingRoleAssignment = GetRoleAssignement(targetObject, principal);
                    if (existingRoleAssignment == null)
                    {
                        Logger.WriteWarning("Progress: Unable to add \"{0}\" to target object \"{1}\".", memberName, itemName);
                    }
                    else
                    {
                        Logger.Write("Progress: Added \"{0}\" to target object \"{1}\".", memberName, itemName);
                    }
                }
            }
        }

        /// <summary>
        /// Adds the role definitions.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="roleAssignment">The role assignment.</param>
        /// <param name="roleAssignmentElement">The role assignment element.</param>
        /// <returns></returns>
        private static bool AddRoleDefinitions(SPWeb web, SPRoleAssignment roleAssignment, XmlElement roleAssignmentElement)
        {
            bool modified = false;
            foreach (XmlElement roleDefinitionElement in roleAssignmentElement.SelectNodes("RoleDefinitionBindings/RoleDefinition"))
            {
                string name = roleDefinitionElement.GetAttribute("Name");
                if (name == "Limited Access")
                    continue;

                SPRoleDefinition existingRoleDef = null;
                try
                {
                    existingRoleDef = web.RoleDefinitions[name];
                }
                catch (Exception) { }
                if (existingRoleDef == null)
                {
                    Logger.Write("Progress: Adding new Role Definition \"{0}\".", name);

                    SPBasePermissions perms = SPBasePermissions.EmptyMask;
                    foreach (string perm in roleDefinitionElement.GetAttribute("BasePermissions").Split(','))
                    {
                        perms = perms | (SPBasePermissions)Enum.Parse(typeof(SPBasePermissions), perm, true);
                    }
                    existingRoleDef = new SPRoleDefinition();
                    existingRoleDef.Name = name;
                    existingRoleDef.BasePermissions = perms;
                    existingRoleDef.Description = roleDefinitionElement.GetAttribute("Description");
                    existingRoleDef.Order = int.Parse(roleDefinitionElement.GetAttribute("Order"));
                    existingRoleDef.Update();

                    SPWeb tempWeb = web;
                    while (!tempWeb.HasUniqueRoleDefinitions)
                        tempWeb = tempWeb.ParentWeb;

                    tempWeb.RoleDefinitions.Add(existingRoleDef);
                }
                if (!roleAssignment.RoleDefinitionBindings.Contains(existingRoleDef))
                {
                    roleAssignment.RoleDefinitionBindings.Add(existingRoleDef);
                    modified = true;
                }
            }
            List<SPRoleDefinition> roleDefsToRemove = new List<SPRoleDefinition>();
            foreach (SPRoleDefinition roleDef in roleAssignment.RoleDefinitionBindings)
            {
                if (roleDef.Name == "Limited Access")
                    continue;

                bool found = false;
                foreach (XmlElement roleDefinitionElement in roleAssignmentElement.SelectNodes("RoleDefinitionBindings/RoleDefinition"))
                {
                    if (roleDef.Name == roleDefinitionElement.GetAttribute("Name"))
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    roleDefsToRemove.Add(roleDef);
                    modified = true;
                }
            }
            foreach (SPRoleDefinition roleDef in roleDefsToRemove)
            {
                Logger.Write("Progress: Removing '{0}' from '{1}'", roleDef.Name, roleAssignment.Member.Name);
                roleAssignment.RoleDefinitionBindings.Remove(roleDef);
            }
            return modified;
        }

        /// <summary>
        /// Gets the role assignement.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="securableObject">The securable object.</param>
        /// <param name="memberName">Name of the member.</param>
        /// <param name="userName">Name of the user.</param>
        /// <returns></returns>
        private static SPRoleAssignment GetRoleAssignement(SPWeb web, SPSecurableObject securableObject, string memberName, string userName)
        {
            SPPrincipal principal = GetPrincipal(web, memberName, userName);
            return GetRoleAssignement(securableObject, principal);
        }

        /// <summary>
        /// Gets the role assignement.
        /// </summary>
        /// <param name="securableObject">The securable object.</param>
        /// <param name="principal">The principal.</param>
        /// <returns></returns>
        private static SPRoleAssignment GetRoleAssignement(SPSecurableObject securableObject, SPPrincipal principal)
        {
            SPRoleAssignment ra = null;
            try
            {
                ra = securableObject.RoleAssignments.GetAssignmentByPrincipal(principal);
            }
            catch (ArgumentException)
            { }
            return ra;
        }

        /// <summary>
        /// Gets the principal.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="memberName">Name of the member.</param>
        /// <returns></returns>
        internal static SPPrincipal GetPrincipal(SPWeb web, string memberName)
        {
            return GetPrincipal(web, memberName, null);
        }

        /// <summary>
        /// Gets the principal.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="memberName">Name of the member.</param>
        /// <param name="loginName">Name of the login.</param>
        /// <returns></returns>
        internal static SPPrincipal GetPrincipal(SPWeb web, string memberName, string loginName)
        {
            foreach (SPPrincipal p in web.SiteUsers)
            {
                if (p.Name.ToLowerInvariant() == memberName.ToLowerInvariant())
                    return p;
            }
            foreach (SPPrincipal p in web.SiteGroups)
            {
                if (p.Name.ToLowerInvariant() == memberName.ToLowerInvariant())
                    return p;
            }

            try
            {
                SPPrincipal principal;
                if (!string.IsNullOrEmpty(loginName) && Microsoft.SharePoint.Utilities.SPUtility.IsLoginValid(web.Site, loginName))
                {
                    // We have a user   
                    Logger.Write("Progress: Adding user '{0}' to site.", loginName);
                    principal = web.EnsureUser(loginName);
                }
                else
                {
                    // We have a group   

                    SPGroup groupToAdd = null;
                    try
                    {
                        groupToAdd = web.SiteGroups[memberName];
                    }
                    catch (SPException)
                    {
                    }
                    if (groupToAdd != null)
                    {
                        // The group exists, so get it   
                        principal = groupToAdd;
                    }
                    else
                    {
                        // The group didn't exist so we need to create it:   
                        //  Create it:  
                        Logger.Write("Progress: Adding group '{0}' to site.", memberName);
                        web.SiteGroups.Add(memberName, web.Site.Owner, web.Site.Owner, string.Empty);
                        //  Get it:   
                        principal = web.SiteGroups[memberName];
                    }
                }
                return principal;
            }
            catch (Exception ex)
            {
                Logger.Write("WARNING: Unable to add member to site: {0}\r\n{1}", memberName, Utilities.FormatException(ex));
            }
            return null;
        }
 
    }
}
