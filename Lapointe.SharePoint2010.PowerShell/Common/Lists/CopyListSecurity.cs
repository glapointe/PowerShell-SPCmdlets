using System;
using System.Text;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    internal static class CopyListSecurity
    {
        /// <summary>
        /// Copies the security.
        /// </summary>
        /// <param name="sourceList">The source list.</param>
        /// <param name="targetList">The target list.</param>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="includeItemSecurity">if set to <c>true</c> [include item security].</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        internal static void CopySecurity(SPList sourceList, SPList targetList, SPWeb targetWeb, bool includeItemSecurity, bool quiet)
        {
            if (!quiet)
                Logger.Write("Start Time: {0}.", DateTime.Now.ToString());

            try
            {
                if (sourceList.WriteSecurity != targetList.WriteSecurity)
                    targetList.WriteSecurity = sourceList.WriteSecurity;

                if (sourceList.ReadSecurity != targetList.ReadSecurity)
                    targetList.ReadSecurity = sourceList.ReadSecurity;

                // Set the security on the list itself.
                SetObjectSecurity(targetWeb, sourceList.ParentWeb, sourceList, targetList, quiet, targetList.RootFolder.ServerRelativeUrl);

                // Set the security on any folders in the list.
                SetFolderSecurity(targetWeb, sourceList, targetList, quiet);

                if (includeItemSecurity)
                {
                    // Set the security on list items.
                    SetListItemSecurity(targetWeb, sourceList, targetList, quiet);
                }


                if (sourceList.AnonymousPermMask64 != targetList.AnonymousPermMask64 && targetList.HasUniqueRoleAssignments)
                    targetList.AnonymousPermMask64 = sourceList.AnonymousPermMask64;

                if (sourceList.AllowEveryoneViewItems != targetList.AllowEveryoneViewItems)
                    targetList.AllowEveryoneViewItems = sourceList.AllowEveryoneViewItems;


                targetList.Update();
            }
            catch (Exception ex)
            {
                Logger.Write(Utilities.FormatException(ex));
            }
            finally
            {
                if (!quiet)
                    Logger.Write("Finish Time: {0}.\r\n", DateTime.Now.ToString());
            }
        }

        /// <summary>
        /// Sets the list item security.
        /// </summary>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="sourceList">The source list.</param>
        /// <param name="targetList">The target list.</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        private static void SetListItemSecurity(SPWeb targetWeb, SPList sourceList, SPList targetList, bool quiet)
        {
            SPField fileRef = default(SPField);
            try
            {
                // See if we have a FileLeafRef field (throws ArgumentException if not found - can't rely on display name)
                fileRef = sourceList.Fields.GetFieldByInternalName("FileLeafRef");
            }
            catch { }

            if (fileRef == default(SPField))
            {
                // Couldn't find a FileLeafRef field so use the index
                for (int i = 0; i < sourceList.ItemCount; i++)
                {
                    if (i > targetList.ItemCount)
                        break;

                    SPListItem targetItem = targetList.Items[i];
                    if (targetItem == null)
                        continue;

                    SetObjectSecurity(targetWeb, sourceList.ParentWeb, sourceList.Items[i], targetItem, quiet, string.Format("{0} (ID={1})", i, sourceList.Items[i].ID));
                }
            }
            else
            {
                // Found a FileLeafRef field so use it (this will be the filename for doc libs and the ID for lists).
                foreach (SPListItem sourceItem in sourceList.Items)
                {
                    SPQuery query = new SPQuery();
                    query.Query = string.Format(@"<Where><Eq><FieldRef Name=""FileLeafRef"" /><Value Type=""Text"">{0}</Value></Eq></Where>", System.Security.SecurityElement.Escape(sourceItem[fileRef.Id].ToString()));

                    SPListItemCollection items = targetList.GetItems(query);

                    foreach (SPListItem targetItem in items)
                    {
                        SetObjectSecurity(targetWeb, sourceList.ParentWeb, sourceItem, targetItem, quiet, sourceItem[fileRef.Id].ToString());
                    }
                }
            }
        }

        /// <summary>
        /// Sets the folder security.
        /// </summary>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="sourceList">The source list.</param>
        /// <param name="targetList">The target list.</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        private static void SetFolderSecurity(SPWeb targetWeb, SPList sourceList, SPList targetList, bool quiet)
        {
            foreach (SPListItem sourceFolder in sourceList.Folders)
            {
                SPListItem targetFolder = null;
                foreach (SPListItem f in targetList.Folders)
                {

                    if (f.Folder.ServerRelativeUrl.Substring(targetList.RootFolder.ServerRelativeUrl.Length) ==
                        sourceFolder.Folder.ServerRelativeUrl.Substring(sourceList.RootFolder.ServerRelativeUrl.Length))
                    {
                        targetFolder = f;
                        break;
                    }
                }
                if (targetFolder == null)
                    continue;

                SetObjectSecurity(targetWeb, sourceList.ParentWeb, sourceFolder, targetFolder, quiet, targetFolder.Folder.ServerRelativeUrl);
            }
        }



        /// <summary>
        /// Sets the object security.
        /// </summary>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="sourceWeb">The source web.</param>
        /// <param name="sourceObject">The source object.</param>
        /// <param name="targetObject">The target object.</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <param name="itemName">Name of the item.</param>
        internal static void SetObjectSecurity(SPWeb targetWeb, SPWeb sourceWeb, SPSecurableObject sourceObject, SPSecurableObject targetObject, bool quiet, string itemName)
        {
            if (!sourceObject.HasUniqueRoleAssignments && targetObject.HasUniqueRoleAssignments)
            {
                if (!quiet)
                    Logger.Write("Progress: Setting target object to inherit permissions from parent for \"{0}\".", itemName);
                targetObject.ResetRoleInheritance();
                return;
            }
            if (sourceObject.HasUniqueRoleAssignments && !targetObject.HasUniqueRoleAssignments)
            {
                if (!quiet)
                    Logger.Write("Progress: Breaking target object inheritance from parent for \"{0}\".", itemName);
                targetObject.BreakRoleInheritance(false);
            }
            else if (!sourceObject.HasUniqueRoleAssignments && !targetObject.HasUniqueRoleAssignments)
            {
                if (!quiet)
                    Logger.Write("Progress: Ignoring \"{0}\".  Target object and source object both inherit from parent.", itemName);
                return; // Both are inheriting so don't change.
            }

            foreach (SPRoleAssignment ra in sourceObject.RoleAssignments)
            {
                SPRoleAssignment existingRoleAssignment = GetRoleAssignement(targetObject, ra.Member);

                if (existingRoleAssignment != null)
                {
                    CopyRoleDefinitionBindings(targetWeb, ra, existingRoleAssignment, quiet);
                }
                else
                {
                    existingRoleAssignment = GetRoleAssignement(targetWeb, ra.Member);
                    if (existingRoleAssignment != null)
                    {
                        if (targetWeb.IsRootWeb &&
                            existingRoleAssignment.RoleDefinitionBindings.Count == 1 &&
                            existingRoleAssignment.RoleDefinitionBindings[0].Name == "Limited Access")
                        {
                            SPWeb tempSourceWeb = sourceWeb;
                            while (!tempSourceWeb.HasUniqueRoleAssignments)
                                tempSourceWeb = tempSourceWeb.ParentWeb;

                            SPRoleAssignment tempRa = GetRoleAssignement(tempSourceWeb, ra.Member);

                            CopyRoleDefinitionBindings(targetWeb, tempRa, existingRoleAssignment, quiet);
                        }
                        else
                            CopyRoleDefinitionBindings(targetWeb, ra, existingRoleAssignment, quiet);

                        targetObject.RoleAssignments.Add(existingRoleAssignment);
                        //if (targetObject is SPList)
                        //    ((SPList)targetObject).Update();

                        existingRoleAssignment = GetRoleAssignement(targetObject, ra.Member);

                        if (existingRoleAssignment == null)
                        {
                            Logger.Write("Progress: Unable to add \"{0}\" to target object \"{1}\".", ra.Member.ToString(), itemName);
                        }
                        else
                        {
                            if (!quiet)
                                Logger.Write("Progress: Added \"{0}\" to target object \"{1}\".", ra.Member.ToString(), itemName);
                        }
                        continue;
                    }

                    SPPrincipal principal = Common.Lists.ImportListSecurity.GetPrincipal(targetWeb, ra.Member.ToString());
                    if (principal == null)
                    {
                        if (ra.Member is SPUser)
                        {
                            targetWeb.AllUsers.Add(ra.Member.ToString(), null, null, null);
                            principal = targetWeb.AllUsers[ra.Member.ToString()];
                        }
                    }
                    if (principal == null)
                    {
                        Logger.Write("Progress: Unable to add \"{0}\" to target object \"{1}\".", ra.Member.ToString(), itemName);
                        continue;
                    }
                    SPRoleAssignment newRA = new SPRoleAssignment(principal);

                    CopyRoleDefinitionBindings(targetWeb, ra, newRA, quiet);

                    if (newRA.RoleDefinitionBindings.Count == 0)
                    {
                        if (!quiet)
                            Logger.Write("Progress: Unable to add \"{0}\" to target object \"{1}\" (principals with only \"Limited Access\" cannot be added).", ra.Member.ToString(), itemName);
                        continue;
                    }

                    targetObject.RoleAssignments.Add(newRA);

                    existingRoleAssignment = GetRoleAssignement(targetObject, ra.Member);
                    if (existingRoleAssignment == null)
                    {
                        Logger.Write("Progress: Unable to add \"{0}\" to target object \"{1}\".", ra.Member.ToString(), itemName);
                    }
                    else
                    {
                        if (!quiet)
                            Logger.Write("Progress: Added \"{0}\" to target object \"{1}\".", ra.Member.ToString(), itemName);
                    }
                }
            }
        }

        /// <summary>
        /// Copies the role definition bindings.
        /// </summary>
        /// <param name="targetWeb">The target web.</param>
        /// <param name="sourceRoleAssignment">The source role assignment.</param>
        /// <param name="targetRoleAssignment">The target role assignment.</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        private static void CopyRoleDefinitionBindings(SPWeb targetWeb, SPRoleAssignment sourceRoleAssignment, SPRoleAssignment targetRoleAssignment, bool quiet)
        {
            bool modified = false;
            foreach (SPRoleDefinition rd in sourceRoleAssignment.RoleDefinitionBindings)
            {
                if (rd.Name == "Limited Access")
                    continue;

                SPRoleDefinition existingRoleDef = null;
                try
                {
                    existingRoleDef = targetWeb.RoleDefinitions[rd.Name];
                }
                catch (SPException) { }
                if (existingRoleDef == null)
                {
                    existingRoleDef = new SPRoleDefinition();
                    existingRoleDef.BasePermissions = rd.BasePermissions;
                    existingRoleDef.Description = rd.Description;
                    existingRoleDef.Name = rd.Name;


                    SPWeb tempWeb = targetWeb;
                    while (!tempWeb.HasUniqueRoleDefinitions)
                        tempWeb = tempWeb.ParentWeb;

                    tempWeb.RoleDefinitions.Add(existingRoleDef);
                    existingRoleDef = tempWeb.RoleDefinitions[rd.Name];

                    if (!quiet)
                        Logger.Write("Progress: Added \"{0}\" role definition to web \"{1}\".", rd.Name, targetWeb.ServerRelativeUrl);
                }
                if (!targetRoleAssignment.RoleDefinitionBindings.Contains(targetWeb.RoleDefinitions[existingRoleDef.Name]))
                {
                    modified = true;
                    targetRoleAssignment.RoleDefinitionBindings.Add(targetWeb.RoleDefinitions[existingRoleDef.Name]);
                    if (!quiet)
                        Logger.Write("Progress: Added \"{0}\" role definition to target role assignment \"{1}\".", existingRoleDef.Name, targetRoleAssignment.Member.Name);
                }

            }
            if (modified)
                targetRoleAssignment.Update();
        }

        /// <summary>
        /// Gets the role assignement.
        /// </summary>
        /// <param name="securableObject">The securable object.</param>
        /// <param name="principal">The principal.</param>
        /// <returns></returns>
        private static SPRoleAssignment GetRoleAssignement(SPSecurableObject securableObject, SPPrincipal principal)
        {
            SPSecurableObject so = securableObject;
            if (!so.HasUniqueRoleAssignments)
                so = so.FirstUniqueAncestorSecurableObject;

            foreach (SPRoleAssignment r in so.RoleAssignments)
            {
                if (r.Member.ToString().ToLowerInvariant() == principal.ToString().ToLowerInvariant())
                    return r;
            }
            return null;
        }
    }
}
