using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Xml;
using System.IO;
using Microsoft.SharePoint.Utilities;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    internal static class ExportListSecurity
    {

        public static XmlTextWriter OpenXmlWriter(StringBuilder sb)
        {
            Logger.Write("Start Time: {0}.", DateTime.Now.ToString());

            XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
            xmlWriter.Formatting = Formatting.Indented;

            xmlWriter.WriteStartElement("Lists");

            return xmlWriter;
        }

        public static void CloseXmlWriter(XmlTextWriter xmlWriter, string outputFile, StringBuilder sb)
        {
            xmlWriter.WriteEndElement(); // Lists

            xmlWriter.Flush();
            xmlWriter.Close();

            if (!string.IsNullOrEmpty(outputFile))
                File.WriteAllText(outputFile, sb.ToString());

            Logger.Write("Finish Time: {0}.\r\n", DateTime.Now.ToString());
        }

        public static void ExportSecurity(string outputFile, bool includeItemSecurity, SPWeb web)
        {
            StringBuilder sb = new StringBuilder();
            XmlTextWriter xmlWriter = OpenXmlWriter(sb);

            foreach (SPList list in web.Lists)
            {
                ExportSecurity(list, web, xmlWriter, includeItemSecurity);
            }

            CloseXmlWriter(xmlWriter, outputFile, sb);
        }

        public static void ExportSecurity(string outputFile, bool includeItemSecurity, SPList list)
        {
            StringBuilder sb = new StringBuilder();
            XmlTextWriter xmlWriter = OpenXmlWriter(sb);

            ExportSecurity(list, list.ParentWeb, xmlWriter, includeItemSecurity);

            CloseXmlWriter(xmlWriter, outputFile, sb);
        }

        /// <summary>
        /// Exports the security.
        /// </summary>
        /// <param name="outputFile">The output file.</param>
        /// <param name="scope">The scope.</param>
        /// <param name="url">The URL.</param>
        /// <param name="includeItemSecurity">if set to <c>true</c> [include item security].</param>
        public static void ExportSecurity(string outputFile, string scope, string url, bool includeItemSecurity)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb())
            {
                if (scope == "list")
                {
                    SPList list = Utilities.GetListFromViewUrl(web, url);

                    if (list == null)
                        throw new SPException("List was not found.");
                    ExportSecurity(outputFile, includeItemSecurity, list);
                }
                else
                {
                    ExportSecurity(outputFile, includeItemSecurity, web);
                }
            }
        }

        /// <summary>
        /// Exports the security.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="web">The web.</param>
        /// <param name="xmlWriter">The XML writer.</param>
        /// <param name="includeItemSecurity">if set to <c>true</c> [include item security].</param>
        public static void ExportSecurity(SPList list, SPWeb web, XmlTextWriter xmlWriter, bool includeItemSecurity)
        {
            try
            {
                Logger.Write("Progress: Processing list \"{0}\".", list.RootFolder.ServerRelativeUrl);

                xmlWriter.WriteStartElement("List");
                xmlWriter.WriteAttributeString("WriteSecurity", list.WriteSecurity.ToString());
                xmlWriter.WriteAttributeString("ReadSecurity", list.ReadSecurity.ToString());
                xmlWriter.WriteAttributeString("AnonymousPermMask64", ((int)list.AnonymousPermMask64).ToString());
                xmlWriter.WriteAttributeString("AllowEveryoneViewItems", list.AllowEveryoneViewItems.ToString());
                xmlWriter.WriteAttributeString("Url", list.RootFolder.Url);

                // Write the security for the list itself.
                WriteObjectSecurity(list, xmlWriter);

                // Write the security for any folders in the list.
                WriteFolderSecurity(list, xmlWriter);

                // Write the security for any items in the list.
                if (includeItemSecurity)
                    WriteItemSecurity(list, xmlWriter);

                xmlWriter.WriteEndElement(); // List
            }
            finally
            {
                Logger.Write("Progress: Finished processing list \"{0}\".", list.RootFolder.ServerRelativeUrl);
            }
        }

        /// <summary>
        /// Writes the folder security.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="xmlWriter">The XML writer.</param>
        public static void WriteFolderSecurity(SPList list, XmlTextWriter xmlWriter)
        {
            foreach (SPListItem folder in list.Folders)
            {
                Logger.Write("Progress: Processing folder \"{0}\".", folder.Url);

                xmlWriter.WriteStartElement("Folder");
                xmlWriter.WriteAttributeString("Url", folder.Url);
                WriteObjectSecurity(folder, xmlWriter);
                xmlWriter.WriteEndElement(); // Folder
            }
        }

        /// <summary>
        /// Writes the item security.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="xmlWriter">The XML writer.</param>
        public static void WriteItemSecurity(SPList list, XmlTextWriter xmlWriter)
        {
            foreach (SPListItem item in list.Items)
            {
                Logger.Write("Progress: Processing item \"{0}\".", item.ID.ToString());

                xmlWriter.WriteStartElement("Item");
                xmlWriter.WriteAttributeString("Id", item.ID.ToString());
                WriteObjectSecurity(item, xmlWriter);
                xmlWriter.WriteEndElement(); // Item
            }
        }

        /// <summary>
        /// Writes the object security.
        /// </summary>
        /// <param name="sourceObject">The source object.</param>
        /// <param name="xmlWriter">The XML writer.</param>
        public static void WriteObjectSecurity(SPSecurableObject sourceObject, XmlTextWriter xmlWriter)
        {
            xmlWriter.WriteAttributeString("HasUniqueRoleAssignments", sourceObject.HasUniqueRoleAssignments.ToString());

            if (!sourceObject.HasUniqueRoleAssignments)
                return;

            //xmlWriter.WriteRaw(sourceObject.RoleAssignments.Xml);

            xmlWriter.WriteStartElement("RoleAssignments");
            foreach (SPRoleAssignment ra in sourceObject.RoleAssignments)
            {
                xmlWriter.WriteStartElement("RoleAssignment");
                xmlWriter.WriteAttributeString("Member", ra.Member.Name);

                SPPrincipalType pType = SPPrincipalType.None;
                if (ra.Member is SPUser)
                {
                    pType = SPPrincipalType.User;
                    xmlWriter.WriteAttributeString("LoginName", ((SPUser)ra.Member).LoginName);
                }
                else if (ra.Member is SPGroup)
                {
                    pType = SPPrincipalType.SharePointGroup;
                }

                xmlWriter.WriteAttributeString("PrincipalType", pType.ToString());

                xmlWriter.WriteStartElement("RoleDefinitionBindings");
                foreach (SPRoleDefinition rd in ra.RoleDefinitionBindings)
                {
                    if (rd.Name == "Limited Access")
                        continue;

                    xmlWriter.WriteStartElement("RoleDefinition");
                    xmlWriter.WriteAttributeString("Name", rd.Name);
                    xmlWriter.WriteAttributeString("Description", rd.Description);
                    xmlWriter.WriteAttributeString("Order", rd.Order.ToString());
                    xmlWriter.WriteAttributeString("BasePermissions", rd.BasePermissions.ToString());
                    xmlWriter.WriteEndElement(); //RoleDefinition
                }
                xmlWriter.WriteEndElement(); //RoleDefinitionBindings
                xmlWriter.WriteEndElement(); //RoleAssignment
            }
            xmlWriter.WriteEndElement(); //RoleAssignments
        }

    }
}
