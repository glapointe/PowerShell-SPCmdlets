using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.ContentTypes
{
    class ExportContentTypes
    {
        /// <summary>
        /// Exports the specified content type(s).
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="contentTypeGroup">The content type group.</param>
        /// <param name="contentTypeName">Name of the content type.</param>
        /// <param name="excludeParentFields">if set to <c>true</c> [exclude parent fields].</param>
        /// <param name="includeFieldDefinitions">if set to <c>true</c> [include field definitions].</param>
        /// <param name="includeListBindings">if set to <c>true</c> [include list bindings].</param>
        /// <param name="listName">Name of the list.</param>
        /// <param name="removeEncodedSpaces">if set to <c>true</c> [remove encoded spaces].</param>
        /// <param name="featureSafe">if set to <c>true</c> [feature safe].</param>
        /// <param name="outputFile">The output file.</param>
        public static void Export(string url, string contentTypeGroup, string contentTypeName, bool excludeParentFields, bool includeFieldDefinitions, bool includeListBindings, string listName, bool removeEncodedSpaces, bool featureSafe, string outputFile)
        {
            StringBuilder sb = new StringBuilder();
            XmlTextWriter xmlWriter = OpenXmlWriter(sb);

            try
            {
                Export(url, contentTypeGroup, contentTypeName, excludeParentFields, includeFieldDefinitions, includeListBindings, listName, removeEncodedSpaces, featureSafe, xmlWriter);
            }
            finally
            {
                CloseXmlWriter(xmlWriter, outputFile, sb);
            }
        }

        public static void Export(string url, string contentTypeGroup, string contentTypeName, bool excludeParentFields, bool includeFieldDefinitions, bool includeListBindings, string listName, bool removeEncodedSpaces, bool featureSafe, XmlTextWriter xmlWriter)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            {
                Export(web, contentTypeGroup, contentTypeName, excludeParentFields, includeFieldDefinitions, includeListBindings, listName, removeEncodedSpaces, featureSafe, xmlWriter);
            }
        }

        public static void Export(SPWeb[] webs, string contentTypeGroup, string contentTypeName, bool excludeParentFields, bool includeFieldDefinitions, bool includeListBindings, string listName, bool removeEncodedSpaces, bool featureSafe, string outputFile)
        {
            StringBuilder sb = new StringBuilder();
            XmlTextWriter xmlWriter = OpenXmlWriter(sb);

            try
            {
                for (int i = 0; i < webs.Length; i++)
                {
                    Export(webs[i], contentTypeGroup, contentTypeName, excludeParentFields, includeFieldDefinitions, includeListBindings, listName, removeEncodedSpaces, featureSafe, xmlWriter);
                }
            }
            finally
            {
                CloseXmlWriter(xmlWriter, outputFile, sb);
            }
        }

        public static void Export(SPWeb web, string contentTypeGroup, string contentTypeName, bool excludeParentFields, bool includeFieldDefinitions, bool includeListBindings, string listName, bool removeEncodedSpaces, bool featureSafe, XmlTextWriter xmlWriter)
        {
            Dictionary<Guid, SPField> ctFields = new Dictionary<Guid, SPField>();

            SPContentTypeCollection availableContentTypes;

            if (listName != null)
            {
                SPList list = web.Lists[listName];
                availableContentTypes = list.ContentTypes;
            }
            else
            {
                availableContentTypes = web.AvailableContentTypes;
            }
            List<SPContentType> contentTypes = new List<SPContentType>();

            // Gather up all the content types we want to export out.
            if (contentTypeName != null)
            {
                SPContentType ct = availableContentTypes[contentTypeName];
                if (ct == null)
                {
                    throw new SPException("The content type specified could not be found.");
                }
                else
                {
                    contentTypes.Add(ct);
                }
            }
            else
            {
                // Loop through all the source content types and create them at the target.
                foreach (SPContentType ct in availableContentTypes)
                {
                    if (contentTypeGroup == null || ct.Group.ToLowerInvariant() == contentTypeGroup.ToLowerInvariant())
                    {
                        contentTypes.Add(ct);
                    }
                }
            }

            if (includeFieldDefinitions)
            {
                // If we're including field definitions then we want to show them first as they'll need to appear first when using within a Feature
                foreach (SPContentType ct in contentTypes)
                {
                    SPContentType parentCT = ct.Parent;
                    foreach (SPField field in ct.Fields)
                    {
                        // If the parent content type contains the current field and the user wants to exclude parent fields then continue to the next field.
                        if (parentCT != null && excludeParentFields && parentCT.Fields.ContainsField(field.InternalName))
                            continue;

                        if (!ctFields.ContainsKey(field.Id))
                            ctFields.Add(field.Id, field);
                    }
                }
                xmlWriter.WriteString("\r\n\r\n");
                foreach (SPField field in ctFields.Values)
                {
                    xmlWriter.WriteRaw(Utilities.GetFieldSchema(field, featureSafe, removeEncodedSpaces));
                    xmlWriter.WriteString("\r\n");
                }
            }

            xmlWriter.WriteString("\r\n");

            foreach (SPContentType ct in contentTypes)
            {
                WriteContentTypeXml(ct, excludeParentFields, removeEncodedSpaces, xmlWriter, featureSafe);
            }

            if (includeListBindings)
                GetListBindings(web, contentTypes, xmlWriter);
        }




        /// <summary>
        /// Writes the content type XML.
        /// </summary>
        /// <param name="ct">The ct.</param>
        /// <param name="excludeParentFields">if set to <c>true</c> [exclude parent fields].</param>
        /// <param name="removeEncodedSpaces">if set to <c>true</c> [remove encoded spaces].</param>
        /// <param name="xmlWriter">The XML writer.</param>
        /// <param name="featureSafe">if set to <c>true</c> [feature safe].</param>
        private static void WriteContentTypeXml(SPContentType ct, bool excludeParentFields, bool removeEncodedSpaces, XmlTextWriter xmlWriter, bool featureSafe)
        {
            SPContentType parentCT = ct.Parent;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(ct.SchemaXml);
            if (featureSafe)
            {
                xmlDoc.DocumentElement.SetAttribute("Version", "0");
                xmlDoc.DocumentElement.SetAttribute("Inherits", "TRUE");
            }
            XmlElement fieldRefsElement = xmlDoc.CreateElement("FieldRefs");
            xmlDoc.DocumentElement.AppendChild(fieldRefsElement);
            foreach (XmlElement field in xmlDoc.SelectNodes("//Fields/Field"))
            {
                // If the parent content type contains the current field and the user wants to exclude parent fields then continue to the next field.
                if (parentCT != null && excludeParentFields && parentCT.Fields.ContainsField(field.GetAttribute("Name")))
                    continue;

                XmlElement fieldRefElement = xmlDoc.CreateElement("FieldRef");
                fieldRefElement.SetAttribute("ID", field.GetAttribute("ID"));

                string name = field.GetAttribute("Name");
                if (name.Contains(Utilities.ENCODED_SPACE) && removeEncodedSpaces)
                    name = name.Replace(Utilities.ENCODED_SPACE, string.Empty);

                fieldRefElement.SetAttribute("Name", name);
                fieldRefsElement.AppendChild(fieldRefElement);
            }
            xmlDoc.DocumentElement.RemoveChild(xmlDoc.SelectSingleNode("//Fields"));

            xmlWriter.WriteString("\r\n");
            xmlWriter.WriteRaw(xmlDoc.OuterXml);
        }

        /// <summary>
        /// Gets the list bindings.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="contentTypes">The content types.</param>
        /// <param name="xmlWriter">The XML writer.</param>
        private static void GetListBindings(SPWeb web, List<SPContentType> contentTypes, XmlTextWriter xmlWriter)
        {
            foreach (SPList list in web.Lists)
            {
                foreach (SPContentType listCT in list.ContentTypes)
                {
                    foreach (SPContentType ct in contentTypes)
                    {
                        //if ((listCT.Scope != ct.Scope && listCT.Parent.Id == ct.Id) || listCT.Id == ct.Id)
                        if (listCT.Name == ct.Name && listCT.Group == ct.Group)
                        {
                            xmlWriter.WriteStartElement("ContentTypeBinding");
                            xmlWriter.WriteAttributeString("ContentTypeId", ct.Id.ToString());
                            xmlWriter.WriteAttributeString("ListUrl", list.RootFolder.ServerRelativeUrl);
                            xmlWriter.WriteEndElement(); // ContentTypeBinding
                        }
                    }
                }
            }

            foreach (SPWeb subWeb in web.Webs)
            {
                try
                {
                    GetListBindings(subWeb, contentTypes, xmlWriter);
                }
                finally
                {
                    subWeb.Dispose();
                }
            }
        }

        public static XmlTextWriter OpenXmlWriter(StringBuilder sb)
        {
            XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
            xmlWriter.Formatting = Formatting.Indented;

            xmlWriter.WriteStartElement("Elements");
            xmlWriter.WriteAttributeString("xmlns", "http://schemas.microsoft.com/sharepoint/");

            return xmlWriter;
        }

        public static void CloseXmlWriter(XmlTextWriter xmlWriter, string outputFile, StringBuilder sb)
        {
            xmlWriter.WriteEndElement(); // Elements

            xmlWriter.Flush();
            xmlWriter.Close();

            if (!string.IsNullOrEmpty(outputFile))
                File.WriteAllText(outputFile, sb.ToString());
        }
    }
}
