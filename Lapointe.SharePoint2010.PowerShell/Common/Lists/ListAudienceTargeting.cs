using System;
using Microsoft.SharePoint;
using System.Xml;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    internal class ListAudienceTargeting
    {

        /// <summary>
        /// Sets the whether audience targeting is enabled or not.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="enabled">if set to <c>true</c> [enabled].</param>
        internal static void SetTargeting(string url, bool enabled)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb())
            {
                SPList list = Utilities.GetListFromViewUrl(web, url);

                SetTargeting(enabled, list);

            }
        }

        internal static void SetTargeting(bool enabled, SPList list)
        {
            if (list == null)
                throw new SPException("List was not found.");

            SPField targetingField = GetTargetingField(list);
            if (enabled && (targetingField == null))
            {
                string createFieldAsXml = CreateFieldAsXml();
                list.Fields.AddFieldAsXml(createFieldAsXml);
                list.Update();
            }
            else if (!enabled && (targetingField != null))
            {
                list.Fields.Delete(targetingField.InternalName);
                list.Update();
            }
        }

        /// <summary>
        /// Gets the targeting field.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <returns></returns>
        private static SPField GetTargetingField(SPList list)
        {
            SPField field = null;
            try
            {
                field = list.Fields[new Guid("61cbb965-1e04-4273-b658-eedaa662f48d")];
            }
            catch (ArgumentException)
            {
            }
            return field;
        }

        /// <summary>
        /// Gets the field as XML.
        /// </summary>
        /// <returns></returns>
        private static string CreateFieldAsXml()
        {
            XmlElement element = new XmlDocument().CreateElement("Field");
            element.SetAttribute("ID", "61cbb965-1e04-4273-b658-eedaa662f48d");
            element.SetAttribute("Type", "TargetTo");
            element.SetAttribute("Name", "TargetTo");
            element.SetAttribute("DisplayName", "Target Audiences");
            element.SetAttribute("Required", "FALSE");
            return element.OuterXml;
        }
    }
}
