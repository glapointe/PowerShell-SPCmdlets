using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    public class AddList
    {
        /// <summary>
        /// Adds a list to the web specified by the URL.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="urlName">Name of the URL.</param>
        /// <param name="title">The title.</param>
        /// <param name="desc">The desc.</param>
        /// <param name="featureId">The feature id.</param>
        /// <param name="templateType">Type of the template.</param>
        /// <param name="docTemplateType">Type of the doc template.</param>
        /// <returns></returns>
        public static SPList Add(string url, string urlName, string title, string desc, Guid featureId, int templateType, string docTemplateType)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
            {
                return Add(web.Lists, urlName, title, desc,
                                     featureId,
                                     templateType,
                                     docTemplateType);
            }
        }

        /// <summary>
        /// Adds a list to the specified list collection.
        /// </summary>
        /// <param name="lists">The lists collection to add the list to.</param>
        /// <param name="urlName">Name of the list to use in the URL.</param>
        /// <param name="title">The title of the list.</param>
        /// <param name="description">The description.</param>
        /// <param name="featureId">The feature id that the list template is associated with.</param>
        /// <param name="templateType">Type of the template.</param>
        /// <param name="docTemplateType">Type of the doc template.</param>
        /// <returns></returns>
        public static SPList Add(
            SPListCollection lists,
            string urlName,
            string title,
            string description,
            Guid featureId,
            int templateType,
            string docTemplateType)
        {
            if (docTemplateType == "")
                docTemplateType = null;

            Guid guid = lists.Add(title, description, urlName, featureId.ToString("D"), templateType, docTemplateType);
            return lists[guid];
        }

        
    }
}
