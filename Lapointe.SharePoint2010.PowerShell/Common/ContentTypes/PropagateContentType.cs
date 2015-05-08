using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Text;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.ContentTypes
{
    public static class PropagateContentType
    {

        /// <summary>
        /// Processes the content type.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="contentTypeName">Name of the content type.</param>
        /// <param name="verbose">if set to <c>true</c> [verbose].</param>
        /// <param name="updateFields">if set to <c>true</c> [update fields].</param>
        /// <param name="removeFields">if set to <c>true</c> [remove fields].</param>
        public static void Execute(SPSite site, string contentTypeName, bool updateFields, bool removeFields)
        {
            try
            {
                Logger.Write("Pushing content type changes to lists for '" + contentTypeName + "'");
                // get the site collection specified
                using (SPWeb rootWeb = site.RootWeb)
                {

                    //Get the source site content type
                    SPContentType sourceCT = rootWeb.AvailableContentTypes[contentTypeName];
                    if (sourceCT == null)
                    {
                        throw new ArgumentException("Unable to find Content Type named \"" + contentTypeName + "\"");
                    }
                    Execute(sourceCT, updateFields, removeFields);
                }
                return;
            }
            catch (Exception ex)
            {
                Logger.WriteException(new System.Management.Automation.ErrorRecord(new SPException("Unhandled error occured.", ex), null, System.Management.Automation.ErrorCategory.NotSpecified, null));
                throw;
            }
            finally
            {
                Logger.Write("Finished pushing content type changes to lists for '" + contentTypeName + "'");
            }
        }

        public static void Execute(SPContentType sourceCT, bool updateFields, bool removeFields)
        {
            using (SPSite site = new SPSite(sourceCT.ParentWeb.Site.ID))
            {
                IList<SPContentTypeUsage> ctUsageList = SPContentTypeUsage.GetUsages(sourceCT);
                foreach (SPContentTypeUsage ctu in ctUsageList)
                {
                    if (!ctu.IsUrlToList)
                        continue;

                    SPWeb web = null;
                    try
                    {
                        try
                        {
                            string webUrl = ctu.Url;
                            if (webUrl.ToLowerInvariant().Contains("_catalogs/masterpage"))
                                webUrl = webUrl.Substring(0, webUrl.IndexOf("/_catalogs/masterpage"));

                            web = site.OpenWeb(webUrl);
                        }
                        catch (SPException ex)
                        {
                            Logger.WriteWarning("Unable to open host web of content type\r\n{0}", ex.Message);
                            continue;
                        }
                        if (web != null)
                        {

                            SPList list = web.GetList(ctu.Url);
                            SPContentType listCT = list.ContentTypes[ctu.Id];
                            ProcessContentType(list, sourceCT, listCT, updateFields, removeFields);
                        }
                    }
                    finally
                    {
                        if (web != null)
                            web.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// Processes the content type.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="sourceCT">The source CT.</param>
        /// <param name="listCT">The list CT.</param>
        /// <param name="updateFields">if set to <c>true</c> [update fields].</param>
        /// <param name="removeFields">if set to <c>true</c> [remove fields].</param>
        private static void ProcessContentType(SPList list, SPContentType sourceCT, SPContentType listCT, bool updateFields, bool removeFields)
        {
            if (listCT == null)
                return;

            if (listCT.ReadOnly)
            {
                Logger.WriteWarning("Unable to update read-only content type ({0}: {1})", listCT.Name, list.RootFolder.ServerRelativeUrl);
                return;
            }

            if (listCT.Sealed)
            {
                Logger.WriteWarning("Unable to update read-only content type ({0}: {1})", listCT.Name, list.RootFolder.ServerRelativeUrl);
                return;
            }

            Logger.Write("PROGRESS: Processing content type on list:" + list.RootFolder.ServerRelativeUrl);

            if (updateFields)
            {
                UpdateListFields(list, listCT, sourceCT);
            }

            if (removeFields)
            {
                //Find the fields to delete

                //Copy collection to avoid modifying enumeration as we go through it
                List<SPFieldLink> listFieldLinks = new List<SPFieldLink>();
                foreach (SPFieldLink listFieldLink in listCT.FieldLinks)
                {
                    listFieldLinks.Add(listFieldLink);
                }

                foreach (SPFieldLink listFieldLink in listFieldLinks)
                {
                    if (!FieldExist(sourceCT, listFieldLink))
                    {
                        Logger.Write("PROGRESS: Removing field \"{0}\" from Content Type on \"{1}\"...", listFieldLink.Name, list.RootFolder.ServerRelativeUrl);
                        listCT.FieldLinks.Delete(listFieldLink.Id);
                        listCT.Update();
                    }
                }
            }

            //Find/add the fields to add
            foreach (SPFieldLink sourceFieldLink in sourceCT.FieldLinks)
            {
                if (!FieldExist(sourceCT, sourceFieldLink))
                {
                    Logger.WriteWarning("Failed to add field \"{0}\" on list \"{1}\" field does not exist (in .Fields[]) on source Content Type", sourceFieldLink.Name, list.RootFolder.ServerRelativeUrl);
                }
                else
                {
                    if (!FieldExist(listCT, sourceFieldLink))
                    {
                        //Perform double update, just to be safe
                        // (but slow)
                        Logger.Write("PROGRESS: Adding field \"{0}\" to Content Type on \"{1}\"...", sourceFieldLink.Name, list.RootFolder.ServerRelativeUrl);
                        if (listCT.FieldLinks[sourceFieldLink.Id] != null)
                        {
                            listCT.FieldLinks.Delete(sourceFieldLink.Id);
                            listCT.Update();
                        }
                        if (!list.Fields.ContainsField(sourceFieldLink.Name))
                        {
                            list.Fields.Add(sourceCT.Fields.GetFieldByInternalName(sourceFieldLink.Name));
                            list.Update();
                        }
                        listCT.FieldLinks.Add(new SPFieldLink(sourceCT.Fields[sourceFieldLink.Id]));
                        listCT.Update();
                    }
                }
            }
            listCT.DocumentTemplate = sourceCT.DocumentTemplate;

            // Reorder the fields.
            try
            {
                Logger.Write("PROGRESS: Reordering fields...");
                // Store the field order so that we can reorder.
                List<string> fields = new List<string>();
                foreach (SPField field in sourceCT.Fields)
                {
                    if (!field.Hidden && field.Reorderable)
                        fields.Add(field.InternalName);
                }
                listCT.FieldLinks.Reorder(fields.ToArray());
                Logger.Write("PROGRESS: Finished reordering fields.");
            }
            catch
            {
                Logger.WriteWarning("Unable to set field order.");
            }
            listCT.Update();
        }

        /// <summary>
        /// Updates the fields of the list content type (listCT) with the
        /// fields found on the source content type (courceCT).
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="listCT">The list CT.</param>
        /// <param name="sourceCT">The source CT.</param>
        private static void UpdateListFields(SPList list, SPContentType listCT, SPContentType sourceCT)
        {
            Logger.Write("PROGRESS: Starting to update fields...");
            foreach (SPFieldLink sourceFieldLink in sourceCT.FieldLinks)
            {
                //has the field changed? If not, continue.
                if (listCT.FieldLinks[sourceFieldLink.Id] != null && listCT.FieldLinks[sourceFieldLink.Id].SchemaXml == sourceFieldLink.SchemaXml)
                {
                    Logger.Write("PROGRESS: Doing nothing to field \"{0}\" from Content Type on: \"{1}\"", sourceFieldLink.Name, list.RootFolder.ServerRelativeUrl);
                    continue;
                }

                if (!FieldExist(sourceCT, sourceFieldLink))
                {
                    Logger.Write("PROGRESS: Doing nothing to field: \"{0}\" on list \"{1}\" field does not exist (in .Fields[]) on source Content Type.", sourceFieldLink.Name, list.RootFolder.ServerRelativeUrl);
                    continue;

                }

                if (listCT.FieldLinks[sourceFieldLink.Id] != null)
                {

                    Logger.Write("PROGRESS: Deleting field \"{0}\" from Content Type on \"{1}\"...", sourceFieldLink.Name, list.RootFolder.ServerRelativeUrl);

                    listCT.FieldLinks.Delete(sourceFieldLink.Id);
                    listCT.Update();
                }

                Logger.Write("PROGRESS: Adding field \"{0}\" from Content Type on \"{1}\"...", sourceFieldLink.Name, list.RootFolder.ServerRelativeUrl);

                listCT.FieldLinks.Add(new SPFieldLink(sourceCT.Fields[sourceFieldLink.Id]));
                //Set displayname, not set by previous operation
                listCT.FieldLinks[sourceFieldLink.Id].DisplayName = sourceCT.FieldLinks[sourceFieldLink.Id].DisplayName;
                listCT.Update();
                Logger.Write("PROGRESS: Done updating fields.");
            }
        }

        /// <summary>
        /// Fields the exist.
        /// </summary>
        /// <param name="contentType">Type of the content.</param>
        /// <param name="fieldLink">The field link.</param>
        /// <returns></returns>
        private static bool FieldExist(SPContentType contentType, SPFieldLink fieldLink)
        {
            try
            {
                //will throw exception on missing fields
                return contentType.Fields[fieldLink.Id] != null;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
