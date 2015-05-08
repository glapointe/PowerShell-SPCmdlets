using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
#if MOSS
using Microsoft.Office.RecordsManagement.InformationPolicy;
#endif
using Microsoft.SharePoint;
using Microsoft.SharePoint.StsAdmin;
using Microsoft.SharePoint.Workflow;
using System.Runtime.InteropServices;

namespace Lapointe.SharePoint.PowerShell.Common.ContentTypes
{
    class CopyContentTypes
    {
        private bool _copyWorkflows = true;
        private bool _copyColumns = true;
        private bool _copyDocConversions = true;
        private bool _copyDocInfoPanel = true;
        private bool _copyPolicies = true;
        private bool _copyDocTemplate = true;

        private SPContentTypeCollection _availableTargetContentTypes;
        private SPContentTypeCollection _targetContentTypes;
        private SPFieldCollection _targetFields;

        public CopyContentTypes()
        {
        }

        public CopyContentTypes(bool copyWorkflows, bool copyColumns, bool copyDocConversions, bool copyDocInfoPanel, bool copyPolicies, bool copyDocTemplate)
        {
            _copyWorkflows = copyWorkflows;
            _copyColumns = copyColumns;
            _copyDocConversions = copyDocConversions;
            _copyDocInfoPanel = copyDocInfoPanel;
            _copyPolicies = copyPolicies;
            _copyDocTemplate = copyDocTemplate;
        }

        public void Copy(string sourceUrl, string targetUrl)
        {
            Copy(sourceUrl, targetUrl, null);
        }

        public void Copy(SPContentType sourceCT, SPWeb target)
        {
            Copy(sourceCT, target.Url);
        }

        public void Copy(SPContentType sourceCT, string targetUrl)
        {
            // Make sure the source exists if it was specified.
            if (sourceCT == null)
            {
                throw new SPException("The source content type could not be found.");
            }

            SPFieldCollection sourceFields = sourceCT.ParentWeb.Fields;

            // Get the target content type and fields.
            GetAvailableTargetContentTypes(targetUrl);
            if (_availableTargetContentTypes[sourceCT.Name] == null)
            {
                Logger.Write("Progress: Source content type '{0}' does not exist on target - creating content type...", sourceCT.Name);

                CreateContentType(targetUrl, sourceCT, sourceFields);
            }
            else
            {
                throw new SPException("Content type already exists at target.");
            }
        }

        public void Copy(SPWeb source, SPWeb target, string sourceContentTypeName)
        {
            Copy(source.Url, target.Url, sourceContentTypeName);
        }

        public void Copy(string sourceUrl, string targetUrl, string sourceContentTypeName)
        {
            SPContentTypeCollection availableSourceContentTypes;
            SPFieldCollection sourceFields;

            // Get the source content type and fields.
            using (SPSite site = new SPSite(sourceUrl))
            {
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(sourceUrl)])
                {
                    Logger.Write("Progress: Getting existing source fields and content types...");

                    availableSourceContentTypes = web.AvailableContentTypes;
                    sourceFields = web.Fields;

                    // Get the target content type and fields.
                    GetAvailableTargetContentTypes(targetUrl);

                    // Make sure the source exists if it was specified.
                    if (sourceContentTypeName != null)
                    {
                        SPContentType sourceCT = availableSourceContentTypes[sourceContentTypeName];
                        if (sourceCT == null)
                        {
                            throw new SPException("The source content type could not be found.");
                        }
                        Logger.Write("Progress: Source content type found.");

                        if (_availableTargetContentTypes[sourceCT.Name] == null)
                        {
                            Logger.Write("Progress: Source content type '{0}' does not exist on target - creating content type...", sourceCT.Name);

                            CreateContentType(targetUrl, sourceCT, sourceFields);
                        }
                        else
                        {
                            throw new SPException("Content type already exists at target.");
                        }
                    }
                    else
                    {
                        // Loop through all the source content types and create them at the target.
                        foreach (SPContentType sourceCT in availableSourceContentTypes)
                        {
                            if (_availableTargetContentTypes[sourceCT.Name] == null)
                            {
                                Logger.Write("Progress: Source content type '{0}' does not exist on target - creating content type...", sourceCT.Name);

                                CreateContentType(
                                    targetUrl,
                                    sourceCT, 
                                    sourceFields);

                                // Reset the fields and content types.
                                GetAvailableTargetContentTypes(targetUrl);
                            }
                            else
                            {
                                Logger.Write("Progress: Source content type '{0}' exists on target - skipping.", sourceCT.Name);
                            }
                        }
                    }
                }
            }

        }

        private void GetAvailableTargetContentTypes(string targetUrl)
        {
            using (SPSite site = new SPSite(targetUrl))
            {
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(targetUrl)])
                {
                    Logger.Write("Progress: Getting existing target fields and content types...");
                    
                    _availableTargetContentTypes = web.AvailableContentTypes;
                    _targetContentTypes = web.ContentTypes;
                    _targetFields = web.Fields;
                }
            }
        }

        /// <summary>
        /// Creates the content type.
        /// </summary>
        /// <param name="targetUrl">The target URL.</param>
        /// <param name="sourceCT">The source content type.</param>
        /// <param name="sourceFields">The source fields.</param>
        private void CreateContentType(string targetUrl, SPContentType sourceCT, SPFieldCollection sourceFields)
        {
            // Make sure any parent content types exist - they have to be there before we can create this content type.
            if (_availableTargetContentTypes[sourceCT.Parent.Id] == null)
            {
                Logger.Write("Progress: Parent of content type '{0}' does not exist - creating...", sourceCT.Name);

                CreateContentType(targetUrl, sourceCT.Parent, sourceFields);

                // Reset the fields and content types.
                GetAvailableTargetContentTypes(targetUrl);
            }

            Logger.Write("Progress: Creating content type '{0}'...", sourceCT.Name);

            // Create a new content type using information from the source content type.
            SPContentType newCT = new SPContentType(sourceCT.Id, _targetContentTypes, sourceCT.Name);

            Logger.Write("Progress: Setting fields for content type '{0}'...", sourceCT.Name);

            // Set all the core properties for the content type.
            newCT.Group = sourceCT.Group;
            newCT.Hidden = sourceCT.Hidden;
            newCT.NewDocumentControl = sourceCT.NewDocumentControl;
            newCT.NewFormTemplateName = sourceCT.NewFormTemplateName;
            newCT.NewFormUrl = sourceCT.NewFormUrl;
            newCT.ReadOnly = sourceCT.ReadOnly;
            newCT.RequireClientRenderingOnNew = sourceCT.RequireClientRenderingOnNew;
            newCT.Description = sourceCT.Description;
            newCT.DisplayFormTemplateName = sourceCT.DisplayFormTemplateName;
            newCT.DisplayFormUrl = sourceCT.DisplayFormUrl;
            newCT.EditFormTemplateName = sourceCT.EditFormTemplateName;
            newCT.EditFormUrl = sourceCT.EditFormUrl;
            newCT.MobileDisplayFormUrl = sourceCT.MobileDisplayFormUrl;
            newCT.MobileEditFormUrl = sourceCT.MobileEditFormUrl;
            newCT.MobileNewFormUrl = sourceCT.MobileNewFormUrl;
            newCT.RequireClientRenderingOnNew = sourceCT.RequireClientRenderingOnNew;

            Logger.Write("Progress: Adding content type '{0}' to collection...", sourceCT.Name);

            // Add the content type to the content types collection and update all the settings.
            _targetContentTypes.Add(newCT);
            newCT.Update();

            // Add all the peripheral items

            try
            {
                if (_copyColumns)
                {
                    Logger.Write("Progress: Adding site columns for content type '{0}'...", sourceCT.Name);

                    AddSiteColumns(newCT, sourceCT, sourceFields);
                }

                if (_copyWorkflows)
                {
                    Logger.Write("Progress: Adding workflow associations for content type '{0}'...", sourceCT.Name);

                    AddWorkflowAssociations(newCT, sourceCT);
                }

                if (_copyDocTemplate)
                {
                    Logger.Write("Progress: Adding document template for content type '{0}'...", sourceCT.Name);

                    AddDocumentTemplate(newCT, sourceCT);
                }

                if (_copyDocConversions)
                {
                    Logger.Write("Progress: Adding document conversion settings for content type '{0}'...", sourceCT.Name);

                    AddDocumentConversionSettings(newCT, sourceCT);
                }

                if (_copyPolicies)
                {
                    Logger.Write("Progress: Adding information rights policies for content type '{0}'...", sourceCT.Name);

                    AddInformationRightsPolicies(newCT, sourceCT);
                }

                if (_copyDocInfoPanel)
                {
                    Logger.Write("Progress: Adding document information panel for content type '{0}'...", sourceCT.Name);

                    AddDocumentInfoPanelToContentType(sourceCT, newCT);
                }
            }
            finally
            {
                newCT.ParentWeb.Site.Dispose();
                newCT.ParentWeb.Dispose();
            }
        }

        /// <summary>
        /// Adds the document template.
        /// </summary>
        /// <param name="targetCT">The target content type.</param>
        /// <param name="sourceCT">The source content type.</param>
        private void AddDocumentTemplate(SPContentType targetCT, SPContentType sourceCT)
        {
            if (string.IsNullOrEmpty(sourceCT.DocumentTemplate))
                return;

            // Add the document template.
            SPFile sourceFile = null;
            try
            {
                sourceFile = sourceCT.ResourceFolder.Files[sourceCT.DocumentTemplate];
            }
            catch (ArgumentException) {}
            if (sourceFile != null && !string.IsNullOrEmpty(sourceFile.Name))
            {
                SPFile targetFile = targetCT.ResourceFolder.Files.Add(sourceFile.Name, sourceFile.OpenBinary(), true);
                targetCT.DocumentTemplate = targetFile.Name;
                targetCT.Update();
            }
            else
            {
                targetCT.DocumentTemplate = sourceCT.DocumentTemplate;
                targetCT.Update();
            }
        }

        /// <summary>
        /// Adds the workflow associations.
        /// </summary>
        /// <param name="targetCT">The target content type.</param>
        /// <param name="sourceCT">The source content type.</param>
        private void AddWorkflowAssociations(SPContentType targetCT, SPContentType sourceCT)
        {
            // Remove the default workflows - we're going to add from the source.
            while (targetCT.WorkflowAssociations.Count > 0)
            {
                targetCT.WorkflowAssociations.Remove(targetCT.WorkflowAssociations[0]);
            }

            // Add workflows.
            foreach (SPWorkflowAssociation wf in sourceCT.WorkflowAssociations)
            {
                Logger.Write("Progress: Adding workflow '{0}' to content type...", wf.Name);
                targetCT.WorkflowAssociations.Add(SPWorkflowAssociation.ImportFromXml(targetCT.ParentWeb.Site.RootWeb, wf.ExportToXml()));
            }
            targetCT.Update();
        }

        /// <summary>
        /// Adds the site columns.
        /// </summary>
        /// <param name="targetCT">The target content type.</param>
        /// <param name="sourceCT">The source content type.</param>
        /// <param name="sourceFields">The source fields.</param>
        private void AddSiteColumns(SPContentType targetCT, SPContentType sourceCT, SPFieldCollection sourceFields)
        {
            // Store the field order so that we can reorder after adding all the fields.
            List<string> fields = new List<string>();
            foreach (SPField field in sourceCT.Fields)
            {
                if (!field.Hidden && field.Reorderable)
                    fields.Add(field.InternalName);
            }
            // Add any columns associated with the content type.
            foreach (SPFieldLink field in sourceCT.FieldLinks)
            {
                // First we need to see if the column already exists as a Site Column.
                SPField sourceField;
                try
                {
                    // First try and find the column via the ID
                    sourceField = sourceFields[field.Id];
                }
                catch
                {
                    try
                    {
                        // Couldn't locate via ID so now try the name
                        sourceField = sourceFields[field.Name];
                    }
                    catch
                    {
                        sourceField = null;
                    }
                }
                if (sourceField == null)
                {
                    // Couldn't locate by ID or name - it could be due to casing issues between the linked version of the name and actual field
                    // (for example, the Contact content type stores the name for email differently: EMail for the field and Email for the link)
                    foreach (SPField f in sourceCT.Fields)
                    {
                        if (field.Name.ToLowerInvariant() == f.InternalName.ToLowerInvariant())
                        {
                            sourceField = f;
                            break;
                        }
                    }
                }
                if (sourceField == null)
                {
                    Logger.WriteWarning("Unable to add column '{0}' to content type.", field.Name);
                    continue;
                }

                if (!_targetFields.ContainsField(sourceField.InternalName))
                {
                    Logger.Write("Progress: Adding column '{0}' to site columns...", sourceField.InternalName);

                    // The column does not exist so add the Site Column.
                    _targetFields.Add(sourceField);
                }

                // Now that we know the column exists we can add it to our content type.
                if (targetCT.FieldLinks[sourceField.InternalName] == null) // This should always be true if we're here but I'm keeping it in as a safety check.
                {
                    Logger.Write("Progress: Associating content type with site column '{0}'...", sourceField.InternalName);
                    
                    // Now add the reference to the site column for this content type.
                    try
                    {
                        targetCT.FieldLinks.Add(field);
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteWarning("Unable to add field '{0}' to content type: {1}", sourceField.InternalName, ex.Message);
                    }
                }
            }
            // Save the fields so that we can reorder them.
            targetCT.Update(true);

            // Reorder the fields.
            try
            {
                targetCT.FieldLinks.Reorder(fields.ToArray());
            }
            catch
            {
                Logger.WriteWarning("Unable to set field order.");
            }
            targetCT.Update(true);
        }

        /// <summary>
        /// Adds the information rights policies.
        /// </summary>
        /// <param name="targetCT">The target content type.</param>
        /// <param name="sourceCT">The source content type.</param>
        private void AddInformationRightsPolicies(SPContentType targetCT, SPContentType sourceCT)
        {
#if MOSS
            // Set information rights policy - must be done after the new content type is added to the collection.
            using (Policy sourcePolicy = Policy.GetPolicy(sourceCT))
            {
                if (sourcePolicy != null)
                {
                    PolicyCatalog catalog = new PolicyCatalog(targetCT.ParentWeb.Site);
                    PolicyCollection policyList = catalog.PolicyList;

                    Policy tempPolicy = null;
                    try
                    {
                        tempPolicy = policyList[sourcePolicy.Id];
                        if (tempPolicy == null)
                        {
                            XmlDocument exportedSourcePolicy = sourcePolicy.Export();
                            try
                            {
                                Logger.Write("Progress: Adding policy '{0}' to content type...", sourcePolicy.Name);

                                PolicyCollection.Add(targetCT.ParentWeb.Site, exportedSourcePolicy.OuterXml);
                            }
                            catch (Exception ex)
                            {
                                if (ex is NullReferenceException || ex is SEHException)
                                    throw;
                                // Policy already exists
                                Logger.WriteException(new System.Management.Automation.ErrorRecord(new SPException("An error occured creating the information rights policy: {0}", ex), null, System.Management.Automation.ErrorCategory.NotSpecified, exportedSourcePolicy));
                            }
                        }
                        Logger.Write("Progress: Associating content type with policy '{0}'...", sourcePolicy.Name);
                        // Associate the policy with the content type.
                        Policy.CreatePolicy(targetCT, sourcePolicy);
                    }
                    finally
                    {
                        if (tempPolicy != null)
                            tempPolicy.Dispose();
                    }
                }
                targetCT.Update();
            }
#endif
        }

        /// <summary>
        /// Adds the document conversion settings.
        /// </summary>
        /// <param name="targetCT">The target content type.</param>
        /// <param name="sourceCT">The source content type.</param>
        private void AddDocumentConversionSettings(SPContentType targetCT, SPContentType sourceCT)
        {
            // Add document conversion settings if document converisons are enabled.
            // ParentWeb and Site will be disposed later.
            if (targetCT.ParentWeb.Site.WebApplication.DocumentConversionsEnabled)
            {
                // First, handle the xml that describes what is enabled and each setting
                string sourceDocConversionXml = sourceCT.XmlDocuments["urn:sharePointPublishingRcaProperties"];
                if (sourceDocConversionXml != null)
                {
                    XmlDocument sourceDocConversionXmlDoc = new XmlDocument();
                    sourceDocConversionXmlDoc.LoadXml(sourceDocConversionXml);

                    targetCT.XmlDocuments.Delete("urn:sharePointPublishingRcaProperties");
                    targetCT.XmlDocuments.Add(sourceDocConversionXmlDoc);
                }
                // Second, handle the xml that describes what is excluded (disabled).
                sourceDocConversionXml =
                    sourceCT.XmlDocuments["http://schemas.microsoft.com/sharepoint/v3/contenttype/transformers"];
                if (sourceDocConversionXml != null)
                {
                    XmlDocument sourceDocConversionXmlDoc = new XmlDocument();
                    sourceDocConversionXmlDoc.LoadXml(sourceDocConversionXml);

                    targetCT.XmlDocuments.Delete("http://schemas.microsoft.com/sharepoint/v3/contenttype/transformers");
                    targetCT.XmlDocuments.Add(sourceDocConversionXmlDoc);
                }
                targetCT.Update();
            }
        }

        /// <summary>
        /// Adds the document info panel to the content type.
        /// </summary>
        /// <param name="sourceCT">The source content type.</param>
        /// <param name="targetCT">The target content type.</param>
        private void AddDocumentInfoPanelToContentType(SPContentType sourceCT, SPContentType targetCT)
        {
            XmlDocument sourceXmlDoc = null;
            string sourceXsnLocation;
            bool isCached;
            bool openByDefault;
            string scope;

            // We first need to get the XML which describes the custom information panel.
            string str = sourceCT.XmlDocuments["http://schemas.microsoft.com/office/2006/metadata/customXsn"];
            if (!string.IsNullOrEmpty(str))
            {
                sourceXmlDoc = new XmlDocument();
                sourceXmlDoc.LoadXml(str);
            }
            if (sourceXmlDoc != null)
            {
                // We found settings for a custom information panel so grab those settings for later use.
                XmlNode node;
                string innerText;
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(sourceXmlDoc.NameTable);
                nsmgr.AddNamespace("cust", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
                sourceXsnLocation = sourceXmlDoc.SelectSingleNode("/cust:customXsn/cust:xsnLocation", nsmgr).InnerText;
                node = sourceXmlDoc.SelectSingleNode("/cust:customXsn/cust:cached", nsmgr);
                isCached = (node != null) && (node.InnerText == bool.TrueString);
                innerText = sourceXmlDoc.SelectSingleNode("/cust:customXsn/cust:openByDefault", nsmgr).InnerText;
                openByDefault = !string.IsNullOrEmpty(innerText) && (innerText == bool.TrueString);
            }
            else
                return;

            // This should never be null but just in case...
            if (sourceXsnLocation == null)
                return;

            // Grab the source file and add it to the target resource folder.
            SPFile sourceFile = null;
            try
            {
                sourceFile = sourceCT.ResourceFolder.Files[sourceXsnLocation];
            }
            catch (ArgumentException)
            {
            }
            if (sourceFile != null)
            {
                SPFile file2 = targetCT.ResourceFolder.Files.Add(targetCT.ParentWeb.Url + "/" + sourceFile.Url, sourceFile.OpenBinary(), true);

                // Get the target and scope to use in the xsn for the custom information panel.
                string targetXsnLocation = targetCT.ParentWeb.Url + "/" + file2.Url;
                scope = targetCT.ParentWeb.Site.MakeFullUrl(targetCT.Scope);

                XmlDocument targetXmlDoc = BuildCustomInformationPanelXml(targetXsnLocation, isCached, openByDefault, scope);
                // Delete the existing doc so that we can add the new one.
                targetCT.XmlDocuments.Delete("http://schemas.microsoft.com/office/2006/metadata/customXsn");
                targetCT.XmlDocuments.Add(targetXmlDoc);
            }
            targetCT.Update();
        }

        /// <summary>
        /// Builds the custom information panel XML.
        /// </summary>
        /// <param name="targetXsnLocation">The target XSN location.</param>
        /// <param name="isCached">if set to <c>true</c> [is cached].</param>
        /// <param name="openByDefault">if set to <c>true</c> [open by default].</param>
        /// <param name="scope">The scope.</param>
        /// <returns></returns>
        private XmlDocument BuildCustomInformationPanelXml(string targetXsnLocation, bool isCached, bool openByDefault, string scope)
        {
            XmlDocument document = new XmlDocument();
            XmlNode newChild = document.CreateNode(XmlNodeType.Element, "customXsn", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            document.AppendChild(newChild);
            XmlNode node2 = document.CreateNode(XmlNodeType.Element, "xsnLocation", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            newChild.AppendChild(node2);
            XmlNode node3 = document.CreateNode(XmlNodeType.Text, "xsnLocationText", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            node3.Value = targetXsnLocation;
            node2.AppendChild(node3);
            node2 = document.CreateNode(XmlNodeType.Element, "cached", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            newChild.AppendChild(node2);
            node3 = document.CreateNode(XmlNodeType.Text, "cachedText", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            node3.Value = isCached.ToString();
            node2.AppendChild(node3);
            node2 = document.CreateNode(XmlNodeType.Element, "openByDefault", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            newChild.AppendChild(node2);
            node3 = document.CreateNode(XmlNodeType.Text, "openByDefaultText", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            node3.Value = openByDefault.ToString();
            node2.AppendChild(node3);
            node2 = document.CreateNode(XmlNodeType.Element, "xsnScope", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            newChild.AppendChild(node2);
            node3 = document.CreateNode(XmlNodeType.Text, "xsnScopeText", "http://schemas.microsoft.com/office/2006/metadata/customXsn");
            node3.Value = scope;
            node2.AppendChild(node3);
            return document;
        }

    }
}
