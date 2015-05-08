using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;

namespace Lapointe.SharePoint.PowerShell.Common.ManagedMetadata
{
    public class ImportTerms
    {
        private XmlDocument _xml;
        private bool _whatIf;

        public ImportTerms(XmlDocument xml, bool whatIf)
        {
            _xml = xml;
            _whatIf = whatIf;
            if (xml.DocumentElement == null)
                throw new Exception("The XML provided does not include a root element.");
        }

        public void Import(TaxonomySession ts)
        {
            if (ts == null)
                throw new ArgumentNullException("ts", "The TaxonomySession object is null.");

            XmlNodeList termStoreNodes = _xml.SelectNodes("//TermStore");
            if (termStoreNodes == null || termStoreNodes.Count == 0)
                return;

            foreach (XmlElement termStoreElement in termStoreNodes)
            {
                string termStoreName = termStoreElement.GetAttribute("Name");
                Logger.Write("Importing Term Store: {0}", termStoreName);

                TermStore termStore = ts.TermStores[termStoreName];
                if (termStore == null)
                {
                    Logger.WriteWarning("Unable to locate target Term Store: {0}", termStoreName);
                    continue;
                }

                XmlNodeList groupNodes = termStoreElement.SelectNodes("./Groups/Group");
                if (groupNodes == null || groupNodes.Count == 0)
                {
                    Logger.WriteWarning("No Group elements were defined in the import file for the Term Store.");
                    continue;
                }
                foreach (XmlElement groupElement in groupNodes)
                {
                    Import(groupElement, termStore);
                }
                if (!_whatIf)
                    termStore.CommitAll();
            }
        }

        public void Import(TermStore parentTermStore)
        {
            if (parentTermStore == null)
                throw new ArgumentNullException("parentTermStore", "The parent TermStore object is null.");
            
            if (_xml.DocumentElement.Name == "Groups")
            {
                XmlNodeList groupNodes = _xml.SelectNodes("./Groups/Group");
                if (groupNodes == null || groupNodes.Count == 0)
                    return;

                foreach (XmlElement groupElement in groupNodes)
                {
                    Import(groupElement, parentTermStore);
                }
            }
            else if (_xml.DocumentElement.Name == "Group")
            {
                Import(_xml.DocumentElement, parentTermStore);
            }
            parentTermStore.CommitAll();
        }

        public void Import(Group parentGroup)
        {
            if (parentGroup == null)
                throw new ArgumentNullException("parentGroup", "The parent Group object is null.");

            if (_xml.DocumentElement.Name == "TermSets")
            {
                XmlNodeList termSetNodes = _xml.SelectNodes("./TermSets/TermSet");
                if (termSetNodes == null || termSetNodes.Count == 0)
                {
                    Logger.WriteWarning("No Term Set elements were defined in the import file for the Group.");
                    return;
                }

                foreach (XmlElement termSetElement in termSetNodes)
                {
                    Import(termSetElement, parentGroup);
                }
            }
            else if (_xml.DocumentElement.Name == "TermSet")
            {
                Import(_xml.DocumentElement, parentGroup);
            }
            parentGroup.TermStore.CommitAll();
        }

        public void Import(TermSet parentTermSet)
        {
            if (parentTermSet == null)
                throw new ArgumentNullException("parentTermSet", "The parent TermSet object is null.");

            XmlNodeList termNodes;
            if (_xml.DocumentElement.Name == "Terms")
            {
                termNodes = _xml.SelectNodes("./Terms/Term");
                if (termNodes == null || termNodes.Count == 0)
                {
                    Logger.WriteWarning("No Term elements were defined in the import file for the Term Set.");
                    return;
                }
                foreach (XmlElement termElement in termNodes)
                {
                    Import(termElement, parentTermSet);
                }
            }
            else if (_xml.DocumentElement.Name == "Term")
            {
                Import(_xml.DocumentElement, parentTermSet);
            }
            parentTermSet.TermStore.CommitAll();
        }


        public void Import(Term parentTerm)
        {
            if (parentTerm == null)
                throw new ArgumentNullException("parentTerm", "The parent Term object is null.");

            XmlNodeList termNodes;
            if (_xml.DocumentElement.Name == "Terms")
            {
                termNodes = _xml.SelectNodes("./Terms/Term");
                if (termNodes == null || termNodes.Count == 0)
                {
                    Logger.WriteWarning("No Term elements were defined in the import file for the Term.");
                    return;
                }
                foreach (XmlElement termElement in termNodes)
                {
                    Import(termElement, parentTerm);
                }
            }
            else if (_xml.DocumentElement.Name == "Term")
            {
                Import(_xml.DocumentElement, parentTerm);
            }
            parentTerm.TermStore.CommitAll();
        }

        private void Import(XmlElement groupElement, TermStore parentTermStore)
        {
            string groupName = groupElement.GetAttribute("Name");
            Group group = null;
            if (bool.Parse(groupElement.GetAttribute("IsSiteCollectionGroup")))
            {
                XmlNodeList siteCollectionIdNodes = groupElement.SelectNodes("./SiteCollectionAccessIds/SiteCollectionAccessId");
                if (siteCollectionIdNodes != null)
                {
                    foreach (XmlElement siteCollectionIdElement in siteCollectionIdNodes)
                    {
                        SPSite site = null;
                        if (!string.IsNullOrEmpty(siteCollectionIdElement.GetAttribute("Url")))
                        {
                            try
                            {
                                site = new SPSite(siteCollectionIdElement.GetAttribute("Url"));
                            }
                            catch
                            {
                                Logger.WriteWarning("Unable to locate a Site Collection at {0}", siteCollectionIdElement.GetAttribute("Url"));
                            }
                        }
                        else
                        {
                            try
                            {
                                site = new SPSite(new Guid(siteCollectionIdElement.GetAttribute("Id")));
                            }
                            catch
                            {
                                Logger.WriteWarning("Unable to locate a Site Collection with ID {0}", siteCollectionIdElement.GetAttribute("Id"));
                            }
                        }
                        if (site != null)
                        {
                            try
                            {
                                if (group == null)
                                {
                                    group = parentTermStore.GetSiteCollectionGroup(site);
                                }
                                if (group != null && group.IsSiteCollectionGroup)
                                {
                                    group.AddSiteCollectionAccess(site.ID);
                                }
                            }
                            catch (MissingMethodException)
                            {
                                Logger.WriteWarning("Unable to retrieve or add Site Collection group. SharePoint 2010 Service Pack 1 or greater is required. ID={0}, Url={1}", siteCollectionIdElement.GetAttribute("Id"), siteCollectionIdElement.GetAttribute("Url"));
                            }
                            finally
                            {
                                site.Dispose();
                            }
                        }
                    }
                }
            }
            try
            {
                if (group == null)
                    group = parentTermStore.Groups[groupName];
            }
            catch (ArgumentException) {}

            if (group == null)
            {
                Logger.Write("Creating Group: {0}", groupName);
#if SP2010
                group = parentTermStore.CreateGroup(groupName);
#else
                // Updated provided by John Calvert
                if (!string.IsNullOrEmpty(groupElement.GetAttribute("Id")))
                {
                    Guid id = new Guid(groupElement.GetAttribute("Id"));
                    group = parentTermStore.CreateGroup(groupName, id);
                }
                else
                    group = parentTermStore.CreateGroup(groupName);
                // End update
#endif
                group.Description = groupElement.GetAttribute("Description");
            }
            parentTermStore.CommitAll();//TEST

            XmlNodeList termSetNodes = groupElement.SelectNodes("./TermSets/TermSet");
            if (termSetNodes != null && termSetNodes.Count > 0)
            {
                foreach (XmlElement termSetElement in termSetNodes)
                {
                    Import(termSetElement, group);
                }
            }
        }

        private void Import(XmlElement termSetElement, Group parentGroup)
        {
            string termSetName = termSetElement.GetAttribute("Name");
            TermSet termSet = null;
            try
            {
                termSet = parentGroup.TermSets[termSetName];
            }
            catch (ArgumentException) {}

            if (termSet == null)
            {
                Logger.Write("Creating Term Set: {0}", termSetName);

                int lcid = parentGroup.TermStore.WorkingLanguage;
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("Id")))
                {
                    Guid id = new Guid(termSetElement.GetAttribute("Id"));
                    termSet = parentGroup.CreateTermSet(termSetName, id, lcid);
                }
                else
                    termSet = parentGroup.CreateTermSet(termSetName, lcid);

                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("Contact")))
                    termSet.Contact = termSetElement.GetAttribute("Contact");
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("Description")))
                    termSet.Description = termSetElement.GetAttribute("Description");
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("CustomSortOrder")))
                    termSet.CustomSortOrder = termSetElement.GetAttribute("CustomSortOrder");
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("IsAvailableForTagging")))
                    termSet.IsAvailableForTagging = bool.Parse(termSetElement.GetAttribute("IsAvailableForTagging"));
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("Owner")))
                    termSet.Owner = termSetElement.GetAttribute("Owner");

                termSet.IsOpenForTermCreation = true;
            }

#if SP2013
            // Updated provided by John Calvert
            XmlNodeList propertyNodes = termSetElement.SelectNodes("./CustomProperties/CustomProperty");
            if (propertyNodes != null && propertyNodes.Count > 0)
            {
                foreach (XmlElement propertyElement in propertyNodes)
                {
                    termSet.SetCustomProperty(propertyElement.GetAttribute("Name"),
                        propertyElement.GetAttribute("Value"));
                }
            }
            // End update
#endif

            parentGroup.TermStore.CommitAll();//TEST

            XmlNodeList termsNodes = termSetElement.SelectNodes("./Terms/Term");
            if (termsNodes != null && termsNodes.Count > 0)
            {
                foreach (XmlElement termElement in termsNodes)
                {
                    Import(termElement, termSet);
                }
            }
            if (!string.IsNullOrEmpty(termSetElement.GetAttribute("IsOpenForTermCreation")) && termSet.TermStore.OrphanedTermsTermSet != termSet)
                termSet.IsOpenForTermCreation = bool.Parse(termSetElement.GetAttribute("IsOpenForTermCreation"));
        }

        private void Import(XmlElement termElement, TermSetItem parentTermSetItem)
        {
            string termName = termElement.GetAttribute("Name");
            Term term = null;
            try
            {
                term = parentTermSetItem.Terms[termName];
            }
            catch (ArgumentException) {}

            if (term == null)
            {
                if (!string.IsNullOrEmpty(termElement.GetAttribute("IsSourceTerm")) &&
                    !bool.Parse(termElement.GetAttribute("IsSourceTerm")))
                {
                    string[] sourceTermInfo = termElement.GetAttribute("SourceTerm").Split('|');

                    Term sourceTerm = parentTermSetItem.TermStore.GetTerm(new Guid(sourceTermInfo[0]));
                    if (sourceTerm == null)
                    {
                        TermCollection sourceTerms = parentTermSetItem.TermStore.GetTerms(sourceTermInfo[1], true,
                                                                                          StringMatchOption.ExactMatch,
                                                                                          1, false);
                        if (sourceTerms != null && sourceTerms.Count > 0)
                            sourceTerm = sourceTerms[0];
                    }
                    if (sourceTerm != null)
                    {
                        Logger.Write("Creating Reference Term: {0}", termName); 
                        term = parentTermSetItem.ReuseTerm(sourceTerm, false);
                    }
                    else
                        Logger.WriteWarning("The Source Term, {0}, was not found. {1} will be created without linking.", sourceTermInfo[1], termName);
                }
                if (term == null)
                {
                    Logger.Write("Creating Term: {0}", termName);
                    
                    int lcid = parentTermSetItem.TermStore.WorkingLanguage;

                    if (!string.IsNullOrEmpty(termElement.GetAttribute("Id")))
                    {
                        Guid id = new Guid(termElement.GetAttribute("Id"));
                        term = parentTermSetItem.CreateTerm(termName, lcid, id);
                    }
                    else
                        term = parentTermSetItem.CreateTerm(termName, lcid);

                    if (!string.IsNullOrEmpty(termElement.GetAttribute("CustomSortOrder")))
                        term.CustomSortOrder = termElement.GetAttribute("CustomSortOrder");
                    if (!string.IsNullOrEmpty(termElement.GetAttribute("IsAvailableForTagging")))
                        term.IsAvailableForTagging = bool.Parse(termElement.GetAttribute("IsAvailableForTagging"));
                    if (!string.IsNullOrEmpty(termElement.GetAttribute("Owner")))
                        term.Owner = termElement.GetAttribute("Owner");

                    if (!string.IsNullOrEmpty(termElement.GetAttribute("IsDeprecated")) &&
                        bool.Parse(termElement.GetAttribute("IsDeprecated")))
                        term.Deprecate(true);
                }
            }

            XmlNodeList descriptionNodes = termElement.SelectNodes("./Descriptions/Description");
            if (descriptionNodes != null && descriptionNodes.Count > 0)
            {
                foreach (XmlElement descriptionElement in descriptionNodes)
                {
                    term.SetDescription(descriptionElement.GetAttribute("Value"),
                        int.Parse(descriptionElement.GetAttribute("Language")));
                }
            }

            XmlNodeList propertyNodes = termElement.SelectNodes("./CustomProperties/CustomProperty");
            if (propertyNodes != null && propertyNodes.Count > 0)
            {
                foreach (XmlElement propertyElement in propertyNodes)
                {
                    term.SetCustomProperty(propertyElement.GetAttribute("Name"), 
                        propertyElement.GetAttribute("Value"));
                }
            }

#if SP2013
            // Updated provided by John Calvert
            XmlNodeList localpropertyNodes = termElement.SelectNodes("./LocalCustomProperties/LocalCustomProperty");
            if (localpropertyNodes != null && localpropertyNodes.Count > 0)
            {
                foreach (XmlElement localpropertyElement in localpropertyNodes)
                {
                    term.SetLocalCustomProperty(localpropertyElement.GetAttribute("Name"),
                        localpropertyElement.GetAttribute("Value"));
                }
            }
            // End update
#endif

            XmlNodeList labelNodes = termElement.SelectNodes("./Labels/Label");
            if (labelNodes != null && labelNodes.Count > 0)
            {
                foreach (XmlElement labelElement in labelNodes)
                {
                    string labelValue = labelElement.GetAttribute("Value");
                    int lcid = int.Parse(labelElement.GetAttribute("Language"));
                    bool isDefault = bool.Parse(labelElement.GetAttribute("IsDefaultForLanguage"));
                    Label label = term.GetAllLabels(lcid).FirstOrDefault(currentLabel => currentLabel.Value == labelValue);
                    if (label == null)
                    {
                        term.CreateLabel(labelValue, lcid, isDefault);
                    }
                    else
                    {
                        if (isDefault && !label.IsDefaultForLanguage)
                            label.SetAsDefaultForLanguage();
                    }
                }
            }
            parentTermSetItem.TermStore.CommitAll();//TEST

            XmlNodeList termsNodes = termElement.SelectNodes("./Terms/Term");
            if (termsNodes != null && termsNodes.Count > 0)
            {
                foreach (XmlElement childTermElement in termsNodes)
                {
                    Import(childTermElement, term);
                }
            }
        }

    }
}
