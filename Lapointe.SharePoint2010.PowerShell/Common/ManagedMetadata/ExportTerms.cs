using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;

namespace Lapointe.SharePoint.PowerShell.Common.ManagedMetadata
{
    public class ExportTerms
    {
        private XmlDocument _xml = new XmlDocument();

        
        private XmlElement AddTermStoreElement(XmlElement parent, TermStore termStore)
        {
            XmlElement element = _xml.CreateElement("TermStore");
            if (parent == null)
                _xml.AppendChild(element);
            else
                parent.AppendChild(element);

            element.SetAttribute("Name", termStore.Name);
            element.SetAttribute("Id", termStore.Id.ToString());
            element.SetAttribute("IsOnline", termStore.IsOnline.ToString());
            element.SetAttribute("WorkingLanguage", termStore.WorkingLanguage.ToString());
            element.SetAttribute("DefaultLanguage", termStore.DefaultLanguage.ToString());
            element.SetAttribute("SystemGroup", termStore.SystemGroup.Id.ToString());

            return element;
        }

        private XmlElement GetTermStoreElement(XmlElement parent, TermStore termStore)
        {
            if (parent == null)
                return AddTermStoreElement(parent, termStore);

            XmlElement termStoreElement = parent.SelectSingleNode("./TermStore[@Id='" + termStore.Id + "']") as XmlElement;
            if (termStoreElement == null)
                termStoreElement = AddTermStoreElement(parent, termStore);
            return termStoreElement;
        }

        private XmlElement GetGroupElement(XmlElement parent, Group group)
        {
            if (parent == null)
                return AddGroupElement(parent, group);

            XmlElement groupElement = parent.SelectSingleNode("./Groups/Group[@Id='" + group.Id + "']") as XmlElement;
            if (groupElement == null)
                groupElement = AddGroupElement(parent, group);

            return groupElement;
        }

        private XmlElement AddGroupElement(XmlElement parent, Group group)
        {
            XmlElement element = _xml.CreateElement("Group");

            if (parent == null)
                _xml.AppendChild(element);
            else
            {
                if (parent.Name == "Groups")
                    parent.AppendChild(element);
                else
                {
                    XmlElement groupsElement = parent.SelectSingleNode("./Groups") as XmlElement;
                    if (groupsElement == null)
                    {
                        groupsElement = _xml.CreateElement("Groups");
                        parent.AppendChild(groupsElement);
                    }
                    groupsElement.AppendChild(element);
                }
            }

            element.SetAttribute("Id", group.Id.ToString());
            element.SetAttribute("Name", group.Name);
            element.SetAttribute("Description", group.Description);
            element.SetAttribute("CreatedDate", group.CreatedDate.ToString());
            element.SetAttribute("LastModifiedDate", group.LastModifiedDate.ToString());
            element.SetAttribute("IsSystemGroup", group.IsSystemGroup.ToString());
            element.SetAttribute("IsSiteCollectionGroup", group.IsSiteCollectionGroup.ToString());
            if (group.SiteCollectionAccessIds.Count > 0)
            {
                XmlElement siteColIdsElement = _xml.CreateElement("SiteCollectionAccessIds");
                element.AppendChild(siteColIdsElement);
                foreach (Guid id in group.SiteCollectionAccessIds)
                {
                    XmlElement idElement = _xml.CreateElement("SiteCollectionAccessId");
                    idElement.SetAttribute("Id", id.ToString());
                    try
                    {
                        using (SPSite site = new SPSite(id))
                            idElement.SetAttribute("Url", site.Url);
                    }
                    catch {}
                    siteColIdsElement.AppendChild(idElement);
                }
            }
            return element;
        }


        private XmlElement GetTermSetElement(XmlElement parent, TermSet termSet)
        {
            if (parent == null)
                return AddTermSetElement(parent, termSet);

            XmlElement termSetElement = parent.SelectSingleNode("./TermSets/TermSet[@Id='" + termSet.Id + "']") as XmlElement;
            if (termSetElement == null)
                termSetElement = AddTermSetElement(parent, termSet);

            return termSetElement;
        }


        private XmlElement AddTermSetElement(XmlElement parent, TermSet termSet)
        {
            XmlElement element = _xml.CreateElement("TermSet");
            if (parent == null)
                _xml.AppendChild(element);
            else
            {
                if (parent.Name == "TermSets")
                    parent.AppendChild(element);
                else
                {
                    XmlElement termSetsElement = parent.SelectSingleNode("./TermSets") as XmlElement;
                    if (termSetsElement == null)
                    {
                        termSetsElement = _xml.CreateElement("TermSets");
                        parent.AppendChild(termSetsElement);
                    }
                    termSetsElement.AppendChild(element);
                }
            }

            element.SetAttribute("Id", termSet.Id.ToString());
            element.SetAttribute("Name", termSet.Name);
            element.SetAttribute("Description", termSet.Description);
            element.SetAttribute("CreatedDate", termSet.CreatedDate.ToString());
            element.SetAttribute("LastModifiedDate", termSet.LastModifiedDate.ToString());
            element.SetAttribute("Contact", termSet.Contact);
            element.SetAttribute("Owner", termSet.Owner);
            element.SetAttribute("IsAvailableForTagging", termSet.IsAvailableForTagging.ToString());
            element.SetAttribute("IsOpenForTermCreation", termSet.IsOpenForTermCreation.ToString());
            element.SetAttribute("CustomSortOrder", termSet.CustomSortOrder);

#if SP2013
            // Updated provided by John Calvert
            XmlElement propertiesElement = _xml.CreateElement("CustomProperties");
            element.AppendChild(propertiesElement);
            foreach (string key in termSet.CustomProperties.Keys)
            {
                XmlElement propertyElement = _xml.CreateElement("CustomProperty");
                propertiesElement.AppendChild(propertyElement);
                propertyElement.SetAttribute("Name", key);
                propertyElement.SetAttribute("Value", termSet.CustomProperties[key]);
            }
            // End update
#endif
            return element;
        }



        private XmlElement GetTermElement(XmlElement parent, Term term)
        {
            if (parent == null)
                return AddTermElement(parent, term);

            XmlElement termElement = parent.SelectSingleNode("./Terms/Term[@Id='" + term.Id + "']") as XmlElement;
            if (termElement == null)
                termElement = AddTermElement(parent, term);

            return termElement;
        }


        private XmlElement AddTermElement(XmlElement parent, Term term)
        {
            XmlElement element = _xml.CreateElement("Term");
            if (parent == null)
                _xml.AppendChild(element);
            else
            {
                if (parent.Name == "TermSets")
                    parent.AppendChild(element);
                else
                {
                    XmlElement termsElement = parent.SelectSingleNode("./Terms") as XmlElement;
                    if (termsElement == null)
                    {
                        termsElement = _xml.CreateElement("Terms");
                        parent.AppendChild(termsElement);
                    }
                    termsElement.AppendChild(element);
                }
            }
            
            element.SetAttribute("Id", term.Id.ToString());
            element.SetAttribute("Name", term.Name);
            element.SetAttribute("CreatedDate", term.CreatedDate.ToString());
            element.SetAttribute("LastModifiedDate", term.LastModifiedDate.ToString());
            element.SetAttribute("Owner", term.Owner);
            element.SetAttribute("IsDeprecated", term.IsDeprecated.ToString());
            element.SetAttribute("IsAvailableForTagging", term.IsAvailableForTagging.ToString());
            element.SetAttribute("IsKeyword", term.IsKeyword.ToString());
            element.SetAttribute("IsReused", term.IsReused.ToString());
            element.SetAttribute("IsRoot", term.IsRoot.ToString());
            element.SetAttribute("IsSourceTerm", term.IsSourceTerm.ToString());
            element.SetAttribute("CustomSortOrder", term.CustomSortOrder);
            element.SetAttribute("SourceTerm", term.SourceTerm.Id + "|" + term.SourceTerm.Name);

            XmlElement descriptionsElement = _xml.CreateElement("Descriptions");
            element.AppendChild(descriptionsElement);
            foreach (int lcid in term.TermStore.Languages)
            {
                XmlElement descriptionElement = _xml.CreateElement("Description");
                descriptionsElement.AppendChild(descriptionElement);
                descriptionElement.SetAttribute("Language", lcid.ToString());
                descriptionElement.SetAttribute("Value", term.GetDescription(lcid));
            }

            XmlElement propertiesElement = _xml.CreateElement("CustomProperties");
            element.AppendChild(propertiesElement);
            foreach (string key in term.CustomProperties.Keys)
            {
                XmlElement propertyElement = _xml.CreateElement("CustomProperty");
                propertiesElement.AppendChild(propertyElement);
                propertyElement.SetAttribute("Name", key);
                propertyElement.SetAttribute("Value", term.CustomProperties[key]);
            }

#if SP2013
            // Updated provided by John Calvert
            XmlElement localpropertiesElement = _xml.CreateElement("LocalCustomProperties");
            element.AppendChild(localpropertiesElement);
            foreach (string key in term.LocalCustomProperties.Keys)
            {
                XmlElement localpropertyElement = _xml.CreateElement("LocalCustomProperty");
                localpropertiesElement.AppendChild(localpropertyElement);
                localpropertyElement.SetAttribute("Name", key);
                localpropertyElement.SetAttribute("Value", term.LocalCustomProperties[key]);
            }
            // End update
#endif

            return element;
        }


        private XmlElement AddLabelElement(XmlElement termElement, Label label)
        {
            XmlElement labelsElement = termElement.SelectSingleNode("./Labels") as XmlElement;
            if (labelsElement == null)
            {
                labelsElement = _xml.CreateElement("Labels");
                termElement.AppendChild(labelsElement);
            }

            XmlElement element = _xml.CreateElement("Label");
            labelsElement.AppendChild(element);
            element.SetAttribute("Value", label.Value);
            element.SetAttribute("Language", label.Language.ToString());
            element.SetAttribute("IsDefaultForLanguage", label.IsDefaultForLanguage.ToString());

            return element;
        }

        public XmlDocument Export(TaxonomySession ts)
        {
            if (ts == null)
                throw new ArgumentNullException("ts", "The TaxonomySession object is null.");

            XmlElement element = _xml.CreateElement("TermStores");
            _xml.AppendChild(element);

            foreach (TermStore termStore in ts.TermStores)
            {
                Export(element, termStore);
            }
            return _xml;
        }

        public XmlDocument Export(TermStore termStore)
        {
            if (termStore == null)
                throw new ArgumentNullException("termStore", "The TermStore object is null.");

            XmlElement termStoreElement = AddTermStoreElement(null, termStore);
            foreach (Group group in termStore.Groups)
            {
                Export(termStoreElement, group);
            }
            return _xml;
        }

        
        public XmlDocument Export(Group group)
        {
            if (group == null)
                throw new ArgumentNullException("group", "The Group object is null.");

            XmlElement groupElement = AddGroupElement(null, group);
            foreach (TermSet termSet in group.TermSets)
            {
                Export(groupElement, termSet);
            }
            return _xml;
        }

        public XmlDocument Export(TermSet termSet)
        {
            if (termSet == null)
                throw new ArgumentNullException("termSet", "The TermSet object is null.");

            XmlElement termSetElement = AddTermSetElement(null, termSet);
            foreach (Term term in termSet.Terms)
            {
                Export(termSetElement, term);
            }
            return _xml;
        }

        public XmlDocument Export(Term term)
        {
            if (term == null)
                throw new ArgumentNullException("term", "The Term object is null.");

            XmlElement termElement = AddTermElement(null, term);
            foreach (Label label in term.Labels)
            {
                AddLabelElement(termElement, label);
            }

            foreach (Term childTerm in term.Terms)
            {
                Export(termElement, childTerm);
            }
            return _xml;
        }

        private void Export(XmlElement parentElement, TermStore termStore)
        {
            if (termStore == null)
                throw new ArgumentNullException("termStore", "The TermStore object is null.");

            XmlElement termStoreElement = GetTermStoreElement(parentElement, termStore);
            foreach (Group termSet in termStore.Groups)
            {
                Export(termStoreElement, termSet);
            }

        }
        
        private void Export(XmlElement parentElement, Group group)
        {
            XmlElement groupElement = GetGroupElement(parentElement, group);
            foreach (TermSet termSet in group.TermSets)
            {
                Export(groupElement, termSet);
            }

        }

        private void Export(XmlElement parentElement, TermSet termSet)
        {
            XmlElement termSetElement = GetTermSetElement(parentElement, termSet);
            foreach (Term term in termSet.Terms)
            {
                Export(termSetElement, term);
            }
        }

        private void Export(XmlElement parentElement, Term term)
        {
            XmlElement termElement = GetTermElement(parentElement, term);
            foreach (Label label in term.Labels)
            {
                AddLabelElement(termElement, label);
            }

            foreach (Term childTerm in term.Terms)
            {
                Export(termElement, childTerm);
            }
        }

        
    }
}
