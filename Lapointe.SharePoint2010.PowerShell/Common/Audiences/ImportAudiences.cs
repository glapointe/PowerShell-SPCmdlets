using System;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Common.Audiences
{
    internal class ImportAudiences
    {
        /// <summary>
        /// Imports the specified XML.
        /// </summary>
        /// <param name="xml">The XML.</param>
        /// <param name="context">The context.</param>
        /// <param name="deleteExisting">if set to <c>true</c> [delete existing].</param>
        /// <param name="compile">if set to <c>true</c> [compile].</param>
        /// <param name="mapFile">The map file.</param>
        internal static void Import(string xml, SPServiceContext context, bool deleteExisting, bool compile, string mapFile)
        {
            AudienceManager manager = new AudienceManager(context);

            XmlDocument audiencesDoc = new XmlDocument();
            audiencesDoc.LoadXml(xml);

            XmlNodeList audienceElements = audiencesDoc.SelectNodes("//Audience");
            if (audienceElements == null)
                throw new ArgumentException("The input file does not contain any audience elements.");

            StringBuilder sb = new StringBuilder();
            XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.WriteStartElement("Replacements");

            Dictionary<Guid, string> ids = new Dictionary<Guid, string>();
            if (deleteExisting)
            {
                Logger.Write("Progrss: Deleting existing audiences.");
                foreach (Audience au in manager.Audiences)
                    ids.Add(au.AudienceID, au.AudienceName);

                foreach (Guid id in ids.Keys)
                {
                    if (id == Guid.Empty)
                        continue;

                    string name = manager.Audiences[id].AudienceName;
                    manager.Audiences.Remove(id);
                }
            }

            foreach (XmlElement audienceElement in audienceElements)
            {
                string audienceName = audienceElement.GetAttribute("AudienceName");
                string audienceDesc = audienceElement.GetAttribute("AudienceDescription");

                Audience audience;
                bool updatedAudience = false;
                if (manager.Audiences.AudienceExist(audienceName))
                {
                    Logger.Write("Progress: Updating audience {0}.", audienceName);
                    audience = manager.Audiences[audienceName];
                    audience.AudienceDescription = audienceDesc ?? "";
                    updatedAudience = true;
                }
                else
                {
                    // IMPORTANT: the create method does not do a null check but the methods that load the resultant collection assume not null.
                    Logger.Write("Progress: Creating audience {0}.", audienceName);
                    audience = manager.Audiences.Create(audienceName, audienceDesc ?? "");
                }

                audience.GroupOperation = (AudienceGroupOperation)Enum.Parse(typeof (AudienceGroupOperation),
                                                     audienceElement.GetAttribute("GroupOperation"));

                audience.OwnerAccountName = audienceElement.GetAttribute("OwnerAccountName");

                audience.Commit();

                if (updatedAudience && audience.AudienceID != Guid.Empty)
                {
                    // We've updated an existing audience.
                    xmlWriter.WriteStartElement("Replacement");
                    xmlWriter.WriteElementString("SearchString", string.Format("(?i:{0})", (new Guid(audienceElement.GetAttribute("AudienceID")).ToString().ToUpper())));
                    xmlWriter.WriteElementString("ReplaceString", audience.AudienceID.ToString().ToUpper());
                    xmlWriter.WriteEndElement(); // Replacement
                }
                else if (!updatedAudience && audience.AudienceID != Guid.Empty && ids.ContainsValue(audience.AudienceName))
                {
                    // We've added a new audience which we just previously deleted.
                    xmlWriter.WriteStartElement("Replacement");
                    foreach (Guid id in ids.Keys)
                    {
                        if (ids[id] == audience.AudienceName)
                        {
                            xmlWriter.WriteElementString("SearchString", string.Format("(?i:{0})", id.ToString().ToUpper()));
                            break;
                        }
                    }
                    xmlWriter.WriteElementString("ReplaceString", audience.AudienceID.ToString().ToUpper());
                    xmlWriter.WriteEndElement(); // Replacement
                }

                XmlElement rulesElement = (XmlElement)audienceElement.SelectSingleNode("rules");
                if (rulesElement == null || rulesElement.ChildNodes.Count == 0)
                {
                    audience.AudienceRules = new ArrayList();
                    audience.Commit();
                    continue;
                }


                string rules = rulesElement.OuterXml;
                Logger.Write("Progress: Adding rules to audience {0}.", audienceName);
                AddAudienceRule.AddRules(context, audienceName, rules, true, compile, false, AppendOp.AND);
            }

            xmlWriter.WriteEndElement(); // Replacements

            if (!string.IsNullOrEmpty(mapFile))
            {
                xmlWriter.Flush();
                File.WriteAllText(mapFile, sb.ToString());
            }
        }

    }
}
