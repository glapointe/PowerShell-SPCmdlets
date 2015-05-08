using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.Services.Protocols;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.SharePoint.PowerShell.Common.Lists;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.WebPartPages;
using Storage = Lapointe.SharePoint.PowerShell.WebPartPagesWebService.Storage;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet("Replace", "SPWebPartType", SupportsShouldProcess = true),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Replaces instances of one web part type with another web part type.")]
    [RelatedCmdlets(typeof(SPCmdletGetFile))]
    [Example(Code = "PS C:\\> Replace-SPWebPartType -File \"http://server_name/pages/default.aspx\" -OldType \"Microsoft.SharePoint.Publishing.WebControls.ContentByQueryWebPart, Microsoft.SharePoint.Publishing, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" -NewType \"MyContentByQueryWebPart, MyCompany.SharePoint.WebParts, Version=1.0.0.0, Culture=neutral, PublicKeyToken=4ec4b9177b831752\" -Publish",
        Remarks = "This example replaces all instances of the web part who's class name is ContentByQueryWebPart with the web part who's class name is MyContentByQueryWebPart.")]
    public class SPCmdletReplaceWebPartType : SPCmdletCustom
    {
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Alias(new string[] { "Url", "Page" })]
        public SPFilePipeBind File { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The web part type to replace.")]
        public TypePipeBind OldType { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The web part type to replace the old type with.")]
        public TypePipeBind NewType { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The web part title to restrict the replacement to.")]
        public string Title { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Additional properties to set or override after copying the old web part properties.")]
        public Hashtable Properties { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "If specified the page will be published after adjusting the Web Part.")]
        public SwitchParameter Publish { get; set; }

        protected override void InternalProcessRecord()
        {
            bool test = false;
            ShouldProcessReason reason;
            if (!base.ShouldProcess(null, null, null, out reason))
            {
                if (reason == ShouldProcessReason.WhatIf)
                {
                    test = true;
                }
            }

            string fileUrl = File.FileUrl;

            SetWebPart(fileUrl, OldType.Read(), NewType.Read(), Title, Properties, Publish, test);

            base.InternalProcessRecord();
        }

        internal static void SetWebPart(string url, Type oldType, Type newType, string title, Hashtable properties, bool publish, bool test)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb()) // The url contains a filename so AllWebs[] will not work unless we want to try and parse which we don't
            {
                bool cleanupContext = false;
                try
                {
                    if (HttpContext.Current == null)
                    {
                        cleanupContext = true;
                        HttpRequest httpRequest = new HttpRequest("", web.Url, "");
                        HttpContext.Current = new HttpContext(httpRequest, new HttpResponse(new StringWriter()));
                        SPControl.SetContextWeb(HttpContext.Current, web);
                    }


                    SPFile file = web.GetFile(url);

                    // file.Item will throw "The object specified does not belong to a list." if the url passed
                    // does not correspond to a file in a list.

                    bool checkIn = false;
                    if (file.InDocumentLibrary && !test)
                    {
                        if (!Utilities.IsCheckedOut(file.Item) || !Utilities.IsCheckedOutByCurrentUser(file.Item))
                        {
                            file.CheckOut();
                            checkIn = true;
                            // If it's checked out by another user then this will throw an informative exception so let it do so.
                        }
                    }

                    SPLimitedWebPartManager manager = null;
                    try
                    {
                        List<WebPart> webParts = Utilities.GetWebPartsByType(web, url, oldType, out manager);
                        foreach (var oldWebPart in webParts)
                        {
                            if (oldWebPart.IsClosed) continue;

                            string wpTitle = oldWebPart.Title;
                            if (string.IsNullOrEmpty(wpTitle)) wpTitle = oldWebPart.DisplayTitle;

                            if (!string.IsNullOrEmpty(title) && 
                                (oldWebPart.DisplayTitle.ToLowerInvariant() != title.ToLowerInvariant() &&
                                oldWebPart.Title.ToLowerInvariant() != title.ToLowerInvariant()))
                            {
                                continue;
                            }
                            Logger.Write("Replacing web part \"{0}\"...", wpTitle);
                            string zone = manager.GetZoneID(oldWebPart);
                            WebPart newWebPart = (WebPart)Activator.CreateInstance(newType);
                            if (SetProperties(oldWebPart, newWebPart, properties))
                            {
                                Logger.WriteWarning("An error was encountered setting web part properties so try one more time in case the error is the result of a sequencing issue.");
                                if (SetProperties(oldWebPart, newWebPart, properties))
                                {
                                    Logger.WriteWarning("Unable to set all properties for web part.");
                                }
                            }
                            if (!test)
                                manager.DeleteWebPart(oldWebPart);

                            try
                            {
                                if (!test)
                                {
                                    manager.AddWebPart(newWebPart, zone, oldWebPart.ZoneIndex);
                                }
                            }
                            catch (Exception)
                            {
                                ServicePointManager.ServerCertificateValidationCallback += delegate { return true; };

                                // We've not already added the web part so use the web service to do this.
                                using (WebPartPagesWebService.WebPartPagesWebService svc = new WebPartPagesWebService.WebPartPagesWebService())
                                {
                                    // We failed adding via the OM so try the web service as a fall back.
                                    svc.Url = manager.Web.Url + "/_vti_bin/WebPartPages.asmx";
                                    svc.Credentials = CredentialCache.DefaultCredentials;

                                    try
                                    {
                                        // Add the web part to the web service.  We use a web service because many
                                        // web parts require the SPContext.Current variables to be set which are
                                        // not set when run from a command line.
                                        StringBuilder sb = new StringBuilder();
                                        XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
                                        xmlWriter.Formatting = Formatting.Indented;
                                        manager.ExportWebPart(newWebPart, xmlWriter);
                                        xmlWriter.Flush();

                                        svc.AddWebPartToZone(url, sb.ToString(), Storage.Shared, zone, oldWebPart.ZoneIndex);

                                    }
                                    catch (SoapException ex)
                                    {
                                        throw new Exception(ex.Detail.OuterXml);
                                    }
                                }
                            }
                            finally
                            {
                                oldWebPart.Dispose();
                                newWebPart.Dispose();
                            }
                            if (zone == "wpz" && file.InDocumentLibrary)
                            {
                                foreach (SPField field in file.Item.Fields)
                                {
                                    if (!field.ReadOnlyField && field is SPFieldMultiLineText && ((SPFieldMultiLineText)field).WikiLinking && file.Item[field.Id] != null)
                                    {
                                        string content = file.Item[field.Id].ToString();
                                        if (content.Contains(oldWebPart.ID.Replace("_", "-").Substring(2)))
                                        {
                                            Logger.Write("Replacing web part identifier in text field \"{0}\"...", field.InternalName);
                                            if (!test)
                                            {
                                                file.Item[field.Id] = content.Replace(oldWebPart.ID.Replace("_", "-").Substring(2), newWebPart.ID.Replace("_", "-").Substring(2));
                                                file.Item.SystemUpdate();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (manager != null)
                        {
                            manager.Web.Dispose();
                            manager.Dispose();
                        }

                        if (!test)
                        {
                            if (checkIn)
                                file.CheckIn("Checking in changes to page due to web part being replaced with a different type.");
                            if (publish && file.InDocumentLibrary)
                            {
                                PublishItems pi = new PublishItems();
                                pi.PublishListItem(file.Item, file.Item.ParentList, false, "Replace-SPWebPartType", "Checking in changes to page due to web part being replaced with a different type.", null);
                            }
                        }
                    }
                }
                finally
                {
                    if (HttpContext.Current != null && cleanupContext)
                    {
                        HttpContext.Current = null;
                    }
                }

            }
        }

        internal static bool SetProperties(object from, object to, Hashtable properties)
        {
            bool hasErrors = false;
            PropertyInfo[] fromFields = from.GetType().GetProperties();
            PropertyInfo[] toFields = to.GetType().GetProperties();
            for (int f = 0; f < fromFields.Length; f++)
            {
                PropertyInfo fromField = fromFields[f];
                foreach (PropertyInfo toField in toFields)
                {
                    if (toField.Name == fromField.Name && toField.PropertyType == fromField.PropertyType)
                    {
                        if (toField.CanWrite)
                        {
                            Logger.Write("Setting property {0}...", fromField.Name);
                            try
                            {
                                toField.SetValue(to, fromField.GetValue(from, null), null);
                            }
                            catch (Exception ex)
                            {
                                // Doesn't matter that these known fields errored out.
                                if (from is System.Web.UI.Control && fromField.Name == "AppRelativeTemplateSourceDirectory")
                                    break;

#if MOSS
                                if (from is Microsoft.SharePoint.Publishing.WebControls.ContentByQueryWebPart && fromField.Name == "Data")
                                    break; 
#endif
                                Logger.WriteWarning("Unable to set property {0}: {1}", fromField.Name, ex.Message);
                                hasErrors = true;
                            }
                        }
                        break;
                    }
                }
            }
            if (properties != null)
            {
                for (int f = 0; f < toFields.Length; f++)
                {
                    PropertyInfo toField = toFields[f];
                    if (properties.ContainsKey(toField.Name))
                    {
                        if (properties[toField.Name] != null && properties[toField.Name].GetType() == toField.PropertyType)
                        {
                            Logger.Write("Setting property {0}...", toField.Name);
                            try
                            {
                                toField.SetValue(to, properties[toField.Name], null);
                            }
                            catch (Exception ex)
                            {
                                Logger.WriteWarning("Unable to set property {0}: {1}", toField.Name, ex.Message);
                                hasErrors = true;
                            }
                        }
                    }
                }
            }
            return hasErrors;
        }
    }
}
