using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
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
    [Cmdlet("Copy", "SPWebPart", SupportsShouldProcess = true, DefaultParameterSetName = "WebPartTitle"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Copies a web part from one file to another file.")]
    [RelatedCmdlets(typeof(SPCmdletGetFile))]
    [Example(Code = "PS C:\\> Copy-SPWebPart -SourceFile \"http://server_name/pages/default.aspx\" -TargetFile \"http://server_name/pages/test.aspx\" -WebPartTitle \"My Web Part\" -Publish",
        Remarks = "This example replaces all instances of the web part who's class name is ContentByQueryWebPart with the web part who's class name is MyContentByQueryWebPart.")]
    public class SPCmdletCopyWebPart : SPCmdletCustom
    {

        [Parameter(Mandatory = true,
            ParameterSetName = "WebPartID",
            HelpMessage = "The ID of the Web Part to copy.")]
        public string WebPartId { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "WebPartTitle",
            HelpMessage = "The title of the Web Part to copy.")]
        public string WebPartTitle { get; set; }


        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        public SPFilePipeBind SourceFile { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        public SPFilePipeBind TargetFile { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "If specified the target page will be published after copying the Web Part.")]
        public SwitchParameter Publish { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The name of the web part zone to copy the web part to.")]
        public string Zone { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The index within the web part zone to copy the web part to.")]
        public int? ZoneIndex { get; set; }


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

            string sourceFileUrl = SourceFile.FileUrl;
            string targetFileUrl = TargetFile.FileUrl;

            string zone = null;
            WebPart sourceWebPart = GetSourceWebPart(sourceFileUrl, WebPartTitle, WebPartId, out zone);
            if (string.IsNullOrEmpty(Zone))
                Zone = zone;

            SetWebPart(sourceWebPart, targetFileUrl, Zone, ZoneIndex, Publish, test);

            base.InternalProcessRecord();
        }
        internal static WebPart GetSourceWebPart(string url, string webPartTitle, string webPartId, out string zone)
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
                    SPLimitedWebPartManager manager = null;
                    try
                    {
                        WebPart webPart = null;
                        if (!string.IsNullOrEmpty(webPartTitle))
                            webPart = Utilities.GetWebPartByTitle(web, url, webPartTitle, out manager);
                        else if (!string.IsNullOrEmpty(webPartId))
                            webPart = Utilities.GetWebPartById(web, url, webPartId, out manager);

                        if (manager != null)
                            zone = manager.GetZoneID(webPart);
                        else
                            zone = null;
                        return webPart;
                    }
                    finally
                    {
                        if (manager != null)
                        {
                            manager.Web.Dispose();
                            manager.Dispose();
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
        internal static void SetWebPart(WebPart sourceWebPart, string targetUrl, string zone, int? zoneIndex, bool publish, bool test)
        {
            if (sourceWebPart.IsClosed)
            {
                sourceWebPart.Dispose();
                throw new Exception("The source web part is closed and cannot be copied.");
            }

            int zoneIndex1 = sourceWebPart.ZoneIndex;
            if (zoneIndex.HasValue)
                zoneIndex1 = zoneIndex.Value;
            Guid storageKey = Guid.NewGuid();
            string id = StorageKeyToID(storageKey);

            using (SPSite site = new SPSite(targetUrl))
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


                    SPFile file = web.GetFile(targetUrl);

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

                    SPLimitedWebPartManager manager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                    try
                    {
                        string wpTitle = sourceWebPart.Title;
                        if (string.IsNullOrEmpty(wpTitle)) wpTitle = sourceWebPart.DisplayTitle;

                        Logger.Write("Copying web part \"{0}\"...", wpTitle);
                        WebPart newWebPart = (WebPart)Activator.CreateInstance(sourceWebPart.GetType());
                        if (SPCmdletReplaceWebPartType.SetProperties(sourceWebPart, newWebPart, null))
                        {
                            Logger.WriteWarning("An error was encountered setting web part properties so try one more time in case the error is the result of a sequencing issue.");
                            if (SPCmdletReplaceWebPartType.SetProperties(sourceWebPart, newWebPart, null))
                            {
                                Logger.WriteWarning("Unable to set all properties for web part.");
                            }
                        }

                        try
                        {
                            if (!test)
                            {
                                newWebPart.ID = id;

                                manager.AddWebPart(newWebPart, zone, zoneIndex1);
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

                                    svc.AddWebPartToZone(targetUrl, sb.ToString(), Storage.Shared, zone, zoneIndex1);

                                }
                                catch (SoapException ex)
                                {
                                    throw new Exception(ex.Detail.OuterXml);
                                }
                            }
                        }
                        finally
                        {
                            sourceWebPart.Dispose();
                            newWebPart.Dispose();
                        }
                        if (zone == "wpz" && file.InDocumentLibrary)
                        {
                            foreach (SPField field in file.Item.Fields)
                            {
                                if (!field.ReadOnlyField && field is SPFieldMultiLineText && ((SPFieldMultiLineText)field).WikiLinking)
                                {
                                    string content = null;
                                    if (file.Item[field.Id] != null)
                                        content = file.Item[field.Id].ToString();

                                    string div = string.Format(CultureInfo.InvariantCulture, "<div class=\"ms-rtestate-read ms-rte-wpbox\" contentEditable=\"false\"><div class=\"ms-rtestate-read {0}\" id=\"div_{0}\"></div><div style='display:none' id=\"vid_{0}\"/></div>", new object[] { storageKey.ToString("D") });
                                    content += div;
                                    Logger.Write("Adding web part to text field \"{0}\"...", field.InternalName);
                                    if (!test)
                                    {
                                        file.Item[field.Id] = content;
                                        file.Item.SystemUpdate();
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
                                pi.PublishListItem(file.Item, file.Item.ParentList, false, "Copy-SPWebPart", "Checking in changes to page due to web part being copied from another page.", null);
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
        internal static string StorageKeyToID(Guid storageKey)
        {
            if (!(Guid.Empty == storageKey))
            {
                return ("g_" + storageKey.ToString().Replace('-', '_'));
            }
            return string.Empty;
        }
    }
}
