using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.WebPartPages;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;
using Lapointe.SharePoint.PowerShell.Common.Pages;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebParts
{
    [Cmdlet("Import", "SPWebPart", DefaultParameterSetName = "File"),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Imports a web part to a page.")]
    [RelatedCmdlets(typeof(SPCmdletGetLimitedWebPartManager), typeof(Pages.SPCmdletGetPublishingPage), typeof(Lists.SPCmdletGetFile))]
    [Example(Code = "PS C:\\> Import-SPWebPart -File \"http://portal/pages/default.aspx\" -Identity \"c:\\my.webpart\" -WebPartTitle \"My Web part\" -Zone \"Left\" -ZoneIndex 0 -DeleteExisting -Publish",
        Remarks = "This example adds a web part to the page http://portal/pages/default.aspx.")]
    public class SPCmdletImportWebPart : SPNewCmdletBaseCustom<WebPart>
    {
        [Parameter(Mandatory = true,
           ParameterSetName = "Manager_File",
           ValueFromPipeline = true,
           ValueFromPipelineByPropertyName = true,
           Position = 0,
           HelpMessage = "The URL to a web part page or an instance of an SPLimitedWebPartManager object.")]
        [Parameter(Mandatory = true,
           ParameterSetName = "Manager_File_WikiPage",
           ValueFromPipeline = true,
           ValueFromPipelineByPropertyName = true,
           Position = 0,
           HelpMessage = "The URL to a web part page or an instance of an SPLimitedWebPartManager object.")]
        [Parameter(Mandatory = true,
           ParameterSetName = "Manager_Assembly",
           ValueFromPipeline = true,
           ValueFromPipelineByPropertyName = true,
           Position = 0,
           HelpMessage = "The URL to a web part page or an instance of an SPLimitedWebPartManager object.")]
        [Parameter(Mandatory = true,
           ParameterSetName = "Manager_Assembly_WikiPage",
           ValueFromPipeline = true,
           ValueFromPipelineByPropertyName = true,
           Position = 0,
           HelpMessage = "The URL to a web part page or an instance of an SPLimitedWebPartManager object.")]
        public SPLimitedWebPartManagerPipeBind Manager { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "File_File",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "File_File_WikiPage",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "File_Assembly",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "File_Assembly_WikiPage",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to a web part page or an instance of an SPFile object.")]
        [Alias(new string[] { "Url", "Page" })]
        public SPFilePipeBind File { get; set; }

        [Parameter(Mandatory = false, 
            HelpMessage = "The title to set the web part to.")]
        public string WebPartTitle { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "File_File",
            HelpMessage = "The name of the web part zone to add the web part to.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Manager_File",
            HelpMessage = "The name of the web part zone to add the web part to.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "File_Assembly",
            HelpMessage = "The name of the web part zone to add the web part to.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Manager_Assembly",
            HelpMessage = "The name of the web part zone to add the web part to.")]
        public string Zone { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "File_File",
            HelpMessage = "The index within the web part zone to add the web part to.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_File",
            HelpMessage = "The index within the web part zone to add the web part to.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "File_Assembly",
            HelpMessage = "The index within the web part zone to add the web part to.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_Assembly",
            HelpMessage = "The index within the web part zone to add the web part to.")]
        public int ZoneIndex { get; set; }

        [Parameter(
            ParameterSetName = "File_File_WikiPage",
            Mandatory = true,
            HelpMessage = "The zone to add the web part to.")]
        [Parameter(
            ParameterSetName = "Manager_File_WikiPage",
            Mandatory = true,
            HelpMessage = "The zone to add the web part to.")]
        [Parameter(
            ParameterSetName = "File_Assembly_WikiPage",
            Mandatory = true,
            HelpMessage = "The zone to add the web part to.")]
        [Parameter(
            ParameterSetName = "Manager_Assembly_WikiPage",
            Mandatory = true,
            HelpMessage = "The zone to add the web part to.")]
        public int Row { get; set; }

        [Parameter(
            ParameterSetName = "File_File_WikiPage",
            Mandatory = true,
            HelpMessage = "The zone index to add the web part to.")]
        [Parameter(
            ParameterSetName = "Manager_File_WikiPage",
            Mandatory = true,
            HelpMessage = "The zone index to add the web part to.")]
        [Parameter(
            ParameterSetName = "File_Assembly_WikiPage",
            Mandatory = true,
            HelpMessage = "The zone index to add the web part to.")]
        [Parameter(
            ParameterSetName = "Manager_Assembly_WikiPage",
            Mandatory = true,
            HelpMessage = "The zone index to add the web part to.")]
        public int Column { get; set; }

        [Parameter(
            ParameterSetName = "File_File_WikiPage",
            Mandatory = false,
            HelpMessage = "Add some space before the web part.")]
        [Parameter(
            ParameterSetName = "Manager_File_WikiPage",
            Mandatory = false,
            HelpMessage = "Add some space before the web part.")]
        [Parameter(
            ParameterSetName = "File_Assembly_WikiPage",
            Mandatory = false,
            HelpMessage = "Add some space before the web part.")]
        [Parameter(
            ParameterSetName = "Manager_Assembly_WikiPage",
            Mandatory = false,
            HelpMessage = "Add some space before the web part.")]
        public SwitchParameter AddSpace { get; set; }

        [Parameter(Mandatory = false, 
            HelpMessage = "If specified, the page will be published after importing the web part.")]
        public SwitchParameter Publish { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "File_File",
            HelpMessage = "If specified, any existing web parts with the same name as the imported one will be deleted.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_File",
            HelpMessage = "If specified, any existing web parts with the same name as the imported one will be deleted.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "File_Assembly",
            HelpMessage = "If specified, any existing web parts with the same name as the imported one will be deleted.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_Assembly",
            HelpMessage = "If specified, any existing web parts with the same name as the imported one will be deleted.")]
        public SwitchParameter DeleteExisting { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "File_File",
            HelpMessage = "Each key represents the name of a string to replace with the specified value.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_File",
            HelpMessage = "Each key represents the name of a string to replace with the specified value.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "File_Assembly",
            HelpMessage = "Each key represents the name of a string to replace with the specified value.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_Assembly",
            HelpMessage = "Each key represents the name of a string to replace with the specified value.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "File_File_WikiPage",
            HelpMessage = "Each key represents the name of a string to replace with the specified value.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_File_WikiPage",
            HelpMessage = "Each key represents the name of a string to replace with the specified value.")]
        public PropertiesPipeBind CustomReplaceText { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "File_File",
            HelpMessage = "The path to the web part file or a valid web part file in the form of an XmlDocument object.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Manager_File",
            HelpMessage = "The path to the web part file or a valid web part file in the form of an XmlDocument object.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "File_File_WikiPage",
            HelpMessage = "The path to the web part file or a valid web part file in the form of an XmlDocument object.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Manager_File_WikiPage",
            HelpMessage = "The path to the web part file or a valid web part file in the form of an XmlDocument object.")]
        public XmlDocumentPipeBind Identity { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "File_Assembly",
            HelpMessage = "Specify the full name of the assembly containing the web part class to add.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_Assembly",
            HelpMessage = "Specify the full name of the assembly containing the web part class to add.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "File_Assembly_WikiPage",
            HelpMessage = "Specify the full name of the assembly containing the web part class to add.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_Assembly_WikiPage",
            HelpMessage = "Specify the full name of the assembly containing the web part class to add.")]
        public string Assembly { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "File_Assembly",
            HelpMessage = "Specify the full name of the web part class to add.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_Assembly",
            HelpMessage = "Specify the full name of the web part class to add.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "File_Assembly_WikiPage",
            HelpMessage = "Specify the full name of the web part class to add.")]
        [Parameter(Mandatory = false,
            ParameterSetName = "Manager_Assembly_WikiPage",
            HelpMessage = "Specify the full name of the web part class to add.")]
        public string TypeName { get; set; }


        [Parameter(Mandatory = false, HelpMessage = "The chrome settings for the web part.")]
        public PartChromeType ChromeType
        {
            get
            {
                if (Fields["ChromeType"] == null)
                    return PartChromeType.Default;
                return (PartChromeType)Fields["ChromeType"];
            }
            set { Fields["ChromeType"] = value; }
        }

        protected override WebPart CreateDataObject()
        {
            string pageUrl = null;
            switch (ParameterSetName)
            {
                case "Manager_File":
                case "Manager_Assembly":
                case "Manager_File_WikiPage":
                case "Manager_Assembly_WikiPage":
                    pageUrl = Manager.PageUrl;
                    break;
                case "File_File":
                case "File_Assembly":
                case "File_File_WikiPage":
                case "File_Assembly_WikiPage":
                    pageUrl = File.FileUrl;
                    break;
            }
            if (pageUrl == null)
                throw new SPCmdletObjectNotFoundException("No page was specified.");

            SPFile file = (new SPFilePipeBind(pageUrl)).Read();
            Hashtable props = null;
            if (CustomReplaceText != null)
                props = CustomReplaceText.Read();

            switch (ParameterSetName)
            {
                case "File_Assembly":
                case "Manager_Assembly":
                    return AddWebPart(file, null, Assembly, TypeName, WebPartTitle, Zone, ZoneIndex, DeleteExisting, props, ChromeType, Publish);
                case "File_File":
                case "Manager_File":
                    return AddWebPart(file, Identity.Read().OuterXml, null, null, WebPartTitle, Zone, ZoneIndex, DeleteExisting, props, ChromeType, Publish);
                case "File_File_WikiPage":
                case "Manager_File_WikiPage":
                    return WikiPageUtilities.AddWebPartToWikiPage(file.Item, Identity.Read().OuterXml, WebPartTitle, Row, Column, AddSpace, props, ChromeType, Publish);
                case "File_Assembly_WikiPage":
                case "Manager_Assembly_WikiPage":
                    var wp = AddWebPart(Assembly, TypeName);
                    return WikiPageUtilities.AddWebPartToWikiPage(file.Item, wp, WebPartTitle, Row, Column, AddSpace, ChromeType, Publish);
            }

            return null;
        }


        /// <summary>
        /// Adds the web part.
        /// </summary>
        /// <param name="file">The page.</param>
        /// <param name="webPartXml">The web part XML file.</param>
        /// <param name="webPartTitle">The web part title.</param>
        /// <param name="zone">The zone.</param>
        /// <param name="zoneId">The zone id.</param>
        /// <param name="deleteWebPart">if set to <c>true</c> [delete web part].</param>
        /// <param name="customReplaceText">The custom replace text.</param>
        /// <param name="chromeType">Type of the chrome.</param>
        /// <param name="publish">if set to <c>true</c> [publish].</param>
        /// <returns></returns>
        public WebPart AddWebPart(SPFile file, string webPartXml, string assembly, string typeName, string webPartTitle, string zone, int zoneId, bool deleteWebPart, Hashtable customReplaceText, PartChromeType chromeType, bool publish)
        {
            bool cleanupContext = false;
            bool checkBackIn = false;

            if (file.InDocumentLibrary)
            {
                if (!Utilities.IsCheckedOut(file.Item) || !Utilities.IsCheckedOutByCurrentUser(file.Item))
                {
                    checkBackIn = true;
                    file.CheckOut();
                }
                // If it's checked out by another user then this will throw an informative exception so let it do so.
            }

            if (HttpContext.Current == null)
            {
                cleanupContext = true;
                HttpRequest httpRequest = new HttpRequest("", file.Item.ParentList.ParentWeb.Url, "");
                HttpContext.Current = new HttpContext(httpRequest, new HttpResponse(new StringWriter()));
                SPControl.SetContextWeb(HttpContext.Current, file.Item.ParentList.ParentWeb);
            }

            string url = file.Item.ParentList.ParentWeb.Site.MakeFullUrl(file.ServerRelativeUrl);
            using (SPLimitedWebPartManager manager = file.Item.ParentList.ParentWeb.GetLimitedWebPartManager(url, PersonalizationScope.Shared))
            {
                try
                {
                    WebPart wp;
                    if (!string.IsNullOrEmpty(webPartXml))
                        wp = AddWebPart(manager, file, webPartXml, customReplaceText);
                    else
                        wp = AddWebPart(assembly, typeName);

                    if (!string.IsNullOrEmpty(webPartTitle))
                    {
                        wp.Title = webPartTitle;
                    }
                    webPartTitle = wp.Title;
                    wp.ChromeType = chromeType;

                    // Delete existing web part with same title so that we only have the latest version on the page
                    foreach (WebPart wpTemp in manager.WebParts)
                    {
                        try
                        {
                            if (wpTemp.Title == wp.Title)
                            {
                                if (deleteWebPart)
                                {
                                    manager.DeleteWebPart(wpTemp);
                                    break;
                                }
                                else
                                    continue;
                            }
                        }
                        finally
                        {
                            wpTemp.Dispose();
                        }
                    }

                    try
                    {
                        manager.AddWebPart(wp, zone, zoneId);
                    }
                    catch (Exception)
                    {
                        System.Net.ServicePointManager.ServerCertificateValidationCallback += delegate { return true; };

                        // We've not already added the web part so use the web service to do this.
                        using (WebPartPagesWebService.WebPartPagesWebService svc = new WebPartPagesWebService.WebPartPagesWebService())
                        {
                            // We failed adding via the OM so try the web service as a fall back.
                            svc.Url = manager.Web.Url + "/_vti_bin/WebPartPages.asmx";
                            svc.Credentials = System.Net.CredentialCache.DefaultCredentials;

                            try
                            {
                                // Add the web part to the web service.  We use a web service because many
                                // web parts require the SPContext.Current variables to be set which are
                                // not set when run from a command line.
                                StringBuilder sb = new StringBuilder();
                                XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
                                xmlWriter.Formatting = Formatting.Indented;
                                manager.ExportWebPart(wp, xmlWriter);
                                xmlWriter.Flush();

                                svc.AddWebPartToZone(url, sb.ToString(), WebPartPagesWebService.Storage.Shared, zone, zoneId);
                            }
                            catch (System.Web.Services.Protocols.SoapException ex)
                            {
                                throw new Exception(ex.Detail.OuterXml);
                            }
                        }
                    }
                    return wp;
                }
                finally
                {
                    if (cleanupContext)
                    {
                        HttpContext.Current = null;
                    }
                    if (manager != null)
                    {
                        manager.Web.Dispose();
                        manager.Dispose();
                    }

                    if (file.InDocumentLibrary && Utilities.IsCheckedOut(file.Item) && (checkBackIn || publish))
                        file.CheckIn("Checking in changes to page due to new web part being added: " + webPartTitle);

                    if (publish && file.InDocumentLibrary && file.Item.ParentList.EnableMinorVersions)
                    {
                        try
                        {
                            file.Publish("Publishing changes to page due to new web part being added: " + webPartTitle);
                            if (file.Item.ModerationInformation != null)
                            {
                                file.Approve("Approving changes to page due to new web part being added: " + webPartTitle);
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteWarning("Unable to publish or approve file: " + ex.Message);
                        }
                    }

                }
           }

        }

        internal static WebPart AddWebPart(SPLimitedWebPartManager manager, SPFile file, string webPartXml, Hashtable customReplaceText)
        {
            WebPart wp;
            XmlTextReader reader = null;
            try
            {
                webPartXml = webPartXml.Replace("${siteCollection}", file.Item.ParentList.ParentWeb.Site.Url);
                webPartXml = webPartXml.Replace("${site}", file.Item.ParentList.ParentWeb.Url);
                webPartXml = webPartXml.Replace("${webTitle}", HttpUtility.HtmlEncode(file.Item.ParentList.ParentWeb.Title));
                if (customReplaceText != null)
                {
                    foreach (string key in customReplaceText.Keys)
                    {
                        webPartXml = webPartXml.Replace(key, HttpUtility.HtmlEncode(customReplaceText[key].ToString()));
                    }
                }
                reader = new XmlTextReader(new StringReader(webPartXml));

                string err;
                wp = manager.ImportWebPart(reader, out err);
                if (!string.IsNullOrEmpty(err))
                    throw new Exception(err);

            }
            finally
            {
                if (reader != null)
                    reader.Close();
            }
            return wp;
        }

        /// <summary>
        /// Adds the web part.
        /// </summary>
        /// <param name="assemblyName">Name of the assembly.</param>
        /// <param name="typeName">Name of the type.</param>
        /// <returns></returns>
        private static WebPart AddWebPart(string assemblyName, string typeName)
        {
            //now try loading the assembly
            Assembly wpAsm;
            WebPart xWp;
            try
            {
                wpAsm = System.Reflection.Assembly.Load(assemblyName);
            }
            catch (Exception asmEx)
            {
                //some error logging here about invalid assembly name
                throw new SPException("Error loading assembly name " + assemblyName + ":\r\nt" + Utilities.FormatException(asmEx));
            }

            if (wpAsm == null)
                throw new SPException("Error loading assembly name " + assemblyName);

            //try creating an instance of the class
            try
            {
                xWp = (WebPart)wpAsm.CreateInstance(typeName);
            }
            catch (Exception instEx)
            {
                //some error logging here about invalid class name
                throw new SPException("Error creating instance of class " + typeName + ":\r\n" + Utilities.FormatException(instEx));
            }

            if (xWp == null)
                throw new SPException("Error creating instance of class " + typeName);

            return xWp;
        }
    }
}
