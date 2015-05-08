using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using Microsoft.SharePoint.Deployment;
using System.IO;
using Microsoft.SharePoint.Administration.Backup;
using System.Collections;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Import", "SPWeb2", SupportsShouldProcess = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("The Import-SPWeb2 cmdlet imports a site collection, Web application, list, or library. This cmdlet extends the capabilities of the Import-SPWeb cmdlet by exposing additional parameters.")]
    [RelatedCmdlets(typeof(SPCmdletExportWeb2), ExternalCmdlets = new[] { "Export-SPWeb", "Import-SPWeb" })]
    public class SPCmdletImportWeb2 : SPCmdletExportImport
    {
        // Fields
        private string m_webName;
        private string m_webParentUrl;

        [Parameter(Mandatory = true, 
            ValueFromPipeline = true, 
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to be exported.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Identity
        {
            get
            {
                return base.GetProp<SPWebPipeBind>("Identity");
            }
            set
            {
                base.SetProp("Identity", value);
            }
        }

        [Parameter(Mandatory = false)]
        public SwitchParameter RetainObjectIdentity { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter SuppressAfterEvents { get; set; }

        [Parameter(Mandatory = false)]
        public SPUpdateVersions UpdateVersions { get; set; }

        protected override void InternalProcessRecord()
        {
            string path = base.Path;
            if (!base.NoFileCompression.IsPresent)
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path))
                {
                    throw new SPException(SPResource.GetString("FileNotFoundExceptionMessage", new object[] { path }));
                }
            }
            else if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
            {
                throw new SPException(SPResource.GetString("DirectoryNotFoundExceptionMessage", new object[] { path }));
            }
            string url = this.Identity.Read().Url;
            if (base.ShouldProcess(string.Format("ShouldProcessImportWeb,{0},{1}", url, base.Path )))
            {
                SPImportSettings settings = new SPImportSettings();
                SPImport import = new SPImport(settings);
                base.SetDeploymentSettings(settings);
                if (base.IncludeUserSecurity.IsPresent)
                {
                    settings.IncludeSecurity = SPIncludeSecurity.All;
                    settings.UserInfoDateTime = SPImportUserInfoDateTimeOption.ImportAll;
                }
                settings.SuppressAfterEvents = SuppressAfterEvents.IsPresent;
                settings.UpdateVersions = this.UpdateVersions;
                char[] trimChars = new char[] { '/' };
                if (url[url.Length - 1] == '/')
                {
                    url = url.TrimEnd(trimChars);
                }
                settings.RetainObjectIdentity = RetainObjectIdentity.IsPresent;
                settings.SiteUrl = url;
                using (SPSite site = new SPSite(url))
                {
                    string str5;
                    Utilities.SplitUrl(Utilities.ConvertToServiceRelUrl(Utilities.GetServerRelUrlFromFullUrl(url), site.ServerRelativeUrl), out str5, out this.m_webName);
                    this.m_webParentUrl = site.ServerRelativeUrl;
                    if (!string.IsNullOrEmpty(str5))
                    {
                        if (!this.m_webParentUrl.EndsWith("/"))
                        {
                            this.m_webParentUrl = this.m_webParentUrl + "/";
                        }
                        this.m_webParentUrl = this.m_webParentUrl + str5;
                    }
                }
                if (this.m_webName == null)
                {
                    this.m_webName = string.Empty;
                }
                EventHandler<SPDeploymentEventArgs> handler = new EventHandler<SPDeploymentEventArgs>(this.OnStarted);
                import.Started += handler;
                try
                {
                    import.Run();
                }
                finally
                {
                    if (!base.NoLogFile.IsPresent)
                    {
                        Console.WriteLine();
                        Console.WriteLine(SPResource.GetString("ExportOperationLogFile", new object[0]));
                        Console.WriteLine("\t{0}", settings.LogFilePath);
                        Console.WriteLine();
                    }
                }
            }
        }

        private void OnStarted(object sender, SPDeploymentEventArgs args)
        {
            SPImportObjectCollection rootObjects = args.RootObjects;
            if (rootObjects.Count != 0)
            {
                if (rootObjects.Count == 1)
                {
                    if (rootObjects[0].Type == SPDeploymentObjectType.Web)
                    {
                        rootObjects[0].TargetParentUrl = this.m_webParentUrl;
                        rootObjects[0].TargetName = this.m_webName;
                    }
                    else
                    {
                        rootObjects[0].TargetParentUrl = this.m_webParentUrl + this.m_webName;
                    }
                }
                else
                {
                    bool rootFound = false;
                    for (int i = 0; i < rootObjects.Count; i++)
                    {
                        rootObjects[i].TargetParentUrl = this.m_webParentUrl;
                        if (rootObjects[i].Type == SPDeploymentObjectType.Web)
                        {
                            if (rootFound)
                            {
                                throw new SPException(SPResource.GetString("ImportOperationMultipleRoots", new object[0]));
                            }
                            rootFound = true;
                            rootObjects[i].TargetName = this.m_webName;
                            return;
                        }
                    }
                }
            }
        }

    }

}
