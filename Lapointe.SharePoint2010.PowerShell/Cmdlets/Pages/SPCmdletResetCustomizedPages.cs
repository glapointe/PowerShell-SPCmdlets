using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using System.IO;
using System.Collections;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Publishing;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Pages
{
    [Cmdlet(VerbsCommon.Reset, "SPCustomizedPages", SupportsShouldProcess = true, DefaultParameterSetName = "SPWeb"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Pages")]
    [CmdletDescription("Resets customized (unghosted) pages to their site definition.")]
    [RelatedCmdlets(typeof(SPCmdletGetCustomizedPages), typeof(Lists.SPCmdletGetList), ExternalCmdlets = new[] { "Get-SPWeb", "Get-SPSite", "Get-SPWebApplication" })]
    [Example(Code = "PS C:\\> Get-SPWeb http://server_name | Reset-SPCustomizedPages",
        Remarks = "This example resets all unghosted pages in http://server_name without updating any child webs.")]
    [Example(Code = "PS C:\\> Get-SPSite http://server_name | Reset-SPCustomizedPages",
        Remarks = "This example resets all unghosted pages in http://server_name including all child webs.")]
    [Example(Code = "PS C:\\> Reset-SPCustomizedPages -File http://server_name/pages/default.aspx",
        Remarks = "This example resets the unghosted page http://server_name/pages/default.aspx.")]
    public class SPCmdletResetCustomizedPages : SPCmdletCustom
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the pages to reset.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [Parameter(ParameterSetName = "SPList",
            Mandatory = false,
            HelpMessage = "Specifies the URL or GUID of the Web containing the pages to reset.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        /// <summary>
        /// Gets or sets the site.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPSite",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The site containing the pages to reset. All sub-webs will be iterated through.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [ValidateNotNull]
        public SPSitePipeBind Site { get; set; }

        [Parameter(ParameterSetName = "SPFile",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to the file to reset.")]
        public SPFilePipeBind File { get; set; }

        [Parameter(ParameterSetName = "SPList",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The URL to the list containing the files to reset.")]
        public SPListPipeBind List { get; set; }

        [Parameter(ParameterSetName = "SPWebApplication",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The web application containing the files to reset.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        public SPWebApplicationPipeBind WebApplication { get; set; }

        /// <summary>
        /// Gets or sets whether to recurse all webs
        /// </summary>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = false,
            HelpMessage = "Recurse all Webs.")]
        public SwitchParameter Recurse { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Abort execution if an error is encountered.")]
        public SwitchParameter HaltOnError { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Use of the Force parameter is not supported and should only be used as a last resort.")]
        public SwitchParameter Force { get; set; }

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
            if (test)
                Logger.Verbose = true;

            bool force = Force;
            bool haltOnError = HaltOnError.IsPresent;

            switch (ParameterSetName)
            {
                case "SPFile":

                    SPFile file = File.Read();
                    try
                    {
                        if (!file.Exists)
                        {
                            throw new FileNotFoundException(string.Format("File '{0}' not found.", File.FileUrl), File.FileUrl);
                        }

                        Common.Pages.ReGhostFile.Reghost(file.Web.Site, file.Web, file, force, haltOnError);
                    }
                    finally
                    {
                        if (file != null)
                        {
                            file.Web.Dispose();
                            file.Web.Site.Dispose();
                        }
                    }
                    break;
                case "SPList":
                    SPList list = null;
                    SPWeb web1 = null;

                    try
                    {
                        if (Web != null)
                        {
                            web1 = Web.Read();
                            if (web1 == null)
                                throw new FileNotFoundException(string.Format("The specified site could not be found. {0}", Web.WebUrl), Web.WebUrl);
                            list = List.Read(web1);
                        }
                        else
                            list = List.Read();

                        if (list == null)
                            throw new FileNotFoundException(string.Format("The specified list could not be found. {0}", List.ListUrl), List.ListUrl);

                        Common.Pages.ReGhostFile.ReghostFilesInList(list.ParentWeb.Site, list.ParentWeb, list, force, haltOnError);
                    }
                    finally
                    {
                        if (web1 != null)
                        {
                            web1.Dispose();
                            web1.Site.Dispose();
                        }
                    }
                    break;
                case "SPWeb":
                    bool recurseWebs = Recurse.IsPresent;
                    SPWeb web2 = Web.Read();
                    if (web2 == null)
                        throw new FileNotFoundException(string.Format("The specified site could not be found. {0}", Web.WebUrl), Web.WebUrl);

                    try
                    {
                        Common.Pages.ReGhostFile.ReghostFilesInWeb(web2.Site, web2, recurseWebs, force, haltOnError);
                    }
                    finally
                    {
                        if (web2 != null)
                        {
                            web2.Dispose();
                            web2.Site.Dispose();
                        }
                    }
                    break;
                case "SPSite":
                    SPSite site1 = Site.Read();
                    if (site1 == null)
                        throw new FileNotFoundException(string.Format("The specified site collection could not be found. {0}", Site.SiteUrl), Site.SiteUrl);

                    try
                    {
                        Common.Pages.ReGhostFile.ReghostFilesInSite(site1, force, haltOnError);
                    }
                    finally
                    {
                        if (site1 != null)
                            site1.Dispose();
                    }
                    break;
                case "SPWebApplication":
                    SPWebApplication webApp = WebApplication.Read();
                    if (webApp == null)
                        throw new FileNotFoundException("The specified web application could not be found.");
 
                    
                    Logger.Write("Progress: Analyzing files in web application '{0}'.", webApp.GetResponseUri(SPUrlZone.Default).ToString());

                    foreach (SPSite site2 in webApp.Sites)
                    {
                        try
                        {
                            Common.Pages.ReGhostFile.ReghostFilesInSite(site2, force, haltOnError);
                        }
                        finally
                        {
                            site2.Dispose();
                        }
                    }
                    break;

            }
        }

       
    }
}
