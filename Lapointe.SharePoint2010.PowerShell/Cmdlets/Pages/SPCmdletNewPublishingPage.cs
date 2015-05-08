using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using Microsoft.SharePoint.Publishing;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.Collections;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Pages
{
    [Cmdlet(VerbsCommon.New, "SPPublishingPage", SupportsShouldProcess = true), 
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Pages")]
    [CmdletDescription("Creates a new publishing page.")]
    [RelatedCmdlets(typeof(SPCmdletGetPublishingPage), typeof(SPCmdletGetPublishingPageLayout), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> $page = Get-SPWeb http://server_name | New-SPPublishingPage -PageName \"custom.aspx\"  -Title \"Custom Page \" -PageLayout \"articleleft.aspx\"",
        Remarks = "This example creates a page named custom.aspx at http://server_name.")]
    [Example(Code = "PS C:\\> $data = @{\"FieldName1\"=\"Field Value 1\";\"FieldName2\"=\"Field Value 2\"}\r\nPS C:\\> $page = Get-SPWeb http://server_name | New-SPPublishingPage -PageName \"custom.aspx\"  -Title \"Custom Page \" -PageLayout \"articleleft.aspx\" -FieldData $data",
        Remarks = "This example creates a page named custom.aspx at http://server_name and sets two additional fields.")]
    public class SPCmdletNewPublishingPage : SPNewCmdletBaseCustom<PublishingPage>
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to create the publishing page in.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        /// <summary>
        /// Gets or sets the page name.
        /// </summary>
        [Parameter(Mandatory = true,
            HelpMessage = "The name of the publishing page to create. Example: default.aspx")]
        [ValidateNotNullOrEmpty]
        public string PageName { get; set; }


        [ValidateNotNullOrEmpty, Parameter(Mandatory = false)]
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the url to the publishing page.
        /// </summary>
        /// <value>The name of the contentType.</value>
        [Parameter(Mandatory = true, ValueFromPipeline = false,
            HelpMessage = "The filename of the page layout to use.")]
        [ValidateNotNullOrEmpty]
        public SPPageLayoutPipeBind PageLayout { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Key/Value pairs of field names and values to set after the page is created.\r\n\r\nExample: -FieldData @{\"FieldName1\"=\"Field Value 1\";\"FieldName2\"=\"Field Value 2\"}")]
        public Hashtable FieldData { get; set; }

        protected override PublishingPage CreateDataObject()
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

            Dictionary<string, string> dict = new Dictionary<string,string>();
            if (FieldData != null && FieldData.Count > 0)
            {
                foreach (object key in FieldData.Keys)
                {
                    dict.Add(key.ToString(), FieldData[key].ToString());
                }
            }

            using (SPWeb web = Web.Read())
            {
                PageLayout pageLayout = PageLayout.Read(web);
                return Common.Pages.CreatePublishingPage.CreatePage(web, PageName, Title, pageLayout.Name, dict, test);
            }
        }

       
    }
}
