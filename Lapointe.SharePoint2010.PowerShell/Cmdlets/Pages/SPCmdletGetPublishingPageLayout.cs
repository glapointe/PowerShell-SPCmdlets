using System;
using System.Collections.Generic;
using System.Management.Automation;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint.Publishing;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Pages
{
    [Cmdlet(VerbsCommon.Get, "SPPublishingPageLayout", SupportsShouldProcess = false, DefaultParameterSetName = "SPWeb"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Pages")]
    [CmdletDescription("Retrieves all publishing page layouts from the specified source.")]
    [RelatedCmdlets(typeof(SPCmdletGetPublishingPage), typeof(SPCmdletRepairPageLayoutUrl), typeof(ContentTypes.SPCmdletGetContentType), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> $layouts = Get-SPWeb http://server_name | Get-SPPublishingPageLayout",
        Remarks = "This example returns back all page layouts from http://server_name.")]
    [Example(Code = "PS C:\\> $layout = Get-SPPublishingPageLayout \"http://server_name/_catalogs/masterpage/articleleft.aspx\"",
        Remarks = "This example returns back the articleleft.aspx page layout from the http://server_name web.")]
    [Example(Code = "PS C:\\> $layout = $Get-SPWeb http://server_name | Get-SPPublishingPageLayout -PageLayout \"articleleft.aspx\"",
        Remarks = "This example returns back the articleleft.aspx page layout from http://server_name.")]
    public class SPCmdletGetPublishingPageLayout : SPGetCmdletBaseCustom<PageLayout>
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "PageLayout",
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to retrieve publishing page layouts from.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to retrieve publishing page layouts from.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        /// <summary>
        /// Gets or sets the page name.
        /// </summary>
        [Parameter(ParameterSetName = "PageLayout",
            Mandatory = true,
            HelpMessage = "The name or path to the page layout to return. If the value is the page name only then the Web parameter must be provided.\r\n\r\nExample: ArticleLeft.aspx or http://<site>/_catalogs/masterpage/articleleft.aspx.")]
        [ValidateNotNull]
        public SPPageLayoutPipeBind[] PageLayout { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The associated content type to filter the returned page layouts by.")]
        [Alias("ContentType")]
        public SPContentTypePipeBind AssociatedContentType { get; set; }

        protected override IEnumerable<PageLayout> RetrieveDataObjects()
        {
            List<PageLayout> layouts = new List<PageLayout>();

            SPWeb web = null;
            if (Web != null)
            {
                web = Web.Read();
                
                if (!PublishingWeb.IsPublishingWeb(web))
                    throw new ArgumentException("The specified web is not a publishing web.");

                AssignmentCollection.Add(web);
                AssignmentCollection.Add(web.Site);
            }

            switch (ParameterSetName)
        	{
                case "PageLayout":
                    foreach (SPPageLayoutPipeBind pipeBind in PageLayout)
                    {
                        PageLayout layout = pipeBind.Read(web);
                        SPContentType ct1 = null;
                        if (AssociatedContentType != null)
                        {
                            ct1 = AssociatedContentType.Read(layout.ListItem.Web);
                        }
                        if (ct1 == null || ct1.Id == layout.AssociatedContentType.Id)
                            WriteResult(layout);
                    }
                    break;
                case "SPWeb":
                    PublishingWeb pubWeb = PublishingWeb.GetPublishingWeb(web);
                    SPContentType ct2 = null;
                    if (AssociatedContentType != null)
                    {
                        ct2 = AssociatedContentType.Read(web);
                    }
                    if (ct2 == null)
                        WriteResult(pubWeb.GetAvailablePageLayouts());
                    else
                        WriteResult(pubWeb.GetAvailablePageLayouts(ct2.Id));
                    break;
        		default:
                    break;
        	}

            return null;
        }
    }
}
