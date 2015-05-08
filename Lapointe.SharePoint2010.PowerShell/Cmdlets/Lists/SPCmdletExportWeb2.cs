using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System;
using Microsoft.SharePoint.Deployment;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Export", "SPWeb2", SupportsShouldProcess = true),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("The Export-SPWeb2 cmdlet exports a site collection, Web application, list, or library. This cmdlet extends the capabilities of the Export-SPWeb cmdlet by exposing additional parameters.")]
    [RelatedCmdlets(typeof(SPCmdletImportWeb2), ExternalCmdlets = new[] {"Export-SPWeb", "Import-SPWeb"})]
    public class SPCmdletExportWeb2 : SPCmdletExport
    {
        public override SPExportObject ExportObject
        {
            get
            {
                SPWeb web = this.Identity.Read();
                SPExportObject obj2 = new SPExportObject();
                if (string.IsNullOrEmpty(base.ItemUrl))
                {
                    obj2.Id = web.ID;
                    obj2.Type = SPDeploymentObjectType.Web;
                    obj2.Url = web.ServerRelativeUrl;
                }
                else
                {
                    SPList list = web.GetList(base.ItemUrl);
                    if (list == null)
                    {
                        throw new SPException(SPResource.GetString("ExportOperationInvalidUrl", new object[0]));
                    }
                    obj2.Id = list.ID;
                    obj2.Type = SPDeploymentObjectType.List;
                }
                obj2.ExcludeChildren = false;
                return obj2;
            }
        }

        [Parameter(Mandatory = true, 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
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

        public override SPSite Site
        {
            get
            {
                using (SPWeb web = this.Identity.Read())
                {
                    return web.Site;
                }
            }
        }

        public override string SiteUrl
        {
            get
            {
                return this.Identity.Read().Url;
            }
        }
    }

}
