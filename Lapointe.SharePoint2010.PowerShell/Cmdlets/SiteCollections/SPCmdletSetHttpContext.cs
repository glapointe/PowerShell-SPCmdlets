using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.Win32;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation.Internal;
using Lapointe.SharePoint.PowerShell.Common;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using System.Web;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.SiteCollections
{
    [Cmdlet(VerbsCommon.Set, "SPHttpContext", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true,RequireUserFarmAdmin = false)]
    [CmdletGroup("Site Collections")]
    [CmdletDescription("Sets the HTTP Context for a given Site and returns a new SPWeb object associated with the new context. When done with a given Site be sure to run Clear-SPHttpContext in order to remove the context.", "Many native SharePoint objects, such as Web Parts, require a valid HTTP Context in order for certain properties to be available. When working within PowerShell there is no HTTP Context so it is necessary to create one. This is accomplished by setting the System.Web.HttpContext.Current static property to a valid object instantiated using the provided SPWeb object. Once created there are SharePoint specific properties that must be set on this object. This cmdlet sets the following (where $context is the HttpContext object): \r\n\t1. $context.Items[\"FormDigestValidated\"] = true;\r\n\t2. $context.User = System.Threading.Thread.CurrentPrincipal;\r\n\t3. $context.Items[\"Microsoft.SharePoint.SPServiceContext\"] = Microsoft.SharePoint.SPServiceContext.GetContext($web.Site);\r\n\t4. Microsoft.SharePoint.Web.Controls.SPControl.SetContextSite($context, $web.Site);\r\n\t5. Microsoft.SharePoint.Web.Controls.SPControl.SetContextWeb($context, $web);\r\nWhen you are finished working with a given Site run Clear-SPHttpContext to set the HttpContext object back to null.")]
    [RelatedCmdlets(typeof(SPCmdletClearHttpContext), ExternalCmdlets = new[] { "Get-SPWeb" })]
    [Example(Code = "PS C:\\> $web = Set-SPHttpContext \"http://demo/\"\r\nPS C:\\> Write-Host \"do something with the web\"\r\nPS C:\\> Clear-SPHttpContext",
        Remarks = "This example sets the HttpContext to http://demo.")]
    public class SPCmdletSetHttpContext : SPSetCmdletBaseCustom<SPWeb>
    {
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The web to set the HTTP Context to.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Identity { get; set; }

        protected override void InternalValidate()
        {
            if (Identity != null)
                DataObject = Identity.Read();
        }

        protected override void UpdateDataObject()
        {
            HttpRequest httpRequest = new HttpRequest("", DataObject.Url, "");
            System.IO.StringWriter writer = new System.IO.StringWriter();
            HttpResponse httpResponse = new HttpResponse(writer);
            HttpContext httpContext = new HttpContext(httpRequest, httpResponse);
            HttpContext.Current = httpContext;
            Microsoft.SharePoint.WebControls.SPControl.SetContextSite(httpContext, DataObject.Site);
            Microsoft.SharePoint.WebControls.SPControl.SetContextWeb(httpContext, DataObject);
            httpContext.Items["HttpHandlerSPWebApplication"] = DataObject.Site.WebApplication;
            httpContext.Items["Microsoft.SharePoint.Admin.GlobalAdmin"] = new SPGlobalAdmin();
            DataObject.AllowUnsafeUpdates = true;
            httpContext.Items["FormDigestValidated"] = true;
            httpContext.User = System.Threading.Thread.CurrentPrincipal;
            httpContext.Items["Microsoft.SharePoint.SPServiceContext"] = Microsoft.SharePoint.SPServiceContext.GetContext(DataObject.Site);

            WriteObject(DataObject);
        }
    }
}
