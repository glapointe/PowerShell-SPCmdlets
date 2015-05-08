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
    [Cmdlet(VerbsCommon.Clear, "SPHttpContext", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Site Collections")]
    [CmdletDescription("Clears the HTTP Context. Run this cmdlet when finished with an HTTP Context created by Set-SPHttpContext.", "Clears the HTTP Context. Run this cmdlet when finished with an HTTP Context created by Set-SPHttpContext. The HTTP Context is cleared by setting the static System.Web.HttpContext.Current property to $null (this can be done explicitly or via this cmdlet).")]
    [RelatedCmdlets(typeof(SPCmdletSetHttpContext))]
    [Example(Code = "PS C:\\> $web = Set-SPHttpContext \"http://demo/\"\r\nPS C:\\> Write-Host \"do something with the web\"\r\nPS C:\\> Clear-SPHttpContext",
        Remarks = "This example clears the HttpContext.")]
    public class SPCmdletClearHttpContext : SPCmdletCustom
    {
        protected override void InternalProcessRecord()
        {
            base.InternalProcessRecord();
            System.Web.HttpContext.Current = null;
        }
    }
}
