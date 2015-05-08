using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Utilities;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.SiteCollections
{
    [Cmdlet(VerbsCommon.Move, "SPWeb", SupportsShouldProcess = false),
        SPCmdlet(RequireLocalFarmExist = true)]
    [CmdletGroup("Site Collections")]
    [CmdletDescription("Move an SPWeb from one URL to another.")]
    [RelatedCmdlets(typeof(SPCmdletDeleteList), typeof(SPCmdletCopyList), typeof(SPCmdletCopyListSecurity),
        typeof(SPCmdletExportListSecurity), ExternalCmdlets = new[] { "Get-SPWeb", "Start-SPAssignment", "Stop-SPAssignment" })]
    [Example(Code = "PS C:\\> Move-SPWeb -Identity \"http://portal/foo/bar\" -Parent \"http://portal/\"",
        Remarks = "This example moves the Web located at http://portal/foo/bar to http://portal/bar.")]
    [Example(Code = "PS C:\\> Move-SPWeb -Identity \"http://portal/foo/bar\" -Parent \"http://portal/\" -UrlName \"foobar\"",
        Remarks = "This example moves the Web located at http://portal/foo/bar to http://portal/foobar.")]
    [Example(Code = "PS C:\\> Move-SPWeb -Identity \"http://portal/foo/bar\" -UrlName \"foobar\"",
        Remarks = "This example moves the Web located at http://portal/foo/bar to http://portal/foo/foobar.")]
    public class SPCmdletMoveWeb : SPSetCmdletBaseCustom<SPWeb>
    {

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            HelpMessage = "Specifies the URL or GUID of the existing Web to move.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Identity { get; set; }


        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            HelpMessage = "Specifies the URL of the parent Web to move the Web to. This will not change the URL name of the current Web.")]
        public SPWebPipeBind Parent { get; set; }

        [Parameter(Mandatory = false,
            ParameterSetName = "Target",
            ValueFromPipeline = true,
            HelpMessage = "Specifies the URL name of the Web. If not specified then the current URL name is used. For example, if moving http://portal/foo/bar, \"bar\" is the URL name of the Web.")]
        public string UrlName { get; set; }


        [Parameter(HelpMessage = "When moving to a new Site Collection, halts the export or import of the Web if an warning occurs.")]
        public SwitchParameter HaltOnWarning { get; set; }

        [Parameter(HelpMessage = "When moving to a new Site Collection, halts the export or import of the Web if an error occurs.")]
        public SwitchParameter HaltOnFatalError { get; set; }

        [Parameter(HelpMessage = "When moving to a new Site Collection, includes user security.")]
        public SwitchParameter IncludeUserSecurity { get; set; }

        [Parameter(HelpMessage = "When moving to a new Site Collection, disable the firing of \"After\" events when creating or modifying list items.")]
        public SwitchParameter SuppressAfterEvents { get; set; }

        [Parameter(HelpMessage = "When moving to a new Site Collection, retains the identity of all (lists and sub-webs).")]
        public SwitchParameter RetainObjectIdentity { get; set; }

        [Parameter(HelpMessage = "When moving to a new Site Collection, the -TempPath parameter specifies where the export files will be stored.")]
        [ValidateDirectoryExists]
        public string TempPath { get; set; }


        protected override void InternalValidate()
        {
            base.InternalValidate();
            if (Parent == null && string.IsNullOrEmpty(UrlName))
                throw new SPCmdletException("You must specify either the Parent and/or the UrlName parameter.");
        }

        protected override void UpdateDataObject()
        {
            SPWeb sourceWeb = null;
            SPWeb parentWeb = null;
            try
            {
                sourceWeb = Identity.Read();
                if (Parent != null)
                {
                    parentWeb = Parent.Read();

                    if (sourceWeb.ID == parentWeb.ID)
                    {
                        throw new SPCmdletException("Source web and parent web cannot be the same.");
                    }
                    if (sourceWeb.ParentWeb != null && sourceWeb.ParentWeb.ID == parentWeb.ID && string.IsNullOrEmpty(UrlName))
                    {
                        throw new Exception("Parent web specified matches the source web's current parent - move is not necessary. Specify the -UrlName parameter to change the URL of the Web within the parent Web.");
                    }
                    if (sourceWeb.IsRootWeb && sourceWeb.Site.ID == parentWeb.Site.ID)
                    {
                        throw new SPCmdletException("Cannot move root web within the same site collection.");
                    }
                    if (sourceWeb.IsRootWeb && RetainObjectIdentity)
                    {
                        // If we allow retainobjectidentity when moving a root web the import will attempt to import over
                        // the parent web's site collection which would be really bad.
                        throw new SPCmdletException("Cannot move a root web when \"-Retainobjectidentity\" is used.");
                    }

                    if (sourceWeb.Site.ID == parentWeb.Site.ID)
                    {
                        string urlName = UrlName;
                        if (string.IsNullOrEmpty(urlName))
                            urlName = sourceWeb.Name;
                        Common.SiteCollections.MoveWeb.MoveWebWithinSite(sourceWeb, parentWeb, urlName);
                    }
                    else
                    {
                        Common.SiteCollections.MoveWeb.MoveWebOutsideSite(sourceWeb, parentWeb, UrlName, RetainObjectIdentity, HaltOnWarning, HaltOnFatalError, IncludeUserSecurity, SuppressAfterEvents, TempPath);
                    }
                }
                else
                {
                    // All we're doing is changing the URL name of the web.
                    Common.SiteCollections.MoveWeb.MoveWebWithinSite(sourceWeb, sourceWeb.ParentWeb, UrlName);
                }
            }
            finally
            {
                if (sourceWeb != null)
                    sourceWeb.Dispose();
                if (parentWeb != null)
                    parentWeb.Dispose();
            }
        }

    }
}
