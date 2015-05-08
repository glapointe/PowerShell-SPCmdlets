using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.New, "SPList", SupportsShouldProcess = true),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Creates a new list.")]
    [RelatedCmdlets(typeof(SPCmdletGetList))]
    [Example(Code = "PS C:\\> $list = Get-SPWeb http://server_name | New-SPList -UrlName \"MyDocs\" -Title \"My Docs\" -FeatureId \"00BFEA71-E717-4E80-AA17-D0C71B360101\" -TemplateType 101 -DocTemplateType 100",
        Remarks = "This example creates a standard document library at http://server_name/MyDocs.")]
    public class SPCmdletNewList : SPNewCmdletBaseCustom<SPList>
    {
        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to create the list in.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }


        [ValidateNotNullOrEmpty, Parameter(Mandatory = true, 
            HelpMessage = "The list title. Example: \"My List\"")]
        public string Title { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The URL name of the list to create. Example: \"MyList\"")]
        [ValidateNotNullOrEmpty]
        public string UrlName { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The feature ID to which the list definition belongs.")]
        public Guid FeatureId { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "An integer corresponding to the list definition type.")]
        public int TemplateType { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The list description.")]
        public string Description { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The ID for the document template type.")]
        public int DocTemplateType { get; set; }

        protected override SPList CreateDataObject()
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


            using (SPWeb web = Web.Read())
            {
                if (!test)
                    return Common.Lists.AddList.Add(web.Lists, UrlName, Title, Description, FeatureId, TemplateType, DocTemplateType.ToString());
            }
            return null;
        }
    }
}
