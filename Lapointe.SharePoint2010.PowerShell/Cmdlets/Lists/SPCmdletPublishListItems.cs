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
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Publish", "SPListItems", SupportsShouldProcess = true),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Publish any draft or pending items.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), ExternalCmdlets = new[] {"Get-SPWeb", "Get-SPSite", "Get-SPWebApplication", "Get-SPFarm"})]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/pages\" | Publish-SPListItems -ListItem 1",
        Remarks = "This example publishes item with ID 1 in the pages library located at http://server_name.")]
    [Example(Code = "PS C:\\> Publish-SPListItems -ListItem \"http://server_name/pages/default.aspx\"",
        Remarks = "This example publishes http://server_name/pages/default.aspx.")]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/pages\" | Publish-SPListItems",
        Remarks = "This example publishes all items in the pages library located at http://server_name.")]
    [Example(Code = "PS C:\\> Get-SPWeb \"http://server_name/\" | Publish-SPListItems",
        Remarks = "This example publishes all items in the web located at http://server_name. This will no recurse through sub-webs.")]
    [Example(Code = "PS C:\\> Get-SPSite \"http://server_name/\" | Publish-SPListItems",
        Remarks = "This example publishes all items in the site located at http://server_name. This will recurse through sub-webs.")]
    [Example(Code = "PS C:\\> Get-SPWebApplication \"http://server_name/\" | Publish-SPListItems",
        Remarks = "This example publishes all items in the web application located at http://server_name. This will recurse through all sites and sub-webs.")]
    [Example(Code = "PS C:\\> Get-SPFarm | Publish-SPListItems",
        Remarks = "This example publishes all items in the farm. This will recurse through all web applications, sites, and sub-webs.")]
    public class SPCmdletPublishListItems : SPCmdletCustom
    {
        Common.Lists.PublishItems _itemsPublisher;

        [Parameter(ParameterSetName = "List",
            HelpMessage = "Take over ownership of any files that do not have an existing check-in.")]
        [Parameter(ParameterSetName = "Web",
            HelpMessage = "Take over ownership of any files that do not have an existing check-in.")]
        [Parameter(ParameterSetName = "Site",
            HelpMessage = "Take over ownership of any files that do not have an existing check-in.")]
        [Parameter(ParameterSetName = "WebApplication",
            HelpMessage = "Take over ownership of any files that do not have an existing check-in.")]
        [Parameter(ParameterSetName = "Farm",
            HelpMessage = "Take over ownership of any files that do not have an existing check-in.")]
        [Alias("TakeOver")]
        public SwitchParameter TakeOverFilesWithNoCheckIn { get; set; }

        [Parameter(ParameterSetName = "List",
            HelpMessage = "A regular expression to match against the file name. For example, to only publish Word and Excel files use the following expression: \"\\.((docx)|(xlsx))$\". If specified, list items are ignored (only files are published).")]
        [Parameter(ParameterSetName = "Web",
            HelpMessage = "A regular expression to match against the file name. For example, to only publish Word and Excel files use the following expression: \"\\.((docx)|(xlsx))$\". If specified, list items are ignored (only files are published).")]
        [Parameter(ParameterSetName = "Site",
            HelpMessage = "A regular expression to match against the file name. For example, to only publish Word and Excel files use the following expression: \"\\.((docx)|(xlsx))$\". If specified, list items are ignored (only files are published).")]
        [Parameter(ParameterSetName = "WebApplication",
            HelpMessage = "A regular expression to match against the file name. For example, to only publish Word and Excel files use the following expression: \"\\.((docx)|(xlsx))$\". If specified, list items are ignored (only files are published).")]
        [Parameter(ParameterSetName = "Farm",
            HelpMessage = "A regular expression to match against the file name. For example, to only publish Word and Excel files use the following expression: \"\\.((docx)|(xlsx))$\". If specified, list items are ignored (only files are published).")]
        public string FilterExpression { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The filename to save all details to.")]
        [ValidateDirectoryExistsAndValidFileName]
        public string LogFile { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "ListItem",
            Position = 0,
            HelpMessage = "The list item to publish.")]
        public SPListItemPipeBind ListItem { get; set; }

        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "ListItem",
            Position = 1,
            HelpMessage = "The list containing the item to publish.")]
        public SPListPipeBind ParentList { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "List",
            Position = 0,
            HelpMessage = "The list containing the items to publish.")]
        public SPListPipeBind List { get; set; }


        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Web",
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the list items to publish.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }


        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Site",
            Position = 0,
            HelpMessage = "The site containing the list items to publish.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }


        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "WebApplication",
            Position = 0,
            HelpMessage = "The web application containing the list items to publish.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        public SPWebApplicationPipeBind WebApplication { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Farm",
            Position = 0,
            HelpMessage = "A valid SPFarm object. All items in the farm will be published.")]
        public SPFarmPipeBind Farm { get; set; }

        [Parameter(HelpMessage = "A comment to associated with each checked in, published, and approved item.")]
        public string Comment { get; set; }

        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();
            Logger.LogFile = LogFile;

            _itemsPublisher = new Common.Lists.PublishItems();
        }

        protected override void InternalEndProcessing()
        {
            base.InternalEndProcessing();

            WriteResult(string.Format("Finished Process: {0} Errors, {1} Items(s) Checked In, {2} Item(s) Published, {3} Item(s) Approved",
                _itemsPublisher.TaskCounts.Errors, _itemsPublisher.TaskCounts.Checkin, _itemsPublisher.TaskCounts.Publish, _itemsPublisher.TaskCounts.Approve));
        }

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


            switch (ParameterSetName)
            {
                case "WebApplication":
                    SPWebApplication webApp1 = WebApplication.Read();
                    _itemsPublisher.Publish(webApp1, test, Comment, TakeOverFilesWithNoCheckIn, FilterExpression);
                    break;
                case "Site":
                    using (SPSite site = Site.Read())
                    {
                        _itemsPublisher.Publish(site, test, Comment, TakeOverFilesWithNoCheckIn, FilterExpression);
                    }
                    break;
                case "Web":
                    using (SPWeb web = Web.Read())
                    {
                        try
                        {
                            _itemsPublisher.Publish(web, test, Comment, TakeOverFilesWithNoCheckIn, FilterExpression);
                        }
                        finally
                        {
                            web.Site.Dispose();
                        }
                    }
                    break;
                case "List":
                    SPList list = List.Read();
                    try
                    {
                        _itemsPublisher.Publish(list, test, Comment, TakeOverFilesWithNoCheckIn, FilterExpression);
                    }
                    finally
                    {
                        list.ParentWeb.Dispose();
                        list.ParentWeb.Site.Dispose();
                    }
                    break;
                case "ListItem":
                    SPListItem item = null;
                    if (ParentList != null)
                    {
                        list = ParentList.Read();
                        item = ListItem.Read(list);
                    }
                    else
                    {
                        item = ListItem.Read();
                    }
                    try
                    {
                        _itemsPublisher.Publish(item, test, Comment);
                    }
                    finally
                    {
                        item.ParentList.ParentWeb.Dispose();
                        item.ParentList.ParentWeb.Site.Dispose();
                    }
                    break;
                default:
                    foreach (SPService svc in Farm.Read().Services)
                    {
                        if (!(svc is SPWebService))
                            continue;

                        foreach (SPWebApplication webApp2 in ((SPWebService)svc).WebApplications)
                        {
                            _itemsPublisher.Publish(webApp2, test, Comment, TakeOverFilesWithNoCheckIn, FilterExpression);
                        }
                    }
                    break;
            }

        }

    }
}
