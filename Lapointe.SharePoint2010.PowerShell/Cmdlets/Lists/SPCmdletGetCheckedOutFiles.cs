using System;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.SharePoint.PowerShell.Common.Lists;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Utilities;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    public class SPCheckedOutFile
    {
        internal Microsoft.SharePoint.SPCheckedOutFile FileWithNoCheckIn { get; set; }
        internal SPFile File { get; set; }
        public Guid SiteId { get; internal set; }
        public Guid WebId { get; internal set; }
        public Guid ListId { get; internal set; }
        public int ListItemId { get; internal set; }
        public SPUser CheckedOutBy { get; internal set; }
        public string CheckedOutByEmail { get; internal set; }
        public int CheckedOutById { get; internal set; }
        public string CheckedOutByName { get; internal set; }
        public string Url { get; internal set; }
        public string DirName { get; internal set; }
        public string ImageUrl { get; internal set; }
        public string LeafName { get; internal set; }
        public long Length { get; internal set; }
        public DateTime TimeLastModified { get; internal set; }
        public void TakeOverCheckOut()
        {
            if (FileWithNoCheckIn != null)
                FileWithNoCheckIn.TakeOverCheckOut();
        }
        public void Delete()
        {
            if (FileWithNoCheckIn != null)
                FileWithNoCheckIn.Delete();
            else if (File != null)
                File.Delete();
        }
        public void CheckIn()
        {
            CheckIn(null);
        }
        public void CheckIn(string comment)
        {
            SPListItem item = GetListItem();
            
            item.File.CheckIn(comment);
        }
        public void Publish()
        {
            Publish(null);
        }
        public void Publish(string comment)
        {
            PublishItems pi = new PublishItems();
            SPListItem item = GetListItem();
            pi.PublishListItem(item, item.ParentList, false, "Get-SPCheckedOutFiles", comment, null);
            if (pi.TaskCounts.Errors > 0)
                throw new Exception("An error occurred publishing the item.");
        }
        private SPListItem GetListItem()
        {
            SPListItem item = null;
            if (FileWithNoCheckIn != null)
            {
                FileWithNoCheckIn.TakeOverCheckOut();
                using (SPSite site = new SPSite(SiteId))
                using (SPWeb web = site.OpenWeb(WebId))
                {
                    item = web.Lists[ListId].GetItemById(ListItemId);
                }
            }
            else if (File != null)
                item = File.Item;

            if (item == null)
                throw new Exception("Unable to retrieve list item.");

            return item;
        }
    }

    [SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true, RequireUserMachineAdmin = false)]
    [CmdletGroup("Lists")]
    [Cmdlet(VerbsCommon.Get, "SPCheckedOutFiles", DefaultParameterSetName = "SPSite")]
    [CmdletDescription("Retrieves check out details for a given List, Web, or Site Collection.")]
    [RelatedCmdlets(typeof(SPCmdletGetFile))]
    [Example(Code = "PS C:\\> Get-SPCheckedOutFiles -Site \"http://server_name/\"",
        Remarks = "This example outputs a list of files that are checked out for the given Site Collection")]
    public class SPCmdletGetCheckedOutFiles : SPGetCmdletBaseCustom<SPCheckedOutFile>
    {
        [Parameter(Mandatory = true, 
            ParameterSetName = "SPSite",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Site to inspect.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "SPWeb",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web to inspect.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPWeb object.")]
        [Parameter(Mandatory = false, 
            ParameterSetName = "SPList",
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 1,
            HelpMessage = "Specifies the URL or GUID of the Web containing the list.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The list whose checked out files are to be returned. \r\n\r\nThe value must be a valid URL in the form http://server_name/lists/listname or /lists/listname. If a server relative URL is provided then the Web parameter must be provided.",
            ParameterSetName = "SPList")]
        public SPListPipeBind List { get; set; }

        [Parameter(ParameterSetName = "SPWeb",
            HelpMessage = "Excludes all child sites and only considers the specified site.")]
        public SwitchParameter ExcludeChildWebs { get; set; }

        protected override IEnumerable<SPCheckedOutFile> RetrieveDataObjects()
        {
            switch (ParameterSetName)
            {
                case "SPSite":
                    using (SPSite site = Site.Read())
                    {
                        foreach (SPWeb web in site.AllWebs)
                        {
                            InspectWeb(web, false);
                            web.Dispose();
                        }
                    }
                    break;
                case "SPWeb":
                    using (SPWeb web = Web.Read())
                    {
                        InspectWeb(web, !ExcludeChildWebs);
                    }
                    break;
                case "SPList":
                    SPWeb parentWeb = null;
                    if (Web != null)
                        parentWeb = Web.Read();
                    SPList list = List.Read(parentWeb);
                    try
                    {
                        InspectLibrary(list);
                    }
                    finally
                    {
                        if (list != null)
                        {
                            list.ParentWeb.Dispose();
                            list.ParentWeb.Site.Dispose();
                        }
                        if (parentWeb != null)
                        {
                            parentWeb.Dispose();
                            parentWeb.Site.Dispose();
                        }
                    }
                    break;
            }
            return null;
        }


        private void InspectWeb(SPWeb web, bool includeChildren)
        {
            foreach (SPList list in web.Lists)
            {
                InspectLibrary(list);
            }
            if (includeChildren)
            {
                foreach (SPWeb child in web.Webs)
                {
                    InspectWeb(child, true);
                }
            }
        }

        private void InspectLibrary(SPList list)
        {
            if (!(list is SPDocumentLibrary))
                return;

            SPDocumentLibrary docLib = (SPDocumentLibrary) list;
            foreach (Microsoft.SharePoint.SPCheckedOutFile file in docLib.CheckedOutFiles)
            {
                var details = new SPCheckedOutFile()
                                  {
                                      FileWithNoCheckIn = file,
                                      File = null,
                                      Url = list.ParentWeb.Site.MakeFullUrl(SPUrlUtility.CombineUrl(list.ParentWeb.ServerRelativeUrl, file.Url)),
                                      SiteId = list.ParentWeb.Site.ID,
                                      WebId = list.ParentWeb.ID,
                                      ListId = list.ID,
                                      ListItemId = file.ListItemId,
                                      CheckedOutBy = file.CheckedOutBy,
                                      CheckedOutByEmail = file.CheckedOutByEmail,
                                      CheckedOutById = file.CheckedOutById,
                                      CheckedOutByName = file.CheckedOutByName,
                                      DirName = file.DirName,
                                      ImageUrl = file.ImageUrl,
                                      LeafName = file.LeafName,
                                      Length = file.Length,
                                      TimeLastModified = file.TimeLastModified
                                  };
                WriteObject(details);
            }
            foreach (SPListItem item in docLib.Items)
            {
                if (!Utilities.IsCheckedOut(item))
                    continue;
                if (docLib.CheckedOutFiles.Any(f => f.ListItemId == item.ID))
                    continue;

                SPFile file = item.File;
                var details = new SPCheckedOutFile()
                {
                    FileWithNoCheckIn = null,
                    File = file,
                    Url = list.ParentWeb.Site.MakeFullUrl(SPUrlUtility.CombineUrl(list.ParentWeb.ServerRelativeUrl, file.Url)),
                    SiteId = list.ParentWeb.Site.ID,
                    WebId = list.ParentWeb.ID,
                    ListId = list.ID,
                    ListItemId = item.ID,
                    CheckedOutBy = file.CheckedOutByUser,
                    CheckedOutByEmail = file.CheckedOutByUser.Email,
                    CheckedOutById = file.CheckedOutByUser.ID,
                    CheckedOutByName = file.CheckedOutByUser.Name,
                    DirName = file.ParentFolder.Url,
                    ImageUrl = "/_layouts/images/" + SPUtility.MapToIcon(list.ParentWeb, file.Name, string.Empty),
                    LeafName = file.Name,
                    Length = file.Length,
                    TimeLastModified = file.TimeLastModified
                };
                WriteObject(details);
            }
        }
    }
}
