using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Workflow;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    public class PublishItems
    {
        public Counts TaskCounts = new Counts();

        #region Publish Methods

        /// <summary>
        /// Publishes all list items for a given web application.
        /// </summary>
        /// <param name="webApp">The web app.</param>
        /// <param name="test">If true test the change only (don't make any changes).</param>
        /// <param name="comment">The comment.</param>
        /// <param name="takeOver">if set to <c>true</c> [take over].</param>
        /// <param name="filterExpression">The filter expression.</param>
        public void Publish(SPWebApplication webApp, bool test, string comment, bool takeOver, string filterExpression)
        {
            Logger.Write("Processing Web Application: " + webApp.DisplayName);

            foreach (SPSite site in webApp.Sites)
            {
                try
                {
                    Publish(site, test, comment, takeOver, filterExpression);
                }
                finally
                {
                    site.Dispose();
                }
            }

            Logger.Write("Finished Processing Web Application: " + webApp.DisplayName + "\r\n");
        }

        /// <summary>
        /// Publishes all list items for a given site collection.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="test">If true test the change only (don't make any changes).</param>
        /// <param name="comment">The comment.</param>
        /// <param name="takeOver">if set to <c>true</c> [take over].</param>
        /// <param name="filterExpression">The filter expression.</param>
        public void Publish(SPSite site, bool test, string comment, bool takeOver, string filterExpression)
        {
            Logger.Write("Processing Site: " + site.ServerRelativeUrl);

            foreach (SPWeb web in site.AllWebs)
            {
                try
                {
                    Publish(web, test, comment, takeOver, filterExpression);
                }
                finally
                {
                    web.Dispose();
                }
            }

            Logger.Write("Finished Processing Site: " + site.ServerRelativeUrl + "\r\n");
        }

        /// <summary>
        /// Publishes all list items for a given web.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="test">If true test the change only (don't make any changes).</param>
        /// <param name="comment">The comment.</param>
        /// <param name="takeOver">if set to <c>true</c> [take over].</param>
        /// <param name="filterExpression">The filter expression.</param>
        public void Publish(SPWeb web, bool test, string comment, bool takeOver, string filterExpression)
        {
            Logger.Write("Processing Web: " + web.ServerRelativeUrl);

            foreach (SPList list in web.Lists)
            {
                Publish(list, test, comment, takeOver, filterExpression);
            }

            Logger.Write("Finished Processing Web: " + web.ServerRelativeUrl + "\r\n");
        }

        /// <summary>
        /// Publishes all list items for a given list.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="test">If true test the change only (don't make any changes).</param>
        /// <param name="comment">The comment.</param>
        /// <param name="takeOver">if set to <c>true</c> [take over].</param>
        /// <param name="filterExpression">The filter expression.</param>
        public void Publish(SPList list, bool test, string comment, bool takeOver, string filterExpression)
        {
            Logger.Write("Processing List: " + list.DefaultViewUrl);

            string source = "\"Publish-SPListItems\"";

            foreach (SPListItem item in list.Items)
            {
                PublishListItem(item, list, test, source, comment, filterExpression);
            }

            foreach (SPListItem item in list.Folders)
            {
                PublishListItem(item, list, test, source, comment, filterExpression);
            }

            if (list is SPDocumentLibrary && takeOver)
            {
                foreach (SPCheckedOutFile file in ((SPDocumentLibrary)list).CheckedOutFiles)
                {
                    file.TakeOverCheckOut();
                    SPListItem item = list.GetItemById(file.ListItemId);
                    PublishListItem(item, list, test, source, comment, filterExpression);
                }
            }

            Logger.Write("Finished Processing List: " + list.DefaultViewUrl + "\r\n");
        }

        /// <summary>
        /// Publishes the specified item.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="test">If true test the change only (don't make any changes).</param>
        /// <param name="comment">The comment.</param>
        public void Publish(SPListItem item, bool test, string comment)
        {
            Logger.Write("Processing Item: " + item.Url);

            string source = "\"Publish-SPListItems\"";

            PublishListItem(item, item.ParentList, test, source, comment, null);

            Logger.Write("Finished Processing Item: " + item.Url + "\r\n");
        }

        #endregion

        #region Primary Worker Methods

        /// <summary>
        /// Publishes the list item.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="list">The list.</param>
        /// <param name="test">If true test the change only (don't make any changes).</param>
        /// <param name="source">The source.</param>
        /// <param name="comment">The comment.</param>
        /// <param name="filterExpression">The filter expression.</param>
        public void PublishListItem(SPListItem item, SPList list, bool test, string source, string comment, string filterExpression)
        {
            if (TaskCounts == null)
                TaskCounts = new Counts();

            string title = item.ID.ToString();
            if (item.Fields.ContainsField("Title"))
                title = item.Title;

            try
            {
                item = item.ParentList.GetItemById(item.ID);
                if (item.File == null && !string.IsNullOrEmpty(filterExpression))
                    return;

                if (item.File != null)
                {
                    if (!string.IsNullOrEmpty(filterExpression))
                    {
                        string fileName = item.File.Name;
                        Regex regex = new Regex(filterExpression, RegexOptions.IgnoreCase);
                        if (!regex.IsMatch(fileName))
                            return;
                    }

                    // We first need to handle the case in which we have a file which means that
                    // we have to deal with the possibility that the file may be checked out.
                    if (item.Level == SPFileLevel.Checkout)
                    {
                        // The file is checked out so we now need to check it in - we'll do a major
                        // checkin which will result in it being published.
                        if (!test)
                        {
                            item.File.CheckIn(comment??"Checked in by " + source, SPCheckinType.MajorCheckIn);
                            // We need to get the File's version of the SPListItem so that we get the changes.
                            // Calling item.Update() will fail because the file is no longer checked out.
                            // If workflow is supported this should now be in a pending state.
                            // Re-retrieve the item to avoid save conflict errors.
                            item = item.ParentList.GetItemById(item.ID);
                        }
                        TaskCounts.Checkin++;
                        TaskCounts.Publish++; // The major checkin causes it to be published so we'll track that as well.
                        Logger.Write("Checked in item: {0} ({1})", title, item.Url);
                    }
                    else if (item.Level == SPFileLevel.Draft && item.ModerationInformation == null)
                    {
                        // The file isn't checked out but it is in a draft state so we need to publish it.
                        if (!test)
                        {
                            if (Utilities.IsCheckedOut(item))
                            {
                                item.File.CheckIn(comment ?? "Checked in by " + source, SPCheckinType.MajorCheckIn);
                                TaskCounts.Checkin++;
                                Logger.Write("Checked in item: {0} ({1})", title, item.Url);
                            }
                            if (item.ParentList.EnableMinorVersions)
                                item.File.Publish(comment ?? "Published by " + source);
                            // We need to get the File's version of the SPListItem so that we get the changes.
                            // Calling item.Update() will fail because the file is no longer checked out.
                            // If workflow is supported this should now be in a pending state.
                            // Re-retrieve the item to avoid save conflict errors.
                            item = item.ParentList.GetItemById(item.ID);
                        }
                        TaskCounts.Publish++;
                        Logger.Write("Published item: {0} ({1})", title, item.Url);
                    }
                    else if (item.Level == SPFileLevel.Published && Utilities.IsCheckedOut(item))
                    {
                        // This technically shouldn't be possible but apparently it is.
                        if (!test)
                        {
                            item.File.CheckIn(comment ?? "Checked in by " + source, SPCheckinType.MajorCheckIn);
                            if (item.ParentList.EnableMinorVersions)
                                item.File.Publish(comment ?? "Published by " + source);
                            item = item.ParentList.GetItemById(item.ID);
                        }
                        TaskCounts.Checkin++;
                        Logger.Write("Checked in item: {0} ({1})", title, item.Url);
                        TaskCounts.Publish++;
                        Logger.Write("Published item: {0} ({1})", title, item.Url);
                    }
                }
            }
            catch (Exception ex)
            {
                TaskCounts.Errors++;
                Logger.WriteException(new System.Management.Automation.ErrorRecord(new SPException("An error occured checking in an item", ex), null, System.Management.Automation.ErrorCategory.NotSpecified, item));
            }

            if (item.ModerationInformation != null)
            {
                // If ModerationInformation is not null then the item supports content approval.
                if (item.File == null &&
                    (item.ModerationInformation.Status == SPModerationStatusType.Draft ||
                    item.ModerationInformation.Status == SPModerationStatusType.Pending))
                {
                    // If content approval is supported but no file is associated with the item then we have
                    // to treat it differently.  We simply set the status information directly.
                    try
                    {
                        if (!test)
                        {
                            // Because the SPListItem object has no direct approval method we have to 
                            // set the information directly (there's no SPFile object to use).
                            CancelWorkflows(false, list, item);
                            item.ModerationInformation.Status = SPModerationStatusType.Approved;
                            item.ModerationInformation.Comment = comment ?? "Approved by " + source;
                            item.Update(); // Because there's no SPFile object we don't have to worry about the item being checkedout for this to succeed as you can't check it out.
                            // Re-retrieve the item to avoid save conflict errors.
                            item = item.ParentList.GetItemById(item.ID);
                        }
                        TaskCounts.Approve++;
                        Logger.Write("Approved item: {0} ({1})", title, item.Url);
                    }
                    catch (Exception ex)
                    {
                        TaskCounts.Errors++;
                        Logger.WriteException(new System.Management.Automation.ErrorRecord(new SPException("An error occured approving an item.", ex), null, System.Management.Automation.ErrorCategory.NotSpecified, item));
                    }
                }
                else
                {
                    // The item supports content approval and we have an SPFile object to work with.
                    try
                    {
                        if (item.ModerationInformation.Status == SPModerationStatusType.Pending)
                        {
                            // The item is pending so it's already been published - we just need to approve.
                            if (!test)
                            {
                                // Cancel any workflows.
                                CancelWorkflows(false, list, item);
                                item.File.Approve(comment ?? "Approved by " + source);
                                // Re-retrieve the item to avoid save conflict errors.
                                item = item.ParentList.GetItemById(item.ID);
                            }
                            TaskCounts.Approve++;
                            Logger.Write("Approved item: {0} ({1})", title, item.Url);
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskCounts.Errors++;
                        Logger.WriteException(new System.Management.Automation.ErrorRecord(new SPException("An error occured approving an item.", ex), null, System.Management.Automation.ErrorCategory.NotSpecified, item));
                    }

                    try
                    {
                        if (item.ModerationInformation.Status == SPModerationStatusType.Draft)
                        {
                            // The item is in a draft state so we have to first publish it and then approve it.
                            if (!test)
                            {
                                if (Utilities.IsCheckedOut(item))
                                {
                                    item.File.CheckIn(comment ?? "Checked in by " + source, SPCheckinType.MajorCheckIn);
                                    TaskCounts.Checkin++;
                                    Logger.Write("Checked in item: {0} ({1})", title, item.Url);
                                }
                                if (item.ParentList.EnableMinorVersions)
                                    item.File.Publish(comment ?? "Published by " + source);
                                // Cancel any workflows.
                                CancelWorkflows(test, list, item);
                                item.File.Approve(comment ?? "Approved by " + source);
                                // We don't need to re-retrieve the item as we're now done with it.
                            }
                            TaskCounts.Publish++;
                            TaskCounts.Approve++;
                            Logger.Write("Published item: {0} ({1})", title, item.Url);
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskCounts.Errors++;
                        Logger.WriteException(new System.Management.Automation.ErrorRecord(new SPException("An error occured approving an item.", ex), null, System.Management.Automation.ErrorCategory.NotSpecified, item));
                    }
                }
            }
        }

        /// <summary>
        /// Cancels the workflows.  This code is a re-engineering of the code that Microsoft uses
        /// when approving an item via the browser.  That code is in Microsoft.SharePoint.ApplicationPages.ApprovePage.
        /// </summary>
        /// <param name="test">If true test the change only (don't make any changes).</param>
        /// <param name="list">The list.</param>
        /// <param name="item">The item.</param>
        private void CancelWorkflows(bool test, SPList list, SPListItem item)
        {
            if (list.DefaultContentApprovalWorkflowId != Guid.Empty &&
                item.DoesUserHavePermissions((SPBasePermissions.ApproveItems |
                                              SPBasePermissions.EditListItems)))
            {
                // If the user has rights to do so then we need to cancel any workflows that
                // are associated with the item.
                SPSecurity.RunWithElevatedPrivileges(
                    delegate
                    {
                        foreach (SPWorkflow workflow in item.Workflows)
                        {
                            if (workflow.ParentAssociation.Id !=
                                list.DefaultContentApprovalWorkflowId)
                            {
                                continue;
                            }
                            SPWorkflowManager.CancelWorkflow(workflow);
                            Logger.Write("Cancelling workflow {0} for item: {1} ({2})", workflow.WebId.ToString(), item.ID.ToString(), item.Url);
                        }
                    });
            }
        }

        #endregion


        /// <summary>
        /// Data class for storing counts of various actions performed.
        /// </summary>
        public class Counts
        {
            public int Checkin = 0;
            public int Approve = 0;
            public int Publish = 0;
            public int Errors = 0;
        }

    }
}
