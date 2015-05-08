using System;
using System.IO;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using System.Management.Automation;

namespace Lapointe.SharePoint.PowerShell.Common.Pages
{
    class ReGhostFile
    {

        /// <summary>
        /// Reghosts the files in site.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="throwOnError">if set to <c>true</c> [throw on error].</param>
        public static void ReghostFilesInSite(SPSite site, bool force, bool throwOnError)
        {
            Logger.Write("Progress: Analyzing files in site collection '{0}'.", site.Url);
            foreach (SPWeb web in site.AllWebs)
            {
                try
                {
                    ReghostFilesInWeb(site, web, false, force, throwOnError);
                }
                finally
                {
                    web.Dispose();
                }
            }
        }

        /// <summary>
        /// Reghosts the files in web.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        /// <param name="recurseWebs">if set to <c>true</c> [recurse webs].</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="throwOnError">if set to <c>true</c> [throw on error].</param>
        public static void ReghostFilesInWeb(SPSite site, SPWeb web, bool recurseWebs, bool force, bool throwOnError)
        {
            Logger.Write("Progress: Analyzing files in web '{0}'.", web.Url);
            foreach (SPFile file in web.Files)
            {
                if (file.CustomizedPageStatus != SPCustomizedPageStatus.Customized && !force)
                    continue;

                Reghost(site, web, file, force, throwOnError);
            }
            foreach (SPList list in web.Lists)
            {
                ReghostFilesInList(site, web, list, force, throwOnError);
            }

            if (recurseWebs)
            {
                foreach (SPWeb childWeb in web.Webs)
                {
                    try
                    {
                        ReghostFilesInWeb(site, childWeb, true, force, throwOnError);
                    }
                    finally
                    {
                        childWeb.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// Reghosts the files in list.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        /// <param name="list">The list.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="throwOnError">if set to <c>true</c> [throw on error].</param>
        public static void ReghostFilesInList(SPSite site, SPWeb web, SPList list, bool force, bool throwOnError)
        {
            if (list.BaseType != SPBaseType.DocumentLibrary)
                return;

            Logger.Write("Progress: Analyzing files in list '{0}'.", list.RootFolder.ServerRelativeUrl);

            foreach (SPListItem item in list.Items)
            {
                if (item.File == null)
                    continue;

                Reghost(site, web, item.File, force, throwOnError);
            }
        }

        /// <summary>
        /// Reghosts the specified file.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <param name="web">The web.</param>
        /// <param name="file">The file.</param>
        /// <param name="force">if set to <c>true</c> [force].</param>
        /// <param name="throwOnError">if set to <c>true</c> [throw on error].</param>
        public static void Reghost(SPSite site, SPWeb web, SPFile file, bool force, bool throwOnError)
        {
            try
            {
                string fileUrl = site.MakeFullUrl(file.ServerRelativeUrl);
                if (file.CustomizedPageStatus != SPCustomizedPageStatus.Customized && !force)
                {
                    Logger.Write("Progress: " + file.ServerRelativeUrl + " was not unghosted (customized).");
                    return;
                }
                if (file.CustomizedPageStatus != SPCustomizedPageStatus.Customized && force)
                {
                    if (!string.IsNullOrEmpty((string)file.Properties["vti_setuppath"]))
                    {
                        file.Properties["vti_hasdefaultcontent"] = "false";

                        string setupPath = (string)file.Properties["vti_setuppath"];
                        string rootPath = Utilities.GetGenericSetupPath("Template");

                        if (!File.Exists(Path.Combine(rootPath, setupPath)))
                        {
                            string message = "The template file (" + Path.Combine(rootPath, setupPath) +
                                             ") does not exist so re-ghosting (uncustomizing) will not be possible.";

                            // something's wrong with the setup path - lets see if we can fix it
                            // Try and remove a leading locale if present
                            setupPath = "SiteTemplates\\" + setupPath.Substring(5);
                            if (File.Exists(Path.Combine(rootPath, setupPath)))
                            {
                                message += "  It appears that a possible template match does exist at \"" +
                                           Path.Combine(rootPath, setupPath) +
                                           "\" however this tool currently is not able to handle pointing the file to the correct template path.  This scenario is most likely due to an upgrade from SPS 2003.";

                                // We found a matching file so reset the property and update the file.
                                // ---  I wish this would work but it simply doesn't - something is preventing the
                                //      update from occuring.  Manipulating the database directly results in a 404
                                //      when attempting to load the "fixed" page so there's gotta be something beyond
                                //      just updating the setuppath property.
                                //file.Properties["vti_setuppath"] = setupPath;
                                //file.Update();
                            }
                            throw new FileNotFoundException(message, setupPath);
                        }
                    }
                }
                Logger.Write("Progress: Re-ghosting (uncustomizing) '{0}'", fileUrl);
                file.RevertContentStream();

                file = web.GetFile(fileUrl);
                if (file.CustomizedPageStatus == SPCustomizedPageStatus.Customized)
                {
                    // Still unsuccessful so take measures further
                    if (force)
                    {
                        object request = Utilities.GetSPRequestObject(web);

                        // I found some cases where calling this directly was the only way to force the re-ghosting of the file.
                        // I think the trick is that it's not updating the file properties after doing the revert (the
                        // RevertContentStream method will call SPRequest.UpdateFileOrFolderProperties() immediately after the 
                        // RevertContentStreams call but ommitting the update call seems to make a difference.
                        Utilities.ExecuteMethod(request, "RevertContentStreams",
                                                new[] { typeof(string), typeof(string), typeof(bool) },
                                                new object[] { web.Url, file.Url, file.CheckOutStatus != SPFile.SPCheckOutStatus.None });


                        Utilities.ExecuteMethod(file, "DirtyThisFileObject", new Type[] { }, new object[] { });

                        file = web.GetFile(fileUrl);

                        if (file.CustomizedPageStatus == SPCustomizedPageStatus.Customized)
                        {
                            throw new SPException("Unable to re-ghost (uncustomize) file " + file.ServerRelativeUrl);
                        }
                        Logger.Write("Progress: " + file.ServerRelativeUrl + " was re-ghosted (uncustomized)!");
                        return;
                    }
                    throw new SPException("Unable to re-ghost (uncustomize) file " + file.ServerRelativeUrl);
                }
                Logger.Write("Progress: " + file.ServerRelativeUrl + " was re-ghosted (uncustomized)!");
            }
            catch (Exception ex)
            {
                if (throwOnError)
                {
                    throw;
                }
                Logger.WriteException(new ErrorRecord(ex, null, ErrorCategory.NotSpecified, file));
            }

        }

    }
}
