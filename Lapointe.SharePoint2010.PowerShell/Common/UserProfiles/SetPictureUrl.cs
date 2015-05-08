using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Text;
using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.UserProfiles
{
    public class SetPictureUrl
    {
        /// <summary>
        /// Sets the pictures to the specified path for all user profiles.
        /// </summary>
        /// <param name="profManager">The prof manager.</param>
        /// <param name="path">The path.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <param name="ignoreMissingData">if set to <c>true</c> [ignore missing data].</param>
        /// <param name="validateUrl">if set to <c>true</c> validate the url.</param>
        public static void SetPictures(UserProfileManager profManager, string path, bool overwrite, bool ignoreMissingData, bool validateUrl)
        {
            foreach (UserProfile profile in profManager)
            {
                SetPicture(profile, path, overwrite, ignoreMissingData, validateUrl);
            }
        }

        /// <summary>
        /// Sets the picture URL for the specfied user.
        /// </summary>
        /// <param name="profManager">The prof manager.</param>
        /// <param name="username">The username.</param>
        /// <param name="path">The path.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <param name="ignoreMissingData">if set to <c>true</c> [ignore missing data].</param>
        /// <param name="validateUrl">if set to <c>true</c> validate the url.</param>
        public static void SetPicture(UserProfileManager profManager, string username, string path, bool overwrite, bool ignoreMissingData, bool validateUrl)
        {
            if (!String.IsNullOrEmpty(username))
            {
                if (!profManager.UserExists(username))
                {
                    throw new SPException("The username specified cannot be found.");
                }
                UserProfile profile = profManager.GetUserProfile(username);
                SetPicture(profile, path, overwrite, ignoreMissingData, validateUrl);
            }
            else
                throw new ArgumentNullException("username", "The username parameter cannot be null or empty.");
        }

        /// <summary>
        /// Sets the picture.
        /// </summary>
        /// <param name="up">Up.</param>
        /// <param name="path">The path.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <param name="ignoreMissingData">if set to <c>true</c> [ignore missing data].</param>
        /// <param name="validateUrl">if set to <c>true</c> validate the url.</param>
        public static void SetPicture(UserProfile up, string path, bool overwrite, bool ignoreMissingData, bool validateUrl)
        {
            if (up["PictureURL"].Value != null && !String.IsNullOrEmpty(up["PictureURL"].Value.ToString()) && !overwrite)
            {
                Logger.WriteWarning("\"{0}\" already contains a picture URL. Specify -Overwrite to replace existing settings.",
                                    up["AccountName"].Value.ToString());
                return;
            }
            if (String.IsNullOrEmpty(path))
            {
                path = String.Empty;
            }
            else
            {
                if (path.Contains("$(username)") || path.Contains("@(username)"))
                {
                    path = path.Replace("@(username)", "$(username)");
                    if (up["UserName"] != null && up["UserName"].Value != null)
                        path = path.Replace("$(username)", up["UserName"].Value.ToString());
                    else
                    {
                        if (up["AccountName"] != null && up["AccountName"].Value != null)
                            path = path.Replace("$(username)", up["AccountName"].Value.ToString().Split('\\')[1]);
                        else
                        {
                            if (!ignoreMissingData)
                                throw new ArgumentException(String.Format("Unable to determine username from existing profile data ({0}).", up.ID));
                            return;
                        }
                    }
                }

                if (path.Contains("$(domain)") || path.Contains("@(domain)"))
                {
                    path = path.Replace("@(domain)", "$(domain)");
                    if (up["AccountName"] != null && up["AccountName"].Value != null)
                        path = path.Replace("$(domain)", up["AccountName"].Value.ToString().Split('\\')[0]);
                    else
                    {
                        if (!ignoreMissingData)
                            throw new ArgumentException(String.Format("Unable to determine domain from existing profile data ({0}).", up.ID));
                        return;
                    }
                }

                if (path.Contains("$(email)") || path.Contains("@(email)"))
                {
                    path = path.Replace("@(email)", "$(email)");
                    if (up["WorkEmail"] != null && up["WorkEmail"].Value != null)
                        path = path.Replace("$(email)", up["WorkEmail"].Value.ToString());
                    else
                    {
                        if (!ignoreMissingData)
                            throw new ArgumentException(String.Format("Unable to determine email from existing profile data ({0}).", up.ID));
                        return;
                    }
                }

                if (path.Contains("$(firstname)") || path.Contains("@(firstname)"))
                {
                    path = path.Replace("@(firstname)", "$(firstname)");
                    if (up["FirstName"] != null && up["FirstName"].Value != null)
                        path = path.Replace("$(firstname)", up["FirstName"].Value.ToString());
                    else
                    {
                        if (!ignoreMissingData)
                            throw new ArgumentException(String.Format("Unable to determine first name from existing profile data ({0}).", up.ID));
                        return;
                    }
                }

                if (path.Contains("$(lastname)") || path.Contains("@(lastname)"))
                {
                    path = path.Replace("@(lastname)", "$(lastname)");
                    if (up["LastName"] != null && up["LastName"].Value != null)
                        path = path.Replace("$(lastname)", up["LastName"].Value.ToString());
                    else
                    {
                        if (!ignoreMissingData)
                            throw new ArgumentException(String.Format("Unable to determine lastname from existing profile data ({0}).", up.ID));
                        return;
                    }
                }

                if (path.Contains("$(employeeid)") || path.Contains("@(employeeid)"))
                {
                    path = path.Replace("@(employeeid)", "$(employeeid)");
                    if (up["EmployeeID"] != null && up["EmployeeID"].Value != null)
                    {
                        path = path.Replace("$(employeeid)", up["EmployeeID"].Value.ToString());
                    }
                    else
                    {
                        if (!ignoreMissingData)
                            throw new ArgumentException(String.Format("Unable to determine Employee ID from existing profile data ({0}).", up.ID));
                        return;
                    }
                }
            }

            if (validateUrl && !String.IsNullOrEmpty(path))
            {
                Logger.Write("Validating URL \"{0}\" for \"{1}\".", path, up["AccountName"].Value.ToString());

                try
                {
                    //Create a request for the URL. 
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(path);
                    request.AllowAutoRedirect = false;
                    request.Credentials = CredentialCache.DefaultCredentials;
                    HttpWebResponse serverResponse = (HttpWebResponse)request.GetResponse();
                    if (serverResponse.StatusCode != HttpStatusCode.OK)
                    {
                        Logger.Write("Unable to find picture. Setting PictureURL property to empty string.");
                        path = String.Empty;
                    }
                    serverResponse.Close();
                }
                catch (Exception ex)
                {
                    Exception ex1 = new Exception(String.Format("Exception occured validating URL \"{0}\" (property not updated):\r\n{1}", path, ex.Message), ex);
                    Logger.WriteException(new ErrorRecord(ex1, null, ErrorCategory.InvalidData, up));
                    return;
                }
            }

            Logger.Write("Setting picture for \"{0}\" to \"{1}\".", up["AccountName"].Value.ToString(), path);

            up["PictureURL"].Value = path;
            up.Commit();
        }
    }
}
