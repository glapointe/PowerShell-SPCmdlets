using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Xml;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.StsAdmin;
using Microsoft.SharePoint.Utilities;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    public class ReplaceFieldValues
    {
        public List<SearchReplaceData> SearchStrings = new List<SearchReplaceData>();
        public List<string> FieldName = new List<string>();
        public bool Publish;
        public bool Quiet;
        public bool Test;
        public string LogFile;
        public bool UseInternalFieldName;

        #region ReplaceValues Methods

        /// <summary>
        /// Replaces the values.
        /// </summary>
        /// <param name="webApp">The web app.</param>
        public void ReplaceValues(SPWebApplication webApp)
        {
            Log("Processing Web Application: " + webApp.DisplayName);

            foreach (SPSite site in webApp.Sites)
            {
                try
                {
                    ReplaceValues(site);
                }
                finally
                {
                    site.Dispose();
                }
            }

            Log("Finished Processing Web Application: " + webApp.DisplayName + "\r\n");
        }

        /// <summary>
        /// Replaces the values.
        /// </summary>
        /// <param name="site">The site.</param>
        public void ReplaceValues(SPSite site)
        {
            Log("Processing Site: " + site.ServerRelativeUrl);

            foreach (SPWeb web in site.AllWebs)
            {
                try
                {
                    ReplaceValues(web);
                }
                finally
                {
                    web.Dispose();
                }
            }

            Log("Finished Processing Site: " + site.ServerRelativeUrl + "\r\n");
        }

        /// <summary>
        /// Replaces the values.
        /// </summary>
        /// <param name="web">The web.</param>
        public void ReplaceValues(SPWeb web)
        {
            Log("Processing Web: " + web.Url);

            foreach (SPList list in web.Lists)
            {
                ReplaceValues(list);
            }

            Log("Finished Processing Web: " + web.Url + "\r\n");
        }

        /// <summary>
        /// Replaces the values.
        /// </summary>
        /// <param name="list">The list.</param>
        public void ReplaceValues(SPList list)
        {
            Log("Processing List: " + list.ParentWeb.Site.MakeFullUrl(list.DefaultViewUrl));

            foreach (SPListItem item in list.Items)
            {
                try
                {
                    if (item.File != null && Utilities.IsCheckedOut(item) && !Utilities.IsCheckedOutByCurrentUser(item))
                    {
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    Log("WARNING: Unable to test item " + item.ID + ": " + ex.Message);
                    continue;
                }
                bool wasCheckedOut = true;
                bool modified = false;

                foreach (SPField field in list.Fields)
                {
                    try
                    {
                        if (item[field.Id] == null || field.ReadOnlyField)
                            continue;
                    }
                    catch (Exception ex)
                    {
                        if (field.InternalName != "Facilities")
                            Log("WARNING: Unable to read field " + field.Id + " (" + field.InternalName + "): " + ex.Message);
                        continue;
                    }
                    //if ((list.Title.ToLowerInvariant() == "upgrade area url mapping" ||
                    //     list.DefaultViewUrl.ToLowerInvariant().IndexOf(UPGRADE_AREA_URL_LIST) >= 0)
                    //    && field.Title == "V2ServerRelativeUrl")
                    //    continue; // We don't want to change this url because then external links will break.

                    Type fieldType = item[field.Id].GetType();

                    if (fieldType != typeof(string))
                        continue; // We're only going to work with strings.

                    string fieldName = field.Title.ToLowerInvariant();
                    if (UseInternalFieldName)
                        fieldName = field.InternalName.ToLowerInvariant();

                    if (FieldName == null || FieldName.Count == 0 || FieldName.Contains(fieldName))
                    {
                        string result = (string)item[field.Id];
                        bool fieldMatchFound = false;
                        foreach (SearchReplaceData searchReplaceData in SearchStrings)
                        {
                            bool isMatch = searchReplaceData.SearchString.IsMatch(result);

                            if (!isMatch)
                                continue;

                            fieldMatchFound = true;
                            result = searchReplaceData.SearchString.Replace(result, searchReplaceData.ReplaceString);
                        }
                        if (result == (string)item[field.Id])
                            fieldMatchFound = false;

                        if (fieldMatchFound)
                        {
                            Log(string.Format("Match found: List={0}, Field={1}, Replacement={2} => {3}", SPUrlUtility.CombineUrl(list.ParentWeb.Url, item.Url),
                                              field.Title, item[field.Id], result));
                            if (!Test)
                            {
                                if (item.File != null && item.File.CheckOutStatus == SPFile.SPCheckOutStatus.None)
                                {
                                    //item.File.CheckOut();
                                    wasCheckedOut = false;
                                }
                                try
                                {
                                    item[field.Id] = result;
                                }
                                catch (Exception ex)
                                {
                                    string msg = string.Format("\r\nWARNING: Unable to set field value.\r\nList={0}, Field={1}, Replacement={2} => {3}",
                                        item.Url, field.Title, item[field.Id], result);
                                    if (LogFile == null)
                                        msg += string.Format("\r\n{0}\r\n{1}", ex.Message, ex.StackTrace);

                                    if (Quiet)
                                        Console.WriteLine(msg + "\r\nSee log file for me details.");

                                    if (LogFile != null && Quiet)
                                        msg += string.Format("\r\n{0}\r\n{1}", ex.Message, ex.StackTrace);

                                    Log(msg);
                                }
                                modified = true;
                            }
                        }
                    }
                }
                if (!Test)
                {
                    if (modified)
                    {
                        try
                        {
                            try
                            {
                                Log("Progress: Attempting System Update to save changes...");
                                item.SystemUpdate();
                            }
                            catch (Exception)
                            {
                                Log("Progress: System Update Failed, attempting check out and update to save changes...");
                                if (!wasCheckedOut)
                                {
                                    item.File.CheckOut();
                                    item.Update();
                                }
                                else
                                    throw;
                            }
                        }
                        catch (Exception ex)
                        {
                            string msg = string.Format("\r\nWARNING: Unable to set item values.\r\nList={0}", item.Url);

                            if (LogFile == null)
                                msg += string.Format("\r\n{0}\r\n{1}", ex.Message, ex.StackTrace);

                            if (Quiet)
                                Console.WriteLine(msg + "\r\nSee log file for me details.");

                            if (LogFile != null)
                                msg += string.Format("\r\n{0}\r\n{1}", ex.Message, ex.StackTrace);

                            Log(msg);
                        }
                    }

                    if (Utilities.IsCheckedOut(item))
                    {
                        if (modified && item.File != null)
                            item.File.CheckIn(
                                "Checking in changes to list item due to automated search and replace", SPCheckinType.MajorCheckIn);
                        else if (!wasCheckedOut && item.File != null)
                            item.File.UndoCheckOut();
                    }

                    if (modified && Publish && !wasCheckedOut)
                    {
                        Common.Lists.PublishItems pi = new Common.Lists.PublishItems();
                        pi.PublishListItem(item, list, Test, "Replace-SPFieldValues", "Replaced values via Replace-SPFieldValues", null);
                    }
                }
            }
            Log("Finished Processing List: " + list.ParentWeb.Site.MakeFullUrl(list.DefaultViewUrl) + "\r\n");
        }

        #endregion

        #region Utility Methods

        /// <summary>
        /// Parses the input file.
        /// </summary>
        /// <param name="settings">The settings.</param>
        /// <param name="inputFile">The input file.</param>
        /// <param name="delimiter">The delimiter.</param>
        /// <param name="isXml">if set to <c>true</c> [is XML].</param>
        public void ParseInputFile(string inputFile, string delimiter, bool isXml)
        {
            if (!isXml)
            {
                if (string.IsNullOrEmpty(inputFile))
                    throw new ArgumentNullException("inputFile");

                string[] inputLines = File.ReadAllLines(inputFile);

                if (inputLines.Length > 0)
                {
                    foreach (string line in inputLines)
                    {
                        string[] arguments = line.Split(delimiter.ToCharArray(), 2);
                        if (arguments.Length == 2)
                        {
                            Log(string.Format("Processing inputfile line: SearchString=[{0}], ReplaceString=[{1}]",
                                              arguments[0], arguments[1]));
                            SearchStrings.Add(new SearchReplaceData(arguments[0], arguments[1]));
                        }
                        else
                            throw new ArgumentException(
                                string.Format("The search and replace string contains too many or too few delimiters: {0}", line));
                    }
                }
            }
            else
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(inputFile);
                ParseInputFile(xmlDoc);
            }
        }

        public void ParseInputFile(XmlDocument xmlDoc)
        {
            foreach (XmlElement replacement in xmlDoc.SelectNodes("//Replacement"))
            {
                XmlElement search = (XmlElement)replacement.SelectSingleNode("./SearchString");
                XmlElement replace = (XmlElement)replacement.SelectSingleNode("./ReplaceString");

                if (search == null)
                    throw new SPException(
                        string.Format("The Replacement node did not contain a SearchString child node: {0}",
                                      replacement.InnerXml));

                if (replace == null)
                    throw new SPException(
                        string.Format("The Replacement node did not contain a ReplaceString child node: {0}",
                                      replacement.InnerXml));

                Log(string.Format("Processing inputfile node: SearchString=[{0}], ReplaceString=[{1}]",
                                  search.InnerText, replace.InnerText));

                SearchStrings.Add(new SearchReplaceData(search.InnerText, replace.InnerText));
            }
        }

        /// <summary>
        /// Logs the specified message using the provided settings.
        /// </summary>
        /// <param name="message">The message.</param>
        private void Log(string message)
        {
            try
            {
                if (!Quiet)
                    Logger.Write(message);
            }
            catch 
            { }
            try
            {
                if (LogFile != null)
                {
                    File.AppendAllText(LogFile, DateTime.Now + ": " + message + "\r\n");
                }
            }
            catch { }
        }

        #endregion

        #region Internal Classes

        public class SearchReplaceData
        {
            public Regex SearchString;
            public string ReplaceString;

            public SearchReplaceData(string searchString, string replaceString)
            {
                SearchString = new Regex(searchString);
                ReplaceString = replaceString;
            }
        }
        #endregion

    }
}
