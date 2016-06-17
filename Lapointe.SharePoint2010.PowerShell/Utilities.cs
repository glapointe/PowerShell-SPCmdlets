using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebPartPages;
using WebPart = System.Web.UI.WebControls.WebParts.WebPart;
using System.Management.Automation;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Taxonomy;
using Microsoft.SharePoint.WebControls;

namespace Lapointe.SharePoint.PowerShell
{
    public class Utilities
    {
        internal const string ENCODED_SPACE = "_x0020_";

        /// <summary>
        /// Gets all bindings.
        /// </summary>
        /// <value>All bindings.</value>
        internal static BindingFlags AllBindings
        {
            get
            {
                return BindingFlags.CreateInstance |
                BindingFlags.FlattenHierarchy |
                BindingFlags.GetField |
                BindingFlags.GetProperty |
                BindingFlags.IgnoreCase |
                BindingFlags.Instance |
                BindingFlags.InvokeMethod |
                BindingFlags.NonPublic |
                BindingFlags.Public |
                BindingFlags.SetField |
                BindingFlags.SetProperty |
                BindingFlags.Static;
            }
        }
        public static string GetGenericSetupPath(string subDir)
        {
#if SP2010
            return GetGenericSetupPath(subDir, 14);
#else
            return GetGenericSetupPath(subDir, 15);
#endif
        }
        public static string GetGenericSetupPath(string subDir, int desiredVersion)
        {
#if SP2010
            return SPUtility.GetGenericSetupPath(subDir);
#else
            return SPUtility.GetVersionedGenericSetupPath(subDir, desiredVersion);
#endif
        }

        public static SPServiceContext GetUserProfileServiceContext(string appName, SPSiteSubscriptionIdentifier subId)
        {
            return GetServiceContext("User Profile Service", appName, subId);
        }
        public static SPServiceApplication GetUserProfileServiceApplication(string appName)
        {
            return GetServiceApplication("User Profile Service", appName);
        }
        public static SPServiceApplication GetUserProfileServiceApplication(SPServiceContext context)
        {
            return GetServiceApplication("User Profile Service", context);
        }

        public static SPSiteSubscriptionIdentifier GetSiteSubscriptionId(Guid id)
        {
            if (id == Guid.Empty)
                return SPSiteSubscriptionIdentifier.Default;
            return new SPSiteSubscriptionIdentifier(id);
        }
        public static SPServiceApplication GetServiceApplication(string typeName, string appName)
        {
            foreach (SPService svc in SPFarm.Local.Services)
            {
                if (string.IsNullOrEmpty(svc.TypeName))
                    continue;

                if (svc.TypeName.ToLowerInvariant() == typeName.ToLowerInvariant())
                {

                    foreach (SPServiceApplication svcApp in svc.Applications)
                    {
                        if (string.IsNullOrEmpty(svcApp.DisplayName))
                            continue;

                        if (svcApp.DisplayName.ToLowerInvariant() == appName.ToLowerInvariant())
                        {
                            return svcApp;
                        }
                    }
                    break;
                }
            }
            return null;
        }
        public static SPServiceApplication GetServiceApplication(string typeName, Guid appId)
        {
            foreach (SPService svc in SPFarm.Local.Services)
            {
                if (string.IsNullOrEmpty(svc.TypeName))
                    continue;

                if (svc.TypeName.ToLowerInvariant() == typeName.ToLowerInvariant())
                {

                    foreach (SPServiceApplication svcApp in svc.Applications)
                    {
                        if (string.IsNullOrEmpty(svcApp.DisplayName))
                            continue;

                        if (svcApp.Id == appId)
                        {
                            return svcApp;
                        }
                    }
                    break;
                }
            }
            return null;
        }

        public static SPServiceApplication GetServiceApplication(string typeName, SPServiceContext context)
        {
            foreach (SPService svc in SPFarm.Local.Services)
            {
                if (string.IsNullOrEmpty(svc.TypeName))
                    continue;

                if (svc.TypeName.ToLowerInvariant() == typeName.ToLowerInvariant())
                {

                    foreach (SPServiceApplication svcApp in svc.Applications)
                    {
                        if (string.IsNullOrEmpty(svcApp.DisplayName))
                            continue;

                        // Test using the default proxy group and default site subscription - the most common
                        SPServiceContext testContext = SPServiceContext.GetContext(svcApp.ServiceApplicationProxyGroup, SPSiteSubscriptionIdentifier.Default);
                        if (testContext == context)
                        {
                            return svcApp;
                        }
                        // Test with each proxy group.
                        foreach (SPServiceApplicationProxyGroup proxyGroup in new SPServiceApplicationProxyGroupCollection(SPFarm.Local))
                        {
                            testContext = SPServiceContext.GetContext(proxyGroup, SPSiteSubscriptionIdentifier.Default);
                            if (testContext == context)
                            {
                                return svcApp;
                            }
                            // Test with each site subscription
                            foreach (SPSiteSubscription siteSub in SPFarm.Local.SiteSubscriptions)
                            {
                                testContext = SPServiceContext.GetContext(proxyGroup, siteSub.Id);
                                if (testContext == context)
                                {
                                    return svcApp;
                                }
                            }

                        }
                    }
                    break;
                }
            }
            return null;
        }

 

        public static SPServiceApplicationProxyGroup GetProxyGroup(SPWebApplication webApplication)
        {
            SPServiceApplicationProxyGroup serviceApplicationProxyGroup = null;
            if (null != webApplication)
            {
                serviceApplicationProxyGroup = webApplication.ServiceApplicationProxyGroup;
                if (null == serviceApplicationProxyGroup)
                {
                    serviceApplicationProxyGroup = SPFarm.Local.GetChild<SPServiceApplicationProxyGroup>("");
                    if (null == serviceApplicationProxyGroup)
                    {
                        serviceApplicationProxyGroup = new SPServiceApplicationProxyGroup("", SPFarm.Local);
                    }
                }
            }
            return serviceApplicationProxyGroup;
        }

 

        public static SPServiceContext GetServiceContext(SPServiceApplication svcApp, SPSiteSubscriptionIdentifier subId)
        {
            if (svcApp != null)
                return SPServiceContext.GetContext(svcApp.ServiceApplicationProxyGroup, subId);

            return null;
        }

        public static SPServiceContext GetServiceContext(string typeName, string appName, SPSiteSubscriptionIdentifier subId)
        {
            SPServiceApplication svcApp = GetServiceApplication(typeName, appName);
            if (svcApp != null)
                return SPServiceContext.GetContext(svcApp.ServiceApplicationProxyGroup, subId);
            
            return null;
        }

        internal static bool ValidateEmail(string email)
        {
            Regex regex = new Regex(@"^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$", RegexOptions.IgnoreCase);
            return regex.Match(email).Success;
        }


        private class SPPrefixComparer : IComparer<string>
        {
            // Methods
            public int Compare(string a, string b)
            {
                if (b.Length > a.Length)
                {
                    return 1;
                }
                return -1;
            }

            public bool Equals(string a, string b)
            {
                return false;
            }

            public int GetHashCode(string s)
            {
                return s.Length;
            }
        }



        private static SPPrefix[] SortPrefixes(SPPrefixCollection prefixes)
        {
            int count = prefixes.Count;
            SortedList<string, SPPrefix> list = new SortedList<string, SPPrefix>(count, new SPPrefixComparer());
            SPPrefix[] array = new SPPrefix[count];
            foreach (SPPrefix prefix in prefixes)
            {
                list.Add(prefix.Name, prefix);
            }
            list.Values.CopyTo(array, 0);
            return array;
        }
 

 

        internal static string FindSiteRoot(SPPrefixCollection prefixes, string serverRelativeRequestPath)
        {
            SPPrefix[] sortedPrefixes = SortPrefixes(prefixes);
            string str = null;
            string strMain = serverRelativeRequestPath.TrimStart(new char[] { '/' });
            foreach (SPPrefix prefix in sortedPrefixes)
            {
                int length;
                int num2;
                if (StsStartsWith(strMain, prefix.Name))
                {
                    length = prefix.Name.Length;
                    num2 = length + 1;
                    

                    if (prefix.PrefixType == SPPrefixType.ExplicitInclusion)
                    {
                        if (((serverRelativeRequestPath.Length < (num2 + 1)) || (serverRelativeRequestPath[num2] != '/')) && (length != 0))
                        {
                            strMain = strMain.TrimEnd(new char[] { '/' });
                            if (!string.Equals(strMain, prefix.Name, StringComparison.CurrentCultureIgnoreCase))
                            {
                                continue;
                            }
                        }
                        if (serverRelativeRequestPath.Length > length)
                        {
                            str = serverRelativeRequestPath.Substring(0, length + 1).Trim();
                            if (str.Length > 1)
                            {
                                str = str.TrimEnd(new char[] { '/' });
                            }
                        }
                        return str;
                    }
                    else if (prefix.PrefixType == SPPrefixType.WildcardInclusion)
                    {
                        if (((serverRelativeRequestPath.Length > (num2 + 1)) && (serverRelativeRequestPath[num2] == '/')) || (string.IsNullOrEmpty(prefix.Name) && (serverRelativeRequestPath.Length != 1)))
                        {
                            int index = serverRelativeRequestPath.IndexOf('/', num2 + 1);
                            if (index < 0)
                            {
                                return serverRelativeRequestPath;
                            }
                            return serverRelativeRequestPath.Substring(0, index);
                        }
                    }
                    else
                        continue;
                }
            }
            return null;
        }



        /// <summary>
        /// Gets the SPRequest object.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <returns></returns>
        internal static object GetSPRequestObject(SPWeb web)
        {
            return GetPropertyValue(web, "Request");
        }

        /// <summary>
        /// Gets the property value.
        /// </summary>
        /// <param name="o">The object whose property is to be retrieved.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <returns></returns>
        internal static object GetPropertyValue(object o, string propertyName)
        {
            return o.GetType().GetProperty(propertyName, AllBindings).GetValue(o, null);
        }

        /// <summary>
        /// Sets the property value.
        /// </summary>
        /// <param name="o">The object whose property is to be set.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="value">The value to set the property to.</param>
        internal static void SetPropertyValue(object o, string propertyName, object value)
        {
            try
            {
                o.GetType().GetProperty(propertyName, AllBindings).SetValue(o, value, null);
            }
            catch (AmbiguousMatchException)
            {
                Type t = o.GetType();
                while (true)
                {
                    try
                    {
                        t.GetProperty(propertyName, AllBindings | BindingFlags.DeclaredOnly).SetValue(o, value, null);
                        break;
                    }
                    catch (NullReferenceException)
                    {
                        if (t.BaseType == null)
                            return;

                        t = t.BaseType;
                    }
                }
            }
        }

        internal static void SetPropertyValue(object o, Type type, string propertyName, object value)
        {
            try
            {
                type.GetProperty(propertyName, AllBindings).SetValue(o, value, null);
            }
            catch (AmbiguousMatchException)
            {
                Type t = type;
                while (true)
                {
                    try
                    {
                        t.GetProperty(propertyName, AllBindings | BindingFlags.DeclaredOnly).SetValue(o, value, null);
                        break;
                    }
                    catch (NullReferenceException)
                    {
                        if (t.BaseType == null)
                            return;

                        t = t.BaseType;
                    }
                }
            }
        }

        internal static object GetFieldValue(object o, string fieldName)
        {
            return o.GetType().GetField(fieldName, AllBindings).GetValue(o);
        }

        internal static void SetFieldValue(object o, Type type, string fieldName, object value)
        {
            type.GetField(fieldName, AllBindings).SetValue(o, value);
        }

        /// <summary>
        /// Executes the method.
        /// </summary>
        /// <param name="objectType">The type.</param>
        /// <param name="methodName">Name of the method.</param>
        /// <param name="parameterTypes">The parameter types.</param>
        /// <param name="parameterValues">The parameter values.</param>
        /// <returns></returns>
        internal static object ExecuteMethod(Type objectType, string methodName, Type[] parameterTypes, object[] parameterValues)
        {
            return ExecuteMethod(objectType, null, methodName, parameterTypes, parameterValues);
        }

        /// <summary>
        /// Executes the method.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="methodName">Name of the method.</param>
        /// <param name="parameterTypes">The parameter types.</param>
        /// <param name="parameterValues">The parameter values.</param>
        /// <returns></returns>
        internal static object ExecuteMethod(object obj, string methodName, Type[] parameterTypes, object[] parameterValues)
        {
            return ExecuteMethod(obj.GetType(), obj, methodName, parameterTypes, parameterValues);
        }

        /// <summary>
        /// Executes the method.
        /// </summary>
        /// <param name="objectType">The type.</param>
        /// <param name="obj">The obj.</param>
        /// <param name="methodName">Name of the method.</param>
        /// <param name="parameterTypes">The parameter types.</param>
        /// <param name="parameterValues">The parameter values.</param>
        /// <returns></returns>
        internal static object ExecuteMethod(Type objectType, object obj, string methodName, Type[] parameterTypes, object[] parameterValues)
        {
            MethodInfo methodInfo = objectType.GetMethod(methodName, AllBindings, null, parameterTypes, null);
            try
            {
                return methodInfo.Invoke(obj, parameterValues);
            }
            catch (TargetInvocationException ex)
            {
                // Get and throw the real exception.
                throw ex.InnerException;
            }
        }

        /// <summary>
        /// Processes the RPC results.
        /// </summary>
        /// <param name="results">The results.</param>
        internal static void ProcessRpcResults(string results)
        {
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(results);
            XmlElement errorText = (XmlElement)xml.SelectSingleNode("//ErrorText");
            if (errorText != null)
            {
                throw new SPException(errorText.InnerText + "(" + xml.DocumentElement.GetAttribute("Code") + ")");
            }
        }

        /// <summary>
        /// Gets the web part by id.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="url">The URL.</param>
        /// <param name="id">The id.</param>
        /// <param name="manager">The web part manager.</param>
        /// <returns></returns>
        internal static WebPart GetWebPartById(SPWeb web, string url, string id, out SPLimitedWebPartManager manager)
        {
            manager = web.GetLimitedWebPartManager(url, PersonalizationScope.Shared);
            WebPart wp = manager.WebParts[id];
            if (wp == null)
            {
                manager.Web.Dispose(); // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                manager.Dispose();
                manager = web.GetLimitedWebPartManager(url, PersonalizationScope.User);
                wp = manager.WebParts[id];
            }
            return wp;
        }

        internal static List<WebPart> GetWebPartsByType(SPWeb web, string url, Type type, out SPLimitedWebPartManager manager)
        {
            manager = web.GetLimitedWebPartManager(url, PersonalizationScope.Shared);
            List<WebPart> foundParts = new List<WebPart>();
            List<WebPart> toDispose = new List<WebPart>();
            try
            {
                foreach (WebPart wp in manager.WebParts)
                {
                    if (wp.GetType() == type)
                    {
                        foundParts.Add(wp);
                    }
                    else
                    {
                        toDispose.Add(wp);
                    }
                }
            }
            finally
            {
                foreach (WebPart tempWP in toDispose)
                {
                    tempWP.Dispose();
                }
            }
            return foundParts;
        }

        /// <summary>
        /// Gets the web part by title.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="url">The URL.</param>
        /// <param name="title">The title.</param>
        /// <param name="manager">The web part manager.</param>
        /// <returns></returns>
        internal static WebPart GetWebPartByTitle(SPWeb web, string url, string title, out SPLimitedWebPartManager manager)
        {
            manager = web.GetLimitedWebPartManager(url, PersonalizationScope.Shared);
            List<WebPart> foundParts = new List<WebPart>();
            WebPart wp = null;
            try
            {
                foreach (WebPart tempWP in manager.WebParts)
                {
                    if (tempWP.DisplayTitle.ToLowerInvariant() == title.ToLowerInvariant() ||
                        tempWP.Title.ToLowerInvariant() == title.ToLowerInvariant())
                    {
                        foundParts.Add(tempWP);
                        wp = tempWP;
                    }
                }
                if (foundParts.Count == 0)
                {
                    manager.Web.Dispose();
                    // manager.Dispose() does not dispose of the SPWeb object and results in a memory leak.
                    manager.Dispose();
                    manager = web.GetLimitedWebPartManager(url, PersonalizationScope.User);
                    foreach (WebPart tempWP in manager.WebParts)
                    {
                        if (tempWP.DisplayTitle.ToLowerInvariant() == title.ToLowerInvariant() ||
                            tempWP.Title.ToLowerInvariant() == title.ToLowerInvariant())
                        {
                            foundParts.Add(tempWP);
                            wp = tempWP;
                        }
                    }
                }
                if (foundParts.Count > 1)
                {
                    string msg =
                        "Found more than one web part matching the specified title.  Use the ID instead:\r\n\r\n";
                    XmlDocument xmlDoc = new XmlDocument();
                    string tempXml = null;
                    foreach (WebPart tempWP in foundParts)
                    {
                        tempXml += Common.WebParts.EnumPageWebParts.GetWebPartDetailsMinimal(tempWP, manager);
                    }
                    xmlDoc.LoadXml("<MatchingWebParts>" + tempXml + "</MatchingWebParts>");
                    throw new SPException(msg + GetFormattedXml(xmlDoc));
                }
            }
            finally
            {
                foreach (WebPart tempWP in foundParts)
                {
                    if (wp.ID != tempWP.ID)
                        tempWP.Dispose();
                }
            }
            return wp;
        }

        /// <summary>
        /// Gets the formatted XML.
        /// </summary>
        /// <param name="xmlDoc">The XML doc.</param>
        /// <returns></returns>
        internal static string GetFormattedXml(XmlDocument xmlDoc)
        {
            StringBuilder sb = new StringBuilder();

            XmlTextWriter xmlWriter = new XmlTextWriter(new StringWriter(sb));
            xmlWriter.Formatting = Formatting.Indented;
            xmlDoc.WriteContentTo(xmlWriter);
            xmlWriter.Flush();

            return sb.ToString();
        }

        /// <summary>
        /// The LookupAccountSid function accepts a security identifier (SID) as input. It retrieves the name of the account for this SID and the name of the first domain on which this SID is found.
        /// </summary>
        /// <param name="systemName">A pointer to a null-terminated character string that specifies the target computer. This string can be the name of a remote computer. If this string is NULL, the account name translation begins on the local system. If the name cannot be resolved on the local system, this function will try to resolve the name using domain controllers trusted by the local system. Generally, specify a value for systemName only when the account is in an untrusted domain and the name of a computer in that domain is known.</param>
        /// <param name="sid">The SID to look up in binary format</param>
        /// <param name="name">A pointer to a buffer that receives a null-terminated string that contains the account name that corresponds to the sid parameter.</param>
        /// <param name="nameBuffer">On input, specifies the size, in TCHARs, of the name buffer. If the function fails because the buffer is too small or if nameBuffer is zero, nameBuffer receives the required buffer size, including the terminating null character.</param>
        /// <param name="domainName">A pointer to a buffer that receives a null-terminated string that contains the name of the domain where the account name was found.</param>
        /// <param name="domainNameBuffer">On input, specifies the size, in TCHARs, of the domainName buffer. If the function fails because the buffer is too small or if domainNameBuffer is zero, domainNameBuffer receives the required buffer size, including the terminating null character.</param>
        /// <param name="accountType">A pointer to a variable that receives a SID_NAME_USE value that indicates the type of the account.</param>
        /// <returns></returns>
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool LookupAccountSid(string systemName, byte[] sid, StringBuilder name, ref int nameBuffer, StringBuilder domainName, ref int domainNameBuffer, ref SID_NAME_USE accountType);

        /// <summary>
        /// The SID_NAME_USE enumeration type contains values that specify the type of a security identifier (SID).
        /// Needed for TryGetNT4StyleAccountName and the unmanaged LookupAccountSid method.
        /// </summary>
        public enum SID_NAME_USE
        {
            SidTypeAlias = 4,
            SidTypeComputer = 9,
            SidTypeDeletedAccount = 6,
            SidTypeDomain = 3,
            SidTypeGroup = 2,
            SidTypeInvalid = 7,
            SidTypeUnknown = 8,
            SidTypeUser = 1,
            SidTypeWellKnownGroup = 5
        }

        /// <summary>
        /// Determines whether [is login valid] [the specified STR login name].
        /// </summary>
        /// <param name="strLoginName">Name of the STR login.</param>
        /// <param name="bIsUserAccount">if set to <c>true</c> [b is user account].</param>
        /// <returns>
        /// 	<c>true</c> if [is login valid] [the specified STR login name]; otherwise, <c>false</c>.
        /// </returns>
        internal static bool IsLoginValid(string strLoginName, out bool bIsUserAccount)
        {
            bIsUserAccount = false;
            object request = GetPropertyValue(SPFarm.Local, "Request");

            MethodInfo methodInfo = typeof(SPUtility).GetMethod("IsLoginValid", AllBindings, null, new Type[] { request.GetType(), typeof(string), typeof(bool).MakeByRefType() }, null);
            try
            {
                object[] args = new object[] { request, strLoginName, bIsUserAccount };
                bool isLoginValid = (bool)methodInfo.Invoke(null, args);
                bIsUserAccount = (bool)args[2];
                return isLoginValid;
            }
            catch (TargetInvocationException ex)
            {
                // Get and throw the real exception.
                throw ex.InnerException;
            }

        }

        /// <summary>
        /// Attempts to get the login name formated as "domain\login".  Note that this has been reverse engineered
        /// from the SPUtility.TryGetNT4StyleAccountName method which is marked as internal.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="webApp">The web app.</param>
        /// <returns></returns>
        internal static string TryGetNT4StyleAccountName(string input, SPWebApplication webApp)
        {
            if (string.IsNullOrEmpty(input) || input.IndexOf(':') >= 0 || input.IndexOf('\\') >= 0)
            {
                return input;
            }
            try
            {
                // Use the web application as a check to determine if we are using windows authentication.
                // If we're not then return back the input as the site is most likely setup to use a custom provider
                // so we wouldn't expect to be able to retrieve an NT account name.
                bool isWindowsAuth = false;
                if (webApp != null)
                {
                    foreach (SPIisSettings setting in webApp.IisSettings.Values)
                    {
                        if (setting.AuthenticationMode == AuthenticationMode.Windows)
                        {
                            isWindowsAuth = true;
                            break;
                        }
                    }
                    if (!isWindowsAuth)
                        return input;
                }

                SecurityIdentifier identifier = (SecurityIdentifier)new NTAccount(input).Translate(typeof(SecurityIdentifier));
                byte[] binaryForm = new byte[identifier.BinaryLength];
                identifier.GetBinaryForm(binaryForm, 0);


                StringBuilder domainNameStringBuilder = new StringBuilder(0x100);
                StringBuilder userNameStringBuilder = new StringBuilder(0x100);
                SID_NAME_USE sidTypeInvalid = SID_NAME_USE.SidTypeUnknown;
                int cbDomainName = 0x100;
                int cbName = 0x100;
                if (!LookupAccountSid(null, binaryForm, userNameStringBuilder, ref cbName, domainNameStringBuilder, ref cbDomainName, ref sidTypeInvalid))
                {
                    throw new Win32Exception();
                }
                string domainName = domainNameStringBuilder.ToString();
                string userName = userNameStringBuilder.ToString().ToLowerInvariant();
                int index = userName.IndexOf('@');
                if (index > 0)
                {
                    userName = userName.Substring(0, index);
                }
                return (domainName + '\\' + userName);

            }
            catch (Exception)
            {
                return input;
            }
        }


        /// <summary>
        /// Gets the field schema.
        /// </summary>
        /// <param name="field">The field.</param>
        /// <param name="featureSafe">if set to <c>true</c> [feature safe].</param>
        /// <param name="removeEncodedSpaces">if set to <c>true</c> [remove encoded spaces].</param>
        /// <returns></returns>
        public static string GetFieldSchema(SPField field, bool featureSafe, bool removeEncodedSpaces)
        {
            string schema = field.SchemaXml;
            if (field.InternalName.Contains(ENCODED_SPACE) && removeEncodedSpaces)
            {
                schema = schema.Replace(string.Format("Name=\"{0}\"", field.InternalName),
                                        string.Format("Name=\"{0}\"", field.InternalName.Replace(ENCODED_SPACE, string.Empty)));
            }
            if (featureSafe)
            {
                XmlDocument schemaDoc = new XmlDocument();
                schemaDoc.LoadXml(schema);
                XmlElement fieldElement = schemaDoc.DocumentElement;

                // Remove the Version attribute
                if (fieldElement.HasAttribute("Version"))
                    fieldElement.RemoveAttribute("Version");

                // Remove the Aggregation attribute
                if (fieldElement.HasAttribute("Aggregation"))
                    fieldElement.RemoveAttribute("Aggregation");

                // Remove the Customization attribute
                if (fieldElement.HasAttribute("Customization"))
                    fieldElement.RemoveAttribute("Customization");

                // Fix the UserSelectionMode attribute
                if (fieldElement.HasAttribute("UserSelectionMode"))
                {
                    if (fieldElement.GetAttribute("UserSelectionMode") == "PeopleAndGroups")
                        fieldElement.SetAttribute("UserSelectionMode", "1");
                    else if (fieldElement.GetAttribute("UserSelectionMode") == "PeopleOnly")
                        fieldElement.SetAttribute("UserSelectionMode", "0");
                }
                schema = schemaDoc.OuterXml;
            }
            return schema;
        }

        /// <summary>
        /// Gets the field.
        /// </summary>
        /// <param name="listViewUrl">The list view URL.</param>
        /// <param name="fieldName">Name of the field.</param>
        /// <param name="fieldTitle">The field title.</param>
        /// <param name="useFieldName">if set to <c>true</c> [use field name].</param>
        /// <param name="useFieldTitle">if set to <c>true</c> [use field title].</param>
        /// <returns></returns>
        internal static SPField GetField(string listViewUrl, string fieldName, string fieldTitle, bool useFieldName, bool useFieldTitle)
        {
            SPList list = GetListFromViewUrl(listViewUrl);

            if (list == null)
            {
                throw new Exception("List not found.");
            }

            SPField field = null;
            try
            {
                // It's possible to have more than one field in the fields collection with the same display name.  The Fields collection
                // will merely return back the first item it finds with a name matching the one specified so it's really a rather useless
                // way of retrieving a field as it's extremely misleading and could lead to someone inadvertantly messing things up.
                // I provide the ability to use the display name for convienence but don't rely on it for anything.
                if (useFieldTitle)
                    field = list.Fields[fieldTitle];
            }
            catch (ArgumentException)
            {
            }

            if (field != null || useFieldName)
            {
                int count = 0;
                string foundFields = string.Empty;
                // If the user specified the display name we need to make sure that only one field exists matching that display name.
                // If they specified the internal name then we need to loop until we find a match.
                foreach (SPField temp in list.Fields)
                {
                    if (useFieldName && (temp.InternalName.ToLowerInvariant() == fieldName.ToLowerInvariant() || temp.Id.ToString().Replace("{", "").Replace("}", "").ToLowerInvariant() == fieldName.ToLowerInvariant().Replace("{", "").Replace("}", "")))
                    {
                        field = temp;
                        break;
                    }
                    else if (useFieldTitle && temp.Title == fieldTitle)
                    {
                        count++;
                        foundFields += "\t" + temp.Title + " = " + temp.InternalName + "\r\n";
                    }
                }
                if (useFieldTitle && count > 1)
                {
                    throw new Exception("More than one field was found matching the display name specified:\r\n\r\n\tDisplay Name = Internal Name\r\n\t----------------------------\r\n" +
                                        foundFields +
                                        "\r\nUse \"-fieldinternalname\" to delete based on the internal name of the field.");
                }
            }

            if (field == null)
                throw new Exception("Field not found.");
            return field;
        }




        /// <summary>
        /// Splits the path file.
        /// </summary>
        /// <param name="fullPathFile">The full path file.</param>
        /// <param name="path">The path.</param>
        /// <param name="filename">The filename.</param>
        internal static void SplitPathFile(string fullPathFile, out string path, out string filename)
        {
            FileInfo info = new FileInfo(fullPathFile);
            path = info.Directory.FullName;
            filename = info.Name;
        }


        /// <summary>
        /// Gets the list from the view URL.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns></returns>
        internal static SPList GetListFromViewUrl(string url)
        {
            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb())
            {
                return GetListFromViewUrl(web, url);
            }
        }

        /// <summary>
        /// Gets the list from the view URL.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="url">The URL.</param>
        /// <returns></returns>
        internal static SPList GetListFromViewUrl(SPWeb web, string url)
        {
            url = SPEncode.UrlDecodeAsUrl(url);

            SPList list = null;
            if (url.ToLowerInvariant().EndsWith(".aspx"))
            {
                try
                {
                    list = web.GetListFromWebPartPageUrl(url);
                }
                catch (SPException)
                {
                    // This block is redundant - if the above fails this should also fail - I left it here for legacy reasons only.
                    foreach (SPList tempList in web.Lists)
                    {
                        foreach (SPView view in tempList.Views)
                        {
                            if (url.ToLower() == SPEncode.UrlDecodeAsUrl(web.Site.MakeFullUrl(view.ServerRelativeUrl)).ToLower())
                            {
                                list = tempList;
                                break;
                            }
                        }
                        if (list != null)
                            break;
                    }
                }
            }
            else
            {
                try
                {
                    SPFolder folder = web.GetFolder(url);
                    if (folder == null)
                        return null;
                    list = web.Lists[folder.ParentListId];
                }
                catch (Exception)
                {
                    list = null;
                }
            }
            return list;
        }

        /// <summary>
        /// Gets the list by URL.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="listUrlName">Name of the list URL.</param>
        /// <param name="isDocLib">if set to <c>true</c> [is doc lib].</param>
        /// <returns></returns>
        internal static SPList GetListByUrl(SPWeb web, string listUrlName, bool isDocLib)
        {
            string strUrl;
            if (!web.IsRootWeb)
            {
                strUrl = web.ServerRelativeUrl;
            }
            else
            {
                strUrl = string.Empty;
            }
            if (!isDocLib)
            {
                strUrl = strUrl + "/Lists/" + listUrlName;
            }
            else
            {
                strUrl = strUrl + "/" + listUrlName;
            }
            return web.GetList(strUrl);
        }


        /// <summary>
        /// Runs the operation.
        /// </summary>
        /// <param name="args">The arguments to pass into STSADM.</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <returns></returns>
        internal static int RunStsAdmOperation(string args, bool quiet)
        {
            string stsadmPath = Path.Combine(Utilities.GetGenericSetupPath("BIN"), "stsadm.exe");

            return RunCommand(stsadmPath, args, quiet);
        }

        /// <summary>
        /// Runs the command.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="args">The args.</param>
        /// <param name="quiet">if set to <c>true</c> [quiet].</param>
        /// <returns></returns>
        internal static int RunCommand(string fileName, string args, bool quiet)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = true;
            startInfo.FileName = fileName;
            startInfo.Arguments = args;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardInput = true;
            startInfo.UseShellExecute = false;

            Process proc = new Process();
            try
            {
                proc.ErrorDataReceived += new DataReceivedEventHandler(Process_ErrorDataReceived);
                if (!quiet)
                    proc.OutputDataReceived += new DataReceivedEventHandler(Process_OutputDataReceived);

                proc.StartInfo = startInfo;
                proc.Start();
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();
                proc.WaitForExit();

                return proc.ExitCode;
            }
            finally
            {
                proc.Close();
            }
        }

        /// <summary>
        /// Handles the ErrorDataReceived event of the Process control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Diagnostics.DataReceivedEventArgs"/> instance containing the event data.</param>
        static void Process_ErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            Console.WriteLine(e.Data);
        }

        /// <summary>
        /// Handles the OutputDataReceived event of the Process control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Diagnostics.DataReceivedEventArgs"/> instance containing the event data.</param>
        static void Process_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (e.Data == "Operation completed successfully.")
                return;

            Console.WriteLine(e.Data);
        }

        /// <summary>
        /// Determines whether a string starts with another string.  This is taken from Microsoft.SharePoint.Utilities.SPUtility
        /// and is needed by ConvertToServiceRelUrl.
        /// </summary>
        /// <param name="strMain">The STR main.</param>
        /// <param name="strBegining">The STR begining.</param>
        /// <returns></returns>
        internal static bool StsStartsWith(string strMain, string strBegining)
        {
            return CultureInfo.InvariantCulture.CompareInfo.IsPrefix(strMain, strBegining, CompareOptions.IgnoreCase);
        }


        /// <summary>
        /// compare strings.
        /// </summary>
        /// <param name="str1">The STR1.</param>
        /// <param name="str2">The STR2.</param>
        /// <returns></returns>
        internal static bool StsCompareStrings(string str1, string str2)
        {
            CompareInfo compareInfo = CultureInfo.InvariantCulture.CompareInfo;
            return (0 == compareInfo.Compare(str1, str2, CompareOptions.IgnoreCase));
        }



        /// <summary>
        /// Splits the URL.
        /// </summary>
        /// <param name="fullOrRelativeUri">The full or relative URI.</param>
        /// <param name="dirName">Name of the dir.</param>
        /// <param name="leafName">Name of the leaf.</param>
        internal static void SplitUrl(string fullOrRelativeUri, out string dirName, out string leafName)
        {
            if (fullOrRelativeUri != null)
            {
                if ((fullOrRelativeUri.Length > 0) && ('/' == fullOrRelativeUri[0]))
                {
                    fullOrRelativeUri = fullOrRelativeUri.Substring(1);
                }
            }
            else
            {
                dirName = string.Empty;
                leafName = string.Empty;
                return;
            }
            int length = fullOrRelativeUri.LastIndexOf('/');
            if (-1 != length)
            {
                dirName = fullOrRelativeUri.Substring(0, length);
                leafName = fullOrRelativeUri.Substring(length + 1);
            }
            else
            {
                dirName = string.Empty;
                if (fullOrRelativeUri.Length > 0)
                {
                    if ('/' == fullOrRelativeUri[0])
                    {
                    }
                    leafName = fullOrRelativeUri.Substring(1);
                }
                else
                {
                    leafName = string.Empty;
                }
            }
        }



        /// <summary>
        /// Converts to service rel URL.  This is taken from Microsoft.SharePoint.Utilities.SPUtility.
        /// </summary>
        /// <param name="strUrl">The STR URL.</param>
        /// <param name="strBaseUrl">The STR base URL.</param>
        /// <returns></returns>
        internal static string ConvertToServiceRelUrl(string strUrl, string strBaseUrl)
        {
            if (((strBaseUrl == null) || !StsStartsWith(strBaseUrl, "/")) || ((strUrl == null) || !StsStartsWith(strUrl, "/")))
            {
                throw new ArgumentException();
            }
            if ((strUrl.Length > 1) && (strUrl[strUrl.Length - 1] == '/'))
            {
                strUrl = strUrl.Substring(0, strUrl.Length - 1);
            }
            if ((strBaseUrl.Length > 1) && (strBaseUrl[strBaseUrl.Length - 1] == '/'))
            {
                strBaseUrl = strBaseUrl.Substring(0, strBaseUrl.Length - 1);
            }
            if (!StsStartsWith(strUrl, strBaseUrl))
            {
                throw new ArgumentException();
            }
            if (strBaseUrl != "/")
            {
                if (strUrl.Length != strBaseUrl.Length)
                {
                    return strUrl.Substring(strBaseUrl.Length + 1);
                }
                return "";
            }
            return strUrl.Substring(1);
        }

        /// <summary>
        /// Concats the server relative urls.
        /// </summary>
        /// <param name="firstPart">The first part.</param>
        /// <param name="secondPart">The second part.</param>
        /// <returns></returns>
        internal static string ConcatServerRelativeUrls(string firstPart, string secondPart)
        {
            firstPart = firstPart.TrimEnd('/');
            secondPart = secondPart.TrimStart('/');
            return (firstPart + "/" + secondPart);
        }

        /// <summary>
        /// Gets the server relative URL from full URL.
        /// </summary>
        /// <param name="url">The STR URL.</param>
        /// <returns></returns>
        internal static string GetServerRelUrlFromFullUrl(string url)
        {
            int index = url.IndexOf("//");
            if ((index < 0) || (index == (url.Length - 2)))
            {
                throw new ArgumentException();
            }
            int startIndex = url.IndexOf('/', index + 2);
            if (startIndex < 0)
            {
                return "/";
            }
            string str = url.Substring(startIndex);
            if (str.IndexOf("?") >= 0)
                str = str.Substring(0, str.IndexOf("?"));

            if (str.IndexOf(".aspx") > 0)
                str = str.Substring(0, str.LastIndexOf("/"));

            if ((str.Length > 1) && (str[str.Length - 1] == '/'))
            {
                return str.Substring(0, str.Length - 1);
            }
            return str;
        }

        /// <summary>
        /// Gets the checked out user id.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <returns></returns>
        internal static string GetCheckedOutUserId(SPItem item)
        {
            if (item is SPListItem)
            {
                SPListItem item2 = (SPListItem)item;
                if (item2.File != null)
                {
                    if (item2.File.CheckedOutByUser == null)
                        return null;

                    return item2.File.CheckedOutByUser.LoginName;
                }
                if (item2.ParentList.BaseType == SPBaseType.DocumentLibrary)
                {
                    return (string)item["CheckoutUser"];
                }
            }
            return null;
        }

        /// <summary>
        /// Determines whether [is checked out by current user] [the specified item].
        /// </summary>
        /// <param name="item">The item.</param>
        /// <returns>
        /// 	<c>true</c> if [is checked out by current user] [the specified item]; otherwise, <c>false</c>.
        /// </returns>
        internal static bool IsCheckedOutByCurrentUser(SPItem item)
        {
            string user = GetCheckedOutUserId(item);
            if (string.IsNullOrEmpty(user))
                return false;
            if (user.Contains("|"))
                user = user.Split('|')[1];
            return ((Environment.UserDomainName + "\\" + Environment.UserName).ToLowerInvariant() == user.ToLowerInvariant());
        }

        /// <summary>
        /// Determines whether the list item is checked out..
        /// </summary>
        /// <param name="item">The item.</param>
        /// <returns>
        /// 	<c>true</c> if is checked out; otherwise, <c>false</c>.
        /// </returns>
        internal static bool IsCheckedOut(SPItem item)
        {
            return !string.IsNullOrEmpty(GetCheckedOutUserId(item));
        }

        /// <summary>
        /// Ensures the aspx.
        /// </summary>
        /// <param name="relativeUrl">The relative URL.</param>
        /// <param name="allowMasterPage">if set to <c>true</c> [allow master page].</param>
        /// <param name="throwException">if set to <c>true</c> [throw exception].</param>
        /// <returns></returns>
        internal static bool EnsureAspx(string relativeUrl, bool allowMasterPage, bool throwException)
        {
            if (relativeUrl == null)
            {
                if (throwException)
                    throw new ArgumentNullException();
                else
                    return false;
            }
            string extension = Path.GetExtension(relativeUrl);
            if (!string.IsNullOrEmpty(extension))
            {
                extension = extension.Substring(1);
            }
            if (!string.IsNullOrEmpty(extension))
            {
                if (CompareStrings(extension, "aspx"))
                {
                    return true;
                }
                if (allowMasterPage)
                {
                    if (CompareStrings(extension, "master"))
                    {
                        return true;
                    }
                    if (CompareStrings(extension, "ascx"))
                    {
                        return true;
                    }
                }
            }
            if (throwException)
                throw new SPException(string.Format("Url is not a valid aspx or master page: {0}", relativeUrl));
            else
                return false;
        }

        /// <summary>
        /// Compares the strings.
        /// </summary>
        /// <param name="str1">The STR1.</param>
        /// <param name="str2">The STR2.</param>
        /// <returns></returns>
        internal static bool CompareStrings(string str1, string str2)
        {
            CompareInfo compareInfo = CultureInfo.InvariantCulture.CompareInfo;
            return (0 == compareInfo.Compare(str1, str2, CompareOptions.IgnoreCase));
        }



        /// <summary>
        /// Determines whether this instance [can convert to from] the specified converter.
        /// </summary>
        /// <param name="converter">The converter.</param>
        /// <param name="type">The type.</param>
        /// <returns>
        /// 	<c>true</c> if this instance [can convert to from] the specified converter; otherwise, <c>false</c>.
        /// </returns>
        internal static bool CanConvertToFrom(TypeConverter converter, Type type)
        {
            return ((((converter != null) && converter.CanConvertTo(type)) && converter.CanConvertFrom(type)) && !(converter is ReferenceConverter));
        }

        /// <summary>
        /// Formats the exception.
        /// </summary>
        /// <param name="ex">The exception object.</param>
        /// <returns></returns>
        internal static string FormatException(Exception ex)
        {
            if (ex == null) return "";

            string msg = "";

            msg += "\r\nError Type:      " + ex.GetType();
            msg += "\r\nError Message:   " + ex.Message; //Get the error message
            msg += "\r\nError Source:    " + ex.Source;  //Source of the message
            msg += "\r\nError TargetSite:" + ex.TargetSite; //Method where the error occurred
            foreach (DictionaryEntry de in ex.Data)
            {
                msg += "\r\n" + de.Key + ": " + de.Value;
            }
            msg += "\r\nError Stack Trace:\r\n" + ex.StackTrace; //Stack Trace of the error
            if (ex.InnerException != null)
                msg += "\r\nInner Exception:\r\n" + FormatException(ex.InnerException);

            return msg;
        }

        /// <summary>
        /// Creates the secure string.
        /// </summary>
        /// <param name="strIn">The string to convert.</param>
        /// <returns></returns>
        internal static SecureString CreateSecureString(string strIn)
        {
            if (strIn != null)
            {
                SecureString str = new SecureString();
                foreach (char ch in strIn)
                {
                    str.AppendChar(ch);
                }
                str.MakeReadOnly();
                return str;
            }
            return null;
        }


        internal static string ConvertToUnsecureString(System.Security.SecureString value)
        {
            IntPtr unmanagedString = System.Runtime.InteropServices.Marshal.SecureStringToGlobalAllocUnicode(value);
            string unsecureString = System.Runtime.InteropServices.Marshal.PtrToStringUni(unmanagedString);
            System.Runtime.InteropServices.Marshal.ZeroFreeGlobalAllocUnicode(unmanagedString);

            return unsecureString;
        }

        internal static bool ValidateCredentials(PSCredential cred)
        {
            IntPtr userHandle = new IntPtr(0);
            try
            {
                bool returnValue = NativeMethods.LogonUser(
                    cred.UserName.Split('\\')[1],
                    cred.UserName.Split('\\')[0],
                    Utilities.ConvertToUnsecureString(cred.Password),
                    NativeMethods.LOGON32_LOGON_INTERACTIVE,
                    NativeMethods.LOGON32_PROVIDER_DEFAULT,
                    ref userHandle
                    );
                return returnValue;
            }
            finally
            {
                NativeMethods.CloseHandle(userHandle);
            }
        }

        internal static TaxonomySession GetTaxonomySessionFromTermStore(TermStore termStore)
        {
            var taxSession = Utilities.GetPropertyValue(termStore, "Session");
            try
            {
                SPSite site = Utilities.GetPropertyValue(taxSession, "SiteOrNull") as SPSite;
                return new TaxonomySession(site, true);
            }
            catch
            {
                var context = Utilities.GetPropertyValue(taxSession, "Context");
                SPSite site = Utilities.GetPropertyValue(context, "SiteOrNull") as SPSite;
                return new TaxonomySession(site, true);
            }
        }
    }
}
