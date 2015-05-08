using System;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.WebApplications
{
    public class DeleteWebApp : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteWebApp"/> class.
        /// </summary>
        public DeleteWebApp()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator()));
            parameters.Add(new SPParam("deleteiiswebsite", "iis"));
            parameters.Add(new SPParam("deletecontentdb", "db"));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nDeletes a web application.\r\n\r\nParameters:\r\n");
            sb.Append("\t-url <url>\r\n");
            sb.Append("\t[-deleteiiswebsite]\r\n");
            sb.Append("\t[-deletecontentdb]\r\n");

            Init(parameters, sb.ToString());
        }

        #region ISPStsadmCommand Members

        /// <summary>
        /// Gets the help message.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <returns></returns>
        public override string GetHelpMessage(string command)
        {
            return HelpMessage;
        }

        /// <summary>
        /// Runs the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public override int Execute(string command, System.Collections.Specialized.StringDictionary keyValues, out string output)
        {
            output = string.Empty;

            

            string url = Params["url"].Value;
            bool deleteContent = Params["deletecontentdb"].UserTypedIn;
            bool deleteIis = Params["deleteiiswebsite"].UserTypedIn;

            SPWebApplication webApp = SPWebApplication.Lookup(new Uri(url));

            foreach (SPIisSettings iis in webApp.IisSettings.Values)
            {
                try
                {
                    DirectoryInfo path = iis.Path;
                    if (!path.Exists)
                        throw new Exception();
                }
                catch (Exception)
                {
                    iis.Path = new DirectoryInfo("c:\\");
                    webApp.Update();
                }
            }

            if (webApp.IisSettings.Count <= 0 && deleteContent)
            {
                DeleteContentDBs(webApp);
                webApp.Delete();
                return (int)ErrorCodes.NoError;
            }

            int index = 0;
            string[] serverComments = new string[webApp.IisSettings.Count];
            string[] vdirs = new string[webApp.IisSettings.Count];
            foreach (SPIisSettings iisSetting in webApp.IisSettings.Values)
            {
                vdirs[index] = iisSetting.Path.ToString();
                serverComments[index] = iisSetting.ServerComment;
                index++;
            }

            // webApp.Unprovision() does not allow us to specify whether we want the site deleted so we have to use the internal version.
            // SPWebApplication.UnprovisionIisWebSites(deleteIis, serverComments, webApp.ApplicationPool.Name);
            MethodInfo unprovisionIisWebSites = webApp.GetType().GetMethod("UnprovisionIisWebSites",
                                            BindingFlags.NonPublic | BindingFlags.Public |
                                            BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.Static,
                                            null, new Type[] {typeof(bool), typeof(string[]), typeof(string)}, null);

            unprovisionIisWebSites.Invoke(webApp, new object[] { deleteIis, serverComments, webApp.ApplicationPool.Name });


            // SPSolution.RetractSolutions(webApp.Farm, webApp.Id, vdirs, serverComments, true);
            MethodInfo retractSolutions = typeof(SPSolution).GetMethod("RetractSolutions",
                                BindingFlags.NonPublic | BindingFlags.Public |
                                BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.Static,
                                null, new Type[] {typeof(SPFarm), typeof(Guid), typeof(string[]), typeof(string[]), typeof(bool)}, null);

            retractSolutions.Invoke(null, new object[] { webApp.Farm, webApp.Id, vdirs, serverComments, true });

            
            if (SPFarm.Local.TimerService.Instances.Count > 1)
            {
                // SPIisWebsiteUnprovisioningJobDefinition is an internal class so we need to use reflection to set it.

                // SPIisWebsiteUnprovisioningJobDefinition jobDef = new SPIisWebsiteUnprovisioningJobDefinition(deleteIis, serverComments, webApp.ApplicationPool.Name, vdirs, webApp.Id, true);
                Type sPIisWebsiteUnprovisioningJobDefinitionType = Type.GetType("Microsoft.SharePoint.Administration.SPIisWebsiteUnprovisioningJobDefinition, Microsoft.SharePoint, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c");

                ConstructorInfo sPIisWebsiteUnprovisioningJobDefinitionConstructor =
                    sPIisWebsiteUnprovisioningJobDefinitionType.GetConstructor(
                        BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.Public,
                        null,
                        new Type[] {typeof(bool), typeof(string[]), typeof(string), typeof(string[]), typeof(Guid), typeof(bool)}, null);
                object jobDef = sPIisWebsiteUnprovisioningJobDefinitionConstructor.Invoke(new object[] { deleteIis, serverComments, webApp.ApplicationPool.Name, vdirs, webApp.Id, true });


                // jobDef.Schedule = new SPOneTimeSchedule(DateTime.Now);
                PropertyInfo scheduleProp = sPIisWebsiteUnprovisioningJobDefinitionType.GetProperty("Schedule",
                                                                BindingFlags.FlattenHierarchy |
                                                                BindingFlags.NonPublic |
                                                                BindingFlags.Instance |
                                                                BindingFlags.InvokeMethod |
                                                                BindingFlags.GetProperty |
                                                                BindingFlags.Public);

                scheduleProp.SetValue(jobDef, new SPOneTimeSchedule(DateTime.Now), null);

                // jobDef.Update();
                MethodInfo update = sPIisWebsiteUnprovisioningJobDefinitionType.GetMethod("Update",
                                                      BindingFlags.NonPublic |
                                                      BindingFlags.Public |
                                                      BindingFlags.Instance |
                                                      BindingFlags.InvokeMethod |
                                                      BindingFlags.FlattenHierarchy,
                                                      null,
                                                      new Type[] { }, null);


                update.Invoke(jobDef, new object[] { });
            }

            if (deleteContent)
                DeleteContentDBs(webApp);

            webApp.Delete();


            return (int)ErrorCodes.NoError;
        }

        #endregion


        /// <summary>
        /// Deletes the content databases.
        /// </summary>
        /// <param name="webApp">The web app.</param>
        private static void DeleteContentDBs(SPWebApplication webApp)
        {
            foreach (SPContentDatabase db in webApp.ContentDatabases)
            {
                db.Unprovision();
            }
        }

    }
}
