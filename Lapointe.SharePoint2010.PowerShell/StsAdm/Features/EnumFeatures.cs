using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Globalization;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Features
{
    public class EnumFeatures : SPOperation
    {
        public EnumFeatures()
        {
            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("url", "url", true, null, new SPUrlValidator(), "Please specify the site collection"));
            string regex = "^Farm$|^Site$|^Web$|^WebApplication$";
            parameters.Add(new SPParam("scope", "s", false, null, new SPRegexValidator(regex + "|" + regex.ToLower())));
            parameters.Add(new SPParam("showhidden", "h", false, null, null));
            parameters.Add(new SPParam("xml", "x", false, null, null));

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nReturns the list of features and their activation status.\r\n\r\nParameters:");
            sb.Append("\r\n\t-url <site collection url>");
            sb.Append("\r\n\t[-scope <Farm | Site | Web | WebApplication>]");
            sb.Append("\r\n\t[-showhidden]");
            sb.Append("\r\n\t[-xml]");
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
        public override int Execute(string command, StringDictionary keyValues, out string output)
        {
            output = string.Empty;

            string url = Params["url"].Value.TrimEnd('/');
            bool asXml = Params["xml"].UserTypedIn;
            bool showHidden = Params["showhidden"].UserTypedIn;

            SPFeatureScope scope = SPFeatureScope.ScopeInvalid;
            if (Params["scope"].UserTypedIn)
                scope = (SPFeatureScope)Enum.Parse(typeof (SPFeatureScope), Params["scope"].Value, true);

            DataTable dt = BuildFeatureListDataTable(url, scope, showHidden);
            if (asXml)
            {
                DataSet features = new DataSet("Features");
                features.Tables.Add(dt);

                output += features.GetXml();
            }
            else
            {
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    i++;
                    if (scope == SPFeatureScope.ScopeInvalid)
                        output += string.Format("{0}. {1}: {2} ({3} - {4})\r\n", i, row["DisplayName"], row["Title"], row["Scope"], row["Status"]);
                    else
                        output += string.Format("{0}. {1}: {2} ({3})\r\n", i, row["DisplayName"], row["Title"], row["Status"]);
                }
            }
            return (int)ErrorCodes.NoError;
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Builds the feature list data table.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="scope">The scope.</param>
        /// <param name="showHidden">if set to <c>true</c> [show hidden].</param>
        /// <returns></returns>
        private static DataTable BuildFeatureListDataTable(string url, SPFeatureScope scope, bool showHidden)
        {
            string[] dataTableColumns = new string[] { "FeatureId", "Title", "DisplayName", "Description", "Status", "Scope" };
            DataTable dtblFeatureList = new DataTable();
            dtblFeatureList.TableName = "Feature";

            for (int index = 0; index < dataTableColumns.Length; index++ )
            {
                string columnName = dataTableColumns[index];
                DataColumn column = new DataColumn(columnName, typeof(string));
                dtblFeatureList.Columns.Add(column);
            }

            AddFeaturesToTable(ref dtblFeatureList, url, scope, showHidden);
            dtblFeatureList.Locale = CultureInfo.CurrentUICulture;
            dtblFeatureList.DefaultView.Sort = "Title ASC";
            return dtblFeatureList;
        }


        /// <summary>
        /// Adds the features to table.
        /// </summary>
        /// <param name="dtblFeatureList">The data table feature list.</param>
        /// <param name="url">The URL.</param>
        /// <param name="scope">The scope.</param>
        /// <param name="showHidden">if set to <c>true</c> [show hidden].</param>
        private static void AddFeaturesToTable(ref DataTable dtblFeatureList, string url, SPFeatureScope scope, bool showHidden)
        {
            CultureInfo info;
            Dictionary<SPFeatureScope, SPFeatureCollection> activeFeatures = new Dictionary<SPFeatureScope, SPFeatureCollection>();

            // We need to get the SPSite and SPWeb so that we can get the culture info and any active features if a scope was passed in.
            using (SPSite site = new SPSite(url))
            {
                using (SPWeb web = site.AllWebs[Utilities.GetServerRelUrlFromFullUrl(url)])
                {
                    info = new CultureInfo((int)web.Language, false);

                    activeFeatures[SPFeatureScope.Farm] = SPWebService.ContentService.Features;
                    activeFeatures[SPFeatureScope.WebApplication] = site.WebApplication.Features;
                    activeFeatures[SPFeatureScope.Site] = site.Features;
                    activeFeatures[SPFeatureScope.Web] = web.Features;
                }
            }

            foreach (SPFeatureDefinition definition in SPFarm.Local.FeatureDefinitions)
            {
                try
                {
                    Guid featureID = definition.Id;
                    // If the scope is marked as invalid then that's our flag that it wasn't provided so we're going to 
                    // list everything regardless of scope and show those that are active for the Web scope.
                    if (definition.Scope != scope && scope != SPFeatureScope.ScopeInvalid)
                        continue;

                    if (definition.Hidden && !showHidden)
                    {
                        continue;
                    }
                    if (!definition.SupportsLanguage(info))
                    {
                        continue;
                    }

                    bool isActive = false;
                    if (activeFeatures[definition.Scope] != null)
                        isActive = (activeFeatures[definition.Scope][featureID] != null);

                    DataRow row = BuildDataRowFromFeatureDefinition(dtblFeatureList, info, definition, isActive);
                    dtblFeatureList.Rows.Add(row);
                }
                catch (SPException)
                {
                    continue;
                }

            }
        }

        /// <summary>
        /// Builds the data row from a feature definition.
        /// </summary>
        /// <param name="dtblFeatures">The data table features.</param>
        /// <param name="languageCulture">The language culture.</param>
        /// <param name="featdef">The feat definition</param>
        /// <param name="fActive">if set to <c>true</c> [f active].</param>
        /// <returns></returns>
        private static DataRow BuildDataRowFromFeatureDefinition(DataTable dtblFeatures, CultureInfo languageCulture, SPFeatureDefinition featdef, bool fActive)
        {
            DataRow row = dtblFeatures.NewRow();
            row["FeatureId"] = featdef.Id.ToString();
            row["Title"] = featdef.GetTitle(languageCulture);
            row["DisplayName"] = featdef.DisplayName;
            row["Description"] = featdef.GetDescription(languageCulture);
            row["Scope"] = featdef.Scope.ToString();
            if (!fActive)
            {
                row["Status"] = "Inactive";
            }
            else
            {
                row["Status"] = "Active";
            }
            return row;
        }
        #endregion
    }
}
