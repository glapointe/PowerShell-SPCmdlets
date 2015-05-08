using System.Collections.Generic;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Farm
{
    [Cmdlet(VerbsCommon.Set, "SPDeveloperDashboard", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Farm")]
    [CmdletDescription("Sets the Developer Dashboard Settings.")]
    [RelatedCmdlets(typeof(SPCmdletGetDeveloperDashboard))]
    [Example(Code = "PS C:\\> Set-SPDeveloperDashboard -DisplayLevel OnDemand -TraceEnabled $true",
        Remarks = "This example enables the developer dashboard.")]
    [Example(Code = "PS C:\\> Set-SPDeveloperDashboard -RequiredPermissions \"ManageWeb,ManageSubwebs\"",
        Remarks = "This example sets the required permissions to view the developer dashboard.")]
    public class SPCmdletSetDeveloperDashboard : SPSetCmdletBaseCustom<SPDeveloperDashboardSettings>
    {
        public SPCmdletSetDeveloperDashboard()
        {
            SPDeveloperDashboardSettings dash = SPWebService.ContentService.DeveloperDashboardSettings;
            AutoLaunchEnabled = dash.AutoLaunchEnabled;
            DisplayLevel = dash.DisplayLevel;
            MaximumCriticalEventsToTrack = dash.MaximumCriticalEventsToTrack;
            MaximumSQLQueriesToTrack = dash.MaximumSQLQueriesToTrack;
            RequiredPermissions = dash.RequiredPermissions;
            TraceEnabled = dash.TraceEnabled;
            AdditionalEventsToTrack = ((List<string>) dash.AdditionalEventsToTrack).ToArray();
        }

        [Parameter(HelpMessage = "Indicates whether the developer dashboard can be auto launched.")]
        public bool AutoLaunchEnabled { get; set; }

        [Parameter(HelpMessage = "Indicates whether the developer dashboard is set to Off, On, or On Demand.")]
        public SPDeveloperDashboardLevel DisplayLevel { get; set; }

        [Parameter(HelpMessage = "The maximum number of critical events and asserts that will be recorded in a single transaction (i.e. one request or timer job). If a single transaction has more than this number of asserts the remainder will be ignored. This can be set to 0 to disable assert tracking.")]
        public int MaximumCriticalEventsToTrack { get; set; }

        [Parameter(HelpMessage = "The maximum number of SQL queries that will be recorded in a single transaction (i.e. one request or timer job). If a single transaction executes more than this number of requests the query will be counted but the query call stack and text will not be kept. ")]
        public int MaximumSQLQueriesToTrack { get; set; }

        [Parameter(HelpMessage = "A permission mask defining the permissions required to see the developer dashboard. This defaults to SPBasePermissions.AddAndCustomizePages.")]
        public SPBasePermissions RequiredPermissions { get; set; }

        [Parameter(HelpMessage = "Whether a link to display full verbose trace will be available at the bottom of the page when the developer dashboard is launched or not.")]
        public bool TraceEnabled { get; set; }

        [Parameter(HelpMessage = "A list of URL tags to track in addition to events with severity above High. ")]
        public string[] AdditionalEventsToTrack { get; set; }

        protected override void UpdateDataObject()
        {
            SPDeveloperDashboardSettings dash = SPWebService.ContentService.DeveloperDashboardSettings;

            dash.AutoLaunchEnabled = AutoLaunchEnabled;
            dash.DisplayLevel = DisplayLevel;
            dash.MaximumCriticalEventsToTrack = MaximumCriticalEventsToTrack;
            dash.MaximumSQLQueriesToTrack = MaximumSQLQueriesToTrack;
            dash.RequiredPermissions = RequiredPermissions;
            dash.TraceEnabled = TraceEnabled;
            dash.AdditionalEventsToTrack.Clear();
            ((List<string>)dash.AdditionalEventsToTrack).AddRange(AdditionalEventsToTrack);

            dash.Update();
        }
    }
}
