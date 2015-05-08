using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Common.Lists;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.Set, "SPListOverlay", SupportsShouldProcess = false),
         SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Sets calendar overlays for the given list.")]
    [RelatedCmdlets(typeof(SPCmdletGetList))]
    [Example(Code = "PS C:\\> Get-SPList \"http://server_name/lists/MyCalendar\" | Set-SPListOverlay -TargetList \"http://server_name/lists/MyOverlayCalendar\" -Color \"Pink\" -ClearExisting",
        Remarks = "This example adds the MyOverlayCalendar calendar as an overlay to the MyCalendar list.")]
    public class SPCmdletSetListOverlay : SPSetCmdletBaseCustom<SPList>
    {
        [Parameter(Mandatory = true,
            HelpMessage = "The color to use for the overlay calendar.",
            ParameterSetName = "Single")]
        [Parameter(Mandatory = true,
            HelpMessage = "The color to use for the overlay calendar.",
            ParameterSetName = "Exchange")]
        public CalendarOverlayColor Color { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The calendar list to add the overlays to.",
            ValueFromPipeline = true,
            Position = 0)]
        public SPListPipeBind TargetList { get; set; }

        [Parameter(HelpMessage = "The name of the view to add the overlays to. If not specified the default view will be used.")]
        public string ViewName { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "Single",
            HelpMessage = "The calendar list to add as an overlay.",
            Position = 1,
            ValueFromPipeline = true)]
        public SPListPipeBind OverlayList { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "Multiple",
            HelpMessage = "The calendar lists to add as an overlay.",
            Position = 1)]
        public SPListPipeBind[] OverlayLists { get; set; }

        [Parameter(HelpMessage = "The title to give the overlay calendar when viewed in the target calendar.",
            ParameterSetName = "Single")]
        [Parameter(Mandatory = true,
            HelpMessage = "The title to give the overlay calendar when viewed in the target calendar.",
            ParameterSetName = "Exchange")]
        public string OverlayTitle { get; set; }

        [Parameter(HelpMessage = "The description to give the overlay calendar when viewed in the target calendar.",
            ParameterSetName = "Single")]
        [Parameter(HelpMessage = "The description to give the overlay calendar when viewed in the target calendar.",
            ParameterSetName = "Exchange")]
        public string OverlayDescription { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "Outlook Web Access URL.",
            ParameterSetName = "Exchange")]
        public string OwaUrl { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "Exchange Web Service URL.",
            ParameterSetName = "Exchange")]
        public string WebServiceUrl { get; set; }

        [Parameter(HelpMessage = "Don't always show the calendar overlay.")]
        public SwitchParameter DoNotAlwaysShow { get; set; }

        [Parameter(HelpMessage = "Clear existing overlays. If not specified then all overlays will be appended to the list of existing overlays (up until 10 - anything after 10 will be ignored)")]
        public SwitchParameter ClearExisting { get; set; }

        protected override void UpdateDataObject()
        {
            SPList targetList = TargetList.Read();
            try
            {
                switch (ParameterSetName)
                {
                    case "Exchange":
                        SetListOverlay.AddCalendarOverlay(targetList, ViewName, OwaUrl, WebServiceUrl, OverlayTitle, OverlayDescription, Color, !DoNotAlwaysShow, ClearExisting);
                        break;
                    case "Single":
                        SPList overlayList = OverlayList.Read();
                        try
                        {
                            SetListOverlay.AddCalendarOverlay(targetList, ViewName, overlayList, OverlayTitle, OverlayDescription, Color, !DoNotAlwaysShow, ClearExisting);
                        }
                        finally
                        {
                            if (overlayList != null)
                            {
                                overlayList.ParentWeb.Dispose();
                                overlayList.ParentWeb.Site.Dispose();
                            }
                        }
                        break;
                    case "Multiple":
                        if (OverlayLists.Length > 10)
                            throw new SPException("You can only have 10 calendar overlays per list.");

                        for (int i = 0; i < OverlayLists.Length; i++)
                        {
                            SPList ol = OverlayLists[i].Read();
                            try
                            {
                                SetListOverlay.AddCalendarOverlay(targetList, ViewName, ol, OverlayTitle, OverlayDescription, (CalendarOverlayColor)i, !DoNotAlwaysShow, ClearExisting && i == 0);
                            }
                            finally
                            {
                                if (ol != null)
                                {
                                    ol.ParentWeb.Dispose();
                                    ol.ParentWeb.Site.Dispose();
                                }
                            }
                        }
                        break;
                }
            }
            finally
            {
                if (targetList != null)
                {
                    targetList.ParentWeb.Dispose();
                    targetList.ParentWeb.Site.Dispose();
                }
            }
        }

    }
}
