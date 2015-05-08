using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.Win32;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation.Internal;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebApplications
{
    [Cmdlet(VerbsCommon.Set, "SPBackConnectionHostNames", SupportsShouldProcess = true, DefaultParameterSetName = "Local"), 
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Web Applications")]
    [CmdletDescription("Sets the back connection host names on the local server or all servers in the farm. An IIS reset will be necessary after the cmdlet completes.")]
    [RelatedCmdlets(ExternalCmdlets = new[] {"Get-SPWebApplication"})]
    [Example(Code = "PS C:\\> Set-SPBackConnectionHostNames",
        Remarks = "This example sets the back connection host names for all web applications in the farm on the local server only.")]
    [Example(Code = "PS C:\\> Set-SPBackConnectionHostNames -UpdateFarm -FarmCredentials (Get-Credential)",
        Remarks = "This example sets the back connection host names for all web applications in the farm on all servers in the farm using the provided credentials.")]
    public class SPCmdletSetBackConnectionHostNames : SPSetCmdletBaseCustom<PSObject>
    {
        #region Parameters

        [Parameter(ParameterSetName = "UpdateFarm",
            Mandatory = true, 
            Position = 1,
            HelpMessage = "Update all servers in the farm.")]
        public SwitchParameter UpdateFarm { get; set; }

        [Parameter(ParameterSetName = "UpdateFarm",
            Mandatory = true, 
            Position = 2, 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The credentials to use on the remote servers to perform the update.")]
        public PSCredential FarmCredentials  { get; set; }

        [Parameter(ParameterSetName = "UpdateFarm",
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The web application whose URLs will be added to the set of back connection host names.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        [Parameter(ParameterSetName = "Local",
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The web application whose URLs will be added to the set of back connection host names.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        public SPWebApplicationPipeBind WebApplication {  get; set; }

        #endregion

        protected override void UpdateDataObject()
        {
            List<string> urls = new List<string>();
            if (WebApplication != null)
            {
                Common.WebApplications.SetBackConnectionHostNames.GetUrls(urls, WebApplication.Read());
            }
            else
                urls = Common.WebApplications.SetBackConnectionHostNames.GetUrls();

            string shouldProcessMsg = "Update the local server with appropriate back connection host names.";
            if (UpdateFarm.IsPresent)
                shouldProcessMsg = "Update all servers with appropriate back connection host names.";

            ShouldProcessReason reason;
            if (!base.ShouldProcess(shouldProcessMsg, null, null, out reason))
            {
                if (reason == ShouldProcessReason.WhatIf)
                {
                    base.WriteResult(urls);
                }
                return;
            }

            if (!UpdateFarm.IsPresent)
                Common.WebApplications.SetBackConnectionHostNames.SetBackConnectionRegKey(urls);
            else
            {
                SPTimerService timerService = SPFarm.Local.TimerService;
                if (null == timerService)
                {
                    throw new SPException("The Farms timer service cannot be found.");
                }
                Common.WebApplications.SetBackConnectionHostNamesTimerJob job = new Common.WebApplications.SetBackConnectionHostNamesTimerJob(timerService);

                job.SubmitJob(FarmCredentials.UserName,
                    Utilities.ConvertToUnsecureString(FarmCredentials.Password),
                    urls);

                WriteResult("Timer job successfully created.");
            }
        }


        protected override void InternalValidate()
        {
            if (UpdateFarm.IsPresent && FarmCredentials == null)
            {
                FarmCredentials = Host.UI.PromptForCredential("Farm Credentials Required",
                    "Please provide an account with rights to update the registry on each server.", null, Environment.UserDomainName);

                if (FarmCredentials == null)
                    throw new SPCmdletException("A valid username with rights to edit the registry is required.");
                else
                {
                    if (!Utilities.ValidateCredentials(FarmCredentials))
                    {
                        throw new SPCmdletException("Invalid Username or password");
                    }

                }
            }
        }

    }
}
