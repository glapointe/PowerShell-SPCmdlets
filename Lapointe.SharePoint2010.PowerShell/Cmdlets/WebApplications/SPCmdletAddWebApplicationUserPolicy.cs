using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WebApplications
{
    [Cmdlet(VerbsCommon.Add, "SPWebApplicationUserPolicy", SupportsShouldProcess = true), 
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Web Applications")]
    [CmdletDescription("Adds a new user policy to the specified web application.")]
    [RelatedCmdlets(ExternalCmdlets = new[] {"Get-SPWebApplication", "Get-SPWeb", "Get-SPUser"})]
    [Example(Code = "PS C:\\> Add-SPWebApplicationUserPolicy -WebApplication http://portal -UserLogin \"user\\domain\" -RoleName \"Full Read\",\"Deny Write\"",
        Remarks = "This example grants user\\domain full read and deny write access to http://portal.")]
    public class SPCmdletAddWebApplicationUserPolicy : SPCmdletCustom
    {
        [Parameter(ParameterSetName = "SPUser",
            Mandatory = true, 
            ValueFromPipeline = true, 
            ValueFromPipelineByPropertyName = true)]
        public SPUserPipeBind User { get; set; }
        
        [Parameter(ParameterSetName = "SPUser",
            Mandatory = false,
            HelpMessage = "Specifies the URL or GUID of the Web to which the user belongs.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(ParameterSetName = "Login",
            Mandatory = true,
            HelpMessage = "The user login in the form domain\\user.")]
        public string UserLogin { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The web application to which the user policy will be added.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        public SPWebApplicationPipeBind WebApplication { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The zone to add the policy to. To add to all zones to not set this parameter. Valid values are \"Custom\", \"Default\", \"Intranet\", \"Internet\", \"Extranet\".")]
        public SPUrlZone[] Zone { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "The roles name(s) to assign the user to. Default values are \"Full Control\", \"Full Read\",\"Deny Write\", and \"Deny All\".")]
        public string[] RoleName { get; set; }

        protected override void InternalProcessRecord()
        {
            SPWebApplication webApp = WebApplication.Read();
            string login = "";
            if (ParameterSetName == "SPUser")
            {
                if (User != null)
                {
                    SPUser user;
                    try
                    {
                        if (Web != null)
                            user = User.Read(Web.Read());
                        else
                            user = User.Read();
                    }
                    catch
                    {
                        throw new SPCmdletException("The SPUser object is not valid.");
                    }
                    login = user.LoginName;
                }
            }
            else
            {
                login = UserLogin;
            }

            if (Zone == null)
            {
                Common.WebApplications.AddUserPolicyForWebApp.AddUserPolicy(login, null, RoleName, webApp);
            }
            else
            {
                Common.WebApplications.AddUserPolicyForWebApp.AddUserPolicy(login, null, RoleName, webApp, Zone);
            }
        }
    }
}
