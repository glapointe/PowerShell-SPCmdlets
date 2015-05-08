using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.UserProfiles
{
    [Cmdlet(VerbsCommon.Set, "SPProfilePictureUrl", SupportsShouldProcess = false),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("User Profiles")]
    [CmdletDescription("Sets the picture URL path for user profiles. The following variables may be used for dynamic replacement: \"@(username)\", \"@(domain)\", \"@(email)\", \"@(firstname)\", \"@(lastname)\", \"@(employeeid)\".")]
    [Example(Code = "PS C:\\> Set-SPProfilePictureUrl -UserProfileServiceApplication \"30daa535-b0fe-4d10-84b0-fb04029d161a\" -Username \"domain\\username\" -Path \"http://intranet/hr/EmployeePictures/@(username).jpg\" -Overwrite -ValidateUrl",
        Remarks = "This example sets the picture url of a user in the user profile service application with ID \"30daa535-b0fe-4d10-84b0-fb04029d161a\".")]
    [RelatedCmdlets(ExternalCmdlets = new[] { "Get-SPServiceApplication" })]
    public class SPCmdletSetProfilePictureUrl : SPSetCmdletBaseCustom<UserProfile>
    {
        [Parameter(ParameterSetName = "Username_UPA", 
            Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true,
            HelpMessage = "Specifies the service application that contains the user profiles to update.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a service application (for example, ServiceApp1); or an instance of a valid SPServiceApplication object.")]
        [ValidateNotNull]
        public SPServiceApplicationPipeBind UserProfileServiceApplication { get; set; }

        [Parameter(ParameterSetName = "Username_UPA", 
            Mandatory = false,
            HelpMessage = "Specifies the site subscription containing the user profiles to update.\r\n\r\nThe type must be a valid URL, in the form http://server_name; a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a site subscription (for example, SiteSubscription1); or an instance of a valid SiteSubscription object.")]
        public SPSiteSubscriptionPipeBind SiteSubscription { get; set; }

        [Parameter(ParameterSetName = "Username_SPSite",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Site to use for retrieving the service context. Use this parameter when the service application is not associated with the default proxy group or more than one custom proxy groups.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind ContextSite { get; set; }


        [Parameter(ParameterSetName = "Username_UPA",
            Mandatory = false, 
            HelpMessage = "The username corresponding to the user profile to update in the form domain\\username.")]
        [Parameter(ParameterSetName = "Username_SPSite",
            Mandatory = false,
            HelpMessage = "The username corresponding to the user profile to update in the form domain\\username.")]
        [ValidateNotNullOrEmpty]
        public string Username { get; set; }

        [Parameter(ParameterSetName = "UserProfile",
            Mandatory = true,
            HelpMessage = "The user profile to update.")]
        [ValidateNotNullOrEmpty]
        public UserProfile UserProfile { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "Path to new photo (i.e., \"http://intranet/hr/EmployeePictures/@(username).jpg\") - leave blank to clear.The path to the images. To substitute dynamic data use the following strings variables within the path: @(username), @(domain), @(email), @(firstname), @(lastname), @(employeeid). The variable names are case sensitive.")]
        [ValidateNotNullOrEmpty]
        public string Path { get; set; }

        [Parameter(HelpMessage = "If provided then existing data values will be overwritten. If omitted then any profile objects with existing data will be ignored.")]
        public SwitchParameter Overwrite { get; set; }

        [Parameter(HelpMessage = "If specified then do not error if a specified variable value cannot be found. Note that if the value is not found then the property value will not be set.")]
        public SwitchParameter IgnoreMissingData { get; set; }

        [Parameter(HelpMessage = "If specified then perform a web request to see if the resultant URL is valid. If the result is not valid then the property value will be set to an empty string.")]
        public SwitchParameter ValidateUrl { get; set; }

        protected override void UpdateDataObject()
        {
            SPServiceContext context = null;
            UserProfileManager profManager = null;
            switch (ParameterSetName)
            {
                case "Username_UPA":
                    SPSiteSubscriptionIdentifier subId;
                    if (SiteSubscription != null)
                    {
                        SPSiteSubscription siteSub = SiteSubscription.Read();
                        subId = siteSub.Id;
                    }
                    else
                        subId = SPSiteSubscriptionIdentifier.Default;

                    SPServiceApplication svcApp = UserProfileServiceApplication.Read();
                    context = SPServiceContext.GetContext(svcApp.ServiceApplicationProxyGroup, subId);
                    profManager = new UserProfileManager(context);

                    if (string.IsNullOrEmpty(Username))
                        Common.UserProfiles.SetPictureUrl.SetPictures(profManager, Path, Overwrite, IgnoreMissingData, ValidateUrl);
                    else
                        Common.UserProfiles.SetPictureUrl.SetPicture(profManager, Username, Path, Overwrite, IgnoreMissingData, ValidateUrl);

                    break;
                case "Username_SPSite":
                    using (SPSite site = ContextSite.Read())
                    {
                        context = SPServiceContext.GetContext(site);
                    }
                    profManager = new UserProfileManager(context);

                    if (string.IsNullOrEmpty(Username))
                        Common.UserProfiles.SetPictureUrl.SetPictures(profManager, Path, Overwrite, IgnoreMissingData, ValidateUrl);
                    else
                        Common.UserProfiles.SetPictureUrl.SetPicture(profManager, Username, Path, Overwrite, IgnoreMissingData, ValidateUrl);

                    break;
                case "UserProfile":
                    Common.UserProfiles.SetPictureUrl.SetPicture(UserProfile, Path, Overwrite, IgnoreMissingData, ValidateUrl);
                    break;
            }
        }

    }
}
