using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Users
{
    [Cmdlet(VerbsCommon.Remove, "SPAllUsers", SupportsShouldProcess = true),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Users")]
    [CmdletDescription("Deletes all Site Collection users. Will not delete site administrators.")]
    [Example(Code = "PS C:\\> Remove-SPAllUsers -Site \"http://demo\"",
        Remarks = "This example removes all site users from the http://demo site.")]
    public class SPCmdletRemoveAllUsers : SPCmdletCustom
    {

        [Parameter(Mandatory = true, ParameterSetName = "SPSite",
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 0,
        HelpMessage = "Specifies the URL or GUID of the Site whose users will be deleted.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        protected override void InternalProcessRecord()
        {
            base.InternalProcessRecord();
            bool test = false;
            ShouldProcessReason reason;
            if (!base.ShouldProcess(null, null, null, out reason))
            {
                if (reason == ShouldProcessReason.WhatIf)
                {
                    test = true;
                }
            }
            if (test)
                Logger.Verbose = true;

            if (Site != null && !test)
            {
                using (SPSite site = Site.Read())
                using (SPWeb web = site.OpenWeb())
                {
                    int offsetIndex = 0;
                    int count = 0;
                    int err = 0;
                    WriteVerbose("Starting user deletion...");
                    while (web.SiteUsers.Count > offsetIndex)
                    {

                        if (web.SiteUsers[offsetIndex].IsSiteAdmin || web.SiteUsers[offsetIndex].ID == web.CurrentUser.ID)
                        {
                            offsetIndex++;
                            continue;
                        }
                        WriteVerbose(string.Format("Progress: Deleting {0}", web.SiteUsers[offsetIndex].LoginName));
                        try
                        {
                            web.SiteUsers.Remove(offsetIndex);
                            count++;
                        }
                        catch (Exception ex)
                        {
                            err++;
                            offsetIndex++;
                            WriteError(new Exception("Unable to delete user.", ex), ErrorCategory.NotSpecified, web);
                        }
                    }
                    WriteVerbose(string.Format("Finished user deletion.  {0} Users deleted, {1} errors.", count.ToString(), err.ToString()));
                }
            }
        }


    }
}
