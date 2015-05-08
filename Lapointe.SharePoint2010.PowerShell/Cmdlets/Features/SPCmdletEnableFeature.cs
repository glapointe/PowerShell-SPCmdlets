using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint;
using Lapointe.SharePoint.PowerShell.Common.Features;
using System.ComponentModel;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Features
{
    [Cmdlet("Enable", "SPFeature2", SupportsShouldProcess = false),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Features")]
    [CmdletDescription("Activates a feature or features at a given scope.")]
    [RelatedCmdlets(typeof(SPCmdletDisableFeature),
        ExternalCmdlets = new[] {"Disable-SPFeature", "Enable-SPFeature", "Get-SPFeature"})]
    [Example(Code = "PS C:\\> Get-SPFeature MyFeature | Enable-SPFeature2 -Url \"http://server_name\" -ActivateAtScope WebApplication",
        Remarks = "This example activates MyFeature in the http://server_name web application. If the Feature is scoped to a site collection or web then it will activate at all sites or webs.")]
    [Example(Code = "PS C:\\> Get-SPFeature MyFeature | Enable-SPFeature2 -Url \"http://server_name\" -IgnoreNonActive -ActivateAtScope WebApplication",
        Remarks = "This example activates MyFeature in the http://server_name web application where it is already active thus causing any feature activation code to rerun.")]
    public class SPCmdletEnableFeature : SPSetCmdletBaseCustom<SPFeatureDefinition>
    {
        bool farmLevelFeature = false;
        SPSite m_Site;

        
        [Parameter(Mandatory = false,
            HelpMessage = "If provided, the cmdlet outputs the Feature definition object after enabling.")]
        public SwitchParameter PassThru { get; set; }

        [Parameter(Mandatory = true,
            HelpMessage = "Specifies the name of the Feature or GUID to activate.\r\n\r\nThe type must be the name of the Feature folder located in the 14\\Template\\Features folder or GUID, in the form 21d186e1-7036-4092-a825-0eb6709e9281.")]
        public SPFeatureDefinitionPipeBind Identity { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Specifies the URL of the Web application, Site Collection, or Site for which the Feature is being activated.\r\n\r\nThe type must be a valid URL; for example, http://server_name.")]
        public string Url { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Forces the activation of a Feature. This causes any custom code associated with the Feature to rerun.")]
        public SwitchParameter Force { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Ignore scopes where the feature is not already activated.")]
        public SwitchParameter IgnoreNonActive { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "Activate features at the specified scope only. Valid values are Farm, WebApplication, Site, Web, and Feature")]
        public ActivationScope? ActivateAtScope { get; set; }


        protected override void InternalValidate()
        {
            if (this.Url != null)
            {
                this.Url = this.Url.Trim();
            }
            try
            {
                base.DataObject = this.Identity.Read();
                this.farmLevelFeature = true;
            }
            catch (SPCmdletPipeBindException)
            {
            }
            if (!this.farmLevelFeature)
            {
                if (string.IsNullOrEmpty(this.Url))
                {
                    throw new SPCmdletException("The specified feature is not a farm scoped feature. Please specify a valid Url.");
                }
                try
                {
                    this.m_Site = new SPSitePipeBind(this.Url).Read(false);
                    base.DataObject = this.Identity.Read(this.m_Site, true);
                }
                catch (SPCmdletPipeBindException)
                {
                    throw new SPCmdletException("The specified site collection could not be found.");
                } 
            }
        }


        protected override void UpdateDataObject()
        {
            SPFeatureDefinition featureDef = null;
            ActivationScope scope = ActivationScope.Feature;

            if (ActivateAtScope.HasValue)
                scope = ActivateAtScope.Value;

            if (farmLevelFeature)
            {
                featureDef = SPFarm.Local.FeatureDefinitions[base.DataObject.Id];
            }
            else
            {
                featureDef = base.DataObject;
            }

            try
            {
                Logger.Write("Started at {0}", DateTime.Now.ToString());
                Guid featureId = featureDef.Id;
                FeatureHelper fh = new FeatureHelper();
                fh.ActivateDeactivateFeatureAtScope(featureDef, scope, true, Url, Force.IsPresent, IgnoreNonActive.IsPresent);
            }
            finally
            {
                Logger.Write("Finished at {0}\r\n", DateTime.Now.ToString());
            }

            if ((featureDef != null) && (PassThru.IsPresent))
            {
                base.DataObject = featureDef;
                base.WriteResult(base.DataObject);
            }
        }

    }

}
