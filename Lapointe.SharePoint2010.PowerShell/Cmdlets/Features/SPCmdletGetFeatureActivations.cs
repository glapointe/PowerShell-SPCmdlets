using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Features
{
    [Cmdlet(VerbsCommon.Get, "SPFeatureActivations", SupportsShouldProcess = false, DefaultParameterSetName = "SPFeature"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Features")]
    [CmdletDescription("Retrieves all publishing pages from the specified source.")]
    [RelatedCmdlets(typeof(SPCmdletEnableFeature), typeof(SPCmdletDisableFeature), ExternalCmdlets = new[] { "Get-SPFeature", "Enable-SPFeature", "Disable-SPFeature" })]
    [Example(Code = "PS C:\\> $features = Get-SPFeatureActivations -Identity \"TeamCollab\" | select @{Expression={$_.Parent.Url}}",
        Remarks = "This example returns back all Feature activations for the Site-scoped Feature TeamCollab and displays the URL of the Site where the Feature is activated.")]
    [Example(Code = "PS C:\\> $features = Get-SPFeatureActivations -Solution \"MyCustomSolution.wsp\"",
        Remarks = "This example returns back all Feature activations for all Features defined by the MyCustomSolution.wsp Solution Package.")]
    public class SPCmdletGetFeatureActivations : SPGetCmdletBaseCustom<SPFeature>
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "SPFeature",
            HelpMessage = "Specifies the name of the Feature or GUID to activate.\r\n\r\nThe type must be the name of the Feature folder located in the 14\\Template\\Features folder or GUID, in the form 21d186e1-7036-4092-a825-0eb6709e9281.")]
        [Alias("Feature")]
        [ValidateNotNull]
        public SPFeatureDefinitionPipeBind[] Identity { get; set; }

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "SPSolution",
            HelpMessage = "Specifies the SharePoint Slution Package containing the Features whose activations will be retrieved.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of a SharePoint Solution (for example, SPSolution1); or an instance of a valid SPSolution object.")]
        [ValidateNotNull]
        public SPSolutionPipeBind[] Solution { get; set; }

        [Parameter(Mandatory = false, Position = 1)]
        public SwitchParameter NeedsUpgrade { get; set; }

        protected override IEnumerable<SPFeature> RetrieveDataObjects()
        {
            List<SPFeature> features = new List<SPFeature>();

            switch (ParameterSetName)
            {
                case "SPFeature":
                    foreach (SPFeatureDefinitionPipeBind fdpb in Identity)
                    {
                        SPFeatureDefinition feature = fdpb.Read();
                        GetFeatureActivations(features, feature, NeedsUpgrade);
                    }
                    break;
                case "SPSolution":
                    foreach (SPSolutionPipeBind spb in Solution)
                    {
                        SPSolution solution = spb.Read();
                        List<SPFeatureDefinition> featureDefinitions = SPFarm.Local.FeatureDefinitions.Where(fd => fd.SolutionId == solution.Id).ToList();
                        foreach (SPFeatureDefinition feature in featureDefinitions)
                        {
                            GetFeatureActivations(features, feature, NeedsUpgrade);
                        }
                    }
                    break;
            }
            return features;
        }

        /// <summary>
        /// Gets the feature activations.
        /// </summary>
        /// <param name="features">The features.</param>
        /// <param name="fd">The fd.</param>
        /// <param name="needsUpgrade">if set to <c>true</c> [needs upgrade].</param>
        private void GetFeatureActivations(List<SPFeature> features, SPFeatureDefinition fd, bool needsUpgrade)
        {
            switch (fd.Scope) 
            {
                case SPFeatureScope.Farm:
                    features.AddRange(SPWebService.AdministrationService.QueryFeatures(fd.Id, needsUpgrade));
                    break;
            
                case SPFeatureScope.WebApplication:
                    features.AddRange(SPWebService.QueryFeaturesInAllWebServices(fd.Id, needsUpgrade));
                    break;
                case SPFeatureScope.Site:
                    foreach (SPService svc in SPFarm.Local.Services)
                    {
                        if (!(svc is SPWebService))
                            continue;

                        foreach (SPWebApplication webApp in ((SPWebService) svc).WebApplications)
                        {
                            features.AddRange(webApp.QueryFeatures(fd.Id, needsUpgrade));
                        }
                    }
                    break;
                case SPFeatureScope.Web:
                    foreach (SPService svc in SPFarm.Local.Services)
                    {
                        if (!(svc is SPWebService))
                            continue;

                        foreach (SPWebApplication webApp in ((SPWebService) svc).WebApplications)
                        {
                            foreach (SPSite site in webApp.Sites)
                            {
                                features.AddRange(site.QueryFeatures(fd.Id, needsUpgrade));
                                site.Dispose();
                            }
                        }
                    }
                    break;
            }
        }
    }
}
