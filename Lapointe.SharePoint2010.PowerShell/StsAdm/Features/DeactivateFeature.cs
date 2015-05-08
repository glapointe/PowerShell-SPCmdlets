using System;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Lapointe.SharePoint.PowerShell.StsAdm.SPValidators;
using System.Collections.Specialized;
using Lapointe.SharePoint.PowerShell.Common.Features;
using Microsoft.SharePoint.StsAdmin;

namespace Lapointe.SharePoint.PowerShell.StsAdm.Features
{
    public class DeactivateFeature : SPOperation
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeactivateFeature"/> class.
        /// </summary>
        public DeactivateFeature()
        {
            SPEnumValidator scopeValidator = new SPEnumValidator(typeof(ActivationScope));

            SPParamCollection parameters = new SPParamCollection();
            parameters.Add(new SPParam("filename", "f", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("name", "n", false, null, new SPNonEmptyValidator()));
            parameters.Add(new SPParam("id", "id", false, null, new SPGuidValidator()));
            parameters.Add(new SPParam("url", "url", false, null, new SPUrlValidator()));
            parameters.Add(new SPParam("force", "force"));
            parameters.Add(new SPParam("ignorenonactive", "ignore"));
            parameters.Add(new SPParam("scope", "s", false, "Feature", scopeValidator));


            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n\r\nDeactivates a feature at a given scope.\r\n\r\nParameters:");
            sb.Append("\r\n\t{-filename <relative path to Feature.xml> |");
            sb.Append("\r\n\t -name <feature folder> |");
            sb.Append("\r\n\t -id <feature Id>}");
            sb.AppendFormat("\r\n\t[-scope <{0}> (defaults to Feature)]", scopeValidator.DisplayValue);
            sb.Append("\r\n\t[-url <url>]");
            sb.Append("\r\n\t[-force]");
            sb.Append("\r\n\t[-ignorenonactive]");

            Init(parameters, sb.ToString());
        }

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
        /// Executes the specified command.
        /// </summary>
        /// <param name="command">The command.</param>
        /// <param name="keyValues">The key values.</param>
        /// <param name="output">The output.</param>
        /// <returns></returns>
        public override int Execute(string command, StringDictionary keyValues, out string output)
        {
            output = string.Empty;
            Logger.Verbose = true;

            ActivationScope scope = ActivationScope.Feature;
            if (Params["scope"].UserTypedIn)
                scope = (ActivationScope)Enum.Parse(typeof(ActivationScope), Params["scope"].Value.ToLowerInvariant(), true);

            bool force = Params["force"].UserTypedIn;
            bool ignoreNonActive = Params["ignorenonactive"].UserTypedIn;
            if (ignoreNonActive)
                force = true;

            string url = null;
            if (Params["url"].UserTypedIn)
                url = Params["url"].Value.TrimEnd('/');

            try
            {
                Logger.Write("Started at {0}", DateTime.Now.ToString());
                Guid featureId = FeatureHelper.GetFeatureIdFromParams(Params);
                FeatureHelper fh = new FeatureHelper();
                fh.ActivateDeactivateFeatureAtScope(scope, featureId, false, url, force, ignoreNonActive);
            }
            finally
            {
                Logger.Write("Finished at {0}\r\n", DateTime.Now.ToString());
            }

            return (int)ErrorCodes.NoError;
        }



        /// <summary>
        /// Validates the specified key values.
        /// </summary>
        /// <param name="keyValues">The key values.</param>
        public override void Validate(StringDictionary keyValues)
        {
            SPTriParameterValidator.Validate("name", Params["name"].Value, "id", Params["id"].Value, "filename",
                                             Params["filename"].Value);

            base.Validate(keyValues);
        }


    }
}
