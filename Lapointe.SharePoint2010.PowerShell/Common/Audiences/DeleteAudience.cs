using System.Collections;
using System.Collections.Specialized;
using System.Text;
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
using Microsoft.SharePoint;
using System;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Common.Audiences
{
    internal class DeleteAudience
    {
        /// <summary>
        /// Deletes the specified audience or all audience rules for the specified audience.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="audienceName">Name of the audience.</param>
        /// <param name="deleteRulesOnly">if set to <c>true</c> [delete rules only].</param>
        internal static void Delete(SPServiceContext context, string audienceName, bool deleteRulesOnly)
        {
            AudienceManager manager = new AudienceManager(context);

            if (!manager.Audiences.AudienceExist(audienceName))
            {
                throw new SPException("Audience name does not exist");
            }

            Audience audience = manager.Audiences[audienceName];

            if (audience.AudienceRules != null && deleteRulesOnly)
            {
                audience.AudienceRules = new ArrayList();
                if (audience.GroupOperation == AudienceGroupOperation.AUDIENCE_MIX_OPERATION)
                {
                    // You can't change from mixed mode using the property without setting some internal fields.
                    object audienceInfo = Utilities.GetFieldValue(audience, "m_AudienceInfo");
                    Utilities.SetPropertyValue(audienceInfo, "NewGroupOperation", AudienceGroupOperation.AUDIENCE_OR_OPERATION);
                    Utilities.SetFieldValue(audience, typeof (Audience), "m_AuidenceGroupOperationChanged", true);
                }
                audience.Commit();
                return;
            }
            if (!deleteRulesOnly)
                manager.Audiences.Remove(audience.AudienceID);
        }

    }
}
