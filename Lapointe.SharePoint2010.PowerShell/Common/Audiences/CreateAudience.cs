using System;
using System.Collections.Specialized;
using System.Text;
using Microsoft.Office.Server;
using Microsoft.Office.Server.Audience;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace Lapointe.SharePoint.PowerShell.Common.Audiences
{
    public enum RuleEnum
    {
        None, Any, All, Mix
    }

    internal class CreateAudience
    {
        /// <summary>
        /// Creates the specified audience.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="audienceName">Name of the audience.</param>
        /// <param name="description">The description.</param>
        /// <param name="rule">The rule.</param>
        /// <param name="owner">The owner.</param>
        /// <param name="update">if set to <c>true</c> [update].</param>
        /// <returns></returns>
        internal static Audience Create(SPServiceContext context, string audienceName, string description, RuleEnum? rule, string owner, bool update)
        {
            AudienceManager manager = new AudienceManager(context);
            AudienceCollection auds = manager.Audiences;

            Audience audience;
            if (auds.AudienceExist(audienceName))
            {
                if (update)
                {
                    audience = auds[audienceName];
                    audience.AudienceDescription = description ?? "";
                }
                else
                    throw new SPException("Audience name already exists");
            }
            else
                audience = auds.Create(audienceName, description??"");// IMPORTANT: the create method does not do a null check but the methods that load the resultant collection assume not null.

            if (update && rule.HasValue && audience.GroupOperation != AudienceGroupOperation.AUDIENCE_MIX_OPERATION)
            {
                if (rule.Value == RuleEnum.Any && audience.GroupOperation != AudienceGroupOperation.AUDIENCE_OR_OPERATION)
                    audience.GroupOperation = AudienceGroupOperation.AUDIENCE_OR_OPERATION;
                else if (rule.Value == RuleEnum.All && audience.GroupOperation != AudienceGroupOperation.AUDIENCE_AND_OPERATION)
                    audience.GroupOperation = AudienceGroupOperation.AUDIENCE_AND_OPERATION;
            }
            else
            {
                if (audience.GroupOperation != AudienceGroupOperation.AUDIENCE_MIX_OPERATION)
                {
                    if (rule.HasValue)
                    {
                        if (rule == RuleEnum.Any)
                            audience.GroupOperation = AudienceGroupOperation.AUDIENCE_OR_OPERATION;
                        else if (rule == RuleEnum.All)
                            audience.GroupOperation = AudienceGroupOperation.AUDIENCE_AND_OPERATION;
                    }
                    else
                        audience.GroupOperation = AudienceGroupOperation.AUDIENCE_OR_OPERATION;
                }
            }
            if (!string.IsNullOrEmpty(owner))
            {
                audience.OwnerAccountName = owner;
            }
            else
            {
                audience.OwnerAccountName = string.Empty;
            }
            audience.Commit();
            return audience;
        }

    }
}
