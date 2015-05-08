using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers;
using Microsoft.SharePoint;

namespace Lapointe.SharePoint.PowerShell.Common.Lists
{
    public class AddEventReceiver
    {
        public enum TargetEnum
        {
            Site, List, ContentType
        }

        /// <summary>
        /// Adds an event receiver to the specified target
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="contentTypeName">Name of the content type.</param>
        /// <param name="target">The target.</param>
        /// <param name="assembly">The assembly.</param>
        /// <param name="className">Name of the class.</param>
        /// <param name="type">The type.</param>
        /// <param name="sequence">The sequence.</param>
        /// <param name="name">The name.</param>
        public static void Add(string url, string contentTypeName, TargetEnum target, string assembly, string className, SPEventReceiverType type, int sequence, string name)
        {

            using (SPSite site = new SPSite(url))
            using (SPWeb web = site.OpenWeb())
            {
                SPContentType contentType = null;
                SPEventReceiverDefinitionCollection eventReceivers;
                if (target == TargetEnum.List)
                {
                    SPList list = Utilities.GetListFromViewUrl(web, url);

                    if (list == null)
                    {
                        throw new Exception("List not found.");
                    }
                    eventReceivers = list.EventReceivers;
                }
                else if (target == TargetEnum.Site)
                    eventReceivers = web.EventReceivers;
                else
                {
                    try
                    {
                        contentType = web.AvailableContentTypes[contentTypeName];
                    }
                    catch (ArgumentException)
                    {
                    }
                    if (contentType == null)
                        throw new SPSyntaxException("The specified content type could not be found.");

                    eventReceivers = contentType.EventReceivers;
                }
                SPEventReceiverDefinition def = Add(eventReceivers, type, assembly, className, name);
                if (sequence >= 0)
                {
                    def.SequenceNumber = sequence;
                    def.Update();
                }
                if (contentType != null)
                {
                    try
                    {
                        contentType.Update((contentType.ParentList == null));
                    }
                    catch (Exception ex)
                    {
                        Exception ex1 = new Exception("An error occured updating the content type.  Most likely the content type was updated but changes may not have been pushed down to any children.", ex);
                        Logger.WriteException(new ErrorRecord(ex1, null, ErrorCategory.NotSpecified, contentType));
                    }
                }
            }
        }

        /// <summary>
        /// Adds an event receiver to a the specified event receiver definition collection.
        /// </summary>
        /// <param name="eventReceivers">The event receivers.</param>
        /// <param name="eventReceiverType">Type of the event receiver.</param>
        /// <param name="assembly">The assembly.</param>
        /// <param name="className">Name of the class.</param>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public static SPEventReceiverDefinition Add(SPEventReceiverDefinitionCollection eventReceivers, SPEventReceiverType eventReceiverType, string assembly, string className, string name)
        {
            SPEventReceiverDefinition def = GetEventReceiver(eventReceivers, eventReceiverType, assembly, className);
            if (def == null)
            {
                eventReceivers.Add(eventReceiverType, assembly, className);
                def = GetEventReceiver(eventReceivers, eventReceiverType, assembly, className);
                if (def != null && !String.IsNullOrEmpty(name))
                {
                    def.Name = name;
                    def.Update();
                }
                return def;
            }
            return def;
        }

        /// <summary>
        /// Gets the event receiver.
        /// </summary>
        /// <param name="eventReceivers">The event receivers.</param>
        /// <param name="eventReceiverType">Type of the event receiver.</param>
        /// <param name="assembly">The assembly.</param>
        /// <param name="className">Name of the class.</param>
        /// <returns></returns>
        private static SPEventReceiverDefinition GetEventReceiver(SPEventReceiverDefinitionCollection eventReceivers, SPEventReceiverType eventReceiverType, string assembly, string className)
        {
            foreach (SPEventReceiverDefinition erd in eventReceivers)
            {
                if (erd.Assembly.ToLower() == assembly.ToLower() && erd.Class.ToLower() == className.ToLower() && erd.Type == eventReceiverType)
                {
                    return erd;
                }
            }
            return null;
        }
    }
}
