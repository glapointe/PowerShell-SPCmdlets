using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.ContentTypes;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.SharePoint.PowerShell.Common.Lists;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet(VerbsCommon.Add, "SPEventReceiver", SupportsShouldProcess = false),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Adds an event receiver to a web, list, or content type.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), typeof(SPCmdletGetContentType), ExternalCmdlets = new[] {"Get-SPWeb"})]
    [Example(Code = "PS C:\\> Get-SPWeb http://demo | Add-SPEventReceiver -Name \"My Cool Event Receivers\" -Assembly \"Falchion.SharePoint.MyCoolProject, Version=1.0.0.0, Culture=neutral, PublicKeyToken=3216c23aba16db08\" -ClassName \"Falchion.SharePoint.MyCoolProject.MyCoolEventReceiver\" -Type \"WebProvisioned\",\"WebDeleted\"",
        Remarks = "This example adds WebProvisioned and WebDeleted event receivers to the http://demo site.")]
    [Example(Code = "PS C:\\> Get-SPList http://demo/Lists/MyList | Add-SPEventReceiver -Name \"My Cool Event Receivers\" -Assembly \"Falchion.SharePoint.MyCoolProject, Version=1.0.0.0, Culture=neutral, PublicKeyToken=3216c23aba16db08\" -ClassName \"Falchion.SharePoint.MyCoolProject.MyCoolEventReceiver\" -Type \"ItemUpdating\",\"ItemAdding\"",
        Remarks = "This example adds ItemUpdating and ItemAdding event receivers to the http://demo/Lists/MyList list.")]
    [Example(Code = "PS C:\\> Get-SPWeb http://demo | Get-SPContentType -Identity \"My Content Type\" | Add-SPEventReceiver -Name \"My Cool Event Receivers\" -Assembly \"Falchion.SharePoint.MyCoolProject, Version=1.0.0.0, Culture=neutral, PublicKeyToken=3216c23aba16db08\" -ClassName \"Falchion.SharePoint.MyCoolProject.MyCoolEventReceiver\" -Type \"ItemUpdating\",\"ItemAdding\"",
        Remarks = "This example adds ItemUpdating and ItemAdding event receivers to the \"My Content Type\" content type.")]
    public class SPCmdletAddEventReceiver : SPCmdletCustom
    {
        public SPCmdletAddEventReceiver()
        {
            Sequence = -1;
        }

        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The web to add the event receiver to.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        public SPWebPipeBind Web { get; set; }

        /// <summary>
        /// Gets or sets the list.
        /// </summary>
        /// <value>The list.</value>
        [Parameter(ParameterSetName = "SPList",
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The list to add the event receiver to.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        public SPListPipeBind List { get; set; }

        /// <summary>
        /// Gets or sets the name of the content type.
        /// </summary>
        /// <value>The name of the contentType.</value>
        [Parameter(ParameterSetName = "SPContentType",
            Mandatory = true, 
            Position = 0, 
            ValueFromPipeline = false,
            HelpMessage = "The name of the content type to return. The type must be a valid content type name; a valid content type ID, in the form 0x0123...; or an instance of an SPContentType object.")]
        [ValidateNotNull]
        public SPContentTypePipeBind ContentType { get; set; }

        [Parameter(Mandatory = true,
            Position = 1,
            HelpMessage = "The name to give to the event receiver. The name has no significance but can be useful when later listing the event receivers.")]
        public string Name { get; set; }

        [Parameter(Mandatory = true,
            Position = 2,
            HelpMessage = "The fully qualified assembly name containing the event receiver class to add.")]
        public string Assembly { get; set; }

        [Parameter(Mandatory = true,
            Position = 3,
            HelpMessage = "The fully qualified class name of the event receiver to add.")]
        public string ClassName { get; set; }

        [Parameter(Mandatory = true,
            Position = 4,
            HelpMessage = "The event type to add.  The command does not validate that you are adding the correct type for the specified target or that the specified class contains handlers for the type specified.")]
        public SPEventReceiverType[] Type { get; set; }

        [Parameter(Mandatory = true,
            Position = 5,
            HelpMessage = "The sequence number specifies the order of execution of the event receiver.")]
        public int Sequence { get; set; }

        protected override void InternalProcessRecord()
        {
            List<SPEventReceiverDefinition> def = new List<SPEventReceiverDefinition>();
            SPContentType ct = null;
            switch (ParameterSetName)
            {
                case "SPWeb":
                    SPWeb web = Web.Read();
                    if (web != null)
                    {
                        foreach (SPEventReceiverType type in Type)
                            def.Add(AddEventReceiver.Add(web.EventReceivers, type, Assembly, ClassName, Name));
                        web.Update();
                    }
                    break;
                case "SPList":
                    SPList list = List.Read();

                    if (list != null)
                        foreach (SPEventReceiverType type in Type)
                            def.Add(AddEventReceiver.Add(list.EventReceivers, type, Assembly, ClassName, Name));

                    break;
                case "SPContentType":
                    ct = ContentType.Read();
                    if (ct != null)
                        foreach (SPEventReceiverType type in Type)
                            def.Add(AddEventReceiver.Add(ct.EventReceivers, type, Assembly, ClassName, Name));

                    break;
            }
            if (def.Count > 0)
            {
                if (Sequence >= 0)
                {
                    foreach (SPEventReceiverDefinition erd in def)
                    {
                        if (erd == null) continue;
                        erd.SequenceNumber = Sequence;
                        erd.Update();
                    }
                }
            }

            if (ct != null)
            {
                try
                {
                    ct.Update((ct.ParentList == null));
                }
                catch (Exception ex)
                {
                    Exception ex1 = new Exception("An error occured updating the content type.  Most likely the content type was updated but changes may not have been pushed down to any children.", ex);
                    Logger.WriteException(new ErrorRecord(ex1, null, ErrorCategory.NotSpecified, ct));
                }
            }
        }

    }
}
