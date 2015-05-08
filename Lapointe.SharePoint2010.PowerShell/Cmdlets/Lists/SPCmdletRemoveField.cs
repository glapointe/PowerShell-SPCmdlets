using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Remove", "SPField", SupportsShouldProcess = true),
        SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("The Remove-SPField cmdlet deletes the specified field from the given list or web.")]
    [Example(Code = "PS C:\\> Remove-SPField -Web \"http://demo\" -InternalFieldName \"MyField\"",
        Remarks = "This example deletes the \"MyField\" field from the web \"http://demo\".")]
    [Example(Code = "PS C:\\> Remove-SPField -List \"http://demo/documents\" -InternalFieldName \"MyField\"",
        Remarks = "This example deletes the \"MyField\" field from the list \"http://demo/documents\".")]
    [Example(Code = "PS C:\\> (Get-SPList \"http://demo/documents\").Fields.GetFieldByInternalName(\"MyField\") | Remove-SPField -InternalFieldName \"MyField\"",
        Remarks = "This example deletes the \"MyField\" field from the list \"http://demo/documents\".")]
    [RelatedCmdlets(typeof(SPCmdletGetList))]
    public class SPCmdletRemoveField : SPRemoveCmdletBaseCustom<SPField>
    {

        /// <summary>
        /// Gets or sets the web.
        /// </summary>
        /// <value>The web.</value>
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The web to remove the field from.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
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
            HelpMessage = "The list to remove the field from.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        [ValidateNotNull]
        public SPListPipeBind List { get; set; }

        /// <summary>
        /// Gets or sets the field.
        /// </summary>
        /// <value>The list.</value>
        [Parameter(ParameterSetName = "SPField",
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The field to remove.\r\n\r\nThe value must be a valid SPField object.")]
        [ValidateNotNull]
        public SPField Field { get; set; }

        [Parameter(ParameterSetName = "SPList",
            Mandatory = true,
            Position = 1,
            ValueFromPipeline = false,
            ValueFromPipelineByPropertyName = false,
            HelpMessage = "The internal name of the field to remove from the list.")]
        [Parameter(ParameterSetName = "SPWeb",
            Mandatory = true,
            Position = 1,
            ValueFromPipeline = false,
            ValueFromPipelineByPropertyName = false,
            HelpMessage = "The internal name of the field to remove from the web.")]
        [ValidateNotNullOrEmpty]
        public string InternalFieldName { get; set; }

        [Parameter(HelpMessage = "Used to force deletion if AllowDeletion property is false.")]
        public SwitchParameter Force { get; set; }

        private SPField GetField(SPFieldCollection fields, string name)
        {
            SPField field = null;
            try
            {
                field = fields.GetFieldByInternalName(name);
            }
            catch {}
            return field;
        }
        protected override void InternalValidate()
        {
            switch (ParameterSetName)
            {
                case "SPWeb":
                    SPWeb web = Web.Read();
                    if (web != null)
                        DataObject = GetField(web.Fields, InternalFieldName);
                    break;
                case "SPList":
                    SPList list = List.Read();
                    if (list != null)
                        DataObject = GetField(list.Fields, InternalFieldName);
                    break;
                case "SPField":
                        DataObject = Field;
                    break;
            }


            if (DataObject == null)
            {
                WriteError(new PSArgumentException("A field with the name " + InternalFieldName + " could not be found."), ErrorCategory.InvalidArgument, null);
                SkipProcessCurrentRecord();
            }
        }

        protected override void DeleteDataObject()
        {
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
            
            if (DataObject != null)
            {
                SPField field = DataObject;
                if (field.ReadOnlyField && Force)
                {
                    if (!test)
                    {
                        field.ReadOnlyField = false;
                        field.Update();
                    }
                    else
                        Logger.Write("Field is marked as Read Only and must be changed to delete.");
                }
                if (!field.CanBeDeleted)
                {
                    if (field.FromBaseType)
                    {
                        throw new Exception(
                            "The field is derived from a base type and cannot be deleted. You must delete the field from the base type.");
                    }
                    if (field.Sealed)
                    {
                        if (Force)
                        {
                            if (!test)
                            {
                                field.Sealed = false;
                                field.Update();
                            }
                            else
                                Logger.Write("Field is marked as Sealed and must be changed to delete.");
                        }
                        else
                            throw new Exception("This field is sealed and cannot be deleted - specify \"-Force\" to ignore this setting and attempt deletion regardless.");
                    }
                    else if (field.AllowDeletion.HasValue && !field.AllowDeletion.Value && !Force)
                    {
                        throw new Exception(
                            "Field is marked as not allowing deletion - specify \"-Force\" to ignore this setting and attempt deletion regardless.");
                    }
                    else if (field.AllowDeletion.HasValue && !field.AllowDeletion.Value && Force)
                    {
                        if (!test)
                        {
                            field.AllowDeletion = true;
                            field.Update();
                        }
                        else
                            Logger.Write("Field is marked as not to Allow Deletion and must be changed to delete.");

                    }
                    else
                    {
                        throw new Exception("Field cannot be deleted.");
                    }
                }
                if (field.Hidden)
                {
                    if (Force)
                    {
                        if (field.CanToggleHidden)
                        {
                            if (!test)
                            {
                                field.Hidden = false;
                                field.Update();
                            }
                            else
                                Logger.Write("Field is marked as Hidden and must be changed to delete.");
                        }
                        else
                        {
                            if (!test)
                            {
                                MethodInfo setFieldBoolValue = field.GetType().GetMethod("SetFieldBoolValue",
                                        BindingFlags.NonPublic | BindingFlags.Public |
                                        BindingFlags.Instance | BindingFlags.InvokeMethod,
                                        null, new Type[] { typeof(string), typeof(bool) }, null);

                                //field.SetFieldBoolValue("Hidden", false);
                                setFieldBoolValue.Invoke(field, new object[] { "Hidden", false });
                                //field.SetFieldBoolValue("CanToggleHidden", true);
                                setFieldBoolValue.Invoke(field, new object[] { "CanToggleHidden", true });
                                field.Update();
                            }
                            else
                                Logger.Write("Field is marked as Hidden and must be changed to delete.");

                        }
                    }
                    else
                        throw new Exception(
                            "You cannot delete hidden fields - specify \"-Force\" to ignore this restriction and attempt deletion regardless.");
                }
                if (!test)
                    field.Delete();
            }
        }
    }
}
