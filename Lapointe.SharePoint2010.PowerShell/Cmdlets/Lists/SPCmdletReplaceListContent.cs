using System.Text;
using System.Collections.Generic;
using Lapointe.SharePoint.PowerShell.Common.Lists;
using Microsoft.SharePoint;
using Microsoft.SharePoint.PowerShell;
using Microsoft.SharePoint.Administration;
using System.Management.Automation;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Lapointe.SharePoint.PowerShell.Cmdlets.SPValidators;
using Lapointe.PowerShell.MamlGenerator.Attributes;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Replace", "SPListContent", SupportsShouldProcess = true, DefaultParameterSetName="List_ParameterInput"),
    SPCmdlet(RequireLocalFarmExist = true, RequireUserMachineAdmin = false, RequireUserFarmAdmin = true)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Replaces all occurances of the search string with the replacement string. Supports the use of regular expressions. Use -WhatIf to verify your replacements before executing.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), ExternalCmdlets = new[] {"Get-SPFarm", "Get-SPWebApplication", "Get-SPSite", "Get-SPWeb"})]
    [Example(Code = "PS C:\\> Get-SPWeb http://portal | Replace-SPListContent -SearchString \"(?i:old product name)\" -ReplaceString \"New Product Name\" -Publish",
        Remarks = "This example does a case-insensitive search for \"old product name\" and replaces with \"New Product Name\" and publishes the changes after completion.")]
    public class SPCmdletReplaceListContent : SPCmdletCustom
    {
        [Parameter(Mandatory = true,
            ParameterSetName = "List_InputFile",
            HelpMessage = "A file with search and replace strings, seperated by a \"|\" character (each search and replace string should be on a separate line).")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Web_InputFile",
            HelpMessage = "A file with search and replace strings, seperated by a \"|\" character (each search and replace string should be on a separate line).")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Site_InputFile",
            HelpMessage = "A file with search and replace strings, seperated by a \"|\" character (each search and replace string should be on a separate line).")]
        [Parameter(Mandatory = true,
            ParameterSetName = "WebApplication_InputFile",
            HelpMessage = "A file with search and replace strings, seperated by a \"|\" character (each search and replace string should be on a separate line).")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Farm_InputFile",
            HelpMessage = "A file with search and replace strings, seperated by a \"|\" character (each search and replace string should be on a separate line).")]
        [ValidateDirectoryExistsAndValidFileName]
        public string InputFile { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "List_InputFile",
            HelpMessage = "The delimiter used within the input file specified by the -InputFile parameter (the default is \"|\" if not specified).")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Web_InputFile",
            HelpMessage = "The delimiter used within the input file specified by the -InputFile parameter (the default is \"|\" if not specified).")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Site_InputFile",
            HelpMessage = "The delimiter used within the input file specified by the -InputFile parameter (the default is \"|\" if not specified).")]
        [Parameter(Mandatory = true,
            ParameterSetName = "WebApplication_InputFile",
            HelpMessage = "The delimiter used within the input file specified by the -InputFile parameter (the default is \"|\" if not specified).")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Farm_InputFile",
            HelpMessage = "The delimiter used within the input file specified by the -InputFile parameter (the default is \"|\" if not specified).")]
        public string InputFileDelimiter { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "List_XmlInputFile",
            HelpMessage = "An XML file or XmlDocument containing the replacements to make: <Replacements><Replacement><SearchString>string</SearchString><ReplaceString>string</ReplaceString></Replacement></Replacements>")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Web_XmlInputFile",
            HelpMessage = "An XML file or XmlDocument containing the replacements to make: <Replacements><Replacement><SearchString>string</SearchString><ReplaceString>string</ReplaceString></Replacement></Replacements>")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Site_XmlInputFile",
            HelpMessage = "An XML file or XmlDocument containing the replacements to make: <Replacements><Replacement><SearchString>string</SearchString><ReplaceString>string</ReplaceString></Replacement></Replacements>")]
        [Parameter(Mandatory = true,
            ParameterSetName = "WebApplication_XmlInputFile",
            HelpMessage = "An XML file or XmlDocument containing the replacements to make: <Replacements><Replacement><SearchString>string</SearchString><ReplaceString>string</ReplaceString></Replacement></Replacements>")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Farm_XmlInputFile",
            HelpMessage = "An XML file or XmlDocument containing the replacements to make: <Replacements><Replacement><SearchString>string</SearchString><ReplaceString>string</ReplaceString></Replacement></Replacements>")]
        public XmlDocumentPipeBind XmlInputFile { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "List_ParameterInput",
            HelpMessage = "A regular expression search string.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Web_ParameterInput",
            HelpMessage = "A regular expression search string.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Site_ParameterInput",
            HelpMessage = "A regular expression search string.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "WebApplication_ParameterInput",
            HelpMessage = "A regular expression search string.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Farm_ParameterInput",
            HelpMessage = "A regular expression search string.")]
        public string SearchString { get; set; }

        [Parameter(Mandatory = true,
            ParameterSetName = "List_ParameterInput",
            HelpMessage = "The string to replace the match with.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Web_ParameterInput",
            HelpMessage = "The string to replace the match with.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Site_ParameterInput",
            HelpMessage = "The string to replace the match with.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "WebApplication_ParameterInput",
            HelpMessage = "The string to replace the match with.")]
        [Parameter(Mandatory = true,
            ParameterSetName = "Farm_ParameterInput",
            HelpMessage = "The string to replace the match with.")]
        public string ReplaceString { get; set; }



        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The list whose content will be replaced.\r\n\r\nThe value must be a valid URL in the form http://server_name",
            ParameterSetName = "List_ParameterInput")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The list whose content will be replaced.\r\n\r\nThe value must be a valid URL in the form http://server_name",
            ParameterSetName = "List_XmlInputFile")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The list whose content will be replaced.\r\n\r\nThe value must be a valid URL in the form http://server_name",
            ParameterSetName = "List_InputFile")]
        public SPListPipeBind List { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Web_ParameterInput",
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Web_XmlInputFile",
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Web_InputFile",
            Position = 0,
            HelpMessage = "Specifies the URL or GUID of the Web containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid name of Microsoft SharePoint Foundation 2010 Web site (for example, MySPSite1); or an instance of a valid SPWeb object.")]
        public SPWebPipeBind Web { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Site_ParameterInput",
            Position = 0,
            HelpMessage = "The site containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Site_XmlInputFile",
            Position = 0,
            HelpMessage = "The site containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Site_InputFile",
            Position = 0,
            HelpMessage = "The site containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid GUID, in the form 12345678-90ab-cdef-1234-567890bcdefgh; a valid URL, in the form http://server_name; or an instance of a valid SPSite object.")]
        public SPSitePipeBind Site { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "WebApplication_ParameterInput",
            Position = 0,
            HelpMessage = "The web application containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "WebApplication_XmlInputFile",
            Position = 0,
            HelpMessage = "The web application containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "WebApplication_InputFile",
            Position = 0,
            HelpMessage = "The web application containing the lists whose content will be replaced.\r\n\r\nThe type must be a valid URL, in the form http://server_name; or an instance of a valid SPWebApplication object.")]
        public SPWebApplicationPipeBind WebApplication { get; set; }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Farm_ParameterInput",
            Position = 0,
            HelpMessage = "Provide the SPFarm object to replace matching content in all lists throughout the farm.")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Farm_XmlInputFile",
            Position = 0,
            HelpMessage = "Provide the SPFarm object to replace matching content in all lists throughout the farm.")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            ParameterSetName = "Farm_InputFile",
            Position = 0,
            HelpMessage = "Provide the SPFarm object to replace matching content in all lists throughout the farm.")]
        public SPFarmPipeBind Farm { get; set; }



        [Parameter(Mandatory = false,
            HelpMessage = "Publish or check-in the item after updating the contents.")]
        public SwitchParameter Publish { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The log file to store all change records to.")]
        [ValidateDirectoryExistsAndValidFileName]
        public string LogFile { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The name of the field to update.")]
        public string[] FieldName { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "If specified then the internal name of the field will be used; otherwise the display name will be used.")]
        public SwitchParameter UseInternalFieldName { get; set; }


        protected override void InternalBeginProcessing()
        {
            base.InternalBeginProcessing();

            Logger.LogFile = LogFile;
        }

        protected override void InternalProcessRecord()
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

            ReplaceFieldValues rfv = new ReplaceFieldValues();
            rfv.LogFile = LogFile;
            rfv.UseInternalFieldName = UseInternalFieldName;
            if (FieldName != null)
            {
                rfv.FieldName.AddRange(FieldName);
            }
            rfv.Publish = Publish.IsPresent;
            rfv.Test = test;
            rfv.Quiet = !Logger.Verbose;

            if (ParameterSetName.EndsWith("_InputFile"))
            {
                rfv.ParseInputFile(InputFile, InputFileDelimiter, false);
            }
            else if (ParameterSetName.EndsWith("_XmlInputFile"))
            {
                rfv.ParseInputFile(XmlInputFile.Read());
            }
            else if (ParameterSetName.EndsWith("_ParameterInput"))
            {
                rfv.SearchStrings.Add(new ReplaceFieldValues.SearchReplaceData(SearchString, ReplaceString));
            }
            if (rfv.SearchStrings.Count == 0)
                throw new SPCmdletException("No search strings were specified.");

            switch (ParameterSetName)
            {
                case "WebApplication_ParameterInput":
                case "WebApplication_XmlInputFile":
                case "WebApplication_InputFile":
                    SPWebApplication webApp1 = WebApplication.Read();
                    if (webApp1 == null)
                        throw new SPException("Web Application not found.");
                    rfv.ReplaceValues(webApp1);
                    break;
                case "Site_ParameterInput":
                case "Site_XmlInputFile":
                case "Site_InputFile":
                    using (SPSite site = Site.Read())
                    {
                        rfv.ReplaceValues(site);
                    }
                    break;
                case "Web_ParameterInput":
                case "Web_XmlInputFile":
                case "Web_InputFile":
                    using (SPWeb web = Web.Read())
                    {
                        try
                        {
                            rfv.ReplaceValues(web);
                        }
                        finally
                        {
                            web.Site.Dispose();
                        }
                    }
                    break;
                case "List_ParameterInput":
                case "List_XmlInputFile":
                case "List_InputFile":
                    SPList list = List.Read();
                    try
                    {
                        rfv.ReplaceValues(list);
                    }
                    finally
                    {
                        list.ParentWeb.Dispose();
                        list.ParentWeb.Site.Dispose();
                    }
                    break;
                default:
                    SPFarm farm = Farm.Read();
                    foreach (SPService svc in farm.Services)
                    {
                        if (!(svc is SPWebService))
                            continue;

                        foreach (SPWebApplication webApp2 in ((SPWebService)svc).WebApplications)
                        {
                            rfv.ReplaceValues(webApp2);
                        }
                    }
                    break;
            }
        }

    }
}
