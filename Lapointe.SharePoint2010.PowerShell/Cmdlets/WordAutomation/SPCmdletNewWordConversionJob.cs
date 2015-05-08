using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePoint.PowerShell.Cmdlets.Lists;
using Lapointe.SharePoint.PowerShell.Cmdlets.PipeBindObjects;
using Microsoft.SharePoint.PowerShell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Microsoft.Office.Word.Server.Conversions;
using Microsoft.SharePoint;
using Microsoft.Office.Word.Server.Service;
using Microsoft.SharePoint.Administration;
using System.Threading;

namespace Lapointe.SharePoint.PowerShell.Cmdlets.WordAutomation
{
    [Cmdlet(VerbsCommon.New, "SPWordConversionJob", SupportsShouldProcess = true, DefaultParameterSetName = "Library"),
     SPCmdlet(RequireLocalFarmExist = true, RequireUserFarmAdmin = false)]
    [CmdletGroup("Word Automation")]
    [CmdletDescription("Creates a new conversion job to convert one or more documents from one format to another.", "Creates a new conversion job to convert one or more documents from one format to another. This cmdlet leverages the Word Automation Service to do the conversion. Using Word Automation Services, you can convert from Open XML WordprocessingML to other document formats. For example, you may want to convert many documents to the PDF format and spool them to a printer or send them by e-mail to your customers. Or, you can convert from other document formats (such as HTML or Word 97-2003 binary documents) to Open XML word-processing documents. By default the timer job that does the conversions is schedule to run every 15 minutes - you can change this by editing the \"Word Automation Services Timer Job\" timer job (or manually tell it to run immediately). If you specified the Wait parameter then it may take up to 15 minutes to respond if you have not changed this default.")]
    [RelatedCmdlets(typeof(SPCmdletGetList), typeof(SPCmdletGetWordConversionJobStatus))]
    [Example(Code = "PS C:\\> New-SPWordConversionJob -InputList \"http://server_name/WordDocs\" -OutputList \"http://server_name/PDFDocs\" -OutputFormt PDF -OutputSaveBehavior AppendIfPossible -Wait",
        Remarks = "This example converts the documents located in http://server_name/WordDocs and stores the converted items in http://server_name/PDFDocs.")]
    [Example(Code = "PS C:\\> New-SPWordConversionJob -InputFile \"http://server_name/WordDocs/report.docx\" -OutputFile \"http://server_name/PDFDocs/report.pdf\" -OutputFormt PDF -OutputSaveBehavior AppendIfPossible -Wait",
        Remarks = "This example converts the document located in http://server_name/WordDocs/report.docx and stores the converted item as http://server_name/PDFDocs/report.pdf.")]
    [Example(Code = "PS C:\\> New-SPWordConversionJob -InputFolder \"http://server_name/WordDocs/Reports\" -OutputFolder \"http://server_name/PDFDocs/Reports\" -OutputFormt PDF -OutputSaveBehavior AppendIfPossible -Wait",
        Remarks = "This example converts the document located in http://server_name/WordDocs/Reports and stores the converted items in http://server_name/PDFDocs/Reports.")]
    public class SPCmdletNewWordConversionJob : SPNewCmdletBase<ConversionJob>
    {
        public SPCmdletNewWordConversionJob()
        {
            MarkupView = MarkupTypes.Comments | MarkupTypes.Ink | MarkupTypes.Text | MarkupTypes.Formatting;
            OutputFormat = SaveFormat.Automatic;
            OutputSaveBehavior = SaveBehavior.AppendIfPossible;
            CompatibilityMode = Microsoft.Office.Word.Server.Conversions.CompatibilityMode.MaintainCurrentSetting;
            RevisionState = Microsoft.Office.Word.Server.Conversions.RevisionState.FinalShowingMarkup;
        }

        [Parameter(ParameterSetName = "Library",
        Mandatory = true,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 0,
        HelpMessage = "The input library whose items will be converted and copied to the output list.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPListPipeBind InputList { get; set; }

        [Parameter(ParameterSetName = "Library", 
        Mandatory = true,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 1,
        HelpMessage = "The output library where the converted items will be stored.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPListPipeBind OutputList { get; set; }

        [Parameter(ParameterSetName = "File",
        Mandatory = true,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 0,
        HelpMessage = "The input file that will be converted and copied to the output file.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPFilePipeBind InputFile { get; set; }

        [Parameter(ParameterSetName = "File",
        Mandatory = true,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 1,
        HelpMessage = "The output file where the converted item will be copied to.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public string OutputFile { get; set; }


        [Parameter(ParameterSetName = "Folder",
        Mandatory = true,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 0,
        HelpMessage = "The input library folder whose items will be converted and copied to the output list.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPFolderPipeBind InputFolder { get; set; }

        [Parameter(ParameterSetName = "Folder",
        Mandatory = true,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true,
        Position = 1,
        HelpMessage = "The output library folder where the converted items will be stored.\r\n\r\nThe value must be a valid URL in the form http://server_name")]
        public SPFolderPipeBind OutputFolder { get; set; }

        [Parameter(ParameterSetName = "Folder",
        HelpMessage = "Recursively convert all items in the folder's sub-folders.")]
        public SwitchParameter Recurse { get; set; }

        [Parameter(Mandatory = true,
        Position = 3,
        HelpMessage = "Specifies the Save format for this conversion job. The extension associated with the Save format specified here is appended to the output file if the extension is not already present. For example, when converting to PDF, a document with the output file name http://contoso/output/output.pdf would remain unchanged; a document with the output file name http://contoso/output/output.docx would be changed to http://contoso/output/output.docx.pdf.")]
        public SaveFormat OutputFormat { get; set; }

        [Parameter(HelpMessage = "Specifies the behavior that should be applied when saving converted files to existing file names.")]
        public SaveBehavior OutputSaveBehavior { get; set; }

        [Parameter(HelpMessage = "Specifies the appropriate compatibility mode for the output file. If the file is not an Open XML File format document, this setting is ignored. MaintainCurrentSetting specifies that files maintain their compatibility mode as follows:\r\n\t• Binary files and files in Word 97-2003 compatibility mode stay in that mode.\r\n\t• Word 2007 documents or documents in Word 2007 compatibility mode stay in that mode.\r\n\t• Documents upgraded to Word 2010 stay upgraded.")]
        public CompatibilityMode CompatibilityMode { get; set; }

        [Parameter(HelpMessage = "Indicates whether the document is saved with an added thumbnail. Setting this property has the same effect as checking Save Thumbnail in the Save As dialog in Word.")]
        public SwitchParameter AddThumbnail { get; set; }

        [Parameter(HelpMessage = "Indicates whether any fields in the document are automatically updated when the document is opened.")]
        public SwitchParameter UpdateFields { get; set; }

        [Parameter(HelpMessage = "Indicates whether fonts used within the document are obfuscated and saved within the output. This setting uses the same obfuscation mechanisms as Word.")]
        public SwitchParameter EmbedFonts { get; set; }

        [Parameter(HelpMessage = "Specifies the type(s) of markup that should be shown in the document.The possible values for this property correspond to the options available on the ribbon (Tracking group, Show Markup dropdown). The default value for this property is @(Comments, Ink, Text, Formatting), which means all types of markup are shown.")]
        public MarkupTypes MarkupView { get; set; }

        [Parameter(HelpMessage = "Specifies the visibility of revisions in the document.")]
        public RevisionState RevisionState { get; set; }

        [Parameter(HelpMessage = "Indicates whether to restrict the characters that are included in the embedded font to only those characters that are required by the current document.")]
        public SwitchParameter SubsetEmbeddedFonts { get; set; }

        [Parameter(HelpMessage = "If specified then the cmdlet will block until the conversion completes.")]
        public SwitchParameter Wait { get; set; }

        protected override ConversionJob CreateDataObject()
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

            SPWeb contextWeb = null;
            object input = null;
            object output = null;
            if (ParameterSetName == "Library")
            {
                input = InputList.Read();
                output = OutputList.Read();
                if (input != null)
                    contextWeb = ((SPList)input).ParentWeb;
            }
            else if (ParameterSetName == "Folder")
            {
                input = InputFolder.Read();
                output = OutputFolder.Read();
                if (input != null)
                    contextWeb = ((SPFolder)input).ParentWeb;
            }
            else if (ParameterSetName == "File")
            {
                input = InputFile.Read();
                if (!((SPFile)input).Exists)
                    throw new Exception("The specified input file does not exist.");

                output = OutputFile;
                if (input != null)
                    contextWeb = ((SPFile)input).ParentFolder.ParentWeb;
                if (contextWeb != null)
                    input = contextWeb.Site.MakeFullUrl(((SPFile)input).ServerRelativeUrl);
            }
            if (input == null)
                throw new Exception("The input can not be a null or empty value.");
            if (output == null)
                throw new Exception("The output can not be a null or empty value.");

            WordServiceApplicationProxy proxy = GetWordServiceApplicationProxy(contextWeb.Site.WebApplication);

            ConversionJobSettings settings = new ConversionJobSettings();
            settings.OutputFormat = OutputFormat;
            settings.OutputSaveBehavior = OutputSaveBehavior;
            settings.UpdateFields = UpdateFields;
            settings.AddThumbnail = AddThumbnail;
            settings.CompatibilityMode = CompatibilityMode;
            settings.EmbedFonts = EmbedFonts;
            settings.SubsetEmbeddedFonts = SubsetEmbeddedFonts;
            settings.MarkupView = MarkupView;
            settings.RevisionState = RevisionState;
            ConversionJob job = new ConversionJob(proxy, settings);
            job.UserToken = contextWeb.CurrentUser.UserToken;
            if (ParameterSetName == "Library")
            {
                job.AddLibrary((SPList)input, (SPList)output);
            }
            else if (ParameterSetName == "Folder")
            {
                job.AddFolder((SPFolder)input, (SPFolder)output, Recurse);
            }
            else if (ParameterSetName == "File")
            {
                job.AddFile((string)input, (string)output);
            }

            job.Start();

            if (Wait)
            {
                ConversionJobStatus jobStatus = null;
                do
                {
                    Thread.Sleep(1000);
                    jobStatus = new ConversionJobStatus(proxy, job.JobId, null);
                } while (jobStatus.Failed == 0 && jobStatus.Succeeded == 0);
            }
            return job;
        }


        internal static WordServiceApplicationProxy GetWordServiceApplicationProxy(SPWebApplication webApp)
        {
            if (!webApp.ServiceApplicationProxyGroup.ContainsType(typeof(WordServiceApplicationProxy)))
            {
                throw new Exception("No Word Automation Service Application Proxy associated with this Web Application");
            }
            else
            {
                foreach (SPServiceApplicationProxy proxy in webApp.ServiceApplicationProxyGroup.Proxies)
                {
                    if (proxy.GetType() == typeof(WordServiceApplicationProxy))
                        return (WordServiceApplicationProxy)proxy;
                }
                return null;
            }
        }


    }
}
