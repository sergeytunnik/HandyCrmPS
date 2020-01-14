using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsData.Import, "CRMSolution")]
    [OutputType("System.Collections.Generic.Dictionary<string, Guid>")]
    public class ImportCrmSolutionCommand : CrmCmdletBase
    {
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string Path { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter ConvertToManaged { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter OverwriteCustomizations { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter PublishWorkflows { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter SkipProductUpdateDependencies { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter SkipRibbonMetadataProcessing { get; set; }

        [Parameter(Mandatory = false)]
        [Alias("Holding")]
        public SwitchParameter HoldingSolution { get; set; }

        [Parameter(Mandatory = false)]
        [Alias("Async")]
        public SwitchParameter Asynchronous { get; set; }

        protected override void ProcessRecord()
        {
            var absoluteFilePath = GetAbsoluteFilePath(Path);
            var fileBytes = File.ReadAllBytes(absoluteFilePath);
            var importJobId = Guid.NewGuid();
            var asyncJobId = Guid.Empty;

            ImportSolutionRequest impSolutionRequest = new ImportSolutionRequest()
            {
                ConvertToManaged = ConvertToManaged,
                CustomizationFile = fileBytes,
                HoldingSolution = HoldingSolution,
                ImportJobId = importJobId,
                OverwriteUnmanagedCustomizations = OverwriteCustomizations,
                PublishWorkflows = PublishWorkflows,
                SkipProductUpdateDependencies = SkipProductUpdateDependencies,
                SkipRibbonMetadataProcessing = SkipRibbonMetadataProcessing
            };

            WriteVerbose(string.Format("Starting importing solution from {0}", absoluteFilePath));

            if (Asynchronous)
            {
                var response = (ExecuteAsyncResponse)Connection.Execute(
                    new ExecuteAsyncRequest
                    {
                        Request = impSolutionRequest
                    });

                asyncJobId = response.AsyncJobId;
            }
            else
            {
                Connection.Execute(impSolutionRequest);
                WriteVerbose("Finished importing solution");
            }

            WriteObject(new Dictionary<string, Guid>()
            {
                {"importjobid", importJobId },
                {"asyncjobid", asyncJobId }
            });
        }
    }
}
