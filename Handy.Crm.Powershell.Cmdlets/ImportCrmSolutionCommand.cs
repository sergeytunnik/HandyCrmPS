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
        private string AbsolutePath
        {
            get
            {
                return System.IO.Path.IsPathRooted(Path)
                    ? Path
                    : System.IO.Path.GetFullPath(System.IO.Path.Combine(SessionState.Path.CurrentLocation.Path, Path));
            }
        }

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
        [Alias("Holding")]
        public SwitchParameter HoldingSolution { get; set; }

        [Parameter(Mandatory = false)]
        [Alias("Async")]
        public SwitchParameter Asynchronous { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            byte[] fileBytes = File.ReadAllBytes(AbsolutePath);
            Guid importJobId = Guid.NewGuid();
            Guid asyncJobId = Guid.Empty;

            ImportSolutionRequest impSolutionRequest = new ImportSolutionRequest()
            {
                ConvertToManaged = ConvertToManaged,
                CustomizationFile = fileBytes,
                HoldingSolution = HoldingSolution,
                ImportJobId = importJobId,
                OverwriteUnmanagedCustomizations = OverwriteCustomizations,
                PublishWorkflows = PublishWorkflows,
                SkipProductUpdateDependencies = SkipProductUpdateDependencies
            };

            WriteVerbose(string.Format("Starting importing solution from {0}", AbsolutePath));

            if (Asynchronous)
            {
                var response = (ExecuteAsyncResponse)organizationService.Execute(
                    new ExecuteAsyncRequest
                    {
                        Request = impSolutionRequest
                    });

                asyncJobId = response.AsyncJobId;
            }
            else
            {
                organizationService.Execute(impSolutionRequest);
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
