using Microsoft.Crm.Sdk.Messages;
using System;
using System.IO;
using System.Management.Automation;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsData.Import, "CRMTranslation")]
    [OutputType(typeof(Guid))]
    public class ImportCrmTranslationCommand : CrmCmdletBase
    {
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string FilePath { get; set; }

        protected override void ProcessRecord()
        {
            var importJobId = Guid.NewGuid();
            var fileBytes = File.ReadAllBytes(GetAbsoluteFilePath(FilePath));

            var request = new ImportTranslationRequest()
            {
                ImportJobId = importJobId,
                TranslationFile = fileBytes
            };

            Connection.Execute(request);

            WriteObject(importJobId);
        }
    }
}
