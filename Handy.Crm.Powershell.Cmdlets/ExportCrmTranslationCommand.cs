using Microsoft.Crm.Sdk.Messages;
using System.IO;
using System.Management.Automation;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsData.Export, "CRMTranslation")]
    public class ExportCrmTranslationCommand : CrmCmdletBase
    {
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string SolutionName { get; set; }

        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string FilePath { get; set; }

        protected override void ProcessRecord()
        {
            var request = new ExportTranslationRequest()
            {
                SolutionName = SolutionName
            };

            var response = (ExportTranslationResponse)Connection.Execute(request);

            var absoluteFilePath = GetAbsoluteFilePath(FilePath);
            WriteVerbose(string.Format("Saving translations ({0}) at {1}", SolutionName, absoluteFilePath));
            File.WriteAllBytes(absoluteFilePath, response.ExportTranslationFile);
        }
    }
}
